"""VisioClient: every COM call in the codebase lives in this module.

Rules that keep this correct:
- All methods are called on the ComWorker's STA thread (see runtime.py).
- Methods return plain dicts/lists only — COM proxies must never escape.
- win32com is imported lazily inside the default app factory, so this module
  imports fine on macOS and tests inject a FakeVisio app factory instead.
- COM collections are 1-based and iterated explicitly via Count/Item so a
  duck-typed fake works without implementing COM enumerators.
"""

from __future__ import annotations

import difflib
import os
import re
import sys
from typing import Any, Callable, Optional

from . import constants as C
from . import routing
from .errors import VisioMcpError, hresult_of, translate_com_error
from .models import DropSpec, hex_to_rgb_formula

_CONTROL_CHARS = dict.fromkeys(i for i in range(32) if i not in (9, 10, 13))


def _clean_text(raw: Any) -> str:
    """Visio shape text can embed field-escape control chars; strip them."""
    return str(raw).translate(_CONTROL_CHARS) if raw else ""


def _is_com_error(exc: BaseException) -> bool:
    return type(exc).__name__ == "com_error"


def _items(collection):
    """Iterate a 1-based COM collection via Count/Item."""
    for i in range(1, int(collection.Count) + 1):
        yield collection.Item(i)


def _icons_dir() -> str:
    """Folder where users drop labeled custom icons (VISIO_MCP_ICONS_DIR)."""
    return os.environ.get(C.ICONS_DIR_ENV) or C.DEFAULT_ICONS_DIR


def _norm_label(text: Any) -> str:
    """Normalize a label/query so 'Private Endpoint', 'private_endpoint' and
    'private-endpoint' all compare equal."""
    chars = ["-" if ch in " _-" else ch for ch in str(text).strip().lower()]
    normalized = "".join(chars)
    while "--" in normalized:
        normalized = normalized.replace("--", "-")
    return normalized.strip("-")


def _list_local_icons() -> list[dict]:
    """Image files in the local icon folder, as {name, path, ext} rows."""
    directory = _icons_dir()
    icons: list[dict] = []
    try:
        entries = sorted(os.listdir(directory))
    except OSError:
        return icons
    for filename in entries:
        full = os.path.join(directory, filename)
        stem, ext = os.path.splitext(filename)
        if ext.lower() in C.ICON_EXTS and os.path.isfile(full):
            icons.append({"name": stem, "path": full, "ext": ext.lower().lstrip(".")})
    return icons


def _ext_rank(path: str) -> int:
    ext = os.path.splitext(path)[1].lower()
    return C.ICON_EXTS.index(ext) if ext in C.ICON_EXTS else len(C.ICON_EXTS)


def _find_local_icon(name: str) -> Optional[str]:
    """Resolve a drop name to a local icon file path (None if no clear match).

    Prefers an exact normalized-label match, then a unique-label substring
    match; when a label exists in several formats the ICON_EXTS order decides.
    """
    if not name or not str(name).strip():
        return None
    target = _norm_label(name)
    icons = _list_local_icons()
    exact = [ic["path"] for ic in icons if _norm_label(ic["name"]) == target]
    if exact:
        return sorted(exact, key=_ext_rank)[0]
    subs = [ic for ic in icons if target and target in _norm_label(ic["name"])]
    if len({_norm_label(ic["name"]) for ic in subs}) == 1:
        return sorted((ic["path"] for ic in subs), key=_ext_rank)[0]
    return None


def _stencil_files_under(root: str) -> list[str]:
    """All .vss/.vssx files under root, case-insensitive (stencil filenames
    are conventionally uppercase)."""
    hits = []
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            if fn.lower().endswith((".vssx", ".vss")):
                hits.append(os.path.join(dirpath, fn))
    return sorted(hits)


def default_app_factory():
    """Attach to a running Visio instance, or launch a new visible one."""
    if sys.platform != "win32":
        raise VisioMcpError(
            "This server drives Microsoft Visio via COM and must run on Windows "
            "with Visio desktop installed. It cannot control Visio from "
            f"{sys.platform!r}."
        )
    import win32com.client  # noqa: PLC0415 — Windows-only import by design

    try:
        return win32com.client.GetActiveObject("Visio.Application")
    except Exception as exc:
        if not _is_com_error(exc) or hresult_of(exc) != C.MK_E_UNAVAILABLE:
            raise translate_com_error(exc, "attaching to Visio") from exc
    try:
        return win32com.client.DispatchEx("Visio.Application")
    except Exception as exc:
        if _is_com_error(exc):
            raise translate_com_error(exc, "launching Visio") from exc
        raise


class VisioClient:
    def __init__(self, app_factory: Optional[Callable[[], Any]] = None):
        self._app_factory = app_factory or default_app_factory
        self._app_obj: Any = None

    # -- plumbing -----------------------------------------------------------

    def release(self) -> None:
        """Drop COM references. Never quits Visio — the user keeps their app."""
        self._app_obj = None

    def _app(self):
        if self._app_obj is not None:
            try:
                _ = self._app_obj.Version  # liveness probe: user may have quit Visio
            except Exception:
                self._app_obj = None
        if self._app_obj is None:
            app = self._app_factory()
            app.Visible = True
            try:
                # 7 = IDNO: auto-answer modal prompts so COM calls never hang
                app.AlertResponse = 7
            except Exception:
                pass
            self._app_obj = app
        return self._app_obj

    def _guard(self, context: str, fn: Callable[[], Any]) -> Any:
        """Run fn, translating com_error into an actionable VisioMcpError."""
        try:
            return fn()
        except VisioMcpError:
            raise
        except Exception as exc:
            if _is_com_error(exc):
                raise translate_com_error(exc, context) from exc
            raise

    def _drawing_doc(self):
        app = self._app()
        doc = None
        try:
            doc = app.ActiveDocument
        except Exception:
            doc = None
        if doc is not None and int(doc.Type) == C.VIS_TYPE_DRAWING:
            return doc
        for d in _items(app.Documents):
            if int(d.Type) == C.VIS_TYPE_DRAWING:
                return d
        raise VisioMcpError(
            "No drawing document is open in Visio — call create_document or "
            "open_document first."
        )

    def _page(self, page_name: Optional[str] = None):
        app = self._app()
        doc = self._drawing_doc()
        if page_name is None:
            try:
                page = app.ActivePage
                if page is not None:
                    return page
            except Exception:
                pass
            return doc.Pages.Item(1)
        for p in _items(doc.Pages):
            if str(p.Name) == page_name or str(p.NameU) == page_name:
                return p
        names = [str(p.Name) for p in _items(doc.Pages)]
        raise VisioMcpError(f"No page named {page_name!r}. Pages: {names}")

    def _shape_by_id(self, page, shape_id: int):
        try:
            return page.Shapes.ItemFromID(int(shape_id))
        except Exception:
            raise VisioMcpError(
                f"Shape {shape_id} not found on page {str(page.Name)!r} — call "
                "get_page_state to list current shape ids."
            ) from None

    def _stencil_docs(self) -> list:
        app = self._app()
        return [d for d in _items(app.Documents) if int(d.Type) == C.VIS_TYPE_STENCIL]

    # -- status / documents --------------------------------------------------

    def _builtin_content_dir(self) -> Optional[str]:
        """Visio ships ~130 stencils (incl. Azure/AWS) under
        <install>\\Visio Content\\<locale>\\ — resolvable by bare filename."""
        try:
            base = str(self._app().Path)
        except Exception:
            return None
        root = os.path.join(base, "Visio Content")
        return root if os.path.isdir(root) else None

    def _builtin_stencil_files(self) -> list[str]:
        root = self._builtin_content_dir()
        return _stencil_files_under(root) if root else []

    def status(self) -> dict:
        app = self._app()
        docs, stencils = [], []
        for d in _items(app.Documents):
            entry = {"name": str(d.Name), "path": str(d.FullName)}
            if int(d.Type) == C.VIS_TYPE_STENCIL:
                stencils.append(entry)
            else:
                docs.append(entry)
        active_doc = active_page = None
        try:
            if app.ActiveDocument is not None:
                active_doc = str(app.ActiveDocument.Name)
            if app.ActivePage is not None:
                active_page = str(app.ActivePage.Name)
        except Exception:
            pass
        builtin_files = self._builtin_stencil_files()
        cloud = [
            os.path.basename(f)
            for f in builtin_files
            if any(k in os.path.basename(f).lower() for k in ("azure", "aws"))
        ]
        return {
            "visio_version": str(app.Version),
            "documents": docs,
            "stencils": stencils,
            "active_document": active_doc,
            "active_page": active_page,
            "my_shapes_path": str(app.MyShapesPath),
            "builtin_stencils_path": self._builtin_content_dir(),
            "builtin_stencil_count": len(builtin_files),
            "builtin_cloud_stencils": cloud[:60],
            "note": (
                "Built-in stencils open by bare filename, e.g. "
                "open_stencil('AZURESTORAGE_U.VSSX') — no My Shapes install needed."
                if cloud else
                "No built-in Azure/AWS stencils detected; download stencil packs "
                "into the My Shapes folder for cloud icons."
            ),
        }

    def create_document(self, template: Optional[str] = None, measurement: str = "us") -> dict:
        app = self._app()
        name = template or ""
        if name and "." not in os.path.basename(name):
            # bare family name like 'BASFLO' -> BASFLO_U.VSTX / BASFLO_M.VSTX
            name = f"{name}_{'M' if measurement == 'metric' else 'U'}.VSTX"

        def _add():
            return app.Documents.Add(name)

        doc = self._guard(
            f"creating document from template {name!r}" if name else "creating blank document",
            _add,
        )
        return {
            "document": str(doc.Name),
            "pages": [str(p.Name) for p in _items(doc.Pages)],
            "note": "Coordinates are in inches, origin at the page bottom-left; "
                    "drop point is the shape center.",
        }

    def open_document(self, path: str) -> dict:
        app = self._app()
        abs_path = os.path.abspath(os.path.expanduser(path))
        if not os.path.exists(abs_path):
            raise VisioMcpError(f"File not found: {abs_path}")
        doc = self._guard(f"opening {abs_path}", lambda: app.Documents.Open(abs_path))
        return {
            "document": str(doc.Name),
            "path": str(doc.FullName),
            "pages": [str(p.Name) for p in _items(doc.Pages)],
            "note": "Call get_page_state to learn the shape ids on each page.",
        }

    def save_document(self, path: Optional[str] = None) -> dict:
        doc = self._drawing_doc()
        if path is None:
            existing = str(doc.Path or "")
            if not existing:
                raise VisioMcpError(
                    "This document has never been saved — pass an absolute file "
                    "path ending in .vsdx to save_document."
                )
            self._guard("saving document", doc.Save)
            return {"path": str(doc.FullName)}
        abs_path = os.path.abspath(os.path.expanduser(path))
        if not abs_path.lower().endswith((".vsdx", ".vsdm", ".vsd")):
            abs_path += ".vsdx"
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        self._guard(f"saving document to {abs_path}", lambda: doc.SaveAs(abs_path))
        return {"path": str(doc.FullName)}

    def export_page_png(self, path: str, page_name: Optional[str] = None) -> dict:
        page = self._page(page_name)
        abs_path = os.path.abspath(os.path.expanduser(path))
        if not abs_path.lower().endswith(".png"):
            abs_path += ".png"
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        self._guard(f"exporting page to {abs_path}", lambda: page.Export(abs_path))
        if not os.path.exists(abs_path):
            raise VisioMcpError(f"Visio reported success but no file was written at {abs_path}")
        return {"path": abs_path, "page": str(page.Name)}

    # -- stencils / masters ---------------------------------------------------

    def open_stencil(self, name_or_path: str) -> dict:
        app = self._app()
        query = name_or_path.strip()
        candidates: list[str] = []
        if os.path.isabs(query):
            candidates.append(query)
        else:
            candidates.append(query)  # Visio resolves bare names via its search paths
            my_shapes = str(app.MyShapesPath)
            candidates.append(os.path.join(my_shapes, query))
            if "." not in os.path.basename(query):
                candidates.append(os.path.join(my_shapes, query + ".vssx"))
                candidates.append(os.path.join(my_shapes, query + ".vss"))
            # last resort: fuzzy filename match under My Shapes, then Visio's
            # built-in content folder (which ships Azure/AWS stencils)
            fuzzy_pool = _stencil_files_under(my_shapes) + self._builtin_stencil_files()
            for hit in fuzzy_pool:
                if query.lower() in os.path.basename(hit).lower():
                    candidates.append(hit)

        errors: list[str] = []
        for cand in candidates:
            try:
                doc = app.Documents.OpenEx(cand, C.STENCIL_OPEN_FLAGS)
                return {
                    "stencil": str(doc.Name),
                    "path": str(doc.FullName),
                    "master_count": int(doc.Masters.Count),
                    "note": "Use find_masters to list its shapes.",
                }
            except Exception as exc:
                errors.append(f"{cand}: {exc}")
        raise VisioMcpError(
            f"Could not open a stencil matching {name_or_path!r}. Tried Visio's "
            f"search paths, {str(app.MyShapesPath)!r}, and the built-in Visio "
            "Content folder. Call visio_status to see builtin_cloud_stencils "
            "(Visio ships Azure/AWS stencils out of the box); for other vendor "
            "packs, download and unzip into the My Shapes folder. "
            f"Attempts: {errors[:3]}"
        )

    def _master_sources(self) -> list[tuple[str, Any]]:
        """(source_name, masters_collection) for all open stencils + the doc itself."""
        sources = [(str(d.Name), d.Masters) for d in self._stencil_docs()]
        try:
            doc = self._drawing_doc()
            sources.append((f"(document) {doc.Name}", doc.Masters))
        except VisioMcpError:
            pass
        return sources

    def find_masters(self, query: Optional[str] = None, stencil: Optional[str] = None) -> dict:
        sources = self._master_sources()
        if stencil is not None:
            sources = [s for s in sources if stencil.lower() in s[0].lower()]
            if not sources:
                raise VisioMcpError(
                    f"No open stencil matches {stencil!r} — call open_stencil first, "
                    "or omit the stencil argument to search all open stencils."
                )
        total_masters = sum(int(m.Count) for _, m in sources)
        rows = []
        q = (query or "").lower()
        for source_name, masters in sources:
            for m in _items(masters):
                name = str(m.Name)
                if q and q not in name.lower():
                    continue
                rows.append({"master": name, "stencil": source_name})
                if len(rows) >= 100:
                    return {"masters": rows, "truncated": True}
        # Local custom icons are discoverable alongside masters (unless a
        # specific stencil was requested, which only applies to Visio masters).
        local_rows = []
        if stencil is None:
            qn = _norm_label(query or "")
            for ic in _list_local_icons():
                if q and q not in ic["name"].lower() and (
                    not qn or qn not in _norm_label(ic["name"])
                ):
                    continue
                local_rows.append(
                    {"master": ic["name"], "stencil": "(local icon)", "path": ic["path"]}
                )
        for row in local_rows:
            rows.append(row)
            if len(rows) >= 100:
                return {"masters": rows, "truncated": True}
        if total_masters == 0 and not local_rows:
            raise VisioMcpError(
                "No stencils are open — call open_stencil first (e.g. "
                "open_stencil('BASFLO_U.VSSX') for basic flowchart shapes)."
            )
        return {"masters": rows, "truncated": False}

    def list_local_icons(self) -> dict:
        """Report the local custom-icon folder and the labeled files in it.

        No COM needed — this just inspects the folder users drop icons into.
        """
        directory = _icons_dir()
        try:
            os.makedirs(directory, exist_ok=True)
        except OSError:
            pass
        return {
            "icons_dir": directory,
            "exists": os.path.isdir(directory),
            "icons": _list_local_icons(),
            "note": (
                "Save labeled image files here (e.g. 'subnet.svg', "
                "'private-endpoint.png'); the label is the file name without "
                "its extension. Labels are discoverable via find_masters and "
                "can be placed with drop_shape('<label>', x, y) when Visio has "
                "no built-in master."
            ),
        }

    def _find_master(self, master: str, stencil: Optional[str] = None, _alias_depth: int = 0):
        sources = self._master_sources()
        if stencil is not None:
            sources = [s for s in sources if stencil.lower() in s[0].lower()]
        all_names: list[str] = []
        substring_hits = []
        for _, masters in sources:
            for m in _items(masters):
                name = str(m.Name)
                all_names.append(name)
                if name.lower() == master.lower():
                    return m
                if master.lower() in name.lower():
                    substring_hits.append(m)
        if len(substring_hits) == 1:
            return substring_hits[0]
        if len(substring_hits) > 1:
            names = sorted({str(m.Name) for m in substring_hits})[:8]
            raise VisioMcpError(
                f"Master {master!r} is ambiguous — matches {names}. Use the exact name."
            )
        key = master.strip().lower()
        if key in C.MASTER_HINTS:
            raise VisioMcpError(f"No master named {master!r}: {C.MASTER_HINTS[key]}")
        alias = C.MASTER_ALIASES.get(key)
        if alias and _alias_depth == 0:
            try:
                return self._find_master(alias, stencil, _alias_depth=1)
            except VisioMcpError:
                pass  # alias target not in the open stencils either
        suggestions = difflib.get_close_matches(master, all_names, n=5, cutoff=0.4)
        hint = f" Closest matches: {suggestions}." if suggestions else ""
        alias_hint = (
            f" This is usually drawn with the {alias!r} master — open the "
            "stencil that contains it." if alias else ""
        )
        raise VisioMcpError(
            f"No master named {master!r} found in the open stencils.{alias_hint}{hint} "
            "Call find_masters to browse, or open_stencil to load the stencil "
            "that contains it."
        )

    # -- shapes ---------------------------------------------------------------

    def _shape_info(self, shape) -> dict:
        return {
            "shape_id": int(shape.ID),
            "name": str(shape.NameU),
            "x": round(float(shape.CellsU("PinX").ResultIU), 4),
            "y": round(float(shape.CellsU("PinY").ResultIU), 4),
            "width_in": round(float(shape.CellsU("Width").ResultIU), 4),
            "height_in": round(float(shape.CellsU("Height").ResultIU), 4),
            "text": _clean_text(shape.Text),
        }

    def drop_shapes(self, specs: list[dict], page_name: Optional[str] = None) -> dict:
        app = self._app()
        page = self._page(page_name)
        parsed = [DropSpec.model_validate(s) for s in specs]
        results = []
        scope = self._guard("starting undo scope", lambda: app.BeginUndoScope("Drop shapes"))
        try:
            for spec in parsed:
                try:
                    m = self._find_master(spec.master, spec.stencil)
                except VisioMcpError:
                    icon_path = _find_local_icon(spec.master)
                    if icon_path is None:
                        raise
                    w, h = spec.width_in, spec.height_in
                    if w is None and h is None:
                        w = C.LOCAL_ICON_DEFAULT_IN
                    info = self.import_image(icon_path, spec.x, spec.y, w, h, page_name)
                    results.append({**info, "master": spec.master, "source": "local_icon"})
                    continue
                shape = self._guard(
                    f"dropping master {spec.master!r}",
                    lambda m=m, s=spec: page.Drop(m, s.x, s.y),
                )
                if spec.width_in is not None:
                    shape.CellsU("Width").ResultIU = spec.width_in
                if spec.height_in is not None:
                    shape.CellsU("Height").ResultIU = spec.height_in
                if spec.text is not None:
                    shape.Text = spec.text
                results.append({**self._shape_info(shape), "master": str(m.Name)})
            self._guard("closing undo scope", lambda: app.EndUndoScope(scope, True))
        except Exception:
            try:
                app.EndUndoScope(scope, False)
            except Exception:
                pass
            raise
        return {"page": str(page.Name), "shapes": results}

    def import_image(
        self,
        path: str,
        x: float,
        y: float,
        width_in: Optional[float] = None,
        height_in: Optional[float] = None,
        page_name: Optional[str] = None,
    ) -> dict:
        """Import an external image file (SVG/PNG/JPG/EMF/...) as a page shape.

        Used for icons that have no Visio master (e.g. Azure's 'Subnet' glyph)
        or vendor logos. `path` must be a LOCAL file; remote URLs are fetched
        by the server tool before this runs. (x, y) is the shape center in
        inches. When only one of width/height is given the other is derived
        from the image's native aspect ratio.
        """
        if not path or not str(path).strip():
            raise VisioMcpError("path must not be empty")
        if not os.path.isfile(path):
            raise VisioMcpError(
                f"Image file not found: {path!r}. Pass an absolute path to a "
                "local image file (download remote icons first, or use the "
                "import_image tool with an http(s) URL)."
            )
        for name, value in (("width_in", width_in), ("height_in", height_in)):
            if value is not None and value <= 0:
                raise VisioMcpError(f"{name} must be positive, got {value}")
        page = self._page(page_name)

        def _import():
            shape = page.Import(path)
            native_w = float(shape.CellsU("Width").ResultIU)
            native_h = float(shape.CellsU("Height").ResultIU)
            w, h = width_in, height_in
            if w is not None and h is None and native_w:
                h = w * (native_h / native_w)
            elif h is not None and w is None and native_h:
                w = h * (native_w / native_h)
            if w is not None:
                shape.CellsU("Width").ResultIU = w
            if h is not None:
                shape.CellsU("Height").ResultIU = h
            shape.CellsU("PinX").ResultIU = x
            shape.CellsU("PinY").ResultIU = y
            return shape

        shape = self._guard(f"importing image {os.path.basename(path)!r}", _import)
        return {**self._shape_info(shape), "imported_from": path}

    def update_shape(
        self,
        shape_id: int,
        page_name: Optional[str] = None,
        text: Optional[str] = None,
        x: Optional[float] = None,
        y: Optional[float] = None,
        width_in: Optional[float] = None,
        height_in: Optional[float] = None,
    ) -> dict:
        page = self._page(page_name)
        shape = self._shape_by_id(page, shape_id)

        def _apply():
            if text is not None:
                shape.Text = text
            if x is not None:
                shape.CellsU("PinX").ResultIU = x
            if y is not None:
                shape.CellsU("PinY").ResultIU = y
            if width_in is not None:
                shape.CellsU("Width").ResultIU = width_in
            if height_in is not None:
                shape.CellsU("Height").ResultIU = height_in

        self._guard(f"updating shape {shape_id}", _apply)
        return self._shape_info(shape)

    def style_shape(
        self,
        shape_id: int,
        page_name: Optional[str] = None,
        fill_color: Optional[str] = None,
        line_color: Optional[str] = None,
        text_color: Optional[str] = None,
        line_weight_pt: Optional[float] = None,
        font_size_pt: Optional[float] = None,
        bold: Optional[bool] = None,
        line_pattern: Optional[str] = None,
    ) -> dict:
        if line_pattern is not None and line_pattern not in C.LINE_PATTERNS:
            raise VisioMcpError(
                f"line_pattern must be one of {sorted(C.LINE_PATTERNS)}, got {line_pattern!r}"
            )
        page = self._page(page_name)
        shape = self._shape_by_id(page, shape_id)

        def _apply():
            # FormulaForceU writes through cell guards — Visio's built-in
            # container/theme masters guard LineColor/Char.* cells, and plain
            # FormulaU raises 'Cell is guarded' on them
            if line_pattern is not None:
                shape.CellsU("LinePattern").FormulaForceU = str(C.LINE_PATTERNS[line_pattern])
            if fill_color is not None:
                shape.CellsU("FillForegnd").FormulaForceU = hex_to_rgb_formula(fill_color)
            if line_color is not None:
                shape.CellsU("LineColor").FormulaForceU = hex_to_rgb_formula(line_color)
            if text_color is not None:
                shape.CellsU("Char.Color").FormulaForceU = hex_to_rgb_formula(text_color)
            if line_weight_pt is not None:
                shape.CellsU("LineWeight").FormulaForceU = f"{line_weight_pt} pt"
            if font_size_pt is not None:
                shape.CellsU("Char.Size").FormulaForceU = f"{font_size_pt} pt"
            if bold is not None:
                current = int(shape.CellsU("Char.Style").ResultIU)
                shape.CellsU("Char.Style").FormulaForceU = str(
                    (current | 1) if bold else (current & ~1)
                )

        self._guard(f"styling shape {shape_id}", _apply)
        return {"shape_id": int(shape.ID), "styled": True}

    def delete_shapes(self, shape_ids: list[int], page_name: Optional[str] = None) -> dict:
        page = self._page(page_name)
        for sid in shape_ids:
            shape = self._shape_by_id(page, sid)
            self._guard(f"deleting shape {sid}", shape.Delete)
        return {"deleted": len(shape_ids)}

    # -- connectors -----------------------------------------------------------

    @staticmethod
    def _fmt_fraction(value: float) -> str:
        """0.5 -> '0.5', 0.25 -> '0.25' — formula-friendly, no trailing zeros."""
        text = f"{value:.4f}".rstrip("0").rstrip(".")
        return text or "0"

    def _edge_point_formulas(self, side: str, pos: float) -> tuple[str, str]:
        """(X formula, Y formula) for the point at `pos` along `side`, in the
        shape's local coordinates (0..Width, 0..Height)."""
        p = self._fmt_fraction(pos)
        if side == "left":
            return "0", f"Height*{p}"
        if side == "right":
            return "Width", f"Height*{p}"
        if side == "bottom":
            return f"Width*{p}", "0"
        return f"Width*{p}", "Height"  # top

    @staticmethod
    def _side_tolerance(length: float) -> float:
        """Allowance for a point being considered on an edge, scaled to the
        edge's own length (see C.SIDE_POINT_TOLERANCE_FRAC)."""
        return max(C.SIDE_POINT_TOLERANCE_MIN_IN,
                   abs(length) * C.SIDE_POINT_TOLERANCE_FRAC)

    @staticmethod
    def _edge_fraction(side: str, x: float, y: float,
                       width: float, height: float) -> Optional[float]:
        """Position (0..1) of a local point on `side`, or None if it is not on
        that edge. Mirrors _edge_point_formulas in reverse.

        The perpendicular distance is compared against a fraction of the
        dimension being crossed (width for left/right, height for
        bottom/top), so a point near a corner on a large shape is not claimed
        by the adjacent edge.
        """
        if side in ("left", "right"):
            edge_x = 0.0 if side == "left" else width
            if not height or abs(x - edge_x) > VisioClient._side_tolerance(width):
                return None
            return y / height
        edge_y = 0.0 if side == "bottom" else height
        if not width or abs(y - edge_y) > VisioClient._side_tolerance(height):
            return None
        return x / width

    @staticmethod
    def _side_of_point(x: float, y: float, width: float,
                       height: float) -> tuple[Optional[str], Optional[float]]:
        """(side, pos) for a connection point's local coordinates.

        Corner points match two edges; the first in bottom/top/left/right
        order wins (a corner is reported as bottom/top). Near-edge matching
        uses a per-dimension relative tolerance, so a point on one edge is
        not misreported as lying on the perpendicular edge merely because it
        is close to a corner in absolute inches.
        """
        candidates: list[tuple[str, float]] = []
        if width:
            tol_y = VisioClient._side_tolerance(height)
            if abs(y) <= tol_y:
                candidates.append(("bottom", x / width))
            if abs(y - height) <= tol_y:
                candidates.append(("top", x / width))
        if height:
            tol_x = VisioClient._side_tolerance(width)
            if abs(x) <= tol_x:
                candidates.append(("left", y / height))
            if abs(x - width) <= tol_x:
                candidates.append(("right", y / height))
        return candidates[0] if candidates else (None, None)

    @staticmethod
    def _free_edge_pos(taken: list[float], edge_len: float,
                       side: str, shape_id: int) -> float:
        """Widest well-spaced free slot on a 0..1 edge, inset from the corners.

        With nothing taken this is 0.5 — the round-1 midpoint behaviour.
        Further slots step away from the midpoint by at least
        EDGE_MIN_SLOT_GAP_IN physically, alternating below/above, so short
        edges hold few points instead of packing them together. Raises when
        every candidate is closer than that gap to an occupied position.
        """
        if edge_len <= 0:
            return 0.5
        min_gap = C.EDGE_MIN_SLOT_GAP_IN / edge_len
        lo, hi = C.EDGE_POS_MIN, C.EDGE_POS_MAX
        usable = min(0.5 - lo, hi - 0.5)
        steps = int(usable / min_gap + 1e-9) if min_gap > 0 else 0
        candidates = [0.5]
        for step in range(1, steps + 1):
            # unrounded: rounding before the gap check can shave a candidate
            # just under the minimum; the returned value is rounded below
            candidates.append(0.5 - step * min_gap)
            candidates.append(0.5 + step * min_gap)
        points = [t for t in taken if 0.0 <= t <= 1.0]
        tolerance = min(1e-4, min_gap / 2.0)  # absorb rounding, never the gap
        for candidate in candidates:
            if all(abs(candidate - point) >= min_gap - tolerance for point in points):
                return round(candidate, 4)
        capacity = 1 + 2 * steps
        raise VisioMcpError(
            f"The {side} edge of shape {shape_id} is full: at most {capacity} "
            f"auto-slotted connector point(s) fit with "
            f"{C.EDGE_MIN_SLOT_GAP_IN} in spacing ({edge_len:.2f} in edge). "
            "Use another side, an explicit from_pos/to_pos, or 'auto' "
            "dynamic glue."
        )

    @staticmethod
    def _connect_point_row(connect) -> Optional[int]:
        """Zero-based Connection Points row glued by a Connect row, or None
        for dynamic glue (PinX/PinY) and unreadable targets.

        ToPart >= 100 identifies a connection point (row = ToPart - 100);
        the ToCell name ('Connections.X3') is the fallback.
        """
        try:
            to_part = int(connect.ToPart)
        except Exception:
            to_part = None
        if to_part is not None and to_part >= 100:
            return to_part - 100
        try:
            match = re.match(r"Connections\.X(\d+)$", str(connect.ToCell.Name))
        except Exception:
            return None
        return int(match.group(1)) - 1 if match else None

    def _used_edge_positions(self, page, shape, side: str) -> list[float]:
        """Fractions of `side` already occupied by statically glued connector
        ends. Dynamic glue does not occupy a fixed point and is ignored."""
        width = float(shape.CellsU("Width").ResultIU)
        height = float(shape.CellsU("Height").ResultIU)
        section = C.VIS_SECTION_CONNECTION_PTS
        used: set[float] = set()
        for other in _items(page.Shapes):
            try:
                if not int(other.OneD):
                    continue
                connects = list(_items(other.Connects))
            except Exception:
                continue
            for connect in connects:
                try:
                    if int(connect.ToSheet.ID) != int(shape.ID):
                        continue
                except Exception:
                    continue
                row = self._connect_point_row(connect)
                if row is None:
                    continue
                try:
                    x = float(shape.CellsSRC(section, row, C.VIS_CNNCT_X).ResultIU)
                    y = float(shape.CellsSRC(section, row, C.VIS_CNNCT_Y).ResultIU)
                except Exception:
                    continue
                frac = self._edge_fraction(side, x, y, width, height)
                if frac is not None and 0.0 <= frac <= 1.0:
                    used.add(round(frac, 4))
        return sorted(used)

    def _resolve_edge_pos(self, page, shape, side: str, pos: Optional[float],
                          extra_taken: tuple[float, ...] = ()) -> Optional[float]:
        """Position to use for one endpoint: None for dynamic glue, the given
        fraction when explicit, otherwise the widest free slot on that side."""
        if side == "auto":
            return None
        if pos is not None:
            return float(pos)
        width = float(shape.CellsU("Width").ResultIU)
        height = float(shape.CellsU("Height").ResultIU)
        edge_len = height if side in ("left", "right") else width
        taken = self._used_edge_positions(page, shape, side) + list(extra_taken)
        return self._free_edge_pos(taken, edge_len, side, int(shape.ID))

    def _edge_connection_cell(self, shape, side: str, pos: float):
        """(Connection Points cell, actual position) at `pos` along `side`.

        An existing point within EDGE_POS_MATCH_TOLERANCE of `pos` is reused;
        the returned position is read back from that point, so callers report
        where the endpoint really is, not where they asked. Matching is
        fraction-based (relative to the edge), never an absolute inch
        tolerance — on a small icon 0.01 in would merge distinct positions.
        The point is created from Width/Height when the edge has none.

        Explicit sides pin the connector with static glue to this point —
        dynamic glue (side='auto') lets Visio choose sides from the geometry
        and cannot be directed through Cell.GlueTo. Row order in the
        Connection Points section differs between masters (BASFLO's Process
        stores left/right/bottom/top, Decision left/right/top/bottom and
        Start/End bottom/top/left/right), so rows are matched by their local
        X/Y coordinates, never by index.
        """
        width = float(shape.CellsU("Width").ResultIU)
        height = float(shape.CellsU("Height").ResultIU)
        pos = round(float(pos), 4)  # keep formula and duplicate check in sync

        section = C.VIS_SECTION_CONNECTION_PTS
        for row in range(int(shape.RowCount(section))):
            x = float(shape.CellsSRC(section, row, C.VIS_CNNCT_X).ResultIU)
            y = float(shape.CellsSRC(section, row, C.VIS_CNNCT_Y).ResultIU)
            frac = self._edge_fraction(side, x, y, width, height)
            if frac is not None and abs(frac - pos) <= C.EDGE_POS_MATCH_TOLERANCE:
                return shape.CellsSRC(section, row, C.VIS_CNNCT_X), round(frac, 4)

        x_formula, y_formula = self._edge_point_formulas(side, pos)
        row = int(shape.AddRow(section, C.VIS_ROW_LAST, C.VIS_TAG_DEFAULT))
        x_cell = shape.CellsSRC(section, row, C.VIS_CNNCT_X)
        shape.CellsSRC(section, row, C.VIS_CNNCT_Y).FormulaU = y_formula
        x_cell.FormulaU = x_formula
        return x_cell, pos

    def _glue_target(self, shape, side: str, pos: Optional[float]):
        """(Cell, actual position) a connector endpoint is glued to: PinX and
        None for dynamic glue ('auto'), a point on the side otherwise."""
        if side == "auto":
            return shape.CellsU("PinX"), None
        return self._edge_connection_cell(shape, side, 0.5 if pos is None else pos)

    def _connect_endpoint_info(self, connect) -> dict:
        """Glue details for one connector endpoint (a Connect row).

        Reports whether the end is dynamically glued ('dynamic') or glued to
        a connection point ('static'), and for static glue the edge, the
        position along it (0..1) and the 1-based connection-point row.
        """
        info = {"glue": None, "side": None, "pos": None, "point": None}
        row = self._connect_point_row(connect)
        if row is None:
            try:
                cell_name = str(connect.ToCell.Name)
            except Exception:
                cell_name = ""
            if cell_name in ("PinX", "PinY"):
                info["glue"] = "dynamic"
            return info
        info["glue"] = "static"
        info["point"] = row + 1
        try:
            target = connect.ToSheet
            width = float(target.CellsU("Width").ResultIU)
            height = float(target.CellsU("Height").ResultIU)
            section = C.VIS_SECTION_CONNECTION_PTS
            x = float(target.CellsSRC(section, row, C.VIS_CNNCT_X).ResultIU)
            y = float(target.CellsSRC(section, row, C.VIS_CNNCT_Y).ResultIU)
        except Exception:
            return info
        side, pos = self._side_of_point(x, y, width, height)
        info["side"] = side
        info["pos"] = round(pos, 4) if pos is not None else None
        return info

    # -- route control: geometry, diagnostics, manual routes ------------------

    @staticmethod
    def _shape_bbox(shape) -> tuple[float, float, float, float]:
        """(left, bottom, right, top) in page inches."""
        px = float(shape.CellsU("PinX").ResultIU)
        py = float(shape.CellsU("PinY").ResultIU)
        lx = float(shape.CellsU("LocPinX").ResultIU)
        ly = float(shape.CellsU("LocPinY").ResultIU)
        w = float(shape.CellsU("Width").ResultIU)
        h = float(shape.CellsU("Height").ResultIU)
        return (px - lx, py - ly, px - lx + w, py - ly + h)

    def _connector_polyline(self, shape) -> tuple[list[tuple[float, float]], bool]:
        """Routed polyline in page inches, or ([Begin, End], False).

        Geometry rows are local coordinates; ``Shape.XYToPage`` applies the
        shape's angle/flip, so dynamic connectors (angle 0) and rotated
        manual 1-D routes are both handled. The geometry is accepted only
        when its first/last vertex lands on Begin/End.
        """
        try:
            begin = (float(shape.CellsU("BeginX").ResultIU),
                     float(shape.CellsU("BeginY").ResultIU))
            end = (float(shape.CellsU("EndX").ResultIU),
                   float(shape.CellsU("EndY").ResultIU))
        except Exception:
            return [], False
        points: list[tuple[float, float]] = []
        try:
            section = C.VIS_SECTION_FIRST_COMPONENT
            for row in range(1, int(shape.RowCount(section))):
                lx = float(shape.CellsSRC(section, row, C.VIS_GEOM_X).ResultIU)
                ly = float(shape.CellsSRC(section, row, C.VIS_GEOM_Y).ResultIU)
                px, py = shape.XYToPage(lx, ly)
                points.append((float(px), float(py)))
        except Exception:
            return [begin, end], False
        if (len(points) < 2
                or routing.distance(points[0], begin) > 0.05
                or routing.distance(points[-1], end) > 0.05):
            return [begin, end], False
        return points, True

    def _connector_label_rect(self, shape) -> Optional[tuple[float, float, float, float]]:
        """Approximate page-space text-block rectangle, or None with no text."""
        try:
            if not _clean_text(shape.Text).strip():
                return None
            tx = float(shape.CellsU("TxtPinX").ResultIU)
            ty = float(shape.CellsU("TxtPinY").ResultIU)
            lx = float(shape.CellsU("TxtLocPinX").ResultIU)
            ly = float(shape.CellsU("TxtLocPinY").ResultIU)
            tw = float(shape.CellsU("TxtWidth").ResultIU)
            th = float(shape.CellsU("TxtHeight").ResultIU)
            cx, cy = shape.XYToPage(tx - lx + tw / 2.0, ty - ly + th / 2.0)
            return (float(cx) - tw / 2.0, float(cy) - th / 2.0,
                    float(cx) + tw / 2.0, float(cy) + th / 2.0)
        except Exception:
            return None

    @staticmethod
    def _is_manual_connector(shape) -> bool:
        try:
            return bool(int(shape.CellsU(f"User.{C.ROUTE_USER_MARKER}").ResultIU))
        except Exception:
            return False

    @staticmethod
    def _set_user_cell(shape, name: str, value: str) -> None:
        """Write one User row, adding it when the shape does not have it."""
        try:
            shape.AddNamedRow(C.VIS_SECTION_USER, name, 0)
        except Exception:
            pass  # row already exists on a rebuild
        shape.CellsU(f"User.{name}").FormulaU = str(value)

    @staticmethod
    def _get_user_cell(shape, name: str, default: Optional[float] = None):
        try:
            return float(shape.CellsU(f"User.{name}").ResultIU)
        except Exception:
            return default

    @classmethod
    def _mark_manual_connector(cls, shape) -> None:
        try:
            cls._set_user_cell(shape, C.ROUTE_USER_MARKER, "1")
        except Exception:
            pass  # the route still works; only diagnostics grouping is affected

    @staticmethod
    def _cell_formula(cell) -> str:
        """Formula text of a cell, falling back to its value.

        Real Visio always has a formula; the fallback keeps the duck-typed
        test fake (which only records ResultIU for some writes) faithful.
        """
        formula = str(cell.FormulaU).strip()
        if formula:
            return formula
        return f"{float(cell.ResultIU):g}"

    # Critical cells: if any cannot be copied the conversion must abort and
    # roll back, never delete a connector whose style would be lost.
    _STYLE_REQUIRED = ("LinePattern", "LineWeight", "LineColor",
                       "BeginArrow", "EndArrow")
    _STYLE_OPTIONAL = ("LineColorTrans", "LineCap", "Rounding",
                       "BeginArrowSize", "EndArrowSize")

    def _copy_connector_style(self, source, target) -> dict:
        """Copy native style and text between connector shapes.

        Raises VisioMcpError when a critical cell or the text cannot be
        copied. Callers run inside an undo scope and delete the source only
        after this returns, so a failure leaves the original connector in
        place instead of silently dropping its style.
        """
        copied: dict[str, str] = {}
        for name in self._STYLE_REQUIRED:
            try:
                formula = self._cell_formula(source.CellsU(name))
                target.CellsU(name).FormulaU = formula
                copied[name] = formula
            except Exception as exc:
                raise VisioMcpError(
                    f"could not preserve connector style cell {name} "
                    f"({exc}); the original connector was left unchanged"
                ) from exc
        for name in self._STYLE_OPTIONAL:
            try:
                formula = self._cell_formula(source.CellsU(name))
                target.CellsU(name).FormulaU = formula
                copied[name] = formula
            except Exception:
                continue
        try:
            text = str(source.Text)
            target.Text = text
            copied["Text"] = text
        except Exception as exc:
            raise VisioMcpError(
                f"could not preserve the connector text ({exc}); the "
                "original connector was left unchanged"
            ) from exc
        return copied

    @staticmethod
    def _connector_glue_targets(conn) -> dict[str, Any]:
        """The cells a connector's BeginX/EndX are glued to, by end name."""
        targets: dict[str, Any] = {}
        try:
            for connect in _items(conn.Connects):
                name = str(connect.FromCell.Name)
                if name.startswith("Begin"):
                    targets["begin"] = connect.ToCell
                elif name.startswith("End"):
                    targets["end"] = connect.ToCell
        except Exception:
            pass
        return targets

    def _connector_ends(self, conn) -> tuple[Optional[int], Optional[int],
                                             dict, dict]:
        """(from_id, to_id, from_end, to_end) for a connector's glue rows."""
        from_id = to_id = None
        from_end: dict = {}
        to_end: dict = {}
        try:
            for connect in _items(conn.Connects):
                name = str(connect.FromCell.Name)
                if name.startswith("Begin"):
                    from_id = int(connect.ToSheet.ID)
                    from_end = self._connect_endpoint_info(connect)
                elif name.startswith("End"):
                    to_id = int(connect.ToSheet.ID)
                    to_end = self._connect_endpoint_info(connect)
        except Exception:
            pass
        return from_id, to_id, from_end, to_end

    @staticmethod
    def _manual_local_formulas(point: tuple[float, float]) -> tuple[str, str]:
        """Page-anchored local X/Y formulas for a 1-D line geometry vertex.

        The line's local X axis runs Begin -> End, so a page point P maps to
        localX = dot(P-Begin, End-Begin)/Width and localY = cross(End-Begin,
        P-Begin)/Width. Written as formulas, the vertex stays at the page
        point while the glued endpoints move and across save/reopen.
        """
        x, y = point
        fx = (f"(({x:.6f} in-BeginX)*(EndX-BeginX)"
              f"+({y:.6f} in-BeginY)*(EndY-BeginY))/Width")
        fy = (f"((EndX-BeginX)*({y:.6f} in-BeginY)"
              f"-(EndY-BeginY)*({x:.6f} in-BeginX))/Width")
        return fx, fy

    def _reset_manual_geometry(self, shape) -> None:
        """Leave only the flag row, the MoveTo vertex and the end vertex."""
        section = C.VIS_SECTION_FIRST_COMPONENT
        while int(shape.RowCount(section)) > 3:
            shape.DeleteRow(section, int(shape.RowCount(section)) - 1)
        if int(shape.RowCount(section)) < 3:
            shape.AddRow(section, C.VIS_ROW_LAST, C.VIS_TAG_LINE_TO)

    def _write_manual_geometry(self, shape, points: list[tuple[float, float]]) -> None:
        """Rewrite a 1-D line's vertices to the page-space route.

        Row 1 is the glued begin vertex, the last row the glued end vertex
        (X=Width, Y=0); every interior row is a page-anchored formula.
        """
        if len(points) < 2:
            raise VisioMcpError("a manual route needs at least two points")
        section = C.VIS_SECTION_FIRST_COMPONENT
        self._reset_manual_geometry(shape)
        interior = points[1:-1]
        for _ in interior:
            shape.AddRow(section, C.VIS_ROW_LAST, C.VIS_TAG_LINE_TO)
        rows = 2 + len(interior)  # 2 = MoveTo + end
        for index, point in enumerate(interior):
            row = 2 + index
            fx, fy = self._manual_local_formulas(point)
            shape.CellsSRC(section, row, C.VIS_GEOM_X).FormulaU = fx
            shape.CellsSRC(section, row, C.VIS_GEOM_Y).FormulaU = fy
        shape.CellsSRC(section, rows, C.VIS_GEOM_X).FormulaU = "Width"
        shape.CellsSRC(section, rows, C.VIS_GEOM_Y).FormulaU = "0"

    def _label_page_point(self, points: list[tuple[float, float]], label_pos: float,
                          label_offset_in: float) -> tuple[float, float]:
        point, direction = routing.point_at_fraction(points, label_pos)
        # positive offset = left of travel
        nx, ny = -direction[1], direction[0]
        return (point[0] + nx * label_offset_in, point[1] + ny * label_offset_in)

    def _set_label_on_route(self, shape, label_pos: float,
                            label_offset_in: float, manual: bool) -> None:
        """Place the native text block along the routed path.

        Uses Visio's POINTALONGPATH(Geometry1.Path, travel, offset): travel
        is 0..1 along the route, offset is perpendicular inches (positive =
        left of travel). The label follows later reroutes and persists;
        ResultIU writes do not (SETATREF would recalculate them).
        """
        pos = max(0.0, min(1.0, float(label_pos)))
        offset = float(label_offset_in)
        expression = (f"POINTALONGPATH(Geometry1.Path,{pos:.6f},"
                      f"{offset:.6f} in)")
        try:
            shape.CellsU("TxtPinX").FormulaForceU = f"PNTX({expression})"
            shape.CellsU("TxtPinY").FormulaForceU = f"PNTY({expression})"
        except Exception:
            # fallback: anchor the label to an absolute page point
            points, _ = self._connector_polyline(shape)
            self._set_label_page_point(
                shape, self._label_page_point(points, pos, offset), manual
            )
            return
        if manual:
            shape.CellsU("TxtAngle").FormulaU = "-Angle"
            shape.CellsU("TxtWidth").FormulaU = "MAX(TEXTWIDTH(TheText),5*Char.Size)"
            shape.CellsU("TxtHeight").FormulaU = "TEXTHEIGHT(TheText,TxtWidth)"
            shape.CellsU("TxtLocPinX").FormulaU = "TxtWidth*0.5"
            shape.CellsU("TxtLocPinY").FormulaU = "TxtHeight*0.5"
        else:
            shape.CellsU("TxtAngle").FormulaU = "0 deg"

    def _set_label_page_point(self, shape, point: tuple[float, float],
                              manual: bool) -> None:
        """Pin a connector's native text block to an absolute page point."""
        x, y = point
        if manual:
            fx, fy = self._manual_local_formulas(point)
            shape.CellsU("TxtPinX").FormulaU = fx
            shape.CellsU("TxtPinY").FormulaU = fy
            # the line itself is rotated; counter-rotate the text block
            shape.CellsU("TxtAngle").FormulaU = "-Angle"
            # size the block to the text, not to the whole line
            shape.CellsU("TxtWidth").FormulaU = "MAX(TEXTWIDTH(TheText),5*Char.Size)"
            shape.CellsU("TxtHeight").FormulaU = "TEXTHEIGHT(TheText,TxtWidth)"
            shape.CellsU("TxtLocPinX").FormulaU = "TxtWidth*0.5"
            shape.CellsU("TxtLocPinY").FormulaU = "TxtHeight*0.5"
        else:
            # dynamic connectors are not rotated: local = page - (Pin-LocPin)
            shape.CellsU("TxtPinX").FormulaForceU = (
                f"{x:.6f} in - PinX + LocPinX"
            )
            shape.CellsU("TxtPinY").FormulaForceU = (
                f"{y:.6f} in - PinY + LocPinY"
            )
            shape.CellsU("TxtAngle").FormulaU = "0 deg"

    @staticmethod
    def _validate_waypoints(waypoints) -> Optional[list[tuple[float, float]]]:
        if waypoints is None:
            return None
        if not isinstance(waypoints, (list, tuple)):
            raise VisioMcpError("waypoints must be a list of [x, y] page points")
        if len(waypoints) > C.ROUTE_MAX_WAYPOINTS:
            raise VisioMcpError(
                f"waypoints accepts at most {C.ROUTE_MAX_WAYPOINTS} points, "
                f"got {len(waypoints)}"
            )
        parsed: list[tuple[float, float]] = []
        for index, point in enumerate(waypoints):
            if not isinstance(point, (list, tuple)) or len(point) != 2:
                raise VisioMcpError(
                    f"waypoints[{index}] must be [x, y] in inches, got {point!r}"
                )
            try:
                x, y = float(point[0]), float(point[1])
            except (TypeError, ValueError):
                raise VisioMcpError(
                    f"waypoints[{index}] must contain numbers, got {point!r}"
                ) from None
            if not (abs(x) < 1e7 and abs(y) < 1e7):
                raise VisioMcpError(f"waypoints[{index}] is not a finite point")
            parsed.append((x, y))
        return parsed

    @staticmethod
    def _validate_label_args(label_pos, label_offset_in, has_text: bool) -> None:
        if label_pos is None and label_offset_in is None:
            return
        if not has_text:
            raise VisioMcpError(
                "label_pos/label_offset_in require a label (text) on the connector"
            )
        if label_pos is not None and not 0.0 <= float(label_pos) <= 1.0:
            raise VisioMcpError(
                f"label_pos must be between 0 and 1, got {label_pos!r}"
            )
        if label_offset_in is not None:
            try:
                float(label_offset_in)
            except (TypeError, ValueError):
                raise VisioMcpError(
                    f"label_offset_in must be a number, got {label_offset_in!r}"
                ) from None

    def _create_manual_connector(self, page, shape_from, shape_to, from_side, to_side,
                                 resolved_from, resolved_to, waypoints, label,
                                 end_arrow, begin_arrow, line_pattern,
                                 line_weight_pt, line_color, label_pos,
                                 label_offset_in, actual: dict,
                                 route_style: int = C.ROUTE_RIGHT_ANGLE,
                                 route_ext: int = C.ROUTE_EXT_STRAIGHT) -> Any:
        """Native 1-D line with a page-anchored polyline; endpoints glued."""
        begin = routing.side_point(self._shape_bbox(shape_from), from_side,
                                   resolved_from)
        end = routing.side_point(self._shape_bbox(shape_to), to_side, resolved_to)
        conn = page.DrawLine(begin[0], begin[1], end[0], end[1])
        begin_cell, actual["from"] = self._glue_target(shape_from, from_side,
                                                       resolved_from)
        end_cell, actual["to"] = self._glue_target(shape_to, to_side, resolved_to)
        conn.CellsU("BeginX").GlueTo(begin_cell)
        conn.CellsU("EndX").GlueTo(end_cell)
        actual_begin = (float(conn.CellsU("BeginX").ResultIU),
                        float(conn.CellsU("BeginY").ResultIU))
        actual_end = (float(conn.CellsU("EndX").ResultIU),
                      float(conn.CellsU("EndY").ResultIU))
        points = routing.orthogonal_route(
            actual_begin, actual_end, from_side, to_side, waypoints or (),
            C.ROUTE_MANUAL_STUB_IN,
        )
        self._write_manual_geometry(conn, points)
        self._mark_manual_connector(conn)
        # Remember the routing behaviour the route replaced, so
        # set_connector_route(..., restore_auto=True) can give it back.
        try:
            self._set_user_cell(conn, C.ROUTE_USER_STYLE, str(int(route_style)))
            self._set_user_cell(conn, C.ROUTE_USER_EXT, str(int(route_ext)))
        except Exception:
            pass
        conn.CellsU("EndArrow").FormulaU = str(
            C.ARROW_FILLED_TRIANGLE if end_arrow else C.ARROW_NONE
        )
        conn.CellsU("BeginArrow").FormulaU = str(
            C.ARROW_FILLED_TRIANGLE if begin_arrow else C.ARROW_NONE
        )
        conn.CellsU("LinePattern").ResultIU = C.LINE_PATTERNS[line_pattern]
        if line_weight_pt is not None:
            conn.CellsU("LineWeight").FormulaU = f"{line_weight_pt} pt"
        if line_color is not None:
            conn.CellsU("LineColor").FormulaU = hex_to_rgb_formula(line_color)
        if label:
            conn.Text = label
            pos = 0.5 if label_pos is None else float(label_pos)
            offset = 0.0 if label_offset_in is None else float(label_offset_in)
            self._set_label_on_route(conn, pos, offset, manual=True)
        return conn

    def route_diagnostics(self, page_name: Optional[str] = None,
                          connector_id: Optional[int] = None) -> dict:
        """Read-only routing quality report for one page.

        For every connector: page polyline, length, detour ratio, all
        crossings, every actual collinear-overlap interval, node traversal,
        label conflicts, plus (for manual routes) orthogonality and whether
        a node move left the page-anchored path needing readjustment.
        ``connector_id`` restricts the report to one connector. No mutation.
        """
        page = self._page(page_name)
        connectors: dict[int, dict] = {}
        node_rects: dict[int, tuple[float, float, float, float]] = {}
        label_rects: dict[int, tuple[float, float, float, float]] = {}
        for shape in _items(page.Shapes):
            sid = int(shape.ID)
            try:
                is_connector = bool(int(shape.OneD))
            except Exception:
                is_connector = False
            if is_connector:
                points, geometry_ok = self._connector_polyline(shape)
                from_id = to_id = None
                try:
                    for connect in _items(shape.Connects):
                        name = str(connect.FromCell.Name)
                        if name.startswith("Begin"):
                            from_id = int(connect.ToSheet.ID)
                        elif name.startswith("End"):
                            to_id = int(connect.ToSheet.ID)
                except Exception:
                    pass
                connectors[sid] = {
                    "from_id": from_id,
                    "to_id": to_id,
                    "points": points,
                    "geometry_ok": geometry_ok,
                    "route_kind": "manual" if self._is_manual_connector(shape) else "auto",
                }
                label = self._connector_label_rect(shape)
                if label is not None:
                    label_rects[sid] = label
            else:
                try:
                    if shape.ContainerProperties is not None:
                        continue  # a container is a zone, not an obstacle
                except Exception:
                    pass
                node_rects[sid] = self._shape_bbox(shape)
        report = routing.diagnose_routes(
            connectors, node_rects, label_rects,
            endpoint_eps=C.ROUTE_ENDPOINT_EPS_IN,
            overlap_min=C.ROUTE_OVERLAP_MIN_IN,
            inside_min=C.ROUTE_INSIDE_MIN_IN,
            node_inset=C.ROUTE_NODE_INSET_IN,
            label_area_min=C.ROUTE_LABEL_AREA_MIN_IN2,
            detour_ratio=C.ROUTE_DETOUR_RATIO,
            ortho_tol=C.ROUTE_ORTHO_TOLERANCE_IN,
        )
        if connector_id is not None:
            if connector_id not in connectors:
                available = sorted(connectors)
                raise VisioMcpError(
                    f"Shape {connector_id} is not a connector on page "
                    f"{page.Name!r}; connectors: {available}"
                )
            report = routing.filter_report(report, connector_id)
        items = []
        for sid, info in connectors.items():
            if connector_id is not None and sid != connector_id:
                continue
            data = report["connectors"][sid]
            items.append({
                "connector_id": sid,
                "from_id": info["from_id"],
                "to_id": info["to_id"],
                "route_kind": info["route_kind"],
                "geometry_ok": info["geometry_ok"],
                "orthogonal": data["orthogonal"],
                "needs_readjust": data["needs_readjust"],
                "non_orthogonal_segments": data["non_orthogonal_segments"],
                "readjust_hint": data["readjust_hint"],
                "polyline": [[round(x, 3), round(y, 3)] for x, y in info["points"]],
                "route_length": data["route_length"],
                "straight_length": data["straight_length"],
                "detour_ratio": data["detour_ratio"],
                "crossings": data["crossings"],
                "overlaps": data["overlaps"],
                "through_shapes": data["through_shapes"],
                "label_rect": ([round(v, 3) for v in label_rects[sid]]
                               if sid in label_rects else None),
                "label_conflicts": data["label_conflicts"],
            })
        sheet = page.PageSheet
        return {
            "page": str(page.Name),
            "page_size_in": [
                round(float(sheet.CellsU("PageWidth").ResultIU), 3),
                round(float(sheet.CellsU("PageHeight").ResultIU), 3),
            ],
            "connectors": items,
            "summary": report["summary"],
        }

    def set_route_options(self, separate_lanes: bool = True,
                          page_name: Optional[str] = None) -> dict:
        """Set page-level native connector lane separation.

        LineAdjustFrom=1 (separate all lines) and LineAdjustTo=2 (never
        overlap lines) make Visio's router keep distinct connectors out of
        one lane; this removes short merge trunks (0.15 in) and long
        co-linear overlaps without changing the connector API. The cells are
        page layout settings and persist in the saved document.
        """
        app = self._app()
        page = self._page(page_name)
        sheet = page.PageSheet

        def apply():
            sheet.CellsU("LineAdjustFrom").FormulaU = "1" if separate_lanes else "0"
            sheet.CellsU("LineAdjustTo").FormulaU = "2" if separate_lanes else "0"
            return {
                "line_adjust_from": int(float(
                    sheet.CellsU("LineAdjustFrom").ResultIU)),
                "line_adjust_to": int(float(
                    sheet.CellsU("LineAdjustTo").ResultIU)),
            }

        scope = self._guard("starting undo scope",
                            lambda: app.BeginUndoScope("Set route options"))
        try:
            values = self._guard("setting page route options", apply)
            self._guard("closing undo scope",
                        lambda: app.EndUndoScope(scope, True))
        except Exception:
            try:
                app.EndUndoScope(scope, False)
            except Exception:
                pass
            raise
        return {
            "page": str(page.Name),
            "separate_lanes": bool(separate_lanes),
            **values,
            "note": "Visio's router separates different connectors into "
                    "distinct lanes; manual routes keep exact corridors.",
        }

    def connect_shapes(
        self,
        from_id: int,
        to_id: int,
        label: Optional[str] = None,
        route: str = "right_angle",
        end_arrow: bool = True,
        begin_arrow: bool = False,
        page_name: Optional[str] = None,
        line_pattern: str = "solid",
        line_weight_pt: Optional[float] = None,
        line_color: Optional[str] = None,
        from_side: str = "auto",
        to_side: str = "auto",
        from_pos: Optional[float] = None,
        to_pos: Optional[float] = None,
        waypoints: Optional[list] = None,
        label_pos: Optional[float] = None,
        label_offset_in: Optional[float] = None,
    ) -> dict:
        if route not in C.CONNECTOR_ROUTES:
            raise VisioMcpError(
                f"route must be one of {sorted(C.CONNECTOR_ROUTES)}, got {route!r}"
            )
        if line_pattern not in C.LINE_PATTERNS:
            raise VisioMcpError(
                f"line_pattern must be one of {sorted(C.LINE_PATTERNS)}, got {line_pattern!r}"
            )
        for arg_name, side in (("from_side", from_side), ("to_side", to_side)):
            if side not in C.CONNECTOR_SIDES:
                raise VisioMcpError(
                    f"{arg_name} must be one of {list(C.CONNECTOR_SIDES)}, got {side!r}"
                )
        for side_name, pos_name, side, pos in (
            ("from_side", "from_pos", from_side, from_pos),
            ("to_side", "to_pos", to_side, to_pos),
        ):
            if pos is None:
                continue
            if side == "auto":
                raise VisioMcpError(
                    f"{pos_name} requires an explicit {side_name} "
                    "(top/bottom/left/right); 'auto' uses Visio's dynamic glue "
                    "and has no fixed point on an edge."
                )
            if not 0.0 <= float(pos) <= 1.0:
                raise VisioMcpError(
                    f"{pos_name} must be between 0 and 1 (a fraction along the "
                    f"edge), got {pos!r}"
                )
            if not C.EDGE_CORNER_EPS < float(pos) < 1.0 - C.EDGE_CORNER_EPS:
                raise VisioMcpError(
                    f"{pos_name}={pos!r} is a corner: a corner lies on two "
                    "edges, so the side is ambiguous. Use a position between "
                    f"{C.EDGE_CORNER_EPS} and {1.0 - C.EDGE_CORNER_EPS} "
                    "(e.g. 0.01 or 0.99), or name the adjacent side."
                )
        parsed_waypoints = self._validate_waypoints(waypoints)
        self._validate_label_args(label_pos, label_offset_in, bool(label))
        if parsed_waypoints is not None:
            if route != "right_angle":
                raise VisioMcpError(
                    "waypoints require route='right_angle' (manual routes are "
                    "orthogonal polylines), got route=" + repr(route)
                )
            if from_side == "auto" or to_side == "auto":
                raise VisioMcpError(
                    "waypoints require explicit from_side and to_side "
                    "(top/bottom/left/right): a manual route has no dynamic "
                    "glue to choose a side for it."
                )
        app = self._app()
        page = self._page(page_name)
        shape_from = self._shape_by_id(page, from_id)
        shape_to = self._shape_by_id(page, to_id)

        # Resolve crowded-edge slots before touching the page: the scan only
        # reads existing glue, so nothing needs rolling back if it fails.
        resolved_from = self._guard(
            "resolving the from edge position",
            lambda: self._resolve_edge_pos(page, shape_from, from_side, from_pos),
        )
        extra_taken: tuple[float, ...] = ()
        if (resolved_from is not None and to_side == from_side
                and int(shape_to.ID) == int(shape_from.ID)):
            extra_taken = (resolved_from,)  # self-loop: from and to share the edge
        resolved_to = self._guard(
            "resolving the to edge position",
            lambda: self._resolve_edge_pos(
                page, shape_to, to_side, to_pos, extra_taken
            ),
        )

        # Actual positions may differ from the resolved ones when an existing
        # point within tolerance is reused; report what was really glued.
        actual: dict[str, Optional[float]] = {"from": None, "to": None}

        def _connect():
            if parsed_waypoints is not None:
                return self._create_manual_connector(
                    page, shape_from, shape_to, from_side, to_side,
                    resolved_from, resolved_to, parsed_waypoints, label,
                    end_arrow, begin_arrow, line_pattern, line_weight_pt,
                    line_color, label_pos, label_offset_in, actual,
                )
            conn = page.Drop(app.ConnectorToolDataObject, 0.0, 0.0)
            # 'auto' glues to PinX = dynamic glue: Visio picks the best sides
            # and re-routes as the layout changes. A pinned side glues to that
            # side's connection point instead (static glue): the endpoint
            # follows its shape and re-routes, but keeps the chosen edge and
            # position along it.
            begin_cell, actual["from"] = self._glue_target(
                shape_from, from_side, resolved_from
            )
            end_cell, actual["to"] = self._glue_target(
                shape_to, to_side, resolved_to
            )
            conn.CellsU("BeginX").GlueTo(begin_cell)
            conn.CellsU("EndX").GlueTo(end_cell)
            route_style, route_ext = C.CONNECTOR_ROUTES[route]
            conn.CellsU("ShapeRouteStyle").ResultIU = route_style
            conn.CellsU("ConLineRouteExt").ResultIU = route_ext
            conn.CellsU("EndArrow").FormulaU = str(
                C.ARROW_FILLED_TRIANGLE if end_arrow else C.ARROW_NONE
            )
            conn.CellsU("BeginArrow").FormulaU = str(
                C.ARROW_FILLED_TRIANGLE if begin_arrow else C.ARROW_NONE
            )
            # always set: an explicit 'solid' must override themed defaults
            conn.CellsU("LinePattern").ResultIU = C.LINE_PATTERNS[line_pattern]
            if line_weight_pt is not None:
                conn.CellsU("LineWeight").FormulaU = f"{line_weight_pt} pt"
            if line_color is not None:
                conn.CellsU("LineColor").FormulaU = hex_to_rgb_formula(line_color)
            if label:
                conn.Text = label
                if label_pos is not None or label_offset_in is not None:
                    pos = 0.5 if label_pos is None else float(label_pos)
                    offset = 0.0 if label_offset_in is None else float(label_offset_in)
                    self._set_label_on_route(conn, pos, offset, manual=False)
            return conn

        # One undo scope covers the dropped connector AND any connection point
        # created for it, so a failed glue/style step leaves the page untouched.
        scope = self._guard("starting undo scope", lambda: app.BeginUndoScope("Connect shapes"))
        try:
            conn = self._guard(f"connecting shape {from_id} -> {to_id}", _connect)
            self._guard("closing undo scope", lambda: app.EndUndoScope(scope, True))
        except Exception:
            try:
                app.EndUndoScope(scope, False)  # False = roll the connector back
            except Exception:
                pass
            raise
        manual = parsed_waypoints is not None
        applied_label_pos = None
        applied_label_offset = None
        if label and (manual or label_pos is not None or label_offset_in is not None):
            applied_label_pos = 0.5 if label_pos is None else round(float(label_pos), 4)
            applied_label_offset = (0.0 if label_offset_in is None
                                    else round(float(label_offset_in), 4))
        return {
            "connector_id": int(conn.ID),
            "from_id": from_id,
            "to_id": to_id,
            "route": route,
            "route_kind": "manual" if manual else "auto",
            "from_side": from_side,
            "to_side": to_side,
            "from_pos": round(actual["from"], 4) if actual["from"] is not None else None,
            "to_pos": round(actual["to"], 4) if actual["to"] is not None else None,
            "label_pos": applied_label_pos,
            "label_offset_in": applied_label_offset,
        }

    def set_connector_route(
        self,
        connector_id: int,
        waypoints: Optional[list] = None,
        label_pos: Optional[float] = None,
        label_offset_in: Optional[float] = None,
        page_name: Optional[str] = None,
        restore_auto: bool = False,
    ) -> dict:
        """Re-route an existing connector, move only its label, or restore
        Visio's native automatic routing.

        With ``waypoints``, a dynamic connector is replaced by a native 1-D
        line carrying a page-anchored polyline (same endpoints, sides,
        positions, arrows and style); a manual route is rewritten in place.
        Without ``waypoints`` only the label position is changed.
        ``restore_auto=True`` converts a route-controlled line back into a
        dynamic connector glued to the same connection points, preserving
        text, arrows and line style; an already-automatic connector is left
        alone. Waypoints cannot be combined with ``restore_auto``.
        """
        app = self._app()
        page = self._page(page_name)
        conn = self._shape_by_id(page, connector_id)
        try:
            is_connector = bool(int(conn.OneD))
        except Exception:
            is_connector = False
        if not is_connector:
            raise VisioMcpError(
                f"Shape {connector_id} is not a connector (OneD shape)"
            )
        text = _clean_text(conn.Text)
        self._validate_label_args(label_pos, label_offset_in, bool(text))
        parsed = self._validate_waypoints(waypoints)
        if restore_auto and parsed is not None:
            raise VisioMcpError(
                "restore_auto cannot be combined with waypoints: automatic "
                "routing lets Visio choose the path. Omit waypoints, or leave "
                "restore_auto off to pin an explicit route."
            )
        if (not restore_auto and parsed is None
                and label_pos is None and label_offset_in is None):
            raise VisioMcpError(
                "set_connector_route needs waypoints, label_pos/"
                "label_offset_in, or restore_auto; nothing would change otherwise"
            )

        from_id, to_id, from_end, to_end = self._connector_ends(conn)
        if from_id is None or to_id is None:
            raise VisioMcpError(
                f"Connector {connector_id} is not glued to two shapes; "
                "get_page_state lists its endpoints."
            )
        manual = self._is_manual_connector(conn)
        old_label_rect = self._connector_label_rect(conn)

        # -- restore Visio's native automatic routing ------------------------
        if restore_auto:
            if not manual:
                if label_pos is not None or label_offset_in is not None:
                    pos = 0.5 if label_pos is None else float(label_pos)
                    offset = (0.0 if label_offset_in is None
                              else float(label_offset_in))
                    scope = self._guard(
                        "starting undo scope",
                        lambda: app.BeginUndoScope("Set connector route"))
                    try:
                        self._set_label_on_route(conn, pos, offset, manual=False)
                        self._guard("closing undo scope",
                                    lambda: app.EndUndoScope(scope, True))
                    except Exception:
                        try:
                            app.EndUndoScope(scope, False)
                        except Exception:
                            pass
                        raise
                points, _ = self._connector_polyline(conn)
                return {
                    "connector_id": int(conn.ID),
                    "replaced_id": None,
                    "route_kind": "auto",
                    "restored_auto": False,
                    "note": "connector already uses Visio's automatic routing",
                    "from_id": from_id,
                    "to_id": to_id,
                    "from_side": from_end.get("side"),
                    "to_side": to_end.get("side"),
                    "from_pos": from_end.get("pos"),
                    "to_pos": to_end.get("pos"),
                    "label_pos": (round(0.5 if label_pos is None else float(label_pos), 4)
                                  if (label_pos is not None or label_offset_in is not None)
                                  else None),
                    "label_offset_in": (round(0.0 if label_offset_in is None
                                              else float(label_offset_in), 4)
                                        if (label_pos is not None or label_offset_in is not None)
                                        else None),
                    "polyline": [[round(x, 3), round(y, 3)] for x, y in points],
                }
            targets = self._connector_glue_targets(conn)
            if "begin" not in targets or "end" not in targets:
                raise VisioMcpError(
                    f"Connector {connector_id} has no readable glue targets; "
                    "cannot restore automatic routing without losing its "
                    "topology. get_page_state lists its endpoints."
                )
            route_style = int(self._get_user_cell(
                conn, C.ROUTE_USER_STYLE, C.ROUTE_RIGHT_ANGLE))
            route_ext = int(self._get_user_cell(
                conn, C.ROUTE_USER_EXT, C.ROUTE_EXT_STRAIGHT))
            replaced_id = int(conn.ID)
            scope = self._guard(
                "starting undo scope",
                lambda: app.BeginUndoScope("Restore automatic routing"))
            try:
                new = page.Drop(app.ConnectorToolDataObject, 0.0, 0.0)
                # the same connection points = the same static glue, sides
                # and positions the manual route had
                new.CellsU("BeginX").GlueTo(targets["begin"])
                new.CellsU("EndX").GlueTo(targets["end"])
                new.CellsU("ShapeRouteStyle").ResultIU = route_style
                new.CellsU("ConLineRouteExt").ResultIU = route_ext
                self._copy_connector_style(conn, new)
                if text and (label_pos is not None or label_offset_in is not None):
                    pos = 0.5 if label_pos is None else float(label_pos)
                    offset = (0.0 if label_offset_in is None
                              else float(label_offset_in))
                    self._set_label_on_route(new, pos, offset, manual=False)
                elif text and old_label_rect is not None:
                    center = ((old_label_rect[0] + old_label_rect[2]) / 2.0,
                              (old_label_rect[1] + old_label_rect[3]) / 2.0)
                    self._set_label_page_point(new, center, manual=False)
                new_id = int(new.ID)
                conn.Delete()  # only after every copy above succeeded
                self._guard("closing undo scope",
                            lambda: app.EndUndoScope(scope, True))
            except Exception:
                try:
                    app.EndUndoScope(scope, False)
                except Exception:
                    pass
                raise
            points, _ = self._connector_polyline(new)
            new_from_id, new_to_id, new_from_end, new_to_end = \
                self._connector_ends(new)
            applied_pos = applied_offset = None
            if text and (label_pos is not None or label_offset_in is not None):
                applied_pos = round(0.5 if label_pos is None else float(label_pos), 4)
                applied_offset = round(
                    0.0 if label_offset_in is None else float(label_offset_in), 4)
            return {
                "connector_id": new_id,
                "replaced_id": replaced_id,
                "route_kind": "auto",
                "restored_auto": True,
                "from_id": new_from_id,
                "to_id": new_to_id,
                "from_side": new_from_end.get("side"),
                "to_side": new_to_end.get("side"),
                "from_pos": new_from_end.get("pos"),
                "to_pos": new_to_end.get("pos"),
                "label_pos": applied_pos,
                "label_offset_in": applied_offset,
                "polyline": [[round(x, 3), round(y, 3)] for x, y in points],
            }

        # -- label-only path -------------------------------------------------
        if parsed is None:
            points, _ = self._connector_polyline(conn)
            pos = 0.5 if label_pos is None else float(label_pos)
            offset = 0.0 if label_offset_in is None else float(label_offset_in)
            scope = self._guard("starting undo scope",
                                lambda: app.BeginUndoScope("Set connector route"))
            try:
                self._set_label_on_route(conn, pos, offset, manual=manual)
                self._guard("closing undo scope",
                            lambda: app.EndUndoScope(scope, True))
            except Exception:
                try:
                    app.EndUndoScope(scope, False)
                except Exception:
                    pass
                raise
            return {
                "connector_id": int(conn.ID),
                "replaced_id": None,
                "route_kind": "manual" if manual else "auto",
                "restored_auto": False,
                "from_id": from_id,
                "to_id": to_id,
                "from_side": from_end.get("side"),
                "to_side": to_end.get("side"),
                "from_pos": from_end.get("pos"),
                "to_pos": to_end.get("pos"),
                "label_pos": round(pos, 4),
                "label_offset_in": round(offset, 4),
                "polyline": [[round(x, 3), round(y, 3)] for x, y in points],
            }

        # -- route path ------------------------------------------------------
        from_side = from_end.get("side")
        to_side = to_end.get("side")
        if (from_end.get("glue") != "static" or to_end.get("glue") != "static"
                or from_side not in routing.SIDE_NORMALS
                or to_side not in routing.SIDE_NORMALS):
            raise VisioMcpError(
                "set_connector_route requires a connector pinned to explicit "
                "sides (from_side/to_side). This connector uses dynamic glue; "
                "connect_shapes(..., from_side=..., to_side=..., waypoints=...) "
                "creates a route-controlled connector instead."
            )
        from_pos = float(from_end.get("pos"))
        to_pos = float(to_end.get("pos"))

        end_arrow = bool(int(conn.CellsU("EndArrow").ResultIU))
        begin_arrow = bool(int(conn.CellsU("BeginArrow").ResultIU))
        try:
            route_style = int(float(conn.CellsU("ShapeRouteStyle").ResultIU))
        except Exception:
            route_style = C.ROUTE_RIGHT_ANGLE
        try:
            route_ext = int(float(conn.CellsU("ConLineRouteExt").ResultIU))
        except Exception:
            route_ext = C.ROUTE_EXT_STRAIGHT

        scope = self._guard("starting undo scope",
                            lambda: app.BeginUndoScope("Set connector route"))
        replaced_id = None
        try:
            if manual:
                begin = (float(conn.CellsU("BeginX").ResultIU),
                         float(conn.CellsU("BeginY").ResultIU))
                end = (float(conn.CellsU("EndX").ResultIU),
                       float(conn.CellsU("EndY").ResultIU))
                points = routing.orthogonal_route(
                    begin, end, from_side, to_side, parsed,
                    C.ROUTE_MANUAL_STUB_IN,
                )
                self._write_manual_geometry(conn, points)
                self._mark_manual_connector(conn)
                if text and (label_pos is not None or label_offset_in is not None):
                    pos = 0.5 if label_pos is None else float(label_pos)
                    offset = (0.0 if label_offset_in is None
                              else float(label_offset_in))
                    self._set_label_on_route(conn, pos, offset, manual=True)
            else:
                shape_from = page.Shapes.ItemFromID(from_id)
                shape_to = page.Shapes.ItemFromID(to_id)
                actual = {"from": from_pos, "to": to_pos}
                new = self._create_manual_connector(
                    page, shape_from, shape_to, from_side, to_side,
                    from_pos, to_pos, parsed, text or None,
                    end_arrow, begin_arrow, "solid", None, None,
                    label_pos, label_offset_in, actual,
                    route_style=route_style, route_ext=route_ext,
                )
                # Strict: a partial style copy raises and the undo scope
                # restores the original connector — it is never deleted
                # while some of its style could not be carried over.
                self._copy_connector_style(conn, new)
                if (text and label_pos is None and label_offset_in is None
                        and old_label_rect is not None):
                    center = ((old_label_rect[0] + old_label_rect[2]) / 2.0,
                              (old_label_rect[1] + old_label_rect[3]) / 2.0)
                    self._set_label_page_point(new, center, manual=True)
                points, _ = self._connector_polyline(new)
                replaced_id = int(conn.ID)
                conn.Delete()  # only after every copy above succeeded
                conn = new
            self._guard("closing undo scope",
                        lambda: app.EndUndoScope(scope, True))
        except Exception:
            try:
                app.EndUndoScope(scope, False)
            except Exception:
                pass
            raise

        points, _ = self._connector_polyline(conn)
        applied_pos = None
        applied_offset = None
        if text:
            if label_pos is not None or label_offset_in is not None:
                applied_pos = round(0.5 if label_pos is None else float(label_pos), 4)
                applied_offset = round(
                    0.0 if label_offset_in is None else float(label_offset_in), 4
                )
        return {
            "connector_id": int(conn.ID),
            "replaced_id": replaced_id,
            "route_kind": "manual",
            "restored_auto": False,
            "from_id": from_id,
            "to_id": to_id,
            "from_side": from_side,
            "to_side": to_side,
            "from_pos": round(from_pos, 4),
            "to_pos": round(to_pos, 4),
            "label_pos": applied_pos,
            "label_offset_in": applied_offset,
            "polyline": [[round(x, 3), round(y, 3)] for x, y in points],
        }

    # -- pages ----------------------------------------------------------------

    def pages(self, action: str, name: Optional[str] = None) -> dict:
        app = self._app()
        doc = self._drawing_doc()
        if action == "list":
            active = None
            try:
                active = str(app.ActivePage.Name) if app.ActivePage is not None else None
            except Exception:
                pass
            return {
                "pages": [
                    {"name": str(p.Name), "index": int(p.Index), "active": str(p.Name) == active}
                    for p in _items(doc.Pages)
                ]
            }
        if action == "add":
            page = self._guard("adding page", doc.Pages.Add)
            if name:
                page.Name = name
            return {"page": str(page.Name), "index": int(page.Index)}
        if action == "activate":
            if not name:
                raise VisioMcpError("pages(action='activate') requires a page name")
            page = self._page(name)
            self._guard(
                f"activating page {name!r}",
                lambda: setattr(app.ActiveWindow, "Page", page),
            )
            return {"active_page": str(page.Name)}
        raise VisioMcpError(f"Unknown pages action {action!r}; use list, add, or activate.")

    # -- page size / text / containers ------------------------------------------

    def set_page_size(
        self,
        width_in: Optional[float] = None,
        height_in: Optional[float] = None,
        orientation: Optional[str] = None,
        fit_to_contents: bool = False,
        page_name: Optional[str] = None,
    ) -> dict:
        # validate everything BEFORE mutating the document
        if orientation not in (None, "portrait", "landscape"):
            raise VisioMcpError(
                f"orientation must be 'portrait' or 'landscape', got {orientation!r}"
            )
        for name, value in (("width_in", width_in), ("height_in", height_in)):
            if value is not None and value <= 0:
                raise VisioMcpError(f"{name} must be positive, got {value}")
        page = self._page(page_name)

        def _apply():
            sheet = page.PageSheet
            if fit_to_contents:
                page.ResizeToFitContents()
            if width_in is not None:
                sheet.CellsU("PageWidth").ResultIU = width_in
            if height_in is not None:
                sheet.CellsU("PageHeight").ResultIU = height_in
            w = float(sheet.CellsU("PageWidth").ResultIU)
            h = float(sheet.CellsU("PageHeight").ResultIU)
            if orientation == "landscape" or (orientation is None and w > h):
                sheet.CellsU("PrintPageOrientation").ResultIU = C.ORIENTATION_LANDSCAPE
            else:
                sheet.CellsU("PrintPageOrientation").ResultIU = C.ORIENTATION_PORTRAIT
            return [round(w, 3), round(h, 3)]

        size = self._guard("setting page size", _apply)
        return {"page": str(page.Name), "page_size_in": size}

    def drop_text(
        self,
        text: str,
        x: float,
        y: float,
        width_in: Optional[float] = None,
        height_in: Optional[float] = None,
        font_size_pt: float = 10.0,
        bold: bool = False,
        text_color: Optional[str] = None,
        align: str = "center",
        page_name: Optional[str] = None,
    ) -> dict:
        aligns = {"left": C.ALIGN_LEFT, "center": C.ALIGN_CENTER, "right": C.ALIGN_RIGHT}
        if align not in aligns:
            raise VisioMcpError(f"align must be one of {sorted(aligns)}, got {align!r}")
        if not text.strip():
            raise VisioMcpError("text must not be empty")
        for name, value in (("width_in", width_in), ("height_in", height_in)):
            if value is not None and value <= 0:
                raise VisioMcpError(f"{name} must be positive, got {value}")
        page = self._page(page_name)
        # rough auto-size from the longest line when no explicit size given
        longest = max(len(line) for line in text.splitlines())
        w = width_in if width_in is not None else max(1.0, longest * font_size_pt * 0.009)
        h = height_in if height_in is not None else max(0.3, len(text.splitlines()) * font_size_pt * 0.02)

        def _drop():
            # Visio's standard text-only shape: a borderless, unfilled rectangle
            shape = page.DrawRectangle(x - w / 2, y - h / 2, x + w / 2, y + h / 2)
            shape.CellsU("LinePattern").ResultIU = 0
            shape.CellsU("FillPattern").ResultIU = 0
            shape.Text = text
            shape.CellsU("Char.Size").FormulaU = f"{font_size_pt} pt"
            if bold:
                cell = shape.CellsU("Char.Style")
                cell.ResultIU = int(cell.ResultIU) | 1
            if text_color is not None:
                shape.CellsU("Char.Color").FormulaU = hex_to_rgb_formula(text_color)
            shape.CellsU("Para.HorzAlign").ResultIU = aligns[align]
            return shape

        shape = self._guard("dropping text", _drop)
        return {**self._shape_info(shape), "is_text": True}

    def _container_master(self, master_name: Optional[str]):
        app = self._app()
        path = self._guard(
            "locating the built-in container stencil",
            lambda: app.GetBuiltInStencilFile(C.VIS_STENCIL_CONTAINERS, C.VIS_MS_US),
        )
        # reuse the stencil if already open — reopening an open stencil can
        # raise on real Visio
        base = os.path.basename(str(path))
        stencil = next(
            (d for d in self._stencil_docs() if str(d.Name) == base), None
        )
        if stencil is None:
            stencil = self._guard(
                "opening the built-in container stencil",
                lambda: app.Documents.OpenEx(path, C.STENCIL_OPEN_FLAGS),
            )
        if master_name:
            return self._find_master(master_name, stencil=str(stencil.Name))
        return stencil.Masters.Item(1)

    def add_container(
        self,
        label: str,
        member_ids: list[int],
        master: Optional[str] = None,
        padding_in: float = 0.4,
        page_name: Optional[str] = None,
    ) -> dict:
        if not member_ids:
            raise VisioMcpError("member_ids must list at least one shape to contain")
        app = self._app()
        page = self._page(page_name)
        members = [self._shape_by_id(page, sid) for sid in member_ids]
        container_master = self._container_master(master)

        def _build():
            # bounding box of the members, in page coordinates; the local pin
            # is not necessarily the shape center (icon masters often anchor
            # bottom-center), so subtract LocPin rather than assuming Width/2
            lefts, rights, bottoms, tops = [], [], [], []
            for m in members:
                left = float(m.CellsU("PinX").ResultIU) - float(m.CellsU("LocPinX").ResultIU)
                bottom = float(m.CellsU("PinY").ResultIU) - float(m.CellsU("LocPinY").ResultIU)
                lefts.append(left)
                rights.append(left + float(m.CellsU("Width").ResultIU))
                bottoms.append(bottom)
                tops.append(bottom + float(m.CellsU("Height").ResultIU))
            cx = (min(lefts) + max(rights)) / 2
            cy = (min(bottoms) + max(tops)) / 2
            container = page.Drop(container_master, cx, cy)
            container.CellsU("Width").ResultIU = max(rights) - min(lefts) + 2 * padding_in
            container.CellsU("Height").ResultIU = max(tops) - min(bottoms) + 2 * padding_in
            container.Text = label
            props = container.ContainerProperties
            if props is None:
                raise VisioMcpError(
                    "The dropped shape is not a Visio container — try a different "
                    "container master."
                )
            for m in members:
                props.AddMember(m, C.VIS_MEMBER_ADD_EXPAND_CONTAINER)
            # Drop puts the container on top of its members; push it behind
            # them so an opaque container fill can't hide the shapes
            container.SendToBack()
            return container

        # one undo scope so a mid-build failure can't leave an orphan container
        scope = self._guard("starting undo scope", lambda: app.BeginUndoScope("Add container"))
        try:
            container = self._guard(f"creating container {label!r}", _build)
            self._guard("closing undo scope", lambda: app.EndUndoScope(scope, True))
        except Exception:
            try:
                app.EndUndoScope(scope, False)  # False = roll the drop back
            except Exception:
                pass
            raise
        return {**self._shape_info(container), "member_ids": member_ids}

    def container_members(
        self,
        action: str,
        container_id: int,
        member_ids: list[int],
        page_name: Optional[str] = None,
    ) -> dict:
        if action not in ("add", "remove"):
            raise VisioMcpError(f"action must be 'add' or 'remove', got {action!r}")
        if not member_ids:
            raise VisioMcpError("member_ids must not be empty")
        page = self._page(page_name)
        container = self._shape_by_id(page, container_id)
        members = [self._shape_by_id(page, sid) for sid in member_ids]

        def _apply():
            props = container.ContainerProperties
            if props is None:
                raise VisioMcpError(
                    f"Shape {container_id} is not a container — create one with "
                    "add_container."
                )
            for m in members:
                if action == "add":
                    props.AddMember(m, C.VIS_MEMBER_ADD_EXPAND_CONTAINER)
                else:
                    props.RemoveMember(m)

        self._guard(f"{'adding' if action == 'add' else 'removing'} container members", _apply)
        return {"container_id": container_id, "action": action, "member_ids": member_ids}

    def badge_container(
        self,
        container_id: int,
        icon: str,
        corner: str = "top_left",
        size_in: float = 0.35,
        page_name: Optional[str] = None,
        inset_in: float = 0.15,
    ) -> dict:
        """Place a small header icon just inside a container's corner.

        Resolves `icon` to a built-in master first, then to a labeled image in
        the local icon folder; the badge becomes a container member so it moves
        with the zone. Placement is computed from the container's bounds — no
        caller coordinate math.
        """
        corners = ("top_left", "top_right", "bottom_left", "bottom_right")
        if corner not in corners:
            raise VisioMcpError(f"corner must be one of {list(corners)}, got {corner!r}")
        if size_in <= 0:
            raise VisioMcpError(f"size_in must be positive, got {size_in}")
        app = self._app()
        page = self._page(page_name)
        container = self._shape_by_id(page, container_id)
        props = container.ContainerProperties
        if props is None:
            raise VisioMcpError(
                f"Shape {container_id} is not a container — create one with "
                "add_container, then badge it."
            )

        icon_master = None
        icon_path = None
        try:
            icon_master = self._find_master(icon)
        except VisioMcpError:
            icon_path = _find_local_icon(icon)
            if icon_path is None:
                raise VisioMcpError(
                    f"No Visio master or local icon matches {icon!r}. Open the "
                    "stencil that contains the master, or save a labeled image "
                    f"(e.g. {icon.strip().lower().replace(' ', '-') + '.svg'!r}) "
                    "in the local icon folder (see list_local_icons), then retry."
                )

        # container bounds in page coordinates (LocPin is not always center)
        left = float(container.CellsU("PinX").ResultIU) - float(container.CellsU("LocPinX").ResultIU)
        bottom = float(container.CellsU("PinY").ResultIU) - float(container.CellsU("LocPinY").ResultIU)
        right = left + float(container.CellsU("Width").ResultIU)
        top = bottom + float(container.CellsU("Height").ResultIU)
        half = size_in / 2
        x = (left + inset_in + half) if corner.endswith("left") else (right - inset_in - half)
        y = (top - inset_in - half) if corner.startswith("top") else (bottom + inset_in + half)

        def _build():
            if icon_master is not None:
                badge = page.Drop(icon_master, x, y)
                native_w = float(badge.CellsU("Width").ResultIU)
                native_h = float(badge.CellsU("Height").ResultIU)
                badge.CellsU("Width").ResultIU = size_in
                if native_w:
                    badge.CellsU("Height").ResultIU = size_in * (native_h / native_w)
                badge.CellsU("PinX").ResultIU = x
                badge.CellsU("PinY").ResultIU = y
            else:
                info = self.import_image(icon_path, x, y, size_in, None, page_name)
                badge = self._shape_by_id(page, info["shape_id"])
            props.AddMember(badge, C.VIS_MEMBER_ADD_EXPAND_CONTAINER)
            return badge

        scope = self._guard("starting undo scope", lambda: app.BeginUndoScope("Badge container"))
        try:
            badge = self._guard(f"badging container {container_id} with {icon!r}", _build)
            self._guard("closing undo scope", lambda: app.EndUndoScope(scope, True))
        except Exception:
            try:
                app.EndUndoScope(scope, False)
            except Exception:
                pass
            raise
        return {
            **self._shape_info(badge),
            "container_id": container_id,
            "corner": corner,
            "source": "master" if icon_master is not None else "local_icon",
        }

    # -- layout / introspection ------------------------------------------------

    def auto_layout(
        self,
        style: str = "flowchart_tb",
        spacing_in: float = 0.75,
        resize_page: bool = True,
        page_name: Optional[str] = None,
    ) -> dict:
        if style not in C.LAYOUT_STYLES:
            raise VisioMcpError(
                f"style must be one of {sorted(C.LAYOUT_STYLES)}, got {style!r}"
            )
        page = self._page(page_name)
        place_style, route_style = C.LAYOUT_STYLES[style]

        def _layout():
            sheet = page.PageSheet
            sheet.CellsU("PlaceStyle").ResultIU = place_style
            sheet.CellsU("RouteStyle").ResultIU = route_style
            sheet.CellsU("AvenueSizeX").ResultIU = spacing_in
            sheet.CellsU("AvenueSizeY").ResultIU = spacing_in
            page.Layout()
            if resize_page:
                page.ResizeToFitContents()

        self._guard(f"auto-layout ({style})", _layout)
        return {"page": str(page.Name), "style": style, "note": "Export a PNG to inspect the result."}

    def get_page_state(self, page_name: Optional[str] = None) -> dict:
        page = self._page(page_name)
        shapes = []
        for shape in _items(page.Shapes):
            info = self._shape_info(shape)
            try:
                info["master"] = str(shape.Master.NameU) if shape.Master is not None else None
            except Exception:
                info["master"] = None
            is_connector = False
            try:
                is_connector = bool(int(shape.OneD))
            except Exception:
                pass
            info["is_connector"] = is_connector
            try:
                containers = shape.MemberOfContainers
                info["container_ids"] = [int(c) for c in containers] if containers else []
            except Exception:
                info["container_ids"] = []
            try:
                info["is_container"] = shape.ContainerProperties is not None
            except Exception:
                info["is_container"] = False
            if is_connector:
                from_id = to_id = None
                from_end: dict = {}
                to_end: dict = {}
                try:
                    for connect in _items(shape.Connects):
                        cell_name = str(connect.FromCell.Name)
                        target = int(connect.ToSheet.ID)
                        end_info = self._connect_endpoint_info(connect)
                        if cell_name.startswith("Begin"):
                            from_id, from_end = target, end_info
                        elif cell_name.startswith("End"):
                            to_id, to_end = target, end_info
                except Exception:
                    pass
                info["from_id"] = from_id
                info["to_id"] = to_id
                info["from_glue"] = from_end.get("glue")
                info["from_side"] = from_end.get("side")
                info["from_pos"] = from_end.get("pos")
                info["from_point"] = from_end.get("point")
                info["to_glue"] = to_end.get("glue")
                info["to_side"] = to_end.get("side")
                info["to_pos"] = to_end.get("pos")
                info["to_point"] = to_end.get("point")
                info["route_kind"] = (
                    "manual" if self._is_manual_connector(shape) else "auto"
                )
            shapes.append(info)
        sheet = page.PageSheet
        return {
            "page": str(page.Name),
            "page_size_in": [
                round(float(sheet.CellsU("PageWidth").ResultIU), 3),
                round(float(sheet.CellsU("PageHeight").ResultIU), 3),
            ],
            "shapes": shapes,
        }
