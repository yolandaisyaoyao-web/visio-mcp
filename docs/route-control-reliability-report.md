# 固定路径连接线可靠性修补报告

- 分支：`fix/route-control-reliability`（基于 `feat/connector-route-control` @ `b58fa58`）
- 范围：只修复已识别的可靠性问题；未开发寻路算法、未增加依赖、未合并 `main`、未推送远程。
- 验证环境：Windows 11 + 本地 Visio 16.0（中文，BASFLO_M）。

## 修复内容

1. **固定路径恢复原生自动布线**
   - `set_connector_route(connector_id, restore_auto=True)`：将手动 1-D 线转换回 Visio 动态连接线。
   - 保留起止形状（拓扑）、同一连接点（静态粘附、侧别、位置）、文字、箭头、线型/线宽/颜色，以及转换前的 `ShapeRouteStyle` / `ConLineRouteExt`。
   - 已是自动布线的连接线为幂等无操作；`restore_auto` 与 `waypoints` 互斥。
2. **几何诊断漏报/不准确修复**
   - 交叉：报告一对连接线之间的**全部**不同交点；只有交点同时落在两条连接线的全局起止端（共享挂点）时才忽略，落在折点/内部顶点上的交点不再漏报。
   - 重叠：报告全部**实际重叠区间**（页面坐标端点），同一线段上的相邻重叠片段会合并；`segment` 字段也改为真实区间，不再返回整条线段。
   - 完整性：几何提取失败时 `summary.diagnostics_complete=false`、`geometry_failed_ids` 列出降级连接线，并给出说明；不再把 Begin-End 直线回退结果当作完整几何。
3. **`_clean_path()` 修复**
   - 共线简化只在路径继续同向时删除中间点，反向折返点保留，不再删除“去而复返”的路径。
   - 构建完成后用 `route_violations()` 校验正交性、首段沿 from 侧法向、末段逆 to 侧法向、waypoint 仍在路径上且顺序正确；违反则拒绝写入。
4. **转换样式保留与失败回滚**
   - 动态 ↔ 手动转换统一走 `_copy_connector_style()`；`LinePattern` / `LineWeight` / `LineColor` / 箭头 / 文字复制失败时抛出错误并回滚 undo scope，**先完成全部复制才删除原连接线**。
   - 真实 Visio 验证：注入样式复制失败后原连接线仍在、页面形状数不变、新建的连接点也被撤销。
5. **节点移动后的固定路径诊断**
   - 手动路由报告 `orthogonal`、`non_orthogonal_segments`（含偏差量）与 `needs_readjust`；节点移动导致页面锚定折点与端点错位时明确提示重新调用 `set_connector_route`。
   - `summary.manual_needs_readjust_ids` 汇总需要调整的连接线。
6. **接口与文档**
   - `route_diagnostics(page, connector_id=...)` 支持按连接线 ID 筛选（summary 同步收敛到该线）。
   - README、MCP 工具说明与 docstring 更新；API 只做增量参数/字段，旧调用保持兼容。

## 测试证据

### 单元测试（跨平台 fake Visio）

```
.venv\Scripts\python.exe -m pytest
158 passed, 1 skipped
```

原有 140 项全部通过；新增 18 项针对本次缺陷的测试（折点交叉、多处交叉/重叠、实际重叠区间、
几何失败完整性、反向路径/waypoint 保持、恢复自动布线、样式失败回滚、节点移动非正交、ID 筛选）。

### 真实 Visio 验收

| 套件 | 结果 |
|---|---|
| `scripts/verify_connector_sides.py` | **52/52 PASS** |
| `scripts/regression_algorithm_flowcharts.py` | **功能正确性 58/58 通过**；布线质量观察项 53/72（见下） |
| `scripts/verify_route_control.py` | **54/54 PASS**（原 39 项 + 新增 15 项） |

`verify_route_control.py` 新增 [E] 段覆盖：

- 手动路由 → 移动节点 → 诊断 `needs_readjust=true` → 保存/重开后仍准确 → 重新指定 waypoint 恢复正交。
- `restore_auto` 后：动态连接线、同一 `bottom/top @ 0.5` 静态粘附、文字/样式/箭头保持、拓扑不变；
  保存重开仍保持；随后可再次转换回固定路径。
- `restore_auto` 与“动态 → 手动”转换分别注入样式复制失败：原连接线保留、形状数不变、
  动态转换失败时新建连接点回滚（行数 `[4, 4] → [4, 4]`）。
- 按 `connector_id` 筛选诊断。

### B/C/D 布线改善不退化

- B：分道后 0.15 in 共线重叠 1 → 0；交叉按对不增加（2 → 1）。
- C：分道后 0.15 in 共线重叠 1 → 0；交叉按对不增加（3 → 2）。
- D：6.5 in 完全并线、标签冲突与交叉全部清零；移动 P6、保存重开均保持；拓扑与节点位置不变。
- 三张图的拓扑、节点位置、页面分道设置均持久化。

### 关于综合回归质量观察项的变化

修正后的诊断首次显示 B/C/D 基线中既有的**折点/汇合接触**（例如 B 的 `(6.1, 4.309)`、
C 的 `(5.187, 6.7)`、D 的 `(10.213, 4.5)`）：这些点在旧实现中因距线段端点 0.05 in 以内被忽略。
本次未改动这些基线图的线路几何；功能检查仍为 58/58。分道/固定走廊对它们的效果与交叉对统计
已记录在 `route-control-data.json`。

## 产物

- `D:\桌面\test\visio-mcp-regression\route-control-data.json`
- `route-control-roundtrip-route-controlled.vsdx/.png`
- `B_decisions_merge-route-controlled.vsdx/.png`
- `C_multiple_loops-route-controlled.vsdx/.png`
- `D_tdma_reproduction-route-controlled.vsdx/.png`
- `regression-data.json`（综合回归）
