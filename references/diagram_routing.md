# 科研绘图路由与视觉自检清单 (diagram_routing.md)

`tju-ppt` 自带四条绘图路线，**零外部 Skill 依赖**。必需依赖仅 `python-pptx`、`Pillow`、`numpy`、`lxml`；
路线 C 额外需要 `cairosvg`，路线 D 需要 `matplotlib`；Mermaid 为可选增强。

## 1. 路由决策

| 场景 | 路线 | 工具 | 产物是否可在 PPT 中改字 |
|---|---|---|---|
| ≤ 6 步线性流程、输入→核心→输出 | A1 原生流程 | `add_flow_pipeline` | ✅ |
| 闭环控制 / 迭代优化 / 训练-验证回路 | A2 原生闭环 | `add_feedback_loop` | ✅ |
| 分层系统架构（感知/决策/控制/应用） | A3 原生分层 | `add_layered_architecture` | ✅ |
| 任意两盒子连线、正交折线 | A4 原生连接 | `connect_boxes` / `add_connector` | ✅ |
| 多分支汇合、跨分组连线、技术路线图、算法框图 | C 示意图引擎 | `TjuDiagram`（`tju_svg_diagram.py`） | ❌（保留 `.svg` 与生成脚本可重绘） |
| 曲线、柱状、消融、热力图 | D 数据图 | `tju_plot_style.py` | ❌（保留脚本与数据可重绘） |
| 时序图 / 状态机（机器上已有 `mmdc`） | M 可选 | `mmdc -c references/tju_mermaid_theme.json` | ❌ |
| 论文原图（真实实验照片、平台照片） | — | `add_image_fit(..., border_color=C_BLUE)` | ❌（不做调色板审计） |

判定顺序：**能用 A 就不用 C**（导师可直接改字）；连线需要绕行 / 分组域 / 3 条以上汇入同一节点时改用 C。

## 2. 共同规则

1. **按最终放置尺寸作图**：`TjuDiagram(w_in, h_in)`、`new_figure(w_in, h_in)` 的尺寸 = PPT 中放置框尺寸，
   这样图中 18 pt 文字放进 PPT 后仍是 18 pt；不要先画大图再缩小。
2. **四色**：节点角色 `input`（浅蓝底蓝框）/`core`（深蓝底白字）/`output`（白底红框）/`emphasis`（红底白字）/`plain`/`note`；
   数据图仅 3 种线色，第 4 条起依靠线型 + 标记区分；柱状图用 `TJU_BAR_STYLES`（填充 + 阴影线）。
3. **字体**：西文 Times New Roman、中文 微软雅黑（Linux 无雅黑时自动退到 Noto Sans CJK SC 等黑体）。
   `python scripts/tju_svg_diagram.py` 可打印实际解析到的字体。
4. **上下标**：节点与标签文字同样支持 `_{sub}` / `^{sup}`；**禁止组合重音符**（`x̂`, `ẋ`），
   cairosvg 无法合成会出现方块——带帽变量放到幻灯片的原生 OMML 公式里。

### TjuDiagram 常用图元
- `node(..., dashed=True)`：虚线框（用于待确认 / 休眠 / 占位等非实体状态）。
- `node(..., check=False)`：示意性图元允许与其他节点重叠，不参与重叠与穿框校验。
- `arrow([(x1, y1), (x2, y2), ...], both=...)`：不绑定节点的自由折线箭头，支持双向箭头两端回缩。
- `text(x, y, ..., anchor="middle")`：分区副标题、脚注，支持 `\n` 多行文本并参与出界与压框校验。
- `from_spec(spec)`：JSON IR 接口完整支持 `dashed` / `check` / `align` / `wrap`、`arrows`、`group.pt` 与 `line_spacing`。
- 节点内文字自动按真实字宽折行并均衡行宽，避免单字孤行；①②③、■▲ 等符号自动交给中文字体渲染。

## 3. 已知限制（务必知晓）

- `cairosvg` 只读取 `font-family` 的第一个字体且无字形回退，因此引擎逐段指定字体；若手写 SVG 必须同样处理。
- `matplotlib < 3.6` 无字体回退：含中文的文字整串使用中文字体（其中的西文也会是黑体），纯西文才用 Times。
  升级到 `matplotlib >= 3.6` 后可实现同一标签内中西文分设。
- 路线 C/D 的图片在 PPT 中不可编辑，务必把生成脚本与 `.svg` 一起保存在项目 `figures/` 目录。

## 4. 视觉自检清单（生成后逐项过）

**自动（必须 0 问题）**
- `TjuDiagram.save(strict=True)`：文字溢出、节点重叠、连线穿过无关节点、标签压框/出界、边拐点/自由箭头/分组框出界、组合重音符。
- `python scripts/tju_audit.py deck.pptx --skip 1,2 --images figures/*.png`：
  普通文本与 OMML 公式字号/粗体/四槽位字体、越色与主题色回退、残留 `_{` 与裸数学下标、承上启下、备注、包围盒重叠；自绘图片越色像素 ≤ 0.5%。
- `python scripts/tju_audit.py deck.pptx --skip 1,2 --content`：检查单页密度、纯文字页（原生流程图自动放行）、多行溢出估计、口语化/绝对化表述与缩写首现。

**人工（看 144 dpi 渲染 PNG）**
- 箭头方向符合数据流 / 控制流语义（分层架构默认 `flow="up"` 数据上行）。
- 箭头首尾贴边、无悬空；回环线不压文字。
- 图例不遮挡曲线（必要时 `ax.set_ylim` 留顶部空间或 `ncol` 横排）。
- 同一页里图片文字与正文视觉字号一致（按放置尺寸作图即可保证）。
- 红色元素每栏不超过 1–2 个。
