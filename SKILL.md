---
name: physics-variant-bank
description: Use when creating targeted variant practice sets, exam-review worksheets, or revised solutions and printable Word/PDF handouts from high school physics exams, lecture examples, or error sets（适用于高中物理试卷、例题或错题集的靶向变式出题、解析修订与 Word/PDF 讲义制作）.
---

# 高中物理靶向变式题库

将试卷、例题或错题集转化为以**物理建模、过程演化、临界判断与规律选择**为主线的深度靶向练习，输出排版稳定一致的 Markdown、Word 与 PDF（学生版 + 教师版）。

## 资产分工（按阶段按需读取，单一事实源）

| 路径 | 职责与读取时机 |
| :--- | :--- |
| [references/physics-design.md](references/physics-design.md) | 阶段一必读：实质变式、反同型测试、反数学化、阶梯设问、动态解析与 Markdown 书写底线 |
| [references/drawing-and-qa.md](references/drawing-and-qa.md) | 阶段二必读：二维/三维与双子图路由、4 档预设-嵌宽速查、视觉公理与 PDF 真实渲染终检 |
| [scripts/extract_source_docx.py](scripts/extract_source_docx.py) | 阶段一直接运行（纯标准库）：无损还原原卷/答案 DOCX 中的 OMML 公式并导出嵌入图锚点 |
| [scripts/select_subset.py](scripts/select_subset.py) | 考后抽题直接运行（纯标准库）：从全卷母库按错题题号抽题重排并同步裁剪教师版总览表 |
| [scripts/physics_draw.py](scripts/physics_draw.py) | 阶段二直接导入：300 DPI 纯黑白防压缩制图库（预设画布、燕尾箭头、简体宋体锁死、常用原语） |
| [scripts/build_docx.py](scripts/build_docx.py) | 阶段一 `precheck`（纯标准库）+ 阶段二 `compile`：A4 模板编译、表格防劈裂、OMML 正体锁死与七重终检 |

## 运行环境约定

- **纯标准库脚本**（`extract_source_docx.py`、`select_subset.py`、`build_docx.py precheck`）：直接用 `python3` 运行。
- **制图与编译脚本**（`physics_draw.py`、`build_docx.py compile`）：使用专用虚拟环境 `~/.local/share/physics-variant-bank-venv/bin/python`（或 `uv run --with-requirements requirements.txt python3`），严禁裸跑缺失依赖的系统 `python`。
- **Linux 服务端 PDF 中文字形**：本机已通过 `~/.config/fontconfig/fonts.conf` 将 `宋体/SimSun` 映射至 `Noto Serif CJK SC`、`黑体/SimHei` 映射至 `Noto Sans CJK SC`；跨机器迁移时须同步配置。

## 工作模式与启动规则

- **全中文交互：** 汇报、澄清与交付文档均用简体中文；符号、公式、路径与命令保留规范写法。
- **按范围直接启动：** 用户已指定题号、题量或修订任务时直接执行，不弹窗确认；范围未明确时简要给出 3 个选项：1) 逐题 1:1 变式；2) 挑典型高频易错题生成靶向变式；3) 仅输出考点诊断与原题答案核查表。
- **三模式自适应：**
  1. **整卷/分批直出模式（默认）：** 固定图号 `figs/fig_q{题号:02d}.png`、`figs_sol/sol_q{题号:02d}.png`，合并直接拼接 Markdown 重跑编译；
  2. **先库后卷分步模式（考前建库 $\to$ 考后按错题组卷）：** 阶段一仅保存 `student_bank.md`、`teacher_bank.md`（图位用 `<!-- 待配图：... -->` 占位）并必跑 `build_docx.py precheck`；考后用 `select_subset.py --pick ...` 抽题重排，仅对入选错题补绘配图并编译成品；
  3. **单题状态模式：** 仅当用户明确要求逐题建库或跨会话接续时启用（新序列统一归档至工作区 `题库序列/S00X_YYYYMMDD_卷名/`，不再沿用旧版单题 SVG 目录结构）。

## 固定交付契约

- **产物清单：** `student_bank.md`、`teacher_bank.md`、`figs/fig_qXX.png`、`figs_sol/sol_qXX.png`、`reference.docx`、`{卷名}_学生版.docx/.pdf`、`{卷名}_教师版.docx/.pdf`。
- **学生版排题契约（严禁含原题号提示或答案）：** 卷头 `# 标题` + `建议用时`（成套试卷附 `满分`）；题号用加粗纯文本（`**1.**`～`**N.**`），选择题选项 `A.`～`D.` 顶格独立成段。
  - **成套试卷输入**（原卷自带大题型分节与分值）：保留原卷大题分节、题序与分值；
  - **例题讲义 / 专题集 / 错题集输入**（无大题分节）：严格按原题顺序连续编号，不重排题序、不强行拆大题分节；多选题在题干首部标注 `（多选）`。
- **教师版固定骨架（与学生版 1:1 对应）：**
  1. `## 一、参考答案总览表`（选择题填表，含解答题时表下附简答）
  2. `## 二、原题考点诊断、答案核查与变式映射总览表`（固定 4 列：`题号 | 原题核心考点与高频易错点 | 原题参考答案核查 | 变式核心模型与思维升维切入点`）
  3. `## 三、逐题参考答案与深度详析`（每题固定顺序：**【参考答案】** $\to$ 教师解析图（按需） $\to$ **【分步解析】** $\to$ **【防错点拨】**）

## 两阶段执行流程

### 阶段一：无损提取、核查、实质变式定稿与双途径验算（先定题，后制图）
1. 若原卷/答案为 `.docx`，先运行 `python3 scripts/extract_source_docx.py <源.docx> --out-dir <临时目录>` 还原全部 OMML 公式并核查原图关键数据（严禁仅凭丢失根号/分式的纯文本 `.txt` 判题）；锁定学段边界并独立验算原题参考答案。
2. 读取 [references/physics-design.md](references/physics-design.md)，完成变式设计、三项反同型测试与反数学化审查。
3. 编写独立 Python 脚本用第二途径（如解析推导 vs 数值扫描/仿真）交叉校验全部选项、临界与数值，断言全过后再定稿 `student_bank.md` 与 `teacher_bank.md`，并执行阶段一 Markdown 预检：
   ```bash
   python3 scripts/build_docx.py precheck student_bank.md teacher_bank.md
   ```
   须输出 `[PRECHECK PASS]`；若采用“先库后卷”模式，至此即可交付 Markdown 母库，待考后用 `select_subset.py` 抽题进入阶段二。

### 阶段二：预设化配图、一键编译与真实渲染终检
1. 读取 [references/drawing-and-qa.md](references/drawing-and-qa.md)，按有效维度路由制图（使用专用 venv Python），逐图渲染并目视核验通过后再进入下一步。
2. 统一通过脚本完成编译与全部自动门禁（严禁手拼残缺命令或手改 DOCX XML）：
   ```bash
   VENV_PY=~/.local/share/physics-variant-bank-venv/bin/python
   $VENV_PY scripts/build_docx.py compile student_bank.md "{卷名}_学生版.docx" --expected-imgs N --pdf --preview-dir preview_s
   $VENV_PY scripts/build_docx.py compile teacher_bank.md "{卷名}_教师版.docx" --expected-imgs M --pdf --preview-dir preview_t
   ```
   须全部输出 `[PRECHECK PASS]` 与 `[PASS]`（门禁项由脚本自动校验）。
3. 打开 `preview_s/` 与 `preview_t/` 中的关键 PDF 渲染页目视抽检，核对字号、简体常规宋字形、公式、表格跨页与分页连贯性；如实说明已验证的渲染环境，清理临时预览目录后交付。
