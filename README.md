# 五年级信息科技教案生成器

按固定版式生成「五年级信息科技电子教案」docx，保证**严格两页、不要大片空白**。

支持 **271 教学模式**（导入2分钟 → 教学7分钟 → 反思/拓展1分钟）和 **八种阳光卡** 自动标注。

---

## 快速开始

### 1. 环境要求

| 需要 | 说明 |
|------|------|
| Python 3.8+ | 生成 docx 只需标准库，无需额外安装 |
| pywin32（可选） | `pip install pywin32`，用于 Windows 上通过 Word 精确校验页数 |
| Microsoft Word（可选） | 配合 pywin32 实现真实分页校验 |

**没有 pywin32 / Word 也能用**，脚本会自动退回估算模式生成 docx，只是不能精确保证两页。

### 2. 目录结构

```
wunianji-jiaoan/
├── README.md                    ← 本文件
├── SKILL.md                     ← WorkBuddy AI 专用入口（其他平台可忽略）
├── assets/
│   └── 教案模板.docx             ← 空白版式模板（从参考教案净化而来）
├── references/
│   ├── 格式规范.md               ← 精确版式参数（页面、表格、字体、行距）
│   └── 内容写法.md               ← 271框架、八种阳光卡、各栏目写法示例
└── scripts/
    ├── build_lesson_plan.py     ← 核心生成脚本（通用，不依赖任何平台）
    ├── check_pages.py           ← Word COM 页数校验（需要 pywin32）
    ├── prepare_template.py      ← 模板净化脚本（从参考教案生成空白模板）
    └── test_content.json        ← 示例内容
```

### 3. 生成教案

准备一个 JSON 内容文件（参考 `scripts/test_content.json`），然后运行：

```bash
# 基本用法
python scripts/build_lesson_plan.py --content 内容.json --out 教案.docx

# 用内联 JSON
python scripts/build_lesson_plan.py --json '{"lesson_ordinal":"3","lesson_name":"...","design":[...]}' --out 教案.docx

# 跳过分页校验（快速生成）
python scripts/build_lesson_plan.py --content 内容.json --out 教案.docx --no-check

# 用自定义模板
python scripts/build_lesson_plan.py --content 内容.json --out 教案.docx --template 我的模板.docx
```

### 4. 分页校验

- **有 pywin32 + Word**：脚本自动用 Word COM 读取真实页数，二分搜索课后反思行高，精确保证 2 页。
- **没有 pywin32 / Word**：脚本用内容行数估算行高，生成后建议手动打开 Word 检查页数。
- **macOS / Linux**：目前无法用 Word COM 校验，建议安装 LibreOffice 后手动检查，或使用 `--no-check` 模式。

---

## 在不同 AI 平台上使用

### WorkBuddy AI
SKILL.md 已配置好，直接对话说「帮我写第N课教案」即可触发。

### ChatGPT / Claude / 其他 AI
1. 把 `references/内容写法.md` 的内容发给 AI，让它按 271 模式和阳光卡规则写教案。
2. 把 AI 生成的内容组装成 JSON 格式（参考下方格式说明）。
3. 保存为 `.json` 文件，运行 `build_lesson_plan.py` 生成 docx。

### Cursor / VS Code
1. 打开本目录作为工作区。
2. 在对话中让 AI 按 `references/` 里的规范写教案内容。
3. 用终端运行 `build_lesson_plan.py` 生成 docx。

---

## 内容 JSON 格式

```json
{
  "designer": "张 根",
  "week": "4",
  "teach_date": "",
  "class_no": "五年级     班",
  "lesson_ordinal": "3",
  "lesson_period": "1",
  "lesson_name": "游戏体验寻规律",
  "objectives": [
    "1. 了解算法的基本概念，知道算法是解决问题的一系列明确步骤。",
    "2. 通过游戏体验活动，能够将简单问题分解为有序步骤。",
    "3. 感受算法在生活中的广泛应用，激发学习兴趣。"
  ],
  "key_difficult": [
    "重点：理解算法是解决问题的一系列明确步骤。",
    "难点：将实际问题抽象为算法步骤。"
  ],
  "methods": "任务驱动法、游戏教学法、小组合作法",
  "preparation": "多媒体教室、PPT课件、游戏素材",
  "blackboard": ["第3课 游戏体验寻规律", "算法 → 解决问题的步骤"],
  "personalized": "",
  "reflection": "",
  "design": [
    {"kind": "title", "text": "一、情境导入（2分钟）（乐学与善思）"},
    {"kind": "body", "text": "教师出示走迷宫游戏..."},
    {"kind": "title", "text": "二、探究新知（7分钟）（分析与质疑）"},
    {"kind": "sub", "text": "活动1：走迷宫游戏体验"},
    {"kind": "body", "text": "分组进行走迷宫游戏..."},
    {"kind": "title", "text": "三、总结拓展（1分钟）（总结与反思）"},
    {"kind": "body", "text": "今天我们了解了算法的概念..."}
  ]
}
```

### 字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| designer | string | 设计者，默认"张 根" |
| week | string | 周次 |
| teach_date | string | 授课时间，可留空 |
| class_no | string | 班级 |
| lesson_ordinal | string | 第几课 |
| lesson_period | string | 课时数，默认"1" |
| lesson_name | string | 课题名 |
| objectives | string[] | 教学目标，每条一行 |
| key_difficult | string[] | 重点难点，每条一行 |
| methods | string | 教法学法 |
| preparation | string | 教学准备 |
| blackboard | string[] | 板书设计，每行一条 |
| personalized | string | 个性化设计，可留空 |
| reflection | string | 课后反思，可留空（靠行高撑满第二页） |
| design | object[] | 教学设计正文，每项含 kind 和 text |

### design 的 kind 值

| kind | 样式 | 用途 |
|------|------|------|
| title | 加粗 | 大环节标题（一、二、三） |
| sub | 加粗 | 活动小标题（活动1、活动2） |
| body | 不加粗 | 正文内容 |

---

## 271 教学模式

每课时按 **2:7:1** 时间分配：

| 环节 | 时长 | 内容 |
|------|------|------|
| 导入 | 约2分钟 | 情境导入、复习引入、问题驱动 |
| 教学 | 约7分钟 | 新授、探究、实践（主要环节） |
| 反思/拓展 | 约1分钟 | 总结归纳、反思提升、拓展延伸 |

## 八种阳光卡

在合适的教学环节标注对应阳光卡（在环节标题后用括号注明）：

| 阳光卡 | 适用环节 |
|--------|---------|
| 乐学与善思 | 导入环节 |
| 实践与运用 | 动手操作、上机实践 |
| 劳动与创造 | 创作作品、完成任务 |
| 阅读与表达 | 朗读、分享交流 |
| 分析与质疑 | 提出问题、分析推理 |
| 合作与探究 | 小组合作、探究学习 |
| 总结与反思 | 总结环节、回顾反思 |
| 自主与管理 | 自主学习、课堂管理 |

---

## 自定义模板

如果想用自己的教案做模板：

1. 把你的教案 docx 放到 `assets/` 目录。
2. 运行 `python scripts/prepare_template.py`（修改脚本里的 SRC 路径指向你的文件）。
3. 或者直接用 `--template` 参数指定：`python build_lesson_plan.py --content 内容.json --out 输出.docx --template 你的模板.docx`

模板要求：10行6列表格，结构参考 `references/格式规范.md`。

---

## 常见问题

**Q: 生成的文件只有 1 页 / 超过 2 页怎么办？**
- 检查是否安装了 pywin32（`pip install pywin32`）。
- 没有 pywin32 时脚本用估算行高，可能不精确。手动打开 Word，调整课后反思行高。
- 教学设计正文建议控制在 600-980 字 / 17-27 段。

**Q: Mac / Linux 上能用吗？**
- 能生成 docx（纯 Python 标准库），但无法用 Word COM 校验页数。
- 用 `--no-check` 参数跳过分页校验，生成后用 LibreOffice / WPS 手动检查。

**Q: 模板里的格式跟我的不一样？**
- 用 `prepare_template.py` 从你自己的教案生成新模板。
- 或直接用 `--template` 参数指定你的模板文件。
