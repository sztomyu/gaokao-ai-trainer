# 高考 AI 训练系统（gaokao-ai-trainer）

> 一个无需数据库、开箱即用的高考刷题与学情管理系统。教师端可管理题库、用大模型把 PDF/Word 试卷自动识别入库；学生端按知识点训练、错题重练、按遗忘曲线复测，并给出学情画像。

前后端一体：Flask 单进程同时提供 REST API、静态 SPA 与题目图片服务。数据全部存放在本地 Excel 文件里，**不需要安装 MySQL / Redis / Docker**，拷贝目录即可运行与迁移。

题目中的 **LaTeX 数学公式**、**mhchem 化学方程式**、**题目配图**（函数图像 / 几何图形 / 物理示意图 / 分子结构式）全部在浏览器中原位渲染，前端依赖已本地化，**完全离线可用**。

---

## 功能特性

### 教师端

- **题库管理**：按科目 / 章节 / 知识点 / 题型 / 难度增删改查，支持批量导入
- **AI 试卷识别**：上传 PDF 或 Word，调用视觉大模型逐题提取题干、选项、答案、解析，自动切分知识点
- **配图自动截取**：基于 PyMuPDF 解析 PDF 矢量绘图，按 XY 坐标聚类成图形块，再匹配到最近的题目，导出为 PNG 并嵌入题干
- **公式还原**：解析 docx 内嵌的 OMML 公式对象（分式、上下标、根式），转换为可读文本或 LaTeX，避免公式丢失
- **组卷**：手动选题或按条件自动组卷，支持真题卷 / 模拟卷
- **账号管理**：教师 / 学生双角色，教师接口做角色校验

### 学生端

- **知识点训练**：选定科目 → 章节 → 知识点，连续出题
- **模拟卷**：按条件随机组卷限时作答
- **即时判分**：单选 / 多选 / 填空自动判分，解答题给出参考答案供自评
- **错题本**：自动收集答错题目，支持重练
- **遗忘曲线复测**：已达成的知识点按 `1 / 2 / 4 / 7 / 15 / 30` 天间隔安排复测，到期在首页提醒
- **学情统计**：正确率、知识点掌握度雷达图、薄弱点排序（ECharts 渲染）

---

## 快速开始

### 环境要求

- Python 3.9+
- Node.js 16+（可选，仅用于 `npm run dev` 启动器）

### 安装与运行

```bash
git clone https://github.com/sztomyu/gaokao-ai-trainer.git
cd gaokao-ai-trainer

# 1. 安装依赖
pip install -r requirements-web.txt

# 2. 启动（任选其一）
python server/app.py --port 7100            # 方式一：直接启动
npm run dev -- --port 7100                  # 方式二：npm 启动器，自动探测 Python
```

浏览器打开 <http://127.0.0.1:7100> 即可。

### 默认账号

| 账号 | 密码 | 角色 |
|---|---|---|
| `teacher` | `123456` | 教师 |
| `student` | `123456` | 学生 |

> 首次部署请立即修改默认密码，或直接在 `server/data/users.xlsx` 中删除演示账号后新建。

### 配置 AI 试卷识别（可选）

不带 API Key 时系统照常运行，只是「试卷识别」功能不可用。启用步骤：

1. 以教师账号登录 → 系统设置 → 大模型 API
2. 分别填入 **文本模型** 与 **视觉模型** 的 Base URL、模型名、API Key，点「测试」
3. 保存后会写入 `server/data/settings.json`（该文件已在 `.gitignore` 中，不会被提交）

`server/data/settings.example.json` 给出了配置结构参考，也可直接复制为 `settings.json` 后手动填写。

兼容任何 OpenAI 风格的 `/chat/completions` 接口，例如 DeepSeek、通义千问（DashScope）、本地 vLLM 等。

---

## 数据存储

免数据库，5 个 Excel 文件即全部状态，均位于 `server/data/`：

| 文件 | 内容 |
|---|---|
| `questions.xlsx` | 题目主表（`questions` 子表）+ 知识点树（`knowledge_points` 子表） |
| `users.xlsx` | 账号、密码哈希（SHA-256 加盐）、角色 |
| `papers.xlsx` | 真题卷 / 模拟卷及题目组成 |
| `records.xlsx` | 作答记录、错题 |
| `mastery.xlsx` | 知识点掌握度画像、复测调度 |
| `media/` | 题目配图 |
| `uploads/` | 上传的原始试卷 |

所有写操作经全局锁串行化，写入前自动备份为 `.bak`。

### 内置题库

仓库自带一份可直接使用的题库，**1527 道题**，覆盖 6 个科目：

| 科目 | 题数 | 科目 | 题数 |
|---|---|---|---|
| 物理 | 598 | 英语 | 207 |
| 化学 | 342 | 语文 | 79 |
| 数学 | 226 | 生物 | 75 |

含 257 道配图题、1148 个知识点三元组、309 个章节。可直接使用，也可在教师端清空后导入自己的题库。

---

## 目录结构

```
gaokao-ai-trainer/
├── server/                     # Flask 后端
│   ├── app.py                  #   API 路由 + 静态托管 + 媒体服务
│   ├── excel_store.py          #   数据访问层（Excel 读写）
│   ├── engine.py               #   判分 / 掌握度 / 组卷 / 复测调度
│   ├── study_loop.py           #   学习闭环（错题、相似题、复测）
│   ├── analytics.py            #   学情统计
│   ├── ai_recognize.py         #   大模型试卷识别 + PDF 图形截取
│   ├── data/                   #   Excel 数据 + media/ 配图
│   └── uploads/                #   上传的原始试卷
├── web/                        # 静态 SPA（原生 JS，无构建步骤）
│   ├── index.html
│   ├── css/app.css
│   ├── js/                     #   api / render / teacher / student / app
│   └── vendor/                 #   KaTeX + mhchem + ECharts（本地化，离线可用）
├── scripts/
│   ├── dev.mjs                 #   开发启动器
│   ├── gen_exercise_book.py    #   按知识点导出可打印 Word 习题集（含答案附录）
│   ├── gen_knowledge_map.py    #   导出考点分布统计 Word
│   ├── gen_syllabus.py         #   导出考纲详解 Word
│   ├── import_physics_book.py  #   从 docx 教辅批量导入题目（含 OMML 公式解析）
│   ├── test_student_flow.py    #   学生端全流程接口回归测试
│   └── verify_retest_count.py  #   复测计数一致性校验
├── legacy/                     # v1 tkinter 桌面版源码（参考用，不含数据）
├── requirements-web.txt
└── package.json
```

---

## 常用脚本

```bash
# 按知识点导出某一科目的习题集（正文只有题目，答案与解题步骤集中在书末附录）
python scripts/gen_exercise_book.py --subject 物理 --output "物理高考练习.docx"

# 从 docx 教辅批量导入题目
python scripts/import_physics_book.py

# 学生端全流程回归（需服务已在 7100 端口运行）
python scripts/test_student_flow.py
```

---

## 技术栈

| 层 | 选型 |
|---|---|
| 后端 | Flask（单进程，`threaded=True`） |
| 数据 | openpyxl + Excel 文件 |
| 文档解析 | PyMuPDF（PDF 矢量图形）、python-docx（docx + OMML 公式） |
| 前端 | 原生 JS（无框架、无构建）、KaTeX、mhchem、ECharts |
| 鉴权 | Header `X-Auth-Token`，SHA-256 令牌；教师接口额外校验角色 |

---

## 版权与免责声明

- **代码**以 MIT 协议开源，可自由使用、修改、商用。
- **题库内容与配图**来源于公开的高考真题与教辅材料，**版权归原出题机构与出版方所有**。本项目仅出于学习与研究目的收录，不主张任何权利，也不提供任何形式的商业授权。
- 如果你是权利人并认为本仓库内容侵犯了你的权益，请提交 Issue 或联系作者，我们将在收到通知后**立即删除相关内容**。
- 本项目与任何考试机构、出版社无关联，题目答案与解析由程序或模型生成，**可能存在错误，请以官方答案为准**，不构成备考承诺。

## License

[MIT](LICENSE) © 2026 sztomyu
