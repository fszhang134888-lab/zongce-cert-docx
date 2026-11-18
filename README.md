# zongce-bonus-cert-docx

按广东技术师范大学集成电路与信息工程学院模板制作「学院证明（含名单）」Word 文档并导出 PDF 的 WorkBuddy skill。适用于综测加分证明、活动参与证明、歌手大赛选手/嘉宾/观众名单等场景。

## 它解决什么

这类证明看着简单，实则有一堆隐形规范：标题宋体二号、正文仿宋三号、首行缩进两字符、16pt 字在 15.6pt 文档网格里实得 31.2pt 行距、名单用无框线五列表格、第一列只放班级且与「名单如下：」对齐、两字姓名补全角空格撑成三字宽、任何一格都不许折行……

这个 skill 把这些规范全部固化进脚本，并提供 docx/PDF 双层校验，避免"编辑器里看着对、交上去不对"。

## 安装

把整个目录放到 WorkBuddy 的 skills 目录下：

```bash
# Windows
C:\Users\<你的用户名>\.workbuddy\skills\zongce-bonus-cert-docx\

# macOS / Linux
~/.workbuddy/skills/zongce-bonus-cert-docx/
```

放好后重启会话即可。也可以直接调用脚本，不依赖 WorkBuddy。

### 让别人帮你装（防装错）

只说"帮我装 GitHub 上那个 skill"，AI 会自己猜**装到哪个目录**和**目录叫什么名**——这两处最容易错。把下面整段发给对方的 AI，它装完会回报路径和 name，你一眼就能核对：

```
请从 https://github.com/fszhang134888-lab/zongce-cert-docx 安装一个 Agent Skill。

硬性要求:
1. 目标目录名必须是 zongce-bonus-cert-docx,不是仓库名 zongce-cert-docx。
   因为 SKILL.md 里的 name 字段写的是 zongce-bonus-cert-docx。
2. SKILL.md 必须直接放在这个目录的根部,不要多套一层子文件夹。
3. 装完后自检并回报:
   - 列出该目录下的文件,确认 SKILL.md 存在
   - 告诉我你这台机器上 skill 的实际安装路径
   - 读出 SKILL.md 的 name 字段,确认是 zongce-bonus-cert-docx

依赖:python-docx、pymupdf、Pillow。生成 PDF 需要 Windows + Microsoft Word。
```

三种常见装错方式，对照检查：

| 错法 | 后果 |
|---|---|
| 目录名抄成仓库名 `zongce-cert-docx` | 按目录名注册的软件里，触发词对不上 |
| clone 后多套一层子目录 | SKILL.md 不在根目录，扫不到 |
| 装到了别的 skills 根目录 | 重启后完全没这个 skill |

**最终验证**：让对方的 AI 生成一份综测加分证明样例。能正常出 docx，就是装对了——这比任何路径检查都靠谱。

> 以上假设对方软件支持 Agent Skills 机制（能识别 `SKILL.md`），如 Claude Code、WorkBuddy。若软件只支持 rules / 自定义指令，则没有"安装"概念，需把 `SKILL.md` 内容作为提示词喂给它。

## 用法

```bash
cd scripts

# 1. 出一份样例看看版式
python gen_cert.py --demo 样例.docx

# 2. 导出配置 JSON，填自己的名单
python gen_cert.py --dump-config cert.json

# 3. 生成 docx
python gen_cert.py --config cert.json --out 综测加分证明.docx

# 4. 校验 docx 的字体、字号、表格结构
python verify_fonts.py 综测加分证明.docx --body-font 仿宋

# 5. 用 Word 导出 PDF（Windows，PowerShell）
$app = New-Object -ComObject Word.Application
$doc = $app.Documents.Open("C:\path\综测加分证明.docx", $false, $false)
$doc.ExportAsFixedFormat("C:\path\综测加分证明.pdf", 17)

# 6. 校验 PDF 的实际字体、列坐标、有无折行、落款是否齐平
python verify_pdf.py 综测加分证明.pdf --docx 综测加分证明.docx --title-font 宋体 --body-font 仿宋
```

配置文件 `cert.json` 里：

- `roster`：每个 `{label, names}` 是一个班；姓名超过 4 个自动续行，续行不再重复班级
- 正文支持两个占位符：`{n}` = 名单总人数，`{first}` = 名单第一位姓名
- `pad_two_char_names`：默认 `true`，两字姓名中间补一个全角空格撑成三字宽（学院成品的做法）

## 依赖

- Python 3.9+
- `python-docx`、`pymupdf`、`Pillow`
- Windows + Microsoft Word（用于导出 PDF）；`docx2pdf.ps1` 是 Word COM 调用的留存版本

## 文件结构

| 路径 | 说明 |
|---|---|
| `SKILL.md` | skill 主文档：铁律、模板规范、标准流程、常见坑 |
| `references/template-spec.md` | 官方模板与学院成品的实测数据（含字体/坐标/行距实验记录） |
| `scripts/gen_cert.py` | 生成器：字体探测、列宽规划、表格排版、占位符 |
| `scripts/verify_fonts.py` | docx XML 层校验 |
| `scripts/verify_pdf.py` | PDF 渲染层校验 |
| `scripts/docx2pdf.ps1` | Word COM 转 PDF（执行策略常拦，README 里有直接写法） |
| `assets/` | 官方格式示例、格式说明、早期截图（供对照，非运行时依赖） |

## 说明

仓库中的姓名均为演示用的虚构名。模板文件来自学院公开发放的格式说明，仅用于版式对照。
