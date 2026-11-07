---
name: pdf-auto-analyze
description: >-
  本地 PDF 自动解读技能。当用户要求"分析 / 解读 / 总结这份 PDF"，或在消息中给出本地 .pdf 文件路径时，
  按固定流水线提取全文、识别文档类型（技术笔记 / 课程资料 / 报告）、拆解核心结论与流程、
  整理关键参数表、列出工程风险与坑点，并区分「AI 模型负责」与「外部业务规则负责」，
  最终产出一份整洁的结构化 Markdown 分析报告并落盘。仅处理磁盘上已存在的 PDF，不碰邮箱与网络下载。
whenToUse: >-
  用户提到"分析这个 PDF""帮我解读这份文档""总结这份资料/讲义/报告"，或输入里带有本地 .pdf 路径
  （如 C:\db\docs\xxx.pdf、./paper.pdf）时使用。不要用于邮件附件接收、在线 PDF 抓取或 PDF 编辑。
---

# PDF 自动解读（pdf-auto-analyze）

> 安装方式、可选依赖（pdftotext / Windows OCR / 视觉模型）与已知限制见 [README.md](README.md)。

把一份本地 PDF 变成一份可交付的 Markdown 分析报告：事实、流程、参数、风险四维拆解，外加 AI 与业务规则的职责划分。

## 1. 能力边界（必须遵守）

| 做 | 不做 |
|---|---|
| 读取本地已存在的 PDF（绝对路径或工作区相对路径） | **不接收邮件、不登录邮箱、不抓取附件** |
| 提取全文并按页建立引用锚点 | 不联网下载 PDF、不调用在线 OCR 服务 |
| 输出 `<文件名>.analysis.md` 分析报告 | 不修改、不重写、不删除原始 PDF |
| 用**本地** OCR 还原截图里的文字（见 Step 2.5） | 不调用**在线/云** OCR 服务 |
| 缺依赖时给出明确的安装/替代建议 | 不臆造文档中不存在的数据与参数 |

> 若用户给的是邮件链接、网盘链接、URL：先明确拒绝并说明本技能只处理本地文件，请其先保存到本地再给路径。

## 2. 触发条件

满足任一条即启用：

1. 用户表达"分析 / 解读 / 总结 / 讲一下"某份 PDF；
2. 用户消息里出现以 `.pdf` 结尾的本地路径（Windows 形如 `C:\...\x.pdf`，或相对路径 `docs/x.pdf`）；
3. 用户追问已分析过的 PDF 的细节（复用已提取的文本，不必重复提取）。

## 3. 输入参数

| 参数 | 必需 | 说明 |
|---|---|---|
| `pdf_path` | 是 | PDF 路径。相对路径按当前工作目录解析；不确定时用 `glob` 找 `*.pdf` 再确认 |
| `focus` | 否 | 关注点，如"只看第 3 章的部署方案""重点抽参数表" |
| `output_dir` | 否 | 默认写到 PDF 同目录；用户指定则从其指定 |

若 `pdf_path` 缺失或含糊（多个候选 PDF），先列出候选让用户选，不要盲猜。

## 4. 执行流水线

### Step 0 — 体检（不读内容也要先做）

用 `pwsh` 确认文件存在、大小、是不是真 PDF：

```powershell
$p = 'C:\path\to\doc.pdf'
Test-Path -LiteralPath $p
(Get-Item -LiteralPath $p).Length
# 魔数校验：真 PDF 前 5 字节为 %PDF-（兼容 PS 5.1 / PS 7）
$fs = [System.IO.File]::OpenRead($p); $buf = New-Object byte[] 5
$null = $fs.Read($buf, 0, 5); $fs.Close()
[System.Text.Encoding]::ASCII.GetString($buf)
```

- 非 `%PDF` 开头 → 报"文件不是有效 PDF（可能是改名的 docx/图片/下载失败页）"。
- 大小为 0 → 报"文件为空"。
- 无权限或路径不存在 → 直接进入 §6 失败处理，**不要**换路径硬猜。

### Step 1 — 用 `read` 工具试读（第一优先级，需验真）

按用户要求先 `read` 目标 PDF，但必须判定返回值性质：

| 返回特征 | 判定 | 处理 |
|---|---|---|
| 出现 `obj` / `endobj` / `stream` / `FlateDecode` / 大量控制字符或乱码 | 二进制流，**不是正文** | 丢弃，进入 Step 2 |
| 只有零星可读英文短句、无成段行文 | 仅部分可读 | 作为辅助线索，正文以 Step 2 提取结果为准 |
| 成段可读行文（未压缩的简单 PDF） | 可作正文 | 仍需 Step 2 复核页数与完整性 |

> 铁律：`read` 对 PDF 返回的是原始字节流，绝大多数压缩过的 PDF 会得到乱码。**绝不能**把乱码当正文去写报告。

### Step 2 — 文本提取（优先级链，逐级降级）

**① pdftotext（若安装）**

```powershell
pdftotext -layout -enc UTF-8 "C:\path\doc.pdf" "$env:TEMP\doc.txt"
```

**② Python + pypdf（Windows 上最稳的兜底）**

先探依赖，缺则装（装前告诉用户一句，装失败就往下走）：

```powershell
py -c "import pypdf, sys; print('pypdf', pypdf.__version__)" 2>$null
if ($LASTEXITCODE -ne 0) { py -m pip install -q --disable-pip-version-check pypdf }
```

再用固定脚本提取，**必须带页号与统计**（用于识别扫描件、生成页码引用）：

```powershell
$script = @'
import sys
from pypdf import PdfReader
src, dst = sys.argv[1], sys.argv[2]
r = PdfReader(src)
if r.is_encrypted:
    try:
        r.decrypt("")
    except Exception as e:
        print(f"ENCRYPTED: {e}"); sys.exit(3)
out, stats = [], []
for i, page in enumerate(r.pages, 1):
    try:
        t = page.extract_text() or ""
    except Exception as e:
        t = f"[[EXTRACT_ERROR page {i}: {e}]]"
    stats.append((i, len(t)))
    out.append(f"\n===== [PAGE {i}] =====\n{t}")
open(dst, "w", encoding="utf-8").write("".join(out))
total = sum(n for _, n in stats)
print(f"PAGES={len(stats)} CHARS={total} CHARS_PER_PAGE={total/max(len(stats),1):.0f}")
print("EMPTY_PAGES=" + ",".join(str(i) for i, n in stats if n < 20))
'@
Set-Content -Encoding utf8 "$env:TEMP\pdfx.py" $script
py "$env:TEMP\pdfx.py" "C:\path\doc.pdf" "$env:TEMP\doc.txt"
```

**③ 全部失败** → 进入 §6，输出明确错误提示，不要凭文件名猜内容。

### Step 2.5 — 图片文字还原（可选）

> 两条路线按能力从高到低挑：**① 视觉模型直连**（能描述图形语义）＞ **② Windows 本地 OCR**（只能读文字）。当前会话模型不支持 image 输入时，① 仍可用——走脚本直连 API，不必切换会话模型。

#### 路线 ①：视觉模型直连 API（推荐，能看懂图）

适用：会话模型不支持 `read_image`，或需要理解波形/接线/界面布局而非仅文字。

> **为什么不直接用 `read_image` 看图**：图片一旦作为 image 块进入会话历史，`selectModel` 会拒绝把该会话切回纯文本模型（`dsh-host-apiproxy`：会话含图片 + 目标模型未声明 image → `model-unavailable`），且工具结果裁剪只裁 text 块、图片块原样保留。**结论：图要走脚本转成文字落盘，永远别塞进会话上下文。** 这样任何模型（含便宜的 flash）都能随时接手同一批结论。

```powershell
# 三家环境变量：VLM_BASE_URL / VLM_API_KEY / VLM_MODEL
# 任选一个 OpenAI 兼容端点即可（Moonshot 只是示例，不是必需）
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$base  = if ($env:VLM_BASE_URL) { $env:VLM_BASE_URL.TrimEnd('/') } else { 'https://api.moonshot.cn/v1' }
$model = if ($env:VLM_MODEL)    { $env:VLM_MODEL } else { 'kimi-k3' }

# key 优先级：环境变量 → DSH 凭据库（凭据名按你的服务商改，Moonshot 用 MOONSHOT_API_KEY）
$key = $env:VLM_API_KEY
if (-not $key) {
  $credFile = "$env:USERPROFILE\.dsh\.credentials.yaml"
  if (Test-Path -LiteralPath $credFile) {
    $line = (Get-Content -LiteralPath $credFile) |
            Where-Object { $_ -match '^(MOONSHOT|VLM|OPENAI|DASHSCOPE|ZAI)_API_KEY\s*:' } |
            Select-Object -First 1
    if ($line) { $key = ($line -split ':\s*', 2)[1].Trim() }
  }
}
if (-not $key) { throw '缺少视觉模型 key：请设置 $env:VLM_API_KEY，或写入 ~/.dsh/.credentials.yaml' }

$img = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes('C:\path\to\page.jpg'))
$payload = @{
  model = $model                          # 必须是你账号 /v1/models 里真实存在的 id
  messages = @(@{ role = 'user'; content = @(
    @{ type = 'image_url'; image_url = @{ url = "data:image/jpeg;base64,$img" } },
    @{ type = 'text'; text = '逐字读出图中文字；若是图形，说明坐标轴、曲线走势与关键读数；看不清就说看不清。' }
  )})
  max_tokens = 4096                       # 见下方坑点：推理模型给小预算会返回空正文
} | ConvertTo-Json -Depth 8 -Compress
$r = Invoke-RestMethod -Uri "$base/chat/completions" -Method Post `
     -Headers @{ Authorization = "Bearer $key"; 'Content-Type' = 'application/json' } `
     -Body ([System.Text.Encoding]::UTF8.GetBytes($payload)) -TimeoutSec 300
$r.choices[0].message.content
```

**可替换的端点（都走 `image_url` + `data:` base64，字段名一致）**

| 服务商 | `VLM_BASE_URL` | 示例模型 id | 备注 |
|---|---|---|---|
| Moonshot（示例） | `https://api.moonshot.cn/v1` | `kimi-k3` | Kimi K2/K3 家族只接受 `temperature=1` |
| OpenAI | `https://api.openai.com/v1` | `gpt-5.x` / `gpt-4.1` | 无 temperature 限制 |
| 阿里百炼（兼容模式） | `https://dashscope.aliyuncs.com/compatible-mode/v1` | `qwen-vl-max` / `qwen3.6-plus` | 国内直连 |
| 智谱 | `https://open.bigmodel.cn/api/paas/v4` | `glm-4v-plus` | 国内直连 |
| 本地 vLLM / Ollama | `http://127.0.0.1:8000/v1` | 自部署视觉模型 | 无需 key，`VLM_API_KEY` 随便填 |

**必查的三个坑（实测踩过）**：

1. **模型 id 先验真**：`GET $base/models` 看账号实际开通了哪些。目录/文档里写着的 id 可能已退役或未授权，硬用会报 `Not found the model ... or Permission denied`。
2. **`temperature` 优先不传**：Kimi K2/K3 家族只接受 `1`，传 `0` 直接 400 `invalid temperature: only 1 is allowed`；若某端点报 temperature 相关错误，改成 `1` 或不传。
3. **`max_tokens` 给足（≥4096）**：推理模型把预算耗在思考上时会返回 **HTTP 200 + 空 content**，看着像成功、实际没内容。判据：`finish_reason=length` 且 `message.content` 为空 → 调大预算重试；同时可看 `message.reasoning_content` 是否被填满。

> **可选：让 DSH 会话内也能直接看图**（不推荐日常用，因为会话会被锁定在视觉模型上）。在 `~/.dsh/settings.yaml` 加：

```yaml
llm-pi-ai:
  providers:
    moonshotai-cn:
      apiKeyEnv: MOONSHOT_API_KEY
      models:
        - id: kimi-k3
          input: [text, image]
```

> 之后模型选择器里切到 `kimi-k3` 就能用 `read_image`；但**图片一旦进会话就切不回纯文本模型**（见上方警告），所以更推荐下面的脚本路线。

#### 路线 ②：Windows 本地 OCR（无 key 时的兜底，只出文字）

什么时候做：Step 3 判定"无文字层/扫描件"，或文档严重依赖截图（技术笔记常见），或用户问"图里写了什么"。

**前提判定**：

- 若当前模型**支持 image 输入**（`read_image` 不报 `does not declare image input`）→ 优先直接 `read_image` 看图，能读到图形语义（波形形态、接线、界面布局），OCR 做不到。
- 若模型只支持文本（如 `deepseek-v4-flash`）→ 走 Windows 内置 OCR，只能拿到**图上的字**，必须向用户说明"图形/波形未解读"。

**① 导出内嵌图片**

```powershell
$imgdir = Join-Path $env:TEMP 'pdfimgs'
@'
import sys, os
from pypdf import PdfReader
src, outdir = sys.argv[1], sys.argv[2]
os.makedirs(outdir, exist_ok=True)
r = PdfReader(src); n = 0
for pi, page in enumerate(r.pages, 1):
    try: imgs = list(page.images)
    except Exception as e:
        print(f"page {pi}: image enum failed {e}"); continue
    print(f"PAGE {pi}: {len(imgs)} images")
    for im in imgs:
        n += 1
        ext = os.path.splitext(im.name)[1] or ".png"
        fp = os.path.join(outdir, f"p{pi}_{n}{ext}")
        open(fp, "wb").write(im.data)
        print(f"  {im.name} -> {fp} {len(im.data)} bytes")
print("TOTAL", n)
'@ | Set-Content -Encoding utf8 "$env:TEMP\pdfimg.py"
py "$env:TEMP\pdfimg.py" "C:\path\doc.pdf" $imgdir
```

**② Windows 内置 OCR（离线，需已装中文语言包）**

```powershell
# 先确认可用语言：无 zh-Hans-CN 就别往下走，直接报告"需装 OCR 语言包"
[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages | ForEach-Object { $_.LanguageTag }

$ps = @'
Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue
# 注意：'IAsyncOperation`1' 必须用单引号，双引号里反引号会被当转义符吃掉
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and `
  $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op, $type) { $t = $asTask.MakeGenericMethod($type).Invoke($null, @($op)); $t.Wait(-1) | Out-Null; $t.Result }
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType=WindowsRuntime]
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage((New-Object Windows.Globalization.Language "zh-Hans-CN"))
if (-not $engine) { "ENGINE NULL"; exit 1 }
foreach ($f in Get-ChildItem (Join-Path $env:TEMP 'pdfimgs') -Include *.jpg,*.png,*.jpeg -Recurse | Sort-Object Name) {
  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($f.FullName)) ([Windows.Storage.StorageFile])
  $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bmp = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
  $res = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
  "===== $($f.Name) | $($dec.PixelWidth)x$($dec.PixelHeight) ====="
  foreach ($line in $res.Lines) { $line.Text }
  $stream.Dispose()
}
'@
Set-Content -Encoding utf8 "$env:TEMP\winocr.ps1" $ps
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\winocr.ps1"
```

**③ 结果处理**

1. OCR 中文会被插空格（`摆 杆 从 出 于`），报告里要清理后再引用：`py -c "import re,sys; t=open(sys.argv[1],encoding='utf-8').read(); print(re.sub(r'(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])','',t))" ocr.txt`
2. 度符号常被识别成 `。`/`.`，按语义补正**但必须注明"OCR 偏差，已按语义补正"**。
3. 只出现孤立数字的窄条截图（如 `.065`、`49.`）**不要猜含义**，写进"待确认"。
4. 报告元信息必须写明：`图片 N 张，OCR 识别 M 张，仅还原文字，图形未解读`。

**替代方案（Windows OCR 不可用时）**：`py -m pip install rapidocr_onnxruntime`（离线中文 OCR，质量更好，首次需联网下模型），或让用户把页面导出为文本。

### Step 3 — 质量闸门（决定能不能往下写）

| 指标 | 结论 | 动作 |
|---|---|---|
| `CHARS_PER_PAGE < 80` **且** 空页占比（<20 字符的页）> 60% | 多半是扫描件/图片型 PDF，无文字层 | 报告"该 PDF 无可用文字层（疑似扫描件），需先 OCR 再分析"，并列出已能读到的页 |
| `EMPTY_PAGES` 非空但占比不高 | 少数页无文字层（图表页/扫描插图） | 正常出报告，但在元信息标注缺失页 |
| `EXTRACT_ERROR` 大量出现 | 字体/编码异常 | 报告覆盖页范围，只对成功页出结论 |
| 页数与 §0 体检不一致 | 文件可能损坏 | 提示用户复核文件完整性 |
| 提取成功 | 继续 | 在报告元信息里写明 `覆盖页：1-N/N` |

### Step 4 — 识别文档类型

| 类型 | 识别信号 | 报告侧重点 |
|---|---|---|
| 技术笔记 | 命令行、配置片段、代码块、报错栈、"踩坑"口吻 | 流程与坑点优先，参数需给取值范围 |
| 课程资料 | 章节编号、定义-例题-练习、教学口吻、"回顾/作业" | 知识点层级 + 关键结论清单 |
| 报告/方案 | 摘要-背景-方案-评估-结论、图表编号、指标数字 | 结论先行 + 指标表 + 风险与未决项 |
| 混合 | 多种信号并存 | 按章节分别标注类型，不强行归一 |

### Step 5 — 四维拆解

1. **核心结论（事实层）**：只写文档明确断言的内容，每条尽量带页码锚点 `[p.12]`。区分"文档主张"与"客观事实"，后者若存疑需标注。
2. **关键流程流水线**：按 `输入 → 处理 → 输出` 展开，逐步写清触发条件、判定分支、异常兜底；有先后与依赖的用编号步骤或 mermaid 表示。
3. **关键参数表**：名称 / 取值 / 单位 / 作用 / 出处页码。原文缺单位就写"单位未标注"，不要补。
4. **工程风险与坑点**：风险 / 触发条件 / 影响面 / 缓解措施。原文没写缓解措施时标「文档未给出」。

### Step 6 — 划分「AI 模型职责」与「外部业务规则职责」

仅当文档含技术方案时执行；否则写一句"该文档不涉及技术职责划分"。

| 环节 | 承责方 | 判定依据 | 文档出处 |
|---|---|---|---|
| 例：意图识别 | AI 模型 | 语言理解、概率性输出 | [p.5] |
| 例：额度校验 | 外部业务规则 | 阈值/权限/计费，需确定性可审计 | [p.7] |

判定原则：

- 语义理解、生成、抽取、排序、改写 → **AI 模型**
- 阈值判定、权限、计费、合规校验、状态机流转 → **外部业务规则**
- 两者耦合的环节拆成两行分别标注「模型出候选 / 规则做裁决」，并写明交界处的输入输出契约
- 文档未明确归属的，标注「未明确（需与作者确认）」

### Step 7 — 生成报告并落盘

1. 按 §5 模板成文，用 `write` 保存为 `<PDF 同名>.analysis.md`（默认与 PDF 同目录，路径含空格需原样保留）。
2. 回复用户：报告路径 + 3 条以内 TL;DR + 覆盖页范围 + 遗留问题（若有）。
3. 中文文档就用中文报告；原文术语/参数名保留原文，不要翻译成歧义词。

## 5. 报告模板

```markdown
# 《<文档标题>》分析报告

## 0. 元信息
| 项 | 值 |
|---|---|
| 源文件 | <绝对路径> |
| 页数 / 覆盖页 | N / 1-N |
| 提取方式 | pdftotext / pypdf / read |
| 文档类型 | 技术笔记 / 课程资料 / 报告（可复合） |
| 分析时间 | <日期> |
| 提取质量 | 正常 / 部分页缺失 / 无文字层 |

## 1. 一句话结论
<30 字以内说清这份文档到底在讲什么、结论是什么>

## 2. 核心结论（事实层）
1. <结论> [p.X]
2. <结论> [p.Y]
> 引用原文关键句（可选）

## 3. 关键流程流水线
**整体链路**：<A → B → C>

| # | 步骤 | 输入 | 处理逻辑 | 输出 | 异常分支 |
|---|---|---|---|---|---|
| 1 | | | | | |

## 4. 关键参数表
| 参数 | 取值 | 单位 | 作用 | 出处 |
|---|---|---|---|---|

## 5. 工程风险与坑点
| # | 风险/坑点 | 触发条件 | 影响 | 缓解措施 | 出处 |
|---|---|---|---|---|---|

## 6. 职责划分：AI 模型 vs 外部业务规则
| 环节 | 承责方 | 判定依据 | 出处 |
|---|---|---|---|

## 7. 待确认 / 文档缺口
- <文档没讲清、前后矛盾、缺单位或缺前提的地方>

## 8. 附：章节索引（可选）
| 页码 | 内容 |
|---|---|
```

## 6. 失败处理（必须输出明确错误提示）

| 失败场景 | 检测方式 | 输出给用户的话术 |
|---|---|---|
| 路径不存在 | `Test-Path` 为 false | ❌ 未找到 PDF：`<路径>`。请确认路径（可提供绝对路径，或我先用 glob 搜 `*.pdf`）。 |
| 不是 PDF | 魔数非 `%PDF` | ❌ `<文件>` 不是有效 PDF（文件头为 `<实际魔数>`），可能是改名的 docx/图片或下载未完成。 |
| 文件为空/损坏 | 大小 0 或解析抛 `PdfReadError` | ❌ PDF 无法解析：`<异常信息>`。文件可能损坏或未完整下载，请重新获取。 |
| 加密 | `r.is_encrypted` 且空密码失败 | ❌ 该 PDF 已加密，无法读取正文。请提供密码或解密后的副本。 |
| 无文字层（扫描件） | Step 3 闸门命中 | 先按 Step 2.5 做本地 OCR；仍无字再报 ⚠️「该 PDF 无可用文字层（疑似扫描件/图片版），仅第 X 页有零星文本，我未做内容推测」。 |
| 依赖缺失 | `pdftotext` 与 `pypdf` 均不可用 | ❌ 本机缺少 PDF 解析能力（pdftotext/pypdf 均不可用）。可执行 `py -m pip install pypdf` 后重试，或先把 PDF 另存为文本再发我。 |
| 大文件超时 | 提取脚本超时/单页超长 | ⚠️ 文档较大（N 页），已分析第 1-M 页；如需剩余部分请指定页码范围。 |

**任何失败都不得编造报告内容**：宁可给一句"没读到"，也不要产出看似完整实则虚构的结论。

## 7. 质量红线

1. **参数可溯源**：每个数字/阈值/配置项都要能指回页码；找不到出处的写「未标注」。
2. **不补全**：文档没写的机制、没给的默认值一律不写，放到"待确认"。
3. **区分主张与事实**：文档说的方案 ≠ 已验证的结论，必要时标注「文档主张，未见验证」。
4. **覆盖透明**：元信息里写清页数与覆盖范围，提取失败页要点名。
5. **报告可读**：表格优先于长段落；一节一主题；不要复述全文目录。

## 8. 命令速查

```powershell
# 体检
Test-Path -LiteralPath $p; (Get-Item -LiteralPath $p).Length

# pypdf 可用性 + 安装
py -c "import pypdf; print(pypdf.__version__)"
py -m pip install -q pypdf

# 提取（脚本 $env:TEMP\pdfx.py，见 Step 2）
py "$env:TEMP\pdfx.py" "C:\path\doc.pdf" "$env:TEMP\doc.txt"

# 读取提取结果（用 read 工具，不要用 Get-Content 看长文本）
```

## 9. 最小可用路径（懒人版）

1. `Test-Path` + 魔数校验 →
2. `py -m pip install -q pypdf` 探针 →
3. `py pdfx.py in.pdf out.txt` →
4. 用 `read` 读 `out.txt` 并检查 `CHARS_PER_PAGE` →
5. 按模板出报告 → `write` 落盘 → 回报路径与 TL;DR。
