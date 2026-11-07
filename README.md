# pdf-auto-analyze

> DeepSeek Harness（DSH）技能：给一个本地 PDF，自动产出一份结构化的 Markdown 分析报告。
> 四维拆解：**核心结论 / 键流程流水线 / 关键参数表 / 工程风险坑点**，外加 **AI 模型 vs 外部业务规则** 的职责划分。

一份 PDF 进来，一份 `xxx.analysis.md` 出去。不联网抓文档、不碰邮箱、不改原文件。

---

## 它解决什么问题

技术笔记、课程资料、竞赛/项目报告这类 PDF，常见形态是：**正文没多少字，关键信息都在截图里**。人工读要来回翻，AI 读又容易只读文字层、把图当空气。

本技能把流程固定成一条流水线：

1. **体检** —— 文件存在？大小正常？前 5 字节是 `%PDF` 吗？（防"改名的 docx"和"下载未完成的空壳"）
2. **文本提取** —— `read` 试读验真 → `pdftotext` → `pypdf`，全程带页号与字符统计
3. **质量闸门** —— 按"字符/页 + 空页占比"判断是不是扫描件；命中就不硬写，明确告知需 OCR
4. **图内信息还原** —— 双路线：视觉模型直连（能看懂波形/坐标轴）＞ Windows 离线 OCR（只出文字）
5. **四维拆解** —— 结论带页码锚点、流程按"输入→处理→输出"、参数表标出处、风险表标缓解措施
6. **职责划分** —— 语义理解/生成归 AI 模型，阈值/权限/计费/状态机归外部业务规则，耦合处拆两行
7. **产出报告** —— 按固定模板渲染，落盘 `<PDF 同名>.analysis.md`
8. **失败话术** —— 路径错/非 PDF/加密/扫描件/依赖缺失/超大文件，各有明确的报错文案，**禁止编造内容**

## 它不做什么

- ❌ 不收邮件、不登录邮箱、不抓附件
- ❌ 不联网下载 PDF、不调用在线 OCR 服务
- ❌ 不修改 / 不重写 / 不删除原始 PDF
- ❌ 文档里没写的数据一律不补（缺单位就写"单位未标注"，缺缓解措施就写"文档未给出"）

---

## 安装

```powershell
# 1) 放到 DSH 的技能目录（Windows 下即 C:\Users\<你>\.dsh\skills\）
git clone <本仓库地址> "$env:USERPROFILE\.dsh\skills\pdf-auto-analyze"

# 2) 装 Python 依赖（文本提取主力）
py -m pip install -r "$env:USERPROFILE\.dsh\skills\pdf-auto-analyze\requirements.txt"
```

装完重启 DSH（或刷新页面），技能清单里就会出现 `pdf-auto-analyze`。

### 可选依赖

| 依赖 | 用途 | 装不上怎么办 |
|---|---|---|
| `pypdf`（必需） | 文本提取 + 图片导出 | 装不上就只能走 `pdftotext` |
| `pdftotext`（poppler） | 更精确的版面提取 | 缺它不影响，`pypdf` 兜底 |
| Windows 10/11 + `zh-Hans-CN` OCR 语言包 | 图片内文字（离线免费） | 缺它就走视觉模型路线 |
| 任一 OpenAI 兼容的多模态 API key | 看懂图形语义（波形/界面/接线） | 缺它就只能读文字，读不出图形 |

> 系统要求：**Windows**（PowerShell 5.1 或 7）。因为 OCR 路线用的是 `Windows.Media.Ocr`，Linux/macOS 需要自行改用 tesseract / PaddleOCR。

---

## 怎么用

只要在对话里表达"分析这份 PDF"或给出 `.pdf` 路径，技能就会被触发：

```
你：C:\Users\me\Desktop\部署笔记.pdf 帮我分析一下
```

```
你：把这份技术笔记解读成结构化报告，重点抽参数表和坑点
```

```
你：这份方案里哪些环节该交给模型、哪些必须走业务规则？
```

```
你：分析 C:\docs\scan.pdf
AI：⚠️ 该 PDF 无可用文字层（CHARS_PER_PAGE=2，空页 18/18，疑似扫描件）。
    需先 OCR 再分析，我不做内容推测。
```

## 报告结构

| # | 章节 | 内容 |
|---|---|---|
| 0 | 元信息 | 源文件、页数/覆盖页、提取方式、识别质量、文档类型 |
| 1 | 一句话结论 | 30 字内说清这份文档讲了什么、结论是什么 |
| 2 | 核心结论（事实层） | 逐条结论 + 页码锚点 `[p.12]`，区分"文档主张"与"事实" |
| 3 | 关键流程流水线 | 步骤 / 输入 / 处理逻辑 / 输出 / 异常分支 |
| 4 | 关键参数表 | 参数 / 取值 / 单位 / 作用 / 出处页码 |
| 5 | 工程风险与坑点 | 风险 / 触发条件 / 影响面 / 缓解措施 |
| 6 | 职责划分 | AI 模型 vs 外部业务规则，含判定依据与出处 |
| 7 | 待确认 / 文档缺口 | 文档没讲清、前后矛盾、缺单位、缺前提的地方 |

---

## 关键配置：视觉模型（可选）

技能默认**不需要**任何云端 key：文本靠 `pypdf`，图内文字靠 Windows 自带 OCR，完全离线。

只有当你需要**看懂图形**（波形走势、坐标轴读数、OLED 界面布局、电路接线）时，才配一个多模态模型。任何 OpenAI 兼容端点都行：

```powershell
$env:VLM_BASE_URL = 'https://api.moonshot.cn/v1'   # 换别家就改这里
$env:VLM_MODEL    = 'kimi-k3'
$env:VLM_API_KEY  = 'sk-...'                       # 或写进 ~/.dsh/.credentials.yaml
```

| 服务商 | Base URL | 示例模型 |
|---|---|---|
| Moonshot | `https://api.moonshot.cn/v1` | `kimi-k3` |
| OpenAI | `https://api.openai.com/v1` | `gpt-5.x` / `gpt-4.1` |
| 阿里百炼（兼容模式） | `https://dashscope.aliyuncs.com/compatible-mode/v1` | `qwen-vl-max` |
| 智谱 | `https://open.bigmodel.cn/api/paas/v4` | `glm-4v-plus` |
| 本地 vLLM / Ollama | `http://127.0.0.1:8000/v1` | 自部署视觉模型 |

三个实测踩过的坑（详见 `SKILL.md` Step 2.5）：

1. 模型 id 先用 `GET $base/models` 验真，别照抄文档；
2. `temperature` 优先别传（Kimi K2/K3 只接受 `1`）；
3. `max_tokens` 给足 ≥4096——推理模型会返回 **HTTP 200 + 空正文**，看着成功其实没内容。

## 已知限制

- **扫描件/图片型 PDF**：无文字层时只能拿到图，必须依赖 OCR 或视觉模型；两者都没有就明确报错，不出结论。
- **Windows 独占的 OCR 路线**：用的是 `Windows.Media.Ocr`；非 Windows 环境需自行替换。
- **视觉模型只解决"看得懂"**：波形图的**通道含义、单位**若图上没标，技能不会猜，一律进"待确认"。
- **别用 `read_image` 读图**：图片一旦进入 DSH 会话历史，该会话就无法再切回纯文本模型（会报 `model-unavailable`），且工具结果裁剪不会移除图片块。本技能因此统一走"脚本直连 API → 结果落盘为文字"的路线。
- **超长文档**：默认全量处理；数百页时可指定页码范围。

## 目录结构

```
pdf-auto-analyze/
├── SKILL.md          # 技能主体（DSH 读取的指令）
├── README.md         # 本文件
├── LICENSE           # MIT
└── requirements.txt  # pypdf
```

## License

MIT，见 [LICENSE](LICENSE)。
