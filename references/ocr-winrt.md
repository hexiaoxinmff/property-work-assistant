# Windows 离线 OCR 后备

当平台当前没有可用的图片或扫描 PDF 文字识别能力、任务需要将资料转成可搜索文本，且运行环境为 Windows 时，可检查本机 WinRT 后备：图片使用 `scripts/ocr_images_winrt.ps1`；扫描 PDF 使用 `scripts/ocr_pdf_winrt.ps1`。两者调用 Windows 自带 OCR，不上传文件、不联网、不安装依赖；均需要 Windows PowerShell 5.1 和系统已安装的 `zh-Hans-CN` OCR 语言包。PDF 另依赖系统的 WinRT PDF 渲染组件。运行前确认平台允许调用本机 PowerShell，并让用户知道识别在本机执行。

支持 PNG、JPG/JPEG、BMP、TIF/TIFF。脚本为每张图输出同名 UTF-8 TXT，并生成 `ocr_manifest.json`，记录来源路径、源文件 SHA-256、识别语言、图像尺寸、行数和待人工核对状态。必须指定一个尚不存在的新输出目录；同名输入会被拒绝，避免输出碰撞。源图只读。

示例（从技能目录运行；输出目录必须是尚不存在的新目录）：

```powershell
powershell.exe -NoProfile -File ".\scripts\ocr_images_winrt.ps1" -InputPath "D:\samples\notice.png" -OutputDirectory "D:\outputs\notice_ocr"
```

PDF 示例（仅适用于扫描 PDF，且需新输出目录）：

```powershell
powershell.exe -NoProfile -File ".\scripts\ocr_pdf_winrt.ps1" -InputPath "D:\samples\notice.pdf" -OutputDirectory "D:\outputs\notice_pdf_ocr" -ResolutionDpi 200
```

PDF 脚本逐页用 Windows WinRT PDF 渲染器生成内存位图，再调用系统 OCR；默认 200 DPI，每页限制 40,000,000 像素，可用 `-ResolutionDpi`（96–400）与 `-MaxPixelsPerPage` 调节。每个 PDF 生成同名 UTF-8 TXT，并在输出目录生成 `ocr_manifest.json`，记录源 PDF SHA-256、页数、分辨率、页尺寸、行数及待人工核对状态。必须使用全新输出目录；重复文件名会被拒绝；源 PDF 只读。加密、损坏、超限、系统组件不可用或运行失败时停止并如实报告。

没有语言包或 Windows PowerShell 5.1 时停止并说明。识别文本可能出现错字、漏字、空格或标点误识；此方法不提供置信度、校对或准确率保证。逐页对照原图，尤其人工核对编号、金额、日期、主体、位置、期限、表格、公章及手写内容；不清晰处标为待核实，不能用 OCR 结果直接发函、结算或作为正式业务记录。若 OCR 执行失败或输出异常，保留错误和已有输出，不声称完成，转用经确认的平台视觉能力或报告限制。

本机虚构样例于 2026-11-14 实测：1400×700 中文图片生成 6 行转录，核心中文句子可辨，但编号连字符被识别为“一”、部分标点误识。单页扫描 PDF 在本机以 200 DPI 渲染为 2800×1400，再生成 6 行转录；编号字符及标点也有误识，源 PDF SHA-256 前后相同。两项结果仅证明本机调用链可执行，不代表一般图像质量、准确率或其他机器表现。

