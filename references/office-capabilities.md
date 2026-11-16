# 文件处理和长期偏好

本技能包含 `scripts/office_helper.py`，可以从技能目录用 Python 运行。随包脚本使用 Python 3.9 或更新版本的语法与标准库；本机已测试版本为 3.12，其他版本及依赖的兼容性须实际检查。还需要各命令对应的库：Word 使用 `python-docx`，Excel 使用 `openpyxl`，文字通知图片使用 `Pillow` 和中文字体。偏好功能仅使用 Python 标准库。库或文件访问能力缺失时说明限制。除非使用者在当前任务明确授权，不得通过 pip、系统包管理器、下载程序或其他方式安装、升级或修改依赖；不得因缺依赖而联网。优先使用平台原生文件能力、视觉能力和生图能力。

## Word、Excel 和图片

- `python scripts/office_helper.py read-docx input.docx`：输出正文段落和表格；不保证页眉、文本框、图片文字、修订及版式。DOCX 扫描图需要平台图片理解或 OCR。
- `python scripts/office_helper.py read-xlsx input.xlsx --max-rows 100 --max-cols 40`：分表读取值和公式文本，未计算公式；查看 total_rows、total_columns、more_rows 和 more_columns，分批时使用 --start-row、--start-col 逐块读取，直到所有范围覆盖；指定 --sheet 可按表单独处理。验证结果、行列、合并单元格和公式缓存，不能把公式表达式当作计算结果。
- `python scripts/office_helper.py inspect-image input.png`：只读取图像尺寸等属性，无法识图、识字。要理解图片或识别文字时优先使用平台视觉/OCR；Windows 图片 OCR 缺口可按 [Windows 离线图片 OCR 后备](ocr-winrt.md) 检查本机 WinRT。图片脚本仅处理图片；扫描 PDF 可检查同一指南中的 WinRT PDF 后备。输出必须逐页对照原图人工核验。上述能力均不可用时说明限制并请用户提供文字或其他资料。现场照片的解释仍需人工复核。
- 用 `create-docx`、`create-xlsx`、`create-image` 的 `--spec INPUT.json --output OUTPUT` 新建基础文件；输出不覆盖已有文件。图片仅将用户给定文字排成通知 PNG，不是凭描述生成照片或海报；需要生成或编辑真实图像时使用平台生图工具。不能将基础文件称作保留用户模板。
- 使用已有 Word 模板时按 [Word 模板填充与验收](word-template.md) 定点填写、核对结构和页面；只支持可定位的正文单节点目标，复杂结构按指南处理。修改 Excel 时参照 [受限数值填写](excel-template.md)，仅在支持条件下定点另存并核验；含公式、合并单元格或数据验证的文件须使用其他可用工具并单独验证。格式不支持、依赖缺失或无法验证时如实说明。

输入 JSON 示例（按需要生成，内容均为虚构；`create-docx` 与 `create-image`）：
```json
{"title":"物业通知", "paragraphs":["请关注公共区域清洁安排。", "具体时间：待确认"]}
```
`create-docx` 可选 `"tables": [[["项目", "状态"], ["清洁", "待确认"]]]`。Excel 示例：
```json
{"sheets":[{"name":"跟进清单","rows":[["事项","状态"],["公共区域清洁","待确认"]]}]}
```
新表所有字符串按文本处理，避免将不可信文字当公式执行；需公式、样式或复杂工作簿时使用相应专业工具，验证公式计算。

## 跨会话偏好

在本地文件系统可访问时，优先用用户指定的偏好文件；否则默认 `~/.property-work-assistant/preferences.json`，可用 `PROPERTY_ASSISTANT_PREFERENCES` 或 `--store PATH` 指向用户确认的持久位置。启动相关任务时执行 `python scripts/office_helper.py preferences list` 读取已确认偏好；默认路径在当前操作系统用户目录，须同一账号和可持续存储才能跨会话读取。每次向使用者清楚说明保存位置；不同平台或设备不会自动共享。

只有当使用者明确表示要长期记住某项偏好时，才调用 `preferences add --scope 场景 --text 偏好 --confirmed` 保存。添加前去除个人、商户、合同、金额和敏感资料；涉及业务规则需要有来源和负责人确认，不能由单次案例推断。使用者要求查看或删除时使用 `preferences list` 或 `preferences delete --id ID --confirmed`；修改时先列出并使用 preferences update --id ID --text 新偏好 --confirmed 一次写入；有需要可用 --scope 更新范围，避免先删除后新增导致记录丢失。读写失败应报告并可给出可复制偏好，不得声称已记住；这些文件不在技能包内，不应随技能包分享。

当前任务的明确要求优先，且按适用范围使用偏好。长期偏好不是对事实的记忆，不记录或自动复用客户数据、具体任务金额或未经确认的规则。存储依赖平台能运行 Python 和访问同一持久路径；没有这些权限时不能保证跨会话记忆。
