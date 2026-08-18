# AGENTS.md

## 项目背景

这是"上海小学三年级学生作业"仓库，位于 `C:\mywork\projects\student_homework`，试卷存放在 `三年级/` 目录。
现有科目：语文（chinese_*.doc）、英语（english_*.doc）、数学（math_*.doc），使用沪教版教材。

用户是家长，典型需求：**基于已有复习卷 + 小朋友的错题记录，生成一份新的复习卷**（难度略增、题量更多、覆盖全册知识点、帮助找出薄弱点）。

## 文件约定

- 命名：`<subject>_YYYYMMDD.doc`（如 `math_20260816.doc`、`math_20260818.doc`）。
- **`.doc` 实际是 RTF 文本**（文件头 `{\rtf1 \ansi \ansicpg936`，WPS Office 保存），用 RTF 方式读写，不是 Word 二进制。
- 原卷末尾常有家长手写的"测试结果/错题记录"段落，出题前必须读到文件结尾并归纳错题类型。
- 新卷末尾必须附"参考答案（供家长核对）"，并加一段给家长的"注"：说明覆盖的知识点范围、错题强化题对应位置。
- 用户此前要求新卷"不限定时间和分数"（不要写 时间：xx分钟 满分：xx分）。

## Word 文件读写

详细方法、完整脚本和踩坑记录见 **[word文件读写方法.md](word文件读写方法.md)**，核心要点：

1. **读取**：用其中的提取脚本把 RTF 转 UTF-8 纯文本。中文是 GBK 双字节 `\'xx` 转义，必须连续配对解码；`\uN` 转义要跳过后面的回退字符。
2. **写入**：生成纯 ASCII RTF，所有非 ASCII 字符用 `ch.encode('gbk')` 转 `\'xx`。
3. **致命坑**：`\colortbl` 等目的地控制字必须用花括号包成 `{\colortbl;...}`，否则 Word 打开全空白（自写解析器看不出问题）。`\fonttbl` 同理。
4. **不要二次转义**：字符串里放真实 `\t`/空格，由转义函数统一处理一次。
5. **验证**：用 Word COM 自动化（`New-Object -ComObject Word.Application`）实开文件，检查 `Characters.Count` 是否数千级、能否找到"参考答案"；再用提取脚本回读核对完整性。

## 工具使用说明

- **shell_command（PowerShell）**：文件枚举、字节头检查、Word COM 验证。命令默认不加 `sandbox_permissions`（审批策略为 never，加了会被拒绝）。
- **Python**：嵌入式发行版 `C:\software\python-3.14.3-embed-amd64`。无 striprtf/pandoc/python-docx，不可联网安装；RTF 解析/生成按方法文档手写。
- **apply_patch**：创建、修改、删除文件的唯一方式（UTF-8 安全，中文路径不会乱码）。删除临时脚本用 apply_patch 的 Delete File，`Remove-Item` 可能被沙箱策略拦截。
- **中文路径/乱码**：不要把中文路径经 stdin 管道传给 `python -`（会变 `???`）；脚本先 apply_patch 落盘再运行。PowerShell 管道输出中文乱码多为显示层问题，验证一律读 UTF-8 文件（`Get-Content -Raw -Encoding UTF8`）。
- **Word 锁文件**：目录出现 `~$*.doc` 说明用户正用 Word 打开该文件，不要覆盖写入对应试卷。

## 标准工作流（生成新试卷）

1. 提取原卷文本，读错题记录 → 归纳薄弱知识点。
2. 设计题目：覆盖三年级全册（口算、竖式、递等式、填空、判断、选择、图形、解决问题、拓展），对错题类型设置强化题。
3. 生成脚本中对每个答案写 `assert` 程序化校验，全部通过再生成文件。
4. 生成 `三年级/<subject>_YYYYMMDD.doc`，命名与日期一致。
5. Word COM 实开验证非空白，回读提取文本核对完整性。
6. 用 apply_patch 清理临时脚本，保持目录只留试卷与本文档。

## 仓库状态

- 原卷（chinese/english/math_20260816.doc）已提交 git；新生成的卷默认为未跟踪文件，保留给用户。
- 当前已有 `math_20260818.doc`（第二份数学复习卷），含参考答案与错题强化说明。
