# Word(.doc) 文件读写方法总结

本仓库中的试卷文件（`三年级/*.doc`）**本质是 RTF 文本**（文件头为 `{\rtf1 \ansi \ansicpg936`，由 WPS Office 保存），
因此可以完全绕开 Word 二进制格式，用文本方式读取和生成。本文档总结已验证的方法、可直接复用的脚本和常见坑。

## 0. 结论速览（下次先看这里）

| 事项 | 结论 |
| --- | --- |
| 文件本质 | `.doc` 后缀，RTF 内容，GBK 编码中文（`\'xx` 形式），头为 `{\rtf1 \ansi \ansicpg936` |
| 读取 | Python 写 RTF→纯文本提取脚本（见 §1），结果写 UTF-8 文本再查看 |
| 写入 | Python 生成 ASCII 纯文本 RTF，非 ASCII 字符转成 GBK `\'xx`（见 §2） |
| 致命坑 1 | `\colortbl` 等"目的地"控制字**必须用花括号包起来**：`{\colortbl;...}`，否则 Word 把后面全部正文当颜色表吞掉，文档打开是空白 |
| 致命坑 2 | GBK `\'xx` 是**双字节**，解码必须连续配对；单个字节解码会得到 `?` |
| 验证 | 用 Word COM 自动化打开检查字符数/页数（见 §3），不要只靠自写解析器 |
| 工具 | PowerShell + 嵌入式 Python 3.14（`C:\software\python-3.14.3-embed-amd64`）+ apply_patch |

---

## 1. 读取方法（RTF → 纯文本）

### 1.1 原理

- RTF 文件本身是 ASCII 文本；中文用 `\'xx` 十六进制转义表示 GBK 字节，例如 宋体 = `\'cb\'ce\'cc\'e5`。
- 部分字符用 Unicode 转义 `\uc1 \u183 ?`（`\uN` 后跟 1 个回退字符，需跳过）。
- 段落标记 `\par`、制表符 `\tab`、单元格结束 `\cell` 分别转成换行/制表符。
- 花括号分组 `{...}` 和控制字（`\fs24`、`\pard` 等）在正文提取时忽略。

### 1.2 可直接复用的提取脚本

```python
# -*- coding: utf-8 -*-
"""RTF(.doc) 提取为纯文本。用法：
    python _rtf_extract.py <目录> [文件名前缀]
输出同目录 _extracted.txt（UTF-8）。"""
import os
import re
import sys

BASE = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
PREFIX = sys.argv[2] if len(sys.argv) > 2 else ""
target = None
for name in os.listdir(BASE):
    if name.lower().startswith(PREFIX.lower()) and name.lower().endswith(".doc"):
        target = os.path.join(BASE, name)
        break
if target is None:
    raise SystemExit("no matching .doc found")

raw = open(target, "rb").read().decode("latin-1")


def decode_rtf(s):
    out, i, n = [], 0, len(s)
    while i < n:
        c = s[i]
        if c == "\\":
            j = i + 1
            m = re.match(r"[a-zA-Z]+", s[j:])
            if m:
                word = m.group(0)
                j += len(word)
                m2 = re.match(r"-?\d+", s[j:])
                if m2:
                    j += len(m2.group(0))
                if j < n and s[j] == " ":
                    j += 1
                if word == "par":
                    out.append("\n")
                elif word in ("line", "row"):
                    out.append("\n")
                elif word in ("tab", "cell"):
                    out.append("\t")
                i = j
            elif s[j : j + 1] == "'":
                out.append("\\x")
                i = j + 1
            else:
                if s[j : j + 1] in "\\{}":
                    out.append(s[j])
                i = j + 1
        elif c in "{}":
            i += 1
        else:
            out.append(c)
            i += 1
    return "".join(out)


def decode_escapes(s):
    out, i, n = [], 0, len(s)
    while i < n:
        c = s[i]
        if c == "\\" and s[i + 1 : i + 2] == "x":
            buf = bytearray()
            while s[i : i + 2] == "\\x" and i + 4 <= n:
                try:
                    buf.append(int(s[i + 2 : i + 4], 16))
                except ValueError:
                    break
                i += 4
            out.append(buf.decode("gbk", errors="replace"))
        elif c == "\\" and s[i + 1 : i + 2] == "u":
            m = re.match(r"\\u(-?\d+)", s[i:])
            if m:
                out.append(chr(int(m.group(1))))
                i += len(m.group(0))
                if s[i : i + 2] == "\\x":
                    i += 4
                elif i < n:
                    i += 1
            else:
                i += 1
        elif c == "\\" and s[i + 1 : i + 2] == "\\":
            out.append("\\")
            i += 2
        else:
            out.append(c)
            i += 1
    return "".join(out)


text = re.sub(r"\n{3,}", "\n\n", decode_escapes(decode_rtf(raw)))
out_path = os.path.join(BASE, "_extracted.txt")
open(out_path, "w", encoding="utf-8").write(text)
print(out_path)
```

运行：
```powershell
python _rtf_extract.py "C:\mywork\projects\student_homework\三年级" math_20260816
```

查看提取结果（务必显式 UTF-8，避免控制台乱码误判）：
```powershell
$t = Get-Content '...\_extracted.txt' -Raw -Encoding UTF8
```

### 1.3 读取时的坑

- **GBK 必须双字节配对解码**：`\'cb\'ce\'cc\'e5` 是"宋体"，拆开单字节解全会得到 `?`。
- **中文路径经 stdin 管道传给 `python -` 会变 `???`**：把脚本用 apply_patch 写到磁盘再运行（apply_patch 按 UTF-8 安全写入），不要在命令行里拼接中文路径给 Python。
- **管道显示乱码是显示层问题**：PowerShell 输出被重新编码时中文可能显示为 `ï¿½` 或 `ÈýÄê¼¶`，不代表文件损坏；一律读 UTF-8 文件验证。
- 原卷末尾常有家长手写的错题记录（如"测试结果：第四大题填空，第9小题……"），提取后务必读到文件结尾。

---

## 2. 编写方法（生成 RTF .doc）

### 2.1 原理

- 生成纯 ASCII 的 RTF：所有非 ASCII 字符（含全角标点、`× ÷ ＝ ＞ ＜`、中文）用 GBK 字节转成 `\'xx`。
- 必须给出：`\rtf1\ansi\ansicpg936`、花括号包裹的 `\fonttbl`、**花括号包裹的 `\colortbl`**。
- 段落用 `\pard\plain\lang2052\f1\fs24 ... \par`；标题用 `\qc` 居中、`\b ... \b0` 加粗。
- 文件名沿用约定：`math_YYYYMMDD.doc`。

### 2.2 核心函数（可直接复用）

```python
import os

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "三年级")
OUT_FILE = os.path.join(OUT_DIR, "math_YYYYMMDD.doc")


def rtf_escape(text):
    """把普通字符串转成 RTF 安全文本：ASCII 原样，其余转 GBK \'xx。"""
    out = []
    for ch in text:
        if ch in "\\{}":
            out.append("\\" + ch)
        elif ch == "\t":
            out.append("\\tab ")
        elif ord(ch) < 32:
            continue
        elif ord(ch) < 128:
            out.append(ch)
        else:
            try:
                for b in ch.encode("gbk"):
                    out.append("\\'%02x" % b)
            except UnicodeEncodeError:
                out.append("?")
    return "".join(out)


def para(text, align="ql", bold=False, size=24, after=60, before=0):
    fmt = "\\pard\\plain\\lang2052\\f1\\fs%d\\sa%d" % (size, after)
    if before:
        fmt += "\\sb%d" % before
    fmt += "\\%s " % align
    if bold:
        fmt += "\\b "
    body = rtf_escape(text)
    if bold:
        body += "\\b0 "
    return fmt + body + "\\par "


def build_doc(title, subtitle, sections, answers, note):
    """sections: [(标题, [题目...]), ...]；answers: [答案行...]"""
    rtf = []
    rtf.append("{\\rtf1\\ansi\\ansicpg936\\deff0")
    rtf.append(
        "{\\fonttbl{\\f0\\froman\\fcharset0 Times New Roman;}"
        "{\\f1\\fnil\\fcharset134 " + rtf_escape("宋体") + ";}"
        "{\\f2\\fnil\\fcharset134 " + rtf_escape("黑体") + ";}}"
    )
    # 关键：\colortbl 必须加花括号，否则 Word 里整个文档显示为空白
    rtf.append("{\\colortbl;\\red0\\green0\\blue0;}")
    rtf.append("\\viewkind4\\uc1\\pard\\lang2052\\f1\\fs24 ")
    rtf.append(para(title, align="qc", bold=True, size=36, after=80))
    rtf.append(para(subtitle, align="qc", size=24, after=80))
    for sec_title, items in sections:
        rtf.append(para(sec_title, bold=True, size=26, before=160, after=80))
        for i, q in enumerate(items, 1):
            rtf.append(para("%d. %s" % (i, q), after=100))
    rtf.append("\\page ")
    rtf.append(para("参考答案（供家长核对）", align="qc", bold=True, size=30, after=100))
    for line in answers:
        rtf.append(para(line, after=60))
    rtf.append(para(note, size=20, after=0, before=160))
    body = "".join(rtf) + "}"
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(OUT_FILE, "wb") as f:
        f.write(body.encode("ascii"))
    print("Wrote", OUT_FILE)
```

### 2.3 编写时的坑

- **`\colortbl` 必须写成 `{\colortbl;...}`**（本次实际踩过的坑：漏了花括号，Word 打开全空白，自写解析器却正常）。
- 同理，`\fonttbl`、`\stylesheet` 等目的地控制字都要在花括号内。
- **不要二次转义**：字符串里先放真实制表符 `\t` 和空格，由 `rtf_escape` 统一转一次；不要先拼好 `\tab`、`\'xx` 再整体转义，否则出现 `\\tab`、`\\'a1` 这类字面文本。
- 生成后用 Word COM 验证（见 §3），不要只信自写解析器。
- 所有计算答案生成前用断言程序化校验（口算、竖式、递等式、应用题答案逐一 `assert`）。

---

## 3. 验证方法（Word COM 自动化）

Word 已安装时，用隐藏实例实际打开文件，检查字符/段落/页数并搜索关键内容，能立刻发现"空白文档"类问题：

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open('C:\...\math_20260818.doc', $false, $true)
    "chars=$($doc.Characters.Count) paragraphs=$($doc.Paragraphs.Count) pages=$($doc.ComputeStatistics(2))"
    $txt = $doc.Content.Text
    $txt.Substring(0, [Math]::Min(200, $txt.Length))   # 看开头
    $i = $txt.IndexOf('参考答案')                        # 找答案页
    $doc.Close($false)
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
```

正常特征：`chars` 数千、`paragraphs` 百余、`pages` 数页；异常特征：`chars=1 paragraphs=1`（空白，多半是目的地控制字没加花括号）。

另可把生成的 .doc 再喂给 §1 的提取脚本做一次"回读"，人工核对题目与答案是否完整。

---

## 4. 工具与命令速查

| 用途 | 命令/工具 |
| --- | --- |
| 看文件是否 RTF | PowerShell 读前 16 字节，应见 `{\rtf1`；或 `Get-ChildItem` 看体积 |
| 提取正文 | `python _rtf_extract.py <目录> [前缀]`（脚本见 §1.2） |
| 生成新卷 | Python 脚本（函数见 §2.2），输出 `三年级/math_YYYYMMDD.doc` |
| 程序化校验答案 | 生成脚本里对每个数值写 `assert`，运行通过再生成 |
| Word 实开验证 | PowerShell COM（见 §3） |
| 创建/修改/删除文件 | **apply_patch**（UTF-8 安全；`Remove-Item` 可能被沙箱策略拦截，删除用 apply_patch 的 Delete File） |
| 查看 UTF-8 文本 | `Get-Content -Raw -Encoding UTF8` |

Python 为嵌入式发行版（`C:\software\python-3.14.3-embed-amd64`），无 striprtf/pandoc/python-docx，且环境不可联网安装；按本文档手写解析/生成即可。

---

## 5. 完整工作流（下次照做）

1. `_rtf_extract.py` 提取原卷 → 通读全文，特别注意末尾"测试结果/错题"记录。
2. 归纳错题对应的知识点（例：分数比较、周长与面积、质量单位、对折翻倍）。
3. 设计新卷：覆盖三年级全册各知识点，对错题类型设置"强化题"；难度略增、题量增多；不限定时间/分数（用户此前要求）。
4. 生成脚本中为每个计算答案写 `assert`，全部通过后再写 RTF。
5. 生成 `三年级/math_YYYYMMDD.doc`（卷末附"参考答案（供家长核对）"和给家长的"注"，说明错题强化位置）。
6. 用 Word COM 实开验证非空白，再回读提取文本核对完整性。
7. 用 apply_patch 删除临时脚本，保持目录干净（只留试卷文件和本目录文档）。
