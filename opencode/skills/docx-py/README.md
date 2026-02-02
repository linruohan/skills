# Python-Docx 快速创建 Word

使用 python-docx 库快速创建 Word 文件。

## 前置条件

安装依赖库：

```bash
pip install python-docx
```

## 快速创建 Word

### 场景 1：创建空的 Word 文件（单行命令）

直接生成无任何内容的 .docx 文件：

```bash
python -c "from docx import Document; Document().save('empty.docx')"
```

### 场景 2：创建带文字内容的 Word（最常用，单行命令）

生成后直接包含指定文字：

```bash
python -c "from docx import Document; doc=Document(); doc.add_paragraph('这是Python单行命令创建的Word内容'); doc.save('content.docx')"
```

### 场景 3：从文件内容创建 Word

将文本文件内容写入 Word：

```python
from docx import Document
import pathlib

doc = Document()
content = pathlib.Path('input.txt').read_text(encoding='utf-8')
doc.add_paragraph(content)
doc.save('output.docx')
```

## 关键参数说明

- `Document()`：创建一个空的 Word 文档对象
- `add_paragraph('文字内容')`：添加段落内容
- `save('文件名.docx')`：保存文档为 .docx 格式
- 仅支持 .docx 格式，不支持旧版 .doc 格式

## 总结

- 核心逻辑：`Document()` -> `add_paragraph()` -> `save()` 一步完成 Word 创建
- 终端执行加 `python -c "包裹代码"`
- 必装 python-docx，是最简洁的 Python 创建 Word 的方法
