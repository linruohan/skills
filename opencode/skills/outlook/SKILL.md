---
name: outlook
description: 通过本地 Outlook 客户端发送和读取邮件。使用此功能需要 Windows 系统并安装 Outlook 和 pywin32 库（pip install pywin32）。支持发送带附件的邮件，以及读取收件箱、已发送等文件夹中的邮件。
license: Complete terms in LICENSE.txt
---

# Outlook 邮件处理

此 skill 提供了通过本地 Outlook 客户端发送和读取邮件的功能。

## 前置条件

- **操作系统**: Windows（Outlook COM 接口仅支持 Windows）
- **软件要求**: 已安装 Microsoft Outlook 客户端并配置邮箱账户
- **Python 依赖**: `pip install pywin32 pandas openpyxl`

## 功能概述

### 发送邮件 `send_outlook_mail`

通过本地 Outlook 客户端发送邮件。

#### 参数说明

```python
send_outlook_mail(
    to_addrs,          # 收件人，多个用分号分隔："a@outlook.com;b@163.com"
    subject,           # 邮件主题
    body,              # 邮件正文（纯文本，支持换行符\n）
    cc_addrs="",       # 抄送人，多个用分号分隔，无则留空
    bcc_addrs="",      # 密送人，多个用分号分隔，无则留空
    attachments=None,  # 附件路径，单个字符串或列表：r"C:\test.xlsx" 或 [r"C:\1.txt", r"D:\2.pdf"]
)
```

#### 使用示例

```python
# 发送简单邮件
send_outlook_mail(
    to_addrs="user@example.com",
    subject="测试邮件",
    body="这是一封测试邮件"
)

# 发送带附件和多收件人的邮件
send_outlook_mail(
    to_addrs="user1@example.com;user2@example.com",
    subject="项目报告",
    body="请查收附件中的项目报告",
    cc_addrs="manager@example.com",
    attachments=[r"C:\report.pdf", r"C:\data.xlsx"]
)
```

### 读取邮件 `read_outlook_mails`

从 Outlook 文件夹中读取邮件列表。

#### 参数说明

```python
read_outlook_mails(
    folder_name="收件箱",    # 要读取的文件夹（收件箱/已发送/草稿等）
    read_count=10,           # 读取最新的N封邮件，0表示读取全部
    filter_unread=False      # 是否仅读取未读邮件
)
```

#### 返回值

返回邮件列表，每个邮件为字典格式，包含以下字段：
- `序号`: 邮件编号
- `发件人`: 发件人名称
- `发件人邮箱`: 发件人邮箱地址
- `邮件主题`: 邮件主题
- `发送时间`: 发送时间（格式：YYYY-MM-DD HH:MM:SS）
- `是否未读`: 布尔值，True 表示未读
- `邮件正文（纯文本）`: 邮件正文内容（前500字）
- `是否有附件`: 布尔值，True 表示有附件

#### 使用示例

```python
# 读取收件箱最新10封邮件
mails = read_outlook_mails(folder_name="收件箱", read_count=10)

# 仅读取未读邮件
mails = read_outlook_mails(folder_name="收件箱", filter_unread=True)

# 读取已发送文件夹中的全部邮件
mails = read_outlook_mails(folder_name="已发送", read_count=0)

# 遍历打印邮件信息
for mail in mails:
    print(f"主题: {mail['邮件主题']}")
    print(f"发件人: {mail['发件人']} ({mail['发件人邮箱']})")
    print(f"时间: {mail['发送时间']}")
    print(f"是否未读: {mail['是否未读']}")
    print("-" * 40)
```

### 读取邮件并导出到Excel `read_mails_to_excel`

从 Outlook 文件夹中读取邮件列表并导出为Excel文件。

#### 参数说明

```python
read_mails_to_excel(
    folder_name="收件箱",           # 要读取的文件夹（收件箱/已发送/草稿等）
    read_count=10,                  # 读取最新的N封邮件，0表示读取全部
    filter_unread=False,            # 是否仅读取未读邮件
    output_path=None                # 输出Excel文件路径，默认为桌面
)
```

#### 返回值

返回两个值：
1. 邮件列表（字典格式，同 `read_outlook_mails`）
2. Excel文件保存路径

#### 使用示例

```python
# 读取收件箱最新10封邮件并导出到Excel
mails, file_path = read_mails_to_excel(folder_name="收件箱", read_count=10)
print(f"已保存到: {file_path}")

# 仅读取未读邮件并导出到指定路径
mails, file_path = read_mails_to_excel(
    folder_name="收件箱",
    filter_unread=True,
    output_path=r"C:\Reports\emails.xlsx"
)

# 读取已发送文件夹中的全部邮件并导出
mails, file_path = read_mails_to_excel(
    folder_name="已发送",
    read_count=0
)
```

## 支持的文件夹名称

Outlook 文件夹使用中文命名，支持以下常见文件夹：
- `收件箱` - Inbox
- `已发送` - Sent Items
- `草稿` - Drafts
- `已删除` - Deleted Items
- `垃圾邮件` - Junk Email
- `归档` - Archive
- 自定义文件夹名称

## 注意事项

1. **字符编码**: 此 skill 已配置 UTF-8 编码处理，可正确显示中文和特殊字符
2. **Outlook 必须运行**: 发送邮件时 Outlook 客户端需要处于运行状态
3. **附件路径**: 附件路径必须是绝对路径，且文件必须存在
4. **正文长度**: 读取邮件时，正文默认截取前500字符以避免输出过长
5. **邮件数量**: 读取大量邮件可能会影响性能，建议合理设置 `read_count` 参数

## 错误处理

常见错误及解决方案：

- `pywintypes.com_error`: Outlook 未正确安装或未配置账户
- `FileNotFoundError`: 附件文件路径不存在
- `UnicodeEncodeError`: 字符编码问题（已通过 UTF-8 配置解决）

## 完整示例

```python
# 发送邮件并确认收件箱中的反馈
send_outlook_mail(
    to_addrs="recipient@example.com",
    subject="自动报告",
    body="这是由脚本自动发送的报告\n\n请查收附件。",
    attachments=[r"C:\reports\report.xlsx"]
)

# 检查收件箱中的新邮件
new_mails = read_outlook_mails(folder_name="收件箱", read_count=5, filter_unread=True)
print(f"发现 {len(new_mails)} 封未读邮件")

# 批量导出未读邮件到Excel
import time
time.sleep(10)  # 等待邮件到达
mails, excel_file = read_mails_to_excel(
    folder_name="收件箱",
    read_count=10,
    filter_unread=True
)
print(f"已导出 {len(mails)} 封邮件到: {excel_file}")
```
