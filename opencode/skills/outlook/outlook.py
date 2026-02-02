# -*- coding: utf-8 -*-
import os
import sys
from datetime import datetime

import pandas as pd
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")


def send_outlook_mail(to_addrs, subject, body, cc_addrs="", bcc_addrs="", attachments=None):
    """
    调用本地Outlook客户端发送邮件
    :param to_addrs: 收件人，多个用分号分隔："a@outlook.com;b@163.com"
    :param subject: 邮件主题
    :param body: 邮件正文（纯文本，支持换行符\n）
    :param cc_addrs: 抄送人，多个用分号分隔，无则留空
    :param bcc_addrs: 密送人，多个用分号分隔，无则留空
    :param attachments: 附件路径，单个字符串或列表：r"C:\test.xlsx" 或 [r"C:\1.txt", r"D:\2.pdf"]
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = to_addrs
    mail.CC = cc_addrs
    mail.BCC = bcc_addrs
    mail.Subject = subject
    mail.Body = body

    if attachments:
        att_list = [attachments] if isinstance(attachments, str) else attachments
        for att_path in att_list:
            if os.path.exists(att_path):
                mail.Attachments.Add(Source=att_path)
            else:
                print(f"警告：附件{att_path}不存在，已跳过")

    mail.Send()

    print(f"邮件【{subject}】发送成功！")


def read_outlook_mails(folder_name="收件箱", read_count=10, filter_unread=False):
    """
    读取Outlook邮件列表
    :param folder_name: 要读取的文件夹（收件箱/已发送/草稿等，中文直接填）
    :param read_count: 读取最新的N封邮件，0表示读取全部
    :param filter_unread: 是否仅读取未读邮件
    :return: 邮件列表（字典格式，含核心信息）
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.Folders(1).Folders(folder_name)
    mails = inbox.Items
    mails.Sort("[SentOn]", True)

    if filter_unread:
        mails = mails.Restrict("[UnRead] = True")

    mail_list = []
    read_num = read_count if read_count > 0 else len(mails)
    for i, mail in enumerate(mails):
        if i >= read_num:
            break
        mail_info = {
            "序号": i + 1,
            "发件人": mail.SenderName if hasattr(mail, "SenderName") else "未知发件人",
            "发件人邮箱": mail.SenderEmailAddress if hasattr(mail, "SenderEmailAddress") else "未知",
            "邮件主题": mail.Subject,
            "发送时间": mail.SentOn.strftime("%Y-%m-%d %H:%M:%S") if mail.SentOn else "未知时间",
            "是否未读": mail.UnRead,
            "邮件正文（纯文本）": mail.Body[:500] + "..." if len(mail.Body) > 500 else mail.Body,
            "是否有附件": mail.Attachments.Count > 0,
        }
        mail_list.append(mail_info)
    return mail_list


def read_mails_to_excel(folder_name="收件箱", read_count=10, filter_unread=False, output_path=None):
    """
    读取Outlook邮件列表并导出为Excel文件
    :param folder_name: 要读取的文件夹（收件箱/已发送/草稿等，中文直接填）
    :param read_count: 读取最新的N封邮件，0表示读取全部
    :param filter_unread: 是否仅读取未读邮件
    :param output_path: 输出Excel文件路径，默认为桌面
    :return: (邮件列表, Excel文件保存路径)
    """
    # 读取邮件
    mail_list = read_outlook_mails(folder_name, read_count, filter_unread)

    # 设置输出路径
    if output_path is None:
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(desktop_path, f"未读邮件_{timestamp}.xlsx")

    # 导出为Excel
    df = pd.DataFrame(mail_list)
    df.to_excel(output_path, index=False, engine="openpyxl")

    print(f"已保存 {len(mail_list)} 封邮件到 {output_path}")

    return mail_list, output_path
