#!/usr/bin/env python3
"""
发送邮件脚本 - 由 OpenClaw 调用
用法: python3 send_mail.py <收件人> <主题> <内容文件路径>
"""
import smtplib
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 163 邮箱配置
SMTP_HOST = "smtp.163.com"
SMTP_PORT = 465
SENDER = "skygentwu@163.com"
PASSWORD = "MZMxq33qdbxbR7TH"

def send_mail(to_addr, subject, body):
    msg = MIMEMultipart()
    msg["From"] = SENDER
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(SENDER, PASSWORD)
        server.sendmail(SENDER, [to_addr], msg.as_string())
    print(f"✅ 邮件已发送至 {to_addr}")

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("用法: python3 send_mail.py <收件人> <主题> <内容文件>")
        sys.exit(1)
    
    to_addr = sys.argv[1]
    subject = sys.argv[2]
    with open(sys.argv[3], "r", encoding="utf-8") as f:
        body = f.read()
    
    send_mail(to_addr, subject, body)
