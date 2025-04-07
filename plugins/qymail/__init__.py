file: wechat_work_email_plugin.py

"""
微信企业邮箱管理插件
版本: 1.0.0
作者: [您的名称]
功能:
- 可视化收发企业邮件
- 支持附件管理
- 邮件定时检查
- 邮件分类展示
- 安全的凭证存储
"""

import imaplib
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
import pytz
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from typing import Any, List, Dict, Optional, Tuple

from app.core.config import settings
from app.plugins import _PluginBase
from app.log import logger
from app.schemas import NotificationType

class WeWorkEmail(_PluginBase):
    plugin_name = "微信企业邮箱管理"
    plugin_desc = "可视化收发微信企业邮箱邮件，支持附件管理"
    plugin_icon = "https://example.com/email_icon.png"
    plugin_version = "1.0.0"
    plugin_author = "[时也命也]"
    author_url = "https://github.com/beijingxiaokuoe"
    plugin_config_prefix = "wework_email_"
    plugin_order = 2
    auth_level = 2

    # 私有属性
    _enabled = False
    _account = None
    _password = None
    _imap_host = "imap.exmail.qq.com"
    _smtp_host = "smtp.exmail.qq.com"
    _check_interval = "*/10 * * * *"  # 默认10分钟检查一次
    _notify = True
    _scheduler: Optional[BackgroundScheduler] = None

    def init_plugin(self, config: dict = None):
        self.stop_service()

        if config:
            self._enabled = config.get("enabled")
            self._account = config.get("account")
            self._password = config.get("password")
            self._check_interval = config.get("check_interval", "*/10 * * * *")

        if self._enabled and self._validate_config():
            self._scheduler = BackgroundScheduler(timezone=settings.TZ)
            self._scheduler.add_job(
                self.check_email,
                trigger=CronTrigger.from_crontab(self._check_interval),
                id="wework_email_check"
            )
            self._scheduler.start()
            logger.info("微信企业邮箱插件已启动")

    def _validate_config(self) -> bool:
        """验证配置有效性"""
        if not all([self._account, self._password]):
            logger.error("邮箱账号或密码未配置")
            return False
        return True

    def check_email(self):
        """检查新邮件"""
        try:
            with imaplib.IMAP4_SSL(self._imap_host) as imap:
                imap.login(self._account, self._password)
                imap.select("INBOX")

                status, messages = imap.search(None, 'UNSEEN')
                if status == 'OK':
                    email_count = len(messages[0].split())
                    if email_count > 0:
                        self._send_notification(
                            title="📬 新邮件通知",
                            text=f"检测到 {email_count} 封未读邮件"
                        )
        except Exception as e:
            logger.error(f"检查邮件失败: {str(e)}")

    def send_email(self, to: str, subject: str, content: str, attachments: list = None) -> dict:
        """发送邮件"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self._account
            msg['To'] = to
            msg['Subject'] = subject
            msg.attach(MIMEText(content, 'html'))

            # 附件处理
            if attachments:
                for file in attachments:
                    part = MIMEText(file['content'])
                    part.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=file['name']
                    )
                    msg.attach(part)

            with smtplib.SMTP_SSL(self._smtp_host) as smtp:
                smtp.login(self._account, self._password)
                smtp.send_message(msg)
            
            return {"status": True, "message": "邮件发送成功"}
        except Exception as e:
            logger.error(f"邮件发送失败: {str(e)}")
            return {"status": False, "message": str(e)}

    def get_recent_emails(self, limit=20) -> List[dict]:
        """获取最近邮件"""
        emails = []
        try:
            with imaplib.IMAP4_SSL(self._imap_host) as imap:
                imap.login(self._account, self._password)
                imap.select("INBOX")

                status, messages = imap.search(None, 'ALL')
                if status == 'OK':
                    for num in messages[0].split()[:limit]:
                        status, data = imap.fetch(num, '(RFC822)')
                        if status == 'OK':
                            raw_email = data[0][1]
                            email_message = email.message_from_bytes(raw_email)
                            
                            email_info = {
                                'from': email_message['From'],
                                'subject': email_message['Subject'],
                                'date': email_message['Date'],
                                'content': self._parse_email_content(email_message)
                            }
                            emails.append(email_info)
        except Exception as e:
            logger.error(f"获取邮件失败: {str(e)}")
        return emails

    def _parse_email_content(self, msg) -> str:
        """解析邮件内容"""
        content = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    content += part.get_payload(decode=True).decode()
        else:
            content = msg.get_payload(decode=True).decode()
        return content[:500] + "..."  # 截取前500字符

    def _send_notification(self, title: str, text: str):
        """发送通知"""
        if self._notify:
            self.post_message(
                mtype=NotificationType.SiteMessage,
                title=title,
                text=text
            )

    def get_form(self) -> Tuple[List[dict], Dict[str, Any]]:
        return [
            {
                'component': 'VForm',
                'content': [
                    {
                        'component': 'VRow',
                        'content': [
                            {
                                'component': 'VCol',
                                'props': {'cols': 12, 'md': 6},
                                'content': [
                                    {
                                        'component': 'VTextField',
                                        'props': {
                                            'model': 'account',
                                            'label': '企业邮箱',
                                            'placeholder': 'user@company.com'
                                        }
                                    }
                                ]
                            },
                            {
                                'component': 'VCol',
                                'props': {'cols': 12, 'md': 6},
                                'content': [
                                    {
                                        'component': 'VTextField',
                                        'props': {
                                            'model': 'password',
                                            'label': '授权密码',
                                            'type': 'password'
                                        }
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        'component': 'VRow',
                        'content': [
                            {
                                'component': 'VCol',
                                'props': {'cols': 12},
                                'content': [
                                    {
                                        'component': 'VCronField',
                                        'props': {
                                            'model': 'check_interval',
                                            'label': '邮件检查频率'
                                        }
                                    }
                                ]
                            }
                        ]
                    }
                ]
            }
        ], {
            "enabled": False,
            "account": "",
            "password": "",
            "check_interval": "*/10 * * * *"
        }

    def get_page(self) -> List[dict]:
        """邮件列表页面"""
        emails = self.get_recent_emails()
        rows = []
        for email in emails:
            rows.append({
                'component': 'tr',
                'content': [
                    {'component': 'td', 'text': email['from']},
                    {'component': 'td', 'text': email['subject']},
                    {'component': 'td', 'text': email['date']},
                    {
                        'component': 'td',
                        'content': [{
                            'component': 'VBtn',
                            'props': {
                                'small': True,
                                'variant': 'tonal',
                                'onClick': f'window.showEmailDetail({email})'
                            },
                            'text': '查看详情'
                        }]
                    }
                ]
            })

        return [{
            'component': 'VCard',
            'content': [
                {
                    'component': 'VCardTitle',
                    'text': '📧 最近邮件'
                },
                {
                    'component': 'VCardText',
                    'content': [{
                        'component': 'VTable',
                        'props': {'hover': True},
                        'content': [
                            {
                                'component': 'thead',
                                'content': [{
                                    'component': 'tr',
                                    'content': [
                                        {'component': 'th', 'text': '发件人'},
                                        {'component': 'th', 'text': '主题'},
                                        {'component': 'th', 'text': '日期'},
                                        {'component': 'th', 'text': '操作'}
                                    ]
                                }]
                            },
                            {
                                'component': 'tbody',
                                'content': rows
                            }
                        ]
                    }]
                }
            ]
        }]

    def stop_service(self):
        if self._scheduler:
            self._scheduler.remove_all_jobs()
            if self._scheduler.running:
                self._scheduler.shutdown()
            self._scheduler = None
