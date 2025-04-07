"""
微信企业邮箱插件（MoviePilot V2 规范版）
功能：监控企业邮箱并解析邮件内容执行自动化操作
"""
from typing import Any, Dict, Optional
from email.header import decode_header
from email.message import Message
import imaplib
import email
import re

from app.core.config import settings
from app.core.context import Context
from app.core.event import eventmanager, Event
from app.log import logger
from app.modules.emby import Emby
from app.schemas.types import EventType
from app.plugins._base import BasePlugin

class EnterpriseEmail(BasePlugin):
    # 插件名称
    plugin_name = "企业邮箱助手"
    # 插件描述
    plugin_desc = "监控企业邮箱并执行自动化操作"
    # 插件图标
    plugin_icon = "qymail.png"
    # 插件版本
    plugin_version = "1.1"
    # 插件作者
    plugin_author = "时也命也"
    plugin_settings = {
        "enable": True,
        "auto_start": True
    }

    # 配置模型
    class ConfigModel(BasePlugin.ConfigModel):
        email_server: str = "imap.exmail.qq.com"
        email_port: int = 993
        email_user: str
        email_password: str
        check_interval: int = 300  # 检查间隔（秒）
        command_prefix: str = "/movie"  # 指令前缀

    def __init__(self):
        super().__init__()
        self._imap = None
        self._running = False

    def init(self):
        """插件初始化"""
        if self._config.email_user and self._config.email_password:
            self._imap = imaplib.IMAP4_SSL(
                self._config.email_server,
                self._config.email_port
            )
            try:
                self._imap.login(
                    self._config.email_user,
                    self._config.email_password
                )
                logger.info("企业邮箱登录成功")
            except Exception as e:
                logger.error(f"邮箱登录失败: {str(e)}")
                self._imap = None

    def start(self):
        """启动插件"""
        if not self._running and self._imap:
            self._running = True
            self.check_email_job()

    def stop(self):
        """停止插件"""
        self._running = False
        if self._imap:
            try:
                self._imap.close()
                self._imap.logout()
            except Exception:
                pass

    @eventmanager.register(EventType.PluginReload)
    def reload(self, event: Event):
        """响应插件重载事件"""
        self.stop()
        self.init()
        self.start()

    def check_email_job(self):
        """定时检查邮件任务"""
        if not self._running:
            return

        try:
            self._imap.select("INBOX")
            status, messages = self._imap.search(None, "UNSEEN")
            if status != "OK":
                return

            for num in messages[0].split():
                self.process_email(num)
                
        except Exception as e:
            logger.error(f"检查邮件失败: {str(e)}")
        finally:
            # 定时循环
            self._scheduler.add_job(
                self.check_email_job,
                'interval',
                seconds=self._config.check_interval,
                id="email_check"
            )

    def process_email(self, email_num: str):
        """处理单封邮件"""
        try:
            status, data = self._imap.fetch(email_num, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # 解析邮件信息
            subject = self._decode_header(msg.get("Subject", ""))
            from_ = self._decode_header(msg.get("From", ""))
            content = self._get_text_content(msg)

            logger.info(f"收到新邮件 - 发件人: {from_}, 主题: {subject}")

            # 执行命令处理
            if subject.startswith(self._config.command_prefix):
                self._handle_command(from_, content)

            # 标记为已读
            self._imap.store(email_num, "+FLAGS", "\\Seen")

        except Exception as e:
            logger.error(f"处理邮件失败: {str(e)}")

    def _handle_command(self, sender: str, command: str):
        """处理邮件指令"""
        # 权限验证（示例）
        if not self._is_authorized(sender):
            logger.warning(f"未授权用户尝试执行命令: {sender}")
            return

        # 解析命令参数
        match = re.match(r"/(\w+)\s+(.+)", command)
        if not match:
            return

        cmd_type, params = match.groups()
        
        # 执行对应操作
        if cmd_type == "search":
            # 调用媒体库搜索
            results = Emby().search_media(params)
            # TODO: 发送结果邮件
            logger.info(f"执行搜索命令: {params}")

        elif cmd_type == "download":
            # 触发下载任务
            self.event_manager.send_event(
                EventType.DownloadAdd,
                {
                    "url": params,
                    "user": sender
                }
            )
            logger.info(f"触发下载任务: {params}")

    def _is_authorized(self, sender: str) -> bool:
        """验证发件人权限"""
        return sender in settings.SUPERUSERS

    def _decode_header(self, header: str) -> str:
        """解码邮件头"""
        decoded = decode_header(header)
        return str(decoded[0][0], decoded[0][1] or "utf-8")

    def _get_text_content(self, msg: Message) -> str:
        """提取纯文本内容"""
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode("utf-8")
        return ""

    def get_state(self) -> bool:
        """获取运行状态"""
        return self._running

    @staticmethod
    def get_command() -> Dict[str, Any]:
        """暴露服务接口（示例）"""
        return {
            "cmd": "/email",
            "desc": "邮件服务接口",
            "params": {}
        }
