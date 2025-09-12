#!/usr/bin/env python3
"""
Oracle Quick Open (OQO) Tool
功能：快速打开 Oracle SR/KM文档/Bug/GRP/OLUEK/URL链接，并支持Google搜索
作者：v0279158888
日期：2018-11-28
最近更新：2025-08-19 代码重构和优化
"""

import re
import os
import sys
import datetime
import platform
from urllib import parse
import subprocess
import logging
import json
import time
import argparse
from abc import ABC, abstractmethod
from typing import Optional, Dict, List, Any, Set
from dataclasses import dataclass
from pathlib import Path

# 配置常量
@dataclass
class Config:
    """配置常量类"""
    # URL前缀
    SR_PREFIX = 'https://support.us.oracle.com/oip/faces/secure/srm/srview/SRTechnical.jspx?srNumber='
    DOC_PREFIX = 'https://mosemp.us.oracle.com/epmos/faces/DocContentDisplay?id='
    BUG_PREFIX = 'https://bug.oraclecorp.com/pls/bug/webbug_edit.edit_info_top?rptno='
    JIRA_PREFIX = 'https://jira-sd.mc1.oracleiaas.com/browse/'
    PEOPLE_PREFIX = 'https://people.oracle.com/'
    CMOS_PREFIX = 'https://fa-etmi-saasfaprod1.fa.ocs.oraclecloud.com/fscmUI/faces/deeplink?objType=SVC_SERVICE_REQUEST&objKey=srNumber='
    GRP_PREFIX = 'https://support.us.oracle.com/oip/faces/secure/grp/sch/schedule.jspx?mc=true&oracleEmail='
    
    # 正则表达式模式
    PATTERNS = {
        'sr': r"3-\d{9,12}",
        'cmos': r"4-\d{9,12}",
        'doc': r'(?:(?<=^)|(?<=\s)|(?<=\())\d{1,8}\.\d(?=\s|$|\))',
        'email': r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+',
        'url': r"""(?:^|\s)\b[a-zA-Z]{3,8}://[^\s'"]+[\w=/](?=\s|$)""",
        'bug': r'(?:^|\s)\b\d{7,9}(?=\s|$)',
        'jira': r'(?:(?<=^)|(?<=\s)|(?<=\/))OLUEK-\d{4,5}(?=\s|$)',
        'people': r"(?:^|\s)@[a-z]{3,20}(?=\s|$)"
    }
    
    # 允许的域名
    ALLOWED_DOMAINS = [
        'oracle.com',
        'oraclecloud.com',
        'oracleiaas.com',
        'oraclecorp.com',
        'google.com'
    ]

class PlatformInterface(ABC):
    """平台接口抽象类"""
    
    @abstractmethod
    def get_selected_text(self) -> str:
        """获取选中的文本"""
        pass
    
    @abstractmethod
    def set_clipboard(self, text: str) -> None:
        """设置剪贴板内容"""
        pass
    
    @abstractmethod
    def send_notification(self, message: str) -> None:
        """发送通知"""
        pass
    
    @abstractmethod
    def open_browser(self, url: str) -> bool:
        """打开浏览器"""
        pass

class MacOSInterface(PlatformInterface):
    """macOS平台实现"""
    
    def get_selected_text(self) -> str:
        """获取macOS选中的文本"""
        try:
            # 获取前台应用
            front_app = self._get_front_app()
            
            if front_app == "Google Chrome":
                return self._get_chrome_selection()
            else:
                return self._get_app_selection(front_app)
        except Exception as e:
            logging.error(f"macOS获取选中文本失败: {str(e)}")
            return ""
    
    def _get_front_app(self) -> str:
        """获取前台应用名称"""
        script = '''
            tell application "System Events"
                set frontApp to name of first application process whose frontmost is true
            end tell
        '''
        result = subprocess.run(['osascript', '-e', script], 
                              capture_output=True, text=True, timeout=3)
        return result.stdout.strip()
    
    def _get_chrome_selection(self) -> str:
        """从Chrome获取选中的文本"""
        script = '''
            tell application "Google Chrome"
                try
                    tell active tab of front window
                        execute javascript "window.getSelection().toString()"
                    end tell
                on error errMsg
                    return "ERROR: " & errMsg
                end try
            end tell
        '''
        result = subprocess.run(['osascript', '-e', script],
                             capture_output=True, text=True, timeout=3)
        selected_text = result.stdout.strip()
        
        if selected_text.startswith("ERROR:"):
            logging.error(f"Chrome脚本错误: {selected_text}")
            return ""
        return selected_text
    
    def _get_app_selection(self, app_name: str) -> str:
        """从其他应用获取选中的文本"""
        script = f'''
            tell application "{app_name}"
                try
                    set selectedText to ""
                    tell application "System Events"
                        keystroke "c" using command down
                    end tell
                    delay 0.1
                    set selectedText to (do shell script "pbpaste")
                    return selectedText
                on error
                    return ""
                end try
            end tell
        '''
        result = subprocess.run(['osascript', '-e', script], 
                             capture_output=True, text=True)
        selected_text = result.stdout.strip()
        
        if selected_text:
            return selected_text
        
        # 回退到剪贴板内容
        return self._get_clipboard_content()
    
    def _get_clipboard_content(self) -> str:
        """获取剪贴板内容"""
        result = subprocess.run(['pbpaste'], capture_output=True, text=True)
        return result.stdout.strip()
    
    def set_clipboard(self, text: str) -> None:
        """设置剪贴板内容"""
        subprocess.Popen(["pbcopy"], stdin=subprocess.PIPE).communicate(text.encode('utf-8'))
    
    def send_notification(self, message: str) -> None:
        """发送macOS通知"""
        script = f'''display notification "{message}" with title "OQO Script"'''
        os.system(f"osascript -e '{script}'")
    
    def open_browser(self, url: str) -> bool:
        """在macOS上打开浏览器"""
        try:
            subprocess.run(['/usr/bin/open', '-a', 'Google Chrome', url], 
                         check=True, timeout=5)
            return True
        except subprocess.CalledProcessError:
            return False

class LinuxInterface(PlatformInterface):
    """Linux平台实现"""
    
    def get_selected_text(self) -> str:
        """获取Linux选中的文本"""
        try:
            # 首先尝试xsel
            result = subprocess.run(['xsel', '--primary'], 
                                 capture_output=True, text=True)
            if result.stdout.strip():
                return result.stdout.strip()
            
            # 回退到xclip
            result = subprocess.run(['xclip', '-o', '-selection', 'primary'],
                                 capture_output=True, text=True)
            return result.stdout.strip()
        except Exception as e:
            logging.error(f"Linux获取选中文本失败: {str(e)}")
            return ""
    
    def set_clipboard(self, text: str) -> None:
        """设置剪贴板内容"""
        subprocess.Popen(["xclip", "-i"], stdin=subprocess.PIPE).communicate(text.encode('utf-8'))
    
    def send_notification(self, message: str) -> None:
        """发送Linux通知"""
        os.system(f"notify-send -t 1000 '{message}'")
    
    def open_browser(self, url: str) -> bool:
        """在Linux上打开浏览器"""
        try:
            subprocess.run(['/usr/bin/google-chrome', '--proxy-auto-detect', url], 
                         check=True, timeout=5)
            return True
        except subprocess.CalledProcessError:
            return False

class URLProcessor:
    """URL处理器"""
    
    def __init__(self, config: Config):
        self.config = config
        self._initialize_date_range()
    
    def _initialize_date_range(self):
        """初始化日期范围"""
        now = datetime.datetime.now()
        self.date_key1 = now.strftime("%Y-%m-%d")
        self.date_key2 = (now + datetime.timedelta(days=7)).strftime("%Y-%m-%d")
    
    def extract_items(self, text: str) -> Dict[str, List[str]]:
        """从文本中提取各种项目"""
        items = {}
        for key, pattern in self.config.PATTERNS.items():
            items[key] = re.findall(pattern, text)
        return items
    
    def generate_urls(self, items: Dict[str, List[str]]) -> List[str]:
        """生成URL列表"""
        urls = []
        
        # SR URLs (Service Request)
        for sr in items.get('sr', []):
            urls.append(f"{self.config.SR_PREFIX}{sr.strip()}")
        
        # CMOS URLs (Cloud My Oracle Support)
        for cmos in items.get('cmos', []):
            urls.append(f"{self.config.CMOS_PREFIX}{cmos.strip()}&action=EDIT_IN_TAB")
        
        # Document URLs (Knowledge Base)
        for doc in items.get('doc', []):
            urls.append(f"{self.config.DOC_PREFIX}{doc.strip()}")
        
        # Bug URLs
        for bug in items.get('bug', []):
            urls.append(f"{self.config.BUG_PREFIX}{bug.strip()}")
        
        # JIRA URLs
        for jira in items.get('jira', []):
            urls.append(f"{self.config.JIRA_PREFIX}{jira.strip()}")
        
        # People URLs
        for people in items.get('people', []):
            #people_id = people.strip().lstrip('@')
            people_id = people.strip()
            urls.append(f"{self.config.PEOPLE_PREFIX}{people_id}")
        
        # Email GRP URLs (Group Schedule)
        for email in items.get('email', []):
            grp_url = (f"{self.config.GRP_PREFIX}"
                      f"{email.strip().upper()}"
                      f"&fromDate={self.date_key1}&toDate={self.date_key2}")
            urls.append(grp_url)
        
        # 直接URLs (已经存在的完整URL)
        for url in items.get('url', []):
            # 清理和验证URL
            clean_url = url.strip()
            if self.validate_url(clean_url):
                urls.append(clean_url)
        
        # 去重并排序
        return sorted(list(dict.fromkeys(urls)))
    
    def generate_raw_values(self, items: Dict[str, List[str]]) -> str:
        """生成原始捕获值的纯文本格式"""
        raw_sections = []
        
        for item_type, item_list in items.items():
            if item_list:
                # 去重并排序
                unique_items = sorted(list(dict.fromkeys(item_list)))
                if unique_items:
                    raw_sections.append(f"{item_type.upper()}:")
                    for item in unique_items:
                        raw_sections.append(f"  {item.strip()}")
                    raw_sections.append("")  # 空行分隔
        
        return "\n".join(raw_sections).strip()
    
    def get_first_value(self, items: Dict[str, List[str]]) -> Optional[str]:
        """获取第一个捕获的值（按优先级顺序）"""
        # 定义优先级顺序
        priority_order = ['sr', 'cmos', 'doc', 'bug', 'jira', 'people', 'email', 'url']
        
        for item_type in priority_order:
            item_list = items.get(item_type, [])
            if item_list:
                # 返回第一个值，去除@符号（对于people类型）
                first_value = item_list[0].strip()
                if item_type == 'people':
                    first_value = first_value.lstrip('@')
                return first_value
        
        return None
    
    def validate_url(self, url: str) -> bool:
        """验证URL是否有效"""
        try:
            result = parse.urlparse(url)
            if not all([result.scheme, result.netloc]):
                return False
            if result.scheme not in ['http', 'https']:
                return False
            
            # 检查可疑字符
            if any(c in url for c in ['<', '>', '"', "'", ';', '|', '`']):
                return False
            
            # 检查允许的域名
            return any(domain in result.netloc for domain in self.config.ALLOWED_DOMAINS)
        except Exception:
            return False

class GoogleSearcher:
    """Google搜索器"""
    
    def __init__(self, platform_interface: PlatformInterface):
        self.platform = platform_interface
    
    def search(self, search_text: str) -> None:
        """执行Google搜索"""
        try:
            if not search_text.strip():
                logging.warning("搜索文本为空，跳过搜索")
                return
            
            search_url = self._build_search_url(search_text)
            if self.platform.open_browser(search_url):
                logging.info(f"执行Google搜索: {search_text}")
            else:
                logging.error("打开Google搜索失败")
                
        except Exception as e:
            logging.error(f"Google搜索执行失败: {str(e)}")
    
    def _build_search_url(self, search_text: str) -> str:
        """构建搜索URL"""
        # Oracle相关网站的搜索范围
        search_scope = (
            ' ( '
            # Oracle 产品文档
            'site:docs.oracle.com OR '
            # Oracle Cloud文档
            'site:docs.cloud.oracle.com OR '
            # Oracle 社区和支持
            'site:community.oracle.com OR '
            'site:support.oracle.com OR '
            # Oracle 技术博客
            'site:blogs.oracle.com OR '
            # Oracle 知识库
            'site:support.oracle.com/knowledge OR '
            # My Oracle Support
            'site:support.oracle.com/epmos '
            ') '
        )
        
        # 添加时间限制（最近2年的结果）
        current_year = datetime.datetime.now().year
        time_filter = f' after:{current_year-2}'
        
        # 构建完整的搜索参数
        search_params = {
            'q': search_text + search_scope + time_filter,
            'hl': 'en',  # 设置语言为英文
            'num': '100',  # 显示更多结果
            'safe': 'active'  # 安全搜索
        }
        
        return "https://www.google.com/search?" + parse.urlencode(search_params)

class OQOTool:
    """Oracle Quick Open工具主类"""
    
    def __init__(self, auto_open: bool = True, raw_capture: bool = False, one_value: bool = False):
        """初始化OQO工具"""
        self.auto_open = auto_open
        self.raw_capture = raw_capture
        self.one_value = one_value
        self.config = Config()
        self.platform = self._create_platform_interface()
        self.url_processor = URLProcessor(self.config)
        self.google_searcher = GoogleSearcher(self.platform)
        
        # 配置日志
        self._setup_logging()
    
    def _create_platform_interface(self) -> PlatformInterface:
        """创建平台接口"""
        if platform.system() == 'Darwin':
            return MacOSInterface()
        else:
            return LinuxInterface()
    
    def _setup_logging(self):
        """设置日志系统"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('oqo.log'),
                logging.StreamHandler()
            ]
        )
    
    def run(self):
        """运行主程序"""
        try:
            # 获取剪贴板内容
            clipboard_text = self.platform.get_selected_text()
            if not clipboard_text:
                logging.error("无法获取剪贴板内容")
                self.platform.send_notification("无法获取剪贴板内容")
                return
            
            # 提取并处理项目
            items = self.url_processor.extract_items(clipboard_text)
            
            if self.one_value:
                # 单值模式：只显示第一个捕获的值
                first_value = self.url_processor.get_first_value(items)
                if first_value:
                    self.platform.set_clipboard(first_value)
                    message = f'已提取第一个值: {first_value}'
                    self.platform.send_notification(message)
                    logging.info(message)
                    
                    # 打印第一个值
                    print(first_value)
                else:
                    self.platform.send_notification("未找到匹配的项目")
                    logging.info("未找到匹配的项目")
            elif self.raw_capture:
                # 原始捕获模式：只显示纯文本值
                raw_text = self.url_processor.generate_raw_values(items)
                if raw_text:
                    self.platform.set_clipboard(raw_text)
                    message = f'已提取 >> {sum(len(items.get(k, [])) for k in items.keys())} << 个原始值到剪贴板'
                    self.platform.send_notification(message)
                    logging.info(message)
                    
                    # 打印原始值
                    print("\n" + "="*60)
                    print("提取的原始值:")
                    print("="*60)
                    print(raw_text)
                    print("="*60)
                else:
                    self.platform.send_notification("未找到匹配的项目")
                    logging.info("未找到匹配的项目")
            else:
                # 正常模式：生成URL
                urls = self.url_processor.generate_urls(items)
                
                if urls:
                    # 复制URL到剪贴板
                    self.platform.set_clipboard("\n".join(urls))
                    
                    if self.auto_open:
                        self._open_urls(urls)
                        message = f'打开成功 >> {len(urls)} << 个项目\n已复制到剪贴板'
                    else:
                        message = f'已提取 >> {len(urls)} << 个项目到剪贴板'
                    
                    self.platform.send_notification(message)
                    logging.info(message)
                    
                    # 打印提取的项目统计
                    self._print_extraction_summary(items)
                else:
                    # 执行Google搜索
                    self.google_searcher.search(clipboard_text)
                    logging.info("执行Google搜索")
                
        except Exception as e:
            logging.error(f"程序运行出错: {str(e)}")
            self.platform.send_notification("程序运行出错，请查看日志")
    
    def _print_extraction_summary(self, items: Dict[str, List[str]]) -> None:
        """打印提取项目统计"""
        print("\n" + "="*60)
        print("提取项目统计:")
        print("="*60)
        total_count = 0
        for item_type, item_list in items.items():
            if item_list:
                count = len(item_list)
                total_count += count
                print(f"{item_type.upper()}: {count}")
        print(f"\n总计: {total_count} 个项目")
        print("="*60)
    
    def _open_urls(self, urls: List[str]) -> None:
        """打开URL列表"""
        opened_count = 0
        last_open_time = time.time()
        min_interval = 0.5  # 最小间隔时间（秒）
        
        for url in urls:
            url = url.strip()
            if not url or not self.url_processor.validate_url(url):
                logging.warning(f"跳过无效URL: {url}")
                continue
            
            # 速率限制，避免过快打开URL
            current_time = time.time()
            if current_time - last_open_time < min_interval:
                time.sleep(min_interval - (current_time - last_open_time))
            
            if self.platform.open_browser(url):
                opened_count += 1
                last_open_time = time.time()
                logging.info(f"成功打开URL: {url}")
            else:
                logging.error(f"打开URL失败: {url}")
        
        logging.info(f"总共打开 {opened_count}/{len(urls)} 个URL")

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='Oracle Quick Open Tool')
    parser.add_argument('--no-open', action='store_true',
                      help='只提取URL而不自动打开浏览器')
    parser.add_argument('--raw-capture', action='store_true',
                      help='只显示原始捕获值，不转换为URL格式')
    parser.add_argument('--one-value', action='store_true',
                      help='只显示第一个捕获的值，无额外文本')
    parser.add_argument('--debug', action='store_true',
                      help='启用调试模式')
    
    args = parser.parse_args()
    
    # 设置调试模式
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    try:
        oqo = OQOTool(auto_open=not args.no_open, raw_capture=args.raw_capture, one_value=args.one_value)
        oqo.run()
    except KeyboardInterrupt:
        print("\n程序被用户中断")
        sys.exit(0)
    except Exception as e:
        print(f"程序启动失败: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
