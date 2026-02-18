import os
import re
import sys
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from pathlib import Path
import requests
import datetime
import threading
import webbrowser
import tempfile
from openai import OpenAI

class FileToMindmapApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("文件转思维导图/PPT生成器 + 内容对话")
        self.root.geometry("1100x800")
        
        # 对话历史存储
        self.conversation_history = []
        
        # AI模型默认选择
        self.ai_model_var = tk.StringVar(value="DeepSeek")
        
        # 检查核心依赖可用性
        self.selenium_available = self.check_selenium()
        self.pyperclip_available = self.check_pyperclip()
        
        # 初始化UI、AI客户端、日志
        self.setup_ui()
        self.init_ai_clients()
        self.log("程序启动成功！请选择文件或输入路径开始使用")
        
        # 依赖缺失提示
        if not self.selenium_available:
            self.log("  警告: 未安装selenium库，自动上传功能将受限")

    def check_selenium(self) -> bool:
        """检查selenium及相关组件是否可用（控制浏览器自动上传）"""
        try:
            import selenium
            from selenium import webdriver
            from selenium.webdriver.common.by import By
            return True
        except ImportError:
            return False

    def check_pyperclip(self) -> bool:
        """检查pyperclip是否可用（复制内容到剪贴板）"""
        try:
            import pyperclip
            return True
        except ImportError:
            return False

    def init_ai_clients(self):
        """初始化DeepSeek和Kimi AI客户端（读取环境变量中的API密钥）"""
        # DeepSeek客户端初始化
        self.deepseek_client = None
        deepseek_api_key = os.getenv("DEEPSEEK_API_KEY")
        if deepseek_api_key:
            try:
                # 验证API密钥格式
                if not self._validate_api_key(deepseek_api_key, "DeepSeek"):
                    self.log("  DeepSeek API密钥格式无效")
                else:
                    self.deepseek_client = OpenAI(
                        api_key=deepseek_api_key,
                        base_url="https://api.deepseek.com"
                    )
                    self.log(" DeepSeek客户端初始化成功")
            except Exception as e:
                self.log(f" 初始化DeepSeek客户端失败: {str(e)}")
        else:
            self.log("  警告: 未设置DEEPSEEK_API_KEY环境变量")
        
        # Kimi API配置
        self.kimi_api_key = os.getenv("KIMI_API_KEY")
        self.kimi_api_url = "https://api.moonshot.cn/v1/chat/completions"
        if self.kimi_api_key:
            if not self._validate_api_key(self.kimi_api_key, "Kimi"):
                self.log("  Kimi API密钥格式无效")
            else:
                self.log(" Kimi API密钥已设置")
        else:
            self.log("  警告: 未设置KIMI_API_KEY环境变量")
        
        # 检查API密钥状态
        self.check_api_status()
    
    def _validate_api_key(self, api_key, provider):
        """验证API密钥格式"""
        if not api_key or not isinstance(api_key, str):
            return False
        
        # 不同提供商的API密钥格式验证
        if provider == "DeepSeek":
            # DeepSeek API密钥通常以sk-开头
            return api_key.startswith("sk-") and len(api_key) > 20
        elif provider == "Kimi":
            # Kimi API密钥通常以sk-开头
            return api_key.startswith("sk-") and len(api_key) > 20
        return False

    def check_api_status(self):
        """检查两个AI模型的API可用性"""
        deepseek_status = "可用" if os.getenv("DEEPSEEK_API_KEY") and self.deepseek_client else "不可用"
        kimi_status = "可用" if os.getenv("KIMI_API_KEY") else "不可用"
        self.log(f" API状态检查: DeepSeek({deepseek_status}), Kimi({kimi_status})")
        
        # 无可用API时提示
        if deepseek_status == "不可用" and kimi_status == "不可用":
            self.log(" 警告: 没有可用的API密钥，生成功能将无法使用")
            messagebox.showwarning("API密钥缺失", "请至少设置一个API密钥（DeepSeek或Kimi）才能使用生成功能！")

    def setup_ui(self):
        """创建主界面（选项卡结构：生成功能 + 对话功能）"""
        main_notebook = ttk.Notebook(self.root)
        main_notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 选项卡1：思维导图/PPT生成
        mindmap_frame = ttk.Frame(main_notebook)
        main_notebook.add(mindmap_frame, text='思维导图/PPT生成')
        self.setup_mindmap_tab(mindmap_frame)
        
        # 选项卡2：文件内容对话
        chat_frame = ttk.Frame(main_notebook)
        main_notebook.add(chat_frame, text='文件内容对话')
        self.setup_chat_tab(chat_frame)

    def setup_mindmap_tab(self, parent_frame):
        """配置思维导图/PPT生成选项卡的UI组件"""
        main_frame = ttk.Frame(parent_frame, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 1. 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="1. 选择输入文件", padding="10")
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10), columnspan=2)
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=60).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(file_frame, text="浏览...", command=self.browse_file).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="手动输入路径", command=self.input_path_dialog).grid(row=0, column=2, padx=5)
        
        # 2. 文件内容预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="2. 文件内容预览", padding="10")
        preview_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=10, wrap=tk.WORD)
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 3. 生成选项配置区域
        self.options_frame = ttk.LabelFrame(main_frame, text="3. 生成选项", padding="10")
        self.options_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 输出类型选择（思维导图/PPT大纲）
        ttk.Label(self.options_frame, text="输出类型:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.output_type_var = tk.StringVar(value="思维导图")
        self.output_type_combo = ttk.Combobox(self.options_frame, textvariable=self.output_type_var, 
                                            values=["思维导图", "PPT大纲"], state="readonly", width=15)
        self.output_type_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        self.output_type_combo.bind("<<ComboboxSelected>>", self.on_output_type_changed)
        
        # 生成语言选择
        ttk.Label(self.options_frame, text="生成语言:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.lang_var = tk.StringVar(value="中文")
        ttk.Combobox(self.options_frame, textvariable=self.lang_var, values=["中文", "英文"], state="readonly", width=12).grid(row=0, column=3, sticky=tk.W, padx=(0, 20))
        
        # 思维导图层级选择（默认显示）
        self.depth_label = ttk.Label(self.options_frame, text="思维导图层级:")
        self.depth_label.grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
        self.depth_var = tk.StringVar(value="3 (推荐)")
        self.depth_combo = ttk.Combobox(self.options_frame, textvariable=self.depth_var, 
                                       values=["2", "3 (推荐)", "4", "5"], state="readonly", width=12)
        self.depth_combo.grid(row=0, column=5, sticky=tk.W, padx=(0, 5))
        
        # AI模型选择
        ttk.Label(self.options_frame, text="AI模型:").grid(row=0, column=6, sticky=tk.W, padx=(20, 5))
        self.ai_model_combo = ttk.Combobox(self.options_frame, textvariable=self.ai_model_var, 
                                         values=["DeepSeek", "Kimi"], state="readonly", width=12)
        self.ai_model_combo.grid(row=0, column=7, sticky=tk.W, padx=(0, 5))
        self.ai_model_combo.bind("<<ComboboxSelected>>", self.on_model_changed)
        
        # PPT页数选择（默认隐藏，切换输出类型时显示）
        self.ppt_pages_label = ttk.Label(self.options_frame, text="PPT页数:")
        self.ppt_pages_var = tk.StringVar(value="10-15")
        self.ppt_pages_combo = ttk.Combobox(self.options_frame, textvariable=self.ppt_pages_var, 
                                          values=["5-8", "10-15", "15-20", "20-25"], state="readonly", width=12)
        
        # 4. 功能按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, pady=(10, 5))
        
        # 第一行按钮（生成/保存/清空）
        button_row1 = ttk.Frame(button_frame)
        button_row1.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.generate_btn = ttk.Button(button_row1, text="生成思维导图大纲", command=self.generate_content, state=tk.DISABLED)
        self.generate_btn.grid(row=0, column=0, padx=5)
        self.save_btn = ttk.Button(button_row1, text="保存为Markdown", command=self.save_as_markdown, state=tk.DISABLED)
        self.save_btn.grid(row=0, column=1, padx=5)
        self.clear_btn = ttk.Button(button_row1, text="清空日志", command=self.clear_log)
        self.clear_btn.grid(row=0, column=2, padx=5)
        
        # 第二行按钮（跳转外部工具）
        button_row2 = ttk.Frame(button_frame)
        button_row2.grid(row=1, column=0, sticky=tk.W)
        ttk.Label(button_row2, text="使用在线工具:").grid(row=0, column=0, padx=(0, 5))
        
        # PPT生成按钮（跳转Kimi Slides）
        ppt_button = ttk.Button(button_row2, text="PPT生成", command=self.open_ppt_generator, width=12)
        ppt_button.grid(row=0, column=1, padx=5)
        self.create_tooltip(ppt_button, "点击跳转到 Kimi AI 的PPT生成工具")
        
        # 思维导图生成按钮（跳转XMind）
        mindmap_button = ttk.Button(button_row2, text="思维导图生成", command=self.open_mindmap_generator, width=15)
        mindmap_button.grid(row=0, column=2, padx=5)
        self.create_tooltip(mindmap_button, "点击跳转到 XMind 在线思维导图工具")
        
        # 5. 结果与日志显示区域
        result_frame = ttk.LabelFrame(main_frame, text="4. 生成结果与程序日志", padding="10")
        result_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        result_frame.grid_columnconfigure(0, weight=1)
        result_frame.grid_rowconfigure(0, weight=1)
        
        # 分割结果和日志的面板
        paned = ttk.PanedWindow(result_frame, orient=tk.VERTICAL)
        paned.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 生成结果显示
        result_container = ttk.Frame(paned)
        result_container.grid_columnconfigure(0, weight=1)
        result_container.grid_rowconfigure(0, weight=1)
        self.result_label = ttk.Label(result_container, text="思维导图大纲 (Markdown格式):")
        self.result_label.grid(row=0, column=0, sticky=tk.W, pady=(0,5))
        self.result_text = scrolledtext.ScrolledText(result_container, height=12, wrap=tk.WORD)
        self.result_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        paned.add(result_container, weight=3)
        
        # 程序日志显示
        log_container = ttk.Frame(paned)
        log_container.grid_columnconfigure(0, weight=1)
        log_container.grid_rowconfigure(0, weight=1)
        ttk.Label(log_container, text="程序日志:").grid(row=0, column=0, sticky=tk.W, pady=(0,5))
        self.log_text = scrolledtext.ScrolledText(log_container, height=6, wrap=tk.WORD)
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        paned.add(log_container, weight=1)
        
        # 网格权重配置（自适应窗口大小）
        parent_frame.columnconfigure(0, weight=1)
        parent_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=2)
        main_frame.rowconfigure(4, weight=3)

    def create_tooltip(self, widget, text):
        """为按钮添加鼠标悬浮提示"""
        def enter(event):
            x, y, cx, cy = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 20
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            label = ttk.Label(self.tooltip, text=text, background="#ffffe0", relief="solid", borderwidth=1)
            label.pack()
            
        def leave(event):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
                
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def open_ppt_generator(self):
        """打开Kimi Slides并自动上传生成的PPT大纲（md文件）"""
        if not hasattr(self, 'generated_markdown') or not self.generated_markdown:
            messagebox.showwarning("警告", "请先生成PPT大纲内容！")
            return
            
        # 显示加载状态
        self.log(" 正在准备上传文件...")
        
        # 新线程执行上传（避免阻塞UI）
        def upload_task():
            try:
                # 生成临时md文件
                temp_dir = tempfile.gettempdir()
                temp_file = os.path.join(temp_dir, "ppt_outline.md")
                with open(temp_file, 'w', encoding='utf-8') as f:
                    f.write(self.generated_markdown)
                self.log(f" 已将PPT大纲保存到临时文件: {temp_file}")
                
                # Selenium不可用时使用备用方案
                if not self.selenium_available:
                    self.log("  Selenium不可用，使用备用方案")
                    self.fallback_ppt_generation(temp_file)
                    return
                
                # 执行上传
                self.log(" 正在启动浏览器上传文件...")
                self.upload_to_kimi_slides(temp_file)
                
            except Exception as e:
                error_msg = f" 打开Kimi PPT生成工具失败: {str(e)}"
                self.log(error_msg)
                self.root.after(0, lambda: messagebox.showerror("打开失败", f"无法自动上传文件:\n{str(e)}"))
        
        # 启动上传线程
        upload_thread = threading.Thread(target=upload_task)
        upload_thread.daemon = True
        upload_thread.start()

    def _get_chrome_driver(self):
        """获取Chrome浏览器驱动"""
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.service import Service
            from selenium.webdriver.chrome.options import Options
            import time
            
            # Chrome浏览器配置
            chrome_options = Options()
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--start-maximized')
            chrome_options.add_argument('--ignore-certificate-errors')
            chrome_options.add_argument('--disable-extensions')
            chrome_options.add_argument(f"--user-data-dir={os.path.join(tempfile.gettempdir(), 'xmind_chrome_profile')}")
            
            # 查找chromedriver.exe
            chromedriver_paths = [
                "chromedriver.exe",
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "chromedriver.exe"),
                os.path.join(os.environ.get("PATH", ""), "chromedriver.exe")
            ]
            
            chromedriver_path = None
            for path in chromedriver_paths:
                if os.path.exists(path):
                    chromedriver_path = path
                    break
            
            if not chromedriver_path:
                raise FileNotFoundError(" 未找到chromedriver.exe，请确保它在项目目录或系统PATH中")
            
            self.log(f"  使用chromedriver: {chromedriver_path}")
            
            # 启动Chrome
            try:
                service = Service(executable_path=chromedriver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
                # 设置隐式等待
                driver.implicitly_wait(5)
                return driver
            except Exception as e:
                raise Exception(f"启动Chrome失败: {str(e)}")
        except Exception as e:
            raise Exception(f"获取Chrome驱动失败: {str(e)}")
    
    def _find_upload_element(self, driver):
        """查找上传元素"""
        from selenium.webdriver.common.by import By
        from selenium.common.exceptions import NoSuchElementException, TimeoutException
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        import time
        
        # 尝试多种方式查找上传元素
        upload_element = None
        
        # 等待页面加载完成
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except TimeoutException:
            self.log(" 页面加载超时，尝试直接查找元素")
        
        # 尝试方法1: 使用指定XPath
        try:
            upload_element = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div[2]/div[2]/label/svg")
            if upload_element and upload_element.is_displayed():
                self.log(" 使用指定XPath找到上传元素")
                return upload_element
        except NoSuchElementException:
            self.log(" 指定XPath未找到上传元素，尝试其他方法")
        except Exception as e:
            self.log(f" 使用XPath查找失败: {str(e)}")
        
        # 尝试方法2: 使用多种CSS选择器
        selectors = [
            "input[type='file']", 
            "[class*='upload']", 
            "[data-testid*='upload']",
            "button:contains('上传')", 
            "label:contains('上传')",
            "button", 
            "label", 
            "svg"
        ]
        
        for selector in selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                self.log(f" 使用选择器 '{selector}' 找到 {len(elements)} 个元素")
                for i, element in enumerate(elements):
                    try:
                        if element.is_displayed() and element.is_enabled():
                            # 尝试获取元素文本或属性，判断是否为上传按钮
                            try:
                                text = element.text.lower()
                                if "上传" in text or "upload" in text:
                                    upload_element = element
                                    self.log(f" 使用CSS选择器找到上传元素: {selector} (元素 {i+1})")
                                    return upload_element
                            except:
                                pass
                            
                            # 对于input[type='file']，直接返回
                            if selector == "input[type='file']":
                                upload_element = element
                                self.log(f" 找到文件输入元素: {selector}")
                                return upload_element
                    except Exception as e:
                        self.log(f" 检查元素 {i+1} 时出错: {str(e)}")
                        continue
            except Exception as e:
                self.log(f" 使用选择器 '{selector}' 查找时出错: {str(e)}")
                continue
        
        # 尝试方法3: 查找所有可点击元素，检查是否包含上传相关文本
        try:
            all_buttons = driver.find_elements(By.TAG_NAME, "button")
            for button in all_buttons:
                try:
                    if button.is_displayed() and button.is_enabled():
                        text = button.text.lower()
                        if "上传" in text or "upload" in text:
                            upload_element = button
                            self.log(" 通过按钮文本找到上传元素")
                            return upload_element
                except:
                    continue
        except Exception as e:
            self.log(f" 查找所有按钮时出错: {str(e)}")
        
        self.log(" 未找到上传元素")
        return None
    
    def upload_to_kimi_slides(self, file_path):
        """使用Selenium自动上传md文件到Kimi Slides"""
        try:
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.common.exceptions import TimeoutException, NoSuchElementException
            import time
            
            # 获取驱动并打开Kimi Slides
            driver = self._get_chrome_driver()
            driver.get("https://www.kimi.com/slides")
            self.log(" 已打开Kimi Slides")
            
            # 等待页面加载
            time.sleep(5)
            
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                self.log(" 正在查找上传按钮...")
                
                # 查找上传元素
                upload_element = self._find_upload_element(driver)
                
                # 未找到上传元素时提示手动操作
                if upload_element is None:
                    self.log("  未找到自动上传元素，将打开手动上传页面")
                    self.root.after(0, lambda: messagebox.showinfo("手动上传", 
                        f"已打开Kimi Slides页面。\nMarkdown文件已保存到: {file_path}\n请在网页中手动上传文件。"))
                    return
                
                # 执行上传操作
                upload_element.click()
                time.sleep(2)
                file_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
                
                if file_inputs:
                    file_input = file_inputs[0]
                    file_input.send_keys(file_path)  # 自动输入md文件路径
                    self.log(f" 已上传文件: {file_path}")
                    time.sleep(3)
                    
                    # 尝试点击生成按钮
                    self._try_click_generate_button(driver)
                else:
                    self.root.after(0, lambda: messagebox.showinfo("手动上传", 
                        f"已打开Kimi Slides页面。\nMarkdown文件已保存到: {file_path}\n请在网页中手动上传文件。"))
                    
            except TimeoutException:
                self.log("  等待页面加载超时")
                self.root.after(0, lambda: messagebox.showinfo("手动上传", 
                    f"已打开Kimi Slides页面。\nMarkdown文件已保存到: {file_path}\n请在网页中手动上传文件。"))
                
        except FileNotFoundError as e:
            self.log(f" ChromeDriver未找到: {str(e)}")
            self.root.after(0, lambda: self.fallback_ppt_generation(file_path))
        except Exception as e:
            self.log(f" Selenium启动失败: {str(e)}")
            self.root.after(0, lambda: self.fallback_ppt_generation(file_path))
    
    def _try_click_generate_button(self, driver):
        """尝试点击生成按钮"""
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import TimeoutException
        import time
        
        try:
            xpath_to_click = "/html/body/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div[3]/div[2]/div[3]/div"
            element_to_click = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath_to_click))
            )
            if element_to_click and element_to_click.is_displayed():
                element_to_click.click()
                self.log(f" 成功点击生成按钮")
                time.sleep(3)
                
                # 检测生成状态
                page_source = driver.page_source.lower()
                success_indicators = ["上传成功", "解析完成", "开始生成", "生成中", "processing"]
                for indicator in success_indicators:
                    if indicator.lower() in page_source:
                        self.log(f" 检测到生成开始: {indicator}")
                        self.root.after(0, lambda: messagebox.showinfo("开始生成", "PPT已经开始生成，请等待处理完成。"))
                        break
            else:
                self.log("  生成按钮不可点击")
        except TimeoutException:
            self.log("  等待生成按钮超时")

    def fallback_ppt_generation(self, file_path):
        """PPT生成备用方案：复制内容到剪贴板或提示手动上传"""
        try:
            if self.pyperclip_available:
                import pyperclip
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                pyperclip.copy(content)
                webbrowser.open("https://www.kimi.com/slides")
                self.root.after(0, lambda: messagebox.showinfo("备用方案", 
                    f"已打开Kimi Slides页面。\nPPT大纲内容已复制到剪贴板。\n文件保存路径: {file_path}\n\n请按 Ctrl+V 粘贴内容。"))
            else:
                webbrowser.open("https://www.kimi.com/slides")
                self.root.after(0, lambda: messagebox.showinfo("手动操作", 
                    f"已打开Kimi Slides页面。\nPPT大纲文件保存路径: {file_path}\n\n请手动上传此文件。"))
        except Exception as e:
            self.log(f" 备用方案失败: {str(e)}")
            webbrowser.open("https://www.kimi.com/slides")

    def open_mindmap_generator(self):
        """打开XMind并自动上传生成的思维导图大纲（md文件）"""
        if not hasattr(self, 'generated_markdown') or not self.generated_markdown:
            messagebox.showwarning("警告", "请先生成思维导图大纲内容！")
            return
            
        # 显示加载状态
        self.log(" 正在准备上传文件...")
        
        # 新线程执行上传（避免阻塞UI）
        def upload_task():
            try:
                # 生成临时md文件（带时间戳避免重名）
                temp_dir = tempfile.gettempdir()
                temp_file = os.path.join(temp_dir, f"mindmap_outline_{int(datetime.datetime.now().timestamp())}.md")
                with open(temp_file, 'w', encoding='utf-8') as f:
                    f.write(self.generated_markdown)
                self.log(f" 已将思维导图大纲保存到临时文件: {temp_file}")
                
                # Selenium不可用时使用备用方案
                if not self.selenium_available:
                    self.log("  Selenium不可用，使用备用方案")
                    self.fallback_mindmap_generation(temp_file)
                    return
                
                # 执行上传
                self.log(" 正在启动浏览器上传到XMind...")
                self.upload_to_xmind(temp_file)
                
            except Exception as e:
                error_msg = f" 打开XMind工具失败: {str(e)}"
                self.log(error_msg)
                self.root.after(0, lambda: messagebox.showerror("打开失败", f"无法自动上传文件:\n{str(e)}"))
        
        # 启动上传线程
        upload_thread = threading.Thread(target=upload_task)
        upload_thread.daemon = True
        upload_thread.start()

    def upload_to_xmind(self, file_path):
        """使用Selenium自动上传md文件到XMind在线工作台"""
        try:
            from selenium.webdriver.common.by import By
            from selenium.common.exceptions import NoSuchElementException
            import time
            
            # 获取驱动并打开XMind工作台
            driver = self._get_chrome_driver()
            driver.get("https://app.xmind.cn/home/my-works")
            self.log(" 已打开XMind工作台")
            
            # 等待页面加载和登录
            time.sleep(5)
            self.log(" 请在1秒内完成登录...")
            for i in range(1, 0, -1):
                self.log(f" 等待登录剩余时间: {i}秒")
                time.sleep(1)
            self.log(" 登录等待结束，开始执行上传操作")
            time.sleep(3)

            # 执行上传操作
            self._execute_xmind_upload_steps(driver, file_path)
                
        except FileNotFoundError as e:
            self.log(f" ChromeDriver未找到: {str(e)}")
            self.root.after(0, lambda: self.fallback_mindmap_generation(file_path))
        except Exception as e:
            self.log(f" 自动上传失败: {str(e)}")
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            self.root.after(0, lambda: self.fallback_mindmap_generation(file_path))
    
    def _execute_xmind_upload_steps(self, driver, file_path):
        """执行XMind上传步骤"""
        from selenium.webdriver.common.by import By
        from selenium.common.exceptions import NoSuchElementException
        import time
        
        # 第一步：点击第一个XPath
        self.log("正在查找第一个元素...")
        xpath1 = "/html/body/div[1]/div/div/div/div[4]/div/section[2]/div/button[3]"
            
        try:
                element1 = driver.find_element(By.XPATH, xpath1)
                if element1 and element1.is_displayed() and element1.is_enabled():
                    self.log(f"找到第一个元素: {xpath1}")
                    element1.click()
                    self.log("成功点击第一个元素")
                    time.sleep(2)
                else:
                    self.log("第一个元素不可点击，尝试其他选择器")
                    raise NoSuchElementException("第一个元素不可点击")
        except (NoSuchElementException, Exception) as e:
                self.log(f"使用第一个XPath失败: {str(e)}")
                self.root.after(0, lambda: messagebox.showwarning("操作失败", 
                    f"无法找到第一个元素，请手动操作。\n文件已保存到: {file_path}"))
                return
            
            # 第二步：
        self.log("正在查找第二个元素(textarea)...")
        time.sleep(2)
        
        xpath2 = "/html/body/div[19]/div/div/div[1]/div/div[2]/div[2]/div[2]/div/div"
        
        try:
                element2 = driver.find_element(By.XPATH, xpath2)
                if element2 and element2.is_displayed() and element2.is_enabled():
                    self.log(f"找到第二个元素: {xpath2}")
                    element2.click()
                    self.log("成功点击第二个元素")
                    time.sleep(2)
                else:
                    self.log("第二个元素不可点击，尝试其他选择器")
                    raise NoSuchElementException("第二个元素不可点击")
        except (NoSuchElementException, Exception) as e:
                self.log(f"使用第二个XPath失败: {str(e)}")
                self.root.after(0, lambda: messagebox.showwarning("操作失败", 
                    f"无法找到第二个元素，请手动操作。\n文件已保存到: {file_path}"))
                return

    def fallback_mindmap_generation(self, file_path):
        """思维导图生成备用方案：复制内容到剪贴板或提示手动上传"""
        try:
            if self.pyperclip_available:
                import pyperclip
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                pyperclip.copy(content)
                webbrowser.open("https://app.xmind.cn/home/my-works")
                self.root.after(0, lambda: messagebox.showinfo("备用方案", 
                    f"已打开XMind工作台页面。\n思维导图大纲内容已复制到剪贴板。\n文件保存路径: {file_path}\n\n"
                    "操作步骤:\n1. 登录XMind账号\n2. 新建思维导图\n3. 按Ctrl+V粘贴内容"))
            else:
                webbrowser.open("https://app.xmind.cn/home/my-works")
                self.root.after(0, lambda: messagebox.showinfo("手动操作", 
                    f"已打开XMind工作台页面。\n思维导图大纲文件保存路径: {file_path}\n\n请手动上传此文件。"))
        except Exception as e:
            self.log(f" 备用方案失败: {str(e)}")
            webbrowser.open("https://app.xmind.cn/home/my-works")

    def setup_chat_tab(self, parent_frame):
        """配置文件内容对话选项卡的UI组件"""
        chat_main_frame = ttk.Frame(parent_frame, padding="10")
        chat_main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 1. 文件状态区域
        status_frame = ttk.LabelFrame(chat_main_frame, text="当前文件状态", padding="10")
        status_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.chat_file_status_var = tk.StringVar(value="未加载文件")
        ttk.Label(status_frame, textvariable=self.chat_file_status_var, font=('Arial', 10)).grid(row=0, column=0, sticky=tk.W)
        self.chat_file_length_var = tk.StringVar(value="文件内容长度: 0 字符")
        ttk.Label(status_frame, textvariable=self.chat_file_length_var, font=('Arial', 9)).grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        ttk.Button(status_frame, text="重新加载文件", command=self.reload_file_for_chat, width=15).grid(row=0, column=2, padx=(20, 0))
        
        # 对话AI模型选择
        ttk.Label(status_frame, text="对话AI模型:").grid(row=0, column=3, sticky=tk.W, padx=(20, 5))
        self.chat_ai_model_var = tk.StringVar(value="DeepSeek")
        self.chat_ai_model_combo = ttk.Combobox(status_frame, textvariable=self.chat_ai_model_var, 
                                               values=["DeepSeek", "Kimi"], state="readonly", width=12)
        self.chat_ai_model_combo.grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
        
        # 2. 对话历史显示区域
        chat_history_frame = ttk.LabelFrame(chat_main_frame, text="对话历史", padding="10")
        chat_history_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        chat_history_frame.grid_columnconfigure(0, weight=1)
        chat_history_frame.grid_rowconfigure(0, weight=1)
        self.chat_history_text = scrolledtext.ScrolledText(chat_history_frame, height=20, wrap=tk.WORD, font=('Arial', 10))
        self.chat_history_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 3. 提问输入区域
        input_frame = ttk.LabelFrame(chat_main_frame, text="提问", padding="10")
        input_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.grid_columnconfigure(0, weight=1)
        self.question_var = tk.StringVar()
        question_entry = ttk.Entry(input_frame, textvariable=self.question_var, font=('Arial', 10))
        question_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        question_entry.bind('<Return>', lambda e: self.ask_question())
        self.ask_btn = ttk.Button(input_frame, text="提问", command=self.ask_question, width=10, state=tk.DISABLED)
        self.ask_btn.grid(row=0, column=1, padx=5)
        ttk.Button(input_frame, text="清除对话", command=self.clear_conversation, width=10).grid(row=0, column=2, padx=5)
        
        # 4. 示例问题区域
        examples_frame = ttk.LabelFrame(chat_main_frame, text="示例问题（点击使用）", padding="10")
        examples_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        example_questions = [
            "这个文件主要讲了什么内容？", "总结一下文件的核心要点",
            "文件中提到了哪些重要概念？", "作者的主要观点是什么？",
            "这个文件的结构是怎样的？", "文件中的关键数据或事实有哪些？"
        ]
        for i, question in enumerate(example_questions):
            btn = ttk.Button(examples_frame, text=question, command=lambda q=question: self.use_example_question(q), width=40)
            btn.grid(row=i//2, column=i%2, padx=5, pady=2, sticky=tk.W)
        
        # 网格权重配置
        parent_frame.columnconfigure(0, weight=1)
        parent_frame.rowconfigure(0, weight=1)
        chat_main_frame.columnconfigure(0, weight=1)
        chat_main_frame.rowconfigure(1, weight=1)
        
        # 状态变量初始化
        self.current_file_content = ""
        self.generated_markdown = ""
        self.current_file_path = ""

    def on_output_type_changed(self, event=None):
        """切换输出类型时更新UI（显示/隐藏对应的配置项）"""
        output_type = self.output_type_var.get()
        if output_type == "思维导图":
            # 显示思维导图层级，隐藏PPT页数
            self.depth_label.grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
            self.depth_combo.grid(row=0, column=5, sticky=tk.W, padx=(0, 5))
            self.ppt_pages_label.grid_remove()
            self.ppt_pages_combo.grid_remove()
            # 更新按钮和结果标签文本
            self.generate_btn.config(text="生成思维导图大纲")
            self.result_label.config(text="思维导图大纲 (Markdown格式):")
        elif output_type == "PPT大纲":
            # 显示PPT页数，隐藏思维导图层级
            self.ppt_pages_label.grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
            self.ppt_pages_combo.grid(row=0, column=5, sticky=tk.W, padx=(0, 5))
            self.depth_label.grid_remove()
            self.depth_combo.grid_remove()
            # 更新按钮和结果标签文本
            self.generate_btn.config(text="生成PPT大纲")
            self.result_label.config(text="PPT大纲 (Markdown格式):")

    def on_model_changed(self, event=None):
        """切换AI模型时检查API密钥可用性"""
        ai_model = self.ai_model_var.get()
        self.log(f" 切换AI模型为: {ai_model}")
        if ai_model == "DeepSeek" and not os.getenv("DEEPSEEK_API_KEY"):
            self.log("  警告: DeepSeek API密钥未设置")
            messagebox.showwarning("API密钥警告", "DeepSeek API密钥未设置。\n请设置DEEPSEEK_API_KEY环境变量。")
        elif ai_model == "Kimi" and not os.getenv("KIMI_API_KEY"):
            self.log("  警告: Kimi API密钥未设置")
            messagebox.showwarning("API密钥警告", "Kimi API密钥未设置。\n请设置KIMI_API_KEY环境变量。")

    def browse_file(self):
        """打开文件选择对话框选择输入文件"""
        filetypes = [
            ('文本文件', '*.txt'), ('Markdown', '*.md'), ('Word文档', '*.docx'),
            ('PDF文件', '*.pdf'), ('所有文件', '*.*')
        ]
        filename = filedialog.askopenfilename(title="选择文件", filetypes=filetypes)
        if filename:
            self.file_path_var.set(filename)
            self.load_and_preview_file(filename)

    def input_path_dialog(self):
        """手动输入文件路径的对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("手动输入文件路径")
        dialog.geometry("500x100")
        ttk.Label(dialog, text="请输入文件的完整路径:").pack(pady=(10, 5))
        entry = ttk.Entry(dialog, width=60)
        entry.pack(pady=5)
        entry.focus_set()
        
        def confirm():
            path = entry.get().strip()
            if not path:
                messagebox.showerror("错误", "文件路径不能为空")
                return
            
            # 验证路径安全性
            if not self._validate_file_path(path):
                messagebox.showerror("错误", "文件路径无效或不安全")
                return
            
            if os.path.isfile(path):
                self.file_path_var.set(path)
                self.load_and_preview_file(path)
                dialog.destroy()
            else:
                messagebox.showerror("错误", f"文件不存在或路径无效: {path}")
        
        ttk.Button(dialog, text="确定", command=confirm).pack(pady=5)
        dialog.transient(self.root)
        dialog.grab_set()
    
    def _validate_file_path(self, path):
        """验证文件路径的安全性"""
        try:
            # 移除路径中的特殊字符
            if any(char in path for char in '<>"|?*'):
                return False
            
            # 解析路径
            normalized_path = os.path.normpath(path)
            
            # 检查路径长度
            if len(normalized_path) > 260:  # Windows路径长度限制
                return False
            
            return True
        except:
            return False

    def load_and_preview_file(self, filepath):
        """加载文件并预览内容（支持txt/md/docx/pdf格式）"""
        try:
            self.current_file_path = filepath
            self.log(f" 正在加载文件: {filepath}")
            self.current_file_content = self.extract_file_content(filepath)
            
            # 更新预览区域（显示前500字符）
            self.preview_text.delete(1.0, tk.END)
            preview_content = self.current_file_content[:500] + ("..." if len(self.current_file_content) > 500 else "")
            self.preview_text.insert(1.0, preview_content)
            
            # 更新对话选项卡的文件状态
            self.update_chat_file_status()
            
            self.log(f" 文件加载成功，大小: {len(self.current_file_content)} 字符")
            self.generate_btn.config(state=tk.NORMAL)  # 启用生成按钮
            self.save_btn.config(state=tk.DISABLED)    # 禁用保存按钮（未生成内容）
            self.ask_btn.config(state=tk.NORMAL)       # 启用对话提问按钮
        except Exception as e:
            self.log(f" 加载文件失败: {str(e)}")
            messagebox.showerror("错误", f"无法读取文件:\n{str(e)}")

    def update_chat_file_status(self):
        """更新对话选项卡中的文件状态信息"""
        if hasattr(self, 'current_file_path') and self.current_file_path:
            filename = os.path.basename(self.current_file_path)
            self.chat_file_status_var.set(f"已加载文件: {filename}")
            if hasattr(self, 'current_file_content'):
                content_length = len(self.current_file_content)
                self.chat_file_length_var.set(f"文件内容长度: {content_length} 字符")
                # 显示文件摘要
                self.add_to_chat_history("系统", f"已加载文件: {filename}\n文件内容摘要（前200字符）:\n{self.current_file_content[:200]}...")
        else:
            self.chat_file_status_var.set("未加载文件")
            self.chat_file_length_var.set("文件内容长度: 0 字符")

    def reload_file_for_chat(self):
        """重新加载当前文件（用于对话功能）"""
        if hasattr(self, 'current_file_path') and self.current_file_path:
            try:
                self.current_file_content = self.extract_file_content(self.current_file_path)
                self.update_chat_file_status()
                self.add_to_chat_history("系统", "文件已重新加载")
                self.log(f" 为对话功能重新加载文件: {self.current_file_path}")
            except Exception as e:
                self.add_to_chat_history("系统", f"重新加载文件失败: {str(e)}")
        else:
            self.add_to_chat_history("系统", "没有已加载的文件")

    def extract_file_content(self, filepath):
        """提取不同格式文件的文本内容"""
        # 检查文件大小
        max_file_size = 5 * 1024 * 1024  # 5MB限制
        if os.path.getsize(filepath) > max_file_size:
            raise ValueError(f"文件大小超过限制（最大5MB），当前大小：{os.path.getsize(filepath) / 1024 / 1024:.2f}MB")
        
        # 检查文件扩展名
        ext = Path(filepath).suffix.lower()
        supported_extensions = ['.txt', '.md', '.docx', '.pdf']
        if ext not in supported_extensions:
            self.log(f"警告: 非推荐文件格式 {ext}，将尝试作为纯文本读取")
        
        content = ""
        try:
            if ext == '.txt' or ext == '.md':
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
            elif ext == '.docx':
                try:
                    import docx
                    doc = docx.Document(filepath)
                    content = '\n'.join([para.text for para in doc.paragraphs])
                except ImportError:
                    raise ImportError("请安装 python-docx 库以支持 .docx 文件（命令：pip install python-docx）")
            elif ext == '.pdf':
                try:
                    import PyPDF2
                    with open(filepath, 'rb') as f:
                        pdf_reader = PyPDF2.PdfReader(f)
                        # 限制PDF页数
                        max_pages = 50
                        if len(pdf_reader.pages) > max_pages:
                            self.log(f"警告: PDF文件超过{max_pages}页，将只处理前{max_pages}页")
                        for i, page in enumerate(pdf_reader.pages):
                            if i >= max_pages:
                                break
                            page_text = page.extract_text()
                            if page_text:
                                content += page_text + "\n"
                except ImportError:
                    raise ImportError("请安装 PyPDF2 库以支持 .pdf 文件（命令：pip install PyPDF2）")
            else:
                # 其他格式尝试作为纯文本读取
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
        except UnicodeDecodeError:
            # 编码错误时尝试GBK编码
            with open(filepath, 'r', encoding='gbk', errors='ignore') as f:
                content = f.read()
        except Exception as e:
            raise Exception(f"读取文件时出错: {str(e)}")
        
        # 检查内容是否为空
        if not content.strip():
            raise ValueError("文件内容为空或无法提取文本")
        
        # 限制内容长度
        max_content_length = 100000  # 100,000字符
        if len(content) > max_content_length:
            self.log(f"警告: 文件内容超过{max_content_length}字符，将只使用前{max_content_length}字符")
            content = content[:max_content_length]
        
        return content.strip()

    def _call_ai_api(self, ai_model: str, messages: list, max_tokens: int = 2000, temperature: float = 0.7) -> str:
        """统一调用AI API的方法，包含重试机制
        
        Args:
            ai_model: AI模型名称，支持 "DeepSeek" 和 "Kimi"
            messages: 消息列表，格式为 [{"role": "system/user", "content": "消息内容"}]
            max_tokens: 最大 token 数
            temperature: 生成温度
            
        Returns:
            str: AI 生成的内容
            
        Raises:
            Exception: API调用失败时抛出异常
        """
        import time
        max_retries = 2
        retry_delay = 2
        
        for attempt in range(max_retries + 1):
            try:
                if ai_model == "DeepSeek":
                    # 调用DeepSeek API
                    response = self.deepseek_client.chat.completions.create(
                        model="deepseek-chat",
                        messages=messages,
                        max_tokens=max_tokens,
                        temperature=temperature
                    )
                    return response.choices[0].message.content
                else:  # Kimi
                    # 调用Kimi API
                    headers = {
                        "Content-Type": "application/json",
                        "Authorization": f"Bearer {self.kimi_api_key}"
                    }
                    data = {
                        "model": "moonshot-v1-8k",
                        "messages": messages,
                        "max_tokens": max_tokens,
                        "temperature": temperature,
                        "stream": False
                    }
                    response = requests.post(self.kimi_api_url, headers=headers, json=data, timeout=60)
                    response.raise_for_status()
                    result = response.json()
                    return result["choices"][0]["message"]["content"]
            except requests.exceptions.Timeout as e:
                if attempt < max_retries:
                    self.log(f"{ai_model} API调用超时，{retry_delay}秒后重试 ({attempt + 1}/{max_retries})...")
                    time.sleep(retry_delay)
                    retry_delay *= 2  # 指数退避
                else:
                    raise
            except requests.exceptions.RequestException as e:
                if attempt < max_retries and "50" in str(e):  # 只重试服务器错误
                    self.log(f"{ai_model} API请求失败，{retry_delay}秒后重试 ({attempt + 1}/{max_retries})...")
                    time.sleep(retry_delay)
                    retry_delay *= 2  # 指数退避
                else:
                    raise
            except Exception as e:
                if attempt < max_retries:
                    self.log(f"{ai_model}调用失败，{retry_delay}秒后重试 ({attempt + 1}/{max_retries})...")
                    time.sleep(retry_delay)
                    retry_delay *= 2  # 指数退避
                else:
                    raise
    
    def generate_content(self):
        """调用AI生成思维导图/PPT大纲（核心生成功能）"""
        if not self.current_file_content:
            messagebox.showwarning("警告", "请先选择有效的文件")
            return
        
        # 获取配置参数
        output_type = self.output_type_var.get()
        language = self.lang_var.get()
        ai_model = self.ai_model_var.get()
        file_ext = Path(self.current_file_path).suffix if self.current_file_path else "文件"
        
        # 检查API密钥
        if ai_model == "DeepSeek" and (not os.getenv("DEEPSEEK_API_KEY") or not self.deepseek_client):
            messagebox.showerror("错误", "DeepSeek API密钥未设置或初始化失败")
            return
        if ai_model == "Kimi" and not os.getenv("KIMI_API_KEY"):
            messagebox.showerror("错误", "未设置 KIMI_API_KEY 环境变量")
            return
        
        # 构建生成提示词
        if output_type == "思维导图":
            # 解析思维导图层级
            try:
                depth = int(self.depth_var.get().split()[0])
            except:
                depth = 3
            prompt = f"""请将以下`{file_ext}`文件的内容，分析并整理成一个层次清晰、逻辑完整的思维导图大纲。
**要求:**
1. 输出严格的 Markdown 格式，使用 `#` 表示一级主题，`##` 表示二级主题，依此类推。
2. 总共设计大约 {depth} 个层级。
3. 大纲语言使用{language}。
4. 从文件中提炼核心主题、关键概念、重要论点和支撑细节。
5. 结构要符合思维导图的放射性特点，不要写成纯列表。
6. 只输出Markdown内容，不要有任何解释性前缀或后缀。
**文件内容摘要 (前500字符):**
{self.current_file_content[:500]}...
**请开始输出Markdown格式的思维导图大纲:**"""
        else:  # PPT大纲
            pages_range = self.ppt_pages_var.get()
            prompt = f"""请将以下`{file_ext}`文件的内容，分析并整理成一个适合制作PPT演示文稿的详细大纲。
**要求:**
1. 输出严格的 Markdown 格式。
2. 设计一个完整的PPT结构，包含封面页、目录页、内容页和结束页。
3. 内容页数量控制在{pages_range}页左右。
4. 每页PPT使用`##`作为标题，然后列出该页的要点内容。
5. 每个要点使用`-`或`*`符号表示。
6. 大纲语言使用{language}。
7. 从文件中提炼核心内容，确保逻辑连贯、重点突出。
8. 只输出Markdown内容，不要有任何解释性前缀或后缀。
**建议PPT结构示例:**
## 封面页
- 主标题: [根据内容拟定]
- 副标题: [可选]
- 演讲者/日期: [可选]
## 目录页
1. 主要内容一
2. 主要内容二
3. 主要内容三
## 内容页1: [具体标题]
- 要点1
- 要点2
- 要点3
## 内容页2: [具体标题]
- 要点1
- 要点2
## 结束页
- 总结/致谢/联系方式
**文件内容摘要 (前500字符):**
{self.current_file_content[:500]}...
**请开始输出Markdown格式的PPT大纲:**"""
        
        self.log(f"🚀 正在调用 {ai_model} API 生成{output_type}...")
        self.generate_btn.config(state=tk.DISABLED, text="生成中...")
        self.root.update()
        
        try:
            # 构建消息
            messages = [
                {"role": "system", "content": f"你是一个专业的{output_type}设计师，擅长将复杂信息转化为清晰的结构化大纲。"},
                {"role": "user", "content": prompt}
            ]
            
            # 调用API
            ai_output = self._call_ai_api(ai_model, messages)
            
            # 处理生成结果
            self.generated_markdown = ai_output.strip()
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(1.0, self.generated_markdown)
            self.log(f" {ai_model}生成{output_type}成功！共 {len(self.generated_markdown)} 字符")
            self.save_btn.config(state=tk.NORMAL)  # 启用保存按钮
        except requests.exceptions.Timeout:
            error_msg = f"{ai_model} API调用超时"
            self.log(f"{error_msg}")
            messagebox.showerror("生成失败", f"{error_msg}，请稍后重试")
        except requests.exceptions.RequestException as e:
            error_msg = f"{ai_model} API请求失败: {str(e)}"
            self.log(f"{error_msg}")
            messagebox.showerror("生成失败", error_msg)
        except Exception as e:
            error_msg = f"{ai_model}调用失败: {str(e)}"
            self.log(f" {error_msg}")
            messagebox.showerror("生成失败", f"调用{ai_model} API时出错:\n{str(e)}")
        finally:
            # 恢复生成按钮状态
            output_type_text = "思维导图大纲" if output_type == "思维导图" else "PPT大纲"
            self.generate_btn.config(state=tk.NORMAL, text=f"生成{output_type_text}")

    def ask_question(self):
        """基于加载的文件内容向AI提问"""
        question = self.question_var.get().strip()
        if not question:
            messagebox.showwarning("警告", "请输入问题")
            return
        if not hasattr(self, 'current_file_content') or not self.current_file_content:
            messagebox.showwarning("警告", "请先加载文件")
            return
        
        # 获取对话AI模型并检查API
        ai_model = self.chat_ai_model_var.get()
        if ai_model == "DeepSeek" and (not os.getenv("DEEPSEEK_API_KEY") or not self.deepseek_client):
            messagebox.showerror("错误", "DeepSeek API密钥未设置或初始化失败")
            return
        if ai_model == "Kimi" and not os.getenv("KIMI_API_KEY"):
            messagebox.showerror("错误", "未设置 KIMI_API_KEY 环境变量")
            return
        
        # 添加问题到对话历史
        self.add_to_chat_history("你", question)
        self.question_var.set("")
        self.ask_btn.config(state=tk.DISABLED, text="思考中...")
        
        # 新线程处理AI回答（避免阻塞UI）
        thread = threading.Thread(target=self.process_question, args=(question, ai_model))
        thread.daemon = True
        thread.start()

    def process_question(self, question, ai_model):
        """在新线程中调用AI生成回答"""
        try:
            # 构建提问提示词
            prompt = f"""请基于以下文件内容回答用户的问题。如果问题与文件内容无关，请礼貌地说明。
文件内容摘要（前1500字符）:
{self.current_file_content[:1500]}{'...' if len(self.current_file_content) > 1500 else ''}
用户问题: {question}
请提供准确、简洁的回答，并尽可能引用文件中的具体内容。"""
            
            # 构建消息
            messages = [
                {"role": "system", "content": "你是一个专业的文档分析师，擅长基于文档内容回答用户问题。"},
                {"role": "user", "content": prompt}
            ]
            
            # 调用API
            answer = self._call_ai_api(ai_model, messages, max_tokens=1000).strip()
            
            # 主线程更新UI显示回答
            self.root.after(0, self.display_answer, answer, ai_model)
        except requests.exceptions.Timeout:
            error_msg = f"{ai_model} API调用超时"
            self.root.after(0, self.display_answer, error_msg, ai_model)
        except requests.exceptions.RequestException as e:
            error_msg = f"{ai_model} API请求失败: {str(e)}"
            self.root.after(0, self.display_answer, error_msg, ai_model)
        except Exception as e:
            error_msg = f"{ai_model}回答失败: {str(e)}"
            self.root.after(0, self.display_answer, error_msg, ai_model)
        finally:
            # 恢复提问按钮状态
            self.root.after(0, lambda: self.ask_btn.config(state=tk.NORMAL, text="提问"))

    def display_answer(self, answer, ai_model):
        """在对话历史中显示AI的回答"""
        if "API调用超时" in answer or "API请求失败" in answer or "回答失败" in answer:
            self.add_to_chat_history("系统", answer)
        else:
            self.add_to_chat_history(f"{ai_model}助手", answer)
        self.log(f"📢 {ai_model}已回答用户问题")

    def add_to_chat_history(self, sender, message):
        """添加消息到对话历史（带时间戳）"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {sender}:\n{message}\n{'-'*50}\n"
        self.chat_history_text.insert(tk.END, formatted_message)
        self.chat_history_text.see(tk.END)
        # 保存到对话历史列表
        self.conversation_history.append({
            "sender": sender,
            "message": message,
            "timestamp": timestamp
        })

    def use_example_question(self, question):
        """快速使用示例问题提问"""
        self.question_var.set(question)
        self.ask_question()


    def clear_conversation(self):
        """清空对话历史"""
        self.chat_history_text.delete(1.0, tk.END)
        self.conversation_history.clear()
        self.add_to_chat_history("系统", "对话历史已清空")

    def save_as_markdown(self):
        """将生成的大纲保存为Markdown文件"""
        if not self.generated_markdown:
            messagebox.showwarning("警告", "没有可保存的内容")
            return
        

        # 生成默认文件名
        output_type = self.output_type_var.get()
        if self.current_file_path:
            base_name = Path(self.current_file_path).stem
            default_name = f"{base_name}_思维导图.md" if output_type == "思维导图" else f"{base_name}_PPT大纲.md"
        else:
            default_name = f"{output_type}大纲.md"
            

        # 保存文件对话框
        filepath = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown文件", "*.md"), ("文本文件", "*.txt")],
            initialfile=default_name
        )
        if filepath:
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(self.generated_markdown)
                self.log(f" 文件已保存: {filepath}")
                # 保存成功提示
                if output_type == "思维导图":
                    messagebox.showinfo("保存成功", f"思维导图大纲已保存到:\n{filepath}\n\n您可以将此文件导入 XMind 等思维导图软件。")
                else:
                    messagebox.showinfo("保存成功", f"PPT大纲已保存到:\n{filepath}\n\n您可以根据此大纲制作PPT演示文稿。")
            except Exception as e:
                self.log(f" 保存文件失败: {str(e)}")
                messagebox.showerror("保存失败", f"无法保存文件:\n{str(e)}")

    def log(self, message):
        """记录程序日志（带时间戳）"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        if hasattr(self, 'log_text'):
            self.log_text.insert(tk.END, log_entry)
            self.log_text.see(tk.END)
            self.root.update()
        else:
            print(log_entry.strip())

    def clear_log(self):
        """清空程序日志"""
        if hasattr(self, 'log_text'):
            self.log_text.delete(1.0, tk.END)
            self.log("日志已清空。")

    def run(self):
        """启动应用程序"""
        self.root.mainloop()

# 主程序入口
if __name__ == "__main__":
    # 检查必需依赖库
    required_libs = ['openai', 'requests']
    missing_libs = []
    for lib in required_libs:
        try:
            __import__(lib)
        except ImportError:
            missing_libs.append(lib)
    if missing_libs:
        print(f" 缺少必需的库: {', '.join(missing_libs)}")
        print("请使用以下命令安装:")
        print(f"pip install {' '.join(missing_libs)}")
        sys.exit(1)


    # 检查API密钥（提示用户设置）
    if not os.getenv("DEEPSEEK_API_KEY") and not os.getenv("KIMI_API_KEY"):
        print(" 警告: 未设置 DEEPSEEK_API_KEY 或 KIMI_API_KEY 环境变量")
        print("您可以在程序中选择文件，但生成功能需要至少一个API密钥")
        print("\n 设置方法:")
        print("1. 对于Windows:")
        print("   setx DEEPSEEK_API_KEY \"your-api-key\"")
        print("   setx KIMI_API_KEY \"your-api-key\"")
        print("2. 对于macOS/Linux:")
        print("   export DEEPSEEK_API_KEY=\"your-api-key\"")
        print("   export KIMI_API_KEY=\"your-api-key\"")
        print("3. 重启终端或IDE使环境变量生效")
    
    # 检查Selenium（自动上传依赖）
    try:
        __import__('selenium')
        print(" Selenium已安装")
    except ImportError:
        print("  警告: 未安装selenium库，自动上传功能将受限")
        print("如果需要自动上传功能，请使用以下命令安装:")
        print("pip install selenium")
    
    # 检查ChromeDriver（Selenium依赖）
    if os.path.exists("chromedriver.exe"):
        print(" 找到chromedriver.exe")
    else:
        print("  警告: 未找到chromedriver.exe，请下载与Chrome版本匹配的chromedriver")
        print(" 下载地址: https://chromedriver.chromium.org/")
        print(" 下载后请将chromedriver.exe放在程序目录下")
    
    # 启动应用
    print("\n 程序正在启动...")
    app = FileToMindmapApp()
    app.run()