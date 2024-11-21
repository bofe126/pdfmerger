import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import os
import sys
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import tempfile
import subprocess

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF合并工具")
        self.root.geometry("1200x900")
        
        # 设置 DPI 感知以提高文字清晰度
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
        
        # 设置主题样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # Adobe 红色系列配色 - 柔和版本
        self.colors = {
            'primary': '#C62828',       # 深红主色
            'secondary': '#FFEBEE',     # 超浅红背景
            'accent': '#B71C1C',        # 暗红强调色
            'text': '#37474F',          # 深蓝灰文本
            'light_text': '#78909C',    # 浅蓝灰文本
            'border': '#FFCDD2',        # 柔和红边框
            'hover': '#D32F2F',         # 鲜红悬停
            'bg': '#FFFFFF',            # 纯白背景
            'selected': '#EF9A9A',      # 浅红选中
            'card_bg': '#FAFAFA',       # 灰白卡片背景
            'disabled': '#BDBDBD',      # 禁用状态
        }
        
        # 配置全局样式
        style.configure('.',
            background=self.colors['bg'],
            foreground=self.colors['text'],
            font=('Microsoft YaHei UI', 10)
        )
        
        # 配置按钮样式
        style.configure('Primary.TButton',
            background=self.colors['primary'],
            foreground='white',
            padding=(20, 10),
            font=('Microsoft YaHei UI', 10, 'bold'),
            borderwidth=0
        )
        style.map('Primary.TButton',
            background=[
                ('active', self.colors['hover']),
                ('disabled', self.colors['disabled'])
            ],
            foreground=[
                ('disabled', '#FFFFFF')
            ]
        )
        
        # 配置标签样式
        style.configure('Title.TLabel',
            font=('Microsoft YaHei UI', 16, 'bold'),
            foreground=self.colors['primary'],
            padding=(0, 10)
        )
        
        # 配置页码标签样式
        style.configure('Page.TLabel',
            font=('Microsoft YaHei UI', 10),
            foreground=self.colors['light_text']
        )
        
        # 配置框架样式
        style.configure('Card.TFrame',
            background=self.colors['card_bg'],
            relief='solid',
            borderwidth=1,
            bordercolor=self.colors['border']
        )
        
        # 设置窗口图标
        try:
            if getattr(sys, 'frozen', False):
                # 打包后的路径
                icon_path = os.path.join(sys._MEIPASS, 'icon.png')
            else:
                # 开发时的路径
                icon_path = 'icon.png'
                
            if os.path.exists(icon_path):
                # 同时设置 iconbitmap 和 iconphoto
                img = Image.open(icon_path)
                # 创建临时 ICO 文件
                ico_path = os.path.join(tempfile.gettempdir(), 'temp_icon.ico')
                img.save(ico_path, format='ICO', sizes=[(32, 32)])
                self.root.iconbitmap(ico_path)
                
                # 设置窗口图标
                icon = tk.PhotoImage(file=icon_path)
                self.root.iconphoto(True, icon)
                
                # 清理临时文件
                try:
                    os.remove(ico_path)
                except:
                    pass
                    
        except Exception as e:
            print(f"加载图标失败: {e}")
        
        self.pdf_files = []
        self.current_preview = None
        self.current_page = 0
        self.file_rotations = {}
        self.preview_cache = {}  # 添加预览缓存
        
        # 绑定主窗口大小变化事件
        self.root.bind('<Configure>', self.on_window_resize)
        self.last_window_size = None
        
        self.setup_ui()
        
    def get_poppler_path(self):
        if getattr(sys, 'frozen', False):
            return os.path.join(sys._MEIPASS, 'poppler')
        return r"D:\Program Files\poppler-24.08.0\Library\bin"  # 开发时的路径
        
    def setup_ui(self):
        # 获取 style 对象
        style = ttk.Style()
        
        # 主容器使用网格布局
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=2)
        self.root.grid_rowconfigure(0, weight=1)
        
        # 左侧面板
        left_frame = ttk.Frame(self.root, padding="20")
        left_frame.grid(row=0, column=0, sticky="nsew")
        
        # 标题
        ttk.Label(
            left_frame, 
            text="PDF文件列表",
            style='Title.TLabel'
        ).pack(fill='x', pady=(0, 20))
        
        # 按钮区域
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill='x', pady=(0, 20))
        
        add_btn = ttk.Button(
            btn_frame,
            text="添加文件",
            style='Primary.TButton',
            command=self.add_files
        )
        add_btn.pack(side='left', padx=(0, 10))
        
        merge_btn = ttk.Button(
            btn_frame,
            text="合并PDF",
            style='Primary.TButton',
            command=self.merge_pdfs
        )
        merge_btn.pack(side='left')
        
        # 文件列表区域（使用卡片样式）
        list_frame = ttk.Frame(left_frame, style='Card.TFrame')
        list_frame.pack(fill='both', expand=True)
        
        # 创建带滚动条的文件列表容器
        list_container = ttk.Frame(list_frame)
        list_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 配置滚动条样式
        style.configure("Custom.Vertical.TScrollbar",
            background=self.colors['bg'],
            arrowcolor=self.colors['primary'],
            troughcolor=self.colors['secondary'],
            width=10,
            relief="flat"
        )
        
        style.map("Custom.Vertical.TScrollbar",
            background=[
                ('pressed', self.colors['primary']),
                ('active', self.colors['hover'])
            ]
        )
        
        # 文件列表
        self.file_listbox = tk.Listbox(
            list_container,
            selectmode=tk.SINGLE,
            font=('Microsoft YaHei UI', 10),
            bg=self.colors['bg'],
            fg=self.colors['text'],
            selectbackground=self.colors['selected'],
            selectforeground=self.colors['text'],
            activestyle='none',
            relief='flat',
            borderwidth=0,
            highlightthickness=0
        )
        
        # 滚动条
        scrollbar = ttk.Scrollbar(
            list_container, 
            orient="vertical",
            style="Custom.Vertical.TScrollbar",
            command=self.file_listbox.yview
        )
        self.file_listbox.config(
            yscrollcommand=scrollbar.set,
            highlightbackground=self.colors['border'],
            highlightcolor=self.colors['primary']
        )
        
        # 使用网格布局放置列表和滚动条
        self.file_listbox.grid(row=0, column=0, sticky="nsew", padx=(0, 5))  # 添加右边距
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # 配置网格权重
        list_container.grid_columnconfigure(0, weight=1)
        list_container.grid_rowconfigure(0, weight=1)
        
        # 绑定选择事件
        self.file_listbox.bind('<<ListboxSelect>>', self.on_select_file)
        
        # 文件操作按钮
        control_frame = ttk.Frame(left_frame)
        control_frame.pack(fill='x', pady=(20, 0))
        
        for text, command in [
            ("上移", self.move_up),
            ("下移", self.move_down),
            ("删除", self.remove_file)
        ]:
            btn = ttk.Button(
                control_frame,
                text=text,
                command=command,
                style='Primary.TButton',
                width=6
            )
            btn.pack(side='left', padx=(0, 10))
        
        # 右侧预览面板
        right_frame = ttk.Frame(self.root, padding="20")
        right_frame.grid(row=0, column=1, sticky="nsew")
        right_frame.grid_columnconfigure(0, weight=1)  # 添加列权重
        right_frame.grid_rowconfigure(2, weight=1)     # 预览区域行权重
        
        # 预览标题
        ttk.Label(
            right_frame,
            text="预览",
            style='Title.TLabel'
        ).grid(row=0, column=0, sticky="w")  # 使用网格布局
        
        # 预览区域（使用卡片样式）
        preview_frame = ttk.Frame(right_frame, style='Card.TFrame')
        preview_frame.grid(row=2, column=0, sticky="nsew")
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)  # 添加行权重
        
        # 创建一个容器来包含预览和控制区域
        container = ttk.Frame(preview_frame)
        container.pack(fill='both', expand=True, padx=20, pady=20)
        container.grid_rowconfigure(0, weight=1)  # 预览区域可伸缩
        container.grid_rowconfigure(1, weight=0)  # 控制区域固定高度
        container.grid_columnconfigure(0, weight=1)  # 添加列权重
        
        # 创建预览容器，设置较小的固定高度
        self.preview_container = ttk.Frame(container, style='Card.TFrame', width=600, height=650)
        self.preview_container.grid(row=0, column=0, sticky="")  # 移除 nsew，使用默认居中
        self.preview_container.grid_propagate(False)  # 禁止自动调整大小
        
        # 预览标签
        self.preview_label = ttk.Label(self.preview_container)
        self.preview_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # 底部控制区域
        control_frame = ttk.Frame(container)
        control_frame.grid(row=1, column=0, pady=(20, 0))  # 添加顶部间距
        
        # 页码显示（放在按钮上方）
        self.page_label = ttk.Label(
            control_frame,
            font=('Microsoft YaHei UI', 10),
            foreground=self.colors['light_text']
        )
        self.page_label.pack(pady=(0, 10))
        
        # 按钮容器（居中）
        button_container = ttk.Frame(control_frame)
        button_container.pack(anchor='center')
        
        # 添加按钮
        for text, command in [
            ("上一页", self.prev_page),
            ("下一页", self.next_page),
            ("旋转", self.rotate_page)
        ]:
            btn = ttk.Button(
                button_container,
                text=text,
                command=command,
                style='Primary.TButton',
                width=8
            )
            btn.pack(side='left', padx=5)
    
    def on_window_resize(self, event):
        """只处理主窗口大小变化事件"""
        if event.widget == self.root:  # 只响应主窗口的大小变化
            new_size = (self.root.winfo_width(), self.root.winfo_height())
            
            # 如果主窗口尺发生变化且有当前览的文件
            if new_size != self.last_window_size and self.file_listbox.curselection():
                self.last_window_size = new_size
                # 重新加载预览
                idx = self.file_listbox.curselection()[0]
                self.preview_pdf(self.pdf_files[idx])
    
    def add_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("PDF files", "*.pdf")]
        )
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
    
    def move_up(self):
        idx = self.file_listbox.curselection()
        if not idx or idx[0] == 0:
            return
        
        current_idx = idx[0]
        # 交换文件列表中的位置
        self.pdf_files[current_idx], self.pdf_files[current_idx-1] = \
            self.pdf_files[current_idx-1], self.pdf_files[current_idx]
            
        # 交换显示列表中的位置
        current_text = self.file_listbox.get(current_idx)
        above_text = self.file_listbox.get(current_idx-1)
        
        self.file_listbox.delete(current_idx)
        self.file_listbox.delete(current_idx-1)
        
        self.file_listbox.insert(current_idx-1, current_text)
        self.file_listbox.insert(current_idx, above_text)
        
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.selection_set(current_idx-1)
    
    def move_down(self):
        idx = self.file_listbox.curselection()
        if not idx or idx[0] == self.file_listbox.size()-1:
            return
        
        current_idx = idx[0]
        # 交换文件列表中的位置
        self.pdf_files[current_idx], self.pdf_files[current_idx+1] = \
            self.pdf_files[current_idx+1], self.pdf_files[current_idx]
            
        # 交换显示列表中的位置
        current_text = self.file_listbox.get(current_idx)
        below_text = self.file_listbox.get(current_idx+1)
        
        self.file_listbox.delete(current_idx+1)
        self.file_listbox.delete(current_idx)
        
        self.file_listbox.insert(current_idx, below_text)
        self.file_listbox.insert(current_idx+1, current_text)
        
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.selection_set(current_idx+1)
    
    def remove_file(self):
        idx = self.file_listbox.curselection()
        if not idx:
            return
        
        file_path = self.pdf_files[idx[0]]
        self.pdf_files.pop(idx[0])
        self.file_listbox.delete(idx[0])
        self.preview_label.configure(image='')
        self.page_label.configure(text='')
        
        # 清理相关缓存
        if file_path in self.file_rotations:
            del self.file_rotations[file_path]
        # 清理预览缓存
        self.preview_cache = {k: v for k, v in self.preview_cache.items() 
                            if not k.startswith(file_path)}
    
    def on_select_file(self, event):
        idx = self.file_listbox.curselection()
        if not idx:
            return
        
        file_path = self.pdf_files[idx[0]]
        self.current_page = 0
        self.preview_pdf(file_path)
    
    def preview_pdf(self, pdf_path):
        try:
            # 检查缓存
            cache_key = f"{pdf_path}_{self.current_page}"
            if cache_key in self.preview_cache:
                img = self.preview_cache[cache_key]
                # 获取当前页面的旋转角度
                current_rotation = self.file_rotations.get(pdf_path, {}).get(self.current_page, 0)
                if current_rotation:
                    img = img.rotate(-current_rotation, expand=True, resample=Image.Resampling.BICUBIC)
            else:
                # 创建临时文件路径
                temp_img_path = os.path.join(tempfile.gettempdir(), 'temp_preview.png')
                
                # 建 STARTUPINFO 对象来隐藏窗口
                startupinfo = None
                if os.name == 'nt':  # Windows 系统
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                
                # 使用 pdftoppm 命令行工具转换 PDF 到图像
                poppler_path = self.get_poppler_path()
                pdftoppm_path = os.path.join(poppler_path, 'pdftoppm.exe')
                
                # 构建命令，降低 DPI 以提高速度
                cmd = [
                    pdftoppm_path,
                    '-f', str(self.current_page + 1),
                    '-l', str(self.current_page + 1),
                    '-png',
                    '-singlefile',
                    '-r', '150',  # 降低 DPI
                    pdf_path,
                    os.path.splitext(temp_img_path)[0]
                ]
                
                # 执行命令
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    startupinfo=startupinfo
                )
                process.wait()
                
                # 检查输出文件是否存在
                output_file = f"{os.path.splitext(temp_img_path)[0]}.png"
                if os.path.exists(output_file):
                    img = Image.open(output_file)
                    # 保存到缓存
                    self.preview_cache[cache_key] = img.copy()
                    
                    # 获取当前页面的旋转角度
                    current_rotation = self.file_rotations.get(pdf_path, {}).get(self.current_page, 0)
                    if current_rotation:
                        img = img.rotate(-current_rotation, expand=True, resample=Image.Resampling.BICUBIC)
                    
                    # 清理临时文件
                    try:
                        os.remove(output_file)
                    except:
                        pass
            
            # 调整图像大小
            preview_width = self.preview_container.winfo_width()
            preview_height = self.preview_container.winfo_height()
            
            if preview_width > 1 and preview_height > 1:
                img_ratio = img.width / img.height
                container_ratio = preview_width / preview_height
                
                if img_ratio > container_ratio:
                    new_width = min(preview_width, 800)
                    new_height = int(new_width / img_ratio)
                else:
                    new_height = min(preview_height, 800)
                    new_width = int(new_height * img_ratio)
                
                # 使用更快的重采样方法
                img = img.resize((new_width, new_height), 
                               Image.Resampling.BILINEAR)  # 使用双线性插值
            
            photo = ImageTk.PhotoImage(img)
            self.preview_label.configure(
                image=photo,
                compound='center',
                anchor='center'
            )
            self.preview_label.image = photo
            
            total_pages = len(PdfReader(pdf_path).pages)
            self.page_label.configure(
                text=f"第 {self.current_page + 1} 页，共 {total_pages} 页",
                font=('Microsoft YaHei UI', 10)
            )
        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {str(e)}")
    
    def prev_page(self):
        if not self.file_listbox.curselection():
            return
            
        idx = self.file_listbox.curselection()[0]
        pdf = PdfReader(self.pdf_files[idx])
        
        if self.current_page > 0:
            self.current_page -= 1
            self.preview_pdf(self.pdf_files[idx])
    
    def next_page(self):
        if not self.file_listbox.curselection():
            return
            
        idx = self.file_listbox.curselection()[0]
        pdf = PdfReader(self.pdf_files[idx])
        
        if self.current_page < len(pdf.pages) - 1:
            self.current_page += 1
            self.preview_pdf(self.pdf_files[idx])
    
    def rotate_page(self):
        if not self.file_listbox.curselection():
            return
            
        idx = self.file_listbox.curselection()[0]
        file_path = self.pdf_files[idx]
        
        # 初始化文件的旋转记录（如果不存在）
        if file_path not in self.file_rotations:
            self.file_rotations[file_path] = {}
        
        # 更新当前页面的旋转角度
        current_rotation = self.file_rotations[file_path].get(self.current_page, 0)
        new_rotation = (current_rotation + 90) % 360
        self.file_rotations[file_path][self.current_page] = new_rotation
        
        self.preview_pdf(file_path)
    
    def merge_pdfs(self):
        if not self.pdf_files:
            messagebox.showwarning("警告", "请至少添加一个PDF文件")
            return
            
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if not output_path:
            return
            
        try:
            pdf_writer = PdfWriter()
            
            # 如果只有一个文件，直接处理旋转并保存
            if len(self.pdf_files) == 1:
                pdf_file = self.pdf_files[0]
                pdf_reader = PdfReader(pdf_file)
                page_rotations = self.file_rotations.get(pdf_file, {})
                
                for page_num, page in enumerate(pdf_reader.pages):
                    # 获取该页的旋转角度（如果有）
                    rotation = page_rotations.get(page_num, 0)
                    if rotation:
                        page.rotate(rotation)
                    pdf_writer.add_page(page)
                    
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                    
                messagebox.showinfo("成功", "PDF导出完成！")
                return
            
            # 多个文件的合并逻辑保持不变
            for pdf_file in self.pdf_files:
                pdf_reader = PdfReader(pdf_file)
                page_rotations = self.file_rotations.get(pdf_file, {})
                
                for page_num, page in enumerate(pdf_reader.pages):
                    # 获取该页的旋转角度（如果有）
                    rotation = page_rotations.get(page_num, 0)
                    if rotation:
                        page.rotate(rotation)
                    pdf_writer.add_page(page)
            
            with open(output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
                
            messagebox.showinfo("成功", "PDF合并完成！")
        except Exception as e:
            messagebox.showerror("错误", f"操作失败: {str(e)}")
    
    def convert_png_to_ico(self, png_path):
        """将PNG转换为ICO格式，支持多种尺寸以提高清晰度"""
        img = Image.open(png_path)
        
        # 准备多个尺寸的图标
        icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        icons = []
        
        for size in icon_sizes:
            # 创建一个新的透明背景图像
            icon = Image.new('RGBA', size, (0, 0, 0, 0))
            
            # 计算保持宽高比的新尺寸
            img_copy = img.copy()
            img_copy.thumbnail(size, Image.Resampling.LANCZOS)
            
            # 计算居中位置
            offset = ((size[0] - img_copy.width) // 2,
                     (size[1] - img_copy.height) // 2)
            
            # 将调整后的图像粘贴到透明背景上
            icon.paste(img_copy, offset, img_copy if img_copy.mode == 'RGBA' else None)
            icons.append(icon)
        
        # 在临时目录创建ico文件
        if getattr(sys, 'frozen', False):
            ico_path = os.path.join(tempfile.gettempdir(), 'temp_icon.ico')
        else:
            ico_path = os.path.splitext(png_path)[0] + '.ico'
        
        # 保存多尺寸图标
        icons[0].save(
            ico_path, 
            format='ICO', 
            sizes=[(icon.width, icon.height) for icon in icons],
            append_images=icons[1:]
        )
        return ico_path
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        # 先选中点击的项
        index = self.file_listbox.nearest(event.y)
        if index >= 0:
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(index)
            self.file_listbox.activate(index)
            # 显示菜单
            self.context_menu.post(event.x_root, event.y_root)
    
    def on_drag_start(self, event):
        """开始拖拽"""
        # 获取点击的项
        index = self.file_listbox.nearest(event.y)
        if index >= 0:
            # 保存拖拽数据
            self.drag_data['item'] = self.file_listbox.get(index)
            self.drag_data['index'] = index
            self.drag_data['y'] = event.y
            # 选中该项
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(index)
            self.file_listbox.activate(index)
    
    def on_drag_motion(self, event):
        """拖拽过程中"""
        if self.drag_data['item']:
            # 计算当前位置
            cur_index = self.file_listbox.nearest(event.y)
            if cur_index >= 0 and cur_index != self.drag_data['index']:
                # 移动项目
                old_index = self.drag_data['index']
                self.move_item(old_index, cur_index)
                self.drag_data['index'] = cur_index
    
    def on_drag_release(self, event):
        """结束拖拽"""
        self.drag_data = {'item': None, 'index': None, 'y': 0}
    
    def move_item(self, old_index, new_index):
        """移动列表项"""
        if old_index == new_index:
            return
            
        # 移动文件列表中的项
        item = self.pdf_files.pop(old_index)
        self.pdf_files.insert(new_index, item)
        
        # 移动显示列表中的项
        text = self.file_listbox.get(old_index)
        self.file_listbox.delete(old_index)
        self.file_listbox.insert(new_index, text)
        
        # 保持选中状态
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.selection_set(new_index)
        self.file_listbox.activate(new_index)
        
        # 更新预览
        self.preview_pdf(self.pdf_files[new_index])

# 在程序开始时转换图标
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop() 