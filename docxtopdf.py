import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageFile, UnidentifiedImageError
import win32com.client
import pythoncom
import webbrowser

# 允许加载损坏的图片
ImageFile.LOAD_TRUNCATED_IMAGES = True

class DocToPdfConverter:
    # 支持的文档格式
    SUPPORTED_DOC_FORMATS = [
        ('Word 文档', '*.doc;*.docx'),
        ('Word 模板', '*.dot;*.dotx'),
        ('启用宏的文档', '*.docm;*.dotm'),
        ('富文本格式', '*.rtf'),
        ('纯文本', '*.txt'),
        ('网页格式', '*.htm;*.html'),
        ('OpenDocument', '*.odt'),
        ('XML 文档', '*.xml'),
        ('PDF 文档', '*.pdf')
    ]

    # 支持的图片格式
    SUPPORTED_IMAGE_FORMATS = [
        ('JPEG 图片', '*.jpg;*.jpeg'),
        ('PNG 图片', '*.png'),
        ('BMP 图片', '*.bmp'),
        ('GIF 图片', '*.gif'),
        ('TIFF 图片', '*.tiff')
    ]

    def __init__(self, master):
        self.master = master
        master.title("文档/图片转PDF工具")
        master.geometry("800x750")
        master.configure(bg="#f0f8ff")  # 浅蓝色背景
        
        # 设置窗口图标
        try:
            master.iconbitmap('pdf_icon.ico')  # 如果有图标文件
        except:
            pass

        # 检测办公软件
        self.office_type = self.detect_office()
        if not self.office_type:
            messagebox.showwarning("警告", "未检测到Microsoft Word或WPS，将仅支持图片转PDF功能")

        # ================= 样式配置 =================
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 主背景色
        self.style.configure(".", background="#f0f8ff", foreground="#333333")
        
        # 框架样式
        self.style.configure("TFrame", background="#f0f8ff")
        self.style.configure("TLabelframe", background="#f0f8ff", bordercolor="#4a90e2", relief=tk.GROOVE)
        self.style.configure("TLabelframe.Label", background="#f0f8ff", foreground="#2c3e50", font=("微软雅黑", 10, "bold"))
        
        # 标签样式
        self.style.configure("TLabel", background="#f0f8ff", font=("微软雅黑", 10))
        
        # 按钮样式
        self.style.configure("TButton", font=("微软雅黑", 10), padding=8, relief=tk.RAISED)
        self.style.map("TButton",
            foreground=[('active', 'white'), ('!active', 'white')],
            background=[('active', '#45aaf2'), ('!active', '#2d98da')],
            bordercolor=[('active', '#45aaf2'), ('!active', '#2d98da')]
        )
        
        # 输入框样式
        self.style.configure("TEntry", fieldbackground="white", font=("微软雅黑", 10), padding=6)
        
        # 树形视图样式
        self.style.configure("Treeview", 
            background="white", 
            foreground="#333333",
            rowheight=25,
            fieldbackground="white"
        )
        self.style.configure("Treeview.Heading", 
            background="#4a90e2", 
            foreground="white",
            font=("微软雅黑", 10, "bold")
        )
        self.style.map("Treeview",
            background=[('selected', '#3498db')],
            foreground=[('selected', 'white')]
        )

        # ================= 主界面布局 =================
        # 标题区域
        self.header_frame = ttk.Frame(master)
        self.header_frame.pack(pady=(20, 10), fill=tk.X)
        
        # 主标题
        ttk.Label(
            self.header_frame,
            text="文档/图片转PDF工具",
            font=("微软雅黑", 18, "bold"),
            foreground="#2c3e50",
            justify="center"
        ).pack(pady=(0, 5))
        
        # 副标题
        ttk.Label(
            self.header_frame,
            text="支持多种文档和图片格式转换为PDF",
            font=("微软雅黑", 12),
            foreground="#7f8c8d",
            justify="center"
        ).pack()

        # 文件选择区域
        self.setup_file_section()

        # 输出路径区域
        self.setup_output_section()

        # 文件列表区域
        self.setup_list_section()

        # 操作按钮区域
        self.setup_action_section()

        # 状态栏
        self.setup_status_bar()

        # 初始化变量
        self.output_path = ""
        self.app = None
        self.current_file = ""
        self.supported_doc_exts = self.generate_supported_extensions(self.SUPPORTED_DOC_FORMATS)
        self.supported_image_exts = self.generate_supported_extensions(self.SUPPORTED_IMAGE_FORMATS)

    def generate_supported_extensions(self, formats):
        """生成带点的扩展名集合"""
        exts = set()
        for _, formats_str in formats:
            if formats_str == '*.*':
                continue
            for ext in formats_str.split(';'):
                exts.add(ext.lower().replace("*", ""))
        return exts

    def detect_office(self):
        """检测可用的办公软件"""
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Quit()
            return "word"
        except:
            try:
                wps = win32com.client.Dispatch("Kwps.Application")
                wps.Quit()
                return "wps"
            except:
                return None

    def setup_file_section(self):
        """文件选择区域"""
        file_frame = ttk.LabelFrame(
            self.master,
            text=" 1. 选择文件 ",
            padding=(15, 10))
        file_frame.pack(pady=10, padx=20, fill=tk.X)

        # 文件路径输入框
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(
            path_frame,
            text="文件路径:",
            font=("微软雅黑", 10, "bold")
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        self.entry_path = ttk.Entry(path_frame, width=50)
        self.entry_path.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        # 文件选择按钮
        button_frame = ttk.Frame(file_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(
            button_frame,
            text="选择文档",
            command=self.select_document,
            width=15
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="选择图片",
            command=self.select_image,
            width=15
        ).pack(side=tk.LEFT, padx=5)

    def setup_output_section(self):
        """输出路径区域"""
        output_frame = ttk.LabelFrame(
            self.master,
            text=" 2. 输出设置 ",
            padding=(15, 10)
        )
        output_frame.pack(pady=10, padx=20, fill=tk.X)

        # 输出路径输入框
        output_path_frame = ttk.Frame(output_frame)
        output_path_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(
            output_path_frame,
            text="输出目录:",
            font=("微软雅黑", 10, "bold")
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        self.entry_output = ttk.Entry(output_path_frame, width=50)
        self.entry_output.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        # 选择目录按钮
        ttk.Button(
            output_path_frame,
            text="浏览...",
            command=self.select_output_path,
            width=10
        ).pack(side=tk.LEFT, padx=5)

    def setup_list_section(self):
        """文件列表区域"""
        list_frame = ttk.LabelFrame(
            self.master,
            text=" 3. 当前文件 ",
            padding=(15, 10))
        list_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        # 创建带滚动条的树形视图
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # 垂直滚动条
        y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 水平滚动条
        x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("filename", "path", "type"),
            show="headings",
            height=5,
            yscrollcommand=y_scroll.set,
            xscrollcommand=x_scroll.set
        )
        self.tree.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        
        # 配置滚动条
        y_scroll.config(command=self.tree.yview)
        x_scroll.config(command=self.tree.xview)
        
        # 设置列
        self.tree.heading("filename", text="文件名", anchor=tk.W)
        self.tree.heading("path", text="路径", anchor=tk.W)
        self.tree.heading("type", text="类型", anchor=tk.W)
        
        self.tree.column("filename", width=200, minwidth=150, stretch=tk.YES)
        self.tree.column("path", width=350, minwidth=200, stretch=tk.YES)
        self.tree.column("type", width=100, minwidth=80, stretch=tk.NO)

    def setup_action_section(self):
        """操作按钮区域"""
        action_frame = ttk.Frame(self.master)
        action_frame.pack(pady=20, padx=20, fill=tk.X)
        
        # 左对齐按钮
        left_frame = ttk.Frame(action_frame)
        left_frame.pack(side=tk.LEFT, expand=True)
        
        ttk.Button(
            left_frame,
            text="项目说明",
            command=self.show_project_info,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            left_frame,
            text="查看源码",
            command=self.view_source_code,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            left_frame,
            text="联系作者",
            command=self.contact_author,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        # 右对齐按钮
        right_frame = ttk.Frame(action_frame)
        right_frame.pack(side=tk.RIGHT, expand=True)
        
        ttk.Button(
            right_frame,
            text="支持格式",
            command=self.show_supported_formats,
            width=15
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            right_frame,
            text="开始转换",
            command=self.start_conversion,
            width=15,
            style="Accent.TButton"
        ).pack(side=tk.RIGHT, padx=5)
        
        # 创建强调按钮样式
        self.style.configure("Accent.TButton", 
            background="#2ecc71", 
            foreground="white"
        )
        self.style.map("Accent.TButton",
            background=[('active', '#27ae60'), ('!active', '#2ecc71')],
            foreground=[('active', 'white'), ('!active', 'white')]
        )

    def setup_status_bar(self):
        """状态栏"""
        status_frame = ttk.Frame(self.master, relief=tk.SUNKEN)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)
        
        self.status_label = ttk.Label(
            status_frame,
            text="就绪",
            font=("微软雅黑", 9),
            foreground="#7f8c8d",
            anchor=tk.W
        )
        self.status_label.pack(fill=tk.X, padx=5)

    def show_project_info(self):
        """显示项目说明"""
        messagebox.showinfo(
            "项目说明",
            "📚 文档/图片转PDF工具\n\n"
            "🔹 项目诞生原因：\n"
            " 由一个懒癌晚期的大学生因为受不了某些软件广告式的转换界面\n"
            " 而打造的一款极简工具\n"
            "🔹 主要功能：\n"
            "• 支持Word、Excel、PPT等多种文档格式转PDF\n"
            "• 支持JPG、PNG等常见图片格式转PDF\n"
            "• 简洁直观的用户界面\n\n"
            "🔹 版本: 1.0.0\n"
            "© 2025 文档转换工具"
        )

    def view_source_code(self):
        """查看源码"""
        result = messagebox.askyesno(
            "查看源码",
            "即将跳转到GitHub查看项目源码，是否继续？"
        )
        if result:
            try:
                webbrowser.open("https://github.com/example/docxtopdf")
            except Exception as e:
                messagebox.showerror("错误", f"无法打开网页: {str(e)}")

    def contact_author(self):
        """联系作者"""
        contact_window = tk.Toplevel(self.master)
        contact_window.title("联系作者")
        contact_window.geometry("300x250")
        contact_window.resizable(False, False)
        
        tk.Label(
            contact_window,
            text="📧 联系方式",
            font=("微软雅黑", 12, "bold"),
            pady=10
        ).pack()
        
        tk.Label(
            contact_window,
            text="QQ: 3864095082",
            font=("微软雅黑", 11),
            pady=5
        ).pack()
        
        tk.Label(
            contact_window,
            text="瓦:马枪手胡图图#92533",
            font=("微软雅黑", 11),
            pady=5
        ).pack()
        
        ttk.Button(
            contact_window,
            text="关闭",
            command=contact_window.destroy,
            width=10
        ).pack(pady=10)

    def update_status(self, message):
        """更新状态栏信息"""
        self.status_label.config(text=message)
        self.master.update()

    def select_document(self):
        """选择文档文件"""
        filetypes = []
        for desc, ext in self.SUPPORTED_DOC_FORMATS:
            if ext == '*.*':
                filetypes.append((desc, ext))
            else:
                ext_tuple = tuple(ext.split(';'))
                filetypes.append((desc, ext_tuple))
        
        path = filedialog.askopenfilename(
            title="选择要转换的文档文件",
            filetypes=filetypes,
            defaultextension="*.*"
        )
        if path:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, path)
            self.current_file = path
            self.update_file_list(path, "文档")
            self.update_status(f"已选择文档: {os.path.basename(path)}")

    def select_image(self):
        """选择图片文件"""
        filetypes = []
        for desc, ext in self.SUPPORTED_IMAGE_FORMATS:
            if ext == '*.*':
                filetypes.append((desc, ext))
            else:
                ext_tuple = tuple(ext.split(';'))
                filetypes.append((desc, ext_tuple))
        
        # 添加一个选项以显示所有支持的图片格式
        all_image_exts = []
        for _, ext in self.SUPPORTED_IMAGE_FORMATS:
            all_image_exts.extend(ext.split(';'))
        all_image_exts = tuple(all_image_exts)
        filetypes.insert(0, ("所有支持的图片格式", all_image_exts))
        
        path = filedialog.askopenfilename(
            title="选择要转换的图片文件",
            filetypes=filetypes,
            defaultextension="*.*"
        )
        if path:
            # 验证图片文件是否有效
            if not self.is_valid_image(path):
                return
            
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, path)
            self.current_file = path
            self.update_file_list(path, "图片")
            self.update_status(f"已选择图片: {os.path.basename(path)}")

    def select_output_path(self):
        """选择输出路径"""
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_path = path
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, path)
            self.update_status(f"输出目录设置为: {path}")

    def update_file_list(self, file_path, file_type):
        """更新文件列表显示"""
        self.tree.delete(*self.tree.get_children())
        if file_path:
            self.tree.insert("", "end", values=(os.path.basename(file_path), file_path, file_type))

    def show_supported_formats(self):
        """显示支持格式"""
        doc_formats = "\n".join([
            f"• {desc} ({ext.replace(';', ', ')})"
            for desc, ext in self.SUPPORTED_DOC_FORMATS
        ])
        image_formats = "\n".join([
            f"• {desc} ({ext.replace(';', ', ')})"
            for desc, ext in self.SUPPORTED_IMAGE_FORMATS
        ])
        
        messagebox.showinfo(
            "支持格式",
            f"📄 文档支持格式：\n{doc_formats}\n\n"
            f"🖼️ 图片支持格式：\n{image_formats}\n\n"
            f"💻 办公软件: {'WPS' if self.office_type == 'wps' else 'Microsoft Word' if self.office_type else '无'}\n"
            "ℹ️ 注意：图片转换使用PIL库，文档转换使用Office组件"
        )

    def is_valid_image(self, file_path):
        """验证图片文件是否有效"""
        # 检查文件是否存在
        if not os.path.exists(file_path):
            messagebox.showerror("错误", f"文件不存在: {file_path}")
            self.update_status("错误: 文件不存在")
            return False

        try:
            with Image.open(file_path) as img:
                # 验证文件内容
                img.verify()
                # 重新打开图片以加载数据
                with Image.open(file_path) as img:
                    img.load()  # 加载图片数据
            return True
        except (IOError, UnidentifiedImageError, Image.DecompressionBombError) as e:
            messagebox.showerror("错误", f"无效的图片文件: {str(e)}")
            self.update_status("错误: 无效的图片文件")
            return False

    def convert_image_to_pdf(self, image_path, output_path):
        """将图片转换为PDF"""
        self.update_status(f"正在转换图片到PDF: {os.path.basename(image_path)}...")
        try:
            with Image.open(image_path) as img:
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # 保持原始图片质量
                img.save(
                    output_path, 
                    "PDF", 
                    resolution=100.0,
                    quality=95,
                    save_all=True if hasattr(img, 'is_animated') and img.is_animated else False
                )
            self.update_status(f"图片转换成功: {os.path.basename(image_path)}")
            return True
        except Exception as e:
            messagebox.showerror("错误", f"图片转换失败: {str(e)}")
            self.update_status(f"错误: 图片转换失败 - {str(e)}")
            return False

    def start_conversion(self):
        """开始转换"""
        if not self.current_file:
            messagebox.showwarning("警告", "请先选择要转换的文件！")
            self.update_status("警告: 未选择文件")
            return
        if not self.output_path:
            messagebox.showwarning("警告", "请先选择输出目录！")
            self.update_status("警告: 未选择输出目录")
            return

        try:
            # 获取规范化的文件扩展名
            file_ext = os.path.splitext(self.current_file)[1].lower()
            
            # 生成输出路径
            filename = os.path.basename(self.current_file)
            pdf_name = os.path.splitext(filename)[0] + ".pdf"
            pdf_path = os.path.join(self.output_path, pdf_name)

            # 判断文件类型
            is_image = file_ext in self.supported_image_exts

            if is_image:
                # 图片转PDF
                if not self.convert_image_to_pdf(self.current_file, pdf_path):
                    return
            else:
                # 文档转PDF
                if not self.office_type:
                    messagebox.showerror("错误", "未检测到Office软件，无法转换文档！")
                    self.update_status("错误: 未检测到Office软件")
                    return

                self.update_status(f"正在转换文档到PDF: {filename}...")
                pythoncom.CoInitialize()
                try:
                    app_name = "Kwps.Application" if self.office_type == "wps" else "Word.Application"
                    self.app = win32com.client.Dispatch(app_name)
                    self.app.Visible = False

                    doc = self.app.Documents.Open(self.current_file)
                    doc.SaveAs(pdf_path, FileFormat=17)  # 17是PDF格式
                    doc.Close()
                    self.update_status(f"文档转换成功: {filename}")
                except Exception as e:
                    messagebox.showerror("错误", f"文档转换失败: {str(e)}")
                    self.update_status(f"错误: 文档转换失败 - {str(e)}")
                    return
                finally:
                    if hasattr(self, 'app') and self.app:
                        self.app.Quit()
                    pythoncom.CoUninitialize()

            # 转换成功后询问
            if messagebox.askyesno("完成", f"✅ {filename} 转换成功！\n是否继续转换其他文件？"):
                self.current_file = ""
                self.entry_path.delete(0, tk.END)
                self.update_file_list("", "")
                self.update_status("就绪 - 等待选择新文件")
            else:
                self.master.destroy()

        except Exception as e:
            messagebox.showerror("错误", f"转换过程中发生错误: {str(e)}")
            self.update_status(f"错误: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocToPdfConverter(root)
    root.mainloop()