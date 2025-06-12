# --- 安装说明 ---
# 本程序需要以下库来处理Word文档。
# 请在你的命令行/终端中运行以下命令：
# pip install docx2pdf PyMuPDF

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import re
import os
import csv
from collections import defaultdict
import typing

# --- 依赖导入与检查 ---
try:
    import fitz  # PyMuPDF
    from docx2pdf import convert
    FILE_PROCESSING_AVAILABLE = True
except ImportError:
    FILE_PROCESSING_AVAILABLE = False

# --- 全局常量 ---
STOP_WORDS = {
    'a', 'about', 'above', 'after', 'again', 'against', 'all', 'am', 'an', 'and', 'any', 'are', "aren't", 'as', 'at',
    'be', 'because', 'been', 'before', 'being', 'below', 'between', 'both', 'but', 'by', "can't", 'cannot', 'could',
    "couldn't", 'did', "didn't", 'do', 'does', "doesn't", 'doing', "don't", 'down', 'during', 'each', 'few', 'for',
    'from', 'further', 'had', "hadn't", 'has', "hasn't", 'have', "haven't", 'having', 'he', "he'd", "he'll", "he's",
    'her', 'here', "here's", 'hers', 'herself', 'him', 'himself', 'his', 'how', "how's", 'i', "i'd", "i'll", "i'm",
    "i've", 'if', 'in', 'into', 'is', "isn't", 'it', "it's", 'its', 'itself', "let's", 'me', 'more', 'most',
    "mustn't", 'my', 'myself', 'no', 'nor', 'not', 'of', 'off', 'on', 'once', 'only', 'or', 'other', 'ought', 'our',
    'ours', 'ourselves', 'out', 'over', 'own', 'same', "shan't", 'she', "she'd", "she'll", "she's", 'should',
    "shouldn't", 'so', 'some', 'such', 'than', 'that', "that's", 'the', 'their', 'theirs', 'them', 'themselves',
    'then', 'there', "there's", 'these', 'they', "they'd", "they'll", "they're", "they've", 'this', 'those',
    'through', 'to', 'too', 'under', 'until', 'up', 'very', 'was', "wasn't", 'we', "we'd", "we'll", "we're", "we've",
    'were', "weren't", 'what', "what's", 'when', "when's", 'where', "where's", 'which', 'while', 'who', "who's",
    'whom', 'why', "why's", 'with', "won't", 'would', "wouldn't", 'you', "you'd", "you'll", "you're", "you've",
    'your', 'yours', 'yourself', 'yourselves'
}
WORD_REGEX = r'\b[a-zA-Z-./]+(?:\s*/\s*[a-zA-Z-./]+)*\b'

# ==============================================================================
# 核心业务逻辑层 (Backend)
# ==============================================================================
class IndexerBackend:
    """封装所有核心处理逻辑的类"""
    def __init__(self, status_callback: typing.Callable, progress_callback: typing.Callable):
        self.status_callback = status_callback
        self.progress_callback = progress_callback

    def _normalize_term(self, term: str) -> str:
        """规范化术语：转小写，合并斜杠空格"""
        return re.sub(r'\s*/\s*', '/', term).lower()

    def _extract_terms(self, text: str, mode: str) -> typing.List[str]:
        """根据模式从文本中提取术语"""
        raw_words = re.findall(WORD_REGEX, text)
        normalized_words = [self._normalize_term(word) for word in raw_words]

        if mode == "words":
            return [word for word in normalized_words if word not in STOP_WORDS]
        if mode == "words_no_filter":
            return normalized_words
        if mode == "phrases":
            phrases = re.findall(r'\b[A-Z][a-zA-Z]*(?:\s+(?:of|the|and|for|in|to|on)\s+[A-Z][a-zA-Z]*)*'
                                 r'(?:\s+[A-Z][a-zA-Z]*)*\b', text)
            normalized_phrases = [p.lower() for p in phrases]
            filtered_words = [word for word in normalized_words if word not in STOP_WORDS]
            return list(set(normalized_phrases + filtered_words))
        return []

    def extract_from_docx(self, docx_path: str, mode: str) -> typing.Dict[str, typing.Set[int]]:
        """从DOCX文件中提取术语和其所在的页码"""
        if not FILE_PROCESSING_AVAILABLE:
            raise ImportError("文件处理库 (PyMuPDF, docx2pdf) 未安装。")
        
        term_map = defaultdict(set)
        temp_pdf_path = ""
        try:
            self.status_callback("状态：正在将DOCX转换为PDF...")
            self.progress_callback(0, 100)
            
            temp_pdf_path = os.path.splitext(docx_path)[0] + "_temp.pdf"
            convert(docx_path, temp_pdf_path)
            
            self.status_callback("状态：正在从PDF中提取文本和页码...")
            pdf_doc = fitz.open(temp_pdf_path)
            num_pages = len(pdf_doc)
            for i, page in enumerate(pdf_doc):
                text = page.get_text("text")
                terms = self._extract_terms(text, mode)
                for term in terms:
                    term_map[term].add(i + 1)
                self.progress_callback(int((i + 1) / num_pages * 100), 100)
            
            pdf_doc.close()
            return term_map
        finally:
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)

    def save_results_as_txt(self, output_path: str, term_map: dict):
        """将索引结果保存为TXT文件，内容为Typst格式"""
        self.status_callback("状态：正在生成Typst格式的TXT文件...")
        
        # 按首字母对术语进行分组
        grouped_terms = defaultdict(list)
        for term in sorted(term_map.keys()):
            first_letter = term[0].upper()
            if 'A' <= first_letter <= 'Z':
                grouped_terms[first_letter].append(term)
            else:
                grouped_terms['#'].append(term) # 将非字母开头的归为一类
        
        with open(output_path, 'w', encoding='utf-8') as f:
            # 写入Typst文件头
            f.write('// ===============================================\n')
            f.write('// Typst Index Layout by Mutro\n')
            f.write('// Based on previously approved format\n')
            f.write('// ===============================================\n\n')
            f.write('// 1. Document & Page Setup\n')
            f.write('#set document(title: "Index table", author: "Mutro")\n')
            f.write('#set page(paper: "a4", margin: (x: 2cm, y: 2.5cm))\n\n')
            f.write('// 2. Font Setup\n')
            f.write('#set text(font: ("New Computer Modern", "Noto Serif CJK SC"), size: 10pt)\n\n')
            f.write('// 3. Title Section\n')
            f.write('#align(center)[\n')
            f.write('  #text(size: 24pt, weight: 600)[Index]\n')
            f.write(']\n#v(2em)\n\n')
            f.write('// 4. Main Content\n')
            f.write('#columns(2, gutter: 1.5em)[\n')

            # 按字母顺序写入术语
            sorted_letters = sorted(grouped_terms.keys())
            if '#' in sorted_letters:
                sorted_letters.remove('#')
                sorted_letters.append('#')

            for i, letter in enumerate(sorted_letters):
                # 【修复】确保每个字母块之间有空行，但第一个块之前没有
                if i > 0:
                    f.write('\n')
                
                # 【修复】对'#'进行转义以避免Typst语法错误
                display_letter = "\#" if letter == "#" else letter
                
                f.write(f'  // --- {letter} ---\n')
                f.write(f'  #text(size: 18pt, weight: "bold")[{display_letter}]\n')
                f.write('  #line(length: 100%)\n')
                f.write('  #v(0.8em)\n')
                
                for term in grouped_terms[letter]:
                    pages = ", ".join(map(str, sorted(list(term_map[term]))))
                    # 替换双引号以避免Typst语法错误
                    escaped_term = term.replace('"', '\\"')
                    # 严格遵循参考格式，使用[...]并添加强制换行符 `\\`
                    f.write(f'  #text(weight: "bold")[{escaped_term}]: {pages}\\\n')
            
            # 确保结尾只有一个换行符
            f.write(']')
        
        self.status_callback("状态：索引文件已成功保存！")

# ==============================================================================
# 图形用户界面层 (Frontend)
# ==============================================================================
class SimplifiedIndexerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("索引生成器 BY MUTRO")
        self.root.geometry("550x300")
        self.root.minsize(500, 300)
        
        self.input_filepath = tk.StringVar()
        self.extraction_mode = tk.StringVar(value="words")
        self.processing_thread = None

        self.backend = IndexerBackend(self.update_status, self.update_progress)

        self._create_widgets()

    def _create_widgets(self):
        """创建所有UI控件"""
        main_frame = ttk.Frame(self.root, padding="15 15 15 15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 步骤 1: 文件选择 ---
        file_frame = ttk.LabelFrame(main_frame, text=" 步骤 1: 选择文件 ", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        file_button = ttk.Button(file_frame, text="选择Word文件 (.docx)", command=self.select_file)
        file_button.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.file_label = ttk.Label(file_frame, text="未选择文件", style="Info.TLabel", width=40)
        self.file_label.pack(side=tk.LEFT, padx=(10, 0))

        # --- 步骤 2: 提取模式 ---
        extract_frame = ttk.LabelFrame(main_frame, text=" 步骤 2: 选择提取模式 ", padding="10")
        extract_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Radiobutton(extract_frame, text="提取单个词汇 (过滤通用词)", variable=self.extraction_mode, value="words").pack(anchor=tk.W)
        ttk.Radiobutton(extract_frame, text="提取关键短语 (标题和术语)", variable=self.extraction_mode, value="phrases").pack(anchor=tk.W)
        ttk.Radiobutton(extract_frame, text="筛选版单个词汇 (不过滤)", variable=self.extraction_mode, value="words_no_filter").pack(anchor=tk.W)

        # --- 开始处理 ---
        self.start_button = ttk.Button(main_frame, text="开始生成索引", command=self.start_processing, style="Accent.TButton")
        self.start_button.pack(fill=tk.X, ipady=5)

        # --- 状态栏和进度条 ---
        status_frame = ttk.Frame(self.root, padding="5 2")
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label = ttk.Label(status_frame, text="状态：就绪", anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_bar = ttk.Progressbar(status_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(side=tk.RIGHT)
        
        # --- 样式配置 ---
        style = ttk.Style()
        style.configure("TLabel", foreground="#333")
        style.configure("Info.TLabel", foreground="blue")
        style.configure("Accent.TButton", font=('Helvetica', 10, 'bold'), foreground="white")
        style.map("Accent.TButton", background=[('active', '#005fcc'), ('!disabled', '#007bff')], foreground=[('!disabled', 'white')])

    def select_file(self):
        """打开文件对话框选择Word文档"""
        path = filedialog.askopenfilename(title="选择一个Word文档", filetypes=[("Word 文档", "*.docx")])
        if path:
            self.input_filepath.set(path)
            self.file_label.config(text=f".../{os.path.basename(path)}")

    def start_processing(self):
        """开始处理流程，创建并启动后台线程"""
        if not FILE_PROCESSING_AVAILABLE:
            messagebox.showerror("依赖缺失", "缺少必要的库 (docx2pdf, PyMuPDF)。\n请在命令行运行: pip install docx2pdf PyMuPDF")
            return
            
        if not self.input_filepath.get():
            messagebox.showwarning("操作无效", "请先选择一个Word文件！")
            return
            
        if self.processing_thread and self.processing_thread.is_alive():
            messagebox.showinfo("提示", "正在处理中，请稍候...")
            return

        self.start_button.config(state=tk.DISABLED)
        self.processing_thread = threading.Thread(
            target=self._run_backend_task,
            args=(self.input_filepath.get(), self.extraction_mode.get()),
            daemon=True
        )
        self.processing_thread.start()

    def _run_backend_task(self, docx_path, extraction_mode):
        """后台线程执行的主任务"""
        try:
            term_map = self.backend.extract_from_docx(docx_path, extraction_mode)
            self.root.after(0, self.on_processing_complete, term_map, None)
        except Exception as e:
            self.root.after(0, self.on_processing_complete, None, e)

    def on_processing_complete(self, term_map, error):
        """后台任务完成后的回调函数"""
        self.start_button.config(state=tk.NORMAL)
        self.update_progress(0, 100)

        if error:
            self.update_status(f"错误: {error}", is_error=True)
            messagebox.showerror("处理失败", f"处理过程中发生错误:\n{error}")
            return
        
        self.update_status("状态：处理完成！请选择保存位置。")
        output_path = filedialog.asksaveasfilename(
            title="保存TXT索引文件",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if output_path:
            try:
                self.backend.save_results_as_txt(output_path, term_map)
                messagebox.showinfo("成功", "包含Typst代码的TXT文件已成功保存。\n\n请注意，本程序生成的索引表仅供参考。由于文件格式或其他技术问题，索引表可能存在遗漏或不准确之处。我们正在不断优化程序，但仍建议以具体文件内容为准。本应用及其开发者对索引表的准确性、完整性或由此引发的任何后果不承担责任。")
            except Exception as e:
                self.update_status(f"错误: {e}", is_error=True)
                messagebox.showerror("保存失败", f"无法保存文件:\n{e}")
        else:
            self.update_status("状态：用户取消保存。")

    def update_status(self, text: str, is_error: bool = False):
        """安全地更新状态栏文本"""
        self.status_label.config(text=text, foreground="red" if is_error else "black")

    def update_progress(self, value, maximum):
        """安全地更新进度条"""
        self.progress_bar['maximum'] = maximum
        self.progress_bar['value'] = value
        self.progress_bar.stop() # 确保不确定模式的动画停止

if __name__ == "__main__":
    app_root = tk.Tk()
    app = SimplifiedIndexerApp(app_root)
    app_root.mainloop()
