# -*- coding: utf-8 -*-
"""PLC模块排序生成器 - GUI界面"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_module import generate_module_excel, generate_i_modules, generate_q_modules, generate_iw_modules, generate_qw_modules


class ModuleGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PLC模块排序生成器")
        self.root.geometry("500x450")
        
        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="PLC模块排序生成器", font=("Arial", 14, "bold")).pack(pady=5)
        
        # I模块设置
        i_frame = ttk.LabelFrame(main_frame, text="I模块设置", padding="10")
        i_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(i_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.i_count = tk.IntVar(value=4)
        ttk.Spinbox(i_frame, from_=1, to=32, textvariable=self.i_count, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(i_frame, text="起始字节:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.i_start = tk.IntVar(value=2)
        ttk.Spinbox(i_frame, from_=0, to=1000, textvariable=self.i_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(i_frame, text=f"(共 {self.i_count.get()*16} 点)", foreground="gray").grid(row=0, column=4, sticky=tk.W)
        self.i_label = ttk.Label(i_frame, text="", foreground="blue")
        self.i_label.grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=(5,0))
        
        # Q模块设置
        q_frame = ttk.LabelFrame(main_frame, text="Q模块设置", padding="10")
        q_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(q_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.q_count = tk.IntVar(value=4)
        ttk.Spinbox(q_frame, from_=1, to=32, textvariable=self.q_count, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(q_frame, text="起始字节:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.q_start = tk.IntVar(value=2)
        ttk.Spinbox(q_frame, from_=0, to=1000, textvariable=self.q_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(q_frame, text=f"(共 {self.q_count.get()*16} 点)", foreground="gray").grid(row=0, column=4, sticky=tk.W)
        self.q_label = ttk.Label(q_frame, text="", foreground="blue")
        self.q_label.grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=(5,0))
        
        # IW模块设置
        iw_frame = ttk.LabelFrame(main_frame, text="IW模块设置", padding="10")
        iw_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(iw_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.iw_count = tk.IntVar(value=2)
        ttk.Spinbox(iw_frame, from_=1, to=32, textvariable=self.iw_count, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(iw_frame, text="起始偏移:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.iw_start = tk.IntVar(value=10)
        ttk.Spinbox(iw_frame, from_=0, to=1000, textvariable=self.iw_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(iw_frame, text=f"(共 {self.iw_count.get()*8} 点)", foreground="gray").grid(row=0, column=4, sticky=tk.W)
        self.iw_label = ttk.Label(iw_frame, text="", foreground="blue")
        self.iw_label.grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=(5,0))
        
        # QW模块设置
        qw_frame = ttk.LabelFrame(main_frame, text="QW模块设置", padding="10")
        qw_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(qw_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.qw_count = tk.IntVar(value=2)
        ttk.Spinbox(qw_frame, from_=1, to=32, textvariable=self.qw_count, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(qw_frame, text="起始偏移:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.qw_start = tk.IntVar(value=16)
        ttk.Spinbox(qw_frame, from_=0, to=1000, textvariable=self.qw_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(qw_frame, text=f"(共 {self.qw_count.get()*4} 点)", foreground="gray").grid(row=0, column=4, sticky=tk.W)
        self.qw_label = ttk.Label(qw_frame, text="", foreground="blue")
        self.qw_label.grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=(5,0))
        
        # 输出文件
        out_frame = ttk.Frame(main_frame)
        out_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(out_frame, text="输出文件:").pack(side=tk.LEFT)
        self.output_path = tk.StringVar(value="PLC模块排序表.xlsx")
        ttk.Entry(out_frame, textvariable=self.output_path, width=30).pack(side=tk.LEFT, padx=5)
        ttk.Button(out_frame, text="浏览", command=self.browse_output).pack(side=tk.LEFT)
        
        # 按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="预览", command=self.preview).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="生成", command=self.generate, width=12).pack(side=tk.LEFT, padx=5)
        
        # 预览框
        preview_frame = ttk.LabelFrame(main_frame, text="预览", padding="5")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.preview_text = tk.Text(preview_frame, height=6, font=("Consolas", 9))
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        
        self.preview()
    
    def browse_output(self):
        filename = filedialog.asksaveasfilename(title="保存位置", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if filename:
            self.output_path.set(filename)
    
    def preview(self):
        i_pts = generate_i_modules(self.i_count.get(), self.i_start.get())
        q_pts = generate_q_modules(self.q_count.get(), self.q_start.get())
        iw_pts = generate_iw_modules(self.iw_count.get(), self.iw_start.get())
        qw_pts = generate_qw_modules(self.qw_count.get(), self.qw_start.get())
        
        preview = f"""【I模块】{len(i_pts)}点: {', '.join(i_pts[:8])}...
【Q模块】{len(q_pts)}点: {', '.join(q_pts[:8])}...
【IW模块】{len(iw_pts)}点: {', '.join(iw_pts)}
【QW模块】{len(qw_pts)}点: {', '.join(qw_pts)}"""
        
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(1.0, preview)
    
    def generate(self):
        try:
            output = self.output_path.get()
            if not output.endswith('.xlsx'):
                output += '.xlsx'
            
            generate_module_excel(
                self.i_count.get(), self.i_start.get(),
                self.q_count.get(), self.q_start.get(),
                self.iw_count.get(), self.iw_start.get(),
                self.qw_count.get(), self.qw_start.get(),
                output
            )
            messagebox.showinfo("成功", f"已生成: {output}")
        except Exception as e:
            messagebox.showerror("错误", str(e))


if __name__ == '__main__':
    root = tk.Tk()
    app = ModuleGeneratorApp(root)
    root.mainloop()
