import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess

class MarkItDownGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MarkItDown GUI")
        self.root.geometry("600x400")  # 初始窗口大小
        self.root.resizable(False, False)  # 固定窗口大小

        # 高分辨率屏幕优化
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception as e:
            print("高分辨率适配失败：", e)

        # 输入文件路径
        self.input_file = tk.StringVar()
        tk.Label(root, text="选择输入文件:", font=("Arial", 12)).pack(pady=10)
        tk.Entry(root, textvariable=self.input_file, width=50, font=("Arial", 10)).pack(pady=5)
        tk.Button(root, text="浏览...", command=self.select_input_file, font=("Arial", 10)).pack()

        # 输出文件路径
        self.output_file = tk.StringVar()
        tk.Label(root, text="选择输出文件:", font=("Arial", 12)).pack(pady=10)
        tk.Entry(root, textvariable=self.output_file, width=50, font=("Arial", 10)).pack(pady=5)
        tk.Button(root, text="浏览...", command=self.select_output_file, font=("Arial", 10)).pack()

        # 转换按钮
        tk.Button(root, text="开始转换", command=self.start_conversion, font=("Arial", 12), bg="green", fg="white").pack(pady=20)

    def select_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择输入文件",
            filetypes=(("Word 文件", "*.docx"), ("所有文件", "*.*"))
        )
        if file_path:
            self.input_file.set(file_path)

    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".md",
            filetypes=(("Markdown 文件", "*.md"), ("所有文件", "*.*"))
        )
        if file_path:
            self.output_file.set(file_path)

    def start_conversion(self):
        input_path = self.input_file.get()
        output_path = self.output_file.get()

        # 验证路径是否有效
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("错误", "请输入有效的输入文件路径！")
            return

        if not output_path:
            messagebox.showerror("错误", "请输入有效的输出文件路径！")
            return

        # 调用 MarkItDown 命令
        try:
            command = f"markitdown \"{input_path}\" -o \"{output_path}\""
            result = subprocess.run(command, shell=True, capture_output=True, text=True)

            if result.returncode == 0:
                messagebox.showinfo("成功", "文件转换成功！")
            else:
                messagebox.showerror("错误", f"转换失败：{result.stderr}")
        except Exception as e:
            messagebox.showerror("错误", f"出现意外错误：{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MarkItDownGUI(root)
    root.mainloop()
