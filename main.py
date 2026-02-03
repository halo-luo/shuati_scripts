import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime


class QuizViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("题库查看器 - 带错题本")
        self.root.geometry("1100x850")
        self.root.resizable(True, True)

        self.df = None  # 原题库
        self.current_index = 0
        self.has_answered = False
        self.wrong_questions = []  # 错题本列表：[{index, question, user_answer, correct, basis, time}, ...]

        # 顶部：加载 + 查看错题本 + 导出
        top_frame = tk.Frame(root)
        top_frame.pack(pady=10, fill="x", padx=20)

        self.load_btn = tk.Button(top_frame, text="选择Excel题库", font=("Microsoft YaHei", 12),
                                  command=self.load_excel)
        self.load_btn.pack(side="left", padx=10)

        self.file_label = tk.Label(top_frame, text="尚未加载题库", font=("Microsoft YaHei", 11), fg="gray")
        self.file_label.pack(side="left", padx=20)

        tk.Button(top_frame, text="查看错题本", font=("Microsoft YaHei", 12),
                  command=self.show_wrong_book).pack(side="left", padx=10)

        tk.Button(top_frame, text="导出错题本到Excel", font=("Microsoft YaHei", 12),
                  command=self.export_wrong_book).pack(side="left", padx=10)

        # 题干
        self.question_label = tk.Label(root, text="", wraplength=1050, justify="left",
                                       font=("Microsoft YaHei", 14, "bold"))
        self.question_label.pack(pady=20, padx=30, anchor="w")

        # 选项区域
        self.options_frame = tk.Frame(root)
        self.options_frame.pack(pady=10, padx=50, fill="x")

        self.option_vars = []
        self.option_buttons = []

        # 确认按钮
        self.confirm_btn = tk.Button(root, text="确认答案", font=("Microsoft YaHei", 12), width=12,
                                     command=self.show_answer, state="disabled")
        self.confirm_btn.pack(pady=10)

        # 答案区
        self.answer_frame = tk.Frame(root)
        self.answer_frame.pack(pady=20, padx=50, fill="x")

        self.result_label = tk.Label(self.answer_frame, text="", font=("Microsoft YaHei", 13, "bold"))
        self.result_label.pack(anchor="w")

        self.correct_answer_label = tk.Label(self.answer_frame, text="", font=("Microsoft YaHei", 12), fg="green")
        self.correct_answer_label.pack(anchor="w")

        self.basis_label = tk.Label(self.answer_frame, text="", wraplength=1000, justify="left",
                                    font=("Microsoft YaHei", 11))
        self.basis_label.pack(anchor="w", pady=(10, 0))

        # 导航
        nav_frame = tk.Frame(root)
        nav_frame.pack(pady=30)

        self.prev_btn = tk.Button(nav_frame, text="上一题", font=("Microsoft YaHei", 12), width=12,
                                  command=self.prev_question, state="disabled")
        self.prev_btn.pack(side="left", padx=50)

        self.next_btn = tk.Button(nav_frame, text="下一题", font=("Microsoft YaHei", 12), width=12,
                                  command=self.next_question, state="disabled")
        self.next_btn.pack(side="left", padx=50)

        tk.Label(nav_frame, text="跳到第", font=("Microsoft YaHei", 11)).pack(side="left", padx=(80, 5))

        self.jump_entry = tk.Entry(nav_frame, width=6, font=("Microsoft YaHei", 11), justify="center")
        self.jump_entry.pack(side="left")

        tk.Button(nav_frame, text="跳转", font=("Microsoft YaHei", 11),
                  command=self.jump_to).pack(side="left", padx=5)

        self.status_label = tk.Label(root, text="", font=("Microsoft YaHei", 10), fg="gray")
        self.status_label.pack(pady=10)

        self.wrong_count_label = tk.Label(root, text="错题本：0 道", font=("Microsoft YaHei", 10), fg="orange")
        self.wrong_count_label.pack(pady=5)

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            title="选择题库Excel文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path)
            self.df.columns = self.df.columns.str.strip()
            str_cols = self.df.select_dtypes(include=['object']).columns
            self.df[str_cols] = self.df[str_cols].apply(lambda x: x.str.strip())

            self.file_label.config(text=f"已加载：{file_path.split('/')[-1]}", fg="green")
            self.current_index = 0
            self.wrong_questions = []  # 清空错题本
            self.update_wrong_count()
            self.show_question(0)

            messagebox.showinfo("成功", f"共加载 {len(self.df)} 道题目")

        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取文件：\n{str(e)}")

    def show_question(self, index):
        if self.df is None or len(self.df) == 0:
            return

        if index < 0 or index >= len(self.df):
            return

        self.has_answered = False
        row = self.df.iloc[index]

        for widget in self.options_frame.winfo_children():
            widget.destroy()
        self.option_vars.clear()
        self.option_buttons.clear()

        self.result_label.config(text="")
        self.correct_answer_label.config(text="")
        self.basis_label.config(text="")

        q_text = f"第 {index + 1} 题   {row.get('题型', '未知题型')}\n\n{row.get('题干', '无题干')}"
        self.question_label.config(text=q_text)

        options_str = row.get('选项', '')
        if pd.isna(options_str) or not options_str.strip():
            tk.Label(self.options_frame, text="（本题无选项）", font=("Microsoft YaHei", 11)).pack(anchor="w")
        else:
            options_list = [opt.strip() for opt in str(options_str).split('|') if opt.strip()]
            for i, opt in enumerate(options_list):
                var = tk.BooleanVar(value=False)
                self.option_vars.append(var)
                cmd = lambda v=var: self.on_option_change(v)
                btn = tk.Checkbutton(self.options_frame, text=opt, variable=var,
                                     font=("Microsoft YaHei", 12), anchor="w", justify="left",
                                     command=cmd)
                btn.pack(fill="x", pady=5)
                self.option_buttons.append(btn)

        self.status_label.config(text=f"第 {index + 1} / {len(self.df)} 题    分值：{row.get('试题分数', '1')}")
        self.confirm_btn.config(state="disabled")
        self.prev_btn.config(state="normal" if index > 0 else "disabled")
        self.next_btn.config(state="normal" if index < len(self.df) - 1 else "disabled")

    def on_option_change(self, var):
        any_selected = any(v.get() for v in self.option_vars)
        self.confirm_btn.config(state="normal" if any_selected else "disabled")

    def show_answer(self):
        if self.df is None or self.has_answered:
            return

        row = self.df.iloc[self.current_index]
        correct = str(row.get('答案', '')).strip()

        user_choices = []
        for i, var in enumerate(self.option_vars):
            if var.get():
                opt_text = self.option_buttons[i].cget("text")
                letter = opt_text.split('-')[0].strip() if '-' in opt_text else ""
                user_choices.append(letter)

        user_answer = "".join(user_choices) if user_choices else "未选"

        is_correct = user_answer.strip() == correct.strip()

        result_text = "回答正确！" if is_correct else "回答错误！"
        result_color = "green" if is_correct else "red"
        self.result_label.config(text=result_text, fg=result_color)

        self.correct_answer_label.config(text=f"正确答案：{correct}")
        self.basis_label.config(text=f"题目依据：{row.get('题目依据', '无')}")

        # 如果错题，记录到错题本
        if not is_correct:
            wrong_entry = {
                "序号": row.get("序号", self.current_index + 1),
                "题干": row.get("题干", "无"),
                "选项": row.get("选项", "无"),
                "用户答案": user_answer,
                "答案": correct,
                "题目依据": row.get("题目依据", "无"),
                "做题时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "原题库索引": self.current_index
            }
            self.wrong_questions.append(wrong_entry)
            self.update_wrong_count()

        # 高亮用户选择
        for i, var in enumerate(self.option_vars):
            if var.get():
                btn = self.option_buttons[i]
                opt_text = btn.cget("text")
                letter = opt_text.split('-')[0].strip() if '-' in opt_text else ""
                if letter == correct:
                    btn.config(fg="green", selectcolor="lightgreen")
                else:
                    btn.config(fg="red", selectcolor="pink")

        self.has_answered = True
        self.confirm_btn.config(state="disabled")

    def update_wrong_count(self):
        count = len(self.wrong_questions)
        self.wrong_count_label.config(text=f"错题本：{count} 道")

    def show_wrong_book(self):
        if not self.wrong_questions:
            messagebox.showinfo("错题本", "当前没有错题记录")
            return

        wrong_window = tk.Toplevel(self.root)
        wrong_window.title("错题本")
        wrong_window.geometry("1100x700")

        text_area = tk.Text(wrong_window, wrap="word", font=("Microsoft YaHei", 11))
        text_area.pack(fill="both", expand=True, padx=10, pady=10)

        scrollbar = tk.Scrollbar(wrong_window, command=text_area.yview)
        scrollbar.pack(side="right", fill="y")
        text_area.config(yscrollcommand=scrollbar.set)

        for i, w in enumerate(self.wrong_questions, 1):
            text_area.insert(tk.END, f"错题 {i} (原序号：{w['序号']})\n")
            text_area.insert(tk.END, f"题干：{w['题干']}\n")
            text_area.insert(tk.END, f"选项：{w['选项']}\n")
            text_area.insert(tk.END, f"你的答案：{w['用户答案']}    正确答案：{w['答案']}\n")
            text_area.insert(tk.END, f"题目依据：{w['题目依据']}\n")
            text_area.insert(tk.END, f"做错时间：{w['做题时间']}\n")
            text_area.insert(tk.END, "-" * 80 + "\n\n")

        text_area.config(state="disabled")

    def export_wrong_book(self):
        if not self.wrong_questions:
            messagebox.showinfo("提示", "当前没有错题可导出")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            title="保存错题本"
        )
        if not file_path:
            return

        try:
            df_wrong = pd.DataFrame(self.wrong_questions)
            df_wrong.to_excel(file_path, index=False)
            messagebox.showinfo("成功", f"错题本已保存到：\n{file_path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def next_question(self):
        if self.current_index < len(self.df) - 1:
            self.current_index += 1
            self.show_question(self.current_index)

    def prev_question(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.show_question(self.current_index)

    def jump_to(self):
        try:
            num = int(self.jump_entry.get().strip())
            if 1 <= num <= len(self.df):
                self.current_index = num - 1
                self.show_question(self.current_index)
            else:
                messagebox.showwarning("提示", f"请输入 1 ~ {len(self.df)} 之间的数字")
        except ValueError:
            messagebox.showwarning("提示", "请输入有效数字")


# 启动程序
if __name__ == "__main__":
    root = tk.Tk()
    app = QuizViewer(root)
    root.mainloop()
