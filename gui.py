import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xlwings as xw

selected_images = []
selected_excel = ""
cancel_flag = False


def parse_cell(value):
    """把 A1、b2、AA10 解析成 (row, col)"""
    value = value.strip().upper()
    if not value:
        raise ValueError("起始单元格不能为空")

    letters = ""
    digits = ""
    i = 0

    while i < len(value) and value[i].isalpha():
        letters += value[i]
        i += 1

    while i < len(value) and value[i].isdigit():
        digits += value[i]
        i += 1

    if i != len(value) or not letters or not digits:
        raise ValueError("起始单元格格式类似 A1")

    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)

    row = int(digits)

    if row < 1 or col < 1:
        raise ValueError("行号和列号必须大于 0")

    return row, col


def log(msg):
    log_text.insert(tk.END, msg + "\n")
    log_text.see(tk.END)
    root.update()


def on_select_images():
    global selected_images

    paths = filedialog.askopenfilenames(
        title="请选择图片（可按住 Ctrl 多选，或 Ctrl+A 全选）",
        filetypes=[
            ("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif"),
            ("所有文件", "*.*"),
        ],
    )

    if not paths:
        log("取消选择图片")
        return

    selected_images = [p.replace("/", "\\") for p in paths]
    selected_images.sort()

    label_images_info.config(text=f"已选图片：{len(selected_images)} 张")
    log(f"已选择图片：{len(selected_images)} 张")
    for p in selected_images:
        log(f"  {p}")


def on_select_excel():
    global selected_excel

    path = filedialog.askopenfilename(
        title="请选择 Excel 文件",
        filetypes=[
            ("Excel 文件", "*.xlsx *.xls *.xlsm"),
            ("新版 Excel", "*.xlsx"),
            ("老版 Excel", "*.xls"),
            ("带宏 Excel", "*.xlsm"),
            ("所有文件", "*.*"),
        ],
    )

    if not path:
        log("取消选择 Excel")
        return

    selected_excel = path.replace("/", "\\")
    label_excel_info.config(text=os.path.basename(selected_excel))
    log(f"已选择 Excel：{selected_excel}")


def on_cancel():
    global cancel_flag
    cancel_flag = True
    log("已请求取消，正在停止...")


def insert_images(ws, files, mode, start_row, start_col, gap):
    """把图片依次插入到工作表，并填满单元格"""
    global cancel_flag

    step = gap + 1
    total = len(files)

    progress_bar["maximum"] = total
    progress_bar["value"] = 0

    for i, full_path in enumerate(files):
        if cancel_flag:
            log("已取消插入。")
            return False

        if mode == "行":
            cell = ws.range((start_row + i * step, start_col))
        else:
            cell = ws.range((start_row, start_col + i * step))

        fixed_path = full_path.replace("/", "\\")

        ws.pictures.add(
            fixed_path,
            left=cell.left,
            top=cell.top,
            width=cell.width,
            height=cell.height,
        )

        progress_bar["value"] = i + 1
        log(f"已插入 {i + 1}/{total}：{fixed_path}")

    return True


def on_start():
    global cancel_flag
    cancel_flag = False

    if not selected_images:
        messagebox.showwarning("提示", "请先选择图片")
        return

    if not selected_excel:
        messagebox.showwarning("提示", "请先选择 Excel 文件")
        return

    mode = mode_var.get()

    try:
        start_row, start_col = parse_cell(cell_entry.get())
    except ValueError as e:
        messagebox.showerror("输入错误", str(e))
        return

    try:
        gap = int(gap_entry.get().strip())
        if gap < 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("输入错误", "图片间距必须是大于或等于 0 的整数")
        return

    btn_start.config(state="disabled")
    btn_cancel.config(state="normal")
    progress_bar["value"] = 0

    log("开始插入...")
    log(f"方向={mode}，起始单元格={cell_entry.get().strip().upper()}，间距={gap}")

    app = None
    try:
        app = xw.App(visible=False)
        wb = app.books.open(selected_excel)
        ws = wb.sheets[0]

        finished = insert_images(ws, selected_images, mode, start_row, start_col, gap)

        if finished:
            wb.save()
            log("完成，所有图片已插入。")
            messagebox.showinfo("完成", "所有图片已插入。")
        else:
            log("未保存，已取消。")

        wb.close()

    except Exception as e:
        log(f"出错了：{e}")
        messagebox.showerror("出错了", str(e))

    finally:
        if app is not None:
            app.quit()
        btn_start.config(state="normal")
        btn_cancel.config(state="disabled")


root = tk.Tk()
root.title("Excel 图片批量插入工具")
root.geometry("700x640")

title_label = tk.Label(
    root,
    text="Excel 图片批量插入工具",
    font=("微软雅黑", 16, "bold")
)
title_label.pack(pady=10)

frame_images = tk.Frame(root)
frame_images.pack(fill="x", padx=20, pady=5)

btn_select_images = tk.Button(
    frame_images,
    text="选择图片",
    width=12,
    command=on_select_images
)
btn_select_images.pack(side="left")

label_images_info = tk.Label(
    frame_images,
    text="已选图片：0 张",
    anchor="w"
)
label_images_info.pack(side="left", padx=10)

frame_excel = tk.Frame(root)
frame_excel.pack(fill="x", padx=20, pady=5)

btn_select_excel = tk.Button(
    frame_excel,
    text="选择 Excel",
    width=12,
    command=on_select_excel
)
btn_select_excel.pack(side="left")

label_excel_info = tk.Label(
    frame_excel,
    text="未选择 Excel",
    anchor="w"
)
label_excel_info.pack(side="left", padx=10)

frame_params = tk.Frame(root)
frame_params.pack(fill="x", padx=20, pady=10)

tk.Label(frame_params, text="排列方向：").grid(row=0, column=0, sticky="w")

mode_var = tk.StringVar(value="行")
tk.Radiobutton(frame_params, text="行", variable=mode_var, value="行").grid(row=0, column=1, sticky="w")
tk.Radiobutton(frame_params, text="列", variable=mode_var, value="列").grid(row=0, column=2, sticky="w")

tk.Label(frame_params, text="起始单元格：").grid(row=1, column=0, sticky="w", pady=5)
cell_entry = tk.Entry(frame_params, width=15)
cell_entry.insert(0, "A1")
cell_entry.grid(row=1, column=1, sticky="w", pady=5)

tk.Label(frame_params, text="图片间距：").grid(row=2, column=0, sticky="w", pady=5)
gap_entry = tk.Entry(frame_params, width=15)
gap_entry.insert(0, "0")
gap_entry.grid(row=2, column=1, sticky="w", pady=5)

frame_buttons = tk.Frame(root)
frame_buttons.pack(pady=10)

btn_start = tk.Button(
    frame_buttons,
    text="开始插入",
    width=15,
    height=2,
    command=on_start
)
btn_start.pack(side="left", padx=10)

btn_cancel = tk.Button(
    frame_buttons,
    text="取消",
    width=15,
    height=2,
    command=on_cancel,
    state="disabled"
)
btn_cancel.pack(side="left", padx=10)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=600, mode="determinate")
progress_bar.pack(padx=20, pady=10)

log_text = tk.Text(root, height=14)
log_text.pack(fill="both", expand=True, padx=20, pady=10)

root.mainloop()