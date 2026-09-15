import os
import sys
import tkinter as tk
from tkinter import filedialog
import xlwings as xw

EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".gif")


class UserExit(Exception):
    """用户主动退出"""
    pass


def safe_input(prompt):
    """所有输入都走这里，支持 q / exit / 退出 / Ctrl+C"""
    try:
        value = input(prompt).strip()
    except KeyboardInterrupt:
        print()
        raise UserExit()

    if value.lower() in ("q", "exit", "退出"):
        raise UserExit()

    return value


def ask_image_files():
    """弹出文件选择窗口，让用户多选图片"""
    root = tk.Tk()
    root.withdraw()

    paths = filedialog.askopenfilenames(
        title="请选择图片（可按住 Ctrl 多选，或 Ctrl+A 全选）",
        filetypes=[
            ("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif"),
            ("所有文件", "*.*"),
        ],
    )

    root.destroy()

    if not paths:
        raise UserExit()

    files = [p.replace("/", "\\") for p in paths]
    files.sort()
    return files


def ask_excel_file():
    """弹出文件选择窗口，支持 xlsx / xls / xlsm"""
    root = tk.Tk()
    root.withdraw()

    path = filedialog.askopenfilename(
        title="请选择 Excel 文件（取消则退出）",
        filetypes=[
            ("Excel 文件", "*.xlsx *.xls *.xlsm"),
            ("新版 Excel", "*.xlsx"),
            ("老版 Excel", "*.xls"),
            ("带宏 Excel", "*.xlsm"),
            ("所有文件", "*.*"),
        ],
    )

    root.destroy()

    if not path:
        raise UserExit()

    return path


def ask_mode():
    """循环询问，直到输入 行 或 列"""
    while True:
        mode = safe_input("请输入 行 或 列（输入 q 退出）：").strip()
        if mode in ("行", "列"):
            return mode
        print("输入错误，只能输入 行 或 列，请重新输入。")


def ask_cell():
    """循环询问起始单元格，比如 B2，返回 (row, col)"""
    while True:
        value = safe_input("请输入起始单元格（比如 B2，输入 q 退出）：").strip().upper()

        if not value:
            print("输入不能为空，请重新输入。")
            continue

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
            print("输入错误，格式类似 B2，请重新输入。")
            continue

        col = 0
        for ch in letters:
            col = col * 26 + (ord(ch) - ord("A") + 1)

        row = int(digits)

        if row < 1 or col < 1:
            print("行号和列号必须大于 0，请重新输入。")
            continue

        return row, col


def ask_gap():
    """循环询问，直到输入一个大于或等于 0 的整数"""
    while True:
        value = safe_input("请输入图片间距（0 表示紧挨着，输入 q 退出）：").strip()
        try:
            num = int(value)
        except ValueError:
            print("输入错误，必须输入数字，请重新输入。")
            continue
        if num < 0:
            print("输入错误，必须大于或等于 0，请重新输入。")
            continue
        return num


def insert_images(ws, files, mode, start_row, start_col, gap):
    """把图片依次插入到工作表，并填满单元格"""
    step = gap + 1
    for i, full_path in enumerate(files):
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
        print(f"已插入：{fixed_path}")


def main():
    try:
        files = ask_image_files()
        print("已选择图片：", len(files))

        excel_path = ask_excel_file()
        print("已选择 Excel：", excel_path)

        mode = ask_mode()
        start_row, start_col = ask_cell()
        gap = ask_gap()

        app = xw.App(visible=False)
        try:
            wb = app.books.open(excel_path)
            ws = wb.sheets[0]
            insert_images(ws, files, mode, start_row, start_col, gap)
            wb.save()
            wb.close()
        finally:
            app.quit()

        print("完成，所有图片已插入。")

    except UserExit:
        print("已退出。")
    except FileNotFoundError as e:
        print("文件错误：", e)
    except ValueError as e:
        print("输入错误：", e)
    except Exception as e:
        print("出错了：", e)
        sys.exit(1)


if __name__ == "__main__":
    main()