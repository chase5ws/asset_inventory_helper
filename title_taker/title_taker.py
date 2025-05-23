import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import datetime
import os

def process_file():
    file_path = file_entry.get()
    sheet_name = table_entry.get()

    if not os.path.isfile(file_path):
        messagebox.showerror("錯誤", "檔案不存在，請重新選擇！")
        return

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取 Excel 發生錯誤：\n{e}")
        return

    # 取所有欄位的唯一值，組成舉證
    unique_dict = {col: df[col].dropna().unique() for col in df.columns}
    max_len = max(len(vals) for vals in unique_dict.values())
    for col in unique_dict:
        unique_dict[col] = list(unique_dict[col]) + [''] * (max_len - len(unique_dict[col]))

    result_df = pd.DataFrame(unique_dict)

    # 讓使用者選擇儲存檔案名稱與路徑
    default_filename = datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx'
    out_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile=default_filename,
        title="儲存唯一值 Excel"
    )
    if not out_path:
        return  # 使用者取消儲存

    try:
        result_df.to_excel(out_path, index=False)
    except Exception as e:
        messagebox.showerror("錯誤", f"儲存 Excel 發生錯誤：\n{e}")
        return

    # 詢問是否繼續
    answer = messagebox.askyesno("完成", f"已儲存成檔案：\n{out_path}\n\n是否要繼續？")
    if answer:
        # 清空欄位
        file_entry.delete(0, tk.END)
        table_entry.delete(0, tk.END)
        table_entry.insert(0, "Sheet1")
    else:
        root.destroy()

def browse_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def close_app():
    root.destroy()

# UI 設計
root = tk.Tk()
root.title("Excel title taker")

tk.Label(root, text="Excel 檔案：").grid(row=0, column=0, padx=5, pady=5, sticky='e')
file_entry = tk.Entry(root, width=40)
file_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="選擇檔案", command=browse_file).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Sheet 名稱：").grid(row=1, column=0, padx=5, pady=5, sticky='e')
table_entry = tk.Entry(root, width=40)
table_entry.grid(row=1, column=1, padx=5, pady=5)
table_entry.insert(0, "Sheet1")  # 預設值

tk.Button(root, text="儲存", command=process_file, bg="#4CAF50", fg="white").grid(row=2, column=0, columnspan=2, pady=15, sticky='ew')
tk.Button(root, text="關閉程式", command=close_app, bg="#F44336", fg="white").grid(row=2, column=2, pady=15, sticky='ew')

root.mainloop()
