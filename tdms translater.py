import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from datetime import datetime

def calculate_time_diff():
    # 使用者選擇檔案
    file_path = filedialog.askopenfilename(
        title="選擇 TDMS 檔案",
        filetypes=[("TDMS files", "*.tdms")]
    )
    if not file_path:
        return

    # 使用者輸入機台時間
    input_str = simpledialog.askstring(
        "輸入機台時間",
        "請輸入機台時間（格式：YYYY-MM-DD HH:MM）\n例如：2025-04-22 09:52"
    )
    if not input_str:
        return

    try:
        machine_time = datetime.strptime(input_str, "%Y-%m-%d %H:%M")
    except ValueError:
        messagebox.showerror("格式錯誤", "時間格式錯誤，請使用 YYYY-MM-DD HH:MM")
        return

    real_time = datetime.now()
    delta = real_time - machine_time
    total_seconds = int(delta.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes = remainder // 60

    info = (
        f"📂 檔案名稱：{file_path.split('/')[-1]}\n"
        f"🕒 手動輸入機台時間：{machine_time.strftime('%Y-%m-%d %H:%M')}\n"
        f"⏰ 系統現在時間：{real_time.strftime('%Y-%m-%d %H:%M')}\n"
        f"🧮 時間差（真實時間 - 機台時間）：約 {hours} 小時 {minutes} 分鐘"
    )

    messagebox.showinfo("時間資訊結果", info)

# 明確建立主視窗（並設為頂層）
root = tk.Tk()
root.title("TDMS 時間差工具")
root.geometry("400x200")
root.attributes("-topmost", True)  # 視窗置頂

# 按鈕
btn = tk.Button(root, text="選擇檔案並輸入機台時間", font=("Arial", 14), command=calculate_time_diff)
btn.pack(pady=60)

root.mainloop()
