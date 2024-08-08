import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd  # type: ignore


# ナップサックアルゴリズムを実行する関数
def knapsack_wood_cutting(lengths, counts, max_length):
    dp = [0] * (max_length + 1)
    item_used = [None] * (max_length + 1)

    for i, (length, count) in enumerate(zip(lengths, counts)):
        for _ in range(count):
            for j in range(max_length, length - 1, -1):
                if dp[j - length] + length > dp[j]:
                    dp[j] = dp[j - length] + length
                    item_used[j] = i

    remaining_length = max_length
    used_counts = [0] * len(lengths)
    cut_combination = []

    while remaining_length > 0 and item_used[remaining_length] is not None:
        i = item_used[remaining_length]
        used_counts[i] += 1
        cut_combination.append(lengths[i])
        remaining_length -= lengths[i]

    return (
        used_counts,
        max_length
        - sum(lengths[i] * used_counts[i] for i in range(len(lengths))),
        cut_combination,
    )


# グローバル変数でファイルパスを保存
file_path = None


# ファイル選択と処理を行う関数
def open_file():
    global file_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        # ファイル名を表示
        file_name_label.config(text=f"選択されたファイル: {file_path.split('/')[-1]}")
    else:
        file_name_label.config(text="ファイルが選択されていません")


# ナップサックアルゴリズムを実行して、必要なmax_lengthの本数を計算する関数
def run_algorithm():
    global file_path
    if not file_path:
        messagebox.showwarning("警告", "まずファイルを選択してください")
        return

    try:
        # Excelファイルを読み込む
        data = pd.read_excel(file_path)
        lengths = data["長さ(mm)"].tolist()
        counts = data["本数"].tolist()

        # ユーザーが入力した長材の長さを取得
        max_length_str = max_length_entry.get()
        if not max_length_str.isdigit():
            messagebox.showerror("エラー", "有効な長材の長さを入力してください（整数）")
            return

        max_length = int(max_length_str)

        # 木材のカットを繰り返して必要な本数を計算
        total_used_counts = [0] * len(lengths)
        total_material_count = 0
        all_cut_combinations = []

        result_label.config(state=tk.NORMAL)
        result_label.delete(1.0, tk.END)  # 結果をクリア

        while sum(counts) > 0:
            used_counts, min_waste, cut_combination = knapsack_wood_cutting(
                lengths, counts, max_length
            )

            # 使用した木材をカウントし、全体の使用本数を更新
            if sum(used_counts) == 0:
                break

            for i in range(len(lengths)):
                total_used_counts[i] += used_counts[i]
                counts[i] -= used_counts[i]

            total_material_count += 1
            all_cut_combinations.append((cut_combination, min_waste))

            # 結果を順次表示
            result_text = f"長材 {total_material_count}:\n"
            result_text += (
                f"  カットされた長さ: {', '.join(map(str, cut_combination))}\n"
            )
            result_text += f"  余りの長さ: {min_waste} mm\n\n"
            result_label.insert(tk.END, result_text)
            result_label.see(tk.END)
            result_label.update_idletasks()

        # 全体結果を表示
        summary_text = f"\n必要な長材の本数: {total_material_count} 本\n"
        result_label.insert(tk.END, summary_text)
        result_label.see(tk.END)
        result_label.update_idletasks()

        result_label.config(state=tk.DISABLED)

    except Exception as e:
        messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました:\n{e}")


# GUIのセットアップ
root = tk.Tk()
root.title("木材カット最適化")
root.geometry("850x600")  # 画面サイズを設定

# フォントと色の設定
FONT = ("Helvetica", 12)
HEADING_FONT = ("Helvetica", 14, "bold")
BG_COLOR = "#f0f0f0"
BUTTON_COLOR = "#4CAF50"
BUTTON_TEXT_COLOR = "#ffffff"

root.configure(bg=BG_COLOR)

frame = tk.Frame(root, padx=20, pady=20, bg=BG_COLOR)
frame.pack(padx=20, pady=20, expand=True, fill="both")

heading_label = tk.Label(
    frame, text="木材カット最適化ツール", font=HEADING_FONT, bg=BG_COLOR
)
heading_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))

# 長材の長さの入力をグループ化して中央揃え
max_length_frame = tk.Frame(frame, bg=BG_COLOR)
max_length_frame.grid(row=1, column=0, columnspan=3, pady=10)

max_length_label = tk.Label(
    max_length_frame, text="長材の長さ (mm):", font=FONT, bg=BG_COLOR
)
max_length_label.pack(side="left", padx=(0, 5))

max_length_entry = tk.Entry(
    max_length_frame, font=FONT, justify="center", width=10
)
max_length_entry.pack(side="left")

# ファイル選択ボタンとファイル名表示をグループ化して中央揃え
file_frame = tk.Frame(frame, bg=BG_COLOR)
file_frame.grid(row=2, column=0, columnspan=3, pady=10)

open_button = tk.Button(
    file_frame,
    text="ファイルを開く",
    font=FONT,
    command=open_file,
)
open_button.pack(side="left", padx=(0, 10))

file_name_label = tk.Label(
    file_frame, text="ファイルが選択されていません", font=FONT, bg=BG_COLOR, anchor="w"
)
file_name_label.pack(side="left")

# 実行ボタンを中央揃え
run_button = tk.Button(
    frame,
    text="実行",
    font=FONT,
    command=run_algorithm,
)
run_button.grid(row=3, column=0, columnspan=3, pady=10)

# スクロール可能なテキストウィジェットを作成
result_frame = tk.Frame(frame, bg=BG_COLOR)
result_frame.grid(row=4, column=0, columnspan=3, pady=(20, 0), sticky="nsew")

scrollbar = tk.Scrollbar(result_frame, orient="vertical")
scrollbar.pack(side="right", fill="y")

result_label = tk.Text(
    result_frame,
    wrap="word",
    font=FONT,
    bg=BG_COLOR,
    yscrollcommand=scrollbar.set,
    state=tk.NORMAL,
)
result_label.pack(pady=10, padx=10, expand=True, fill="both")
scrollbar.config(command=result_label.yview)

frame.columnconfigure(0, weight=1)  # 左側の列の拡大を設定
frame.columnconfigure(1, weight=1)  # 中央の列の拡大を設定
frame.columnconfigure(2, weight=1)  # 右側の列の拡大を設定
frame.rowconfigure(4, weight=1)

root.mainloop()
