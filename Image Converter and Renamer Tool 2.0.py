# 版權資訊
# 該程式的構想和設計者：Nichiya Sen
# 程式碼由以下AI：Grok、Gemini、ChatGPT、Claude協助，感謝科技的進步，讓不懂程式碼的新手也能做出可用的程式
# 完成於 2025/5/5
# 感謝您的使用

import os
import json
import datetime
import webbrowser
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog, Toplevel, Text, Scrollbar
from tkinter.ttk import Entry as TtkEntry, Progressbar, Style as TtkStyle, Scrollbar as TtkScrollbar, Style # Added Scrollbar and Style for ttk

# 嘗試載入 tkinterdnd2 模組
dnd_available = False # 新增一個旗標來記錄拖曳功能是否可用
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    dnd_available = True # 如果成功載入，設定旗標為 True
    print("tkinterdnd2 模組載入成功。拖曳功能可用。")
except ImportError:
    print("錯誤：未安裝 tkinterdnd2 模組。請先安裝該模組，才能使用拖放功能。")
    print("您可以嘗試以下指令安裝：pip install tkinterdnd2")
    DND_FILES = None # 確保 DND_FILES 是 None
    # TkinterDnD = object # 這裡不再需要將 TkinterDnD 設定為 object，依賴 dnd_available 旗標判斷


# 嘗試建立主視窗 (先嘗試 TkinterDnD.Tk，失敗則退回 Tk)
root = None
try:
    if dnd_available:
        try:
            root = TkinterDnD.Tk()
            print("成功建立 TkinterDnD 視窗。")
        except TclError as e:
            # 如果建立 TkinterDnD 視窗失敗 (例如環境不支援)，則退回使用標準 Tkinter
            print(f"警告：無法建立 TkinterDnD 視窗 ({e})。退回使用標準 Tkinter。")
            print("拖曳功能將無法使用。")
            root = Tk() # 退回使用標準 Tk 視窗
            dnd_available = False # 拖曳功能不可用
    else:
        # 如果 tkinterdnd2 模組就沒有載入成功，直接嘗試建立標準 Tk 視窗
        root = Tk()
        print("建立標準 Tkinter 視窗。拖曳功能無法使用。")

except TclError as e:
    # 如果建立標準 Tk 視窗也失敗 (表示沒有圖形介面環境)
    print(f"嚴重錯誤：無法建立圖形介面視窗 ({e})。請確認您的環境支援圖形顯示。")
    print("程式將結束。")
    root = None # 確保 root 是 None，阻止後續介面程式碼運行


from PIL import Image, ImageTk, ImageSequence
from collections import OrderedDict
import io  # 引入 io 模組，用於在記憶體中處理圖片
import traceback # 引入 traceback 模組，用於獲取完整的錯誤堆疊訊息

SETTINGS_FILE = "settings.json"
SUPPORTED_FORMATS = (".webp", ".jpg", ".jpeg", ".png", ".gif")

# 檢查是否支援動態 WebP
SUPPORTS_ANIMATED_WEBP = False  # 預設為 False
try:
    # 嘗試創建一個簡單的動態 WebP 來測試環境
    test_img = Image.new("RGB", (1, 1))
    test_img_path = "test.webp"
    test_img.save(test_img_path, "WEBP", save_all=True)
    Image.open(test_img_path).verify()  # 驗證檔案是否有效
    os.remove(test_img_path)
    SUPPORTS_ANIMATED_WEBP = True
except Exception as e:
    print(f"動態 WebP 支援檢查失敗：{e}")

# 語言字典
lang_zh_TW = {
    "title": "圖像轉換與命名工具", # 更新版本號
    "language_label": "語言：",
    "instructions_btn": "使用說明",
    "select_files_btn": "選擇圖片檔案",
    "import_folder_btn": "匯入資料夾",
    "select_output_btn": "選擇輸出資料夾",
    "output_label_default": "未選擇資料夾",
    "format_label": "轉出格式：",
    "add_date": "加上日期",
    "date_position_label": "日期位置：",
    "position_prefix": "前方",
    "position_suffix": "後方",
    "add_number": "自動編號",
    "number_position_label": "編號位置：",
    "prefix_label": "前綴字：",
    "batch_rename_btn": "批次改名 (統一檔名)",
    "reset_btn": "重新設定",
    "convert_btn": "開始轉換",
    "open_output_btn": "開啟輸出資料夾",
    "file_list_label": "檔案清單：",
    "toggle_all": "勾選/取消全部",
    "delete_selected_btn": "刪除選取項目",
    "clear_all_btn": "清除全部",
    "preview_label": "圖片預覽：",
    "instructions_title": "使用說明",
    "instructions_content": """
使用說明：

1. 選擇圖片檔案或匯入資料夾：
   - 點擊「選擇圖片檔案」或「匯入資料夾」，載入要轉換的圖片（支援 .webp、.jpg、.jpeg、.png、.gif）。
   - 或者，將圖片檔案直接拖曳到檔案清單中。

2. 設定轉換選項：
   - 選擇「轉出格式」（JPG、PNG、GIF、WEBP）。
   - 注意：
     - 動態 WebP 和動態 GIF 會根據選擇的格式轉換。
     - 選擇「GIF」：動態圖片轉為 GIF 格式。
     - 選擇「WEBP」：若環境支援，動態圖片轉為動態 WebP 格式，否則轉為 GIF。
   - 勾選「加上日期」或「自動編號」，可為檔名添加日期或編號：
     - 日期格式為「YYYYMMDD」（例如 20250505）。
     - 編號從指定的起始數字開始（預設為 1），根據檔案數量遞增。
     - 選擇「日期位置」或「編號位置」（前方或後方），決定日期或編號出現在檔名前還是檔名後。
   - 輸入「前綴字」，為檔名添加自訂前綴。
   - 重要：日期、編號、前綴的順序取決於你勾選或輸入的順序。例如：
     - 如果先勾選「加上日期」（設為前方），再勾選「自動編號」（設為前方），檔名會是「日期_編號_原始檔名」（例如「20250505_001_image.jpg」）。
     - 如果先輸入「前綴字」（例如「my_」），再勾選「加上日期」（設為前方），檔名會是「前綴_日期_原始檔名」（例如「my_20250505_image.jpg」）。
     - 順序由你勾選或輸入的先後決定，檔案清單中的預覽名稱會即時更新，方便確認。
   - 點擊「批次改名 (統一檔名)」可統一修改選取檔案的檔名。

3. 選擇輸出資料夾：
   - 點擊「選擇輸出資料夾」，指定轉換後檔案的儲存位置。

4. 開始轉換：
   - 點擊「開始轉換」，程式會將選取的圖片轉換為指定格式，並儲存到輸出資料夾。
   - 轉換完成後，可點擊「開啟輸出資料夾」查看結果。

5. 其他功能：
   - 點擊檔案清單中的檔案，可在右側預覽圖片。
   - 勾選「勾選/取消全部」來全選或取消選取檔案。
   - 點擊「刪除選取項目」或「清除全部」來移除檔案。
   - 點擊「重新設定」可恢復所有選項和檔名到初始狀態。

注意事項：
- 確保選擇了輸出資料夾，否則無法開始轉換。
- 轉換過程中若發生錯誤，會在完成後顯示錯誤訊息。
- 如果選擇將動態圖片轉為動態 WebP，但系統環境不支援，將會轉為 GIF。

資源使用建議：
- 本程式在一般使用場景下（載入幾十個檔案，轉換少量圖片）占用記憶體約 50-100 MB，CPU 占用低於 10%，硬碟 I/O 負載低。
- 若處理大量檔案（例如數百個）或大型動態圖片（例如 GIF），記憶體占用可能達到 200-300 MB，轉換期間 CPU 占用可能短暫升高。
- 建議分批處理大量檔案（例如一次轉換 50 個檔案），並避免頻繁預覽大型動態圖片，以降低資源占用。

關於本程式：
該程式的構想和設計者：Nichiya Sen
程式碼由以下AI：Grok、Gemini、ChatGPT、Claude協助，感謝科技的進步，讓不懂程式碼的新手也能做出可用的程式
完成於 2025/5/5
感謝您的使用
"""
}

lang_en = {
    "title": "Image Converter and Renamer Tool", # Update version
    "language_label": "Language:",
    "instructions_btn": "Instructions",
    "select_files_btn": "Select Image Files",
    "import_folder_btn": "Import Folder",
    "select_output_btn": "Select Output Folder",
    "output_label_default": "No Folder Selected",
    "format_label": "Output Format:",
    "add_date": "Add Date",
    "date_position_label": "Date Position:",
    "position_prefix": "Prefix",
    "position_suffix": "Suffix",
    "add_number": "Auto Number",
    "number_position_label": "Number Position:",
    "prefix_label": "Prefix:",
    "batch_rename_btn": "Batch Rename (Unified Name)",
    "reset_btn": "Reset",
    "convert_btn": "Start Conversion",
    "open_output_btn": "Open Output Folder",
    "file_list_label": "File List:",
    "toggle_all": "Select/Deselect All",
    "delete_selected_btn": "Delete Selected Items",
    "clear_all_btn": "Clear All",
    "preview_label": "Image Preview:",
    "instructions_title": "Instructions",
    "instructions_content": """
Instructions:

1. Select Image Files or Import Folder:
   - Click "Select Image Files" or "Import Folder" to load images for conversion (supports .webp, .jpg, .jpeg, .png, .gif).
   - Alternatively, drag and drop image files into the file list.

2. Set Conversion Options:
   - Choose "Output Format" (JPG, PNG, GIF, WEBP).
   - Note:
     - Animated WebP and GIF files will be converted based on the selected format.
     - Select "GIF": Animated images will be converted to GIF format.
     - Select "WEBP": If the environment supports it, animated images will be converted to animated WebP format, otherwise to GIF.
   - Check "Add Date" or "Auto Number" to add a date or number to the file name:
     - Date format is "YYYYMMDD" (e.g., 20250505).
     - Numbering starts from the specified number (default is 1) and increments based on the number of files.
     - Select "Date Position" or "Number Position" (Prefix or Suffix) to determine if the date or number appears before or after the file name.
   - Enter a "Prefix" to add a custom prefix to the file name.
   - Important: The order of date, number, and prefix depends on the sequence in which you enable them. For example:
     - If you check "Add Date" (set to Prefix) first, then "Auto Number" (set to Prefix), the file name will be "Date_Number_OriginalName" (e.g., "20250505_001_image.jpg").
     - If you enter a "Prefix" (e.g., "my_") first, then check "Add Date" (set to Prefix), the file name will be "Prefix_Date_OriginalName" (e.g., "my_20250505_image.jpg").
     - The order is determined by the sequence of your actions, and the preview name in the file list updates instantly for confirmation.
   - Click "Batch Rename (Unified Name)" to uniformly rename selected files.

3. Select Output Folder:
   - Click "Select Output Folder" to specify where converted files will be saved.

4. Start Conversion:
   - Click "Start Conversion" to convert the selected images to the specified format and save them to the output folder.
   - After conversion, click "Open Output Folder" to view the results.

5. Other Features:
   - Click a file in the file list to preview the image on the right.
   - Check "Select/Deselect All" to select or deselect all files.
   - Click "Delete Selected Items" or "Clear All" to remove files.
   - Click "Reset" to restore all options and file names to their initial state.

Notes:
- Ensure an output folder is selected, otherwise conversion cannot start.
- If errors occur during conversion, they will be displayed after completion.
- If you choose to convert animated images to animated WebP but the system environment does not support it, it will be converted to GIF.

Resource Usage Tips:
- In typical usage scenarios (loading a few dozen files, converting a small number of images), the program uses about 50-100 MB of memory, CPU usage is below 10%, and disk I/O is low.
- When handling a large number of files (e.g., hundreds) or large animated images (e.g., GIFs), memory usage may reach 200-300 MB, and CPU usage may spike temporarily during conversion.
- It is recommended to process large batches of files in smaller groups (e.g., convert 50 files at a time) and avoid frequent previews of large animated images to reduce resource usage.

About This Program:
Concept and Designer of this Program: Nichiya Sen
Code Assisted by the following AIs: Grok, Gemini, ChatGPT, Claude. Thanks to Technological Advancements, Even Novices Without Coding Knowledge Can Create Usable Programs
Completed on 2025/5/5
Thank you for using this program
"""
}

lang_ja = {
    "title": "画像変換・リネームツール", # Update version
    "language_label": "言語：",
    "instructions_btn": "使用説明",
    "select_files_btn": "画像ファイルを選択",
    "import_folder_btn": "フォルダをインポート",
    "select_output_btn": "出力フォルダを選択",
    "output_label_default": "フォルダが選択されていません",
    "format_label": "出力形式：",
    "add_date": "日付を追加",
    "date_position_label": "日付の位置：",
    "position_prefix": "前方",
    "position_suffix": "後方",
    "add_number": "自動番号付け",
    "number_position_label": "番号の位置：",
    "prefix_label": "プレフィックス：",
    "batch_rename_btn": "一括リネーム (統一ファイル名)",
    "reset_btn": "リセット",
    "convert_btn": "変換開始",
    "open_output_btn": "出力フォルダを開く",
    "file_list_label": "ファイルリスト：",
    "toggle_all": "全選択/全解除",
    "delete_selected_btn": "選択項目を削除",
    "clear_all_btn": "すべてクリア",
    "preview_label": "画像プレビュー：",
    "instructions_title": "使用説明",
    "instructions_content": """
使用説明：

1. 画像ファイルの選択またはフォルダのインポート：
   - 「画像ファイルを選択」または「フォルダをインポート」をクリックして、変換する画像を読み込みます（対応形式：.webp、.jpg、.jpeg、.png、.gif）。
   - または、画像ファイルを直接ファイルリストにドラッグ＆ドロップします。

2. 設定オプションの設定：
   - 「出力形式」（JPG、PNG、GIF、WEBP）を選択します。
   - 注意：
     - アニメーション WebP および GIF ファイルは選択した形式に基づいて変換されます。
     - 「GIF」を選択：アニメーション画像は GIF 形式に変換されます。
     - 「WEBP」を選択：環境がサポートしている場合、アニメーション画像はアニメーション WebP 形式に変換されます。そうでない場合は GIF に変換されます。
   - 「日付を追加」または「自動番号付け」をチェックして、ファイル名に日付や番号を追加します：
     - 日付形式は「YYYYMMDD」（例：20250505）です。
     - 番号は指定した開始数字（デフォルトは 1）から始まり、ファイル数に応じて増加します。
     - 「日付の位置」または「番号の位置」（前方または後方）を選択して、日付や番号がファイル名の前か後かを決定します。
   - 「プレフィックス」を入力して、ファイル名にカスタムプレフィックスを追加します。
   - 重要：日付、番号、プレフィックスの順序は、選択または入力した順序に依存します。例：
     - 最初に「日付を追加」（前方に設定）をチェックし、次に「自動番号付け」（前方に設定）をチェックした場合、ファイル名は「日付_番号_元のファイル名」（例：「20250505_001_image.jpg」）になります。
     - 最初に「プレフィックス」（例：「my_」）を入力し、次に「日付を追加」（前方に設定）をチェックした場合、ファイル名は「プレフィックス_日付_元のファイル名」（例：「my_20250505_image.jpg」）になります。
     - 順序は選択または入力のタイミングで決まり、ファイルリスト内のプレビュー名が即座に更新され、確認が簡単です。
   - 點擊「批次改名 (統一檔名)」可統一修改選取檔案的檔名。

3. 選擇輸出資料夾：
   - 點擊「選擇輸出資料夾」，指定轉換後檔案的儲存位置。

4. 変換の開始：
   - 點擊「開始轉換」，程式會將選取的圖片轉換為指定格式，並儲存到輸出資料夾。
   - 変換完成後，可點擊「開啟輸出資料夾」查看結果。

5. 其他功能：
   - 點擊檔案清單中的檔案，可在右側預覽圖片。
   - 勾選「勾選/取消全部」來全選或取消選取檔案。
   - 點擊「刪除選取項目」或「清除全部」來移除檔案。
   - 點擊「重新設定」可恢復所有選項和檔名到初始狀態。

注意事項：
- 確保選擇了輸出資料夾，否則無法開始轉換。
- 轉換過程中若發生錯誤，會在完成後顯示錯誤訊息。
- 如果選擇將動態圖片轉為動態 WebP，但系統環境不支援，將會轉為 GIF。

リソース使用建議：
- 通常の使用シナリオ（数十ファイルの読み込み、数枚の画像変換）では、プログラムは約 50～100 MB のメモリを使用し、CPU 使用率は 10% 以下、ディスク I/O 負荷は低いです。
- 多数のファイル（数百ファイルなど）や大型のアニメーション画像（GIF など）を処理する場合、メモリ使用量が 200～300 MB に達し、変換中に CPU 使用率が一時的に上昇することがあります。
- 大量のファイルを少量ずつ（例えば 1 回に 50 ファイル変換）処理し、大型アニメーション画像の頻繁なプレビューを避けることで、リソース使用量を軽減することをお勧めします。

このツールについて：
本プログラムの構想と設計者：Nichiya Sen
コードは以下のAI（Grok, Gemini, ChatGPT, Claude）によって支援されました。技術の進歩に感謝し、プログラミングの知識がない初心者でも使えるプログラムを作成できるようになりました
完成日：2025/5/5
ご利用いただきありがとうございます
"""
}

# 語言映射
languages = {
    "繁體中文": lang_zh_TW,
    "English": lang_en,
    "日本語": lang_ja
}

# 修正了 ImageConverterAndRenamerToolGUI 的定義，使其接受 master 參數
class ImageConverterAndRenamerToolGUI:
    def __init__(self, master):
        self.master = master
        self.current_language = self.load_settings().get("language", "繁體中文")  # 從設定載入語言，預設繁體中文
        self.lang = languages.get(self.current_language, lang_zh_TW) # Use .get with default
        master.title(self.lang["title"])
        master.geometry("1000x600")
        master.resizable(True, True)

        try:
            # 設定視窗圖示 - 請將 icon.ico 替換為你的圖示檔案名稱
            icon_path = "icon.ico"
            if os.path.exists(icon_path):
                master.iconbitmap(icon_path)
            else:
                print(f"警告：圖示檔案 '{icon_path}' 未找到。")
        except Exception as e:
            print(f"設定視窗圖示時發生錯誤：{e}")


        self.input_files = []
        self.initial_bases = {}
        self.original_bases = {}
        self.preview_fullnames = {}
        self.file_vars = {}
        self.file_rows = {}
        self.checkbuttons = {}
        self.action_order = []
        self.output_dir = ""
        self.anim_frames = []
        self.anim_index = 0
        self.anim_id = None
        self.successful_files = []
        self.settings = self.load_settings()
        last = self.settings.get("last_output", "")
        if last and os.path.isdir(last):
            self.output_dir = last

        # 圖片預覽緩存（最多儲存 10 張縮略圖）
        self.preview_cache = OrderedDict()
        self.preview_cache_limit = 10

        # 統一字體：微軟正黑體（zh_TW）、Arial（en, ja），大小 12
        self.font = ("Microsoft JhengHei" if self.current_language == "繁體中文" else "Arial", 12)

        # 設置 ttk 樣式（用於 OptionMenu, Entry, Progressbar, Scrollbar）
        self.style = TtkStyle() # Use TtkStyle
        self.style.configure("TEntry", font=self.font)
        # Corrected TMenubutton style configuration
        self.style.configure("TMenubutton", font=self.font, foreground="#333333", background="#FFFFFF", borderwidth=1)
        self.style.map("TMenubutton", background=[('active', '#EEEEEE')]) # Add hover effect

        # 主框架
        self.main_frame = Frame(master, bg="#FFFFFF")
        self.main_frame.pack(fill=BOTH, expand=True)

        # 左側控制區（使用 Canvas 實現滾動）
        self.fl_container = Frame(self.main_frame, bg="#FFFFFF")
        self.fl_container.grid(row=0, column=0, padx=5, pady=5, sticky="nsew") # 增加 padx/pady

        self.fl_canvas = Canvas(self.fl_container, width=250, height=550, bd=1, relief="solid", bg="#FFFFFF")  # 增加寬度
        self.fl_canvas.pack(side=LEFT, fill=BOTH, expand=True)

        self.fl_vsb = TtkScrollbar(self.fl_container, orient=VERTICAL, command=self.fl_canvas.yview)
        self.fl_vsb.pack(side=RIGHT, fill=Y)
        self.fl_canvas.configure(yscrollcommand=self.fl_vsb.set)

        self.fl = Frame(self.fl_canvas, bg="#FFFFFF")
        self.fl_canvas.create_window((0, 0), window=self.fl, anchor="nw")
        self.fl.bind("<Configure>", lambda e: self.fl_canvas.configure(scrollregion=self.fl_canvas.bbox("all")))

        # === 設定區塊 ===
        self.settings_frame = LabelFrame(self.fl, text="⚙ 設定", font=self.font, fg="#333333", bg="#FFFFFF", padx=5, pady=5) # 使用 LabelFrame
        self.settings_frame.pack(fill="x", padx=2, pady=5) # 增加 pady

        # 語言選擇 (移入設定區塊)
        self.language_label = Label(self.settings_frame, text=self.lang["language_label"], font=self.font, fg="#333333", bg="#FFFFFF")
        self.language_label.pack(anchor="w")
        self.language_var = StringVar(value=self.current_language)
        self.language_menu = OptionMenu(self.settings_frame, self.language_var, *languages.keys(), command=self.change_language)
        self.language_menu.config(font=self.font, fg="#333333", bg="#FFFFFF", bd=1, relief="solid", highlightthickness=0)
        self.language_menu["menu"].config(font=self.font)
        self.language_menu.pack(fill="x", pady=1)

        # 使用說明按鈕 (移入設定區塊)
        self.instructions_btn = Button(self.settings_frame, text=self.lang["instructions_btn"], font=self.font, bg="#FFFFFF", bd=1, relief="solid", fg="#333333", command=self.show_instructions, wraplength=200)
        self.instructions_btn.pack(fill="x", pady=1)
        # === 設定區塊結束 ===


        # 其他控制按鈕和選項 (保持在設定區塊下方)
        self.select_files_btn = Button(self.fl, text=self.lang["select_files_btn"], font=self.font, bg="#FFFFFF", bd=1, relief="solid", fg="#333333", command=self.select_image_files, width=20, wraplength=200)
        self.select_files_btn.pack(fill="x", padx=2, pady=5) # 增加 pady
        self.import_folder_btn = Button(self.fl, text=self.lang["import_folder_btn"], font=self.font, bg="#FFFFFF", bd=1, relief="solid", fg="#333333", command=self.import_folder, width=20, wraplength=200)
        self.import_folder_btn.pack(fill="x", padx=2, pady=5) # 增加 pady
        self.select_output_btn = Button(self.fl, text=self.lang["select_output_btn"], font=self.font, bg="#FFFFFF", bd=1, relief="solid", fg="#333333", command=self.select_output_folder, width=20, wraplength=200)
        self.select_output_btn.pack(fill="x", padx=2, pady=5) # 增加 pady
        self.output_path_label = Label(self.fl, text=self.output_dir or self.lang["output_label_default"], font=self.font, fg="#333333", wraplength=220, bg="#FFFFFF")
        self.output_path_label.pack(fill="x", padx=2, pady=1)

        self.format_label = Label(self.fl, text=self.lang["format_label"], font=self.font, fg="#333333", bg="#FFFFFF")
        self.format_label.pack(anchor="w", padx=2, pady=(10, 2)) # 增加上方 pady
        self.static_format = StringVar(master, "JPG")
        self.static_format.trace_add("write", lambda *a: self.refresh_previews())
        self.format_menu = OptionMenu(self.fl, self.static_format, "JPG", "PNG", "GIF", "WEBP")
        self.format_menu.config(font=self.font, fg="#333333", bg="#FFFFFF", bd=1, relief="solid", highlightthickness=0) # Added highlightthickness=0
        self.format_menu["menu"].config(font=self.font) # Set font for dropdown menu
        self.format_menu.pack(fill="x", padx=2, pady=1)

        self.add_date = BooleanVar()
        self.add_date_cb = Checkbutton(self.fl, text=self.lang["add_date"], font=self.font, variable=self.add_date, fg="#333333", bg="#FFFFFF",
                                       command=lambda: self.toggle_action("date", self.add_date.get()))
        self.add_date_cb.pack(anchor="w", padx=2, pady=(5,0)) # 增加上方 pady
        self.df = Frame(self.fl, bg="#FFFFFF"); self.df.pack(anchor="w", pady=(1,0), padx=2)
        self.date_position_label = Label(self.df, text=self.lang["date_position_label"], font=self.font, fg="#333333", bg="#FFFFFF")
        self.date_position_label.pack(side=LEFT)
        self.date_position = StringVar(value="prefix")
        self.date_prefix_rb = Radiobutton(self.df, text=self.lang["position_prefix"], font=self.font, variable=self.date_position, value="prefix", fg="#333333", bg="#FFFFFF",
                                          command=self.refresh_previews)
        self.date_prefix_rb.pack(side=LEFT)
        self.date_suffix_rb = Radiobutton(self.df, text=self.lang["position_suffix"], font=self.font, variable=self.date_position, value="suffix", fg="#333333", bg="#FFFFFF",
                                          command=self.refresh_previews)
        self.date_suffix_rb.pack(side=LEFT)

        self.nf = Frame(self.fl, bg="#FFFFFF"); self.nf.pack(anchor="w", pady=(5,0), padx=2) # 增加上方 pady
        self.add_number = BooleanVar()
        self.add_number_cb = Checkbutton(self.nf, text=self.lang["add_number"], font=self.font, variable=self.add_number, fg="#333333", bg="#FFFFFF",
                                         command=lambda: self.toggle_action("number", self.add_number.get()))
        self.add_number_cb.pack(side=LEFT)
        self.number_start = StringVar(value="1")
        self.number_entry = TtkEntry(self.nf, width=5, textvariable=self.number_start)
        self.number_entry.pack(side=LEFT, padx=2)
        self.number_start.trace_add("write", lambda *a: self.refresh_previews())
        self.pf = Frame(self.fl, bg="#FFFFFF"); self.pf.pack(anchor="w", pady=(1,0), padx=2)
        self.number_position_label = Label(self.pf, text=self.lang["number_position_label"], font=self.font, fg="#333333", bg="#FFFFFF")
        self.number_position_label.pack(side=LEFT)
        self.number_position = StringVar(value="suffix")
        self.number_prefix_rb = Radiobutton(self.pf, text=self.lang["position_prefix"], font=self.font, variable=self.number_position, value="prefix", fg="#333333", bg="#FFFFFF",
                                            command=self.refresh_previews)
        self.number_prefix_rb.pack(side=LEFT)
        self.number_suffix_rb = Radiobutton(self.pf, text=self.lang["position_suffix"], font=self.font, variable=self.number_position, value="suffix", fg="#333333", bg="#FFFFFF",
                                            command=self.refresh_previews)
        self.number_suffix_rb.pack(side=LEFT)

        self.prefix_label = Label(self.fl, text=self.lang["prefix_label"], font=self.font, fg="#333333", bg="#FFFFFF")
        self.prefix_label.pack(anchor="w", padx=2, pady=(5,0)) # 增加上方 pady
        self.prefix_text = StringVar()
        self.prefix_entry = Entry(self.fl, textvariable=self.prefix_text, font=self.font)
        self.prefix_entry.pack(fill="x", padx=2, pady=(0,1))
        self.prefix_text.trace_add("write", lambda *a: self.toggle_action("prefix", bool(self.prefix_text.get())))

        self.batch_rename_btn = Button(self.fl, text=self.lang["batch_rename_btn"], font=self.font, bg="#FFFFFF", bd=1, relief="solid", fg="#333333", command=self.batch_rename, wraplength=220)
        self.batch_rename_btn.pack(fill="x", padx=2, pady=5) # 增加 pady
        self.reset_btn = Button(self.fl, text=self.lang["reset_btn"], font=self.font, bg="#FFFFFF", bd=1, relief="solid", fg="#333333", command=self.reset_options, wraplength=220)
        self.reset_btn.pack(fill="x", padx=2, pady=5) # 增加 pady

        self.convert_btn = Button(self.fl, text=self.lang["convert_btn"], font=self.font, bg="#555555", fg="#FFFFFF", bd=1, relief="solid", command=self.convert_images, wraplength=220)
        self.convert_btn.pack(fill="x", padx=2, pady=10) # 增加上方 pady

        self.open_folder_btn = Button(self.fl, text=self.lang["open_output_btn"], font=self.font, bg="#FFFFFF", bd=1, relief="solid", fg="#333333", command=self.open_output_folder,
                                     state="normal" if self.output_dir else "disabled", wraplength=220)
        self.open_folder_btn.pack(fill="x", padx=2, pady=5) # 增加 pady

        self.progress = Progressbar(self.fl, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=2, pady=5) # 增加 pady

        # 中間檔案清單（使用 Canvas 和 Frame 嵌入 Checkbutton）
        self.fc = Frame(self.main_frame, bg="#FFFFFF")
        self.fc.grid(row=0, column=1, padx=5, pady=5, sticky="nsew") # 增加 padx/pady 並確保 sticky="nsew"
        self.file_list_label = Label(self.fc, text=self.lang["file_list_label"], font=self.font, fg="#333333", bg="#FFFFFF")
        self.file_list_label.pack(anchor="w", padx=2)
        self.lf = Frame(self.fc, bg="#FFFFFF"); self.lf.pack(fill=BOTH, expand=True)

        # 使用 grid 佈局確保滾輪位置正確
        self.lf.grid_rowconfigure(0, weight=1)
        self.lf.grid_columnconfigure(0, weight=1)
        self.canvas = Canvas(self.lf, width=500, height=250, bd=1, relief="solid", bg="#FFFFFF")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb = TtkScrollbar(self.lf, orient=VERTICAL, command=self.canvas.yview)
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb = TtkScrollbar(self.lf, orient=HORIZONTAL, command=self.canvas.xview)
        self.hsb.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.check_frame = Frame(self.canvas, bg="#FFFFFF")
        self.canvas.create_window((0,0), window=self.check_frame, anchor="nw")
        self.check_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.toggle_all = BooleanVar()
        self.toggle_all_cb = Checkbutton(self.fc, text=self.lang["toggle_all"], font=self.font, variable=self.toggle_all, fg="#333333", bg="#FFFFFF",
                                         command=lambda: [self.toggle_all_files(), self.refresh_previews()])
        self.toggle_all_cb.pack(fill="x", padx=2, pady=(5,1)) # 增加上方 pady
        self.delete_selected_btn = Button(self.fc, text=self.lang["delete_selected_btn"], font=self.font, bg="#FFFFFF", bd=1, relief="solid", fg="#333333", command=self.delete_selected_files, wraplength=200)
        self.delete_selected_btn.pack(fill="x", padx=2, pady=(0,5)) # 增加下方 pady
        self.clear_all_btn = Button(self.fc, text=self.lang["clear_all_btn"], font=self.font, bg="#FFFFFF", bd=1, relief="solid", fg="#333333", command=self.clear_all_files, wraplength=200)
        self.clear_all_btn.pack(fill="x", padx=2, pady=(0,5)) # 增加下方 pady

        # 右側預覽
        self.fr = Frame(self.main_frame, bg="#FFFFFF")
        self.fr.grid(row=0, column=2, padx=5, pady=5, sticky="nsew") # 增加 padx/pady 並確保 sticky="nsew"
        self.preview_label = Label(self.fr, text=self.lang["preview_label"], font=self.font, fg="#333333", bg="#FFFFFF")
        self.preview_label.pack(anchor="w", padx=2)
        # 預覽圖片標籤設定 expand=True 讓它可以填滿可用空間
        self.preview_image_label = Label(self.fr, width=300, height=150, bd=1, relief="solid", bg="#FFFFFF"); self.preview_image_label.pack(padx=2, pady=2, fill=BOTH, expand=True)
        self.fullpath_label = Label(self.fr, text="", font=self.font, fg="#333333", wraplength=280, justify=LEFT, bg="#FFFFFF"); self.fullpath_label.pack(pady=(2,1), padx=2)

        # 設定 main_frame 的欄位權重，讓中間和右側可以隨著視窗放大而擴展
        self.main_frame.grid_columnconfigure(0, weight=0) # 左側控制區不需要隨意放大
        self.main_frame.grid_columnconfigure(1, weight=1) # 檔案清單區塊可以水平放大
        self.main_frame.grid_columnconfigure(2, weight=1) # 預覽區塊可以水平放大
        self.main_frame.grid_rowconfigure(0, weight=1) # 讓主要內容區塊可以垂直放大


        # 綁定滑鼠滾輪事件到頂層視窗
        self.master.bind("<MouseWheel>", self.handle_mousewheel)

        # 為左側控制區和檔案清單的子 widget 綁定滾輪事件
        self.bind_scroll_events(self.fl_container, self.fl_canvas)
        self.bind_scroll_events(self.lf, self.canvas)

        # 確認拖曳功能是否可用，並設定拖放綁定
        if dnd_available:
             self.master.drop_target_register(DND_FILES)
             self.master.dnd_bind('<<Drop>>', self.drop_files)
        # 這裡可以選擇不印出訊息，因為在前面建立視窗時已經印過了
        # else:
        #      print("拖曳功能無法使用。")


    def bind_scroll_events(self, widget, canvas):
        """為 widget 及其所有子 widget 綁定滑鼠滾輪事件"""
        # Added check to prevent binding if widget is destroyed
        if widget and widget.winfo_exists():
            widget.bind("<MouseWheel>", lambda event: self.handle_mousewheel_for_canvas(event, canvas))
            for child in widget.winfo_children():
                self.bind_scroll_events(child, canvas)

    def handle_mousewheel_for_canvas(self, event, canvas):
        """處理指定 Canvas 的滑鼠滾輪事件"""
        # Added check if canvas exists
        if canvas and canvas.winfo_exists():
            scrollregion = canvas.bbox("all")
            if scrollregion:
                height = scrollregion[3] - scrollregion[1]
                if height > canvas.winfo_height():
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return "break"
            return "break"  # 即使沒有可滾動空間，也阻止事件傳播
        return None # Allow default if canvas doesn't exist

    def handle_mousewheel(self, event):
        """處理滑鼠滾輪事件，根據滑鼠位置分發到適當的 Canvas 或 Text"""
        # 獲取滑鼠當前位置（相對於主視窗）
        # Added check if master exists
        if not self.master or not self.master.winfo_exists():
            return None

        x, y = event.x_root - self.master.winfo_rootx(), event.y_root - self.master.winfo_rooty()

        # 檢查滑鼠是否在左側控制區（self.fl_container 及其子 widget）
        if self.fl_container and self.fl_container.winfo_exists(): # Added check
            fl_x, fl_y = self.fl_container.winfo_x(), self.fl_container.winfo_y()
            fl_width, fl_height = self.fl_container.winfo_width(), self.fl_container.winfo_height()
            if (fl_x <= x <= fl_x + fl_width and fl_y <= y <= fl_y + fl_height):
                return self.handle_mousewheel_for_canvas(event, self.fl_canvas)


        # 檢查滑鼠是否在檔案清單（self.lf 及其子 widget）
        if self.lf and self.lf.winfo_exists(): # Added check
            lf_x, lf_y = self.lf.winfo_x(), self.lf.winfo_y()
            lf_width, lf_height = self.lf.winfo_width(), self.lf.winfo_height()
            if (lf_x <= x <= lf_x + lf_width and lf_y <= y <= lf_y + lf_height):
                 return self.handle_mousewheel_for_canvas(event, self.canvas)


        # 如果滑鼠不在可滾動區域，允許事件繼續傳播
        return None

    def change_language(self, selected_language):
        """切換語言並更新介面"""
        self.current_language = selected_language
        self.lang = languages.get(self.current_language, lang_zh_TW) # Use .get with default

        # 更新字體（根據語言選擇適當字體）
        self.font = ("Microsoft JhengHei" if self.current_language == "繁體中文" else "Arial", 12)
        self.style.configure("TEntry", font=self.font)
        # Corrected TMenubutton style configuration
        self.style.configure("TMenubutton", font=self.font, foreground="#333333", background="#FFFFFF", borderwidth=1)
        self.style.map("TMenubutton", background=[('active', '#EEEEEE')]) # Add hover effect


        # Check if master window exists before updating widgets
        if not self.master or not self.master.winfo_exists():
            return

        # Update title
        self.master.title(self.lang["title"])

        # Update left control area
        # 語言選擇和使用說明按鈕已經移入 settings_frame，這裡不再需要單獨更新
        # if self.language_label.winfo_exists(): self.language_label.config(text=self.lang["language_label"], font=self.font)
        # if self.instructions_btn.winfo_exists(): self.instructions_btn.config(text=self.lang["instructions_btn"], font=self.font)

        # 更新設定區塊的標題
        if self.settings_frame.winfo_exists():
             self.settings_frame.config(text="⚙ 設定" if self.current_language == "繁體中文" else "⚙ Settings" if self.current_language == "English" else "⚙ 設定", font=self.font)
             # 更新設定區塊內的語言標籤和使用說明按鈕文字
             if self.language_label.winfo_exists(): self.language_label.config(text=self.lang["language_label"], font=self.font)
             if self.instructions_btn.winfo_exists(): self.instructions_btn.config(text=self.lang["instructions_btn"], font=self.font)
             if self.language_menu.winfo_exists(): # Update font for the OptionMenu itself and its dropdown
                  self.language_menu.config(font=self.font)
                  self.language_menu["menu"].config(font=self.font)


        if self.select_files_btn.winfo_exists(): self.select_files_btn.config(text=self.lang["select_files_btn"], font=self.font)
        if self.import_folder_btn.winfo_exists(): self.import_folder_btn.config(text=self.lang["import_folder_btn"], font=self.font)
        if self.select_output_btn.winfo_exists(): self.select_output_btn.config(text=self.lang["select_output_btn"], font=self.font)
        if self.output_path_label.winfo_exists(): self.output_path_label.config(text=self.output_dir or self.lang["output_label_default"], font=self.font)
        if self.format_label.winfo_exists(): self.format_label.config(text=self.lang["format_label"], font=self.font)
        # Corrected OptionMenu configuration
        if self.format_menu.winfo_exists():
             self.format_menu.config(font=self.font)
             self.format_menu["menu"].config(font=self.font)

        if self.add_date_cb.winfo_exists(): self.add_date_cb.config(text=self.lang["add_date"], font=self.font)
        if self.date_position_label.winfo_exists(): self.date_position_label.config(text=self.lang["date_position_label"], font=self.font)
        if self.date_prefix_rb.winfo_exists(): self.date_prefix_rb.config(text=self.lang["position_prefix"], font=self.font,  bg="#FFFFFF",)
        if self.date_suffix_rb.winfo_exists(): self.date_suffix_rb.config(text=self.lang["position_suffix"], font=self.font,  bg="#FFFFFF",)
        if self.add_number_cb.winfo_exists(): self.add_number_cb.config(text=self.lang["add_number"], font=self.font)
        if self.number_entry.winfo_exists(): self.number_entry.config(font=self.font)
        if self.number_position_label.winfo_exists(): self.number_position_label.config(text=self.lang["number_position_label"], font=self.font,  bg="#FFFFFF",)
        if self.number_prefix_rb.winfo_exists(): self.number_prefix_rb.config(text=self.lang["position_prefix"], font=self.font,  bg="#FFFFFF",)
        if self.number_suffix_rb.winfo_exists(): self.number_suffix_rb.config(text=self.lang["position_suffix"], font=self.font,  bg="#FFFFFF",)
        if self.prefix_label.winfo_exists(): self.prefix_label.config(text=self.lang["prefix_label"], font=self.font)
        if self.prefix_entry.winfo_exists(): self.prefix_entry.config(font=self.font)
        if self.batch_rename_btn.winfo_exists(): self.batch_rename_btn.config(text=self.lang["batch_rename_btn"], font=self.font)
        if self.reset_btn.winfo_exists(): self.reset_btn.config(text=self.lang["reset_btn"], font=self.font)
        if self.convert_btn.winfo_exists(): self.convert_btn.config(text=self.lang["convert_btn"], font=self.font)
        if self.open_folder_btn.winfo_exists(): self.open_folder_btn.config(text=self.lang["open_output_btn"], font=self.font)
        # self.progress.config(font=self.font) # Progressbar does not have 'font' option

        # Update middle file list
        if self.file_list_label.winfo_exists(): self.file_list_label.config(text=self.lang["file_list_label"], font=self.font)
        if self.toggle_all_cb.winfo_exists(): self.toggle_all_cb.config(text=self.lang["toggle_all"], font=self.font)
        if self.delete_selected_btn.winfo_exists(): self.delete_selected_btn.config(text=self.lang["delete_selected_btn"], font=self.font)
        if self.clear_all_btn.winfo_exists(): self.clear_all_btn.config(text=self.lang["clear_all_btn"], font=self.font)

        # Update right preview area
        if self.preview_label.winfo_exists(): self.preview_label.config(text=self.lang["preview_label"], font=self.font)
        if self.fullpath_label.winfo_exists(): self.fullpath_label.config(font=self.font)

        # Update file names display in the file list
        for fp in list(self.file_rows.keys()): # Iterate over a copy of keys
            if fp in self.file_rows: # Check if key still exists
                _, _, orig_lbl, preview_lbl = self.file_rows[fp]
                if orig_lbl and orig_lbl.winfo_exists():
                    orig_lbl.config(font=self.font)
                if preview_lbl and preview_lbl.winfo_exists():
                    preview_lbl.config(font=self.font)


        # Save language preference
        self.settings["language"] = self.current_language
        self.save_settings()


    def show_instructions(self):
        """顯示使用說明視窗"""
        instructions_window = Toplevel(self.master)
        instructions_window.title(self.lang["instructions_title"])
        # Modified initial geometry and made resizable
        instructions_window.geometry("600x700") # Increased initial size
        instructions_window.resizable(True, True) # Made resizable

        # 說明文字
        instructions_text = Text(instructions_window, wrap="word", font=("Microsoft JhengHei" if self.current_language == "繁體中文" else "Arial", 10), height=25, width=50)
        instructions_text.pack(padx=8, pady=8, fill=BOTH, expand=True)

        # 加入滾動條
        scrollbar = Scrollbar(instructions_text)
        scrollbar.pack(side=RIGHT, fill=Y)
        instructions_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=instructions_text.yview)

        # 綁定滑鼠滾輪事件給使用說明視窗
        # Added check if widget exists
        if instructions_text.winfo_exists():
             instructions_text.bind("<MouseWheel>", lambda event: instructions_text.yview_scroll(int(-1 * (event.delta / 120)), "units"))


        # 插入使用說明內容（已包含關於資訊）
        instructions_text.insert(END, self.lang["instructions_content"])
        instructions_text.config(state=DISABLED)  # 設為唯讀

        # 關閉按鈕
        Button(instructions_window, text="關閉" if self.current_language == "繁體中文" else "Close" if self.current_language == "English" else "閉じる",
               font=self.font, command=instructions_window.destroy).pack(pady=8)

    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE,'r',encoding='utf-8') as f: return json.load(f)
            except json.JSONDecodeError:
                print("Settings file is empty or corrupted, creating a new one.")
                return {}
            except Exception as e:
                 print(f"Error loading settings file: {e}")
                 return {}
        return {}

    def save_settings(self):
        try:
            with open(SETTINGS_FILE,'w',encoding='utf-8') as f: json.dump(self.settings,f, indent=4) # Added indent for readability
        except Exception as e:
            print(f"Error saving settings file: {e}")

    def select_image_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Image files","*.webp *.jpg *.jpeg *.png *.gif")])
        self.add_files(files)
        if self.input_files:
            # Added check if the file list is not empty before trying to show preview
            if self.input_files:
                 self.show_preview(self.input_files[-1])


    def import_folder(self):
        d = filedialog.askdirectory()
        if not d: return
        # Filter files to only include supported formats and ensure they are files, not directories
        files = [os.path.join(d, f) for f in os.listdir(d) if os.path.isfile(os.path.join(d, f)) and f.lower().endswith(SUPPORTED_FORMATS)]
        self.add_files(files)
        if self.input_files:
            # Added check if the file list is not empty before trying to show preview
            if self.input_files:
                self.show_preview(self.input_files[-1])


    def add_files(self, files):
        for fp in files: # Iterating through files directly
            # Added check if the file path already exists to prevent duplicates
            if fp in self.file_vars:
                continue

            if not os.path.isfile(fp): # Ensure it's a file
                continue

            base, ext = os.path.splitext(os.path.basename(fp))
            self.input_files.append(fp)
            self.initial_bases[fp] = base
            self.original_bases[fp] = base
            self.preview_fullnames[fp] = ""
            var = IntVar(value=1)  # 自動勾選新檔案
            # Added check if check_frame exists before creating widgets
            if self.check_frame and self.check_frame.winfo_exists():
                row = Frame(self.check_frame, bg="#FFFFFF")
                # Use grid for file list items for better control
                row.grid(sticky="ew", padx=1, pady=1) # Use grid instead of pack
                self.check_frame.grid_columnconfigure(0, weight=1) # Allow the frame to expand

                cb = Checkbutton(row, variable=var, command=lambda path=fp: self.refresh_previews(), bg="#FFFFFF");
                cb.grid(row=0, column=0, sticky="w") # Use grid

                orig_lbl = Label(row, text=os.path.basename(fp), anchor="w", font=self.font, fg="#333333", bg="#FFFFFF");
                orig_lbl.grid(row=0, column=1, sticky="ew") # Use grid
                row.grid_columnconfigure(1, weight=1) # Allow label to expand

                preview_lbl = Label(row, text="", fg="#1E90FF", anchor="w", font=self.font, bg="#FFFFFF");
                preview_lbl.grid(row=0, column=2, sticky="ew") # Use grid
                row.grid_columnconfigure(2, weight=1) # Allow label to expand


                # Bind events to labels
                orig_lbl.bind("<Double-Button-1>", lambda e, path=fp: self.rename_single(path))
                orig_lbl.bind("<Button-1>", lambda e, path=fp: self.show_preview(path))
                preview_lbl.bind("<Button-1>", lambda e, path=fp: self.show_preview(path)) # Bind preview label too

                self.file_vars[fp] = var
                self.checkbuttons[fp] = cb
                self.file_rows[fp] = (row, cb, orig_lbl, preview_lbl)

        self.refresh_previews()
        # Update the scrollable region of the canvas after adding files
        if self.check_frame and self.check_frame.winfo_exists():
            self.canvas.config(scrollregion=self.canvas.bbox("all"))


    def drop_files(self, event):
        files = self.master.tk.splitlist(event.data)
        valid_files = []
        for fp in files:
            if os.path.isdir(fp):
                # Recursively add files from subdirectories if needed, but for now just top level
                for f in os.listdir(fp):
                    f_path = os.path.join(fp, f)
                    if os.path.isfile(f_path) and f.lower().endswith(SUPPORTED_FORMATS):
                        valid_files.append(f_path)
            elif os.path.isfile(fp) and fp.lower().endswith(SUPPORTED_FORMATS):
                valid_files.append(fp)

        self.add_files(valid_files)

        if self.input_files:
            # Added check if the file list is not empty before trying to show preview
            if self.input_files:
                self.show_preview(self.input_files[-1])


    def show_preview(self, filepath):
        if self.anim_id:
            self.master.after_cancel(self.anim_id)
            self.anim_id = None
        self.anim_frames, self.anim_index = [], 0

        # Clear previous preview
        if self.preview_image_label.winfo_exists():
            self.preview_image_label.config(image="")
        if self.fullpath_label.winfo_exists():
            self.fullpath_label.config(text="")

        if not os.path.exists(filepath):
             if self.fullpath_label.winfo_exists():
                  self.fullpath_label.config(text=f"無法預覽：檔案不存在\n{filepath}" if self.current_language == "繁體中文" else f"Cannot preview: File does not exist\n{filepath}" if self.current_language == "English" else f"プレビューできません：ファイルが存在しません\n{filepath}")
             return

        try:
            # Use a simpler cache key
            # Added check if widgets exist before getting dimensions
            target_width = self.preview_image_label.winfo_width() if self.preview_image_label.winfo_exists() else 300
            target_height = self.preview_image_label.winfo_height() if self.preview_image_label.winfo_exists() else 150

            cache_key = (filepath, target_width, target_height)
            if cache_key in self.preview_cache:
                ph = self.preview_cache[cache_key]
                if self.preview_image_label.winfo_exists():
                    self.preview_image_label.config(image=ph); self.preview_image_label.image = ph
                if self.fullpath_label.winfo_exists():
                    self.fullpath_label.config(text=filepath)
                # Move to the end of OrderedDict to signify recent use
                self.preview_cache.move_to_end(cache_key)
                return


            img = Image.open(filepath)
            if self.fullpath_label.winfo_exists():
                self.fullpath_label.config(text=filepath)

            # 計算圖片適配尺寸，保持寬高比
            img_width, img_height = img.size
            ratio = min(target_width / img_width, target_height / img_height)
            new_width = int(img_width * ratio)
            new_height = int(img_height * ratio)

            # Ensure dimensions are at least 1x1
            new_width = max(1, new_width)
            new_height = max(1, new_height)


            # 檢查是否為動畫圖片（在 resize 之前檢查）
            is_animated = getattr(img, "is_animated", False)
            if is_animated:
                self.anim_frames = [] # Clear previous frames
                try:
                    # Limit the number of frames to prevent excessive memory usage for large GIFs
                    max_frames = 50 # Limit to 50 frames for preview
                    for i, frame in enumerate(ImageSequence.Iterator(img)):
                         if i >= max_frames:
                             print(f"Warning: Limiting preview to {max_frames} frames for {os.path.basename(filepath)}")
                             break
                         tmp = frame.copy()
                         # Convert frame to RGBA to handle transparency and pasting onto background
                         tmp = tmp.convert("RGBA")
                         tmp = tmp.resize((new_width, new_height), Image.Resampling.LANCZOS)
                         tmp_background = Image.new('RGBA', (target_width, target_height), (255, 255, 255, 255))
                         offset = ((target_width - new_width) // 2, (target_height - new_height) // 2)
                         tmp_background.paste(tmp, offset, tmp) # Use tmp as mask for transparency
                         self.anim_frames.append(ImageTk.PhotoImage(tmp_background))

                    if self.anim_frames:
                         self._animate()
                    else:
                        # Handle case where animation has no frames (shouldn't happen with valid images)
                         if self.preview_image_label.winfo_exists():
                             self.preview_image_label.config(image="")
                         if self.fullpath_label.winfo_exists():
                             self.fullpath_label.config(text=f"無法預覽：動畫檔案無影格\n{filepath}" if self.current_language == "繁體中文" else f"Cannot preview: Animated file has no frames\n{filepath}" if self.current_language == "English" else f"プレビューできません：アニメーションファイルにフレームがありません\n{filepath}")


                except Exception as anim_e:
                     # Fallback to showing the first frame or an error message
                     print(f"Error processing animation frames for {os.path.basename(filepath)}: {anim_e}")
                     try:
                         # Try to load just the first frame
                         img.seek(0) # Go back to the first frame
                         img = img.convert("RGBA") # Convert to RGBA
                         img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                         background = Image.new('RGBA', (target_width, target_height), (255, 255, 255, 255))
                         offset = ((target_width - new_width) // 2, (target_height - new_height) // 2)
                         background.paste(img, offset, img) # Use img as mask for transparency
                         ph = ImageTk.PhotoImage(background)
                         if self.preview_image_label.winfo_exists():
                             self.preview_image_label.config(image=ph); self.preview_image_label.image = ph
                         if self.fullpath_label.winfo_exists():
                             self.fullpath_label.config(text=filepath)
                         # Add to cache
                         self.preview_cache[cache_key] = ph
                         # Manage cache size
                         if len(self.preview_cache) > self.preview_cache_limit:
                             self.preview_cache.popitem(last=False) # Remove the oldest item

                     except Exception as fallback_e:
                         if self.preview_image_label.winfo_exists():
                             self.preview_image_label.config(image="")
                         if self.fullpath_label.winfo_exists():
                             self.fullpath_label.config(text=f"無法預覽：{str(fallback_e)}\n{filepath}" if self.current_language == "繁體中文" else f"Cannot preview: {str(fallback_e)}\n{filepath}" if self.current_language == "English" else f"プレビューできません：{str(fallback_e)}\n{filepath}")


            else:
                # Convert to RGBA to handle potential transparency issues when pasting onto background
                img = img.convert("RGBA")
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                background = Image.new('RGBA', (target_width, target_height), (255, 255, 255, 255))
                offset = ((target_width - new_width) // 2, (target_height - new_height) // 2)
                background.paste(img, offset, img) # Use img as mask for transparency
                ph = ImageTk.PhotoImage(background)
                if self.preview_image_label.winfo_exists():
                    self.preview_image_label.config(image=ph); self.preview_image_label.image = ph

                # Add to cache
                self.preview_cache[cache_key] = ph
                # Manage cache size
                if len(self.preview_cache) > self.preview_cache_limit:
                    self.preview_cache.popitem(last=False) # Remove the oldest item


            img.close() # Close the image file after processing

        except Exception as e:
            if self.preview_image_label.winfo_exists():
                self.preview_image_label.config(image="")
            if self.fullpath_label.winfo_exists():
                self.fullpath_label.config(text=f"無法預覽：{str(e)}\n{filepath}" if self.current_language == "繁體中文" else f"Cannot preview: {str(e)}\n{filepath}" if self.current_language == "English" else f"プレビューできません：{str(e)}\n{filepath}")


    def _animate(self):
        if not self.anim_frames: return
        # Added check if widget exists before updating
        if self.preview_image_label and self.preview_image_label.winfo_exists():
            frame = self.anim_frames[self.anim_index]
            self.preview_image_label.config(image=frame)
            self.preview_image_label.image = frame
            self.anim_index = (self.anim_index + 1) % len(self.anim_frames)
            self.anim_id = self.master.after(100, self._animate)

    def select_output_folder(self):
        d = filedialog.askdirectory()
        if d:
            self.output_dir = d
            self.settings['last_output'] = d
            if self.output_path_label.winfo_exists(): # Added check
                 self.output_path_label.config(text=d)
            if self.open_folder_btn.winfo_exists(): # Added check
                 self.open_folder_btn.config(state="normal")
            self.save_settings()

    def rename_single(self, fp):
        # Added check if master exists before showing dialog
        if self.master and self.master.winfo_exists():
            new = simpledialog.askstring("修改檔名" if self.current_language == "繁體中文" else "Rename File" if self.current_language == "English" else "ファイル名を変更",
                                     "輸入不含副檔名的新名稱：" if self.current_language == "繁體中文" else "Enter the new name without extension:" if self.current_language == "English" else "拡張子なしの新しい名前を入力：",
                                     initialvalue=self.original_bases.get(fp, "")) # Use get with default for safety

            if new is not None: # Check if user cancelled
                new = new.strip()
                if new: # Only update if new name is not empty after strip
                    # 只更新 original_bases，不修改 initial_bases
                    self.original_bases[fp] = new
                    _,_,orig_lbl,_ = self.file_rows.get(fp)
                    if orig_lbl and orig_lbl.winfo_exists(): # Check if label exists
                         ext = os.path.splitext(fp)[1]
                         orig_lbl.config(text=f"{new}{ext}")
                    self.refresh_previews()


    def batch_rename(self):
        sel = [fp for fp,v in self.file_vars.items() if v.get()==1]
        if not sel:
            # Added check if master exists before showing messagebox
            if self.master and self.master.winfo_exists():
                 messagebox.showwarning("警告", "請先勾選檔案")
            return

        # Added check if master exists before showing dialog
        if self.master and self.master.winfo_exists():
            common = simpledialog.askstring("批次改名" if self.current_language == "繁體中文" else "Batch Rename" if self.current_language == "English" else "一括リネーム",
                                        "輸入不含副檔名的新名稱：" if self.current_language == "繁體中文" else "Enter a unified name without extension:" if self.current_language == "English" else "拡張子なしの新しい名前を入力：")

            if common is not None: # Check if user cancelled
                 common = common.strip()
                 if common: # Only rename if common name is not empty after strip
                     for fp in sel:
                         # 只更新 original_bases，不修改 initial_bases
                         self.original_bases[fp] = common
                         _,_,orig_lbl,_ = self.file_rows.get(fp)
                         if orig_lbl and orig_lbl.winfo_exists(): # Check if label exists
                              ext = os.path.splitext(fp)[1]
                              orig_lbl.config(text=f"{common}{ext}")

                     # Ensure prefix is added to action_order if common is not empty
                     if self.prefix_text.get() and "prefix" not in self.action_order:
                         self.action_order.append("prefix")
                     elif not self.prefix_text.get() and "prefix" in self.action_order:
                          self.action_order.remove("prefix")

                     self.refresh_previews()


    def reset_options(self):
        self.add_date.set(False)
        self.date_position.set("prefix")
        self.add_number.set(False)
        self.number_start.set("1")
        self.number_position.set("suffix")
        self.prefix_text.set("")
        self.action_order.clear()
        for fp in list(self.input_files): # Iterate over a copy of the list
            if fp in self.initial_bases: # Check if the file still exists in initial_bases
                self.original_bases[fp] = self.initial_bases[fp]
                _,_,orig_lbl,_ = self.file_rows.get(fp)
                if orig_lbl and orig_lbl.winfo_exists(): # Check if label exists
                     ext = os.path.splitext(fp)[1]
                     orig_lbl.config(text=f"{self.initial_bases[fp]}{ext}")
            else:
                 # If the file was deleted, remove it from original_bases as well
                 self.original_bases.pop(fp, None)

        self.refresh_previews()

    def toggle_action(self, name, active):
        if active and name not in self.action_order:
            self.action_order.append(name)
        # Maintain the order based on when actions were last enabled/disabled
        # Remove and re-append to move to the end if becoming active
        if not active and name in self.action_order:
            self.action_order.remove(name)
        elif active and name in self.action_order:
             self.action_order.remove(name)
             self.action_order.append(name)

        self.refresh_previews()

    def refresh_previews(self):
        # Clear all preview labels first
        for fp in list(self.file_vars.keys()): # Iterate over a copy of keys
            if fp in self.file_rows: # Check if key still exists
                _,_,_,pl = self.file_rows[fp]
                if pl and pl.winfo_exists(): # Check if label exists
                    pl.config(text="")

        # Filter to only selected files for preview update
        selected_files = [fp for fp, v in self.file_vars.items() if v.get() == 1]

        for idx, fp in enumerate(self.input_files):
            # Only update preview for selected files
            if fp not in selected_files:
                continue

            row_widgets = self.file_rows.get(fp)
            if not row_widgets: # Skip if file row doesn't exist (e.g., deleted)
                continue

            _, _, _, pl = row_widgets

            if not pl or not pl.winfo_exists(): # Check if label exists
                 continue


            base = self.original_bases.get(fp, os.path.splitext(os.path.basename(fp))[0]) # Use get with default
            ds = datetime.datetime.now().strftime("%Y%m%d")
            try: s = int(self.number_start.get())
            except ValueError: s = 1 # Handle non-integer input for number_start
            num = str(s+idx).zfill(len(self.number_start.get()))
            fmt = self.static_format.get().lower()

            try:
                # Determine output extension based on format selection and animation
                final_ext = fmt # Default to selected format
                if os.path.exists(fp): # Check if file still exists before opening
                    try:
                        with Image.open(fp) as img:
                            is_animated = getattr(img, "is_animated", False)
                            if is_animated:
                                if fmt == "webp" and SUPPORTS_ANIMATED_WEBP:
                                    final_ext = "webp"
                                else:
                                    # Fallback to gif if animated webp is not supported or target is gif
                                    final_ext = "gif"
                            else: # Static image
                                # Ensure correct static extension is used if it differs from selected format (e.g. converting GIF to JPG)
                                if fmt in ("jpg", "jpeg", "png", "webp", "gif"):
                                     final_ext = fmt
                                else:
                                    # Fallback to original extension if selected format is not a standard image format
                                    final_ext = os.path.splitext(fp)[1].lower().lstrip(".")

                    except Exception as img_check_e:
                         print(f"Warning: Could not check if {os.path.basename(fp)} is animated: {img_check_e}. Assuming static.")
                         # Fallback to selected format if check fails
                         final_ext = fmt
            except Exception as e:
                 print(f"Error determining final extension for {os.path.basename(fp)}: {e}")
                 # Fallback to selected format if any error occurs
                 final_ext = fmt


            out_name_preview = self.preview_fullnames.get(fp)
            if out_name_preview:
                # Use the calculated preview name and replace its extension with the determined final extension
                base_name, _ = os.path.splitext(out_name_preview)
                out_name = f"{base_name}.{final_ext}"
            else:
                # Fallback if preview name not found, use original base name and final extension
                base_name = self.original_bases.get(fp, os.path.splitext(os.path.basename(fp))[0]) # Use get with default
                out_name = f"{base_name}.{final_ext}"


            out_path = os.path.normpath(os.path.join(self.output_dir, out_name)) # Use self.output_dir

            # Avoid leading/trailing underscores if parts were empty
            name_parts = []
            for act in self.action_order:
                if act=="prefix" and self.prefix_text.get(): name_parts.append(self.prefix_text.get())
                elif act=="date":
                    if self.date_position.get()=="prefix": name_parts.append(ds)
                elif act=="number":
                    if self.number_position.get()=="prefix": name_parts.append(num)

            name_parts.append(base) # Always include the base name

            for act in self.action_order:
                 if act=="date":
                     if self.date_position.get()=="suffix": name_parts.append(ds)
                 elif act=="number":
                     if self.number_position.get()=="suffix": name_parts.append(num)


            # Join with underscore, handle cases where parts might be empty resulting in multiple underscores
            name = "_".join(part for part in name_parts if part)

            # Replace multiple underscores with a single one
            while '__' in name:
                name = name.replace('__', '_')

            # Remove leading/trailing underscores after joining
            name = name.strip('_')

            # Handle case where name might become empty after stripping (e.g. only prefixes/suffixes added)
            if not name:
                name = base # Fallback to original base name if derived name is empty


            final_name = f"{name}.{final_ext}"
            pl.config(text=final_name)
            self.preview_fullnames[fp] = final_name

        # Update the scrollable region of the canvas
        if self.check_frame and self.check_frame.winfo_exists():
            self.canvas.config(scrollregion=self.canvas.bbox("all"))


    def toggle_all_files(self):
        st = 1 if self.toggle_all.get() else 0
        for v in self.file_vars.values(): v.set(st)
        self.refresh_previews()

    def delete_selected_files(self):
        sel = [fp for fp,v in self.file_vars.items() if v.get()==1]
        if not sel: return # Do nothing if no files selected

        for fp in sel:
            # Added check if file row exists before trying to delete
            if fp in self.file_rows:
                row,cb,ol,pl = self.file_rows[fp]
                # Added check if widget exists before destroying
                if row and row.winfo_exists():
                    row.destroy()

                if fp in self.input_files:
                    self.input_files.remove(fp)

                # Remove from all relevant dictionaries
                for d in (self.file_vars, self.initial_bases, self.original_bases, self.preview_fullnames, self.file_rows, self.checkbuttons):
                    d.pop(fp, None)

        # Clear preview if the selected file was the one being previewed
        # Check if self.fullpath_label exists and is not destroyed
        if self.fullpath_label and self.fullpath_label.winfo_exists() and self.fullpath_label.cget("text") in sel:
             # Check if self.preview_image_label exists and is not destroyed
             if self.preview_image_label and self.preview_image_label.winfo_exists():
                 self.preview_image_label.config(image="")
             if self.fullpath_label.winfo_exists():
                 self.fullpath_label.config(text="")
             if self.anim_id:
                 self.master.after_cancel(self.anim_id)
                 self.anim_id = None
             self.anim_frames = []


        self.refresh_previews()
        # Update the scrollable region of the canvas after deleting files
        if self.check_frame and self.check_frame.winfo_exists():
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
        # If no files left, uncheck toggle all and disable delete/clear buttons
        if not self.input_files:
             self.toggle_all.set(False)
             # You might want to disable delete/clear buttons here if needed


    def clear_all_files(self):
        # Destroy all widgets in the check_frame
        if self.check_frame and self.check_frame.winfo_exists():
            for w in self.check_frame.winfo_children():
                w.destroy()

        # Clear all lists and dictionaries
        self.input_files.clear()
        self.file_vars.clear()
        self.initial_bases.clear()
        self.original_bases.clear()
        self.preview_fullnames.clear()
        self.file_rows.clear()
        self.checkbuttons.clear()
        self.successful_files.clear()

        # Clear preview area
        # Check if self.preview_image_label exists and is not destroyed
        if self.preview_image_label and self.preview_image_label.winfo_exists():
            self.preview_image_label.config(image="")
        # Check if self.fullpath_label exists and is not destroyed
        if self.fullpath_label and self.fullpath_label.winfo_exists():
            self.fullpath_label.config(text="")


        # Cancel any ongoing animation
        if self.anim_id:
            self.master.after_cancel(self.anim_id)
            self.anim_id = None
        self.anim_frames = []

        # Clear preview cache
        self.preview_cache.clear()

        # Update scrollable region
        if self.check_frame and self.check_frame.winfo_exists():
            self.canvas.config(scrollregion=self.canvas.bbox("all"))

        # Reset toggle all checkbox
        self.toggle_all.set(False)
        # You might want to disable delete/clear buttons here


    def generate_unique_filename(self, out_path):
        """生成唯一的檔案名稱，自動添加後綴以避免衝突"""
        # Added check if the file exists before generating a unique name
        if not os.path.exists(out_path):
            return out_path

        base, ext = os.path.splitext(out_path)
        counter = 1
        # Add an underscore before the counter to clearly separate it
        new_path = f"{base}_{counter}{ext}"
        while os.path.exists(new_path):
            counter += 1
            new_path = f"{base}_{counter}{ext}"
        return new_path

    def convert_single_image(self, fp, output_dir, sel):
        """單個圖片轉換邏輯，供多執行緒使用"""
        out_path = None # Initialize out_path to None
        error_message = None # Initialize error_message

        try:
            if not os.path.exists(fp):
                error_message = f"檔案 {os.path.basename(fp)} 轉換失敗：原始檔案不存在，可能已被移動或刪除"
                # Using master.after to show messagebox from the main thread
                if self.master and self.master.winfo_exists():
                    self.master.after(0, lambda: messagebox.showerror("錯誤", error_message))
                return fp, False, error_message

            # Validate image file integrity before proceeding
            try:
                with Image.open(fp) as img_check:
                    img_check.verify()
            except Exception as verify_e:
                 error_message = f"檔案 {os.path.basename(fp)} 轉換失敗：圖片檔案損壞或無效 - {verify_e}"
                 if self.master and self.master.winfo_exists():
                     self.master.after(0, lambda: messagebox.showerror("錯誤", error_message))
                 return fp, False, error_message


            with Image.open(fp) as img:
                # 檢查是否為動態檔案（動態 WebP 或動態 GIF）
                is_anim = getattr(img, "is_animated", False)
                # 根據輸出格式決定轉換目標
                input_ext = os.path.splitext(fp)[1].lower().lstrip(".")
                output_format = self.static_format.get().lower()

                # Determine the final output extension based on animation and selected format
                if is_anim:
                    if output_format == "webp" and SUPPORTS_ANIMATED_WEBP:
                         final_ext = "webp"
                    else: # Convert animated to GIF if target is GIF or animated webp not supported
                         final_ext = "gif"
                else: # Static image conversion
                    final_ext = output_format


                out_name_preview = self.preview_fullnames.get(fp)
                if out_name_preview:
                    # Use the calculated preview name and replace its extension with the determined final extension
                    base_name, _ = os.path.splitext(out_name_preview)
                    out_name = f"{base_name}.{final_ext}"
                else:
                    # Fallback if preview name not found, use original base name and final extension
                    base_name = self.original_bases.get(fp, os.path.splitext(os.path.basename(fp))[0]) # Use get with default
                    out_name = f"{base_name}.{final_ext}"


                out_path = os.path.normpath(os.path.join(output_dir, out_name))

                # Check for output path conflicts BEFORE attempting conversion
                # Generate a unique filename immediately if conflict exists
                original_out_path = out_path
                out_path = self.generate_unique_filename(out_path)
                if out_path != original_out_path:
                     conflict_message = f"檔案 {os.path.basename(fp)} 的輸出路徑與現有檔案衝突，已自動調整為 {os.path.basename(out_path)}"
                     # Using master.after to show messagebox from the main thread
                     if self.master and self.master.winfo_exists():
                         self.master.after(0, lambda msg=conflict_message: messagebox.showwarning("路徑衝突", msg))


                # If input and output formats are the same and filename hasn't changed, copy the file
                # Also check if the output path is different from the input path
                if input_ext == final_ext and os.path.normpath(fp) != out_path and os.path.basename(fp) == out_name:
                     try:
                         shutil.copy2(fp, out_path)
                         print(f"Copied {os.path.basename(fp)} to {os.path.basename(out_path)}")
                         return fp, True, None
                     except Exception as e:
                         error_message = f"複製檔案 {os.path.basename(fp)} 失敗: {e}"
                         if self.master and self.master.winfo_exists():
                             self.master.after(0, lambda: messagebox.showerror("錯誤", error_message))
                         return fp, False, error_message


                # Otherwise, perform the conversion
                start_time = datetime.datetime.now()
                try:
                    if is_anim:
                        # Ensure all frames are converted to RGBA before saving
                        frames = []
                        try:
                            for f in ImageSequence.Iterator(img):
                                frames.append(f.copy().convert("RGBA"))
                        except Exception as frame_copy_e:
                            error_message = f"處理動畫影格失敗：{frame_copy_e}"
                            if self.master and self.master.winfo_exists():
                                self.master.after(0, lambda: messagebox.showerror("轉換錯誤", error_message))
                            return fp, False, error_message


                        if frames: # Check if there are any frames
                            if final_ext == "webp":
                                # Animated webp save
                                try:
                                    frames[0].save(out_path, "WEBP", save_all=True, append_images=frames[1:], loop=0, duration=img.info.get("duration", 100))
                                except Exception as webp_anim_e:
                                     # Fallback to GIF and report error if animated webp saving fails unexpectedly
                                     error_message = f"動態 WebP 儲存失敗，嘗試轉為 GIF：{webp_anim_e}"
                                     if self.master and self.master.winfo_exists():
                                         self.master.after(0, lambda: messagebox.showwarning("轉換警告", error_message))
                                     # Try saving as GIF
                                     try:
                                         frames[0].save(out_path, "GIF", save_all=True, append_images=frames[1:], loop=0, duration=img.info.get("duration", 100))
                                         # Note: If original target was .webp and fallback to GIF happens, the output file will be a GIF.
                                         # The filename will still have the .webp extension due to how out_name is generated
                                         # This is a potential inconsistency that could be improved.
                                     except Exception as gif_fallback_e:
                                         error_message = f"動態 WebP 轉 GIF 也失敗：{gif_fallback_e}"
                                         if self.master and self.master.winfo_exists():
                                             self.master.after(0, lambda: messagebox.showerror("轉換錯誤", error_message))
                                         return fp, False, error_message

                            else:  # final_ext == "gif"
                                frames[0].save(out_path, "GIF", save_all=True, append_images=frames[1:], loop=0, duration=img.info.get("duration", 100))
                        else:
                             # Handle animated file with no frames
                             error_message = f"檔案 {os.path.basename(fp)} 轉換失敗：動畫檔案無有效影格"
                             if self.master and self.master.winfo_exists():
                                 self.master.after(0, lambda: messagebox.showerror("錯誤", error_message))
                             return fp, False, error_message

                    else: # Static conversion
                        if final_ext in ("jpg", "jpeg"):
                            img = img.convert("RGB") # Ensure JPG is RGB
                            img.save(out_path, "JPEG", quality=95) # Use quality 95 as a good default
                        elif final_ext == "png":
                            # PNG can handle transparency, convert to RGBA if original has alpha or is paletted with transparency
                            if img.mode in ('P', 'RGBA', 'LA'):
                                img = img.convert("RGBA")
                            img.save(out_path, "PNG")
                        elif final_ext == "webp":
                             # Static webp save - handle transparency
                             if img.mode in ('P', 'RGBA', 'LA'):
                                 img = img.convert("RGBA")
                                 img.save(out_path, "WEBP", quality=95, lossless=True) # Use lossless for images with alpha
                             else:
                                 img = img.convert("RGB")
                                 img.save(out_path, "WEBP", quality=95) # Use quality for lossy RGB

                        elif final_ext == "gif": # Convert static image to GIF
                             img.save(out_path, "GIF") # PIL saves static images as single-frame GIFs
                        else:
                             error_message = f"不支援的輸出格式：{final_ext}"
                             if self.master and self.master.winfo_exists():
                                 self.master.after(0, lambda: messagebox.showerror("錯誤", error_message))
                             return fp, False, error_message

                except Exception as conversion_e:
                    # This catch is for errors during the actual saving process
                    error_message = f"轉換檔案 {os.path.basename(fp)} 失敗：{str(conversion_e)}\n輸出路徑：{out_path}\n追蹤堆疊：{traceback.format_exc()}"
                    if self.master and self.master.winfo_exists():
                         self.master.after(0, lambda: messagebox.showerror("轉換錯誤", error_message))
                    return fp, False, error_message

                end_time = datetime.datetime.now()
                duration = (end_time - start_time).total_seconds()
                print(f"Converted {os.path.basename(fp)} to {os.path.basename(out_path)} in {duration:.2f} seconds")
                return fp, True, None

        except Exception as e:
            # This catch is for errors outside the inner try block, e.g., file opening issues
            error_message = f"處理檔案 {os.path.basename(fp)} 時發生錯誤：{str(e)}\n追蹤堆疊：{traceback.format_exc()}"
            if self.master and self.master.winfo_exists():
                self.master.after(0, lambda: messagebox.showerror("處理錯誤", error_message))
            return fp, False, error_message


    def convert_images(self):
        sel = [fp for fp,v in self.file_vars.items() if v.get()==1]
        if not sel:
            if self.master and self.master.winfo_exists(): # Added check
                 messagebox.showwarning("警告", "請先勾選檔案")
            return
        if not self.output_dir:
            if self.master and self.master.winfo_exists(): # Added check
                messagebox.showerror("錯誤", "先選輸出資料夾")
            return

        output_dir = os.path.normpath(self.output_dir)
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                if self.master and self.master.winfo_exists(): # Added check
                     messagebox.showerror("錯誤", f"無法創建輸出資料夾 {output_dir}：{str(e)}")
                return

        try:
            test_file = os.path.join(output_dir, "test_write_permission.txt")
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
        except Exception as e:
            if self.master and self.master.winfo_exists(): # Added check
                messagebox.showerror("錯誤", f"無法寫入輸出資料夾 {output_dir}：{str(e)}\n請確認是否有寫入權限，或嘗試以管理員權限運行程式")
            return

        success, fail = 0, 0
        processed_count = 0 # <-- Added initialization here
        self.progress["value"] = 0
        # Added check to prevent division by zero if sel is empty (should be caught earlier but good practice)
        total_files = len(sel) if sel else 0
        step = 100 / total_files if total_files > 0 else 0
        conflict_messages = []

        # Disable convert button during conversion
        if self.convert_btn.winfo_exists(): # Added check
            self.convert_btn.config(state="disabled")

        def run_conversion(file_list, output_directory, selected_files):
            """在背景執行緒中執行圖片轉換"""
            nonlocal success, fail, conflict_messages, processed_count  # 宣告要修改的變數
            # processed_count = 0 # Removed initialization here

            for fp in file_list:
                if not self.master or not self.master.winfo_exists():
                    break # 視窗已關閉，停止轉換

                result = self.convert_single_image(fp, output_directory, selected_files)
                if result[1]:
                    success += 1
                    # self.successful_files.append(fp) # No need to store successful files here if not used later
                else:
                    fail += 1
                    # Only append unique conflict messages to avoid duplicates
                    if result[2] and result[2] not in conflict_messages:
                        conflict_messages.append(result[2])

                processed_count += 1
                # Update progress bar on the main thread
                progress_value = (processed_count / total_files) * 100 if total_files > 0 else 0
                if self.master and self.master.winfo_exists():
                    self.master.after(0, lambda val=progress_value: self.progress.config(value=val))


            # Conversion complete, call completion handler on the main thread
            if self.master and self.master.winfo_exists():
                self.master.after(0, self.conversion_complete, success, fail, conflict_messages)

        # 啟動背景執行緒
        # Pass the list of selected files to run_conversion to allow internal path conflict checking
        self.thread = ThreadPoolExecutor(max_workers=4)  # 使用執行緒池
        self.thread.submit(run_conversion, sel, output_dir, sel)


    def update_progress(self):
        """更新進度條 (No longer needed as progress is updated directly in run_conversion)"""
        pass # Keep for compatibility if other parts call it, but logic moved


    def conversion_complete(self, success, fail, conflict_messages):
        """轉換完成後顯示訊息"""
        if self.master and self.master.winfo_exists():
            # Re-enable convert button
            if self.convert_btn.winfo_exists(): # Added check
                 self.convert_btn.config(state="normal")

            self.refresh_previews() # Refresh previews to show potential filename changes due to conflicts
            if conflict_messages:
                # Join conflict messages with newline for better readability
                messagebox.showinfo("轉換完成 - 訊息紀錄", "\n".join(conflict_messages))

            # Show final summary message
            messagebox.showinfo("完成", f"成功 {success} 筆，失敗 {fail} 筆" if self.current_language == "繁體中文" else f"Success: {success}, Failed: {fail}" if self.current_language == "English" else f"成功：{success} 件、失敗：{fail} 件")

        # Optional: Clear successful files list after completion if not longer needed
        # self.successful_files.clear() # Moved clearing inside conversion_complete

        # Reset progress bar to 0 after completion
        if self.progress.winfo_exists():
             self.progress.config(value=0)


    def open_output_folder(self):
        if self.output_dir and os.path.exists(self.output_dir): # Added check if directory exists
            try:
                webbrowser.open(self.output_dir)
            except Exception as e:
                if self.master and self.master.winfo_exists(): # Added check
                     messagebox.showerror("錯誤", f"無法開啟資料夾 {self.output_dir}：{str(e)}")
        elif self.master and self.master.winfo_exists(): # Added check
             messagebox.showwarning("警告", "輸出資料夾不存在或尚未選擇")


# Main execution block
if __name__ == '__main__':
    # Check if root was successfully created (either TkinterDnD.Tk or Tk) in the try/except block above
    if root:
        # Instantiate the GUI class, passing the root window
        app = ImageConverterAndRenamerToolGUI(root)

        # Start the Tkinter event loop
        root.mainloop()
    else:
        # If root is None, it means window creation failed (either TkinterDnD or standard Tk).
        # An error message was already printed.
        pass # 程式結束
