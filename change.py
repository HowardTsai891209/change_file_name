import os
import pandas as pd
from natsort import natsorted

# 定義資料夾路徑和 Excel 檔案路徑
current_folder = os.path.dirname(os.path.abspath(__file__))#預設全部放在同一個資料夾內
audio_folder = os.path.join(current_folder, 'audio')#錄音檔資料夾名稱
excel_file = os.path.join(current_folder, 'list.xlsx')#excel名稱

# 讀取 Excel 檔案並手動設定欄位名稱
df = pd.read_excel(excel_file, sheet_name='1', usecols='A:C', header=None)
df.columns = ['編號', '准考證號', '姓名']

# 獲取音頻檔案列表並按自然排序和修改時間排序
audio_files = natsorted(os.listdir(audio_folder), key=lambda x: os.path.getmtime(os.path.join(audio_folder, x)))

# 確保 Excel 和音頻檔案數量一致
if len(audio_files) != len(df):
    print("音頻檔案數量和 Excel 項目數量不一致")
    exit()

# 依據排序後的音頻檔案，對應修改檔名
for i, audio_file in enumerate(audio_files):
    new_name = f"{df.at[i, '編號']}_{df.at[i, '准考證號']}_{df.at[i, '姓名']}{os.path.splitext(audio_file)[1]}"
    old_path = os.path.join(audio_folder, audio_file)
    new_path = os.path.join(audio_folder, new_name)
    os.rename(old_path, new_path)
    print(f"Renamed: {audio_file} to {new_name}")

print("檔名修改完成")
