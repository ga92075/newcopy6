# pip install numpy pyperclip xlwings keyboard selenium pandas pyinstaller packaging
# Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
# Set-ExecutionPolicy RemoteSigned -Scope Process
# python -m venv venv
# .\venv\Scripts\activate
# https://github.com/upx/upx/releases
# pyinstaller --onefile --exclude-module matplotlib --exclude-module scipy --exclude-module PyQt5 --exclude-module PySide2 --exclude-module tkinter --exclude-module openpyxl --exclude-module xlrd --exclude-module pyarrow --upx-dir . newcopy6.py


import io
import os
import math
import sys
from collections import Counter
import numpy as np
import re
import pyperclip
from selenium.webdriver.edge.service import Service
from datetime import datetime, timedelta
import xlwings as xw
import keyboard
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import Select # 必須導入 Select 類別
from selenium.common.exceptions import TimeoutException, NoSuchWindowException
import time
import random
import pandas as pd
import subprocess         
from packaging.version import parse  
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchWindowException, InvalidSessionIdException
stop_requested = False
Include_anti=True
WAIT_TIMEOUT2 = 0.2
Include_body_weight=True
include_gas=True
Include_drug=True
Include_IO=True
include_tumormarker=True
Include_thyroid = False
WAIT_TIMEOUT = 5
patient=10
rest=1.3
app = xw.apps.active
wb = app.books.active
ws = wb.sheets.active
driver_instance = None



def on_esc_press():
    global stop_requested # 聲明要修改的是全局變數
    stop_requested = True
    print("\n偵測到 'f8' 鍵，正在準備安全退出...")

# 註冊熱鍵，在腳本啟動時執行
keyboard.add_hotkey('f8', on_esc_press)

def get_edge_browser_version():
    """
    【V2.0 - 更可靠】
    使用 'reg query' 命令檢查 Windows 登錄檔，獲取 Edge 瀏覽器的版本。
    這不受安裝路徑影響。
    """
    try:
        # 這是最穩定的登錄檔位置 (查詢目前使用者的版本)
        cmd = r'reg query "HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon" /v version'
        output = subprocess.check_output(cmd, shell=True, text=True, stderr=subprocess.DEVNULL)
        
        # 輸出範例:
        # HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon
        #     version    REG_SZ    128.0.2739.0
        
        # 使用正規表達式抓取版本號
        match = re.search(r'version\s+REG_SZ\s+([\d\.]+)', output)
        if match:
            return match.group(1).strip()

    except Exception:
        # 如果 HKEY_CURRENT_USER 失敗 (例如系統帳戶執行)，嘗試 HKEY_LOCAL_MACHINE
        try:
            cmd = r'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\EdgeUpdate\Clients\{56EB18F5-8B09-4E90-A4F4-1B05D53267E0}" /v pv /t REG_SZ'
            output = subprocess.check_output(cmd, shell=True, text=True, stderr=subprocess.DEVNULL)
            
            # 輸出範例:
            # HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\EdgeUpdate\Clients\{...}
            #     pv    REG_SZ    128.0.2739.0
            
            match = re.search(r'pv\s+REG_SZ\s+([\d\.]+)', output)
            if match:
                return match.group(1).strip()
        except Exception:
            print("錯誤：無法自動偵測 Edge 瀏覽器安裝版本。")
            print("已嘗試查詢 HKEY_CURRENT_USER 和 HKEY_LOCAL_MACHINE 登錄檔，均失敗。")
            return None

    # 如果 'try' 成功但 'match' 失敗
    print("錯誤：解析 Edge 瀏覽器版本時出錯。")
    return None

def get_driver_version(driver_path):
    """
    執行 msedgedriver.exe --version 並解析其版本號。
    """
    if not os.path.exists(driver_path):
        return None
    try:
        cmd = [driver_path, '--version']
        # 執行命令並獲取輸出
        output = subprocess.check_output(cmd, text=True, stderr=subprocess.DEVNULL)
        
        # 輸出範例: Microsoft Edge WebDriver 128.0.2739.0 (...)
        match = re.search(r'Microsoft Edge WebDriver (\d+\.\d+\.\d+\.\d+)', output)
        if match:
            return match.group(1)
    except Exception as e:
        print(f"錯誤：無法執行 {driver_path} --version。")
        print(f"({e})")
        return None
    return None

def check_driver_compatibility(driver_path):
    """
    獨立的函式：檢查 msedgedriver.exe 版本是否與 Edge 瀏覽器相符。
    如果不相符，將會顯示錯誤訊息並中止程式。
    
    參數:
    driver_path (str): 'msedgedriver.exe' 的完整路徑。
    """
    
    # --- 版本相容性檢查 ---
    print("\n--- 版本相容性檢查 ---")
    
    # 檢查 driver 檔案是否存在
    if not os.path.exists(driver_path):
        print(f"錯誤：找不到驅動程式！")
        print(f"請確認 'msedgedriver.exe' 檔案存在於")
        print(f"{driver_path}")
        print(f"請確認 'msedgedriver.exe' 檔案存在於")
        input("\n按 Enter 鍵結束...")
        sys.exit()

    driver_ver = get_driver_version(driver_path)
    browser_ver = get_edge_browser_version()

    if driver_ver and browser_ver:
        print(f"找到驅動程式版本: {driver_ver}")
        print(f"找到瀏覽器版本: {browser_ver}")
        
        driver_major = parse(driver_ver).major
        browser_major = parse(browser_ver).major

        if driver_major != browser_major:
            print("\n" + "="*50)
            print("【錯誤：版本不相符！】")
            print(f"您的 Edge 瀏覽器主版號是: {browser_major}")
            print(f"您的 msedgedriver.exe 主版號是: {driver_major}")
            
            # --- 【依照您的要求修改錯誤訊息】 ---
            print("\n請手動完成以下更新")
            print("1. 從\"https://developer.microsoft.com/zh-tw/microsoft-edge/tools/webdriver?form=MA13LH#downloads\"下載最新版本的msedgedriver.exe ")
            print("2. 打開Edge>右上設定>設定>關於Microsoft Edge>更新最新的Edge版本")
            print("="*50)
            # Add a countdown timer for 10 seconds
            #print("\n倒數10秒，按 Enter 立即繼續運行...")
            for i in range(10, 0, -1):
                print(f"\r剩餘 {i} 秒...", end="", flush=True)
                time.sleep(1)
            print("\r倒數完成!\n")
            try:
                print("\n按 Enter 繼續嘗試運行(不建議)，按esc或任意鍵結束...")
                # 獲取用戶的按鍵
                key = keyboard.read_event(suppress=True).name
                if key != 'enter':
                    print(f"偵測到 '{key}' 鍵，結束程式。")
                    safe_exit()
                    sys.exit()
                print("偵測到 Enter，繼續程式。")
                time.sleep(0.1)  # 等待按鍵放開
            except Exception as e:
                print(f"發生錯誤: {e}")
                # 若 keyboard 模組無法使用，回到基本的 input
                resp = input("\n按 Enter 繼續運行，輸入任何其他鍵結束: ")
                if resp != "":
                    print("輸入非空值，結束程式。")
                    safe_exit()
                    sys.exit()
        
        
        print("------------------------\n")

    else:
        print("警告：無法自動完成版本驗證，將嘗試直接啟動...")
        print("------------------------\n")
    # --- 版本檢查結束 ---

def safe_exit(driver=None):
    """
    執行資源的安靜清理。
    這個版本更加穩健，會檢查 driver 是否存在才進行關閉。
    """
    print("正在執行資源清理程序...")

    # 移除熱鍵監聽
    try:
        if keyboard.is_hooked('esc'):
            keyboard.remove_hotkey('esc')
            print("已移除 'Esc' 鍵熱鍵監聽。")
    except Exception as e:
        print(f"移除熱鍵時發生錯誤: {e}")

    # 關閉 Selenium WebDriver
    if driver:
        try:
            print("正在關閉 Selenium WebDriver...")
            driver.quit()  # 使用 quit() 能確保瀏覽器和驅動進程都被關閉
            print("Selenium WebDriver 已成功關閉。")
        except Exception as e:
            print(f"關閉 WebDriver 時發生錯誤: {e}")
            
    print("清理完成！")


def move_to_col_and_paste(cell=app.selection, col=1, row_offset=0, text=None, use_paste=True):
    """
    通用函數：移動到指定欄位並選擇性粘貼文字
    參數:
    - cell: 起始儲存格 (預設為當前選取)
    - col: 目標欄位 (1=A, 2=B, 3=C, 4=D, 5=E...)
    - row_offset: 行偏移 (0=同一行, 1=下一行, -1=上一行)
    - text: 要粘貼的文字 (None=不粘貼)
    - use_paste: True=使用快捷鍵粘貼(可復原), False=直接賦值(不可復原)
    """
    starttime = time.time()
    while time.time() - starttime < 15:
        try:
            target_cell = ws.cells(cell.row + row_offset, col)
            target_cell.select()
            
            if text is not None:
                target_cell.value = text
            
            return target_cell

        except Exception as e:
            time.sleep(2)
            print(f"excel non accessible, try again in 2 seconds: {e}")


def move_to_lab_col_paste(cell=app.selection, text=None):
    """移動到同row的第4欄(實驗室欄位)並粘貼"""
    return move_to_col_and_paste(cell, col=4, text=text)

def move_to_drug_col_paste(cell=app.selection, text=None):
    """移動到同row的第5欄(藥物欄位)並粘貼"""
    return move_to_col_and_paste(cell, col=5, text=text)

def move_to_row_start(cell=app.selection, column=1, down=0):
    """移動到指定欄位和行偏移"""
    return move_to_col_and_paste(cell, col=column, row_offset=down)

def get_today_date_formatted() -> str:
    """
    獲取今天的日期，並將其格式化為 (月/日) 的字符串。

    Returns:
        str: 格式化後的日期字符串，例如 "(8/3)"。
    """
    # 1. 獲取今天的日期和時間物件
    today = datetime.now()
    
    # 2. 使用 f-string 提取月份和日期，並組合成目標格式
    #    today.month -> 月份 (數字)
    #    today.day   -> 日期 (數字)
    formatted_date = f"({today.month}/{today.day})"
    
    return formatted_date

def process_gas_data():
    """
    從剪貼簿讀取實驗室數據，將數據轉換為數字格式（除了日期列），
    找出每個欄位最後兩個非 NaN 的值，並在相應欄位添加趨勢比較的摘要行。
    輸出時會刪除小數點後的零（例如 12.0 會顯示為 12）。
    
    Returns:
    DataFrame: 處理後的 DataFrame，包含每個欄位最後兩個非 NaN 值
              以及顯示值比較的附加摘要行（放在各自的列中）。
    """
    try:
        # 從剪貼簿讀取數據
        df = pd.read_clipboard()
        df = df.drop(df.index[-1])
        #df['日期'] = pd.to_datetime(df['日期'], format='%y/%m%d %H:%M', errors='coerce')
    except Exception as e:
        print(f"讀取剪貼簿錯誤: {e}")
        return None
    
 
    # 保存第一列（通常是日期列）
    date_col = df.columns[0]
    date_values = df[date_col].copy()
    
    # 將其餘所有列轉換為數值型，無法轉換的設為 NaN
    
    # 將日期列放回
    df[date_col] = date_values
    
    # 創建結果 DataFrame (4行: 2行數據, 2行摘要)
    result_df = pd.DataFrame(index=range(4), columns=df.columns)



    # 填充日期列的最後兩個值
    result_df.loc[0, date_col] = df[date_col].iloc[-2] if len(df) > 1 else ""
    result_df.loc[1, date_col] = df[date_col].iloc[-1] if len(df) > 0 else ""
    result_df.loc[2, date_col] = "趨勢比較"
    result_df.loc[3, date_col] = "最新值"
    
    # 定義一個函數來格式化數值，消除小數點後的零
    def format_number(num):
        if pd.isna(num):
            return np.nan
        
        # 如果是整數值（如 12.0），則轉換為整數（12）
        if float(num).is_integer():
            return int(float(num))
        
        # 否則保留小數
        return float(num)
    
    # 處理每一列（除了日期列）
    for col in df.columns[1:]:
        if col == date_col:
            continue
        
        # 獲取此列的非 NaN 值
        non_nan_values = df[col].dropna()
        
        if len(non_nan_values) >= 2:
            # 獲取最後兩個非 NaN 值
            penultimate_value = non_nan_values.iloc[-2]  # 倒數第二個值
            last_value = non_nan_values.iloc[-1]  # 最後一個值
            
            # 格式化數值（去除小數點後的零）
            penultimate_formatted = format_number(penultimate_value)
            last_formatted = format_number(last_value)
            
            # 添加到結果 DataFrame (數據行)
            result_df.loc[0, col] = penultimate_formatted
            result_df.loc[1, col] = last_formatted
            
            # 創建比較字串並添加到摘要行 (在相應的列)
            comparison = ">" 
            result_df.loc[2, col] = f"{col}: {penultimate_formatted}{comparison}{last_formatted}"
            result_df.loc[3, col] = f"{col}: {last_formatted}"
            
        elif len(non_nan_values) == 1:
            # 如果只有一個值，則兩行都填入相同的值
            single_value = non_nan_values.iloc[0]
            single_formatted = format_number(single_value)
            
            # 添加到結果 DataFrame (數據行)
            result_df.loc[0, col] = single_formatted
            result_df.loc[1, col] = single_formatted
            
            # 添加到摘要行 (在相應的列)
            result_df.loc[2, col] = f"{col}: {single_formatted}"
            result_df.loc[3, col] = f"{col}: {single_formatted}"
        else:
            # 如果沒有非 NaN 值，則所有行都填入 NaN
            result_df.loc[0, col] = np.nan
            result_df.loc[1, col] = np.nan
            result_df.loc[2, col] = np.nan
            result_df.loc[3, col] = np.nan
    



    return result_df

def process_lab_data(df):
    """
    從剪貼簿讀取實驗室數據，將數據轉換為數字格式（除了日期列），
    找出每個欄位最後兩個非 NaN 的值，並在相應欄位添加趨勢比較的摘要行。
    輸出時會刪除小數點後的零（例如 12.0 會顯示為 12）。
    處理">"或"<"符號，將其刪除並保留數字（例如>100會轉換為100）。
    
    Returns:
    DataFrame: 處理後的 DataFrame，包含每個欄位最後兩個非 NaN 值
              以及顯示值比較的附加摘要行（放在各自的列中）。
    """
    try:
        # 從剪貼簿讀取數據
        
        
        
        try:
            df['日期'] = pd.to_datetime(df['日期'], format='%y-%m-%d %H:%M', errors='coerce')
        except:
            df['日期'] = pd.to_datetime(df['日期'], format='%Y/%m/%d', errors='coerce')
        
        df['日期'] = df['日期'].dt.strftime('%Y/%m/%d')

    except Exception as e:
        print(f"讀取剪貼簿錯誤: {e}")
        return None
    has_pct_poct = 'PCT(POCT)' in df.columns
    has_procalcitonin = 'procalcitonin(PCT)' in df.columns
    
    # 如果兩個欄位都存在，則合併它們
    if has_pct_poct and has_procalcitonin:
        # 先將它們都轉換為數值型
        df['PCT(POCT)'] = pd.to_numeric(df['PCT(POCT)'], errors='coerce')
        df['procalcitonin(PCT)'] = pd.to_numeric(df['procalcitonin(PCT)'], errors='coerce')
        
        # 創建新的 PCT 欄位，優先使用非 NaN 的值
        df['PCT'] = df['procalcitonin(PCT)'].combine_first(df['PCT(POCT)'])
        
        # 刪除原始欄位
        df = df.drop(['PCT(POCT)', 'procalcitonin(PCT)'], axis=1)

    # 如果只有其中一個欄位存在，則重命名為 PCT
    elif has_pct_poct:
        df = df.rename(columns={'PCT(POCT)': 'PCT'})
    elif has_procalcitonin:
        df = df.rename(columns={'procalcitonin(PCT)': 'PCT'})

    # 使用字典映射簡化欄位重命名
    columns_mapping = {
        "procalcitonin(PCT)": "PCT",
        "CREA": "Cr",
        "NA": "Na",
        "ALB": "Alb",
        "BILIT": "Tbili",
        "DBILI": "Dbili",
        "NT-ProBNP": "BNP",
        "Amonia(NH3)": "NH3",
        "INR(PT)": "INR",
        "D-dimer": "Dd",
        "FIBRINOGEN": "fib",
        "lactate": "Lac",   
        "CRP2":"CRP",
        "eGFR(C-G)未BSA校正":"eGFR",
        "Glucose":"Glu",
        "FREET4":"fT4",
        "Free Ca++": "fCa",
        "CA": "Ca",
        
    }
    
    # 進行欄位重命名，只重命名存在於DataFrame中的欄位
    df = df.rename(columns={col: new_col for col, new_col in columns_mapping.items() if col in df.columns})
    
    # 保存第一列（通常是日期列）
    date_col = df.columns[0]

    # 處理 ">", "<" 符號並將其餘所有列轉換為數值型
    for col in df.columns[1:]:
        if col != date_col:  # 跳過日期列
            # 先將 "-" 替換為 NaN
            df[col] = df[col].mask(df[col] == '-', np.nan)
            
            # 處理包含 ">" 或 "<" 的值
            df[col] = df[col].apply(lambda x: remove_symbols(x))
            
            try:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            except:
                pass
    
    # 創建結果 DataFrame (4行: 2行數據, 2行摘要)
    result_df = pd.DataFrame(index=range(4), columns=df.columns)

    # 填充日期列的最後兩個值
    result_df.loc[0, date_col] = df[date_col].iloc[-2] if len(df) > 1 else ""
    result_df.loc[1, date_col] = df[date_col].iloc[-2] if len(df) > 1 else ""
    result_df.loc[2, date_col] = "趨勢比較"
    result_df.loc[3, date_col] = "最新值"

    def format_number(num):
        if pd.isna(num):
            return np.nan
        
        try:
            # 如果是整數值（如 12.0），則轉換為整數（12）
            if float(num).is_integer():
                return str(int(float(num)))
        except:
            pass
        # 否則保留小數
        return str(num)
    
    # 處理每一列（除了日期列）
    for col in df.columns:
        if col == date_col:
            continue
        
        # 獲取此列的非 NaN 值
        non_nan_values = df[col].dropna()
        
        if len(non_nan_values) >= 2:
            # 獲取最後兩個非 NaN 值
            penultimate_value = non_nan_values.iloc[-2]  # 倒數第二個值
            last_value = non_nan_values.iloc[-1]  # 最後一個值
            
            # 格式化數值（去除小數點後的零）
            penultimate_formatted = format_number(penultimate_value)
            last_formatted = format_number(last_value)
            
            # 添加到結果 DataFrame (數據行)
            result_df.loc[0, col] = penultimate_formatted
            result_df.loc[1, col] = last_formatted
            
            # 創建比較字串並添加到摘要行 (在相應的列)
            comparison = ">" 
            result_df.loc[2, col] = f"{col}: {penultimate_formatted}{comparison}{last_formatted}"
            result_df.loc[3, col] = f"{col}: {last_formatted}"
            
        elif len(non_nan_values) == 1:
            # 如果只有一個值，則兩行都填入相同的值
            single_value = non_nan_values.iloc[0]
            single_formatted = format_number(single_value)
            
            # 添加到結果 DataFrame (數據行)
            result_df.loc[0, col] = single_formatted
            result_df.loc[1, col] = single_formatted
            
            # 添加到摘要行 (在相應的列)
            result_df.loc[2, col] = f"{col}: {single_formatted}"
            result_df.loc[3, col] = f"{col}: {single_formatted}"
        else:
            # 如果沒有非 NaN 值，則所有行都填入 NaN
            result_df.loc[0, col] = np.nan
            result_df.loc[1, col] = np.nan
            result_df.loc[2, col] = np.nan
            result_df.loc[3, col] = np.nan

    return result_df

def add_pf_to_df(df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Adds a 'P/F' column to the DataFrame based on PO2 and FIO2 values.

    Args:
        df_input: The input DataFrame containing 'PO2' and 'FIO2' columns
                  for at least the first two rows (index 0 and 1).

    Returns:
        A new DataFrame with the 'P/F' column added.
    """
    original_df = df_input.copy() # Store a copy of the original DataFrame

    try:
        df = df_input.copy() # Create a copy to avoid modifying the original DataFrame

        # Ensure PO2 and FIO2 are numeric types
        df['PO2'] = pd.to_numeric(df['PO2'], errors='coerce')
        df['FIO2'] = pd.to_numeric(df['FIO2'], errors='coerce')
        print(df)
        # Restore original values for specific cells, assuming they might be non-numeric initially
        # and you want to preserve them for display if they were strings.
        # This part seems to contradict the `pd.to_numeric` part if the intention is to perform calculations.
        # If the intention is to just preserve the original string values for display in some cases,
        # and numeric for calculations, this needs careful handling.
        # For calculation, these should be numeric. If they were originally strings that couldn't be converted,
        # they would become NaN after `to_numeric(errors='coerce')`.
        # The lines below will overwrite the numeric conversion if original was not convertible and preserve original non-numeric value.
        # If the goal is to perform calculations, these lines should be removed or handled differently.
        # Assuming for now, these are intended to preserve original values where conversion might have failed,
        # but for calculation purposes, we rely on the `to_numeric` output.
        
        if df['FIO2'].iloc[2] is np.nan or df['FIO2'].iloc[3] is np.nan or df['PO2'].iloc[2] is np.nan or df['PO2'].iloc[3] is np.nan:
            return original_df
        
        df['FIO2'].iloc[2]=df_input['FIO2'].iloc[2]
        df['FIO2'].iloc[3]=df_input['FIO2'].iloc[3]
        df['PO2'].iloc[2]=df_input['PO2'].iloc[2]
        df['PO2'].iloc[3]=df_input['PO2'].iloc[3]


        # Calculate  P/F for row 0 and row 1
        # FIO2 is a percentage, so divide by 100 to convert to a decimal
        # Check if FIO2 is zero to avoid division by zero errors
        if df.loc[0, 'FIO2'] != 0:
            df.loc[0, 'P/F'] = df.loc[0, 'PO2'] / (df.loc[0, 'FIO2'] / 100)
        else:
            df.loc[0, 'P/F'] = np.nan # If FIO2 is zero, set to NaN

        if df.loc[1, 'FIO2'] != 0:
            df.loc[1, 'P/F'] = df.loc[1, 'PO2'] / (df.loc[1, 'FIO2'] / 100)
        else:
            df.loc[1, 'P/F'] = np.nan # If FIO2 is zero, set to NaN

        # Get P/F values for row 0 and row 1
        pf_ratio_row0 = df.loc[0, 'P/F']
        pf_ratio_row1 = df.loc[1, 'P/F']

        # Determine the string for row 2 (index 2 in the DataFrame)
        if pd.isna(pf_ratio_row0) or pd.isna(pf_ratio_row1):
            df.loc[2, 'P/F'] = np.nan
        elif pf_ratio_row0 == pf_ratio_row1:
            df.loc[2, 'P/F'] = f"P/F: {pf_ratio_row0:.0f}"
        else:
            df.loc[2, 'P/F'] = f"P/F: {pf_ratio_row0:.0f}>{pf_ratio_row1:.0f}"
       
        # Determine the string for row 3 (index 3 in the DataFrame)
        if pd.isna(pf_ratio_row1):
            df.loc[3, 'P/F'] = np.nan
        else:
            df.loc[3, 'P/F'] = f"P/F: {pf_ratio_row1:.0f}"

        return df

    except Exception as e:
        print(f"An error occurred: {e}")
        return original_df # Return the original DataFrame in case of an error

def process_glucose_data(df, num=3, include_dates=False):
    """
    Reads glucose data from the clipboard (assumed to be tab-separated)
    and processes it. Now, it groups values by date, using '>' for same-day
    values and ',' to separate different days.

    Args:
        num (int): The index of the row from which to start extracting data
                   (inclusive), going backwards to row 0.
        include_dates (bool): If True, the output string will also include
                              the formatted date (YYYY/MM/DD) for each entry.

    Returns:
        str or None: A formatted string of glucose values (e.g., "Glu: 225>169,Lo>Hi")
                     or "Glu: 2025/06/24-170>2025/06/24-204,2025/06/23-254>..." if include_dates is True.
                     Returns None if no valid data is found or an error occurs.
    """
    
    if df.shape[0] <= 2:
        print("The DataFrame has 2 or fewer rows.")
        return None

    if 'Glucose' not in df.columns or '日期' not in df.columns:
        missing_cols = []
        if 'Glucose' not in df.columns:
            missing_cols.append("'Glucose'")
        if '日期' not in df.columns:
            missing_cols.append("'日期'")
        print(f"錯誤：找不到必要的欄位 {', '.join(missing_cols)}。目前欄位: {df.columns.tolist()}")
        return None

    # 移除最後一列如果包含特定關鍵字
    if len(df) > 0 and ('累積報告' in str(df.iloc[-1].values) or '趨勢圖' in str(df.iloc[-1].values)):
        df = df.drop(df.index[-1])
        # 重置索引以防萬一，確保後續 for 迴圈的索引正確
        df = df.reset_index(drop=True) 

    # 將日期轉換為 datetime 物件，錯誤的日期會變為 NaT
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    
    # 格式化日期字串，無法轉換的日期(NaT)會變成 NaT 字串
    df['日期_日only'] = df['日期'].dt.strftime('%Y/%m/%d') # 用於比較日期是否改變
    df['日期_完整格式'] = df['日期'].dt.strftime('%Y/%m/%d') # 用於 include_dates 顯示

    # 用來存放每個日期的血糖值列表，最終會是一個列表的列表
    # 例如: [['170', '204'], ['254', '248', '216']]
    grouped_entries = [] 
    current_day_entries = [] # 存放當前一天的血糖值
    last_date = None # 用於追蹤上一個處理的日期

    # 確保 num 不會超過 DataFrame 的實際行數
    start_index = min(num, len(df) - 1)

    # 從指定索引開始倒序遍歷
    for i in range(start_index, -1, -1):
        try:
            glucose_val = df.loc[i, 'Glucose']
            current_date = df.loc[i, '日期_日only'] # 獲取不含時間的日期字串用於分組
            full_date_str = df.loc[i, '日期_完整格式'] # 獲取完整格式日期用於顯示

            # 跳過血糖值為 NaN 或日期為 NaT 的行 (NaT 的字串形式也是 'NaT')
            # 同時檢查 full_date_str 是否為 'NaT'，確保日期有效
            if pd.isna(glucose_val) or pd.isna(df.loc[i, '日期']): 
                continue

            # 判斷日期是否改變
            if last_date is not None and current_date != last_date:
                # 如果日期改變了，且 current_day_entries 不為空，則將其加入 grouped_entries
                if current_day_entries:
                    grouped_entries.append(current_day_entries)
                current_day_entries = [] # 開始新的日期分組
            
            last_date = current_date # 更新上一個日期

            processed_glucose_val = str(glucose_val).strip() 

            if processed_glucose_val == 'RR Lo':
                processed_glucose_val = 'Lo'
            elif processed_glucose_val == 'RR Hi':
                processed_glucose_val = 'Hi'
            else:
                numeric_val = pd.to_numeric(glucose_val, errors='coerce')
                if pd.isna(numeric_val):
                    continue 
                else:
                    processed_glucose_val = str(int(numeric_val))

            # 組合日期和血糖值，如果需要
            if include_dates:
                # 確保日期不是 'NaT' 字串
                if not pd.isna(df.loc[i, '日期']): 
                     current_day_entries.append(f"{full_date_str}-{processed_glucose_val}")
                else:
                     continue # 日期無效則跳過
            else:
                current_day_entries.append(processed_glucose_val)
                
        except (KeyError, ValueError, TypeError) as e:
            continue

    # 將最後一個日期的條目加入 grouped_entries (如果存在)
    if current_day_entries:
        grouped_entries.append(current_day_entries)

    if not grouped_entries:
        return None 

    # 將內部列表用 '>' 連接，再將外部列表用 ',' 連接
    final_output_parts = []
    for day_values in grouped_entries:
        final_output_parts.append(">".join(day_values))
    
    return "Glu: " + ", ".join(final_output_parts)

def remove_symbols(value):
    """
    從字符串中移除 ">" 或 "<" 符號並保留數字部分。
    
    Args:
        value: 輸入值，可能是字符串、數字或 NaN
        
    Returns:
        處理後的值，如果原值包含 ">" 或 "<"，則返回不帶這些符號的數字部分
    """
    if pd.isna(value):
        return value
    
    # 將值轉換為字符串以便處理
    value_str = str(value)
    
    # 檢查是否包含 ">" 或 "<" 符號
    if value_str.startswith('>') or value_str.startswith('<'):
        # 刪除符號，只保留數字部分
        return value_str[1:].strip()
    
    return value

def check_threshold_condition(df, column, threshold_low=10, threshold_high=30, output=True):
    """
    檢查實驗室數據中特定欄位的最新值是否介於指定閾值之間，
    並根據條件決定返回哪一行的摘要文本。

    參數:
    df (DataFrame): 包含實驗室數據的 DataFrame，預期至少有四行。
    column (str): 要檢查的欄位名。
    threshold_low (float): 閾值下限，預設為 10。
    threshold_high (float): 閾值上限，預設為 30。
    output (bool): 是否啟用結果輸出，預設為 True。

    返回:
    str or None: 根據條件返回對應的摘要文本，若不滿足條件則返回 None。
    """
    try:
        # 從第二行（索引為 1）獲取最新值
        latest_value_str = df.iloc[1][column]

        # 嘗試將數據轉換為浮點數，處理可能包含非數字字符的情況
        try:
            latest_value = float(latest_value_str)
        except (ValueError, TypeError):
            # 如果轉換失敗，嘗試從字符串中提取數字
            match = re.search(r'(\d+\.?\d*)', str(latest_value_str))
            if match:
                latest_value = float(match.group(1))
            else:
                return None  # 如果無法提取數字，則返回 None

        # 根據新的閾值條件判斷
        condition_met = 0

        # 判斷是否滿足主要條件：output為True且數值介於兩閾值之間
        if output and threshold_low < latest_value < threshold_high:
            condition_met = 1
        # 判斷是否滿足副條件：output為False且數值不在兩閾值之間
        elif not output and (latest_value <= threshold_low or latest_value >= threshold_high):
            condition_met = 1
        elif output and (latest_value <= threshold_low or latest_value >= threshold_high):
            condition_met = 2
        else:
            condition_met = -1

        # 根據 condition_met 的值返回對應的摘要文本
        if condition_met == 2:
            return df.iloc[2][column]
        elif condition_met == 1:
            return df.iloc[3][column]
        elif condition_met == 0:
            return df.iloc[3][column]
        elif condition_met == -1:
            return None

    except Exception as e:
        print(f"處理 {column} 欄位時發生錯誤: {e}")
        return None

def apply_conditions_to_dataframe(df, conditions_dict):
    """
    對 DataFrame 應用多個閾值條件
    
    參數:
    df (DataFrame): 包含實驗室數據的 DataFrame
    conditions_dict (dict): 閾值條件字典，格式為 {'column': (threshold, condition)}
                          例如 {'lactate': (20, '<'), 'procalcitonin': (0.5, '>')}
    
    返回:
    DataFrame: 應用所有條件後的 DataFrame
    """
    result = df.copy()
    a = []
    for column, (threshold, condition, output) in conditions_dict.items():
        a.append(check_threshold_condition(result, column, threshold, condition, output))
    
    return a

def extract_height_weight_trends_from_clipboard(df, height=False, weight=True, bsa=True, ibw=False, gender="男"):
    """
    整理DataFrame，輸出格式化字符串
    
    Parameters:
    df: pandas DataFrame，包含日期時間、身高、體重、BSA等欄位
    height: bool，是否處理並輸出身高資料 (預設True)
    weight: bool，是否處理並輸出體重資料 (預設True)
    bsa: bool，是否處理並輸出BSA資料 (預設True)
    ibw: bool，是否計算並輸出理想體重 (預設True)
    gender: str，性別 ("Male" 或 "Female")，用於IBW計算 (預設"Male")
    
    Returns:
    str: 格式化的字符串，包含身高、體重、BSA、IBW和日期資訊
    """

    def extract_height(height_str):
        """從身高字符串中提取數值（cm）"""
        if pd.isna(height_str):
            return np.nan
        height_str = str(height_str)
        if 'cm' in height_str:
            try:
                return float(height_str.replace('cm', ''))
            except:
                return np.nan
        try:
            return float(height_str)
        except:
            return np.nan

    def extract_weight(weight_str):
        """從體重字符串中提取數值（kg）"""
        if pd.isna(weight_str):
            return np.nan
        weight_str = str(weight_str)
        if 'kg' in weight_str:
            try:
                return float(weight_str.replace('kg', ''))
            except:
                return np.nan
        try:
            return float(weight_str)
        except:
            return np.nan

    def extract_bsa(bsa_str):
        """從BSA字符串中提取數值（m2）"""
        if pd.isna(bsa_str):
            return np.nan
        bsa_str = str(bsa_str)
        if 'm2' in bsa_str and bsa_str != 'm2':
            try:
                return float(bsa_str.replace('m2', ''))
            except:
                return np.nan
        try:
            return float(bsa_str)
        except:
            return np.nan

    def calculate_bsa(height_cm, weight_kg):
        """使用Du Bois公式計算BSA"""
        if pd.isna(height_cm) or pd.isna(weight_kg):
            return np.nan
        return math.sqrt((height_cm * weight_kg) / 3600)

    def calculate_ibw(height_cm, gender="Male"):
        """計算理想體重 (IBW)
        男性: IBW = 50 + 2.3 × (身高英寸 - 60)
        女性: IBW = 45.5 + 2.3 × (身高英寸 - 60)
        """
        if pd.isna(height_cm):
            return np.nan
        height_inches = height_cm / 2.54
        if height_inches <= 60:
            return 50 if gender.lower() == "男" else 45.5
        else:
            base_weight = 50 if gender.lower() == "男" else 45.5
            return base_weight + 2.3 * (height_inches - 60)
    
    
    # 複製DataFrame避免修改原始數據
    df_work = df.copy()
    
    # 獲取欄位名稱（支援中英文）
    date_col = None
    height_col = None
    weight_col = None
    bsa_col = None
    
    for col in df_work.columns:
        col_lower = str(col).lower()
        if '日期' in col or '時間' in col or 'date' in col_lower or 'time' in col_lower:
            date_col = col
        elif '身高' in col or 'height' in col_lower:
            height_col = col
        elif '體重' in col or 'weight' in col_lower:
            weight_col = col
        elif 'bsa' in col_lower:
            bsa_col = col
    
    # 處理日期時間
    if date_col is not None:
        try:
            df_work[date_col] = pd.to_datetime(df_work[date_col])
            date_list = df_work[date_col].tolist()
        except:
            date_list = df_work[date_col].tolist()
    else:
        date_list = []
    
    # 處理身高 (即使height=False也要處理，因為IBW和BSA需要身高資料)
    height_list = []
    if height_col is not None:
        df_work['height_clean'] = df_work[height_col].apply(extract_height)
        # 處理同日期重複資料：只保留最新的一筆非None資料
        if date_col is not None:
            df_work['date_only'] = df_work[date_col].dt.date
            height_processed = []
            date_height_map = {}  # 記錄每個日期的最新非None資料
            
            # 從最後一筆開始往前檢查（因為後面的資料較新）
            for i in reversed(range(len(df_work))):
                date_val = df_work.iloc[i]['date_only']
                height_val = df_work.iloc[i]['height_clean']
                
                # 如果這個日期還沒有記錄過，或者當前值不是None且之前記錄的是None
                if date_val not in date_height_map:
                    date_height_map[date_val] = height_val
                elif (height_val is not None and not pd.isna(height_val) and 
                      (date_height_map[date_val] is None or pd.isna(date_height_map[date_val]))):
                    date_height_map[date_val] = height_val
            
            # 重新遍歷，分配最終值
            for i in range(len(df_work)):
                date_val = df_work.iloc[i]['date_only']
                height_val = df_work.iloc[i]['height_clean']
                
                # 如果當前值等於該日期的最佳值，保留；否則設為None
                if height_val == date_height_map[date_val]:
                    height_processed.append(height_val)
                else:
                    height_processed.append(None)
            
            height_list = height_processed
        else:
            height_list = df_work['height_clean'].tolist()
    
    # 處理體重 (即使weight=False也要處理，因為BSA需要體重資料)
    weight_list = []
    if weight_col is not None:
        df_work['weight_clean'] = df_work[weight_col].apply(extract_weight)
        # 處理同日期重複資料：只保留最新的一筆非None資料
        if date_col is not None:
            if 'date_only' not in df_work.columns:
                df_work['date_only'] = df_work[date_col].dt.date
            weight_processed = []
            date_weight_map = {}  # 記錄每個日期的最新非None資料
            
            # 從最後一筆開始往前檢查（因為後面的資料較新）
            for i in reversed(range(len(df_work))):
                date_val = df_work.iloc[i]['date_only']
                weight_val = df_work.iloc[i]['weight_clean']
                
                # 如果這個日期還沒有記錄過，或者當前值不是None且之前記錄的是None
                if date_val not in date_weight_map:
                    date_weight_map[date_val] = weight_val
                elif (weight_val is not None and not pd.isna(weight_val) and 
                      (date_weight_map[date_val] is None or pd.isna(date_weight_map[date_val]))):
                    date_weight_map[date_val] = weight_val
            
            # 重新遍歷，分配最終值
            for i in range(len(df_work)):
                date_val = df_work.iloc[i]['date_only']
                weight_val = df_work.iloc[i]['weight_clean']
                
                # 如果當前值等於該日期的最佳值，保留；否則設為None
                if weight_val == date_weight_map[date_val]:
                    weight_processed.append(weight_val)
                else:
                    weight_processed.append(None)
            
            weight_list = weight_processed
        else:
            weight_list = df_work['weight_clean'].tolist()
    
    # 準備輸出字符串的組件
    result_parts = []
    

    
    # 1. Weight result: 體重變化字符串 (只有在weight=True時顯示)
    if weight:
        weight_clean = [w for w in weight_list if w is not None and not pd.isna(w)]
        if len(weight_clean) >= 3:
            weight_selected = [weight_clean[0], weight_clean[1], weight_clean[-1]]
            weight_result = f"BW: {weight_selected[2]} >> {weight_selected[1]} > {weight_selected[0]}"
        elif len(weight_clean) == 2:
            weight_result = f"BW: {weight_clean[1]} > {weight_clean[0]}"
        elif len(weight_clean) == 1:
            weight_result = f"BW: {weight_clean[0]}"
        else:
            weight_result = None
        
        if weight_result:
            result_parts.append(weight_result)

    # 2. Height: 第一個有效身高值 (只有在height=True時顯示)
    if height:
        height_clean = [h for h in height_list if h is not None and not pd.isna(h)]
        if height_clean:
            result_parts.append(f"Height: {height_clean[0]}")
    
    # 3. BSA: 使用第一個身高和最新體重計算 (只有在bsa=True時顯示，且能計算出有效值)
    if bsa:
        height_clean = [h for h in height_list if h is not None and not pd.isna(h)]
        weight_clean = [w for w in weight_list if w is not None and not pd.isna(w)]
        
        if height_clean and weight_clean:
            # 使用第一個有效身高和最新體重
            calculated_bsa = calculate_bsa(height_clean[0], weight_clean[-1])
            if not pd.isna(calculated_bsa) and calculated_bsa is not None:
                result_parts.append(f"BSA: {calculated_bsa:.2f}")
    
    # 4. IBW: 理想體重 (只有在ibw=True時顯示，且能計算出有效值)
    if ibw:
        height_clean = [h for h in height_list if h is not None and not pd.isna(h)]
        if height_clean:
            calculated_ibw = calculate_ibw(height_clean[0], gender)
            if not pd.isna(calculated_ibw) and calculated_ibw is not None:
                result_parts.append(f"IBW: {calculated_ibw:.1f}")
    
    # 5. Date: 第一筆日期的月/日格式 (放在最後，不用逗號分隔)
    date_part = None
    if date_list and not pd.isna(date_list[0]):
        first_date = pd.to_datetime(date_list[0])
        date_part = f"({first_date.month}/{first_date.day})"
    
    # 組合成最終字符串
    if result_parts:
        if date_part:
            return ", ".join(result_parts) + " " + date_part
        else:
            return ", ".join(result_parts)
    else:
        return date_part if date_part else ""

def process_lab_report():
    """
    從剪貼簿讀取實驗室報告文本，將其轉換為三個 DataFrame：
    1. 敏感性累積報告
    2. 一年內陽性培養結果累積報告
    3. 最近三個月使用抗生素列表
    
    函數假設數據已經複製到剪貼簿中
    
    Returns:
        tuple: 包含三個 DataFrame 的元組
    """
    try:
        # 從剪貼簿讀取文本
        text = pyperclip.paste()
        if not text.strip():
            print("錯誤：剪貼簿為空")
            return None, None, None
            
        # 分割成三個部分的數據
        parts = text.split("一年內陽性培養結果累積報告")
        if len(parts) < 2:
            print("錯誤：找不到'一年內陽性培養結果累積報告'標記")
            return None, None, None
            
        sensitivity_part = parts[0].strip()
        
        remaining = parts[1].strip()
        parts2 = remaining.split("最近三個月使用抗生素列表")
        
        if len(parts2) < 2:
            print("錯誤：找不到'最近三個月使用抗生素列表'標記")
            culture_part = remaining
            antibiotics_part = ""
        else:
            culture_part = parts2[0].strip()
            antibiotics_part = parts2[1].strip()
        
        # 處理敏感性累積報告
        sensitivity_df = None
        if sensitivity_part:
            # 去掉第一行 "敏感性累積報告"
            sensitivity_lines = sensitivity_part.split('\n')
            if len(sensitivity_lines) > 1:
                sensitivity_data = '\n'.join(sensitivity_lines[1:])
                sensitivity_df = pd.read_csv(io.StringIO(sensitivity_data), delimiter='\t')
                # 清理列名（去除前後空格）
                sensitivity_df.columns = sensitivity_df.columns.str.strip()
        
        # 處理一年內陽性培養結果累積報告
        culture_df = None
        if culture_part:
            # 找到表頭行（以 "-" 開頭的行）
            culture_lines = culture_part.split('\n')
            header_line_index = -1
            
            for i, line in enumerate(culture_lines):
                if line.strip().startswith("-"):
                    header_line_index = i
                    break
            
            if header_line_index >= 0 and len(culture_lines) > header_line_index + 1:
                # 提取表頭和數據
                header_line = culture_lines[header_line_index].strip()
                data_lines = culture_lines[header_line_index:]
                
                # 處理表頭
                headers = [h.strip() for h in header_line.split('\t')]
                if headers[0] == "-":
                    headers[0] = "序號"
                
                # 處理數據
                data_text = '\n'.join(data_lines)
                culture_df = pd.read_csv(io.StringIO(data_text), delimiter='\t', index_col=False)
                
                # 如果列數與表頭數不匹配，嘗試調整
                if len(culture_df.columns) == len(headers):
                    culture_df.columns = headers
                else:
                    min_len = min(len(headers), len(culture_df.columns))
                    culture_df.columns = headers[:min_len]
        
        # 處理最近三個月使用抗生素列表
        antibiotics_df = None
        if antibiotics_part:
            # 特殊處理抗生素部分，手動解析以確保第一列不丟失
            antibiotics_lines = antibiotics_part.split('\n')
            
            if len(antibiotics_lines) > 1:
                # 獲取表頭
                header_line = antibiotics_lines[0].strip()
                headers = [h.strip() for h in header_line.split('\t')]
                
                # 手動構建數據列表
                data_rows = []
                for i in range(1, len(antibiotics_lines)):
                    line = antibiotics_lines[i].strip()
                    if not line:  # 跳過空行
                        continue
                        
                    # 分割行並確保每個單元格都被保留
                    cells = line.split('\t')
                    
                    # 確保不會丟失第一個單元格的數據
                    if len(cells) > 0 and cells[0].strip().startswith(' '):
                        cells[0] = cells[0].strip()
                    
                    # 添加到數據列表
                    data_rows.append(cells)
                
                # 創建DataFrame
                antibiotics_df = pd.DataFrame(data_rows)
                
                # 應用標題
                if len(antibiotics_df.columns) <= len(headers):
                    antibiotics_df.columns = headers[:len(antibiotics_df.columns)]
                else:
                    # 如果數據列比標題多
                    extended_headers = headers + [f"未命名列{i}" for i in range(len(headers), len(antibiotics_df.columns))]
                    antibiotics_df.columns = extended_headers
        
        return sensitivity_df, culture_df, antibiotics_df
        
    except Exception as e:
        print(f"處理數據時出錯: {str(e)}")
        return None, None, None

def get_recent_culture_results_string(df, days=30, limit=3):
    """
    Extract recent culture results from a DataFrame.
    Returns a formatted string with date, location, and organism information.
    Keeps only the newest entry when the same location and organism appear multiple times.
    
    Parameters:
    -----------
    df : pandas.DataFrame
        DataFrame containing culture results with columns for date, specimen, and organism
    days : int, default=30
        Number of days to look back from today
    limit : int, default=3
        Maximum number of most recent culture results to return
        
    Returns:
    --------
    str
        Formatted string with the most recent culture results
    """
    # Get today's date for filtering
    today = datetime.now()
    cutoff_date = today - timedelta(days=days)
    
    # Initialize results list
    cultures = []
    
    # Identify the relevant columns - assuming standard column structure
    date_col = '簽收日期時間'
    location_col = '檢體'
    
    # Add organism column to the list (assuming it's the next column after location)
    organism_col = df.columns[list(df.columns).index(location_col) + 1] if location_col in df.columns else None
    
    # If we don't have the necessary columns, return empty string
    if date_col not in df.columns or location_col not in df.columns or organism_col is None:
        return ""
    
    # Track unique combinations of location and organism
    location_organism_dict = {}
    
    # Process each row
    for _, row in df.iterrows():
        # Get values
        date_val = row[date_col]
        location_val = row[location_col]
        organism_val = row[organism_col] if organism_col in row.index else None
        
        # Skip rows with missing data
        if pd.isna(date_val) or pd.isna(organism_val):
            continue
        
        # Convert date to datetime
        try:
            date_obj = pd.to_datetime(date_val)
        except:
            continue
            
        # Filter out dates older than the cutoff
        if date_obj < cutoff_date:
            continue
            
        # Format date as MM/DD
        formatted_date = f"{date_obj.month}/{date_obj.day}"
        
        # Simplify location
        location_str = str(location_val).strip()
        if '(' in location_str:
            simplified_location = location_str.split('(')[0].strip()
        else:
            simplified_location = location_str.split('.')[0] if '.' in location_str else location_str
        
        # Clean organism name
        organism_str = str(organism_val).strip() if organism_val else "Unknown"
        
        # Abbreviate specific organisms using a dictionary for easy expansion
        organism_abbreviations = {
            "Escherichia coli": "E. coli",
            "Stenotrophomonas maltophilia": "S. maltophilia",
            "Enterococcus faecium (VRE)": "E. faecium (VRE)",
            "Acinetobacter baumannii (CRAB)": "A. baumannii (CRAB)",
            "Acinetobacter baumannii": "A. baumannii",
            "Staphylococcus aureus": "S. aureus", 
            "Staphylococcus aureus (MRSA)": "S. aureus (MRSA)", 
            "Enterococcus faecium": "E. faecium", 
            "Enterococcus faecalis": "E. faecalis",
            "Pseudomonas aeruginosa": "Ps. aeruginosa",
            "Staphylococcus aureus ssp aureus": "S. aureus",
            "Klebsiella pneumoniae": "K. pneumonia",
            "Klebsiella pneumoniae ssp pneumoniae": "K. pneumonia",
            "Klebsiella oxytoca": "K. oxytoca",
            "Klebsiella aerogenes": "K. aerogenes",
            # 可以在此處輕鬆添加更多細菌的縮寫
            # "Full Name": "Abbreviated Name",
        }
        
        # Check if current organism is in our abbreviation dictionary
        if organism_str in organism_abbreviations:
            organism_str = organism_abbreviations[organism_str]
        
        # Create a unique key for this location+organism combination
        key = (simplified_location, organism_str)
        
        # Check if we already have this combination
        if key in location_organism_dict:
            # Compare dates - keep only the newer one
            if date_obj > location_organism_dict[key]['date_obj']:
                # Replace with newer entry
                location_organism_dict[key] = {
                    'date_obj': date_obj,
                    'formatted_entry': f"{formatted_date} {simplified_location}: {organism_str}"
                }
        else:
            # Add new entry to dictionary
            location_organism_dict[key] = {
                'date_obj': date_obj,
                'formatted_entry': f"{formatted_date} {simplified_location}: {organism_str}"
            }
    
    # Extract values from dictionary to get unique entries
    cultures = list(location_organism_dict.values())
    
    # Sort by date (newest first) and create result string
    cultures.sort(key=lambda x: x['date_obj'], reverse=True)
    
    # Limit the number of results according to the limit parameter
    limited_cultures = cultures[:limit]
    
    result_string = ", ".join(culture['formatted_entry'] for culture in limited_cultures)
    if result_string == 'nan/nan : Unknown':
        result_string = None
    return result_string

def get_active_antibiotics(df: pd.DataFrame, anti_D = True) -> str:
    """
    從藥物DataFrame中找出正在使用(IN-USE)的抗生素，並以首字母+開始日期月/日的格式返回

    參數:
    df (pandas.DataFrame): 包含藥物資訊的DataFrame，需要有「藥名」、「開始日」和「狀態」欄位

    回傳:
    str: 使用中抗生素的縮寫列表，格式為"首字母月/日"，以逗號分隔
    """
    # 定義藥物名稱的映射關係表，方便日後擴充
    for i in range(1, len(df)):
        if df.loc[i, '藥名'] == '--':
            df.loc[i, '藥名'] = df.loc[i-1, '藥名']

    drug_name_mapping = {
        'teicoplanin': 'Teicoplanin', # Added based on your image example for consistency if needed
        'cefoperazone': 'Cefoperazone', # Added based on your image example
        'piperacillin': 'Tazocin',
        'liposomal': 'AmphotericinB',
        # 未來可在此處添加更多映射關係:
        # 'generic_name': 'brand_name',
    }

    active_antibiotics = []

    # Ensure '開始日' is in datetime format for easy extraction of month and day
    # It's better to convert once for the entire column than row by row if possible
    df['開始日'] = pd.to_datetime(df['開始日'])
    # Remove rows where '藥名' is in the exclusion list
    exclusion_list = ['nystatin', 'entecavir','tenofovir alafenamide', 'letermovir']
    df = df[~df['藥名'].str.lower().isin([name.lower() for name in exclusion_list])]
    # Iterate over DataFrame rows. Using iterrows is generally preferred when
    # you need both index and row content, and when applying row-wise logic.
    today = datetime.now()
    for index, row in df.iterrows():
        medicine_name = str(row['藥名']).strip() # Ensure it's a string and remove leading/trailing whitespace
        status = str(row['狀態']).strip()       # Ensure it's a string and remove leading/trailing whitespace
        start_date = row['開始日']
        print(start_date,today)
        # Check if status is 'IN-USE'
        if status == 'IN-USE':
            # Get the first word of the medicine name
            first_word_of_drug = medicine_name.split(' ')[0].lower()

            # Apply mapping if the first word is in the mapping dictionary
            # If not in mapping, use the capitalized first word as is
            mapped_name = drug_name_mapping.get(first_word_of_drug, medicine_name.split(' ')[0])

            # Format the start date as Month/Day
            month = start_date.strftime('%m').lstrip('0') # Remove leading zero for month
            day = start_date.strftime('%d').lstrip('0')   # Remove leading zero for day
            #start_date = datetime.strptime(start_date, '%Y-%m-%d').date()
            
            #print(today)
            #print(start_date)
            # Create the formatted string
            # Note: Your example output for the similar def had "first_letter month/day-"
            # I'll follow that structure.
            abbr = f"{mapped_name} {month}/{day}-"
            if anti_D:
                days = int((today - start_date).days)+1
                abbr = f"{mapped_name} D{days}"
            active_antibiotics.append(abbr)

    # Join the results with a comma and space, as in your example def's return
    return ", ".join(active_antibiotics)

def stop_the_code():
    global stop_requested, driver_instance # Declare driver_instance as global
    if stop_requested:
        print("停止請求已發出，正在中止程式運行。")
        safe_exit(driver_instance) # Pass the global driver_instance to safe_exit
        # After safe_exit, if you use sys.exit(), the program will terminate.
        # Otherwise, the calling function should handle returning/breaking.
        raise SystemExit("Program terminated by user request.") # Raise an exception to stop execution flow
        
def extract_active_meds(df_meds): # 函數現在接受一個 DataFrame 輸入
    """
    Extract active medications from an input DataFrame and format them according to specifications.
    
    The function expects the DataFrame to contain columns like '學名', '劑量', '頻率',
    and a column that indicates "使用中" status (e.g., a status column or by checking other columns).
    
    Args:
        df_meds (pd.DataFrame): DataFrame containing medication data.
                                Expected columns: '學名', '劑量', '頻率', 
                                and a column for status (or '使用中' checkable in a text column).
                                For simplification, we assume the input df already has these named columns
                                or we can infer them from column index.
    
    Returns:
        str: Formatted string of active medications.
             Returns an error string if processing fails.
    """
    
    # Define text replacement rules
    text_replacements = {
        "Sod bicarbonate": "Bicarbonate", "Taita No.5": "Taita5", "Taita No.3": "Taita3",
        "Cal chloride 2%/10%": "Vitacal", "Insulin glargine": "Toujeo",
        "Insulin glulisine": "Apidra", "Insulin degludec": "Tresiba",
        "Insulin lispro Mix25": "Humalog-mix25", "Ipratropium/Salbutamol": "Combivent",
        "ACETYLCYSTEINE": "Acetylcystein", "Ursodeoxycholic": "Urso", "Insulin R-HM": "R.I. ",
        "Insulin aspart": "Novorapid", "Insulin detemir": "Levemir", "Vitamin C": "Vit.C",
        "Levetiracetam": "Keppra", "Febu tab 80 mg": "Febuxostat", "Ipratropium nebul.": "Atrovent",
        "Clopidogrel": "Plavix", "Dimethicone (Gasmin)": "Gasmin", "Cal. polystyrene": "Kalimate",
        "Potassium Chloride ER tab 750 mg": "Const-K", "Cal. acetate": "Cal.acetate",
        "Pot. gluconate soln 20 mEq/15 ml": "K-glu", "Sennosides FC tab 20 mg": "Through",
        "KCl 20 mEq in NS *inj 100ml": "KCl20/ns ", "KCl 20 mEq in D5W": "KCl20/d5w ",
        "Bisoprolol FC * tab 1.25 mg": "Concor-1.25", "BISOPROLOL FC * tab 5 MG": "Concor-5",
        "PROPRANOLOL * TAB 40 MG VPP": "Propranolol-40", "Propranolol * tab 10 mg VPP": "Propranolol-10",
        "Ezetimibe+Atorvastatin tab 10/20": "Atozet-10/20", "Valsartan * tab 80 mg": "Valsartan-80",
        'Valsartan "Novart" FC *tab 160mg': "Valsartan-160", "Rosuvastatin FC * tab 10 mg": "Rosuvastatin-10",
        "ATORVASTATIN FC * TAB 20 MG": "Atorvastatin-20", "Pitavastatin FC * tab 2 mg": "Pitavastatin-2",
        "Pitavastatin * tab 4 mg": "Pitavastatin-4", "Atorvastatin FC * tab 40 mg": "Atorvastatin-40", 
        'Fentanyl "PPCD"patch#>*12.5ug/hr': 'Fentanyl-12.5', 'Fentanyl "PPCD" patch#>* 25ug/hr': 'Fentanyl-25',
        'Fentanyl "PPCD" patch#>* 50ug/hr': 'Fentanyl-50', 'KCl 10 mEq in D5-1/3S': 'KCl10/d51/3s ', 
        'Tacrolimus * cap 1 mg': 'Tacrolimus-1', 'TACROLIMUS * CAP 1 MG': 'Tacrolimus-1',
        'Tacrolimus PR * cap 1 mg': 'Tacrolimus-1', 'Tacrolimus * cap 0.5 mg': 'Tacrolimus-0.5',
        'Ciclosporine neoral * cap 25 mg': 'Ciclosporine-25', 'CICLOSPORINE NEORAL * CAP 100 MG': 'Ciclosporine-100',
        'Zinc oxide': 'ZnO'
    }
    
    def apply_text_replacements(text):
        """Apply text replacements to the given text"""
        # 確保輸入是字串，避免對非字串類型操作報錯
        if not isinstance(text, str):
            return str(text) 
        for search_text, replacement_text in text_replacements.items():
            if search_text.lower() in text.lower():
                text = text.replace(search_text, replacement_text)
        return text

    try:
        active_meds_list = []
        
        # --- 假設 DataFrame 的欄位名稱或索引 ---
        # 為了簡化，我假設 DataFrame 有明確的欄位名稱。
        # 如果你的 DataFrame 只是從剪貼簿讀取且沒有明確的標頭，
        # 你可能需要像這樣根據索引來取值 (例如 df.iloc[:, 0] 是第一列):
        # col_generic_name = df_meds.iloc[:, 0] 
        # col_dosage = df_meds.iloc[:, 2]
        # col_frequency = df_meds.iloc[:, 5]
        # col_status_check = df_meds.iloc[:, X] # 假設有一個欄位能判斷 '使用中'
        
        # 更推薦的方法是，在呼叫 extract_active_meds 之前，
        # 先將剪貼簿數據處理成帶有明確標頭的 DataFrame。
        # 這裡我假設你傳入的 DataFrame 已經有這些列，
        # 或者其第一個字元包含 '學名'、'劑量'、'頻率' 的含義。
        
        # 為了兼容性，我們先嘗試用常見的欄位名，如果沒有，則用索引
        # 你可能需要根據實際數據調整這些欄位名或索引
        generic_name_col_name = None
        dosage_col_name = None
        frequency_col_name = None
        
        # 嘗試尋找對應的欄位名稱
        for col in df_meds.columns:
            if '學名' in col:
                generic_name_col_name = col
            elif '劑量' in col or 'Dose' in col:
                dosage_col_name = col
            elif '頻率' in col or 'Freq' in col:
                frequency_col_name = col
        
        # 如果沒有找到明確的欄位名稱，就使用索引（這需要你非常清楚數據的結構）
        # 這裡假設你的原始數據在剪貼簿中是 Tab 分隔的，並且列的順序固定
        if generic_name_col_name is None and len(df_meds.columns) > 0:
            generic_name_col_name = df_meds.columns[0] # 第一列是學名
        if dosage_col_name is None and len(df_meds.columns) > 2:
            dosage_col_name = df_meds.columns[2] # 第三列是劑量
        if frequency_col_name is None and len(df_meds.columns) > 5:
            frequency_col_name = df_meds.columns[5] # 第六列是頻率

        if not all([generic_name_col_name, dosage_col_name, frequency_col_name]):
            raise ValueError("DataFrame 中缺少必要的藥物資訊欄位（學名、劑量、頻率）。")
        
        generic_first_word_list = []
        # 遍歷 DataFrame 的每一行
        for index, row in df_meds.iterrows():
            # 檢查是否包含 "使用中" 這個關鍵字 (假設它可能在任何字串欄位中)
            # 更嚴謹的做法是你有一個明確的 status_column，例如 row['狀態'] == '使用中'
            # 這裡我們假設 "使用中" 存在於某個文字欄位中，或是可以直接檢查 row 的字串表示
            
            # 將整行轉換為字串來檢查 '使用中'，這可能不太精確，但與你原先邏輯類似
            # 更佳方式是檢查一個特定的狀態欄位
            if "使用中" not in str(row.values): # 轉換整行值為字串進行檢查
                continue # 如果不包含 "使用中"，則跳過這行

            generic_name = str(row[generic_name_col_name]).strip()
            dosage = str(row[dosage_col_name]).strip()
            frequency = str(row[frequency_col_name]).strip()
            
            # Apply text replacements to the generic name
            generic_name = apply_text_replacements(generic_name)
            
            # Skip medications we want to exclude
            excluded_meds = ["Sod chloride", "DEXTROSE", "Dextrose", "ChlorPHENIRAMINE", "Heparin", "Norm-Saline"]
            if any(generic_name.startswith(excluded) for excluded in excluded_meds):
                continue
                
            # Get the first word of the generic name
            generic_first_word = generic_name.split()[0]
            
            # Special case for certain prefixes that should be preserved
            prefixes = ["Sod", "Pot", "Mag.", "Cal."]
            if generic_first_word in prefixes:
                generic_first_word = generic_first_word # 保持原樣
            
            # Skip medications with frequency "ONCE"
            if frequency == "ONCE":
                if generic_first_word not in generic_first_word_list:
                    pass
                else:
                    continue

            if frequency == "ANES":
                continue 

            if dosage and frequency:
                formatted_med = generic_first_word
                generic_first_word_list.append(generic_first_word)
                # If dosage is 1, omit it
                if dosage != "1" and dosage != "X1":
                    formatted_med += f" {dosage}"
                
                if generic_first_word in generic_first_word_list:
                    # Remove any previous entry in the list that contains generic_first_word
                    for med in active_meds_list:
                        if ((generic_first_word in med) and ("ONCE" in med)):
                            active_meds_list.remove(med)
                            
                
                formatted_med += f" {frequency}"
                active_meds_list.append(formatted_med)


        
        # Join all medications with commas
        active_meds_list.sort()
        result = ", ".join(active_meds_list)
        
        # Copy the result back to clipboard for convenience
        #pyperclip.copy(result)
        
        return result
    
    except Exception as e:
        return f"???????"

def initialize_headless_driver(download_folder_path="C:\\Selenium_Downloads"):
    """
    Initializes and returns a **headless** Edge WebDriver instance, running
    in the background without a visible UI.
    Configures Edge to automatically download PDF files to a specified directory.
    It expects msedgedriver.exe to be in the same directory as the script.
    Handles common WebDriver setup errors.

    Args:
        download_folder_path (str): The absolute path where PDF files should be downloaded.
                                    (Note: This is currently overridden internally).
    
    Returns:
        webdriver.Edge: The initialized headless Edge WebDriver instance.
        None: If an error occurs during initialization.
    """
    download_folder_path = os.path.join(os.getcwd(), "PDF_Downloads")
    
    print(f"確保下載資料夾存在: {download_folder_path}")
    if not os.path.exists(download_folder_path):
        os.makedirs(download_folder_path)
        print(f"已建立下載資料夾: {download_folder_path}")

    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        # 備用方案：以防 __file__ 未定義 (例如在某些 IDE 或 notebook 中)
        try:
            application_path = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            application_path = os.getcwd()

    driver_path = os.path.join(application_path, 'msedgedriver.exe')

    if not os.path.exists(driver_path):
        print(f"錯誤：找不到驅動程式！")
        print("請從\"https://developer.microsoft.com/zh-tw/microsoft-edge/tools/webdriver?form=MA13LH#downloads\"下載最新版本的msedgedriver.exe ")
        print(f"確認 'msedgedriver.exe' 檔案存在於以下路徑：")
        print(f"{application_path}")
        input("\n請將 msedgedriver.exe 放到正確位置後，再重新執行。按 Enter 鍵結束...")
        sys.exit()

    check_driver_compatibility(driver_path)
    # 假設 check_driver_compatibility 存在
    # check_driver_compatibility(driver_path)

    service = Service(executable_path=driver_path)
    options = webdriver.EdgeOptions()
    
    # --- 【新增】Headless 模式選項 ---
    # 使用 "--headless=new" (推薦) 而不是舊的 "--headless"
    options.add_argument("--headless=new")
    # 在 Headless 模式下，通常建議停用 GPU 加速
    options.add_argument("--disable-gpu")
    # 某些網頁依賴視窗大小來渲染，設定一個預設值
    options.add_argument("--window-size=1920,1080")
    # --- Headless 選項結束 ---
    
    # --- 【保留】隱匿模式選項 (Stealth Options) ---
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36 Edg/128.0.0.0")
    # --- 隱匿模式選項結束 ---
    
    # 【保留】配置下載偏好設定
    prefs = {
        "download.default_directory": download_folder_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        # 這一行對於 Headless 模式自動下載 PDF 至關重要
        "plugins.always_open_pdf_externally": True, 
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.notifications": 2,
        'profile.default_content_settings.popups': 2,
    }
    options.add_experimental_option("prefs", prefs)
    
    try:
        driver = webdriver.Edge(service=service, options=options)
        
        # --- 【保留】在 driver 啟動後執行 JS 來隱藏 webdriver 標記 ---
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        print("Edge 瀏覽器已成功初始化 (Headless 模式)，並使用指定的驅動程式與下載設定（已啟用隱匿模式）。")
        
        # --- 【移除】視窗大小調整與定位相關程式碼 ---
        # 在 Headless 模式下，不需要 (也無法) 操作視窗大小或位置
        
        time.sleep(1) # 保留短暫等待，確保瀏覽器完全啟動
        return driver
    except WebDriverException as e:
        print(f"錯誤: 無法啟動 Edge 瀏覽器。請確保：")
        print(f"1. Edge 瀏覽器已安裝。")
        print(f"2. 位於 '{driver_path}' 的 msedgedriver.exe 檔案確實存在。")
        print(f"3. 該 msedgedriver.exe 的版本與您的 Edge 瀏覽器版本相容。")
        print(f"4. 錯誤訊息: {e}")
        return None

def initialize_driver():
    """
    初始化並回傳一個 Edge WebDriver 實例。
    會自動將 PDF 下載到當前工作目錄下的 'downloaded_PDF' 資料夾。
    此版本會明確指定使用當前目錄下的 msedgedriver.exe，並整合了隱匿模式以避免被偵測。
    """
    # --- 【關鍵修正】: 動態定義下載路徑 ---
    download_folder_path = os.path.join(os.getcwd(), "downloaded_PDF")
    
    print(f"確保下載資料夾存在: {download_folder_path}")
    if not os.path.exists(download_folder_path):
        os.makedirs(download_folder_path)
        print(f"已建立下載資料夾: {download_folder_path}")

    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    driver_path = os.path.join(application_path, 'msedgedriver.exe')

    if not os.path.exists(driver_path):
        print(f"錯誤：找不到驅動程式！")
        print("請從\"https://developer.microsoft.com/zh-tw/microsoft-edge/tools/webdriver?form=MA13LH#downloads\"下載最新版本的msedgedriver.exe ")
        print(f"確認 'msedgedriver.exe' 檔案存在於以下路徑：")
        print(f"{application_path}")
        input("\n請將 msedgedriver.exe 放到正確位置後，再重新執行。按 Enter 鍵結束...")
        sys.exit()


    check_driver_compatibility(driver_path)

    service = Service(executable_path=driver_path)
    options = webdriver.EdgeOptions()
    



    # --- 【新增】隱匿模式選項 (Stealth Options) ---
    # 1. 排除 "enable-automation" 開關，移除 "Edge 正由自動化軟體控制" 的提示
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    # 2. 停用自動化擴充功能
    options.add_experimental_option('useAutomationExtension', False)
    # 3. 設定一個常見的 User-Agent，偽裝成普通使用者
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36 Edg/128.0.0.0")
    # --- 隱匿模式選項結束 ---
    
    # 配置下載偏好設定
    prefs = {
        "download.default_directory": download_folder_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.notifications": 2,
        'profile.default_content_settings.popups': 2,
    }
    options.add_experimental_option("prefs", prefs)
    
    try:
        driver = webdriver.Edge(service=service, options=options)
        
        # --- 【新增】在 driver 啟動後執行 JS 來隱藏 webdriver 標記 ---
        # 這是避免被 JS 偵測的關鍵步驟
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        # --- JS 隱藏結束 ---
        
        print("Edge 瀏覽器已成功初始化，並使用指定的驅動程式與下載設定（已啟用隱匿模式）。")
        
        driver.maximize_window()
        time.sleep(0.5)
        screen_width = driver.get_window_size()['width']
        screen_height = driver.get_window_size()['height']
        driver.set_window_size(screen_width * 3 // 4, screen_height * 3 // 4)
        driver.set_window_position(0, 0)
        print(f"瀏覽器視窗已設定為 1/2 大小並置於左上角。")
        
        time.sleep(1)
        return driver
    except WebDriverException as e:
        print(f"錯誤: 無法啟動 Edge 瀏覽器。請確保：")
        print(f"1. Edge 瀏覽器已安裝。")
        print(f"2. 位於 '{driver_path}' 的 msedgedriver.exe 檔案確實存在。")
        print(f"3. 該 msedgedriver.exe 的版本與您的 Edge 瀏覽器版本相容。")
        print(f"4. 錯誤訊息: {e}")
        return None

def navigate_to_url(driver, url):
    """
    Navigates the driver to the specified URL.
    """
    try:
        print(f"Navigating to: {url}")
        driver.get(url)
        return True
    except Exception as e:
        print(f"Error navigating to {url}: {e}")
        return False

def login_to(driver, username, password, wait_timeout):
    """
    Performs the login sequence on the UpToDate website.
    Returns True if successful, False otherwise.
    """
    stop_the_code()
    wait = WebDriverWait(driver, wait_timeout)
    try:

        print("Entering username...")
        username_field = wait.until(EC.visibility_of_element_located((By.ID, "login_name")))
        username_field.send_keys(username)
        
        print("Username entered.")
        

        # Enter password
        print("Entering password...")
        password_field = wait.until(EC.visibility_of_element_located((By.ID, "password")))
        password_field.send_keys(password)
        password_field.send_keys(Keys.ENTER)
        print("Password entered. Login sequence complete.")
         # Wait for login to process and page to load
        return True

    except TimeoutException as e:
        print(f"Login failed: An element was not found or not interactable within {wait_timeout} seconds.")
        print(f"Error details: {e}")
        return False
    except NoSuchElementException as e:
        print(f"Login failed: Element not found in DOM.")
        print(f"Error details: {e}")
        return False
    except Exception as e:
        print(f"An unexpected error occurred during login: {e}")
        return False

def extract_table_data(driver, table_locator, wait_timeout):
    """
    從指定的表格中提取資料，使用 JavaScript 批量獲取所有單元格的文本。
    適用於 read_html 無法處理或表格包含互動元素的情況。
    """
    wait = WebDriverWait(driver, wait_timeout)
    extracted_df = None

    try:
        # Locate the table element
        table_element = wait.until(EC.visibility_of_element_located(table_locator))

        # --- Extract Header ---
        # Re-locate the table element just before executing script for headers
        # This reduces the chance of it becoming stale if a small DOM change occurs
        table_element_for_headers = wait.until(EC.visibility_of_element_located(table_locator))
        header_text_script = """
            var headerCells = arguments[0].querySelectorAll('th, thead td');
            var headers = [];
            for (var i = 0; i < headerCells.length; i++) {
                headers.push(headerCells[i].textContent.trim());
            }
            return headers.filter(h => h !== '');
        """
        header = driver.execute_script(header_text_script, table_element_for_headers)
        expected_header_len = len(header) if header else 0

        # --- Extract Data ---
        # Re-locate the table element just before executing script for data
        table_element_for_data = wait.until(EC.visibility_of_element_located(table_locator))
        data_text_script = """
            var rows = arguments[0].querySelectorAll('tr');
            var tableData = [];
            for (var i = 1; i < rows.length; i++) { // 從第二行開始 (假設第一行為標頭)
                var cells = rows[i].querySelectorAll('th, td');
                var rowCells = [];
                for (var j = 0; j < cells.length; j++) {
                    rowCells.push(cells[j].textContent.trim());
                }
                tableData.push(rowCells);
            }
            return tableData;
        """
        table_data = driver.execute_script(data_text_script, table_element_for_data)

        # Process data rows (same as your original logic)
        processed_table_data = []
        for row_cells in table_data:
            current_row_len = len(row_cells)
            if expected_header_len > 0:
                if current_row_len > expected_header_len:
                    row_cells = row_cells[:expected_header_len]
                elif current_row_len < expected_header_len:
                    row_cells.extend([''] * (expected_header_len - current_row_len))
            processed_table_data.append(row_cells)

        if processed_table_data:
            if header:
                extracted_df = pd.DataFrame(processed_table_data, columns=header)
            else:
                extracted_df = pd.DataFrame(processed_table_data)
                print("注意: 由於未找到或標頭為空，DataFrame 以預設數字欄位名建立。")
        else:
            print("處理後沒有從表格中提取到任何資料。")

    except TimeoutException:
        print(f"錯誤: 在 {wait_timeout} 秒內未找到或未顯示表格。")
        print(f"請檢查表格的 XPath 和當前頁面內容。定位器: {table_locator[1]}。")
    except NoSuchElementException:
        print(f"錯誤: 使用定位器 '{table_locator[1]}' 未在 DOM 中找到表格元素。")
        print("這可能表示定位器或頁面結構有問題。")
    except WebDriverException as e:
        print(f"JavaScript 執行或 WebDriver 相關錯誤: {e}")
    except Exception as e:
        print(f"表格提取過程中發生意外錯誤: {e}")

    return extracted_df

def key_in_field(driver, field_id, content_to_type, wait_timeout, press_enter=True):
    """
    Finds a text input field by its ID, clears it robustly,
    inputs content, and optionally presses Enter.

    Args:
        driver: Selenium WebDriver instance.
        field_id (str): The 'id' attribute of the target input field (e.g., "target01").
        content_to_type (str): The text content to input into the field.
        wait_timeout (int): The maximum number of seconds to wait for the input field to be visible.
        press_enter (bool): If True, presses Enter after typing. Defaults to True.

    Returns:
        bool: True if the operation is successful, False otherwise.
    """
    stop_the_code()
    wait = WebDriverWait(driver, wait_timeout)

    try:
        print(f"Attempting to find input field (id='{field_id}')...")
        input_field = wait.until(EC.visibility_of_element_located((By.ID, field_id)))
        print(f"Input field with ID '{field_id}' found.")

        # --- Robust Clearing Method ---
        # 1. Click the element to ensure it has focus
        input_field.click()
        time.sleep(random.uniform(0.01, 0.3)) # Small delay for focus to register
       
        # 2. Select all existing text (Ctrl+A)
        # Using CONTROL for Windows/Linux compatibility
        input_field.send_keys(Keys.CONTROL + "a")
        print("Sent Ctrl+A to input field.")
        time.sleep(random.uniform(0.01, 0.3)) # Small delay for selection to register

        # 3. Delete the selected text
        input_field.send_keys(Keys.DELETE) # Or Keys.BACKSPACE
        print("Sent DELETE key to clear input field.")
        time.sleep(random.uniform(0.3, 0.5)) # Give it a bit more time to clear visually
       
        # Input the new content
        print(f"Typing '{content_to_type}' into input field (id='{field_id}')...")
        input_field.send_keys(content_to_type)
        print("Content typed.")

        if press_enter:
            print("Pressing Enter...")
            input_field.send_keys(Keys.ENTER)
            print("Enter pressed.")
            # Give time for the page to react/submit

        return True

    except TimeoutException as e:
        print(f"Error: Input field (id='{field_id}') not found or not visible within {wait_timeout} seconds.")
        print(f"Error details: {e}")
        return False
    except NoSuchElementException as e:
        print(f"Error: Input field (id='{field_id}') not found in DOM.")
        print(f"Error details: {e}")
        return False
    except Exception as e:
        print(f"An unexpected error occurred during interaction with input field '{field_id}': {e}")
        return False

    except TimeoutException as e:
        print(f"Error: Search bar (id='???') not found or not visible within {wait_timeout} seconds.")
        print(f"Error details: {e}")
        return False
    except NoSuchElementException as e:
        print(f"Error: Search bar (id='???') not found in DOM.")
        print(f"Error details: {e}")
        return False
    except Exception as e:
        print(f"An unexpected error occurred during search bar interaction: {e}")
        return False

def find_and_click_first_inpatient_date_link(driver = None, inpatient_text: str = "住院", wait_timeout: int = 10) -> str | None:
    """
    在表格中找到第二欄包含 '住院' 文字的第一列，
    點擊該列第一欄中的日期連結，並回傳該連結的文字。

    Args:
        driver: Selenium WebDriver 實例。
        inpatient_text (str): 用於識別住院類型的文字 (預設為 '住院')。
        wait_timeout (int): 等待元素出現的最長秒數。

    Returns:
        str | None: 如果成功點擊，則回傳連結的文字 (例如 '113/09/01')；如果失敗，則回傳 None。
    """
    stop_the_code()
    wait = WebDriverWait(driver, wait_timeout)

    try:
        print(f"正在搜尋第一筆 '{inpatient_text}' 資料及其對應的日期連結...")

        # XPath 表示式：找到一個 <tr>，其第二個 <td> 包含指定文字，然後選取該 <tr> 的第一個 <td> 內的 <a> 標籤
        xpath_expression = (
            f"//tr[td[2][normalize-space()='{inpatient_text}']]/td[1]/a"
        )
        
        link_element_locator = (By.XPATH, xpath_expression)
        
        print(f"正在使用 XPath 尋找元素: {xpath_expression}")
        
        # 等待元素變得可點擊
        link_element = wait.until(
            EC.element_to_be_clickable(link_element_locator)
        )
        
        # 取得連結的文字並儲存起來
        link_name = link_element.text
        
        print(f"找到目標連結: '{link_name}'，準備點擊...")
        
        # 點擊該元素
        link_element.click()
        
        print(f"已成功點擊 '{link_name}'。")
        
        # 回傳儲存的連結文字
        return link_name

    except TimeoutException:
        print(f"錯誤：在 {wait_timeout} 秒內找不到 '{inpatient_text}' 資料，或其日期連結不可點擊。")
        return None
    except NoSuchElementException:
        print(f"錯誤：在 DOM 中找不到 '{inpatient_text}' 資料的對應元素。")
        return None
    except Exception as e:
        print(f"發生未預期的錯誤：{e}")
        return None

def click_specific_link(driver: webdriver.Remote, locator_type: By, locator_value: str, wait_timeout: int,
                        main_window_handle: str = None, close_popups_after_click: bool = False,
                        popup_wait_timeout: float = 0.5) -> bool:
    """
    查找並點擊一個特定的網頁元素。
    如果配置，點擊後會檢查並關閉任何新彈出的視窗，並將焦點切回主視窗。

    Args:
        driver: Selenium WebDriver 實例。
        locator_type: 用於定位的 By 策略 (例如 By.XPATH, By.ID)。
        locator_value (str): 定位器的值。
        wait_timeout (int): 等待元素可點擊的最長秒數。
        main_window_handle (str, optional): 如果點擊可能彈出新視窗，請提供主視窗句柄。
                                            預設為 None。
        close_popups_after_click (bool): 如果為 True，點擊後會檢查並關閉新彈出視窗。
                                         預設為 False。
        popup_wait_timeout (float): 等待新彈出視窗出現的最長秒數，僅在
                                    close_popups_after_click 為 True 時有效。預設 0.5 秒。

    Returns:
        True: 如果元素被成功點擊，且彈窗處理（如果啟用）成功。
        False: 否則。
    """
    stop_the_code()
    time.sleep(random.uniform(0.01, 0.1))
    # stop_the_code() # 根據你的全局設定，保留或移除這個檢查
    wait = WebDriverWait(driver, wait_timeout)
    # 在點擊前記錄所有視窗句柄，用於後續判斷是否有新視窗彈出
    initial_window_handles_before_click = driver.window_handles 

    try:
        print(f"嘗試尋找並點擊元素 (使用 {locator_type} 和值: '{locator_value}')...")
        
        # 等待元素可點擊
        element = wait.until(
            EC.element_to_be_clickable((locator_type, locator_value))
        )
        
        element.click()
        print(f"成功點擊元素: '{locator_value}'.")
        
        # --- 新增的彈窗處理邏輯 ---
        if close_popups_after_click and main_window_handle:
            print("啟用彈窗處理：點擊後檢查新視窗...")
            try:
                # 等待視窗數量增加 (表示有新視窗彈出)
                WebDriverWait(driver, popup_wait_timeout).until(
                    EC.number_of_windows_to_be(len(initial_window_handles_before_click) + 1)
                )
                print(f"偵測到新視窗出現。")

                # 獲取所有當前視窗句柄
                all_current_handles = driver.window_handles
                # 找出所有不在點擊前就存在的句柄，這些就是彈出的視窗
                pop_up_handles = [handle for handle in all_current_handles if handle not in initial_window_handles_before_click]

                if pop_up_handles:
                    print(f"找到 {len(pop_up_handles)} 個彈出視窗準備關閉。")
                    for handle_to_close in pop_up_handles:
                        try:
                            driver.switch_to.window(handle_to_close)
                            print(f"已切換到彈出視窗 (標題: {driver.title}, URL: {driver.current_url})。")
                            driver.close() # 關閉當前焦點所在的視窗/分頁
                            print("彈出視窗已成功關閉。")
                        except NoSuchWindowException:
                            print(f"警告: 彈出視窗 {handle_to_close} 已不存在或已關閉。")
                        except Exception as e:
                            print(f"關閉彈出視窗 {handle_to_close} 時發生錯誤: {e}")
                    
                    # 關閉所有彈出視窗後，確保 WebDriver 的焦點切換回主視窗
                    driver.switch_to.window(main_window_handle)
                    print("已將焦點切換回主視窗。")
                else:
                    print("未偵測到新的彈出視窗。")
            except TimeoutException:
                print(f"在 {popup_wait_timeout} 秒內未偵測到新的彈出視窗。")
                # 如果沒有偵測到新視窗（超時），但 WebDriver 的焦點可能不在主視窗，則嘗試切回
                if driver.current_window_handle != main_window_handle:
                    driver.switch_to.window(main_window_handle)
                    print("已將焦點切換回主視窗（因未偵測到彈窗超時）。")
            except Exception as e:
                print(f"彈窗處理過程中發生意外錯誤: {e}")
                # 在任何錯誤情況下，都嘗試確保焦點回到主視窗
                if main_window_handle and driver.current_window_handle != main_window_handle:
                    driver.switch_to.window(main_window_handle)
                    print("已將焦點切換回主視窗（因彈窗處理錯誤）。")
        # --- 彈窗處理邏輯結束 ---

        return True
    except TimeoutException:
        print(f"錯誤: 元素未在 {wait_timeout} 秒內找到或不可點擊。")
        print(f"定位器: {locator_type}, 值: '{locator_value}'.")
        return False
    except NoSuchElementException:
        print(f"錯誤: 元素未在 DOM 中找到 (使用 {locator_type} 和值: '{locator_value}').")
        return False
    except Exception as e:
        print(f"點擊元素時發生意外錯誤: {e}")
        return False
    
# --- (其他你的函式定義，例如 initialize_driver, navigate_to_url, login_to_uptodate 等都保持不變) ---

# --- 通用下拉選單選取函式定義 ---
def select_option_from_dropdown(driver, dropdown_locator_type, dropdown_locator_value,
                                option_select_by_type, option_value, wait_timeout):
    """
    選取網頁中任何下拉選單的特定選項。

    Args:
        driver: Selenium WebDriver 實例。
        dropdown_locator_type: 定位下拉選單 (<select> 元素) 的 By 策略 (例如 By.ID, By.CLASS_NAME, By.XPATH)。
        dropdown_locator_value (str): 對應 dropdown_locator_type 的值 (例如 ID 字串, Class 名稱字串, XPath 字串)。
        option_select_by_type (str): 選取選項的方式 ('visible_text', 'value', 或 'index')。
        option_value: 欲選取的選項的值。
                      如果是 'visible_text'，傳入字串 (例如 "二週內")。
                      如果是 'value'，傳入字串 (例如 "14")。
                      如果是 'index'，傳入整數 (例如 1)。
        wait_timeout (int): 等待下拉選單元素出現的最長秒數。

    Returns:
        bool: 如果成功選取選項則返回 True，否則返回 False。
    """
    stop_the_code()
    wait = WebDriverWait(driver, wait_timeout)
    
    try:
        print(f"嘗試尋找下拉選單 ({dropdown_locator_type}, '{dropdown_locator_value}')...")
        dropdown_element = wait.until(
            EC.presence_of_element_located((dropdown_locator_type, dropdown_locator_value))
        )
        print("下拉選單已找到。")

        # 創建 Select 物件來操作下拉選單
        select = Select(dropdown_element)

        print(f"正在選取選項 '{option_value}' (透過 {option_select_by_type})...")
        if option_select_by_type == 'visible_text':
            select.select_by_visible_text(str(option_value))
        elif option_select_by_type == 'value':
            select.select_by_value(str(option_value))
        elif option_select_by_type == 'index':
            select.select_by_index(int(option_value))
        else:
            print(f"錯誤：無效的 option_select_by_type '{option_select_by_type}'。必須是 'visible_text', 'value' 或 'index'。")
            return False
            
        print(f"已成功選取選項 '{option_value}'。")
        
        # 留點時間讓頁面內容更新

        return True

    except TimeoutException:
        print(f"錯誤：下拉選單 ({dropdown_locator_type}, '{dropdown_locator_value}') 未在 {wait_timeout} 秒內找到。")
        return False
    except NoSuchElementException:
        print(f"錯誤：下拉選單 ({dropdown_locator_type}, '{dropdown_locator_value}') 或指定選項未在 DOM 中找到。")
        return False
    except Exception as e:
        print(f"選取下拉選單選項時發生意外錯誤：{e}")
        return False

def extract_IO(driver, original_window_handle, wait_timeout=15):
    """
    Switches to the first newly opened window after clicking '連結NIS' (implied click before this function),
    then clicks the '輸出入量查詢' button which opens a second new window,
    switches to the second new window, performs operations, closes it,
    then closes the first new window, and finally switches back to the original window.

    Args:
        driver: Selenium WebDriver instance.
        original_window_handle: The handle of the original browser window.
        wait_timeout (int): The maximum number of seconds to wait for new windows and elements.

    Returns:
        bool: True if all operations in the new windows were successful, False otherwise.
    """
    print("\n--- Entering extract_IO function ---")
    
    first_new_window_handle = None
    second_new_window_handle = None

    try:
        # --- 處理第一個新視窗的開啟和切換 (這是您原有的邏輯) ---
        print("Waiting for the FIRST new window to open...")
        # 等待視窗數量從1變為2
        WebDriverWait(driver, wait_timeout).until(EC.number_of_windows_to_be(2))
        print("First new window detected.")
        
        all_window_handles = driver.window_handles
        for handle in all_window_handles:
            if handle != original_window_handle:
                first_new_window_handle = handle
                break
        
        if not first_new_window_handle:
            print("Error: Could not find the first new window handle after it opened.")
            return None

        driver.switch_to.window(first_new_window_handle)
        print(f"Switched to FIRST new window. Title: {driver.title}, URL: {driver.current_url}")

        # --- 在第一個新視窗中點擊「輸出入量查詢」按鈕，這會打開第二個新視窗 ---
        print("Attempting to click '輸出入量查詢' button in the first new window...")
        # 假設點擊 By.ID, "mm13" 會在新標籤頁或新視窗中開啟內容
        button_clicked = click_specific_link(driver, By.ID, "mm14", wait_timeout)
        
        if not button_clicked:
            print("Failed to click '輸出入量查詢' button in the first new window. Aborting.")
            # 如果點擊失敗，第一視窗可能還開著，確保關閉並切回原視窗
            try:
                driver.close()
                driver.switch_to.window(original_window_handle)
            except Exception as close_e:
                print(f"Error during cleanup after first click failure: {close_e}")
            return None

        print("Successfully clicked '輸出入量查詢' button in the first new window.")
        
        # --- 處理第二個新視窗的開啟和切換 (新增的邏輯) ---
        print("Waiting for the SECOND new window to open (total windows to be 3)...")
        # 現在等待視窗數量從2變為3
        WebDriverWait(driver, wait_timeout).until(EC.number_of_windows_to_be(3))
        print("Second new window detected.")

        all_window_handles_after_second_open = driver.window_handles
        # 找出第二個新視窗的句柄（既不是原始視窗也不是第一個新視窗的句柄）
        for handle in all_window_handles_after_second_open:
            if handle != original_window_handle and handle != first_new_window_handle:
                second_new_window_handle = handle
                break
        
        if not second_new_window_handle:
            print("Error: Could not find the second new window handle after it opened. Aborting.")
            # 如果找不到第二視窗，嘗試關閉第一視窗並切回原視窗
            try:
                driver.close() # 關閉第一視窗
                driver.switch_to.window(original_window_handle)
            except Exception as close_e:
                print(f"Error during cleanup after second window not found: {close_e}")
            return None

        driver.switch_to.window(second_new_window_handle)
        print(f"Switched to SECOND new window. Title: {driver.title}, URL: {driver.current_url}")
        try:
            # --- 在第二個新視窗中執行操作 ---
            if datetime.today().weekday()==0:
                click_specific_link(driver, By.XPATH, "//div[normalize-space()='前一頁']", wait_timeout)
                time.sleep(1 + random.uniform(0.01, 0.2))

            df = extract_table_data(driver, 
                table_locator = (By.XPATH, '(//table[@style="table-layout:fixed;word-break:break-all"])[2]'), 
                wait_timeout=1.5)
            print(df)
            print(df[0])
            output = None
            IO_value = None
            input_value = None
            yesterday_weekday = (datetime.today()- timedelta(days=1)).weekday() +1
            print(f"Yesterday's weekday (1-7): {yesterday_weekday}")
            
            output = find_data_and_clean(df, '排出(cc)', yesterday_weekday)
            IO_value = find_data_and_clean(df, '輸入-排出(cc)', yesterday_weekday)
            input = find_data_and_clean(df, '輸入(cc)', yesterday_weekday)
            stool = find_data_and_clean(df, '排便次數', yesterday_weekday)
            urine = find_data_and_clean(df, '排尿', yesterday_weekday)
            HD_value = find_data_and_clean(df, '透析', yesterday_weekday)
            Drainage = find_data_and_clean(df, '引流', yesterday_weekday)

            today = " " + get_today_date_formatted()
            IOfinal = None
            if urine is None: 
                urine = 0
            if output or input or IO_value is not None:
                IOfinal = f'I/U: {input}/{urine}({IO_value})'

            if stool != '' and stool is not None:
                IOfinal = IOfinal + f" stool: {stool}"
                
            if HD_value != '' and HD_value is not None:
                IOfinal = IOfinal + f" HD: {HD_value}"

            if Drainage != '' and Drainage is not None:
                IOfinal = IOfinal + f" drain: {Drainage}"


            print('排出:',output)
            print('輸入-排出:',IO_value)
        except Exception as e:
            print(f"Error occurred while extracting I/O data: {e}")

        # 為了觀察效果，可以暫停一下，或者等待數據載入 (推薦使用 WebDriverWait)
        

        # --- 關閉第二個新視窗 ---
        print("Closing the SECOND new window...")
        driver.close() 
        print("Second new window closed.")
        
        # --- 切換回第一個新視窗 (因為第二個新視窗關閉後，焦點通常會回到上一個活動視窗) ---
        driver.switch_to.window(first_new_window_handle)
        print(f"Switched back to FIRST new window. Title: {driver.title}")

        # --- 關閉第一個新視窗 ---
        print("Closing the FIRST new window...")
        driver.close()
        print("First new window closed.")

        # --- 切換回原始視窗 ---
        driver.switch_to.window(original_window_handle) 
        print(f"Switched back to ORIGINAL window. Original window title: {driver.title}")
        
        return IOfinal
        
        
       

    except TimeoutException as te:
        print(f"Error: Operation timed out within {wait_timeout} seconds: {te}")
        return None
    except NoSuchWindowException as nswe:
        print(f"Error: A window was not found during the operation: {nswe}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred in extract_IO: {e}")
        return None
    finally:
        # 這個 finally 區塊作為最終保障，確保在任何情況下都能嘗試回到原始視窗。
        # 如果前面已經成功切換回去了，這裡的操作不會有副作用。
        # 如果因為某些原因失敗了，這裡會嘗試修復。
        try:
            # 只有當目前不在原始視窗時才嘗試切換回去
            if driver.current_window_handle != original_window_handle:
                print("Attempting final cleanup: ensuring switch back to original window.")
                driver.switch_to.window(original_window_handle)
                print(f"Final cleanup successful. Switched back to original window. Title: {driver.title}")
        except Exception as cleanup_e:
            print(f"Error during final cleanup (switching back to original window): {cleanup_e}")


def find_data_and_clean(df, target_str, output_col_number=0):
    """
    從 DataFrame 中查找 column 0 中符合特定字串的列，並回傳該列中指定欄位的值。
    在處理數字前，會先清理字串，去除不必要的非數字字符，但會保留開頭的正負號。
    如果找到的值是數字或可以轉換為數字的文字，則四捨五入到最接近的整數，去除小數點，然後轉回文字。
    如果最終結果不是數字，則回傳 None。

    Args:
        df (pd.DataFrame): 輸入的 Pandas DataFrame。
        target_str (str): 想要查找的字串。
        output_col_number (int, optional): 想要回傳值的欄位索引（從 0 開始）。
                                         預設為 0，即第一個欄位。

    Returns:
        str or None: 符合條件的列中，指定 output_col_number 欄位處理過後的值（文字格式）。
                     如果最終結果不是數字，則回傳 None。
                     如果沒有找到符合條件的列，也回傳 None。
    """
    # 固定 target_col_number 為 0
    target_col_number = 0

    # 確保欄位號碼在 DataFrame 的有效範圍內
    if not (0 <= target_col_number < df.shape[1]):
        print(f"錯誤：'查找的column number' {target_col_number} 超出 DataFrame 範圍。")
        return None
    if not (0 <= output_col_number < df.shape[1]):
        print(f"錯誤：'輸出column number' {output_col_number} 超出 DataFrame 範圍。")
        return None

    # 根據欄位號碼獲取欄位名稱
    target_col_name = df.columns[target_col_number]
    output_col_name = df.columns[output_col_number]

    # 篩選出目標字串所在的列
    filtered_df = df[df[target_col_name] == target_str]

    # 如果找到符合條件的列，則處理並回傳值
    if not filtered_df.empty:
        # 獲取第一筆匹配的值
        result_value = filtered_df[output_col_name].iloc[0]

        # 確保值是字串型別以便應用正則表達式，並移除前後空白
        str_value = str(result_value).strip()

        # --- 清理步驟 (保留正負號) ---
        sign = ''
        if str_value.startswith('-'):
            sign = '-'
            str_value = str_value[1:] # 移除負號以便後續清理
        elif str_value.startswith('+'): # 如果明確有正號，也移除
            str_value = str_value[1:]

        # 使用正則表達式，只保留數字和一個小數點
        cleaned_value = re.sub(r'[^\d.]+', '', str_value)

        # 如果有多個點，只保留第一個點
        if cleaned_value.count('.') > 1:
            parts = cleaned_value.split('.')
            cleaned_value = parts[0] + '.' + ''.join(parts[1:])
        
        # 將負號加回 (如果存在)
        cleaned_value = sign + cleaned_value
        # ------------------------

        # 嘗試將清理後的值轉換為數字，四捨五入，然後轉回文字
        try:
            # 嘗試轉換為浮點數
            num_value = float(cleaned_value)
            # 四捨五入到最接近的整數
            rounded_value = round(num_value)
            # 將四捨五入後的整數轉為文字
            final_output_value = str(int(rounded_value))
            return final_output_value
        except (ValueError, TypeError):
            # 如果清理後仍然無法轉換為數字，則回傳 None
            return None
    else:
        # 如果沒有找到匹配的列，則印出訊息並回傳 None
        print(f"找不到 '{target_str}' 在欄位 '{target_col_name}' 中的資料。")
        return None
    
def extract_and_round_number(input_value) -> str:
    """
    接收一個字串、整數或浮點數，提取其中的數字，四捨五入到整數位，
    然後回傳其字串形式。

    - 如果輸入是 str，會使用正規表示式提取第一個數字序列。
    - 能處理數字前後的空格或非數字字元。
    - 如果無法提取或轉換數字，則回傳 'None'。

    :param input_value: 要進行轉換的輸入值 (str, int, float)。
    :return: 四捨五入後的數字字串，或 'None'。
    """
    # 步驟 1: 檢查輸入值是否為 None，如果是則直接回傳 'None'
    if input_value is None:
        return 'None'

    # 步驟 2: 將輸入統一轉換為字串以便處理
    s = str(input_value)

    # 步驟 3: 使用正規表示式從字串中尋找第一個數字序列
    # \d+ 會匹配一個或多個數字。我們尋找第一個出現的數字組合。
    match = re.search(r'\d+', s)

    # 步驟 4: 如果在字串中找到了數字
    if match:
        try:
            # 提取找到的數字字串
            number_str = match.group(0)
            # 將數字字串轉換為浮點數
            number = float(number_str)
            # 四捨五入到最接近的整數
            rounded_number = round(number)
            # 將結果轉為字串並回傳
            return str(rounded_number)
        except (ValueError, TypeError):
            # 如果在轉換過程中出錯（雖然此處機率很低），依然回傳 'None'
            return 'None'
    else:
        # 步驟 5: 如果整個字串都找不到任何數字
        return 'None'
    
def get_first_li_text(driver, wait_timeout):
    """
    定位第一個 <li> 元素，並盡力提取病患的各項資訊。
    找不到的特定資訊會以 None 回傳。
    """
    # --- 步驟 1: 初始化所有變數為 None ---
    # 這個步驟確保了即使後續任何搜尋失敗，這些變數也都有一個預設的 None 值
    name, patient_id, birthday, sex, bed_info = None, None, None, '男', None
    
    wait = WebDriverWait(driver, wait_timeout)
    
    try:
        # 找到包含資訊的 <li> 元素
        first_li_element = wait.until(
            EC.presence_of_element_located((By.TAG_NAME, "li"))
        )
        full_text = first_li_element.text.strip()
        print(f"找到 <li> 元素，內容為: '{full_text}'")

        # --- 步驟 2: 獨立、逐步地搜尋每一項資訊 ---

        # 1. 尋找姓名
        match_name = re.search(r'^([\u4e00-\u9fa5\s]+)', full_text)
        if match_name:
            name = " ".join(match_name.group(1).split())

        # 2. 尋找病歷號 (8位數字)
        match_id = re.search(r'\b(\d{8})\b', full_text)
        if match_id:
            patient_id = match_id.group(1)

        # 3. 尋找生日 (括號內的8位數字)
        match_bday = re.search(r'\((\d{8})\)', full_text)
        if match_bday:
            birthday = match_bday.group(1)

        # 4. 尋找性別
        match_sex = re.search(r'(男性|女性)', full_text)
        if match_sex:
            sex = '男' if match_sex.group(1) == '男性' else '女'

        # 5. 尋找床號
        match_bed = re.search(r'(?:男性|女性)\s*([\s\S]*?)\(住院中\)', full_text)
        if match_bed:
            bed_info = " ".join(match_bed.group(1).split())

    except (TimeoutException, Exception) as e:
        print(f"提取資訊時發生錯誤: {e}")
        # 如果發生任何錯誤，所有變數將維持初始的 None 值

    # --- 步驟 3: 回傳所有變數 ---
    # 無論找到多少資訊，都會回傳一個包含五個元素的元組
    return name, patient_id, birthday, sex, bed_info

def get_value_by_key_from_unnamed_df(df: pd.DataFrame, search_key: str) -> str | None:
    """
    從一個列名是預設數字 (0, 1) 的 Pandas DataFrame 中，根據指定的鍵
    (位於索引 0 的描述文字) 查找並回傳位於索引 1 的對應值。

    Args:
        df (pd.DataFrame): 包含鍵值對的 DataFrame。
                           預期第一列 (索引 0) 包含描述性鍵，
                           第二列 (索引 1) 包含對應的值。
        search_key (str): 要查找的鍵字串，例如 "０２．病房床號："。

    Returns:
        str or None: 如果找到匹配的鍵，則回傳對應的值（已去除首尾空白）；
                     否則回傳 None。
    """
    # 1. 輸入檢查
    if not isinstance(df, pd.DataFrame):
        print("錯誤: 輸入的 'df' 必須是一個 Pandas DataFrame。")
        return None
    
    if df.empty or df.shape[1] < 2:
        print("錯誤: 輸入的 DataFrame 格式不正確或為空。預期至少有兩列。")
        return None

    # 2. 清理 DataFrame 的第一列和搜索鍵以進行穩健匹配
    # 將第一列 (索引 0) 轉換為字串類型，並移除首尾空白和中文冒號 '：'
    # 使用 .assign() 創建一個新的 DataFrame 來避免 SettingWithCopyWarning
    temp_df = df.assign(
        _CleanKey=df.iloc[:, 0].astype(str).str.strip().str.replace('：', '', regex=False).str.strip()
    )
    
    # 清理要查找的 search_key，移除首尾空白和中文冒號 '：'
    # --- HERE IS THE FIX ---
    clean_search_key = search_key.replace('：', '').strip() # Removed regex=False
    # --- END FIX ---

    # 3. 查找匹配的行
    # 在清理過的臨時列中查找與 cleaned search_key 匹配的行
    matching_rows = temp_df[temp_df['_CleanKey'] == clean_search_key]

    if not matching_rows.empty:
        # 如果找到匹配的行，獲取第二列 (索引 1) 的值，並去除首尾空白
        return str(matching_rows.iloc[0, 1]).strip()
    else:
        # 如果沒有找到匹配的鍵，打印錯誤並回傳 None
        print(f"錯誤: 未找到與 '{search_key}' 匹配的鍵。")
        return None
    
def extract_patient_ids_as_list(df: pd.DataFrame) -> list[str]:
    """
    從輸入的 DataFrame 中提取「病歷號」欄位的數字部分，
    去除「New」等文字，並將結果作為一個字串列表回傳。

    Args:
        df (pd.DataFrame): 包含「病歷號」欄位的 DataFrame。

    Returns:
        list[str]: 包含清理後的病歷號 (字串) 的列表。
    """
    if '病歷號' not in df.columns:
        raise ValueError("輸入的 DataFrame 中沒有找到 '病歷號' 欄位。")

    patient_ids_series = df['病歷號']
    cleaned_ids = patient_ids_series.apply(lambda x: re.sub(r'[^\d]', '', str(x)))

    return cleaned_ids.tolist()

def fix_dataframe_rowspan_issues(df):
    """
    修正 Pandas DataFrame 中因 HTML 表格 rowspan 屬性導致的數據缺失問題，
    特別是當 rowspan 導致後續行的數據列左移時。
    
    該函數會識別並填充 '門/住' 列（通常是 DataFrame 的第一列），
    如果其值為空或被其他數據（如藥物名稱）填充，則用上一組的有效值來填充。

    Args:
        df (pd.DataFrame): 從 HTML 表格提取出來的原始 Pandas DataFrame。
                           預期第一列可能包含 '門/住' 的值，並在 rowspan 時為空或左移。

    Returns:
        pd.DataFrame: 修正後的 DataFrame，其中 rowspan 導致的缺失值已被填充，
                      並將左移的數據歸位（如果原始提取導致了左移）。
    """
    if df.empty:
        print("警告: 輸入的 DataFrame 為空，無需修正。")
        return df

    # 複製 DataFrame 以免修改原始數據
    corrected_df = df.copy()

    original_columns = df.columns.tolist()
    
    # 初始化一個空的列表來儲存修正後的行數據
    processed_rows = []
    
    last_valid_type_value = None # 用於追蹤 '門/住' 的上一個有效值

    for index, row in corrected_df.iterrows():
        current_first_col_value = str(row.iloc[0]).strip() # 獲取當前行的第一個值

        if current_first_col_value in ['OPD', 'ADM']:
            # 這是一個新的 '門/住' 組的開始，儲存該值
            last_valid_type_value = current_first_col_value
            processed_rows.append(row.tolist()) # 將當前行完整添加到結果中
        else:
            # 這是一個因為 rowspan 而「左移」的行
            # 它的第一個元素現在是藥物名稱，'門/住' 的值需要從上一個有效值填充
            
            if last_valid_type_value is not None:
              
                new_row_data = [last_valid_type_value] + row.tolist()
             
                adjusted_row = [last_valid_type_value] + row.iloc[0:].tolist() # 將第0列替換
                processed_rows.append(adjusted_row)
            else:
                # 如果沒有上一個有效的 '門/住' 值（例如，DataFrame 的第一行就是左移的），
                # 這種情況不應在正常數據中發生，但為了穩健性，可以填充為空
                print(f"警告: DataFrame 第 {index} 行缺失 '門/住' 值且無前值可填充。")
                processed_rows.append([''] + row.tolist()) # 填充空值

    # 初始化一個空的列表來儲存修正後的行數據
    final_rows_data = []
    
    # 標題行會是原始 DataFrame 的列名
    final_headers = ['門/住'] + df.columns.tolist() # 新增 '門/住' 列

    last_valid_type_value = None

    for index, row in df.iterrows():
        current_row_list = row.tolist()
        first_col_value = str(current_row_list[0]).strip()

        if first_col_value in ['OPD', 'ADM']:
            # 這是新組的開始，第一列就是 '門/住'
            last_valid_type_value = first_col_value
            final_rows_data.append(current_row_list)
        else:
            # 這是被 rowspan 影響的行
            # 原始的第一列 (current_row_list[0]) 實際上是「標準藥物名稱」
            # 我們需要將它推到正確的位置，並在新的第一列填充 '門/住'
            
            if last_valid_type_value is not None:
                # 創建一個新行：[填充的門/住值] + [原始行中除了第一列之外的所有數據]
                # 注意：這裡假設原始 DataFrame 的第一列就是應該被填充的列，
                # 且在 rowspan 情況下，該列包含了應該屬於第二列的數據。
                new_row = [last_valid_type_value] + current_row_list
                final_rows_data.append(new_row)
            else:
                # 異常情況：如果第一行就不是 'OPD'/'ADM' 且沒有上一個值
                print(f"警告: 第 {index} 行數據異常，無法識別 '門/住' 值。數據: {current_row_list}")
                final_rows_data.append([''] + current_row_list) # 插入空值作為佔位

    # 由於我們在 `final_rows_data` 中已經加入了修正後的數據
    # 並且我們在 `final_headers` 中預設了第一列是 '門/住'
    # 現在創建最終的 DataFrame
    if not final_rows_data:
        # 如果沒有任何數據行，只創建一個空 DataFrame 帶有新的列名
        return pd.DataFrame(columns=final_headers)
    

    # 為了兼容這種情況，我們需要檢查原始 DataFrame 的列名，並相應調整 `final_headers`
    if df.columns[0] != '門/住':
        # 假設原始 DF 的第一列是藥物名稱，但實際上應該是門/住的佔位
        # 我們將 '門/住' 插入到最前面，並保持其餘原始列名
        new_column_names = ['門/住'] + df.columns.tolist()
    else:
        new_column_names = df.columns.tolist() # 原始 DF 列名正確

    final_processed_data = []
    
    for row_list in final_rows_data:
        # 確保每一行的長度與新列名數量匹配
        # 如果行太短，填充空字串
        while len(row_list) < len(new_column_names):
            row_list.append('')
        # 如果行太長，截斷
        final_processed_data.append(row_list[:len(new_column_names)])

    # 創建最終的 DataFrame
    return pd.DataFrame(final_processed_data, columns=new_column_names)

def get_chemo_dose_date_looped(driver: webdriver.Edge, original_window_handle: str, wait_timeout: int) -> pd.DataFrame:
    """
    先一次開啟所有新分頁，在每個分頁中設定時間並點擊搜尋，
    等待所有搜尋完成後，再逐一複製表格數據，並將其合併為一個 DataFrame。
    已新增處理所有 DataFrame 都為空時的情況，並確保視窗正確關閉和切換。

    Args:
        driver (webdriver.Edge): Selenium WebDriver 物件。
        original_window_handle (str): 原始視窗的句柄。
        wait_timeout (int): 等待元素出現的超時時間。

    Returns:
        pd.DataFrame: 合併所有數據後的 DataFrame。如果所有嘗試提取的數據都為空，則返回一個帶有預期欄位的空 DataFrame。
    """


    print("\n--- 階段一: 一次開啟所有視窗並執行搜尋 ---")
    first_new_window_handle = None
    driver.switch_to.window(original_window_handle)
    old_window_handles = set(driver.window_handles)

    try:
        if not click_specific_link(driver, By.LINK_TEXT, "癌症資訊", wait_timeout):
            print(f"無法點擊 '癌症資訊' 連結，跳過。")
            return pd.DataFrame() # 如果無法點擊，直接返回空 DataFrame
        if not click_specific_link(driver, By.LINK_TEXT, "癌症藥物治療紀錄查詢", wait_timeout):
            print(f"無法點擊 '癌症藥物治療紀錄查詢' 連結，跳過。")
            return pd.DataFrame() # 如果無法點擊，直接返回空 DataFrame
        # --- 處理第一個新視窗的開啟和切換 (這是您原有的邏輯) ---
        print("Waiting for the FIRST new window to open...")
        # 等待視窗數量從1變為2
        WebDriverWait(driver, wait_timeout).until(EC.number_of_windows_to_be(2))
        print("First new window detected.")
        
        all_window_handles = driver.window_handles
        for handle in all_window_handles:
            if handle != original_window_handle:
                first_new_window_handle = handle
                break
        
        if not first_new_window_handle:
            print("Error: Could not find the first new window handle after it opened.")
            return False

        driver.switch_to.window(first_new_window_handle)
        print(f"Switched to FIRST new window. Title: {driver.title}, URL: {driver.current_url}")

        
    except (WebDriverException, TimeoutException, Exception) as e: # 捕獲更廣泛的異常
        # 在發生錯誤時，嘗試切換回原始視窗，確保驅動程式不會停留在無效的句柄上
        try:
            if original_window_handle in driver.window_handles:
                driver.switch_to.window(original_window_handle)
        except (NoSuchWindowException, InvalidSessionIdException):
            print("錯誤：原始視窗已關閉或無效，無法切換回去。")
        except Exception as switch_e:
            print(f"切換回原始視窗時發生意外錯誤: {switch_e}")
        

    time.sleep(1.3 + random.uniform(0,0.1)) # 等待所有查詢結果載入



    try: 
        current_df = extract_table_data(driver, table_locator=(By.ID, "importASTR_medicine_view"), wait_timeout=wait_timeout)
        current_df = fix_dataframe_rowspan_issues(current_df)
        
    except:
        print(f" 提取到空數據或 '無資料'。")

  
    driver.close() # 關閉當前焦點所在的視窗
    
    print(f" 視窗已關閉。")

    try:
        driver.switch_to.window(original_window_handle)
        print("已切換回原始視窗，準備處理下一個分頁。")
    except (NoSuchWindowException, InvalidSessionIdException):
        print("錯誤：原始視窗已關閉或無效，無法切換回去。後續操作可能受影響。")
            # 如果原始視窗都沒了，可能無法繼續

    return current_df



def get_final_chemo_summary_flexible(dataframe, 
                                     drug_col, 
                                     date_col, 
                                     dose_col, 
                                     proximity_days=14):
    """
    (版本：月/日)
    根據實際的欄位名稱處理化療週期資料，回傳最終化療資訊的字串。
    日期格式為 M/D。
    """
    if dataframe is None or dataframe.empty:
        print("警告: 輸入的 DataFrame 為 None 或空，無法處理。")
        return None

    
 
    df_copy = dataframe.copy()
    df_copy['Drug'] = df_copy[drug_col].str.strip().str.lower()
    df_copy['StartDate'] = pd.to_datetime(df_copy[date_col])
    df_copy['CleanDose'] = df_copy[dose_col].str.extract(r'(\d+)')
    # 根據 '途徑' 欄位進行特殊處理
    # 如果 '途徑' 為 'PO'，則在 'CleanDose' 後加上 '#'
    df_copy.loc[df_copy['途徑'] == 'PO', 'CleanDose'] = df_copy.loc[df_copy['途徑'] == 'PO', 'CleanDose'].astype(str) + '#'
    # 如果 '途徑' 為 'IT'，則在 '途徑' 後加上 '.IT'
    df_copy.loc[df_copy['途徑'] == 'IT', 'CleanDose'] = df_copy.loc[df_copy['途徑'] == 'IT', 'CleanDose'].astype(str) + '.IT'

    df_copy['MedDoseStr'] = df_copy['Drug'] + ' ' + df_copy['CleanDose'] + ' ' + df_copy['頻次']
    # 當 '頻次' 不是 'ONCE' 或 'STAT' 時，將 '醫囑天數（實際天數）' 加入 MedDoseStr
    mask = ~df_copy['頻次'].isin(['ONCE', 'STAT'])
    df_copy.loc[mask, 'MedDoseStr'] = (
        df_copy.loc[mask, 'MedDoseStr'] + '*' +
        df_copy.loc[mask, '醫囑天數（實際天數）'].str.split('(').str[0]
    )
    

    daily_summary = df_copy.groupby('StartDate')['MedDoseStr'].agg(', '.join).reset_index()
    daily_summary = daily_summary.sort_values(by='StartDate').reset_index(drop=True)

    # ===================== 修改關鍵點 =====================
    # 將日期格式從 strftime('%Y-%m-%d') 改成 月/日 的組合
    # .dt.month 和 .dt.day 會取得不帶前導零的數字
    date_str = daily_summary['StartDate'].dt.month.astype(str) + '/' + daily_summary['StartDate'].dt.day.astype(str)
    daily_summary['FullString'] = '(' + date_str + ': ' + daily_summary['MedDoseStr'] + ')'
    # ===================================================

    if len(daily_summary) < 2:
        return daily_summary.iloc[0]['FullString'] if len(daily_summary) == 1 else ""

    last_date = daily_summary.iloc[-1]['StartDate']
    second_last_date = daily_summary.iloc[-2]['StartDate']
    
    if (last_date - second_last_date).days <= proximity_days:
        second_last_string = daily_summary.iloc[-2]['FullString']
        last_string = daily_summary.iloc[-1]['FullString']
        return f"{second_last_string} & {last_string}"
    else:
        last_string = daily_summary.iloc[-1]['FullString']
        return last_string


def scrape_table_text_with_newlines(driver, wait_timeout, table_locator_type=By.XPATH, table_locator_value="//table[@cellspacing='0' and @cellpadding='0' and @width='100%']"):
    """
    等待指定的 table 元素可見後，抓取其所有文字內容，並保留換行。
    
    Args:
        driver: Selenium WebDriver 實例。
        wait_timeout (int): 等待元素可見的最長秒數。
        table_locator_type: 定位 table 元素的 By 策略 (例如 By.XPATH)。
        table_locator_value (str): 對應 table_locator_type 的值。
        
    Returns:
        str: 包含 table 中所有文字內容的原始字串（保留換行）。
             如果找不到 table 或超時，則返回 None。
    """
    try:
        wait = WebDriverWait(driver, wait_timeout)
        
        # 等待 table 元素可見
        print(f"正在等待 table 元素 '{table_locator_value}' 在 {wait_timeout} 秒內可見...")
        table_element = wait.until(
            EC.visibility_of_element_located((table_locator_type, table_locator_value))
        )
        print("Table 元素已找到且可見。")
        
        # 直接獲取 table 元素的全部文字內容，保留原始格式
        all_text = table_element.text
        
        return all_text
        
    except TimeoutException:
        print(f"錯誤：table 元素在 {wait_timeout} 秒內未可見。")
        return None
    except NoSuchElementException:
        print("錯誤：找不到指定的 table 元素。")
        return None
    except Exception as e:
        print(f"提取 table 文字時發生意外錯誤: {e}")
        return None

def extract_assessment_between_markers(text_content: str) -> str | None:
    """
    從一個字串中提取'診斷(Assessment):'和'治療計畫(Plan):'之間的文字。
    如果第一次找到的文字長度小於10個字元，則會繼續尋找下一個符合條件的區塊，
    直到找到一個長度大於或等於10個字元的內容為止。

    參數:
    text_content (str): 包含文字資料的單一字串。

    返回:
    str | None: 如果找到長度大於或等於10的診斷內容，則返回其內容。
                如果所有找到的內容長度都小於10，或找不到任何標記，則返回None。
    """
    # 定義開始和結束的標記
    start_marker = '診斷(Assessment):'
    end_marker = '治療計畫(Plan):'
    
    # 初始化搜尋的起始位置
    search_start_pos = 0

    while True:
        # 尋找開始標記的位置，從上一次結束的位置繼續
        start_index = text_content.find(start_marker, search_start_pos)

        # 如果找不到開始標記，則跳出循環
        if start_index == -1:
            print(f"找不到開始標記: '{start_marker}'")
            return None

        # 從開始標記之後尋找結束標記
        # 我們需要加上開始標記的長度，以確保從正確的位置開始尋找
        start_of_content = start_index + len(start_marker)
        end_index = text_content.find(end_marker, start_of_content)

        # 如果找不到結束標記，則跳出循環
        if end_index == -1:
            print(f"找不到結束標記: '{end_marker}'")
            return None

        # 提取'診斷(Assessment):'和'治療計畫(Plan):'之間的子字串
        extracted_text = text_content[start_of_content:end_index]
        cleaned_text = extracted_text.strip()

        # 檢查清理後的文字長度是否符合要求
        if len(cleaned_text) >= 10:
            return cleaned_text
        else:
            print(f"找到的內容 '{cleaned_text}' 長度小於10，繼續尋找下一個。")
            # 更新搜尋的起始位置，以便下次從結束標記之後開始尋找
            search_start_pos = end_index

    # 如果循環結束後沒有找到符合條件的內容
    return None

def calculate_pf_ratio(df):
  """
  計算P/F ratio並新增一個 'P/F' 欄位到DataFrame中。

  這個函式會處理 'PO2' 和 'FIO2' 欄位中的非數字或遺失值。

  Args:
    df: 包含 'PO2' 和 'FIO2' 欄位的Pandas DataFrame。

  Returns:
    一個新的DataFrame，其中新增了 'P/F' 欄位。
  """
  # 使用 pd.to_numeric 將 'PO2' 和 'FIO2' 欄位轉換為數字
  # errors='coerce' 會將任何無法轉換的值變成 NaN
  po2_numeric = pd.to_numeric(df['PO2'], errors='coerce')
  fio2_numeric = pd.to_numeric(df['FIO2'], errors='coerce')

  # 計算 P/F ratio。Pandas 會自動處理 NaN 值，
  # 任何涉及 NaN 的計算結果都會是 NaN
  # 另外，我們也處理FIO2為0的狀況，避免除以0的錯誤
  df['P/F'] = np.where(fio2_numeric == 0, np.nan, round(po2_numeric / fio2_numeric * 100,0))

  return df


def process_data_with_date_conversion(df, column_name='IG G'):
    """
    處理 DataFrame，將日期欄位轉換為標準格式，並獲取指定欄位中
    最新的兩個數值以及最新值對應的日期。
    
    參數:
    df (DataFrame): 輸入的 DataFrame。
    column_name (str): 要處理的欄位名稱，預設為 'IG G'。
    
    返回:
    tuple: 包含 (最新值, 第二新值, 最新值的日期)。
           如果資料不足或處理失敗，則返回 None。
    """
    try:
        # 複製 DataFrame 以免修改原始資料
        temp_df = df.copy()
        
        # 轉換日期格式
        try:
            temp_df['日期'] = pd.to_datetime(temp_df['日期'], format='%y-%m-%d %H:%M', errors='coerce')
        except:
            temp_df['日期'] = pd.to_datetime(temp_df['日期'], format='%Y/%m/%d', errors='coerce')
        
        # 將日期格式化為 'YYYY/MM/DD'
        temp_df['日期'] = temp_df['日期'].dt.strftime('%Y/%m/%d')
        
        # --- 新增的處理邏輯 ---
        # 處理 'CMV_REALTIME' 欄位中的特定字串值
        if column_name == 'CMV_REALTIME':
            # 將 'CMV not detected' 替換為 -1
            temp_df[column_name] = temp_df[column_name].replace('CMV not detected', -1)
            temp_df[column_name] = np.where(
            temp_df[column_name].astype(str).str.contains('undetectable', case=False, na=False),-1,  temp_df[column_name])
            # 使用正規表達式來提取數字
            # 'CMV detected, <34.5 IU/mL' 轉為 34.5, 然後取整數
            # 'CMV detected, 214 IU/mL' 轉為 214
            # 這裡我們用一個更通用的方式來處理
            temp_df[column_name] = temp_df[column_name].astype(str).str.extract(r'(\d+\.?\d*)').astype(float)
            
            # 處理 "<34.5" 的情況，將其替換為 34
            temp_df[column_name] = temp_df[column_name].mask(temp_df[column_name] == 34.5, 34)

        # 將其他非數字字元替換為 NaN
        temp_df[column_name] = pd.to_numeric(
            temp_df[column_name].replace(['-', ' '], np.nan),
            errors='coerce'
        )
        
        # 刪除目標欄位中所有 NaN 值的行
        temp_df.dropna(subset=[column_name], inplace=True)
        
        # 確保有足夠的數據
        if len(temp_df) < 2:
            print("警告：資料不足，無法取得最新的兩個數值。")
            if len(temp_df) == 1:
                latest_value = temp_df[column_name].iloc[0]
                latest_date = temp_df['日期'].iloc[0]
                return latest_value, None, latest_date
            return None, None, None

        # 取得最新（最後一筆）和第二新（倒數第二筆）的數值
        latest_value = temp_df[column_name].iloc[-1]
        second_latest_value = temp_df[column_name].iloc[-2]

        # 取得最新值對應的日期
        latest_date = temp_df['日期'].iloc[-1]

        return latest_value, second_latest_value, latest_date

    except Exception as e:
        print(f"處理資料時發生錯誤: {e}")
        return None, None, None
    
def format_comparison(data, round_places=None):
    """
    格式化包含舊值、新值和日期的列表為一個易讀的字串。
    
    參數:
    data (list): 包含 [舊值, 新值, 日期] 的列表，舊值或新值可能是 None。
    round_places (int or None): 四捨五入到小數點後指定的位數。若為 None，則不進行四捨五入。
    
    返回:
    str or None: 格式化後的字串，如果舊值和新值都為 None，則返回 None。
    """
    try:
        if not isinstance(data, list) or len(data) != 3:
            raise ValueError("輸入必須是一個包含三個元素的列表 [old, new, date]。")
            
        old_value, new_value, date_str = data
        
        # 處理四捨五入邏輯
        if isinstance(round_places, int):
            if old_value is not None:
                old_value = round(old_value, round_places)
                if round_places == 0:
                    old_value = int(old_value)  # 移除 .0
            if new_value is not None:
                new_value = round(new_value, round_places)
                if round_places == 0:
                    new_value = int(new_value)  # 移除 .0

        # 處理舊值和新值都為 None 的情況
        if old_value is None and new_value is None:
            return None
            
        # 處理舊值為 None 的情況
        if old_value is None:
            return f"{new_value} {date_str}"
            
        # 處理新值為 None 的情況
        if new_value is None:
            return f"{old_value} {date_str}"

        return f"{old_value}>{new_value} {date_str}"


    except Exception as e:
        print(f"格式化比較資料時發生錯誤: {e}")
        return None

def format_date_with_parentheses(date_string):
    """
    將日期字串從 'YYYY/MM/DD' 格式轉換為 '(M/D)' 格式。
    
    參數:
    date_string (str): 格式為 'YYYY/MM/DD' 的日期字串。
    
    返回:
    str or None: 格式化後的日期字串，如果輸入格式無效則返回 None。
    """
    try:
        # 將字串解析為 datetime 物件
        parsed_date = datetime.strptime(date_string, "%Y/%m/%d")
        
        # 格式化為 '(M/D)'，並移除月份和日的前導零
        formatted_date = parsed_date.strftime("(%#m/%#d)")
        
        # 處理非 Windows 系統的前導零問題
        # 在 Linux/macOS 上，使用 lstrip('0').replace('/0', '/')
        if formatted_date[1] == '0' or '/0' in formatted_date:
            formatted_date = formatted_date.replace('(', '').replace(')', '')
            formatted_date = f"({formatted_date.lstrip('0').replace('/0', '/')})"
            
        return formatted_date
        
    except (ValueError, TypeError) as e:
        print(f"錯誤：日期格式無效，無法轉換。錯誤訊息: {e}")
        return None

def get_foley_lines(driver: webdriver.Edge, original_window_handle: str, wait_timeout: int) -> str:
    """
    爬取表格數據，篩選處置名稱/別名中的管路關鍵字，並彙整成計數的字串 (e.g., CVC*2)。
    優先匹配較長的/更具體的管路名稱，並將最終結果標準化。
    """
    
    print("\n--- 階段一: 一次開啟所有視窗並執行搜尋 ---")
    first_new_window_handle = None
    driver.switch_to.window(original_window_handle)
    # ... (省略舊有邏輯，保持不變) ...
    main_window_handle = driver.current_window_handle
    try:
        if not click_specific_link(driver, By.LINK_TEXT, "醫師功能", wait_timeout):
            print(f"無法點擊 '醫師功能' 連結，跳過。")
            return "None" 
        if not click_specific_link(driver, By.LINK_TEXT, "治療處置功能", wait_timeout):
            print(f"無法點擊 '治療處置功能' 連結，跳過。")
            return "None"
            
        # --- 處理第一個新視窗的開啟和切換 ---
        print("Waiting for the FIRST new window to open...")
        WebDriverWait(driver, wait_timeout).until(EC.number_of_windows_to_be(2))
        print("First new window detected.")
        
        all_window_handles = driver.window_handles
        for handle in all_window_handles:
            if handle != original_window_handle:
                first_new_window_handle = handle
                break
        
        if not first_new_window_handle:
            print("Error: Could not find the first new window handle after it opened.")
            return "None"

        driver.switch_to.window(first_new_window_handle)
        print(f"Switched to FIRST new window. Title: {driver.title}, URL: {driver.current_url}")

    except (WebDriverException, TimeoutException, Exception) as e:
        print(f"階段一發生錯誤: {e}")
        try:
            if original_window_handle in driver.window_handles:
                driver.switch_to.window(original_window_handle)
        except (NoSuchWindowException, InvalidSessionIdException):
            print("錯誤：原始視窗已關閉或無效，無法切換回去。")
        except Exception as switch_e:
            print(f"切換回原始視窗時發生意外錯誤: {switch_e}")
        return "None"
        
    # 點擊篩選連結
    click_specific_link(driver, By.XPATH, "//span[text()='僅顯示生效處置']", wait_timeout)

    time.sleep(1.3 + random.uniform(0,0.1)) # 等待所有查詢結果載入

    # --- 階段二: 提取數據並處理 ---
    line_str = "None"

    try: 
        # 獲取 DataFrame
        current_df = extract_table_data(driver, table_locator=(By.CLASS_NAME, "dataTableList"), wait_timeout=wait_timeout)
        
        if current_df is None or current_df.empty or '處置名稱' not in current_df.columns:
             print(f" 提取到空數據或缺少 '處置名稱' 欄位。")
        else:
            # 1. 定義和排序要篩選的管路關鍵字 
            raw_target_lines = [
                '2-Lumen', 'NG Tube', 'Foley', 'CVC', 'Arterial Line', 'A-line', 
                'PICC', 'Chest Tube', 'T-drain', 'JP Tube', 'NJ Tube',
                'ECMO', 'IABP', 'Peripherally Inserted Central Catheter', 'Anal Tube'
            ]
            
            # 依字串長度降序排序，確保先匹配最長的/最精確的名稱
            target_lines_sorted = sorted(raw_target_lines, key=len, reverse=True)
            
            # 用於初步篩選的 pattern (只需包含所有關鍵字即可)
            pattern = '|'.join([re.escape(line) for line in target_lines_sorted])
            
            print(f"\n--- 階段二: 篩選管路數據 (處置名稱/別名) ---")
            
            # 2. 檢查 '處置名稱' 欄位
            mask_name = current_df['處置名稱'].astype(str).str.contains(pattern, case=False, na=False)
            
            # 3. 檢查 '別名' 欄位並結合 
            if '別名' in current_df.columns:
                mask_alias = current_df['別名'].astype(str).str.contains(pattern, case=False, na=False)
                filtered_series = mask_name | mask_alias
            else:
                filtered_series = mask_name

            # 4. 提取所有匹配的行
            matched_df = current_df[filtered_series]
            
            # 5. 彙整找到的標準管路名稱並計數
            line_counts = Counter()
            
            for index, row in matched_df.iterrows():
                matched_in_row = set()
                
                # 遍歷 '處置名稱' 和 '別名'
                fields_to_check = [str(row['處置名稱'])]
                if '別名' in current_df.columns:
                    fields_to_check.append(str(row['別名']))

                for field_text in fields_to_check:
                    # 依長度排序後的管路列表進行匹配 (長字串優先)
                    for line in target_lines_sorted:
                        # 檢查是否包含該管路名稱
                        if re.search(re.escape(line), field_text, re.IGNORECASE):
                            # 找到最長匹配後，記錄下來並跳出當前欄位檢查
                            matched_in_row.add(line)
                            break 
                    
                # 將這一行找到的所有不重複管路計數 +1
                for found_line in matched_in_row:
                    line_counts[found_line] += 1 

            # 6. 組合結果字串 (帶有計數)
            if line_counts:
                formatted_lines = []
                # 依管路名稱排序
                for line in sorted(line_counts.keys()):
                    count = line_counts[line]
                    if count > 1:
                        formatted_lines.append(f"{line}*{count}")
                    else:
                        formatted_lines.append(line)
                        
                line_str = ', '.join(formatted_lines)
                print(f"✅ 找到的管路 (原始): {line_str}")
                
                # 7. 執行不區分大小寫的縮寫替換 (新增邏輯)
                
                replacements = {
                    r'Peripherally Inserted Central Catheter': 'PICC',
                    r'Arterial Line': 'A-line',
                    r'JP Tube': 'JP',
                    r'NG Tube': 'NG'
                }

                # 使用 re.sub() 進行不區分大小寫替換
                for old, new in replacements.items():
                    # 替換時使用 re.escape(old) 確保特殊字符被正確處理
                    # 使用 re.IGNORECASE 旗標來實現不區分大小寫
                    line_str = re.sub(re.escape(old), new, line_str, flags=re.IGNORECASE)
                
                print(f"✅ 找到的管路 (標準化): {line_str}")

            else:
                line_str = None
                print("🔍 未找到任何目標管路。")

# --- 【B】 新增階段三：提取飲食/配方和熱量資訊 (主要修改區) ---
            print(f"\n--- 階段三: 提取飲食/配方與熱量資訊 (別名) ---")
            
            # ==========================================================
            # ❗ 新增 NPO 優先檢查邏輯
            # ==========================================================
            if '別名' in current_df.columns:
                # 遍歷 '別名' 欄位，尋找是否包含獨立的 NPO (使用單詞邊界 \b 確保是整個單詞)
                # 使用 current_df['別名'].astype(str) 確保處理的是字串
                npo_found = current_df['別名'].astype(str).str.contains(r'\bNPO\b', case=True, na=False).any() 
                TPN_found = current_df['別名'].astype(str).str.contains('TPN care', case=True, na=False).any() 
                try_water = current_df['別名'].astype(str).str.contains('Try Water', case=True, na=False).any()    
                try_D5water = current_df['別名'].astype(str).str.contains('Try Glucose Water', case=True, na=False).any()  
                all_day_nutrition = current_df['別名'].astype(str).str.contains('全日營養品', case=True, na=False).any()                  
                oral_nutrition = current_df['別名'].astype(str).str.contains('口服營養補充品', case=True, na=False).any()   

                
                if TPN_found:
                    diet_str = "TPN"
                    print(f"🚨 檢測到 TPN 處置，設定 diet_str='TPN' 並跳過其他飲食/配方搜尋。")
                    # 由於已經找到最高優先級的 NPO，直接跳過後續的飲食/配方分析
                    # 接下來會直接執行窗口關閉和切換邏輯
                    pass # 跳過後續的 found_diets 邏輯

                elif try_water:
                    diet_str = "try water"
                    # 由於已經找到最高優先級的 NPO，直接跳過後續的飲食/配方分析
                    # 接下來會直接執行窗口關閉和切換邏輯
                    pass # 跳過後續的 found_diets 邏輯   

                                  
                elif try_D5water:
                    diet_str = "try D5W"
                    print(f"🚨 檢測到 D5W 處置，設定 diet_str='D5W' 並跳過其他飲食/配方搜尋。")
                    # 由於已經找到最高優先級的 NPO，直接跳過後續的飲食/配方分析
                    # 接下來會直接執行窗口關閉和切換邏輯
                    pass # 跳過後續的 found_diets 邏輯     

                elif npo_found:
                    diet_str = "NPO"
                    print(f"🚨 檢測到 NPO 處置，設定 diet_str='NPO' 並跳過其他飲食/配方搜尋。")
                    # 由於已經找到最高優先級的 NPO，直接跳過後續的飲食/配方分析
                    # 接下來會直接執行窗口關閉和切換邏輯
                    pass # 跳過後續的 found_diets 邏輯
                              

                elif all_day_nutrition:
                    diet_str = "全日營養品"
                    # 由於已經找到最高優先級的 NPO，直接跳過後續的飲食/配方分析
                    # 接下來會直接執行窗口關閉和切換邏輯
                    pass # 跳過後續的 found_diets 邏輯  
            
                elif oral_nutrition:
                    diet_str = "營養補充品"
                    # 由於已經找到最高優先級的 NPO，直接跳過後續的飲食/配方分析
                    # 接下來會直接執行窗口關閉和切換邏輯
                    pass # 跳過後續的 found_diets 邏輯  


                else:
                    # 如果沒有 NPO，才執行原本的飲食/配方尋找邏輯
                    found_diets = set() 
                    
                    # Regex 1: 尋找 (文字 + '飲食'或'配方')。
                    diet_pattern = re.compile(r'(.*?)(飲食|配方)', re.IGNORECASE)
                    
                    # Regex 2: 尋找熱量資訊 (數字 + kcal/Calorie)
                    kcal_pattern = re.compile(r'熱量：(\d+)\s*k?cal', re.IGNORECASE)

                    for index, row in current_df.iterrows():
                        alias_text = str(row.get('別名', ''))
                        
                        if not alias_text or alias_text.lower() == 'nan':
                            continue

                        # 1. 嘗試匹配 '飲食' 或 '配方' 的名稱
                        diet_match = diet_pattern.search(alias_text)
                        
                        if diet_match:
                            # 基礎飲食名稱：將 Group 1 (前面的文字) 和 Group 2 (關鍵字) 連接起來
                            base_diet = (diet_match.group(1).strip() + diet_match.group(2)).strip()
                            
                            if not base_diet:
                                continue
                            

                                    
                            
                            
                            # 2. 在整個字串中尋找熱量資訊 (kcal)
                            
                            alias_text2 = str(row.get('頻次', ''))
                            kcal_match = kcal_pattern.search(alias_text2)

                            if diet_match.group(1).strip() in ('質地調整', '製作處理調整'):
                                try:
                                    texture_pattern = re.compile(r'質地：(.*?)(?=製作處理|\s|\n|\(|$)')
                                    diet_match = texture_pattern.search(alias_text2)
                                    base_diet = diet_match.group(1).strip() + '飲食'
                                    print(base_diet)
                                except:
                                    print('no texture data')

                            final_diet_name = base_diet

                            if kcal_match:
                                kcal_value = kcal_match.group(1) # 熱量數字
                                final_diet_name = f"{base_diet}-{kcal_value}"
                            
                            # 3. 添加到集合中，確保唯一性
                            if final_diet_name:
                                found_diets.add(final_diet_name)
                    

                    
                    # 4. 組合最終字串
                    if found_diets:
                        diet_str = ', '.join(sorted(found_diets))
                        print(f"✅ 找到的飲食/配方: {diet_str}")
                    else:
                        diet_str = None
                        print("🔍 未找到任何飲食或配方處置。")

                    if diet_str:
                        # 1. 刪除 '肝病' 和 '腎病' (替換為空字串實現刪除)
                        diet_str = diet_str.replace('肝病', '').replace('腎病', '')
                        
                        # 2. 替換 '飲食' 為 '餐'
                        diet_str = diet_str.replace('飲食', '餐')

            else:
                print("⚠️ 數據中缺少 '別名' 欄位，跳過飲食/配方提取。")

    except Exception as e:
        print(f" 提取數據或處理邏輯發生錯誤: {e}")
        line_str = None
    

    # 關閉視窗邏輯（保持不變）
    
    # 檢查視窗句柄是否有效，防止 NoSuchWindowException
    try:
        if driver.current_window_handle != original_window_handle:
            driver.close() # 關閉當前焦點所在的視窗
            print(f" 視窗已關閉。")
    except NoSuchWindowException:
        print("警告: 嘗試關閉視窗時，視窗已不存在。")
    except Exception as e:
        print(f"關閉視窗時發生錯誤: {e}")

    # 切換回原始視窗
    try:
        driver.switch_to.window(original_window_handle)
        print("已切換回原始視窗，準備處理下一個分頁。")
    except (NoSuchWindowException, InvalidSessionIdException):
        print("錯誤：原始視窗已關閉或無效，無法切換回去。後續操作可能受影響。")

    return line_str, diet_str

def format_medications_to_columns(med_string):
    """
    將包含多種藥物的字串轉換為兩欄的格式化字串（縱向排列）。

    Args:
        med_string: 包含藥物名稱的單一字串，用逗號分隔。

    Returns:
        一個格式化好的字串，將藥物分為兩欄顯示。
    """
    # 1. 轉換成清單
    medication_list = [med.strip() for med in med_string.split(',')]

    # 2. 如果清單數量為奇數，則在最後補上一個空字串
    if len(medication_list) % 2 != 0:
        medication_list.append("")

    # 3. 計算每一欄的項目數量
    half_length = len(medication_list) // 2
    first_column = medication_list[:half_length]
    second_column = medication_list[half_length:]

    formatted_output = []
    # 4. 迴圈處理並格式化
    for i in range(half_length):
        first_item = first_column[i]
        second_item = second_column[i]
        formatted_output.append(f"{first_item:<23}{second_item}")

    # 5. 回傳最終字串
    return "\n".join(formatted_output)


def main():
    failure_count = 0
    # --- 將 driver 初始化為 None ---
    global driver_instance
    driver = None 
    
    try:

        # Get the handle to the current console window
        #hWnd = ctypes.windll.kernel32.GetConsoleWindow()

        # Minimize the window (SW_MINIMIZE is 6)
        #ctypes.windll.user32.ShowWindow(hWnd, 6)
        gender = "male"
        app = xw.apps.active
        wb = app.books.active
        ws = wb.sheets.active
        current = app.selection
        Drug_col = ws.range('E1')
        Anti="Y"
        BW="Y"
        drug_col_hidden = ws.api.Columns(Drug_col.column).Hidden
        Anti=ws.range('O2')
        BW=ws.range('P2')
        R_exclude = ws.range('N2')
        Medication = ws.range('Q2')
        
        USERNAME = str(ws.range('S1').value)
        PASSWORD = str(ws.range('U1').value)
        CMV_list = None
        IgG_list = None

        
        IgG_str = None
        CBCDC = []
        Glucose = None
        SMAC_string = None
        bloodgas_string = None
        
        bodyweight = None
        R_excluding = False
        VS_excluding = False
        culture_found=1

        stop_the_code()

        
        include_IO = True
        include_drug = Include_drug
        include_anti = Include_anti
        body_weight = Include_body_weight
        include_thyroid = Include_thyroid
        location = None
        patient=10
        foley_including = True
        chemo_including = True
        drug_formating = True
        anti_days = True
        ai_asking = False
        try:
   
            if ws.range('O1').value is not None and ws.range('O1').value.lower()=="x":
                R_excluding=True
                print("不取 R")  
            if ws.range('M1').value is not None and ws.range('M1').value.lower()=="x":
                VS_excluding=True
                print("不取 VS")
            if ws.range('Q1').value is not None and ws.range('Q1').value.lower()=="x":
                foley_including=False 
                print("不取 管路")
            if ws.range('W1').value is not None and ws.range('W1').value.lower()=="x":
                chemo_including=False
                print("不取 Chemo")
            if ws.range('Y1').value is not None and ws.range('Y1').value.lower()=="x":
                drug_formating=False
                print("不藥物排版")
            if ws.range('AA1').value is not None and ws.range('AA1').value.lower()=="x":
                anti_days=False
                print("抗生素用 Month/Day-")
            else:
                print("抗生素用 anti D?")
            if ws.range('AG1').value is not None and ws.range('AG1').value.lower()=="v":
                ai_asking=True
                ws.range('M:M').api.WrapText = False
                print("整理AI提示")
            else:
                print("不整理AI提示")
        except:
            pass

        department = ws.range('AC1').value if ws.range('AC1').value is not None else 'HEMA'
        glu_number = 3
        if department in ['META']:
            glu_number = 7

        
        try:
            ws.api.Cells.Font.Name = 'Calibri'
            if drug_formating:
                ws.range('E:E').api.Font.Name = 'Consolas'
            ws.range('M:M').api.WrapText = False
            ws.range('A2:A25').rows.autofit()
            

        
        except:
            pass

        current = move_to_row_start(current, column=1)
        if current.value is None or len(extract_and_round_number(current.value)) < 5 or len(extract_and_round_number(current.value)) > 18:
            current = ws.range('A2')
            current.select

        
        
        
        
        driver = None
        #if ws.range('AE1').value is not None and ws.range('AE1').value.lower()=='x' :
            #driver = initialize_headless_driver()
        #else:
        driver = initialize_driver()

        navigate_to_url(driver, url = "https://web9.vghtpe.gov.tw/")
        if USERNAME != 'None' and PASSWORD != 'None':
            login_to(driver, USERNAME, PASSWORD, WAIT_TIMEOUT*8)
        main_window_handle = driver.current_window_handle
        click_specific_link(driver, By.XPATH, "//div[contains(@class, 'link__management__item--text') and normalize-space(.)='DRWEBAPP醫師作業（含病歷查詢)']", 60, main_window_handle, True, 1.3)
        df = extract_table_data(driver, 
                table_locator=(By.ID, "patlist"), 
                wait_timeout=WAIT_TIMEOUT)
        if current.value == None:
            patient_list = extract_patient_ids_as_list(df)
            current.value = [[pid] for pid in patient_list]
        ID_number = extract_and_round_number(current.value)
        print(ID_number)
        patient=len(ID_number)
        while patient>4 and patient < 16:
            try:
                target_field_id = "target01"
                IO_str = None
                date_chemo_dose_string = None
                bed_number = None
                name = None
                culture_str= None
                sex = '男'
                VS_name = None
                VS_number = None
                R_name = None
                R_number = None
                Sex_age = None
                Tumor_marker_string = None
                antibiotic_str=None
                Critical = False
                key_in_field(driver, target_field_id, content_to_type = ID_number, 
                            wait_timeout = WAIT_TIMEOUT, press_enter=True)
                time.sleep(0.2)
                df = extract_table_data(driver, 
                    table_locator=(By.XPATH, '//table[.//th[text()="功能"]]'), 
                    wait_timeout=WAIT_TIMEOUT)
                print(df)
                if '@' in df.iloc[0,1]:
                    Critical = True 

                click_specific_link(driver, By.XPATH, "//a[normalize-space()='查詢']", WAIT_TIMEOUT)
                
                
                progress_note = None
                assessment = None
                foley_lines = None
                chemo_dose_date = None
                diet_str = None
                if True:

                    original_window_handle2 = driver.current_window_handle
                    if foley_including:
                        try:
                            foley_lines, diet_str = get_foley_lines(driver, original_window_handle2,  WAIT_TIMEOUT)
                        except:
                            print('failing extracting line and diet data')
                        
                    
                    if chemo_including:
                        chemo_dose_date = get_chemo_dose_date_looped(driver, original_window_handle2,  WAIT_TIMEOUT)
                        print(chemo_dose_date)
                    
                        date_chemo_dose_string = get_final_chemo_summary_flexible(
                        dataframe=chemo_dose_date,
                        drug_col='標準藥物名稱',
                        date_col='開始日期',
                        dose_col='本次使用總劑量（用藥次數）'
                        )
                    print(date_chemo_dose_string)

                click_specific_link(driver, By.XPATH, "//a[text()='護理紀錄']", WAIT_TIMEOUT)
                click_specific_link(driver, By.XPATH, "//img[@alt='關閉']", 0.2)
                click_specific_link(driver, By.XPATH, "(//img[@alt='關閉'])[2]", 0.01)
                click_specific_link(driver, By.XPATH, "(//img[@alt='關閉'])[3]", 0.01)
                click_specific_link(driver, By.ID, "btn_germ", 0.01)
                click_specific_link(driver, By.ID, "btn_pchkau", 0.01)
                click_specific_link(driver, By.XPATH, "//input[@value='確定送出']", 0.01)
                
                
                
 
                if include_IO == True:
                    original_window_handle = driver.current_window_handle
                    
                    click_specific_link(driver, By.XPATH, "//input[@value='連結NIS']", WAIT_TIMEOUT)            
                    IO_str = extract_IO(driver = driver, original_window_handle = driver.current_window_handle, wait_timeout=WAIT_TIMEOUT)
       

                name, ID, birthday, sex, bed_number = get_first_li_text(driver, WAIT_TIMEOUT) 
                click_specific_link(driver, By.XPATH, "//a[text()='病患資料»']", WAIT_TIMEOUT)
                click_specific_link(driver, By.XPATH, "//a[text()='基本資料']", WAIT_TIMEOUT)

                df = extract_table_data(driver, 
                    table_locator=(By.XPATH, "//tbody[./tr/td[contains(text(), '０１．病歷號')]]"), 
                    wait_timeout=WAIT_TIMEOUT)  

                
                bed_number = get_value_by_key_from_unnamed_df(df,"０２．病房床號：")
                VS_name = get_value_by_key_from_unnamed_df(df,"１８．主治醫師：")
                R_name = get_value_by_key_from_unnamed_df(df,"１９．住院醫師：")
                Age = get_value_by_key_from_unnamed_df(df,"０４．生　日　：")
                Age = Age.split('（')[1]
                Age = Age.split('歲')[0]
                Sex_age = f"{Age} {sex}"
                VS_number = ''.join(filter(str.isdigit, str(VS_name.split('(')[1].split(')')[0])))
                VS_name = '(' + VS_name.split('(')[0] + '\n ' + VS_number + ')'
                R_number = ''.join(filter(str.isdigit, str(R_name.split('(')[1].split(')')[0])))
                R_name = '(' + R_name.split('(')[0] + '\n ' + R_number + ')'
                if bed_number is not None:
                    bed_number = bed_number.replace("－", "-")
                    bed_number = bed_number.replace(" ", "")

                

                
                Profile_list = [bed_number,
                                name,
                                ID_number,
                                Sex_age,
                                ]
                if VS_excluding == False:
                    Profile_list.append(VS_name)
                if R_excluding == False:
                    Profile_list.append(R_name)


                time.sleep(0.15)
                    

                click_specific_link(driver, By.XPATH, "//a[text()='累積報告»']", WAIT_TIMEOUT)
                click_specific_link(driver, By.XPATH, "//a[text()='(累積)CBC']", WAIT_TIMEOUT)
                select_option_from_dropdown(driver, By.ID, 'resdtmonth', 
                                            'visible_text', "二週內", 
                                            wait_timeout = WAIT_TIMEOUT)
                time.sleep(random.uniform(0.2, 0.25))
                click_specific_link(driver, By.LINK_TEXT, "WBC", WAIT_TIMEOUT)
                df = extract_table_data(driver, 
                    table_locator=(By.ID, "resdtable"), 
                    wait_timeout=WAIT_TIMEOUT)
                
        
                df = process_lab_data(df)
                
                CBCDC = None
                Plt1 = None
                Plt2 = None
                Plt = None
                SEG_new = None
                if str(df.iloc[1]['WBC']) != 'nan' and str(df.iloc[1]['WBC']) != 'None':
                    CBC = apply_conditions_to_dataframe(df, conditions_dict={
                    'WBC': (3000, 11000, True), 'HGB': (11, 18, True)})
                    SEG_new, SEG, date2 = process_data_with_date_conversion(df, column_name='SEG') 
                    print(SEG_new)
                    CBC[0] = CBC[0] + "(" + str(SEG_new) + ")" if SEG_new is not None else CBC[0]
                    
                    Plt1 = str(float(df.loc[0]['PLT'])/10000) + 'w'
                    Plt2 = str(float(df.loc[1]['PLT'])/10000) + 'w'
                    Plt = Plt2
                    try:
                        if float(df.loc[1]['PLT'])/10000 < 10:
                            Plt = Plt1 + '>' + Plt2
                    except:
                        pass
                    formatted_date=""
                    try:
                        parsed_date = datetime.strptime(str(df.iloc[1]['日期']), "%Y/%m/%d")
                        formatted_date = "(" + str(parsed_date.strftime("%#m/%#d")) + ")"
                    except:
                        print("")
                    CBC.append(Plt)
                    INRdimer = apply_conditions_to_dataframe(df, conditions_dict={
                    'INR': (-100,-99, False),
                    'Dd': (0,1, False),
                    'fib': (-100, -99, False),})
                    filtered_INRdimer = [item for item in INRdimer if not pd.isna(item) and item is not None]
                    filtered_INRdimer = ' '.join(filtered_INRdimer)
                    CBC = '/'.join(CBC)
                    CBCDC = ' '.join(item for item in [CBC, filtered_INRdimer, formatted_date] if not pd.isna(item) and item is not None and item != '')
                    #CBCDC = ' '.join([CBC, filtered_INRdimer, formatted_date])
                    try:
                        CBCDC = CBCDC.replace('WBC: ', '').replace('HGB: ', '')
                        print(CBCDC)
                    except:
                        pass

                
                select_option_from_dropdown(driver, By.ID, 'resdtype',           # SMAC
                                            'visible_text', "(累積)SMAC", 
                                            wait_timeout = WAIT_TIMEOUT) 
                time.sleep(random.uniform(0.2, 0.3))
                click_specific_link(driver, By.LINK_TEXT, "ALB", WAIT_TIMEOUT)
                

                
                df = extract_table_data(driver, 
                    table_locator=(By.ID, "resdtable"), 
                    wait_timeout=WAIT_TIMEOUT)
                
                df = process_lab_data(df)
                print(df)
                formatted_date=""
                try:
                    parsed_date = datetime.strptime(str(df.iloc[1]['日期']), "%Y/%m/%d")
                    formatted_date = "(" + str(parsed_date.strftime("%#m/%#d")) + ")"
                except:
                    print("")
            
                SMAC = apply_conditions_to_dataframe(df, conditions_dict={
                'CRP': (0, 1, True),
                'PCT': (0,0.2, True),
                'LDH': (0,250, True),
                'UA': (0,6, True),
                'Ferritin': (-100,-99, False)
                })
                filtered_list = [item for item in SMAC if not pd.isna(item) and item is not None]
                SMAC_string1 = ' '.join(filtered_list)
                SMAC = apply_conditions_to_dataframe(df, conditions_dict={
                'Na': (133, 147, True),      
                'K': (3, 5.1, True),       
                'Ca': (-100,-99, False),    
                'fCa': (1.1,1.35, False),      
                'IP': (-100,-99, False),       
                'Mg': (-100,-99, False),     
                'CO2':(20,28,True),
                'BUN': (0, 30, True),           
                'Cr': (0, 1.2, True),        
                'eGFR': (-100,-99, False), 
                'ALT': (0, 50, True), 
                'AST': (0, 50, True), 
                'Tbili': (0, 1.2, True), 
                'ALKP': (0, 200, True), 
                'Alb': (3.5, 100, True), 
                'NH3': (0, 50, True), 
                'Lipase': (0,130, True), 
                'Amylase': (0,130, True),    
                'CK': (0, 200, True), 
                'TROP': (0, 0.11, True), 
                'BNP': (0,250, False), 
                'Lac': (0,16, True),    
                })
                select_option_from_dropdown(driver, By.ID, 'resdtype',           # SMAC
                                            'visible_text', "(累積)床邊血糖", 
                                            wait_timeout = WAIT_TIMEOUT)  
                select_option_from_dropdown(driver, By.ID, 'resdtmonth', 
                                            'visible_text', "一週內", 
                                            wait_timeout = WAIT_TIMEOUT)
                time.sleep(random.uniform(0.2, 0.4))
                click_specific_link(driver, By.LINK_TEXT, "Glucose", WAIT_TIMEOUT)
                
                df = extract_table_data(driver, 
                    table_locator=(By.ID, "resdtable"), 
                    wait_timeout=WAIT_TIMEOUT)
                # 累計血糖
                Glucose = process_glucose_data(df, glu_number, False) # glucose

                filtered_list = [item for item in SMAC if not pd.isna(item) and item is not None] 
                SMAC_string2 = ' '.join(filtered_list)
                SMAC_list = [SMAC_string1, SMAC_string2]
                SMAC_string = '\n'.join(item for item in SMAC_list if not pd.isna(item) and item is not None and item != '')
                
                Thyroid_str=''
                if department in ['ALL', 'META', 'MO', 'GI', 'CM']:
                    select_option_from_dropdown(driver, By.ID, 'resdtype',           # SMAC
                                                'visible_text', "(累積)甲狀腺報告", 
                                                wait_timeout = WAIT_TIMEOUT) 
                    select_option_from_dropdown(driver, By.ID, 'resdtmonth', 
                                                'visible_text', "二週內", 
                                                wait_timeout = WAIT_TIMEOUT)
                    time.sleep(random.uniform(0.2, 0.4))
                    click_specific_link(driver, By.LINK_TEXT, "TSH", WAIT_TIMEOUT)
                    df = extract_table_data(driver, 
                        table_locator=(By.ID, "resdtable"), 
                        wait_timeout=WAIT_TIMEOUT)
                    df = process_lab_data(df)
                    Thyroid = apply_conditions_to_dataframe(df, conditions_dict={
                        'TSH': (-100,-99, False), 'fT4': (-100,-99,False)})
                    Thyroid_list = [item for item in Thyroid if not pd.isna(item) and item is not None]
                    Thyroid_str = ' '.join(Thyroid_list)


                if not pd.isna(Glucose) and Glucose is not None:
                    SMAC_string = ' '.join([SMAC_string, Glucose])
                if Thyroid_str.strip() != '':
                    SMAC_string = ' '.join([SMAC_string, Thyroid_str])       

                
                SMAC_string = ' '.join([SMAC_string,formatted_date])

                if department in ['ICU']:
                    SMAC_string = SMAC_string.replace(" ALT", "\nALT").replace(" CK", "\nCK")
                
                print(SMAC_string)
                if include_gas:
                    select_option_from_dropdown(driver, By.ID, 'resdtype',           # SMAC
                                                'visible_text', "(累積)BloodGas", 
                                                wait_timeout = WAIT_TIMEOUT)  
                    select_option_from_dropdown(driver, By.ID, 'resdtmonth', 
                                                'visible_text', "一週內", 
                                                wait_timeout = WAIT_TIMEOUT)
                    time.sleep(random.uniform(0.2, 0.4))
                    click_specific_link(driver, By.LINK_TEXT, "FIO2", WAIT_TIMEOUT)
                    df = extract_table_data(driver, 
                        table_locator=(By.ID, "resdtable"), 
                        wait_timeout=WAIT_TIMEOUT)
                    df = calculate_pf_ratio(df)
                    df = process_lab_data(df)
                    formatted_date=""
                    try:
                        parsed_date = datetime.strptime(str(df.iloc[1]['日期']), "%Y/%m/%d")
                        formatted_date = "(" + str(parsed_date.strftime("%#m/%#d")) + ")"
                    except:
                        print("")
            
                    
                    print(df)
                    bloodgas = apply_conditions_to_dataframe(df, conditions_dict={
                    'FIO2': (-100, -99, False),
                    'PH': (7.35,7.45,True),
                    'PO2': (-100, -99, False),
                    'PCO2': (-100, -99, False),
                    'HCO3': (20, 28, True),
                    'BE': (-100, -99, False),
                    'P/F': (-100, -99, True)})
                    print(bloodgas)
                    bloodgas.append(formatted_date)
                    filtered_list = [item for item in bloodgas if not pd.isna(item) and item is not None]
                    bloodgas_string = ' '.join(filtered_list)

                if include_tumormarker:
                    select_option_from_dropdown(driver, By.ID, 'resdtype',           
                                                'visible_text', "(累積)腫瘤指標", 
                                                wait_timeout = WAIT_TIMEOUT)  
                    select_option_from_dropdown(driver, By.ID, 'resdtmonth', 
                                                'visible_text', "三個月內", 
                                                wait_timeout = WAIT_TIMEOUT) 
                    time.sleep(random.uniform(0.2, 0.4))
                    click_specific_link(driver, By.LINK_TEXT, "AFP", WAIT_TIMEOUT)
                    df = extract_table_data(driver, 
                        table_locator=(By.ID, "resdtable"), 
                        wait_timeout=WAIT_TIMEOUT)
                    df = process_lab_data(df)
                    print(df)
                    Tumor_marker = []
                    if department in ['HEMA']:
                        Tumor_marker = apply_conditions_to_dataframe(df, conditions_dict={
                            'B2M': (0, 2500, True), 
                            'NSE': (0, 100000, False)})  
                        
                    if department in ['ALL', 'MO', 'GI', 'CM']:
                        Tumor_marker = apply_conditions_to_dataframe(df, conditions_dict={
                            'CEA': (0, 5, True),
                            'CA199': (0, 37, True),
                            'CA153': (0, 30, True),
                            'CA125': (0, 35, True),
                            'AFP': (0, 10, True),
                            'PSA': (0, 4, True),
                            'hCG': (0, 5, True),
                            'Free PSA': (0, 0.25, True),
                            'Ca72_4': (0, 6.9, True),
                            'B2M': (0, 2500, True), 
                            'NSE': (0, 100000, False),
                            'SCC': (0, 2, True),
                            'CYFRA21-1': (0, 3.3, True),
                            'NMP22': (0, 10, True),})  
                    
                    filtered_list = []
                    filtered_list = [item for item in Tumor_marker if not pd.isna(item) and item is not None]
                    # Join the list items into a string with comma and space separation
                    Tumor_marker_date = None
                    if filtered_list != []:
                        try: 
                            Tumor_marker_date = df.iloc[0]['日期']
                            parts = Tumor_marker_date.split('/')
                            month = parts[1] 
                            day = parts[2]  
                            formatted_month = month.lstrip('0') # '8'
                            formatted_day = day.lstrip('0')     # '8'

                            # 4. 用 '/' 將處理過的月份和日期組合起來
                            Tumor_marker_date = f"({formatted_month}/{formatted_day})"
                            filtered_list.append(Tumor_marker_date)
                            Tumor_marker_string = ' '.join(filtered_list)
                        except:
                            print("faile tumor marker")

                if department in ['ALL', 'HEMA', 'ICU', 'MO', 'AIR']:
                    IgG_list = None
                    kappa_lambda = None
                    select_option_from_dropdown(driver, By.ID, 'resdtype',           
                                                'visible_text', "(累積)一般生化", 
                                                wait_timeout = WAIT_TIMEOUT)  
                    select_option_from_dropdown(driver, By.ID, 'resdtmonth', 
                                                'visible_text', "一個月內", 
                                                wait_timeout = WAIT_TIMEOUT) 
                    time.sleep(random.uniform(0.2, 0.4))
                    click_specific_link(driver, By.LINK_TEXT, "IG G", WAIT_TIMEOUT)
                    df = extract_table_data(driver, 
                        table_locator=(By.ID, "resdtable"), 
                        wait_timeout=WAIT_TIMEOUT)
                    print(df)
                    IgG_new, IgG_old, date = process_data_with_date_conversion(df, column_name='IG G')  
                    formatted_date = format_date_with_parentheses(date)
                    IgG_list = format_comparison([IgG_old, IgG_new, formatted_date],0)
                    if IgG_list is not None:
                        IgG_list = 'IgG: ' + IgG_list
                    print(IgG_list)
                    time.sleep(0.1+random.uniform(0.1, 0.2))
                    select_option_from_dropdown(driver, By.ID, 'resdtype',           
                                                'visible_text', "(累積)SMAC", 
                                                wait_timeout = WAIT_TIMEOUT)  
                    select_option_from_dropdown(driver, By.ID, 'resdtmonth', 
                                                'visible_text', "一個月內", 
                                                wait_timeout = WAIT_TIMEOUT) 
                    time.sleep(0.1 + random.uniform(0.1, 0.2))
                    click_specific_link(driver, By.LINK_TEXT, "ALB", WAIT_TIMEOUT)
                    df = extract_table_data(driver, 
                        table_locator=(By.ID, "resdtable"), 
                        wait_timeout=WAIT_TIMEOUT)
                    print(df)
                    kl_new, kl_old, date = process_data_with_date_conversion(df, column_name='kappa/lambda')  
                    formatted_date = format_date_with_parentheses(date)
                    kappa_lambda = format_comparison([kl_old, kl_new,  formatted_date])
                    if kappa_lambda is not None:
                        kappa_lambda = 'κ/λ: ' + kappa_lambda
                    print(kappa_lambda)         
                    filtered_IgG = [item for item in [Tumor_marker_string, IgG_list, kappa_lambda] if item is not None]
                    if filtered_IgG != []:
                        Tumor_marker_string = ' '.join(filtered_IgG)
                    else:
                        Tumor_marker_string = None
            
                if department in ['ALL', 'HEMA', 'ICU', 'MO', 'INF', 'GI', 'AIR', 'CM']:
                    CMV_list = None
                    date = None
                    CMV_new = None
                    CMV_old = None

                    select_option_from_dropdown(driver, By.ID, 'resdtype',           
                                                'visible_text', "(累積)伺機感染", 
                                                wait_timeout = WAIT_TIMEOUT)  
                    select_option_from_dropdown(driver, By.ID, 'resdtmonth', 
                                                'visible_text', "一個月內", 
                                                wait_timeout = WAIT_TIMEOUT) 
                    time.sleep(random.uniform(0.2, 0.4))
                    
                    df = extract_table_data(driver, 
                        table_locator=(By.ID, "resdtable"), 
                        wait_timeout=WAIT_TIMEOUT)
                    print(df)
                    df = df.drop(df.index[-1])
                    CMV_new, CMV_old, date = process_data_with_date_conversion(df, column_name='CMV_REALTIME')  
                    formatted_date = format_date_with_parentheses(date)
                    try:
                        CMV_old = int(CMV_old)
                    except:
                        pass
                    try:
                        CMV_new = int(CMV_new)
                    except:
                        pass


                    if CMV_old == 1:
                        CMV_old = '(-)' 
                    if CMV_old == 34:
                        CMV_old = 'low'
                    if CMV_new == 1:
                        CMV_new = '(-)' 
                    if CMV_new == 34:   
                        CMV_new = 'low'                        
                    #CMV_list = format_comparison([CMV_old, CMV_new, formatted_date],0)

                    if CMV_new is not None:
                        CMV_list = 'CMV: ' + str(CMV_old) + '>' + str(CMV_new) + ' ' + str(formatted_date) 

                    if CMV_new == '(-)' or CMV_new == 'low' :
                        CMV_list = 'CMV: ' + str(CMV_new) + ' ' + str(formatted_date)     

                    print(CMV_list)
                    
                    time.sleep(0.1 + random.uniform(0.1, 0.2))                    

                if body_weight==True:                        # body weight       
                    click_specific_link(driver, By.XPATH, "//a[text()='生命徵象']", WAIT_TIMEOUT) 
                    click_specific_link(driver, By.XPATH, "//input[@value='查詢']", WAIT_TIMEOUT)
                    click_specific_link(driver, By.XPATH, "//th[contains(.,'日期時間')]", WAIT_TIMEOUT)   
                    df = extract_table_data(driver, 
                        table_locator = (By.XPATH, "//table[@style='text-align: center;font-size: 12px;font-family: verdana;background: #c0c0c0;border-left: 1px solid #888888;border-bottom: 1px solid #888888;width:700px']"), 
                        wait_timeout=WAIT_TIMEOUT)
                    print(df)
                    try:
                        if department in ['ICU']:
                            bodyweight = extract_height_weight_trends_from_clipboard(df, height=True, weight=True, bsa=False, ibw=True, gender=sex)
                        elif department in ['ALL']:   
                            bodyweight = extract_height_weight_trends_from_clipboard(df, height=True, weight=True, bsa=True, ibw=True, gender=sex)
                        else:
                            bodyweight = extract_height_weight_trends_from_clipboard(df, gender=sex)
                    except:
                        print("Failed to extract body weight trends")

                if include_anti:                             # culture and antibiotics
                    try:
                        click_specific_link(driver, By.XPATH, "//a[text()='用藥紀錄»']", WAIT_TIMEOUT) 
                        click_specific_link(driver, By.LINK_TEXT, "感染抗生素資訊", WAIT_TIMEOUT)
                        culture_df = extract_table_data(driver, 
                            table_locator=(By.XPATH, "//table[@id='resinf01' and caption='一年內陽性培養結果累積報告']"), 
                            wait_timeout=WAIT_TIMEOUT)
                        print('culture:::',culture_df)
                        culture_str=get_recent_culture_results_string(culture_df)
                        antibiotics_df = extract_table_data(driver, 
                            table_locator = (By.ID, "resinf03"), 
                            wait_timeout=WAIT_TIMEOUT)
                        print('anti:::', antibiotics_df)
                        antibiotic_str=get_active_antibiotics(antibiotics_df, anti_days)            
                        print('anti:::', antibiotic_str)     
                    except:
                        print("failing extracting culture and antibiotic data")
                
                print(culture_str)

                culture_str = [item for item in [CMV_list, culture_str] if item is not None]
                if culture_str != []:
                    culture_str = ' '.join(culture_str)
                else:
                    culture_str = None
                result = []
                if diet_str is not None:
                    bodyweight = bodyweight + ', ' + diet_str
                if foley_lines is not None:
                    bodyweight = bodyweight + ', ' + foley_lines
                result = [CBCDC, SMAC_string, bloodgas_string, Tumor_marker_string, bodyweight, IO_str, culture_str, antibiotic_str,'    ']
                print(result)
                result = [item for item in result if not pd.isna(item) and item is not None and item != '' and item != ' ']
                result = "\n".join(result)
                
                drug_str_origin = None
                drug_str = None
                admission_time = None
                if include_drug==True:           # drug
                    click_specific_link(driver, By.XPATH, "//a[text()='用藥紀錄»']", WAIT_TIMEOUT) 
                    click_specific_link(driver, By.LINK_TEXT, "用藥紀錄", WAIT_TIMEOUT)
                    admission_time = find_and_click_first_inpatient_date_link(driver, inpatient_text="住院", wait_timeout=WAIT_TIMEOUT)
                    df = extract_table_data(driver, 
                        table_locator = (By.ID, "udorder"), 
                        wait_timeout=WAIT_TIMEOUT)
                    print("admission time:",admission_time)
                    print(df)
                    drug_str = extract_active_meds(df)
                    
                    if drug_formating:
                        drug_str = format_medications_to_columns(drug_str)
                    if date_chemo_dose_string is not None:
                        drug_str = '\n'.join([date_chemo_dose_string, drug_str, '     '])
                    print(drug_str)
                    current = move_to_drug_col_paste(current, text = drug_str)

                try:
                    admission_time = datetime.strptime(admission_time, '%Y%m%d')
                    admission_time = f"{admission_time.month}/{admission_time.day}入"
                    if Critical == True:
                        admission_time = '@' + admission_time 
                
                    
                    
                except:
                    admission_time = None
                    print('no admission_time')
                    pass

                days_diff = None
                if department in ['ICU']:
                    try:
                        
                        click_specific_link(driver, By.XPATH, "//a[text()='住院資訊»']", WAIT_TIMEOUT) 
                        click_specific_link(driver, By.LINK_TEXT, "轉床轉科", WAIT_TIMEOUT)
                        time.sleep(0.3)
                        df = extract_table_data(driver, table_locator = (By.ID, "plocslist"), wait_timeout=WAIT_TIMEOUT)
                        if df.iloc[0,2] is not None and ('ICU' in str(df.iloc[0,2]) or 'CCU' in str(df.iloc[0,2]) or 'RCU' in str(df.iloc[0,2])):
                        
                            dt = pd.to_datetime(df.iloc[0,0], errors='coerce')
                            days_diff = (datetime.now().date() - dt.date()).days if not pd.isna(dt) else None
                            print(f"日期: {dt}, 與今天相差天數: {days_diff}")
                    except Exception as e:
                        print(f"日期轉換或計算天數差時發生錯誤: {e}")
                        print(days_diff)
                
                try:
                    admission_time = f"ICU day {days_diff}" if days_diff is not None and days_diff >=0 else admission_time
                   
                except:
                    pass                


                current = move_to_row_start(current, column=2)
                if admission_time:
                    Profile_list.append(admission_time)
                current.value = "\n".join(Profile_list)
                current = move_to_lab_col_paste(current, text=result)
                current = move_to_row_start(current, column=3)                    
                if current.value is None or ai_asking:
                    try:     
                        click_specific_link(driver, By.XPATH, "//a[text()='病程紀錄']", WAIT_TIMEOUT)
                        click_specific_link(driver, By.XPATH, '//a[@title="prgdetail"]', 2)
                        progress_note = scrape_table_text_with_newlines(driver, wait_timeout=2, table_locator_type=By.XPATH, table_locator_value="//table[@cellspacing='0' and @cellpadding='0' and @width='100%']")
                          
                        print(progress_note)
                        assessment = extract_assessment_between_markers(progress_note)
                        print('assessment:',assessment)
                        if current.value is None:
                            current.value = assessment
                    except:
                        pass
                
                if ai_asking:
                    lab_ai = []
                    
                    print('======================================')
                    lab_ai = [CBCDC, SMAC_string, bloodgas_string, Tumor_marker_string]
                    lab_ai = [item for item in lab_ai if not pd.isna(item) and item is not None and item != '' and item != ' ']
                    lab_ai = '\n'.join(lab_ai)
                    print(lab_ai)
                    ai_prompt = f'''# Role:
                        你是一位經驗豐富的「資深臨床醫師」與「資深臨床藥師」。你具備臨床醫學、藥物動力學、抗生素管理導向（Antimicrobial Stewardship）以及臨床營養學的專業知識。

                        # Objective:
                        請根據我提供的病人資料（Patient Profile），進行全面的臨床評估。你的目標是找出目前治療計畫中的潛在風險、疏漏或需要優化的部分。

                        # Context & Constraints:
                        - 病人可能有多重共病且正在接受化療或抗生素治療。
                        - 必須特別注意「肝腎功能（根據抽血數值）」對「目前用藥劑量」的影響。
                        - 必須評估「抗生素」是否符合「感染菌種」的敏感性（若無藥敏資料請依經驗法則判斷）。
                        - 必須評估「營養狀態」（根據飲食與體位變化）。
                        - **重要免責聲明：你的回答僅供醫療人員參考，不作為最終醫療指令。**

                        # Input Data:
                        請分析以下病人數據：

                        1. **基本資料**：
                        - 年齡/性別：{Sex_age}
                        - 住院天數：{admission_time}
                        - 入院身高/前次身高/這次身高(或體重變化)：{bodyweight}
                        2. **臨床診斷**：{assessment}
                        3. **客觀數據**：
                        - 關鍵抽血數值 (包含腎功能Cr/eGFR, 肝功能AST/ALT, 白血球WBC/ANC, 血紅素Hb, 電解質等)：{lab_ai}
                        - 感染控制 (菌種/培養結果)：{culture_str}
                        4. **治療現況**：
                        - 昨天輸入輸出量:{IO_str}
                        - 飲食內容：{diet_str}
                        - 管路留置 (CVC/Foley/NG etc.)：{foley_lines}
                        - 目前使用抗生素：{antibiotic_str}
                        - 最近化療藥物 (Regimen/Date)：{date_chemo_dose_string}
                        - 目前其他用藥及劑量：{drug_str_origin}

                        # Tasks & Output Format:
                        請依照以下四個維度進行分析，並以結構化的表格或條列式輸出：

                        ## 1. ⚠️ 應注意 (Attention/Caution)
                        - **交互作用**：檢查目前藥物（含化療藥、抗生素）是否有嚴重交互作用（Drug-Drug Interactions）、需要預防HBV或PJP infection等
                        - **劑量調整**：根據病人的腎功能（eGFR）或肝功能，目前的藥物劑量是否過高或需要調整？
                        - **副作用監測**：針對最近的化療藥物或高風險藥物，目前抽血數值是否有出現預期外的毒性（如骨髓抑制、電解質失衡）？

                        ## 2. ➕ 建議新增 (Addition)
                        - **支持性療法**：基於病人的症狀或化療副作用，是否缺少止吐、止痛、胃藥或電解質補充？
                        - **預防性投藥**：例如是否需要預防肺囊蟲、B肝復發或血栓預防？
                        - **營養介入**：若身高/體重有顯著變化或飲食攝取不足，是否建議新增營養補充品或會診營養師？

                        ## 3. 🔍 潛在缺漏 (Missing/Omission)
                        - **標準治療**：對照主要診斷，是否有符合 Guideline 的標準用藥未被開立？
                        - **管路照護**：若有管路且有感染跡象，是否缺少管路更換或移除的評估？
                        - **檢查缺漏**：基於目前病況，建議追蹤的檢查（如藥物濃度監測 TDM、後續細菌培養）？

                        ## 4. 🔄 建議更改 (Modification)
                        - **抗生素降階/升階**：根據「感染菌種」與「目前抗生素」，判斷是否無效（需改藥）或太強（需降階）？
                        - **給藥途徑**：若病人飲食正常，是否建議將靜脈注射（IV）藥物改為口服（PO）？
                        - **重複用藥**：是否有藥理機轉重複的藥物同時使用？

                        # Reasoning Process:
                        在給出建議前，請先在內心一步步檢視：
                        1. 計算病人的腎功能分期。
                        2. 確認抗生素是否覆蓋檢出的菌種。
                        3. 掃描所有藥物對照 Lab data 的禁忌症。

                        ### 請把所有建議依照重要性條列式一起列出!!
                        ### 5. 最後把重要內容簡短總結
                        請開始分析。
                        '''
                    print(ai_prompt)
                    current = move_to_row_start(current, column=13)
                    current.value = ai_prompt
                    current.api.WrapText = False


                current = move_to_row_start(current, column=1, down=1)
            
                print(str(current.value))
                if current.value == None:
                    global stop_requested
                    stop_requested = True
                    stop_the_code()

                ID_number = extract_and_round_number(current.value)
                print(ID_number)

                patient = len(ID_number)

                
                click_specific_link(driver, By.LINK_TEXT, "回查詢", WAIT_TIMEOUT, main_window_handle, True, 1.3) 
                failure_count = 0

            except: #如果中途遇到阻礙，重新再來一次
                #找不到病歷號就停止
                current = move_to_row_start(current, column=1, down=0)

                print(ID_number, len(ID_number))
                if len(ID_number) < 4: 
                    break
                
                failure_count =failure_count+1
      
                if failure_count>2:
                    print("連續錯誤超過2次，結束程式")
                    break
                # 關閉所有新開的視窗，只保留原始主視窗
                try:
                    current_handles = driver.window_handles
                    for handle in current_handles:
                        if handle != main_window_handle:
                            try:
                                driver.switch_to.window(handle)
                                driver.close()
                            except Exception as close_e:
                                print(f"Error closing window handle {handle}: {close_e}")
                    driver.switch_to.window(main_window_handle)
                except Exception as e:
                    print(f"Error during window cleanup: {e}")
                try:
                    click_specific_link(driver, By.LINK_TEXT, "回查詢", WAIT_TIMEOUT, main_window_handle, True, 1.3) 
                except:
                    pass
                




    
    except SystemExit:
        # 當按下 Esc 時，stop_the_code() 會引發 SystemExit
        print("偵測到使用者請求退出...")
    except Exception as e:
        # 捕捉其他所有可能的錯誤
        print(f"程式運行時發生未預期的錯誤: {e}")
    finally:
        # --- 無論如何，這裡的程式碼都一定會被執行 ---
        print("\n程式即將結束，執行最終清理...")
        if drug_formating:
            ws.range('E:E').api.Font.Name = 'Consolas'
        ws.range('A2:A25').rows.autofit()
        safe_exit(driver) # 傳入局部的 driver 變數進行清理


if __name__ == "__main__":
    main()

