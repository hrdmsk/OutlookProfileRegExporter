#OutlookRegExporter.py
# Outlookのアカウント設定（プロファイル）をレジストリから読み込み、
# ユーザーが選択したプロファイルを個別または一括でエクスポートするツール

# --- ライブラリのインポート ---
import tkinter as tk
from tkinter import ttk  # テーマ付きウィジェット(Combobox, Radiobutton)を使用
from tkinter import filedialog, messagebox
import subprocess
import os
import winreg
import ctypes
import sys

# --- グローバル定数 ---
# Outlookのバージョンごとに異なるプロファイルのレジストリパスを定義
OUTLOOK_PROFILE_KEYS = {
    "Outlook 2016/2019/365": r"Software\Microsoft\Office\16.0\Outlook\Profiles",
    "Outlook 2013": r"Software\Microsoft\Office\15.0\Outlook\Profiles",
    "Outlook 2010": r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles",
}

# --- 関数定義 ---

def is_admin():
    """スクリプトが管理者権限で実行されているかを確認する。"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def get_outlook_profiles_base_key():
    """PCにインストールされているOutlookのProfilesキーのパスを返す。"""
    for version, key_path in OUTLOOK_PROFILE_KEYS.items():
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.CloseKey(key)
            print(f"検知したバージョン: {version}")
            return key_path
        except FileNotFoundError:
            continue
    return None

def get_available_profiles(base_key_path):
    """指定されたレジストリパス配下にあるプロファイル名の一覧を取得する。"""
    profiles = []
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_key_path)
        i = 0
        while True:
            try:
                profile_name = winreg.EnumKey(key, i)
                profiles.append(profile_name)
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
    except Exception as e:
        print(f"プロファイルの読み込み中にエラーが発生しました: {e}")
    return profiles

def export_profile(output_dir, profile_name, base_path):
    """指定された単一のプロファイル設定をエクスポートする内部関数。"""
    full_reg_path = f"HKEY_CURRENT_USER\\{base_path}\\{profile_name}"
    output_file = os.path.join(output_dir, f"OutlookProfile_{profile_name}.reg")
    try:
        command = ['regedit', '/e', output_file, full_reg_path]
        subprocess.run(command, check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        return True, output_file # 成功
    except Exception as e:
        print(f"'{profile_name}'のエクスポートに失敗: {e}")
        return False, str(e) # 失敗

# --- GUIアプリケーションのクラス定義 ---

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Outlook プロファイルエクスポートツール (管理者権限)")
        self.master.geometry("450x250") # ウィンドウサイズを縦に広げる
        self.pack(pady=10, padx=10, fill="x")
        
        self.profiles = []
        self.load_profiles()
        self.create_widgets()
        self.toggle_mode() # 初期状態を設定

    def load_profiles(self):
        """Outlookプロファイルをレジストリから読み込む"""
        base_key = get_outlook_profiles_base_key()
        if base_key:
            self.profiles = get_available_profiles(base_key)

    def create_widgets(self):
        # --- モード選択 ---
        mode_frame = ttk.LabelFrame(self, text="エクスポートモード")
        mode_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        self.export_mode_var = tk.StringVar(value="individual") # デフォルトは個別
        
        ttk.Radiobutton(mode_frame, text="個別エクスポート", variable=self.export_mode_var, value="individual", command=self.toggle_mode).pack(side="left", padx=10, pady=5)
        ttk.Radiobutton(mode_frame, text="一括エクスポート", variable=self.export_mode_var, value="bulk", command=self.toggle_mode).pack(side="left", padx=10, pady=5)

        # --- プロファイル選択 ---
        tk.Label(self, text="エクスポートするプロファイル:").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 2))
        
        self.profile_var = tk.StringVar()
        self.profile_combobox = ttk.Combobox(self, textvariable=self.profile_var, values=self.profiles, state="readonly")
        self.profile_combobox.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        if self.profiles:
            self.profile_combobox.current(0)
        else:
            self.profile_var.set("利用可能なプロファイルがありません")
            self.profile_combobox.config(state="disabled")

        # --- 出力先フォルダ ---
        tk.Label(self, text="出力先フォルダ:").grid(row=3, column=0, columnspan=2, sticky="w", pady=(5, 2))
        
        self.folder_path_var = tk.StringVar()
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        self.folder_path_var.set(desktop_path)
        
        self.path_entry = tk.Entry(self, textvariable=self.folder_path_var, width=50)
        self.path_entry.grid(row=4, column=0, sticky="ew")

        self.browse_button = tk.Button(self, text="参照...", command=self.browse_folder)
        self.browse_button.grid(row=4, column=1, padx=5)
        
        # --- 実行ボタン ---
        self.export_button = tk.Button(self, text="エクスポート実行", command=self.run_export, bg="#007bff", fg="white", height=2)
        self.export_button.grid(row=5, column=0, columnspan=2, pady=10, sticky="ew")

        if not self.profiles:
            self.export_button.config(state="disabled")

    def toggle_mode(self):
        """エクスポートモードに応じてUIの状態を切り替える"""
        if self.export_mode_var.get() == "individual":
            self.profile_combobox.config(state="readonly" if self.profiles else "disabled")
            self.export_button.config(text="選択したプロファイルをエクスポート")
        else: # "bulk"
            self.profile_combobox.config(state="disabled")
            self.export_button.config(text=f"すべてのプロファイル ({len(self.profiles)}件) をエクスポート")

    def browse_folder(self):
        foldername = filedialog.askdirectory()
        if foldername:
            self.folder_path_var.set(foldername)

    def run_export(self):
        output_dir = self.folder_path_var.get()
        if not os.path.isdir(output_dir):
            messagebox.showwarning("警告", "指定されたフォルダが存在しません。")
            return
        
        base_path = get_outlook_profiles_base_key()
        if not base_path:
            messagebox.showerror("エラー", "Outlookのプロファイルパスが見つかりませんでした。")
            return

        mode = self.export_mode_var.get()
        if mode == "individual":
            selected_profile = self.profile_var.get()
            if not selected_profile or selected_profile == "利用可能なプロファイルがありません":
                messagebox.showwarning("警告", "エクスポートするプロファイルを選択してください。")
                return
            
            success, result = export_profile(output_dir, selected_profile, base_path)
            if success:
                messagebox.showinfo("成功", f"プロファイル '{selected_profile}' が正常にエクスポートされました。\nファイル: {result}")
            else:
                messagebox.showerror("エクスポート失敗", f"プロファイルのエクスポートに失敗しました。\nエラー: {result}")
        
        elif mode == "bulk":
            if not self.profiles:
                messagebox.showwarning("警告", "エクスポート対象のプロファイルがありません。")
                return
            
            success_count = 0
            fail_count = 0
            for profile_name in self.profiles:
                success, _ = export_profile(output_dir, profile_name, base_path)
                if success:
                    success_count += 1
                else:
                    fail_count += 1
            
            messagebox.showinfo("一括エクスポート完了", f"処理が完了しました。\n\n成功: {success_count}件\n失敗: {fail_count}件")

# --- メインの実行ブロック ---
if __name__ == "__main__":
    if is_admin():
        root = tk.Tk()
        app = Application(master=root)
        app.mainloop()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
