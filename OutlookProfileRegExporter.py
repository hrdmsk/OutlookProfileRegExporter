#Outlook Profile Reg Exporter.py
# Outlookのアカウント設定（プロファイル）をレジストリから読み込み、
# ユーザーが選択したプロファイルを個別または一括でエクスポートするツール

# --- ライブラリのインポート ---
import tkinter as tk
from tkinter import ttk  # テーマ付きウィジェット(Radiobutton, Scrollbar)を使用
from tkinter import filedialog, messagebox
import subprocess
import os
import winreg
import ctypes
import sys

# --- グローバル定数 ---
ICON_BASE64 = """
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
"""

OUTLOOK_PROFILE_KEYS = {
    "Outlook 2016/2019/365": r"Software\Microsoft\Office\16.0\Outlook\Profiles",
    "Outlook 2013": r"Software\Microsoft\Office\15.0\Outlook\Profiles",
    "Outlook 2010": r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles",
}

# --- 関数定義 ---

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def get_outlook_profiles_base_key():
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
    """単一のプロファイルをエクスポートする"""
    full_reg_path = f"HKEY_CURRENT_USER\\{base_path}\\{profile_name}"
    output_file = os.path.join(output_dir, f"OutlookProfile_{profile_name}.reg")
    try:
        command = ['regedit', '/e', output_file, full_reg_path]
        subprocess.run(command, check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        return True, output_file
    except Exception as e:
        print(f"'{profile_name}'のエクスポートに失敗: {e}")
        return False, str(e)

def export_all_profiles_to_single_file(output_dir, base_path):
    """全てのプロファイルを単一のファイルにエクスポートする"""
    full_reg_path = f"HKEY_CURRENT_USER\\{base_path}"
    output_file = os.path.join(output_dir, "AllOutlookProfiles.reg")
    try:
        command = ['regedit', '/e', output_file, full_reg_path]
        subprocess.run(command, check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        return True, output_file
    except Exception as e:
        print(f"一括エクスポートに失敗: {e}")
        return False, str(e)

# --- GUIアプリケーションのクラス定義 ---

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Outlook Profile Reg Exporter")
        
        self.master.geometry("450x320")
        self.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.profiles = []
        self.load_profiles()
        self.create_widgets()
        self.toggle_mode()

    def load_profiles(self):
        base_key = get_outlook_profiles_base_key()
        if base_key:
            self.profiles = get_available_profiles(base_key)

    def create_widgets(self):
        # --- モード選択 ---
        mode_frame = ttk.LabelFrame(self, text="Export Mode")
        mode_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        self.export_mode_var = tk.StringVar(value="individual")
        
        ttk.Radiobutton(mode_frame, text="個別エクスポート", variable=self.export_mode_var, value="individual", command=self.toggle_mode).pack(side="left", padx=10, pady=5)
        ttk.Radiobutton(mode_frame, text="一括エクスポート", variable=self.export_mode_var, value="bulk", command=self.toggle_mode).pack(side="left", padx=10, pady=5)

        # --- プロファイル選択 (リストボックスに変更) ---
        tk.Label(self, text="エクスポートするプロファイル (複数選択可):").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 2))
        
        list_frame = tk.Frame(self)
        list_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.profile_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, exportselection=False)
        self.profile_listbox.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.profile_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.profile_listbox.config(yscrollcommand=scrollbar.set)

        if self.profiles:
            for profile in self.profiles:
                self.profile_listbox.insert(tk.END, profile)
        else:
            self.profile_listbox.insert(tk.END, "利用可能なプロファイルがありません")
            self.profile_listbox.config(state="disabled")

        # --- 出力先フォルダ ---
        tk.Label(self, text="出力先フォルダ:").grid(row=3, column=0, columnspan=2, sticky="w", pady=(5, 2))
        
        self.folder_path_var = tk.StringVar()
        # デフォルトのパスとして、デスクトップ上の「OutlookReg」フォルダを設定
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'OutlookReg')
        self.folder_path_var.set(desktop_path)
        
        path_frame = tk.Frame(self)
        path_frame.grid(row=4, column=0, columnspan=2, sticky="ew")
        path_frame.grid_columnconfigure(0, weight=1)

        self.path_entry = tk.Entry(path_frame, textvariable=self.folder_path_var)
        self.path_entry.grid(row=0, column=0, sticky="ew")

        self.browse_button = tk.Button(path_frame, text="参照...", command=self.browse_folder)
        self.browse_button.grid(row=0, column=1, padx=5)
        
        # --- 実行ボタン ---
        self.export_button = tk.Button(self, text="エクスポート実行", command=self.run_export, bg="#007bff", fg="white", height=2)
        self.export_button.grid(row=5, column=0, columnspan=2, pady=10, sticky="ew")

        if not self.profiles:
            self.export_button.config(state="disabled")
        
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def toggle_mode(self):
        if self.export_mode_var.get() == "individual":
            self.profile_listbox.config(state="normal" if self.profiles else "disabled")
            self.export_button.config(text="選択したプロファイルをエクスポート")
        else: # "bulk"
            self.profile_listbox.config(state="disabled")
            self.export_button.config(text=f"すべてのプロファイル ({len(self.profiles)}件) を1つのファイルにエクスポート")

    def browse_folder(self):
        foldername = filedialog.askdirectory()
        if foldername:
            self.folder_path_var.set(foldername)

    def run_export(self):
        output_dir = self.folder_path_var.get()
        
        # 出力先フォルダが存在しない場合は作成する
        try:
            os.makedirs(output_dir, exist_ok=True)
        except OSError as e:
            messagebox.showerror("エラー", f"出力先フォルダの作成に失敗しました。\n{e}")
            return
        
        base_path = get_outlook_profiles_base_key()
        if not base_path:
            messagebox.showerror("エラー", "Outlookのプロファイルパスが見つかりませんでした。")
            return

        mode = self.export_mode_var.get()
        if mode == "individual":
            selected_indices = self.profile_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("警告", "エクスポートするプロファイルを1つ以上選択してください。")
                return
            
            selected_profiles = [self.profile_listbox.get(i) for i in selected_indices]
            success_count, fail_count = 0, 0
            for profile_name in selected_profiles:
                success, _ = export_profile(output_dir, profile_name, base_path)
                if success: success_count += 1
                else: fail_count += 1
            
            messagebox.showinfo("個別エクスポート完了", f"処理が完了しました。\n\n成功: {success_count}件\n失敗: {fail_count}件")

        elif mode == "bulk":
            if not self.profiles:
                messagebox.showwarning("警告", "エクスポート対象のプロファイルがありません。")
                return
            
            success, result = export_all_profiles_to_single_file(output_dir, base_path)
            if success:
                messagebox.showinfo("成功", f"すべてのプロファイルが正常にエクスポートされました。\nファイル: {result}")
            else:
                messagebox.showerror("エクスポート失敗", f"一括エクスポートに失敗しました。\nエラー: {result}")

# --- メインの実行ブロック ---
if __name__ == "__main__":
    if is_admin():
        root = tk.Tk()
        app = Application(master=root)
        app.mainloop()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
