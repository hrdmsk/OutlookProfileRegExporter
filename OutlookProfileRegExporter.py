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
