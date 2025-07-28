import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import winreg
import ctypes
import sys

# Outlookのバージョンごとのレジストリパス
OUTLOOK_PROFILE_KEYS = {
    "Outlook 2016/2019/365": r"Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook",
    "Outlook 2013": r"Software\Microsoft\Office\15.0\Outlook\Profiles\Outlook",
    "Outlook 2010": r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook",
}

def is_admin():
    """管理者権限で実行されているかを確認する"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def get_outlook_profile_key():
    """利用可能なOutlookプロファイルのレジストリキーパスを取得する"""
    for version, key_path in OUTLOOK_PROFILE_KEYS.items():
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.CloseKey(key)
            print(f"検知したバージョン: {version}")
            return key_path
        except FileNotFoundError:
            continue
    return None

def export_settings(output_dir):
    """指定されたディレクトリにOutlookの設定をエクスポートする"""
    reg_key_path = get_outlook_profile_key()

    if not reg_key_path:
        messagebox.showerror("エラー", "対応するOutlookプロファイルが見つかりませんでした。")
        return

    full_reg_path = f"HKEY_CURRENT_USER\\{reg_key_path}"
    output_file = os.path.join(output_dir, "OutlookAccount.reg")

    try:
        command = ['regedit', '/e', output_file, full_reg_path]
        result = subprocess.run(command, check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        messagebox.showinfo("成功", f"設定が正常にエクスポートされました。\nファイル: {output_file}")
    except FileNotFoundError:
        messagebox.showerror("エラー", "regedit.exeが見つかりません。Windows環境で実行してください。")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("エクスポート失敗", f"レジストリのエクスポートに失敗しました。\nエラー: {e.stderr}")
    except Exception as e:
        messagebox.showerror("予期せぬエラー", f"エラーが発生しました。\n{e}")

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        # 管理者権限で実行されていることをタイトルで示す
        self.master.title("Outlook 設定エクスポートツール (管理者権限)")
        self.master.geometry("450x150")
        self.pack(pady=10, padx=10)
        self.create_widgets()

    def create_widgets(self):
        self.folder_path_var = tk.StringVar()
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        self.folder_path_var.set(desktop_path)

        tk.Label(self, text="出力先フォルダ:").grid(row=0, column=0, sticky="w", pady=5)
        
        self.path_entry = tk.Entry(self, textvariable=self.folder_path_var, width=50)
        self.path_entry.grid(row=1, column=0, sticky="ew")

        self.browse_button = tk.Button(self, text="参照...", command=self.browse_folder)
        self.browse_button.grid(row=1, column=1, padx=5)
        
        self.export_button = tk.Button(self, text="設定をエクスポート", command=self.run_export, bg="#007bff", fg="white", height=2)
        self.export_button.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")

    def browse_folder(self):
        foldername = filedialog.askdirectory()
        if foldername:
            self.folder_path_var.set(foldername)

    def run_export(self):
        output_dir = self.folder_path_var.get()
        if not os.path.isdir(output_dir):
            messagebox.showwarning("警告", "指定されたフォルダが存在しません。")
            return
        export_settings(output_dir)

if __name__ == "__main__":
    if is_admin():
        # 管理者権限がある場合、GUIアプリケーションを起動
        root = tk.Tk()
        app = Application(master=root)
        app.mainloop()
    else:
        # 管理者権限がない場合、UACプロンプトを表示して自身を再実行
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)