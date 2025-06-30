import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import smtplib
import json
import os
import sys
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class RobocopyScheduler:
    def __init__(self, root):
        self.root = root
        self.root.title("Robocopy スケジューラ (Windows Task Scheduler)")
        self.root.geometry("800x950")  # 高さを少し増やす
        
        # 設定を保存するファイル名
        self.config_file = "robocopy_config.json"
        # タスク名を変数として管理（デフォルト値は後で設定）
        self.task_name_var = tk.StringVar(value="RobocopyBackupTask")
        
        # Robocopyオプション用の変数
        self.option_vars = {}
        # コピーモード用のラジオボタン変数を追加
        self.copy_mode_var = tk.StringVar(value="MIR")  # デフォルトは/MIR
        
        # 認証情報用の変数を追加
        self.source_auth_enabled_var = tk.BooleanVar()
        self.source_username_var = tk.StringVar()
        self.source_password_var = tk.StringVar()
        self.source_domain_var = tk.StringVar()
        
        self.dest_auth_enabled_var = tk.BooleanVar()
        self.dest_username_var = tk.StringVar()
        self.dest_password_var = tk.StringVar()
        self.dest_domain_var = tk.StringVar()
        
        # 認証ウィジェットのリストを初期化
        self.source_auth_widgets = []
        self.dest_auth_widgets = []
        
        self.create_widgets()
        self.load_config()
        self.update_task_status()
    def generate_task_name_from_dest(self):
        """コピー先フォルダからタスク名を自動生成"""
        dest_path = self.dest_var.get()
        if not dest_path:
            return "RobocopyBackupTask"
        
        try:
            # パスの最後の部分（フォルダ名）を取得
            folder_name = os.path.basename(dest_path.rstrip('\\'))
            
            # 空の場合はドライブ名などを使用
            if not folder_name:
                folder_name = dest_path.replace(':', '').replace('\\', '_')
            
            # タスク名に使用できない文字を除去・置換
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                folder_name = folder_name.replace(char, '_')
            
            # 空白を含む場合はアンダースコアに置換
            folder_name = folder_name.replace(' ', '_')
            
            # 最大長制限（Windowsタスク名は最大238文字）
            if len(folder_name) > 50:
                folder_name = folder_name[:50]
            
            return f"{folder_name}-Robocopy"
            
        except Exception as e:
            self.log_message(f"タスク名生成エラー: {str(e)}", "error")
            return "RobocopyBackupTask"

    def update_task_name_from_dest(self, *args):
        """コピー先フォルダ変更時にタスク名を自動更新"""
        if self.dest_var.get():
            new_task_name = self.generate_task_name_from_dest()
            self.task_name_var.set(new_task_name)
    
    def is_network_path(self, path):
        """ネットワークパスかどうかを判定"""
        return path.startswith("\\\\") if path else False
    
    def update_auth_state(self):
        """認証設定の有効/無効を更新"""
        # コピー元の認証設定
        source_path = self.source_var.get()
        if self.is_network_path(source_path):
            # ネットワークパスの場合は認証設定を有効化
            self.enable_auth_widgets(self.source_auth_widgets)
            self.source_auth_enabled_var.set(True)
        else:
            # ローカルパスの場合は認証設定を無効化
            self.disable_auth_widgets(self.source_auth_widgets)
            self.source_auth_enabled_var.set(False)
        
        # コピー先の認証設定
        dest_path = self.dest_var.get()
        if self.is_network_path(dest_path):
            # ネットワークパスの場合は認証設定を有効化
            self.enable_auth_widgets(self.dest_auth_widgets)
            self.dest_auth_enabled_var.set(True)
        else:
            # ローカルパスの場合は認証設定を無効化
            self.disable_auth_widgets(self.dest_auth_widgets)
            self.dest_auth_enabled_var.set(False)
    
    def enable_auth_widgets(self, widget_list):
        """認証設定ウィジェットを有効化"""
        for widget in widget_list:
            if hasattr(widget, 'configure'):
                widget.configure(state='normal')
    
    def disable_auth_widgets(self, widget_list):
        """認証設定ウィジェットを無効化"""
        for widget in widget_list:
            if hasattr(widget, 'configure'):
                widget.configure(state='disabled')
    
    def test_network_connection(self, path_type):
        """ネットワーク接続をテスト"""
        if path_type == "source":
            path = self.source_var.get()
            username = self.source_username_var.get()
            password = self.source_password_var.get()
            domain = self.source_domain_var.get()
        else:  # dest
            path = self.dest_var.get()
            username = self.dest_username_var.get()
            password = self.dest_password_var.get()
            domain = self.dest_domain_var.get()
        
        if not self.is_network_path(path):
            messagebox.showwarning("警告", "ネットワークパスが指定されていません")
            return
        
        if not username:
            messagebox.showerror("エラー", "ユーザー名を入力してください")
            return
        
        try:
            # ネットワークパスからサーバー部分を抽出
            server_path = "\\\\" + path.split("\\")[2]
            
            # 既存の接続を切断
            disconnect_cmd = f'net use "{server_path}" /delete /y'
            subprocess.run(disconnect_cmd, shell=True, capture_output=True)
            
            # 認証情報を構築
            if domain:
                user_part = f"{domain}\\{username}"
            else:
                user_part = username
            
            # net useコマンドで接続テスト
            connect_cmd = f'net use "{server_path}" /user:"{user_part}" "{password}"'
            
            result = subprocess.run(connect_cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            if result.returncode == 0:
                messagebox.showinfo("接続テスト", f"{path_type}への接続に成功しました")
                self.log_message(f"認証テスト成功: {path}")
                
                # テスト後は接続を切断
                subprocess.run(disconnect_cmd, shell=True, capture_output=True)
            else:
                error_msg = result.stderr.strip() if result.stderr else "不明なエラー"
                messagebox.showerror("接続テスト", f"接続に失敗しました:\n{error_msg}")
                self.log_message(f"認証テスト失敗: {path} - {error_msg}", "error")
                
        except Exception as e:
            messagebox.showerror("接続テスト", f"接続テストでエラーが発生しました:\n{str(e)}")
            self.log_message(f"認証テストエラー: {str(e)}", "error")
    
    def connect_network_path(self, path, username, password, domain=""):
        """ネットワークパスに認証情報で接続"""
        if not self.is_network_path(path):
            return True  # ローカルパスの場合は何もしない
        
        if not username:
            return True  # 認証情報がない場合はそのまま実行
        
        try:
            # ネットワークパスからサーバー部分を抽出
            server_path = "\\\\" + path.split("\\")[2]
            
            # 既存の接続を切断
            disconnect_cmd = f'net use "{server_path}" /delete /y'
            subprocess.run(disconnect_cmd, shell=True, capture_output=True)
            
            # 認証情報を構築
            if domain:
                user_part = f"{domain}\\{username}"
            else:
                user_part = username
            
            # net useコマンドで接続
            connect_cmd = f'net use "{server_path}" /user:"{user_part}" "{password}"'
            
            result = subprocess.run(connect_cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            if result.returncode == 0:
                self.log_message(f"ネットワーク接続成功: {server_path}")
                return True
            else:
                error_msg = result.stderr.strip() if result.stderr else "不明なエラー"
                self.log_message(f"ネットワーク接続失敗: {server_path} - {error_msg}", "error")
                return False
                
        except Exception as e:
            self.log_message(f"ネットワーク接続エラー: {str(e)}", "error")
            return False
    
    def disconnect_network_paths(self):
        """使用したネットワークパスの接続を切断"""
        paths_to_disconnect = []
        
        # コピー元がネットワークパスの場合
        source_path = self.source_var.get()
        if self.is_network_path(source_path) and self.source_username_var.get():
            server_path = "\\\\" + source_path.split("\\")[2]
            paths_to_disconnect.append(server_path)
        
        # コピー先がネットワークパスの場合
        dest_path = self.dest_var.get()
        if self.is_network_path(dest_path) and self.dest_username_var.get():
            server_path = "\\\\" + dest_path.split("\\")[2]
            if server_path not in paths_to_disconnect:  # 重複チェック
                paths_to_disconnect.append(server_path)
        
        # 接続を切断
        for server_path in paths_to_disconnect:
            try:
                disconnect_cmd = f'net use "{server_path}" /delete /y'
                subprocess.run(disconnect_cmd, shell=True, capture_output=True)
                self.log_message(f"ネットワーク接続切断: {server_path}")
            except Exception as e:
                self.log_message(f"接続切断エラー: {server_path} - {str(e)}", "error")
    
    def create_widgets(self):
        # スクロール可能なメインフレームを作成
        # Canvasとスクロールバーを作成
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, padding="10")
        
        # スクロール可能フレームの設定
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Canvasにフレームを配置
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # CanvasとScrollbarを配置
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # マウスホイールでスクロールできるようにする
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        
        # main_frameをscrollable_frameに変更
        main_frame = self.scrollable_frame
        
        # Robocopy設定セクション
        robocopy_frame = ttk.LabelFrame(main_frame, text="Robocopy設定", padding="10")
        robocopy_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # コピー元フォルダ
        ttk.Label(robocopy_frame, text="コピー元フォルダ:").grid(row=0, column=0, sticky=tk.W)
        self.source_var = tk.StringVar()
        # パス変更時のイベントを追加
        self.source_var.trace('w', lambda *args: self.update_auth_state())
        ttk.Entry(robocopy_frame, textvariable=self.source_var, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(robocopy_frame, text="参照", 
                  command=self.browse_source).grid(row=0, column=2)
        
        # コピー元認証設定フレーム（ネットワークパス用）
        self.source_auth_frame = ttk.LabelFrame(robocopy_frame, text="コピー元認証設定（ネットワークパス用）", padding="5")
        self.source_auth_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # コピー元認証情報入力欄
        ttk.Label(self.source_auth_frame, text="ユーザー名:").grid(row=0, column=0, sticky=tk.W)
        source_username_entry = ttk.Entry(self.source_auth_frame, textvariable=self.source_username_var, width=20)
        source_username_entry.grid(row=0, column=1, padx=5)
        self.source_auth_widgets.append(source_username_entry)
        
        ttk.Label(self.source_auth_frame, text="パスワード:").grid(row=0, column=2, sticky=tk.W, padx=(10,0))
        source_password_entry = ttk.Entry(self.source_auth_frame, textvariable=self.source_password_var, width=20, show="*")
        source_password_entry.grid(row=0, column=3, padx=5)
        self.source_auth_widgets.append(source_password_entry)
        
        ttk.Label(self.source_auth_frame, text="ドメイン:").grid(row=1, column=0, sticky=tk.W)
        source_domain_entry = ttk.Entry(self.source_auth_frame, textvariable=self.source_domain_var, width=20)
        source_domain_entry.grid(row=1, column=1, padx=5)
        self.source_auth_widgets.append(source_domain_entry)
        
        source_test_button = ttk.Button(self.source_auth_frame, text="接続テスト", 
                  command=lambda: self.test_network_connection("source"))
        source_test_button.grid(row=1, column=2, columnspan=2, padx=10)
        self.source_auth_widgets.append(source_test_button)
        
        # コピー先フォルダ
        ttk.Label(robocopy_frame, text="コピー先フォルダ:").grid(row=2, column=0, sticky=tk.W)
        self.dest_var = tk.StringVar()
        # パス変更時のイベントを追加
        self.dest_var.trace('w', lambda *args: self.update_auth_state())
        ttk.Entry(robocopy_frame, textvariable=self.dest_var, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(robocopy_frame, text="参照", 
                  command=self.browse_dest).grid(row=2, column=2)
        
        # コピー先認証設定フレーム（ネットワークパス用）
        self.dest_auth_frame = ttk.LabelFrame(robocopy_frame, text="コピー先認証設定（ネットワークパス用）", padding="5")
        self.dest_auth_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # コピー先認証情報入力欄
        ttk.Label(self.dest_auth_frame, text="ユーザー名:").grid(row=0, column=0, sticky=tk.W)
        dest_username_entry = ttk.Entry(self.dest_auth_frame, textvariable=self.dest_username_var, width=20)
        dest_username_entry.grid(row=0, column=1, padx=5)
        self.dest_auth_widgets.append(dest_username_entry)
        
        ttk.Label(self.dest_auth_frame, text="パスワード:").grid(row=0, column=2, sticky=tk.W, padx=(10,0))
        dest_password_entry = ttk.Entry(self.dest_auth_frame, textvariable=self.dest_password_var, width=20, show="*")
        dest_password_entry.grid(row=0, column=3, padx=5)
        self.dest_auth_widgets.append(dest_password_entry)
        
        ttk.Label(self.dest_auth_frame, text="ドメイン:").grid(row=1, column=0, sticky=tk.W)
        dest_domain_entry = ttk.Entry(self.dest_auth_frame, textvariable=self.dest_domain_var, width=20)
        dest_domain_entry.grid(row=1, column=1, padx=5)
        self.dest_auth_widgets.append(dest_domain_entry)
        
        dest_test_button = ttk.Button(self.dest_auth_frame, text="接続テスト", 
                  command=lambda: self.test_network_connection("dest"))
        dest_test_button.grid(row=1, column=2, columnspan=2, padx=10)
        self.dest_auth_widgets.append(dest_test_button)
        
        # Robocopyオプション
        ttk.Label(robocopy_frame, text="Robocopyオプション:").grid(row=4, column=0, sticky=(tk.W, tk.N), padx=5, pady=5)
        
        # オプション選択用のフレーム
        options_frame = ttk.Frame(robocopy_frame)
        options_frame.grid(row=4, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # コピーモード選択（ラジオボタン）
        copy_mode_frame = ttk.LabelFrame(options_frame, text="コピーモード", padding="10")
        copy_mode_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(copy_mode_frame, text="/MIR - ミラーモード（削除も同期）", 
                       variable=self.copy_mode_var, value="MIR").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Radiobutton(copy_mode_frame, text="/E - サブディレクトリコピー（削除なし）", 
                       variable=self.copy_mode_var, value="E").grid(row=1, column=0, sticky=tk.W, pady=2)
        
        # その他のオプション（チェックボックス）
        other_options_frame = ttk.LabelFrame(options_frame, text="その他のオプション", padding="10")
        other_options_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # オプション定義（チェックボックス用）- /SECを追加
        options_config = [
            ("no_progress", "/NP", "進行状況を表示しない"),
            ("no_dir_list", "/NDL", "ディレクトリリストを表示しない"),
            ("sec", "/SEC", "NTFS権限もコピー"),
            ("list_only", "/L", "テストモード（実際にコピーしない）"),
            ("enable_log", "LOG", "ログファイルに記録する"),
        ]
        
        # チェックボックスを作成
        for i, (var_name, option, description) in enumerate(options_config):
            self.option_vars[var_name] = tk.BooleanVar()
            row = i // 2
            col = i % 2
            
            frame = ttk.Frame(other_options_frame)
            frame.grid(row=row, column=col, sticky=tk.W, padx=10, pady=2)
            
            ttk.Checkbutton(frame, variable=self.option_vars[var_name]).grid(row=0, column=0)
            ttk.Label(frame, text=f"{option}").grid(row=0, column=1, padx=(5, 0), sticky=tk.W)
            ttk.Label(frame, text=f"({description})", font=('', 8)).grid(row=1, column=1, padx=(5, 0), sticky=tk.W)
        
        # リトライ・ログ設定フレーム
        config_frame = ttk.LabelFrame(options_frame, text="詳細設定", padding="10")
        config_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # リトライ設定
        retry_frame = ttk.Frame(config_frame)
        retry_frame.grid(row=0, column=0, sticky=tk.W)
        
        ttk.Label(retry_frame, text="リトライ回数:").grid(row=0, column=0, sticky=tk.W)
        self.retry_var = tk.StringVar(value="1")
        ttk.Spinbox(retry_frame, from_=0, to=10, textvariable=self.retry_var, width=5).grid(row=0, column=1, padx=5)
        
        ttk.Label(retry_frame, text="リトライ間隔(秒):").grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        self.wait_var = tk.StringVar(value="1")
        ttk.Spinbox(retry_frame, from_=1, to=60, textvariable=self.wait_var, width=5).grid(row=0, column=3, padx=5)
        
        # ログファイル設定
        log_frame = ttk.Frame(config_frame)
        log_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Label(log_frame, text="ログファイル保存先:").grid(row=0, column=0, sticky=tk.W)
        self.log_file_var = tk.StringVar(value="robocopy_log.txt")
        ttk.Entry(log_frame, textvariable=self.log_file_var, width=40).grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E))
        ttk.Button(log_frame, text="参照", command=self.browse_log_file).grid(row=0, column=2, padx=5)
        
        # 追加オプション欄
        custom_frame = ttk.Frame(config_frame)
        custom_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Label(custom_frame, text="追加オプション:").grid(row=0, column=0, sticky=tk.W)
        self.custom_options_var = tk.StringVar()
        ttk.Entry(custom_frame, textvariable=self.custom_options_var, width=50).grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E))
        
        # デフォルト設定：全オプションを有効にする
        self.copy_mode_var.set("MIR")  # デフォルトは/MIR
        self.option_vars["no_progress"].set(True)      # /NP
        self.option_vars["no_dir_list"].set(True)      # /NDL
        self.option_vars["sec"].set(False)              # /SEC（追加）
        self.option_vars["list_only"].set(True)        # /L（テストモード）
        self.option_vars["enable_log"].set(True)       # LOG
        
        # スケジュール設定セクション
        schedule_frame = ttk.LabelFrame(main_frame, text="スケジュール設定", padding="10")
        schedule_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        # タスク名設定（スケジュール設定の最後に追加）
        task_name_frame = ttk.Frame(schedule_frame)
        task_name_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(task_name_frame, text="タスク名:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(task_name_frame, textvariable=self.task_name_var, width=40).grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E))
        ttk.Button(task_name_frame, text="自動生成", 
                  command=self.update_task_name_from_dest).grid(row=0, column=2, padx=5)
        
        # 説明ラベル
        ttk.Label(task_name_frame, text="※ コピー先フォルダを選択すると自動で生成されます", 
                 font=('', 8), foreground='gray').grid(row=1, column=1, sticky=tk.W, padx=5)

        # コピー先フォルダ
        ttk.Label(robocopy_frame, text="コピー先フォルダ:").grid(row=2, column=0, sticky=tk.W)
        self.dest_var = tk.StringVar()
        # パス変更時のイベントを追加（タスク名自動生成も含める）
        self.dest_var.trace('w', lambda *args: (self.update_auth_state(), self.update_task_name_from_dest()))
        ttk.Entry(robocopy_frame, textvariable=self.dest_var, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(robocopy_frame, text="参照", 
                  command=self.browse_dest).grid(row=2, column=2)
        
        # 実行頻度
        ttk.Label(schedule_frame, text="実行頻度:").grid(row=0, column=0, sticky=tk.W)
        self.frequency_var = tk.StringVar(value="DAILY")
        frequency_combo = ttk.Combobox(schedule_frame, textvariable=self.frequency_var,
                                     values=[("DAILY", "毎日"), ("WEEKLY", "毎週")])
        frequency_combo.grid(row=0, column=1, padx=5)
        frequency_combo.state(['readonly'])
        
        # 曜日選択（毎週の場合）
        ttk.Label(schedule_frame, text="曜日:").grid(row=1, column=0, sticky=tk.W)
        self.weekday_var = tk.StringVar(value="MON")
        weekday_combo = ttk.Combobox(schedule_frame, textvariable=self.weekday_var,
                                   values=[("MON", "月曜日"), ("TUE", "火曜日"), ("WED", "水曜日"), 
                                          ("THU", "木曜日"), ("FRI", "金曜日"), ("SAT", "土曜日"), ("SUN", "日曜日")])
        weekday_combo.grid(row=1, column=1, padx=5)
        weekday_combo.state(['readonly'])
        
        # 実行時刻
        ttk.Label(schedule_frame, text="実行時刻:").grid(row=2, column=0, sticky=tk.W)
        time_frame = ttk.Frame(schedule_frame)
        time_frame.grid(row=2, column=1, padx=5)
        
        self.hour_var = tk.StringVar(value="09")
        self.minute_var = tk.StringVar(value="00")
        
        ttk.Spinbox(time_frame, from_=0, to=23, textvariable=self.hour_var, 
                   width=3, format="%02.0f").grid(row=0, column=0)
        ttk.Label(time_frame, text=":").grid(row=0, column=1)
        ttk.Spinbox(time_frame, from_=0, to=59, textvariable=self.minute_var, 
                   width=3, format="%02.0f").grid(row=0, column=2)
        
        # メール設定セクション
        email_frame = ttk.LabelFrame(main_frame, text="メール通知設定", padding="10")
        email_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # メール送信有効化チェックボックス
        self.email_enabled_var = tk.BooleanVar()
        email_check = ttk.Checkbutton(email_frame, text="実行結果をメールで送信", 
                                     variable=self.email_enabled_var,
                                     command=self.toggle_email_settings)
        email_check.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        # SMTPサーバー設定
        self.email_settings_frame = ttk.Frame(email_frame)
        self.email_settings_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        ttk.Label(self.email_settings_frame, text="SMTPサーバー:").grid(row=0, column=0, sticky=tk.W)
        self.smtp_server_var = tk.StringVar(value="smtp.gmail.com")
        ttk.Entry(self.email_settings_frame, textvariable=self.smtp_server_var, width=30).grid(row=0, column=1, padx=5)
        
        ttk.Label(self.email_settings_frame, text="ポート:").grid(row=0, column=2, sticky=tk.W)
        self.smtp_port_var = tk.StringVar(value="587")
        ttk.Entry(self.email_settings_frame, textvariable=self.smtp_port_var, width=10).grid(row=0, column=3, padx=5)

        self.use_ssl_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.email_settings_frame, text="SSLを使用", variable=self.use_ssl_var).grid(row=0, column=4, padx=5)
        
        ttk.Label(self.email_settings_frame, text="送信者メール:").grid(row=1, column=0, sticky=tk.W)
        self.sender_email_var = tk.StringVar()
        ttk.Entry(self.email_settings_frame, textvariable=self.sender_email_var, width=40).grid(row=1, column=1, columnspan=2, padx=5)
        
        ttk.Label(self.email_settings_frame, text="送信者パスワード:").grid(row=2, column=0, sticky=tk.W)
        self.sender_password_var = tk.StringVar()
        ttk.Entry(self.email_settings_frame, textvariable=self.sender_password_var, 
                 width=40, show="*").grid(row=2, column=1, columnspan=2, padx=5)
        
        ttk.Label(self.email_settings_frame, text="送信先メール:").grid(row=3, column=0, sticky=tk.W)
        self.recipient_email_var = tk.StringVar()
        ttk.Entry(self.email_settings_frame, textvariable=self.recipient_email_var, 
                 width=40).grid(row=3, column=1, columnspan=2, padx=5)
        
        # 初期状態ではメール設定を無効化
        self.toggle_email_settings()
        
        # 初期状態では認証設定を無効化
        self.disable_auth_widgets(self.source_auth_widgets)
        self.disable_auth_widgets(self.dest_auth_widgets)
        
        # 制御ボタンセクション
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="設定を保存", 
                  command=self.save_config).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="今すぐ実行", 
                  command=self.run_now).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="タスク作成/更新", 
                  command=self.create_scheduled_task).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="タスク削除", 
                  command=self.delete_scheduled_task).grid(row=0, column=3, padx=5)
        
        # タスクステータス表示
        status_frame = ttk.LabelFrame(main_frame, text="タスクステータス", padding="10")
        status_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # タスク履歴設定セクション
        history_frame = ttk.LabelFrame(main_frame, text="タスク履歴設定", padding="10")
        history_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # 履歴有効化チェックボックス
        self.history_enabled_var = tk.BooleanVar()
        ttk.Checkbutton(history_frame, text="すべてのタスク履歴を有効にする", 
                    variable=self.history_enabled_var).grid(row=0, column=0, sticky=tk.W)

        history_button_frame = ttk.Frame(history_frame)
        history_button_frame.grid(row=1, column=0, pady=5)

        ttk.Button(history_button_frame, text="履歴設定を適用", 
                command=self.apply_task_history).grid(row=0, column=0, padx=5)
        ttk.Button(history_button_frame, text="現在の設定確認", 
                command=self.check_task_history).grid(row=0, column=1, padx=5)
        
        self.task_status_var = tk.StringVar(value="確認中...")
        ttk.Label(status_frame, textvariable=self.task_status_var).grid(row=0, column=0, sticky=tk.W)
        ttk.Button(status_frame, text="ステータス更新", 
                  command=self.update_task_status).grid(row=0, column=1, padx=10)
        
        # ログ表示エリア
        log_frame = ttk.LabelFrame(main_frame, text="実行ログ", padding="10")
        log_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.log_text = tk.Text(log_frame, height=12, width=100)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # ステータスバー
        self.status_var = tk.StringVar(value="準備完了")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
    
    def _on_mousewheel(self, event):
        """マウスホイールでスクロール"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def check_task_history(self):
        """現在のタスク履歴設定を確認"""
        try:
            # より確実な確認方法を使用
            cmd = 'wevtutil gl Microsoft-Windows-TaskScheduler/Operational'
            
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            if result.returncode == 0:
                # enabled: true/false を確認
                if "enabled: true" in result.stdout.lower():
                    self.history_enabled_var.set(True)
                    status_msg = "タスク履歴は現在有効です"
                else:
                    self.history_enabled_var.set(False)
                    status_msg = "タスク履歴は現在無効です"
            else:
                status_msg = "履歴設定の確認に失敗しました"
            
            self.log_message(status_msg)
            messagebox.showinfo("履歴設定確認", status_msg)
            
        except Exception as e:
            error_msg = f"履歴設定確認エラー: {str(e)}"
            self.log_message(error_msg, "error")
            messagebox.showerror("エラー", error_msg)

    def is_admin(self):
        """現在のプロセスが管理者権限で実行されているかチェック"""
        try:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def run_as_admin(self, command):
        """管理者権限でコマンドを実行"""
        try:
            import ctypes
            
            # PowerShellを管理者権限で実行
            result = ctypes.windll.shell32.ShellExecuteW(
                None, 
                "runas",  # 管理者として実行
                "powershell.exe", 
                f"-Command \"{command}\"",
                None, 
                1  # SW_SHOWNORMAL
            )
            
            # 戻り値が32より大きければ成功
            return result > 32
        except Exception as e:
            self.log_message(f"管理者権限実行エラー: {str(e)}", "error")
            return False

    def apply_task_history(self):
        """タスク履歴の有効/無効を設定"""
        try:
            if self.history_enabled_var.get():
                action = "有効化"
                cmd = "wevtutil sl Microsoft-Windows-TaskScheduler/Operational /e:true"
            else:
                action = "無効化"
                cmd = "wevtutil sl Microsoft-Windows-TaskScheduler/Operational /e:false"
            
            # 管理者権限が必要であることをユーザーに通知
            response = messagebox.askyesno(
                "管理者権限が必要", 
                f"タスク履歴の{action}には管理者権限が必要です。\n"
                "UACダイアログが表示されますが、続行しますか？"
            )
            
            if not response:
                self.log_message("操作がキャンセルされました")
                return
            
            # 現在管理者権限があるかチェック
            if self.is_admin():
                # すでに管理者権限がある場合は直接実行
                self.log_message("管理者権限で実行中...")
                result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
                
                if result.returncode == 0:
                    success_msg = f"タスク履歴を{action}しました"
                    self.log_message(success_msg)
                    messagebox.showinfo("成功", success_msg)
                    # 設定確認を自動実行
                    self.check_task_history()
                else:
                    error_msg = f"履歴{action}エラー: {result.stderr}"
                    self.log_message(error_msg, "error")
                    messagebox.showerror("エラー", error_msg)
            else:
                # 管理者権限がない場合はUACダイアログを表示して実行
                self.log_message(f"管理者権限でタスク履歴{action}を実行します...")
                
                if self.run_as_admin(cmd):
                    self.log_message("管理者権限でのコマンド実行を開始しました")
                    messagebox.showinfo(
                        "実行中", 
                        f"管理者権限でタスク履歴{action}を実行中です。\n"
                        "完了後、「現在の設定確認」ボタンで結果を確認してください。"
                    )
                    # 少し待ってから設定確認
                    self.root.after(2000, self.check_task_history)  # 2秒後に確認
                else:
                    error_msg = "管理者権限での実行に失敗しました"
                    self.log_message(error_msg, "error")
                    messagebox.showerror("エラー", error_msg)
                    
        except Exception as e:
            error_msg = f"履歴設定エラー: {str(e)}"
            self.log_message(error_msg, "error")
            messagebox.showerror("エラー", error_msg)

    def browse_source(self):
        """コピー元フォルダを選択"""
        folder = filedialog.askdirectory(title="コピー元フォルダを選択")
        if folder:
            # パスの区切り文字を統一（/ → \）
            folder = folder.replace('/', '\\')
            self.source_var.set(folder)
            # フォルダ選択後に認証設定の状態を更新
            self.update_auth_state()
    
    def browse_dest(self):
        """コピー先フォルダを選択"""
        folder = filedialog.askdirectory(title="コピー先フォルダを選択")
        if folder:
            # パスの区切り文字を統一（/ → \）
            folder = folder.replace('/', '\\')
            self.dest_var.set(folder)
            # フォルダ選択後に認証設定の状態を更新
            self.update_auth_state()
            # タスク名を自動生成
            self.update_task_name_from_dest()
    
    def browse_log_file(self):
        """ログファイルの保存先を選択"""
        file_path = filedialog.asksaveasfilename(
            title="ログファイルの保存先を選択",
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            self.log_file_var.set(file_path)
    
    def toggle_email_settings(self):
        """メール設定の有効/無効を切り替え"""
        if self.email_enabled_var.get():
            # メール設定を有効化
            for widget in self.email_settings_frame.winfo_children():
                widget.configure(state='normal')
        else:
            # メール設定を無効化
            for widget in self.email_settings_frame.winfo_children():
                if isinstance(widget, ttk.Entry):
                    widget.configure(state='disabled')
    
    def log_message(self, message, tag=None):
        """ログにメッセージを追加"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, full_message, tag)
        self.log_text.see(tk.END)
        
        # タグの設定（初回のみ）
        if not hasattr(self, '_tags_configured'):
            self.log_text.tag_config("error", foreground="red")
            self.log_text.tag_config("success", foreground="green")
            self._tags_configured = True
        
        self.root.update()
    
    def build_robocopy_options(self):
        """選択されたオプションからRobocopyのオプション文字列を構築"""
        options = []
        
        # コピーモードを追加（ラジオボタンから取得）
        copy_mode = self.copy_mode_var.get()
        if copy_mode == "MIR":
            options.append("/MIR")
        elif copy_mode == "E":
            options.append("/E")
        
        # その他のチェックボックスで選択されたオプションを追加
        if self.option_vars["no_progress"].get():
            options.append("/NP")
        
        if self.option_vars["no_dir_list"].get():
            options.append("/NDL")
        
        # /SECオプションを追加
        if self.option_vars["sec"].get():
            options.append("/SEC")
        
        if self.option_vars["list_only"].get():
            options.append("/L")
        
        # リトライ設定を追加
        retry_count = self.retry_var.get()
        wait_time = self.wait_var.get()
        if retry_count:
            options.append(f"/R:{retry_count}")
        if wait_time:
            options.append(f"/W:{wait_time}")
        
        # ログファイル設定
        if self.option_vars["enable_log"].get():
            log_file = self.log_file_var.get()
            if log_file:
                # ログファイルパスにスペースが含まれる場合はクォートで囲む
                if " " in log_file:
                    options.append(f'/LOG+:"{log_file}"')
                else:
                    options.append(f"/LOG+:{log_file}")
        
        # カスタムオプションを追加
        custom_options = self.custom_options_var.get().strip()
        if custom_options:
            options.append(custom_options)
        
        return " ".join(options)
    
    def run_robocopy(self):
        """Robocopyを実行"""
        source = self.source_var.get()
        dest = self.dest_var.get()
        options = self.build_robocopy_options()
        
        if not source or not dest:
            self.log_message("エラー: コピー元またはコピー先が指定されていません", "error")
            return False, "コピー元またはコピー先が指定されていません"
        
        # ネットワークパスの認証を実行
        auth_success = True
        
        # コピー元の認証
        if self.is_network_path(source) and self.source_username_var.get():
            self.log_message("コピー元ネットワークパスに接続中...")
            if not self.connect_network_path(
                source, 
                self.source_username_var.get(),
                self.source_password_var.get(),
                self.source_domain_var.get()
            ):
                auth_success = False
                error_msg = "コピー元ネットワークパスへの認証に失敗しました"
                self.log_message(error_msg, "error")
                return False, error_msg
        
        # コピー先の認証
        if self.is_network_path(dest) and self.dest_username_var.get():
            self.log_message("コピー先ネットワークパスに接続中...")
            if not self.connect_network_path(
                dest,
                self.dest_username_var.get(),
                self.dest_password_var.get(),
                self.dest_domain_var.get()
            ):
                auth_success = False
                error_msg = "コピー先ネットワークパスへの認証に失敗しました"
                self.log_message(error_msg, "error")
                return False, error_msg
        
        if not auth_success:
            return False, "ネットワーク認証に失敗しました"
        
        # Robocopyコマンドを構築
        cmd = f'robocopy "{source}" "{dest}" {options}'
        
        try:
            self.log_message(f"Robocopy実行開始: {cmd}")
            self.log_message("-" * 50)  # 区切り線
            
            # リアルタイムで出力を表示するためのプロセス実行
            process = subprocess.Popen(
                cmd, 
                shell=True, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE, 
                text=True, 
                encoding='cp932',
                bufsize=1,  # 行バッファリング
                universal_newlines=True
            )
            
            # 標準出力をリアルタイムで表示
            output_lines = []
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    output_lines.append(output.strip())
                    self.log_message(output.strip())
            
            # プロセスの完了を待つ
            return_code = process.wait()
            
            # エラー出力があれば表示
            stderr_output = process.stderr.read()
            if stderr_output:
                self.log_message("エラー出力:" , "error")
                for line in stderr_output.strip().split('\n'):
                    if line.strip():
                        self.log_message(f"  {line}" , "error")
            
            self.log_message("-" * 50)  # 区切り線
            
            # Robocopyの戻り値を確認（0-7は正常、8以上はエラー）
            if return_code < 8:
                self.log_message(f"Robocopy実行完了（戻り値: {return_code}）")
                success = True
                message = f"バックアップが正常に完了しました。\n戻り値: {return_code}\n\n" + "\n".join(output_lines)
            else:
                self.log_message(f"Robocopyでエラーが発生しました（戻り値: {return_code}）" , "error")
                success = False
                message = f"バックアップでエラーが発生しました。\n戻り値: {return_code}\n\n" + "\n".join(output_lines)
                if stderr_output:
                    message += f"\n\nエラー詳細:\n{stderr_output}"
            
            return success, message
            
        except Exception as e:
            error_msg = f"実行エラー: {str(e)}"
            self.log_message(error_msg , "error")
            return False, error_msg
        finally:
            # 実行後はネットワーク接続を切断
            self.disconnect_network_paths()
    
    def send_email(self, success, message):
        """メールを送信"""
        if not self.email_enabled_var.get():
            return
        
        try:
            smtp_server = self.smtp_server_var.get()
            smtp_port = int(self.smtp_port_var.get())
            sender_email = self.sender_email_var.get()
            sender_password = self.sender_password_var.get()
            recipient_email = self.recipient_email_var.get()
            
            # メール内容を作成
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Subject'] = f"Robocopyバックアップ結果 - {'成功' if success else '失敗'}"
            
            body = f"""
Robocopyバックアップの実行結果をお知らせします。

実行日時: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
結果: {'成功' if success else '失敗'}
コピー元: {self.source_var.get()}
コピー先: {self.dest_var.get()}

詳細:
{message}
            """
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # SMTPサーバーに接続してメール送信
            if self.use_ssl_var.get():
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10)
            else:
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
                server.ehlo()
                if server.has_extn('STARTTLS'):
                    server.starttls()
                    server.ehlo()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            server.quit()
            
            self.log_message("メール送信完了")
            
        except Exception as e:
            self.log_message(f"メール送信エラー: {str(e)}" , "error")
    
    def run_now(self):
        """今すぐrobocopyを実行"""
        self.status_var.set("実行中...")
        success, message = self.run_robocopy()
        
        if self.email_enabled_var.get():
            self.send_email(success, message)
        
        self.status_var.set("実行完了" if success else "実行エラー")
    
    def create_scheduled_task(self):
        """Windowsタスクスケジューラにタスクを作成"""
        if not self.source_var.get() or not self.dest_var.get():
            messagebox.showerror("エラー", "コピー元とコピー先を設定してください")
            return
        
        task_name = self.task_name_var.get().strip()
        if not task_name:
            messagebox.showerror("エラー", "タスク名を入力してください")
            return
        
        # 設定を保存してからタスクを作成
        self.save_config()
        
        # 現在のスクリプトのパスを取得
        script_path = os.path.abspath(__file__)
        python_path = sys.executable
        
        # タスクの実行時刻
        start_time = f"{self.hour_var.get()}:{self.minute_var.get()}"
        
        # スケジュール頻度に応じてコマンドを構築
        if self.frequency_var.get() == "DAILY":
            schedule_type = "/SC DAILY"
        else:  # WEEKLY
            weekday = self.weekday_var.get()
            schedule_type = f"/SC WEEKLY /D {weekday}"
        
        # schtasksコマンドを構築
        cmd = f'''schtasks /CREATE /TN "{task_name}" /TR "\\"{python_path}\\" \\"{script_path}\\" --scheduled" {schedule_type} /ST {start_time} /F'''
        
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            if result.returncode == 0:
                self.log_message(f"タスクを作成しました: {task_name}")
                messagebox.showinfo("成功", f"スケジュールタスク '{task_name}' を作成しました")
                self.update_task_status()
            else:
                error_msg = f"タスク作成エラー: {result.stderr}"
                self.log_message(error_msg , "error")
                messagebox.showerror("エラー", error_msg)
                
        except Exception as e:
            error_msg = f"タスク作成エラー: {str(e)}"
            self.log_message(error_msg , "error")
            messagebox.showerror("エラー", error_msg)

    def delete_scheduled_task(self):
        """スケジュールされたタスクを削除"""
        task_name = self.task_name_var.get().strip()
        if not task_name:
            messagebox.showerror("エラー", "タスク名が入力されていません")
            return
            
        cmd = f'schtasks /DELETE /TN "{task_name}" /F'
        
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            if result.returncode == 0:
                self.log_message(f"タスクを削除しました: {task_name}")
                messagebox.showinfo("成功", f"スケジュールタスク '{task_name}' を削除しました")
                self.update_task_status()
            else:
                error_msg = f"タスク削除エラー: {result.stderr}"
                self.log_message(error_msg , "error")
                messagebox.showerror("エラー", error_msg)
                
        except Exception as e:
            error_msg = f"タスク削除エラー: {str(e)}"
            self.log_message(error_msg , "error")
            messagebox.showerror("エラー", error_msg)

    def update_task_status(self):
        """タスクのステータスを更新"""
        task_name = self.task_name_var.get().strip()
        if not task_name:
            self.task_status_var.set("タスク名が未設定")
            return
            
        cmd = f'schtasks /QUERY /TN "{task_name}"'
        
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            if result.returncode == 0:
                # タスクが存在する場合
                lines = result.stdout.split('\n')
                for line in lines:
                    if 'Next Run Time' in line or '次回の実行時刻' in line:
                        next_run = line.split(':')[1].strip() if ':' in line else "不明"
                        self.task_status_var.set(f"タスク登録済み - 次回実行: {next_run}")
                        return
                
                self.task_status_var.set("タスク登録済み")
            else:
                self.task_status_var.set("タスク未登録")
                
        except Exception as e:
            self.task_status_var.set(f"ステータス確認エラー: {str(e)}")
    
    def save_config(self):
        """設定をファイルに保存"""
        config = {
            'source': self.source_var.get(),
            'dest': self.dest_var.get(),
            'task_name': self.task_name_var.get(),  # タスク名を追加
            'copy_mode': self.copy_mode_var.get(),
            'robocopy_options': {var_name: var.get() for var_name, var in self.option_vars.items()},
            'retry_count': self.retry_var.get(),
            'wait_time': self.wait_var.get(),
            'log_file': self.log_file_var.get(),
            'custom_options': self.custom_options_var.get(),
            'frequency': self.frequency_var.get(),
            'weekday': self.weekday_var.get(),
            'hour': self.hour_var.get(),
            'minute': self.minute_var.get(),
            'email_enabled': self.email_enabled_var.get(),
            'smtp_server': self.smtp_server_var.get(),
            'smtp_port': self.smtp_port_var.get(),
            'sender_email': self.sender_email_var.get(),
            'sender_password': self.sender_password_var.get(),
            'recipient_email': self.recipient_email_var.get(),
            'history_enabled': self.history_enabled_var.get(),
            'use_ssl': self.use_ssl_var.get(),
            # 認証情報を追加
            'source_auth_enabled': self.source_auth_enabled_var.get(),
            'source_username': self.source_username_var.get(),
            'source_password': self.source_password_var.get(),
            'source_domain': self.source_domain_var.get(),
            'dest_auth_enabled': self.dest_auth_enabled_var.get(),
            'dest_username': self.dest_username_var.get(),
            'dest_password': self.dest_password_var.get(),
            'dest_domain': self.dest_domain_var.get()
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            self.log_message("設定を保存しました")
            
        except Exception as e:
            error_msg = f"設定保存エラー: {str(e)}"
            self.log_message(error_msg , "error")
            messagebox.showerror("エラー", error_msg)

    def load_config(self):
        """設定をファイルから読み込み"""
        if not os.path.exists(self.config_file):
            return
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            self.source_var.set(config.get('source', ''))
            self.dest_var.set(config.get('dest', ''))
            
            # タスク名の読み込み（デフォルト値を設定）
            saved_task_name = config.get('task_name', '')
            if saved_task_name:
                self.task_name_var.set(saved_task_name)
            else:
                # 保存されたタスク名がない場合は自動生成
                self.update_task_name_from_dest()
            
            # コピーモードの読み込み
            self.copy_mode_var.set(config.get('copy_mode', 'MIR'))
            
            # 以下は既存のコードと同じ...
            # オプション設定の読み込み
            robocopy_options = config.get('robocopy_options', {})
            for var_name, var in self.option_vars.items():
                var.set(robocopy_options.get(var_name, False))
            
            self.retry_var.set(config.get('retry_count', '1'))
            self.wait_var.set(config.get('wait_time', '1'))
            self.log_file_var.set(config.get('log_file', 'robocopy_log.txt'))
            self.custom_options_var.set(config.get('custom_options', ''))
            
            # 既存の設定読み込み（変更なし）
            self.frequency_var.set(config.get('frequency', 'DAILY'))
            self.weekday_var.set(config.get('weekday', 'MON'))
            self.hour_var.set(config.get('hour', '09'))
            self.minute_var.set(config.get('minute', '00'))
            self.email_enabled_var.set(config.get('email_enabled', False))
            self.smtp_server_var.set(config.get('smtp_server', 'smtp.gmail.com'))
            self.smtp_port_var.set(config.get('smtp_port', '587'))
            self.sender_email_var.set(config.get('sender_email', ''))
            self.sender_password_var.set(config.get('sender_password', ''))
            self.recipient_email_var.set(config.get('recipient_email', ''))
            self.history_enabled_var.set(config.get('history_enabled', False))
            self.use_ssl_var.set(config.get('use_ssl', False))
            
            # 認証情報の読み込み
            self.source_auth_enabled_var.set(config.get('source_auth_enabled', False))
            self.source_username_var.set(config.get('source_username', ''))
            self.source_password_var.set(config.get('source_password', ''))
            self.source_domain_var.set(config.get('source_domain', ''))
            self.dest_auth_enabled_var.set(config.get('dest_auth_enabled', False))
            self.dest_username_var.set(config.get('dest_username', ''))
            self.dest_password_var.set(config.get('dest_password', ''))
            self.dest_domain_var.set(config.get('dest_domain', ''))
            
            self.toggle_email_settings()
            # 認証設定の状態を更新
            self.update_auth_state()
            self.log_message("設定を読み込みました")
            
        except Exception as e:
            self.log_message(f"設定読み込みエラー: {str(e)}" , "error")

def scheduled_execution():
    """スケジュール実行用の関数（コマンドライン引数で呼び出される）"""
    scheduler = RobocopyScheduler(None)
    
    # 設定を読み込み
    scheduler.load_config()
    
    # Robocopyを実行
    success, message = scheduler.run_robocopy()
    
    # メール送信（設定されている場合）
    if scheduler.email_enabled_var.get():
        scheduler.send_email(success, message)
    
    # ログファイルに結果を記録
    log_file = "robocopy_schedule_log.txt"
    with open(log_file, 'a', encoding='utf-8') as f:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"[{timestamp}] {'成功' if success else '失敗'}: {message}\n")

def main():
    # コマンドライン引数をチェック
    if len(sys.argv) > 1 and sys.argv[1] == '--scheduled':
        # スケジュール実行
        scheduled_execution()
    else:
        # GUI実行
        root = tk.Tk()
        app = RobocopyScheduler(root)
        root.mainloop()

if __name__ == "__main__":
    main()