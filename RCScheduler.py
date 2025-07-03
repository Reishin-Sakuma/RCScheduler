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
import glob
import tempfile

# バージョン情報
VERSION = "1.0.0"

class RobocopyScheduler:
    def __init__(self, root):
        self.root = root
        self.root.title(f"RCScheduler - ver.{VERSION}")
        self.root.geometry("900x1000")
        
        # モダンなスタイル設定
        self.setup_modern_style()
        
        # メインウィンドウの背景色を設定（ダークテーマ）
        self.root.configure(bg='#1a1a1a')
        
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
    
    def setup_modern_style(self):
        """モダンなスタイルを設定"""
        style = ttk.Style()
        
        # テーマを設定
        style.theme_use('clam')
        
        # ダークテーマの色の定義
        self.colors = {
            'primary': '#3b82f6',      # 明るい青
            'secondary': '#6b7280',    # グレー
            'success': '#10b981',      # 緑
            'warning': '#f59e0b',      # オレンジ
            'danger': '#ef4444',       # 赤
            'bg_primary': '#1a1a1a',   # メイン背景
            'bg_secondary': '#2d2d2d', # セカンダリ背景
            'bg_tertiary': '#404040',  # 第三背景
            'white': '#ffffff',        # 白
            'dark': '#0f0f0f',         # 最も濃い
            'border': '#404040',       # ボーダー
            'text': '#e5e5e5',         # メインテキスト
            'text_muted': '#a1a1a1',   # 薄いテキスト
            'text_dark': '#ffffff'     # 白テキスト
        }
        
        # フレームスタイル（ダークテーマ）
        style.configure('Modern.TLabelframe', 
                       background=self.colors['bg_secondary'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=self.colors['border'])
        
        style.configure('Modern.TLabelframe.Label',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text_dark'],
                       font=('メイリオ', 10, 'bold'))
        
        style.configure('Modern.TFrame',
                       background=self.colors['bg_secondary'])
        
        # ボタンスタイル
        style.configure('Primary.TButton',
                       font=('メイリオ', 9),
                       padding=(12, 8),
                       background=self.colors['primary'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none')
        
        style.map('Primary.TButton',
                 background=[('active', '#1d4ed8'),
                           ('pressed', '#1e40af')])
        
        style.configure('Secondary.TButton',
                       font=('メイリオ', 9),
                       padding=(12, 8),
                       background=self.colors['secondary'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none')
        
        style.map('Secondary.TButton',
                 background=[('active', '#475569'),
                           ('pressed', '#334155')])
        
        style.configure('Success.TButton',
                       font=('メイリオ', 9),
                       padding=(12, 8),
                       background=self.colors['success'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none')
        
        style.map('Success.TButton',
                 background=[('active', '#059669'),
                           ('pressed', '#047857')])
        
        style.configure('Warning.TButton',
                       font=('メイリオ', 9),
                       padding=(12, 8),
                       background=self.colors['warning'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none')
        
        style.map('Warning.TButton',
                 background=[('active', '#d97706'),
                           ('pressed', '#b45309')])
        
        style.configure('Danger.TButton',
                       font=('メイリオ', 9),
                       padding=(12, 8),
                       background=self.colors['danger'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none')
        
        style.map('Danger.TButton',
                 background=[('active', '#dc2626'),
                           ('pressed', '#b91c1c')])
        
        # エントリスタイル（ダークテーマ）
        style.configure('Modern.TEntry',
                       fieldbackground=self.colors['bg_tertiary'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=self.colors['border'],
                       padding=(8, 8),
                       font=('メイリオ', 9),
                       foreground=self.colors['text'])
        
        style.map('Modern.TEntry',
                 bordercolor=[('focus', self.colors['primary'])],
                 fieldbackground=[('focus', self.colors['bg_tertiary'])])
        
        # コンボボックススタイル（ダークテーマ）
        style.configure('Modern.TCombobox',
                       fieldbackground=self.colors['bg_tertiary'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=self.colors['border'],
                       padding=(8, 8),
                       font=('メイリオ', 9),
                       foreground=self.colors['text'])
        
        style.map('Modern.TCombobox',
                 bordercolor=[('focus', self.colors['primary'])],
                 fieldbackground=[('focus', self.colors['bg_tertiary'])])
        
        # ラベルスタイル（ダークテーマ）
        style.configure('Modern.TLabel',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text'],
                       font=('メイリオ', 9))
        
        style.configure('Title.TLabel',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text_dark'],
                       font=('メイリオ', 18, 'bold'))
        
        style.configure('Subtitle.TLabel',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text_muted'],
                       font=('メイリオ', 10))
        
        # チェックボタンスタイル（ダークテーマ）
        style.configure('Modern.TCheckbutton',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text'],
                       font=('メイリオ', 9),
                       focuscolor='none')
        
        # ラジオボタンスタイル（ダークテーマ）
        style.configure('Modern.TRadiobutton',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text'],
                       font=('メイリオ', 9),
                       focuscolor='none')
        
        # スピンボックススタイル（ダークテーマ）
        style.configure('Modern.TSpinbox',
                       fieldbackground=self.colors['bg_tertiary'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=self.colors['border'],
                       padding=(8, 8),
                       font=('メイリオ', 9),
                       foreground=self.colors['text'])
        
        style.map('Modern.TSpinbox',
                 bordercolor=[('focus', self.colors['primary'])],
                 fieldbackground=[('focus', self.colors['bg_tertiary'])])
        
        # セパレータスタイル（ダークテーマ）
        style.configure('Modern.TSeparator',
                       background=self.colors['border'])
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
        self.canvas = tk.Canvas(self.root, bg='#1a1a1a', highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, padding="20", style='Modern.TFrame')
        
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
        
        # ヘッダーセクション（プログラム名と説明）
        header_frame = ttk.Frame(main_frame, style='Modern.TFrame')
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 25))

        # プログラム名
        title_label = ttk.Label(header_frame, text="RCScheduler", style='Title.TLabel')
        title_label.grid(row=0, column=0, sticky=tk.W)

        # バージョン情報
        version_label = ttk.Label(header_frame, text=f"v{VERSION}", style='Subtitle.TLabel')
        version_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))

        # プログラム説明
        description_label = ttk.Label(header_frame, 
                                    text="Windows用Robocopyスケジューラソフト - フォルダバックアップの自動化\nRobocopyの基本的なコマンド生成からタスクスケジューラ登録、メール通知設定まで一貫して行えます。", 
                                    style='Subtitle.TLabel')
        description_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))

        # 区切り線
        separator = ttk.Separator(main_frame, orient='horizontal', style='Modern.TSeparator')
        separator.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 20))

        # Robocopy設定セクション
        robocopy_frame = ttk.LabelFrame(main_frame, text="📁 Robocopy設定", padding="20", style='Modern.TLabelframe')
        robocopy_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # コピー元フォルダ
        ttk.Label(robocopy_frame, text="コピー元フォルダ:", style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.source_var = tk.StringVar()
        # パス変更時のイベントを追加
        self.source_var.trace('w', lambda *args: self.update_auth_state())
        ttk.Entry(robocopy_frame, textvariable=self.source_var, width=50, style='Modern.TEntry').grid(row=0, column=1, padx=(10, 5), pady=(0, 5))
        ttk.Button(robocopy_frame, text="📂 参照", 
                  command=self.browse_source, style='Secondary.TButton').grid(row=0, column=2, pady=(0, 5))
        
        # コピー元認証設定フレーム（ネットワークパス用）
        self.source_auth_frame = ttk.LabelFrame(robocopy_frame, text="🔐 コピー元認証設定（ネットワークパス用）", padding="15", style='Modern.TLabelframe')
        self.source_auth_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
        
        # コピー元認証情報入力欄
        ttk.Label(self.source_auth_frame, text="ユーザー名:", style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        source_username_entry = ttk.Entry(self.source_auth_frame, textvariable=self.source_username_var, width=20, style='Modern.TEntry')
        source_username_entry.grid(row=0, column=1, padx=(10, 5), pady=(0, 5))
        self.source_auth_widgets.append(source_username_entry)
        
        ttk.Label(self.source_auth_frame, text="パスワード:", style='Modern.TLabel').grid(row=0, column=2, sticky=tk.W, padx=(15, 0), pady=(0, 5))
        source_password_entry = ttk.Entry(self.source_auth_frame, textvariable=self.source_password_var, width=20, show="*", style='Modern.TEntry')
        source_password_entry.grid(row=0, column=3, padx=(10, 5), pady=(0, 5))
        self.source_auth_widgets.append(source_password_entry)
        
        ttk.Label(self.source_auth_frame, text="ドメイン:", style='Modern.TLabel').grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        source_domain_entry = ttk.Entry(self.source_auth_frame, textvariable=self.source_domain_var, width=20, style='Modern.TEntry')
        source_domain_entry.grid(row=1, column=1, padx=(10, 5), pady=(5, 0))
        self.source_auth_widgets.append(source_domain_entry)
        
        source_test_button = ttk.Button(self.source_auth_frame, text="🔍 接続テスト", 
                  command=lambda: self.test_network_connection("source"), style='Secondary.TButton')
        source_test_button.grid(row=1, column=2, columnspan=2, padx=(15, 0), pady=(5, 0))
        self.source_auth_widgets.append(source_test_button)
        
        # コピー先フォルダ
        ttk.Label(robocopy_frame, text="コピー先フォルダ:", style='Modern.TLabel').grid(row=2, column=0, sticky=tk.W, pady=(15, 5))
        self.dest_var = tk.StringVar()
        # パス変更時のイベントを追加（認証状態とタスク名自動更新）
        self.dest_var.trace('w', lambda *args: (self.update_auth_state(), self.update_task_name_from_dest()))
        ttk.Entry(robocopy_frame, textvariable=self.dest_var, width=50, style='Modern.TEntry').grid(row=2, column=1, padx=(10, 5), pady=(15, 5))
        ttk.Button(robocopy_frame, text="📂 参照", 
                  command=self.browse_dest, style='Secondary.TButton').grid(row=2, column=2, pady=(15, 5))
        
        # コピー先認証設定フレーム（ネットワークパス用）
        self.dest_auth_frame = ttk.LabelFrame(robocopy_frame, text="🔐 コピー先認証設定（ネットワークパス用）", padding="15", style='Modern.TLabelframe')
        self.dest_auth_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
        
        # コピー先認証情報入力欄
        ttk.Label(self.dest_auth_frame, text="ユーザー名:", style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        dest_username_entry = ttk.Entry(self.dest_auth_frame, textvariable=self.dest_username_var, width=20, style='Modern.TEntry')
        dest_username_entry.grid(row=0, column=1, padx=(10, 5), pady=(0, 5))
        self.dest_auth_widgets.append(dest_username_entry)
        
        ttk.Label(self.dest_auth_frame, text="パスワード:", style='Modern.TLabel').grid(row=0, column=2, sticky=tk.W, padx=(15, 0), pady=(0, 5))
        dest_password_entry = ttk.Entry(self.dest_auth_frame, textvariable=self.dest_password_var, width=20, show="*", style='Modern.TEntry')
        dest_password_entry.grid(row=0, column=3, padx=(10, 5), pady=(0, 5))
        self.dest_auth_widgets.append(dest_password_entry)
        
        ttk.Label(self.dest_auth_frame, text="ドメイン:", style='Modern.TLabel').grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        dest_domain_entry = ttk.Entry(self.dest_auth_frame, textvariable=self.dest_domain_var, width=20, style='Modern.TEntry')
        dest_domain_entry.grid(row=1, column=1, padx=(10, 5), pady=(5, 0))
        self.dest_auth_widgets.append(dest_domain_entry)
        
        dest_test_button = ttk.Button(self.dest_auth_frame, text="🔍 接続テスト", 
                  command=lambda: self.test_network_connection("dest"), style='Secondary.TButton')
        dest_test_button.grid(row=1, column=2, columnspan=2, padx=(15, 0), pady=(5, 0))
        self.dest_auth_widgets.append(dest_test_button)
        
        # Robocopyオプション
        ttk.Label(robocopy_frame, text="⚙️ Robocopyオプション:", style='Modern.TLabel').grid(row=4, column=0, sticky=(tk.W, tk.N), pady=(15, 5))
        
        # オプション選択用のフレーム
        options_frame = ttk.Frame(robocopy_frame, style='Modern.TFrame')
        options_frame.grid(row=4, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=(10, 0), pady=(15, 5))
        
        # コピーモード選択（ラジオボタン）
        copy_mode_frame = ttk.LabelFrame(options_frame, text="📋 コピーモード", padding="15", style='Modern.TLabelframe')
        copy_mode_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Radiobutton(copy_mode_frame, text="/MIR - ミラーモード（削除も同期）", 
                       variable=self.copy_mode_var, value="MIR", style='Modern.TRadiobutton').grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Radiobutton(copy_mode_frame, text="/E - サブディレクトリコピー（削除なし）", 
                       variable=self.copy_mode_var, value="E", style='Modern.TRadiobutton').grid(row=1, column=0, sticky=tk.W, pady=3)
        
        # その他のオプション（チェックボックス）
        other_options_frame = ttk.LabelFrame(options_frame, text="🔧 その他のオプション", padding="15", style='Modern.TLabelframe')
        other_options_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
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
            
            frame = ttk.Frame(other_options_frame, style='Modern.TFrame')
            frame.grid(row=row, column=col, sticky=tk.W, padx=15, pady=5)
            
            ttk.Checkbutton(frame, variable=self.option_vars[var_name], style='Modern.TCheckbutton').grid(row=0, column=0)
            ttk.Label(frame, text=f"{option}", style='Modern.TLabel', font=('メイリオ', 9, 'bold')).grid(row=0, column=1, padx=(8, 0), sticky=tk.W)
            ttk.Label(frame, text=f"({description})", style='Subtitle.TLabel').grid(row=1, column=1, padx=(8, 0), sticky=tk.W)
        
        # リトライ・ログ設定フレーム
        config_frame = ttk.LabelFrame(options_frame, text="🔧 詳細設定", padding="15", style='Modern.TLabelframe')
        config_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # リトライ設定
        retry_frame = ttk.Frame(config_frame, style='Modern.TFrame')
        retry_frame.grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        
        ttk.Label(retry_frame, text="リトライ回数:", style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W)
        self.retry_var = tk.StringVar(value="1")
        ttk.Spinbox(retry_frame, from_=0, to=10, textvariable=self.retry_var, width=8, style='Modern.TSpinbox').grid(row=0, column=1, padx=(8, 0))
        
        ttk.Label(retry_frame, text="リトライ間隔(秒):", style='Modern.TLabel').grid(row=0, column=2, sticky=tk.W, padx=(15, 0))
        self.wait_var = tk.StringVar(value="1")
        ttk.Spinbox(retry_frame, from_=1, to=60, textvariable=self.wait_var, width=8, style='Modern.TSpinbox').grid(row=0, column=3, padx=(8, 0))
        
        # ログファイル設定
        log_frame = ttk.Frame(config_frame, style='Modern.TFrame')
        log_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(8, 8))
        
        ttk.Label(log_frame, text="ログファイル保存先:", style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.log_file_var = tk.StringVar(value="robocopy_log.txt")
        ttk.Entry(log_frame, textvariable=self.log_file_var, width=40, style='Modern.TEntry').grid(row=0, column=1, padx=(10, 5), sticky=(tk.W, tk.E))
        ttk.Button(log_frame, text="📂 参照", command=self.browse_log_file, style='Secondary.TButton').grid(row=0, column=2)
        
        # 追加オプション欄
        custom_frame = ttk.Frame(config_frame, style='Modern.TFrame')
        custom_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(8, 0))
        
        ttk.Label(custom_frame, text="追加オプション:", style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.custom_options_var = tk.StringVar()
        ttk.Entry(custom_frame, textvariable=self.custom_options_var, width=50, style='Modern.TEntry').grid(row=0, column=1, padx=(10, 0), sticky=(tk.W, tk.E))
        
        # デフォルト設定：全オプションを有効にする
        self.copy_mode_var.set("MIR")  # デフォルトは/MIR
        self.option_vars["no_progress"].set(True)      # /NP
        self.option_vars["no_dir_list"].set(True)      # /NDL
        self.option_vars["sec"].set(False)              # /SEC（追加）
        self.option_vars["list_only"].set(True)        # /L（テストモード）
        self.option_vars["enable_log"].set(True)       # LOG
        
        # スケジュール設定セクション
        schedule_frame = ttk.LabelFrame(main_frame, text="⏰ スケジュール設定", padding="20", style='Modern.TLabelframe')
        schedule_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        # タスク名設定（スケジュール設定の最後に追加）
        task_name_frame = ttk.Frame(schedule_frame, style='Modern.TFrame')
        task_name_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 0))
        
        ttk.Label(task_name_frame, text="タスク名:", style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Entry(task_name_frame, textvariable=self.task_name_var, width=40, style='Modern.TEntry').grid(row=0, column=1, padx=(10, 5), sticky=(tk.W, tk.E), pady=(0, 5))
        ttk.Button(task_name_frame, text="🔧 自動生成", 
                  command=self.update_task_name_from_dest, style='Secondary.TButton').grid(row=0, column=2, pady=(0, 5))
        
        # 説明ラベル
        ttk.Label(task_name_frame, text="※ コピー先フォルダを選択すると自動で生成されます", 
                 style='Subtitle.TLabel').grid(row=1, column=1, sticky=tk.W, padx=(10, 0))
        
        # 実行頻度
        ttk.Label(schedule_frame, text="実行頻度:", style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        self.frequency_var = tk.StringVar(value="毎日")
        self.frequency_combo = ttk.Combobox(schedule_frame, textvariable=self.frequency_var,
                                     values=["毎日", "毎週"], style='Modern.TCombobox')
        self.frequency_combo.grid(row=0, column=1, padx=(10, 0), pady=(0, 8))
        self.frequency_combo.state(['readonly'])
        # 頻度変更時のイベントを追加
        self.frequency_combo.bind('<<ComboboxSelected>>', self.on_frequency_changed)
        
        # 曜日選択（毎週の場合）
        self.weekday_label = ttk.Label(schedule_frame, text="曜日:", style='Modern.TLabel')
        self.weekday_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 8))
        self.weekday_var = tk.StringVar(value="月曜日")
        self.weekday_combo = ttk.Combobox(schedule_frame, textvariable=self.weekday_var,
                                   values=["月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"], style='Modern.TCombobox')
        self.weekday_combo.grid(row=1, column=1, padx=(10, 0), pady=(0, 8))
        self.weekday_combo.state(['readonly'])
        
        # 実行時刻
        ttk.Label(schedule_frame, text="実行時刻:", style='Modern.TLabel').grid(row=2, column=0, sticky=tk.W, pady=(0, 8))
        time_frame = ttk.Frame(schedule_frame, style='Modern.TFrame')
        time_frame.grid(row=2, column=1, padx=(10, 0), pady=(0, 8))
        
        self.hour_var = tk.StringVar(value="09")
        self.minute_var = tk.StringVar(value="00")
        
        ttk.Spinbox(time_frame, from_=0, to=23, textvariable=self.hour_var, 
                   width=5, format="%02.0f", style='Modern.TSpinbox').grid(row=0, column=0)
        ttk.Label(time_frame, text=" : ", style='Modern.TLabel').grid(row=0, column=1, padx=5)
        ttk.Spinbox(time_frame, from_=0, to=59, textvariable=self.minute_var, 
                   width=5, format="%02.0f", style='Modern.TSpinbox').grid(row=0, column=2)
        
        # メール設定セクション
        email_frame = ttk.LabelFrame(main_frame, text="📧 メール通知設定", padding="20", style='Modern.TLabelframe')
        email_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

        # メール送信有効化チェックボックス
        self.email_enabled_var = tk.BooleanVar()
        email_check = ttk.Checkbutton(email_frame, text="実行結果をメールで送信", 
                                    variable=self.email_enabled_var,
                                    command=self.toggle_email_settings, style='Modern.TCheckbutton')
        email_check.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 15))

        # SMTPサーバー設定
        self.email_settings_frame = ttk.Frame(email_frame, style='Modern.TFrame')
        self.email_settings_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))

        ttk.Label(self.email_settings_frame, text="SMTPサーバー:", style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        self.smtp_server_var = tk.StringVar(value="smtp.gmail.com")
        ttk.Entry(self.email_settings_frame, textvariable=self.smtp_server_var, width=30, style='Modern.TEntry').grid(row=0, column=1, padx=(10, 5), pady=(0, 8))

        ttk.Label(self.email_settings_frame, text="ポート:", style='Modern.TLabel').grid(row=0, column=2, sticky=tk.W, padx=(15, 0), pady=(0, 8))
        self.smtp_port_var = tk.StringVar(value="587")
        ttk.Entry(self.email_settings_frame, textvariable=self.smtp_port_var, width=10, style='Modern.TEntry').grid(row=0, column=3, padx=(10, 0), pady=(0, 8))

        # 接続の保護プルダウン（修正版）
        ttk.Label(self.email_settings_frame, text="接続の保護:", style='Modern.TLabel').grid(row=1, column=0, sticky=tk.W, pady=(0, 8))
        self.connection_security_var = tk.StringVar(value="STARTTLS")
        self.connection_security_combo = ttk.Combobox(self.email_settings_frame, 
                                                    textvariable=self.connection_security_var,
                                                    values=["暗号化なし", "STARTTLS", "SSL/TLS"],
                                                    width=15,
                                                    state='readonly', style='Modern.TCombobox')  # この書き方に変更
        self.connection_security_combo.grid(row=1, column=1, padx=(10, 5), pady=(0, 8))

        # 認証方式プルダウン（修正版）
        ttk.Label(self.email_settings_frame, text="認証方式:", style='Modern.TLabel').grid(row=1, column=2, sticky=tk.W, padx=(15, 0), pady=(0, 8))
        self.auth_method_var = tk.StringVar(value="CRAM-MD5")
        self.auth_method_combo = ttk.Combobox(self.email_settings_frame, 
                                            textvariable=self.auth_method_var,
                                            values=["CRAM-MD5", "LOGIN", "PLAIN", "DIGEST-MD5"],
                                            width=15,
                                            state='readonly', style='Modern.TCombobox')  # この書き方に変更
        self.auth_method_combo.grid(row=1, column=3, padx=(10, 0), pady=(0, 8))

        ttk.Label(self.email_settings_frame, text="送信者メール:", style='Modern.TLabel').grid(row=2, column=0, sticky=tk.W, pady=(0, 8))
        self.sender_email_var = tk.StringVar()
        ttk.Entry(self.email_settings_frame, textvariable=self.sender_email_var, width=40, style='Modern.TEntry').grid(row=2, column=1, columnspan=2, padx=(10, 0), pady=(0, 8))

        ttk.Label(self.email_settings_frame, text="送信者パスワード:", style='Modern.TLabel').grid(row=3, column=0, sticky=tk.W, pady=(0, 8))
        self.sender_password_var = tk.StringVar()
        ttk.Entry(self.email_settings_frame, textvariable=self.sender_password_var, 
                width=40, show="*", style='Modern.TEntry').grid(row=3, column=1, columnspan=2, padx=(10, 0), pady=(0, 8))

        ttk.Label(self.email_settings_frame, text="送信先メール:", style='Modern.TLabel').grid(row=4, column=0, sticky=tk.W, pady=(0, 15))
        self.recipient_email_var = tk.StringVar()
        ttk.Entry(self.email_settings_frame, textvariable=self.recipient_email_var, 
                width=40, style='Modern.TEntry').grid(row=4, column=1, columnspan=2, padx=(10, 0), pady=(0, 15))

        # SMTP接続テストボタン
        smtp_test_frame = ttk.Frame(self.email_settings_frame, style='Modern.TFrame')
        smtp_test_frame.grid(row=5, column=0, columnspan=5, pady=(10, 0))

        ttk.Button(smtp_test_frame, text="🔍 SMTP接続テスト", 
                command=self.test_smtp_connection, style='Secondary.TButton').grid(row=0, column=0, padx=(0, 10))
        ttk.Button(smtp_test_frame, text="📧 メールテスト", 
                command=self.send_email_with_flexible_auth, style='Secondary.TButton').grid(row=0, column=1)  # 新しいメソッド名

        # 初期状態ではメール設定を無効化
        self.toggle_email_settings()
        
        # 初期状態では認証設定を無効化
        self.disable_auth_widgets(self.source_auth_widgets)
        self.disable_auth_widgets(self.dest_auth_widgets)
        
        # 制御ボタンセクション
        button_frame = ttk.Frame(main_frame, style='Modern.TFrame')
        button_frame.grid(row=5, column=0, columnspan=2, pady=(25, 20))

        # 上段ボタン
        top_button_frame = ttk.Frame(button_frame, style='Modern.TFrame')
        top_button_frame.grid(row=0, column=0, columnspan=4, pady=(0, 10))

        ttk.Button(top_button_frame, text="💾 設定を保存", 
                  command=self.save_config, style='Primary.TButton').grid(row=0, column=0, padx=5)
        ttk.Button(top_button_frame, text="▶️ 今すぐ実行", 
                  command=self.run_now, style='Success.TButton').grid(row=0, column=1, padx=5)
        ttk.Button(top_button_frame, text="🧪 バッチテスト", 
                  command=self.test_batch_script, style='Warning.TButton').grid(row=0, column=2, padx=5)

        # 下段ボタン
        bottom_button_frame = ttk.Frame(button_frame, style='Modern.TFrame')
        bottom_button_frame.grid(row=1, column=0, columnspan=4, pady=(0, 0))

        ttk.Button(bottom_button_frame, text="📅 タスク作成/更新", 
                  command=self.create_scheduled_task, style='Primary.TButton').grid(row=0, column=0, padx=5)
        ttk.Button(bottom_button_frame, text="🗑️ タスク削除", 
                  command=self.delete_scheduled_task, style='Danger.TButton').grid(row=0, column=1, padx=5)
        
        # タスクステータス表示
        status_frame = ttk.LabelFrame(main_frame, text="📊 タスクステータス", padding="20", style='Modern.TLabelframe')
        status_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

        # タスク履歴設定セクション
        history_frame = ttk.LabelFrame(main_frame, text="📅 タスク履歴設定", padding="20", style='Modern.TLabelframe')
        history_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

        # 履歴有効化チェックボックス
        self.history_enabled_var = tk.BooleanVar()
        ttk.Checkbutton(history_frame, text="すべてのタスク履歴を有効にする", 
                    variable=self.history_enabled_var, style='Modern.TCheckbutton').grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

        history_button_frame = ttk.Frame(history_frame, style='Modern.TFrame')
        history_button_frame.grid(row=1, column=0, pady=(0, 0))

        ttk.Button(history_button_frame, text="⚙️ 履歴設定を適用", 
                command=self.apply_task_history, style='Primary.TButton').grid(row=0, column=0, padx=(0, 10))
        ttk.Button(history_button_frame, text="🔍 現在の設定確認", 
                command=self.check_task_history, style='Secondary.TButton').grid(row=0, column=1)
        
        self.task_status_var = tk.StringVar(value="確認中...")
        ttk.Label(status_frame, textvariable=self.task_status_var, style='Modern.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Button(status_frame, text="🔄 ステータス更新", 
                  command=self.update_task_status, style='Secondary.TButton').grid(row=0, column=1, padx=(15, 0), pady=(0, 5))
        
        # ログ表示エリア
        log_frame = ttk.LabelFrame(main_frame, text="📜 実行ログ", padding="20", style='Modern.TLabelframe')
        log_frame.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
        # ダークテーマ対応のテキストエリア
        self.log_text = tk.Text(log_frame, height=12, width=100, 
                              bg=self.colors['bg_tertiary'], fg=self.colors['text'], 
                              insertbackground=self.colors['text'], selectbackground=self.colors['primary'],
                              font=('メイリオ', 9), relief='solid', borderwidth=1)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # ステータスバー
        self.status_var = tk.StringVar(value="準備完了")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, style='Modern.TLabel', 
                             relief=tk.SUNKEN, padding=(10, 5))
        status_bar.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 0))

        # 初期状態で曜日選択の有効/無効を設定
        self.update_weekday_state()
    
    def on_frequency_changed(self, event=None):
        """実行頻度が変更されたときの処理"""
        self.update_weekday_state()

    def update_weekday_state(self):
        """曜日選択の有効/無効を更新"""
        if self.frequency_var.get() == "毎日":
            # 毎日の場合は曜日選択を無効化
            self.weekday_combo.configure(state='disabled')
            self.weekday_label.configure(foreground='gray')
        else:
            # 毎週の場合は曜日選択を有効化
            self.weekday_combo.configure(state='readonly')
            self.weekday_label.configure(foreground='black')

    def get_frequency_code(self):
        """表示文字列を内部コードに変換"""
        frequency_map = {
            "毎日": "DAILY",
            "毎週": "WEEKLY"
        }
        return frequency_map.get(self.frequency_var.get(), "DAILY")

    def get_weekday_code(self):
        """表示文字列を内部コードに変換"""
        weekday_map = {
            "月曜日": "MON",
            "火曜日": "TUE", 
            "水曜日": "WED",
            "木曜日": "THU",
            "金曜日": "FRI",
            "土曜日": "SAT",
            "日曜日": "SUN"
        }
        return weekday_map.get(self.weekday_var.get(), "MON")

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
        """選択されたオプションからRobocopyのオプション文字列を構築（ログファイル競合対策版）"""
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
        
        # ログファイル設定は削除（バッチファイル内でリダイレクトするため）
        # 以前の /LOG+: オプションは使用しない
        
        # カスタムオプションを追加
        custom_options = self.custom_options_var.get().strip()
        if custom_options:
            options.append(custom_options)
        
        return " ".join(options)
    
    def build_robocopy_options_for_gui(self):
        """GUI実行用のRobocopyオプション（日時付きログファイル）"""
        options = []
        
        # 基本オプションを取得
        base_options = self.build_robocopy_options()
        if base_options:
            options.append(base_options)
        
        # GUI実行時のみログファイル設定を追加
        if self.option_vars["enable_log"].get():
            base_log_file = self.log_file_var.get() if self.log_file_var.get() else "robocopy_gui_log"
            if base_log_file.endswith('.txt'):
                base_log_file = base_log_file[:-4]  # .txt を除去
            
            # 日時付きログファイル名を生成
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            unique_log_file = f"{base_log_file}_{timestamp}.txt"
            
            # ログファイルパスにスペースが含まれる場合はクォートで囲む
            if " " in unique_log_file:
                options.append(f'/LOG+:"{unique_log_file}"')
            else:
                options.append(f"/LOG+:{unique_log_file}")
            
            # GUI用のログファイル名を記録（メール送信で使用するため）
            self._current_gui_log_file = unique_log_file
        
        return " ".join(options)
    
    def escape_batch_password(self, password):
        """バッチファイル用にパスワードの特殊文字をエスケープ"""
        if not password:
            return ""
        
        # バッチファイルで特別な意味を持つ文字をエスケープ
        escaped = password
        
        # ^文字は最初にエスケープ（他のエスケープ文字に影響しないように）
        escaped = escaped.replace('^', '^^')
        
        # %文字をエスケープ（変数展開を防ぐ）
        escaped = escaped.replace('%', '%%')
        
        # &文字をエスケープ（コマンド区切りを防ぐ）
        escaped = escaped.replace('&', '^&')
        
        # <、>文字をエスケープ（リダイレクトを防ぐ）
        escaped = escaped.replace('<', '^<')
        escaped = escaped.replace('>', '^>')
        
        # |文字をエスケープ（パイプを防ぐ）
        escaped = escaped.replace('|', '^|')
        
        # "文字をエスケープ
        escaped = escaped.replace('"', '""')
        
        # スペースが含まれている場合は全体をダブルクォートで囲む
        if ' ' in escaped:
            escaped = f'"{escaped}"'
        
        return escaped

    def escape_batch_string(self, text):
        """バッチファイル用に文字列の特殊文字をエスケープ（パスワード以外用）"""
        if not text:
            return ""
        
        escaped = text
        
        # ^文字は最初にエスケープ
        escaped = escaped.replace('^', '^^')
        
        # %文字をエスケープ
        escaped = escaped.replace('%', '%%')
        
        # &文字をエスケープ
        escaped = escaped.replace('&', '^&')
        
        # <、>文字をエスケープ
        escaped = escaped.replace('<', '^<')
        escaped = escaped.replace('>', '^>')
        
        # |文字をエスケープ
        escaped = escaped.replace('|', '^|')
        
        return escaped

    def escape_powershell_string(self, text):
        """PowerShell用に文字列をエスケープ"""
        if not text:
            return ""
        
        # PowerShellで特別な意味を持つ文字をエスケープ
        escaped = text
        
        # バックスラッシュを最初にエスケープ
        escaped = escaped.replace('\\', '\\\\')
        
        # ダブルクォートをエスケープ
        escaped = escaped.replace('"', '`"')
        
        # シングルクォートをエスケープ
        escaped = escaped.replace("'", "''")
        
        # バッククォートをエスケープ
        escaped = escaped.replace('`', '``')
        
        # ドル記号をエスケープ（変数展開を防ぐ）
        escaped = escaped.replace('$', '`$')
        
        return escaped

    def test_batch_script(self):
        """生成されるバッチファイルをテスト実行"""
        if not self.source_var.get() or not self.dest_var.get():
            messagebox.showerror("エラー", "コピー元とコピー先を設定してください")
            return
        
        task_name = self.task_name_var.get().strip()
        if not task_name:
            messagebox.showerror("エラー", "タスク名を入力してください")
            return
        
        try:
            # 確認ダイアログ
            response = messagebox.askyesno("確認", 
                "バッチスクリプトのテスト実行を行います。\n"
                "この操作では実際にrobocopyが実行されます。\n\n"
                "続行しますか？")
            
            if not response:
                return
            
            self.log_message("バッチスクリプトのテスト実行を開始します...")
            self.status_var.set("テスト実行中...")
            
            # 一時的なバッチファイルを生成
            test_batch_name = f"{task_name}_test"
            test_batch_path = self.generate_batch_script(test_batch_name)
            
            self.log_message(f"テスト用バッチファイルを生成: {test_batch_path}")
            
            # バッチファイルを実行
            self.log_message("バッチファイルを実行中...")
            
            # バッチファイルを新しいコマンドプロンプトで実行（出力を確認できるように）
            cmd = f'start "Robocopy Test" /wait cmd /c ""{test_batch_path}" & echo. & echo 実行完了。何かキーを押すと閉じます。 & pause"'
            
            result = subprocess.run(cmd, shell=True, capture_output=False)
            
            self.log_message("バッチファイルのテスト実行が完了しました")
            self.log_message("詳細な結果はコマンドプロンプトウィンドウとログファイルを確認してください")
            
            # テスト用ファイルを削除するか確認
            file_list = [f"・{os.path.basename(test_batch_path)}"]
            if self.email_enabled_var.get():
                ps_file_name = f"{test_batch_name}_mail.ps1"
                file_list.append(f"・{ps_file_name}")

            cleanup_response = messagebox.askyesno("テストファイル削除", 
                f"テスト用に生成されたファイルを削除しますか？\n\n" + 
                "\n".join(file_list))
            
            if cleanup_response:
                try:
                    # バッチファイルを削除
                    if os.path.exists(test_batch_path):
                        os.remove(test_batch_path)
                        self.log_message(f"テスト用バッチファイルを削除: {test_batch_path}")
                    
                    # PowerShellスクリプトを削除（メール送信が有効な場合）
                    if self.email_enabled_var.get():
                        ps_file = f"{test_batch_name}_mail.ps1"
                        if os.path.exists(ps_file):
                            os.remove(ps_file)
                            self.log_message(f"テスト用PowerShellスクリプトを削除: {ps_file}")
                            
                except Exception as e:
                    self.log_message(f"テストファイル削除エラー: {str(e)}", "error")
            
            self.status_var.set("テスト実行完了")
            
            messagebox.showinfo("テスト完了", 
                "バッチスクリプトのテスト実行が完了しました。\n"
                "実行結果を確認して、問題がなければタスクスケジューラに登録してください。")
            
        except Exception as e:
            error_msg = f"バッチテストエラー: {str(e)}"
            self.log_message(error_msg, "error")
            messagebox.showerror("エラー", error_msg)
            self.status_var.set("テスト実行エラー")
    
    def run_robocopy(self):
        """Robocopyを実行（GUI用）"""
        source = self.source_var.get()
        dest = self.dest_var.get()
        # GUI実行時は専用のオプション関数を使用
        options = self.build_robocopy_options_for_gui()
        
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
        """メールを送信（フレキシブル対応実運用版）"""
        if not self.email_enabled_var.get():
            return
        
        try:
            self.log_message("メール通知送信中...")
            
            # フレキシブル設定でPowerShellスクリプトを生成
            ps_script_path = self.generate_flexible_mail_script(success, message)
            
            # PowerShellスクリプトを実行
            cmd = f'powershell -ExecutionPolicy Bypass -File "{ps_script_path}"'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932', timeout=60)
            
            if result.returncode == 0:
                self.log_message("メール通知送信完了")
                # 詳細ログは出力しない（実運用のため）
            else:
                self.log_message("メール通知送信失敗", "error")
                if result.stderr:
                    for line in result.stderr.strip().split('\n'):
                        if line.strip():
                            self.log_message(f"  ERROR: {line.strip()}", "error")
            
            # 一時ファイルを削除
            try:
                os.unlink(ps_script_path)
            except:
                pass
                            
        except subprocess.TimeoutExpired:
            self.log_message("メール通知送信タイムアウト", "error")
        except Exception as e:
            self.log_message(f"メール送信エラー: {str(e)}", "error")
    
    def run_now(self):
        """今すぐrobocopyを実行"""
        self.status_var.set("実行中...")
        success, message = self.run_robocopy()
        
        if self.email_enabled_var.get():
            self.send_email(success, message)
        
        self.status_var.set("実行完了" if success else "実行エラー")
    
    def create_scheduled_task(self):
        """Windowsタスクスケジューラにタスクを作成（フレキシブル対応版）"""
        if not self.source_var.get() or not self.dest_var.get():
            messagebox.showerror("エラー", "コピー元とコピー先を設定してください")
            return
        
        task_name = self.task_name_var.get().strip()
        if not task_name:
            messagebox.showerror("エラー", "タスク名を入力してください")
            return
        
        try:
            # 設定を保存してからタスクを作成
            self.save_config()
            
            # バッチファイルを生成
            self.log_message("バッチファイルを生成中...")
            batch_path = self.generate_batch_script(task_name)
            
            # メール送信が有効な場合、フレキシブルメールスクリプトも生成
            generated_files = [batch_path]
            if self.email_enabled_var.get():
                self.log_message("メール送信スクリプトを生成中...")
                try:
                    # バッチ専用のメールスクリプトを生成
                    mail_script_path = self.generate_batch_mail_script(task_name)
                    generated_files.append(mail_script_path)
                    self.log_message(f"メールスクリプト生成完了: {mail_script_path}")
                except Exception as e:
                    self.log_message(f"メールスクリプト生成エラー: {str(e)}", "error")
                    messagebox.showwarning("警告", 
                        f"メールスクリプトの生成に失敗しました: {str(e)}\n"
                        "バックアップは実行されますが、メール通知は送信されません。")
            
            # タスクの実行時刻
            start_time = f"{self.hour_var.get()}:{self.minute_var.get()}"
            
            # スケジュール頻度に応じてコマンドを構築
            frequency_code = self.get_frequency_code()
            if frequency_code == "DAILY":
                schedule_type = "/SC DAILY"
            else:  # WEEKLY
                weekday_code = self.get_weekday_code()
                schedule_type = f"/SC WEEKLY /D {weekday_code}"
            
            # schtasksコマンドを構築
            cmd = f'''schtasks /CREATE /TN "{task_name}" /TR "\\"{batch_path}\\"" {schedule_type} /ST {start_time} /F'''
            
            self.log_message("タスクスケジューラに登録中...")
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            if result.returncode == 0:
                self.log_message(f"タスクを作成しました: {task_name}")
                self.log_message(f"実行ファイル: {batch_path}")
                
                file_list = "\n".join([f"・{os.path.basename(f)}" for f in generated_files])
                messagebox.showinfo("成功", 
                    f"スケジュールタスク '{task_name}' を作成しました\n\n"
                    f"生成されたファイル:\n{file_list}")
                self.update_task_status()
            else:
                error_msg = f"タスク作成エラー: {result.stderr}"
                self.log_message(error_msg, "error")
                messagebox.showerror("エラー", error_msg)
                
        except Exception as e:
            error_msg = f"タスク作成エラー: {str(e)}"
            self.log_message(error_msg, "error")
            messagebox.showerror("エラー", error_msg)

    def delete_scheduled_task(self):
        """スケジュールされたタスクを削除（関連ファイルも削除）"""
        task_name = self.task_name_var.get().strip()
        if not task_name:
            messagebox.showerror("エラー", "タスク名が入力されていません")
            return
        
        # 削除確認
        response = messagebox.askyesno("確認", 
            f"タスク '{task_name}' とその関連ファイルを削除しますか？\n\n"
            f"削除されるファイル:\n"
            f"・{task_name}.bat\n"
            f"・{task_name}_batch_mail.ps1（存在する場合）")
        
        if not response:
            return
            
        try:
            # タスクスケジューラからタスクを削除
            cmd = f'schtasks /DELETE /TN "{task_name}" /F'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            task_deleted = False
            if result.returncode == 0:
                self.log_message(f"タスクを削除しました: {task_name}")
                task_deleted = True
            else:
                # タスクが存在しない場合でもファイル削除は続行
                error_msg = f"タスク削除エラー: {result.stderr}"
                self.log_message(error_msg, "error")
            
            # 関連ファイルを削除
            files_deleted = []
            files_failed = []
            
            # バッチファイルを削除
            batch_file = f"{task_name}.bat"
            if os.path.exists(batch_file):
                try:
                    os.remove(batch_file)
                    files_deleted.append(batch_file)
                    self.log_message(f"バッチファイルを削除しました: {batch_file}")
                except Exception as e:
                    files_failed.append(f"{batch_file} ({str(e)})")
                    self.log_message(f"バッチファイル削除エラー: {batch_file} - {str(e)}", "error")
            
            # PowerShellスクリプトを削除
            ps_file = f"{task_name}_batch_mail.ps1"
            if os.path.exists(ps_file):
                try:
                    os.remove(ps_file)
                    files_deleted.append(ps_file)
                    self.log_message(f"PowerShellスクリプトを削除しました: {ps_file}")
                except Exception as e:
                    files_failed.append(f"{ps_file} ({str(e)})")
                    self.log_message(f"PowerShellスクリプト削除エラー: {ps_file} - {str(e)}", "error")
            
            # 結果をユーザーに通知
            message_parts = []
            if task_deleted:
                message_parts.append(f"スケジュールタスク '{task_name}' を削除しました")
            
            if files_deleted:
                message_parts.append(f"削除されたファイル:\n" + "\n".join([f"・{f}" for f in files_deleted]))
            
            if files_failed:
                message_parts.append(f"削除に失敗したファイル:\n" + "\n".join([f"・{f}" for f in files_failed]))
            
            if task_deleted or files_deleted:
                messagebox.showinfo("削除完了", "\n\n".join(message_parts))
            else:
                messagebox.showwarning("警告", "削除対象が見つかりませんでした")
            
            self.update_task_status()
            
        except Exception as e:
            error_msg = f"削除処理エラー: {str(e)}"
            self.log_message(error_msg, "error")
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
            'task_name': self.task_name_var.get(),
            'copy_mode': self.copy_mode_var.get(),
            'robocopy_options': {var_name: var.get() for var_name, var in self.option_vars.items()},
            'retry_count': self.retry_var.get(),
            'wait_time': self.wait_var.get(),
            'log_file': self.log_file_var.get(),
            'custom_options': self.custom_options_var.get(),
            'frequency': self.get_frequency_code(),
            'weekday': self.get_weekday_code(),
            'hour': self.hour_var.get(),
            'minute': self.minute_var.get(),
            'email_enabled': self.email_enabled_var.get(),
            'smtp_server': self.smtp_server_var.get(),
            'smtp_port': self.smtp_port_var.get(),
            'connection_security': self.connection_security_var.get(),  # 新規追加
            'auth_method': self.auth_method_var.get(),                  # 新規追加
            'sender_email': self.sender_email_var.get(),
            'sender_password': self.sender_password_var.get(),
            'recipient_email': self.recipient_email_var.get(),
            'history_enabled': self.history_enabled_var.get(),
            # 認証情報
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
            self.log_message(error_msg, "error")
            messagebox.showerror("エラー", error_msg)

    def load_config(self):
        """設定をファイルから読み込み（新旧設定互換版）"""
        if not os.path.exists(self.config_file):
            return
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 基本設定の読み込み
            self.source_var.set(config.get('source', ''))
            self.dest_var.set(config.get('dest', ''))
            
            # タスク名の読み込み
            saved_task_name = config.get('task_name', '')
            if saved_task_name:
                self.task_name_var.set(saved_task_name)
            else:
                self.update_task_name_from_dest()
            
            # コピーモードの読み込み
            self.copy_mode_var.set(config.get('copy_mode', 'MIR'))
            
            # オプション設定の読み込み
            robocopy_options = config.get('robocopy_options', {})
            for var_name, var in self.option_vars.items():
                var.set(robocopy_options.get(var_name, False))
            
            self.retry_var.set(config.get('retry_count', '1'))
            self.wait_var.set(config.get('wait_time', '1'))
            self.log_file_var.set(config.get('log_file', 'robocopy_log.txt'))
            self.custom_options_var.set(config.get('custom_options', ''))
            
            # スケジュール設定の読み込み
            frequency_code = config.get('frequency', 'DAILY')
            weekday_code = config.get('weekday', 'MON')
            
            frequency_display_map = {"DAILY": "毎日", "WEEKLY": "毎週"}
            weekday_display_map = {
                "MON": "月曜日", "TUE": "火曜日", "WED": "水曜日",
                "THU": "木曜日", "FRI": "金曜日", "SAT": "土曜日", "SUN": "日曜日"
            }
            
            self.frequency_var.set(frequency_display_map.get(frequency_code, "毎日"))
            self.weekday_var.set(weekday_display_map.get(weekday_code, "月曜日"))
            self.hour_var.set(config.get('hour', '09'))
            self.minute_var.set(config.get('minute', '00'))
            
            # メール設定の読み込み（新旧互換性対応）
            self.email_enabled_var.set(config.get('email_enabled', False))
            self.smtp_server_var.set(config.get('smtp_server', 'smtp.gmail.com'))
            self.smtp_port_var.set(config.get('smtp_port', '587'))
            self.sender_email_var.set(config.get('sender_email', ''))
            self.sender_password_var.set(config.get('sender_password', ''))
            self.recipient_email_var.set(config.get('recipient_email', ''))
            
            # 新しい設定（接続の保護・認証方式）の読み込み
            if 'connection_security' in config:
                # 新しい設定ファイルの場合
                self.connection_security_var.set(config.get('connection_security', 'STARTTLS'))
                self.auth_method_var.set(config.get('auth_method', 'CRAM-MD5'))
            else:
                # 古い設定ファイルの場合は変換
                old_use_ssl = config.get('use_ssl', False)
                old_smtp_port = config.get('smtp_port', '587')
                
                # 古いSSL設定を新しい接続の保護設定に変換
                if old_use_ssl:
                    if old_smtp_port == '465':
                        self.connection_security_var.set('SSL/TLS')
                    else:
                        self.connection_security_var.set('SSL/TLS')  # 明示的にSSLが有効だった場合
                else:
                    if old_smtp_port == '587':
                        self.connection_security_var.set('STARTTLS')  # ポート587の一般的な設定
                    else:
                        self.connection_security_var.set('暗号化なし')
                
                # 認証方式はデフォルトをCRAM-MD5に設定
                self.auth_method_var.set('CRAM-MD5')
                
                self.log_message("古い設定ファイルを新しい形式に変換しました")
            
            # タスク履歴設定
            self.history_enabled_var.set(config.get('history_enabled', False))
            
            # 認証情報の読み込み
            self.source_auth_enabled_var.set(config.get('source_auth_enabled', False))
            self.source_username_var.set(config.get('source_username', ''))
            self.source_password_var.set(config.get('source_password', ''))
            self.source_domain_var.set(config.get('source_domain', ''))
            self.dest_auth_enabled_var.set(config.get('dest_auth_enabled', False))
            self.dest_username_var.set(config.get('dest_username', ''))
            self.dest_password_var.set(config.get('dest_password', ''))
            self.dest_domain_var.set(config.get('dest_domain', ''))
            
            # UI状態の更新
            self.toggle_email_settings()
            self.update_auth_state()
            self.update_weekday_state()
            
            self.log_message("設定を読み込みました")
            
        except Exception as e:
            self.log_message(f"設定読み込みエラー: {str(e)}", "error")
            # エラーが発生した場合はデフォルト設定を適用
            self.connection_security_var.set('STARTTLS')
            self.auth_method_var.set('CRAM-MD5')
            self.log_message("デフォルト設定を適用しました")
    
    def test_smtp_connection(self):
        """SMTP接続をテスト（フレキシブル対応版）"""
        smtp_server = self.smtp_server_var.get().strip()
        smtp_port = self.smtp_port_var.get().strip()
        
        if not smtp_server:
            messagebox.showerror("エラー", "SMTPサーバーを入力してください")
            return
        
        if not smtp_port:
            messagebox.showerror("エラー", "ポート番号を入力してください")
            return
        
        try:
            port_num = int(smtp_port)
            if port_num < 1 or port_num > 65535:
                raise ValueError()
        except ValueError:
            messagebox.showerror("エラー", "有効なポート番号を入力してください（1-65535）")
            return
        
        try:
            self.log_message(f"SMTP接続テスト開始: {smtp_server}:{smtp_port}")
            
            # 接続の保護設定を取得
            connection_security = self.connection_security_var.get()
            use_ssl = connection_security == "SSL/TLS"
            
            # PowerShellスクリプトでSMTP接続をテスト
            ps_script_content = f'''# SMTP接続テスト（フレキシブル版）
    $ErrorActionPreference = "Stop"

    $SmtpServer = "{self.escape_powershell_string(smtp_server)}"
    $SmtpPort = {smtp_port}
    $UseSSL = {"$true" if use_ssl else "$false"}
    $ConnectionSecurity = "{connection_security}"

    Write-Output "SMTP接続テスト開始: $SmtpServer`:$SmtpPort"
    Write-Output "接続の保護: $ConnectionSecurity"

    try {{
        # TCP接続テスト
        $tcpTest = Test-NetConnection -ComputerName $SmtpServer -Port $SmtpPort -WarningAction SilentlyContinue
        
        if ($tcpTest.TcpTestSucceeded) {{
            Write-Output "TCP接続成功"
            
            # SMTP接続テスト
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $tcpClient.ReceiveTimeout = 10000
            $tcpClient.SendTimeout = 10000
            $tcpClient.Connect($SmtpServer, $SmtpPort)
            
            if ($tcpClient.Connected) {{
                $stream = $tcpClient.GetStream()
                
                # SSL/TLS処理
                if ($UseSSL) {{
                    Write-Output "SSL/TLS暗号化開始..."
                    $sslStream = New-Object System.Net.Security.SslStream($stream)
                    $sslStream.AuthenticateAsClient($SmtpServer)
                    $stream = $sslStream
                }}
                
                $reader = New-Object System.IO.StreamReader($stream)
                $writer = New-Object System.IO.StreamWriter($stream)
                
                # SMTPサーバーからの初期レスポンス
                $response = $reader.ReadLine()
                Write-Output "SMTPサーバー応答: $response"
                
                # STARTTLS確認
                if ($ConnectionSecurity -eq "STARTTLS" -and -not $UseSSL) {{
                    Write-Output "EHLO送信（STARTTLS確認用）..."
                    $writer.WriteLine("EHLO localhost")
                    $writer.Flush()
                    
                    do {{
                        $response = $reader.ReadLine()
                        Write-Output "EHLO応答: $response"
                        if ($response -match "STARTTLS") {{
                            Write-Output "STARTTLS対応確認"
                        }}
                    }} while ($response.StartsWith("250-"))
                }}
                
                # QUITで終了
                $writer.WriteLine("QUIT")
                $writer.Flush()
                
                $stream.Close()
                $tcpClient.Close()
                
                Write-Output "SMTP接続テスト成功"
                exit 0
            }}
        }} else {{
            Write-Output "TCP接続に失敗しました"
            exit 1
        }}
    }} catch {{
        Write-Output "接続テストエラー: $($_.Exception.Message)"
        exit 1
    }}'''
            
            # 一時的なPowerShellスクリプトファイルを作成
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.ps1', delete=False, encoding='cp932') as temp_file:
                temp_file.write(ps_script_content)
                temp_ps_path = temp_file.name
            
            try:
                # PowerShellスクリプトを実行
                cmd = f'powershell -ExecutionPolicy Bypass -File "{temp_ps_path}"'
                result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
                
                # 結果をログに出力
                if result.stdout:
                    for line in result.stdout.strip().split('\n'):
                        if line.strip():
                            self.log_message(f"  {line.strip()}")
                
                if result.stderr:
                    for line in result.stderr.strip().split('\n'):
                        if line.strip():
                            self.log_message(f"  ERROR: {line.strip()}", "error")
                
                if result.returncode == 0:
                    self.log_message("SMTP接続テスト成功", "success")
                    messagebox.showinfo("接続テスト成功", 
                        f"SMTPサーバーへの接続に成功しました。\n\n"
                        f"サーバー: {smtp_server}\n"
                        f"ポート: {smtp_port}\n"
                        f"接続の保護: {connection_security}")
                else:
                    self.log_message("SMTP接続テスト失敗", "error")
                    messagebox.showerror("接続テスト失敗", 
                        f"SMTPサーバーへの接続に失敗しました。\n\n"
                        f"詳細はログを確認してください。")
            
            finally:
                # 一時ファイルを削除
                try:
                    os.unlink(temp_ps_path)
                except:
                    pass
                    
        except Exception as e:
            error_msg = f"SMTP接続テストエラー: {str(e)}"
            self.log_message(error_msg, "error")
            messagebox.showerror("エラー", error_msg)

    def send_email_with_flexible_auth(self):
        """選択された認証方式でメール送信"""
        # 入力値の検証
        required_fields = [
            (self.smtp_server_var.get().strip(), "SMTPサーバー"),
            (self.smtp_port_var.get().strip(), "ポート番号"),
            (self.sender_email_var.get().strip(), "送信者メール"),
            (self.sender_password_var.get().strip(), "送信者パスワード"),
            (self.recipient_email_var.get().strip(), "送信先メール")
        ]
        
        for value, field_name in required_fields:
            if not value:
                messagebox.showerror("エラー", f"{field_name}を入力してください")
                return
        
        try:
            port_num = int(self.smtp_port_var.get().strip())
            if port_num < 1 or port_num > 65535:
                raise ValueError()
        except ValueError:
            messagebox.showerror("エラー", "有効なポート番号を入力してください（1-65535）")
            return
        
        # 設定確認
        connection_security = self.connection_security_var.get()
        auth_method = self.auth_method_var.get()
        
        response = messagebox.askyesno("確認", 
            f"以下の設定でメールを送信します。\n\n"
            f"SMTPサーバー: {self.smtp_server_var.get()}:{self.smtp_port_var.get()}\n"
            f"接続の保護: {connection_security}\n"
            f"認証方式: {auth_method}\n"
            f"送信先: {self.recipient_email_var.get()}\n\n"
            f"続行しますか？")
        
        if not response:
            return
        
        try:
            self.log_message(f"フレキシブル認証メール送信開始 ({connection_security} + {auth_method})...")
            
            # PowerShellスクリプトを生成
            ps_script_path = self.generate_flexible_mail_script(True, "テストメール送信")
            
            # PowerShellスクリプトを実行
            cmd = f'powershell -ExecutionPolicy Bypass -File "{ps_script_path}"'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932', timeout=60)
            
            # 結果をログに出力
            if result.stdout:
                for line in result.stdout.strip().split('\n'):
                    if line.strip():
                        self.log_message(f"  {line.strip()}")
            
            if result.stderr:
                for line in result.stderr.strip().split('\n'):
                    if line.strip():
                        self.log_message(f"  ERROR: {line.strip()}", "error")
            
            if result.returncode == 0:
                self.log_message("フレキシブル認証メール送信成功", "success")
                messagebox.showinfo("送信成功", 
                    f"メール送信に成功しました！\n\n"
                    f"接続方式: {connection_security}\n"
                    f"認証方式: {auth_method}\n\n"
                    f"送信先メールボックスを確認してください。")
            else:
                self.log_message("フレキシブル認証メール送信失敗", "error")
                messagebox.showerror("送信失敗", 
                    f"メール送信に失敗しました。\n\n"
                    f"詳細はログを確認してください。")
        
        except subprocess.TimeoutExpired:
            error_msg = "メール送信がタイムアウトしました（60秒）"
            self.log_message(error_msg, "error")
            messagebox.showerror("タイムアウト", error_msg)
        except Exception as e:
            error_msg = f"フレキシブル認証メール送信エラー: {str(e)}"
            self.log_message(error_msg, "error")
            messagebox.showerror("エラー", error_msg)

    def generate_flexible_mail_script(self, success, message):
        """フレキシブル認証用メールスクリプトを生成（実運用・テスト両対応）"""
        try:
            ps_filename = f"{self.task_name_var.get()}_flexible_mail.ps1"
            ps_path = os.path.abspath(ps_filename)
            
            # 設定値を準備
            smtp_server = self.escape_powershell_string(self.smtp_server_var.get().strip())
            smtp_port = self.smtp_port_var.get().strip()
            sender_email = self.escape_powershell_string(self.sender_email_var.get().strip())
            sender_password = self.escape_powershell_string(self.sender_password_var.get().strip())
            recipient_email = self.escape_powershell_string(self.recipient_email_var.get().strip())
            
            # 接続・認証設定
            connection_security = self.connection_security_var.get()
            auth_method = self.auth_method_var.get()
            use_ssl = connection_security == "SSL/TLS"
            use_starttls = connection_security == "STARTTLS"
            
            # パス情報
            source_path = self.escape_powershell_string(self.source_var.get())
            dest_path = self.escape_powershell_string(self.dest_var.get())
            
            # メール件名（テスト用か実運用用かを判定）
            if message == "テストメール送信":
                subject = "RCScheduler テストメール - フレキシブル認証"
                is_test = True
            else:
                subject = "Robocopyバックアップ結果 - " + ("成功" if success else "失敗")
                is_test = False
            
            # エスケープされたメッセージ
            escaped_message = self.escape_powershell_string(message)
            
            # フレキシブル認証PowerShellスクリプト
            ps_content = f'''# RCScheduler フレキシブル認証メール送信
    $ErrorActionPreference = "Stop"

    # 設定
    $SmtpServer = "{smtp_server}"
    $SmtpPort = {smtp_port}
    $SenderEmail = "{sender_email}"
    $SenderPassword = "{sender_password}"
    $RecipientEmail = "{recipient_email}"
    $UseSSL = {"$true" if use_ssl else "$false"}
    $UseSTARTTLS = {"$true" if use_starttls else "$false"}
    $AuthMethod = "{auth_method}"
    $ConnectionSecurity = "{connection_security}"
    $IsTest = {"$true" if is_test else "$false"}

    if ($IsTest) {{
        Write-Output "=== フレキシブル認証テストメール送信 ==="
    }} else {{
        Write-Output "=== バックアップ結果メール送信 ==="
    }}

    function ConvertTo-Base64($text) {{
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
        return [System.Convert]::ToBase64String($bytes)
    }}

    function Send-AuthCommand($writer, $reader, $authMethod, $username, $password) {{
        switch ($authMethod) {{
            "CRAM-MD5" {{
                if ($IsTest) {{ Write-Output "CRAM-MD5認証実行中..." }}
                $writer.WriteLine("AUTH CRAM-MD5")
                $writer.Flush()
                $response = $reader.ReadLine()
                
                if ($response.StartsWith("334")) {{
                    $challenge = $response.Substring(4)
                    $challengeBytes = [System.Convert]::FromBase64String($challenge)
                    $challengeText = [System.Text.Encoding]::UTF8.GetString($challengeBytes)
                    
                    $hmac = New-Object System.Security.Cryptography.HMACMD5
                    $hmac.Key = [System.Text.Encoding]::UTF8.GetBytes($password)
                    $hash = $hmac.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($challengeText))
                    $hashHex = [System.BitConverter]::ToString($hash) -replace "-", ""
                    
                    $cramResponse = "$username " + $hashHex.ToLower()
                    $encodedResponse = ConvertTo-Base64 $cramResponse
                    $writer.WriteLine($encodedResponse)
                    $writer.Flush()
                }}
            }}
            "LOGIN" {{
                if ($IsTest) {{ Write-Output "LOGIN認証実行中..." }}
                $writer.WriteLine("AUTH LOGIN")
                $writer.Flush()
                $response = $reader.ReadLine()
                if ($response.StartsWith("334")) {{
                    $writer.WriteLine((ConvertTo-Base64 $username))
                    $writer.Flush()
                    $response = $reader.ReadLine()
                    if ($response.StartsWith("334")) {{
                        $writer.WriteLine((ConvertTo-Base64 $password))
                        $writer.Flush()
                    }}
                }}
            }}
            "PLAIN" {{
                if ($IsTest) {{ Write-Output "PLAIN認証実行中..." }}
                $authString = ConvertTo-Base64("`0$username`0$password")
                $writer.WriteLine("AUTH PLAIN $authString")
                $writer.Flush()
            }}
            "DIGEST-MD5" {{
                if ($IsTest) {{ Write-Output "DIGEST-MD5認証実行中..." }}
                $writer.WriteLine("AUTH DIGEST-MD5")
                $writer.Flush()
                $response = $reader.ReadLine()
                if ($response.StartsWith("334")) {{
                    $writer.WriteLine("")
                    $writer.Flush()
                }}
            }}
        }}
        
        $response = $reader.ReadLine()
        if ($response.StartsWith("235")) {{
            if ($IsTest) {{ Write-Output "$authMethod 認証成功" }}
            return $true
        }} else {{
            if ($IsTest) {{ Write-Output "$authMethod 認証失敗: $response" }}
            return $false
        }}
    }}

    try {{
        # TCP接続
        if ($IsTest) {{ Write-Output "TCP接続中..." }}
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.Connect($SmtpServer, $SmtpPort)
        $stream = $tcpClient.GetStream()
        
        # SSL/TLS処理
        if ($UseSSL) {{
            if ($IsTest) {{ Write-Output "SSL/TLS暗号化中..." }}
            $sslStream = New-Object System.Net.Security.SslStream($stream)
            $sslStream.AuthenticateAsClient($SmtpServer)
            $stream = $sslStream
            Start-Sleep -Milliseconds 500
        }}
        
        $reader = New-Object System.IO.StreamReader($stream)
        $writer = New-Object System.IO.StreamWriter($stream)
        
        # 初期応答
        $response = $reader.ReadLine()
        if ($IsTest) {{ Write-Output "初期応答: $response" }}
        
        # EHLO送信
        $writer.WriteLine("EHLO localhost")
        $writer.Flush()
        do {{
            $response = $reader.ReadLine()
            if ($IsTest) {{ Write-Output "EHLO応答: $response" }}
        }} while ($response.StartsWith("250-"))
        
        # STARTTLS処理
        if ($UseSTARTTLS -and -not $UseSSL) {{
            if ($IsTest) {{ Write-Output "STARTTLS実行中..." }}
            $writer.WriteLine("STARTTLS")
            $writer.Flush()
            $response = $reader.ReadLine()
            
            if ($response.StartsWith("220")) {{
                $sslStream = New-Object System.Net.Security.SslStream($stream)
                $sslStream.AuthenticateAsClient($SmtpServer)
                $stream = $sslStream
                $reader = New-Object System.IO.StreamReader($stream)
                $writer = New-Object System.IO.StreamWriter($stream)
                
                # STARTTLS後のEHLO
                $writer.WriteLine("EHLO localhost")
                $writer.Flush()
                do {{
                    $response = $reader.ReadLine()
                }} while ($response.StartsWith("250-"))
            }}
        }}
        
        # 認証実行
        if (-not (Send-AuthCommand $writer $reader $AuthMethod $SenderEmail $SenderPassword)) {{
            throw "認証失敗"
        }}
        
        # メール送信
        if ($IsTest) {{ Write-Output "メール送信中..." }}
        $writer.WriteLine("MAIL FROM:<$SenderEmail>")
        $writer.Flush()
        $reader.ReadLine()
        
        $writer.WriteLine("RCPT TO:<$RecipientEmail>")
        $writer.Flush()
        $reader.ReadLine()
        
        $writer.WriteLine("DATA")
        $writer.Flush()
        $reader.ReadLine()
        
        # メール内容
        $writer.WriteLine("From: $SenderEmail")
        $writer.WriteLine("To: $RecipientEmail")
        $writer.WriteLine("Subject: {subject}")
        $writer.WriteLine("Date: $(Get-Date -Format 'r')")
        $writer.WriteLine("Content-Type: text/plain; charset=utf-8")
        $writer.WriteLine("")
        
        if ($IsTest) {{
            $writer.WriteLine("RCSchedulerからのフレキシブル認証テストメールです。")
            $writer.WriteLine("")
            $writer.WriteLine("送信日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
            $writer.WriteLine("接続の保護: $ConnectionSecurity")
            $writer.WriteLine("認証方式: $AuthMethod")
            $writer.WriteLine("SMTPサーバー: $SmtpServer`:$SmtpPort")
            $writer.WriteLine("")
            if ("{source_path}" -and "{dest_path}") {{
                $writer.WriteLine("バックアップ設定:")
                $writer.WriteLine("コピー元: {source_path}")
                $writer.WriteLine("コピー先: {dest_path}")
                $writer.WriteLine("")
            }}
            $writer.WriteLine("この設定でバックアップ結果通知が送信されます。")
        }} else {{
            $writer.WriteLine("Robocopyバックアップの実行結果をお知らせします。")
            $writer.WriteLine("")
            $writer.WriteLine("実行日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
            $writer.WriteLine("結果: {'成功' if success else '失敗'}")
            $writer.WriteLine("コピー元: {source_path}")
            $writer.WriteLine("コピー先: {dest_path}")
            $writer.WriteLine("")
            $writer.WriteLine("詳細:")
            $writer.WriteLine("{escaped_message}")
        }}
        $writer.WriteLine(".")
        $writer.Flush()
        
        $response = $reader.ReadLine()
        if ($response.StartsWith("250")) {{
            if ($IsTest) {{ Write-Output "メール送信成功" }} else {{ Write-Output "バックアップ結果メール送信完了" }}
        }} else {{
            throw "メール送信失敗: $response"
        }}
        
        $writer.WriteLine("QUIT")
        $writer.Flush()
        
        Write-Output "=== 送信完了 ==="
        exit 0
        
    }} catch {{
        Write-Output "送信エラー: $($_.Exception.Message)"
        exit 1
    }} finally {{
        if ($stream) {{ $stream.Close() }}
        if ($tcpClient) {{ $tcpClient.Close() }}
    }}'''
            
            # ファイル保存
            with open(ps_path, 'w', encoding='cp932') as f:
                f.write(ps_content)
            
            return ps_path
            
        except Exception as e:
            self.log_message(f"フレキシブルメールスクリプト生成エラー: {str(e)}", "error")
            raise

    def generate_batch_mail_script(self, task_name):
        """バッチ実行用のメールスクリプトを生成"""
        try:
            ps_filename = f"{task_name}_batch_mail.ps1"
            ps_path = os.path.abspath(ps_filename)
            
            # 設定値を準備
            smtp_server = self.escape_powershell_string(self.smtp_server_var.get().strip())
            smtp_port = self.smtp_port_var.get().strip()
            sender_email = self.escape_powershell_string(self.sender_email_var.get().strip())
            sender_password = self.escape_powershell_string(self.sender_password_var.get().strip())
            recipient_email = self.escape_powershell_string(self.recipient_email_var.get().strip())
            
            # 接続・認証設定
            connection_security = self.connection_security_var.get()
            auth_method = self.auth_method_var.get()
            use_ssl = connection_security == "SSL/TLS"
            use_starttls = connection_security == "STARTTLS"
            
            # パス情報
            source_path = self.escape_powershell_string(self.source_var.get())
            dest_path = self.escape_powershell_string(self.dest_var.get())
            
            # バッチ実行用PowerShellスクリプト（引数対応版）
            ps_content = f'''# RCScheduler バッチ実行用メール送信
    param(
        [string]$BackupSuccess = "0",
        [string]$LogFilePath = ""
    )

    $ErrorActionPreference = "Stop"

    # 設定
    $SmtpServer = "{smtp_server}"
    $SmtpPort = {smtp_port}
    $SenderEmail = "{sender_email}"
    $SenderPassword = "{sender_password}"
    $RecipientEmail = "{recipient_email}"
    $UseSSL = {"$true" if use_ssl else "$false"}
    $UseSTARTTLS = {"$true" if use_starttls else "$false"}
    $AuthMethod = "{auth_method}"

    # 件名設定
    if ($BackupSuccess -eq "1") {{
        $Subject = "Robocopyバックアップ結果 - 成功"
        $Result = "成功"
    }} else {{
        $Subject = "Robocopyバックアップ結果 - 失敗"
        $Result = "失敗"
    }}

    function ConvertTo-Base64($text) {{
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
        return [System.Convert]::ToBase64String($bytes)
    }}

    function Send-AuthCommand($writer, $reader, $authMethod, $username, $password) {{
        switch ($authMethod) {{
            "CRAM-MD5" {{
                $writer.WriteLine("AUTH CRAM-MD5")
                $writer.Flush()
                $response = $reader.ReadLine()
                
                if ($response.StartsWith("334")) {{
                    $challenge = $response.Substring(4)
                    $challengeBytes = [System.Convert]::FromBase64String($challenge)
                    $challengeText = [System.Text.Encoding]::UTF8.GetString($challengeBytes)
                    
                    $hmac = New-Object System.Security.Cryptography.HMACMD5
                    $hmac.Key = [System.Text.Encoding]::UTF8.GetBytes($password)
                    $hash = $hmac.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($challengeText))
                    $hashHex = [System.BitConverter]::ToString($hash) -replace "-", ""
                    
                    $cramResponse = "$username " + $hashHex.ToLower()
                    $encodedResponse = ConvertTo-Base64 $cramResponse
                    $writer.WriteLine($encodedResponse)
                    $writer.Flush()
                }}
            }}
            "LOGIN" {{
                $writer.WriteLine("AUTH LOGIN")
                $writer.Flush()
                $response = $reader.ReadLine()
                if ($response.StartsWith("334")) {{
                    $writer.WriteLine((ConvertTo-Base64 $username))
                    $writer.Flush()
                    $response = $reader.ReadLine()
                    if ($response.StartsWith("334")) {{
                        $writer.WriteLine((ConvertTo-Base64 $password))
                        $writer.Flush()
                    }}
                }}
            }}
            "PLAIN" {{
                $authString = ConvertTo-Base64("`0$username`0$password")
                $writer.WriteLine("AUTH PLAIN $authString")
                $writer.Flush()
            }}
        }}
        
        $response = $reader.ReadLine()
        return $response.StartsWith("235")
    }}

    try {{
        # 接続
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.Connect($SmtpServer, $SmtpPort)
        $stream = $tcpClient.GetStream()
        
        # SSL/STARTTLS処理
        if ($UseSSL) {{
            $sslStream = New-Object System.Net.Security.SslStream($stream)
            $sslStream.AuthenticateAsClient($SmtpServer)
            $stream = $sslStream
            Start-Sleep -Milliseconds 500
        }}
        
        $reader = New-Object System.IO.StreamReader($stream)
        $writer = New-Object System.IO.StreamWriter($stream)
        
        # SMTP通信
        $response = $reader.ReadLine()  # 220応答
        
        $writer.WriteLine("EHLO localhost")
        $writer.Flush()
        do {{ $response = $reader.ReadLine() }} while ($response.StartsWith("250-"))
        
        if ($UseSTARTTLS -and -not $UseSSL) {{
            $writer.WriteLine("STARTTLS")
            $writer.Flush()
            $response = $reader.ReadLine()
            if ($response.StartsWith("220")) {{
                $sslStream = New-Object System.Net.Security.SslStream($stream)
                $sslStream.AuthenticateAsClient($SmtpServer)
                $stream = $sslStream
                $reader = New-Object System.IO.StreamReader($stream)
                $writer = New-Object System.IO.StreamWriter($stream)
                
                $writer.WriteLine("EHLO localhost")
                $writer.Flush()
                do {{ $response = $reader.ReadLine() }} while ($response.StartsWith("250-"))
            }}
        }}
        
        # 認証
        if (-not (Send-AuthCommand $writer $reader $AuthMethod $SenderEmail $SenderPassword)) {{
            throw "認証失敗"
        }}
        
        # メール送信
        $writer.WriteLine("MAIL FROM:<$SenderEmail>")
        $writer.Flush()
        $reader.ReadLine()
        
        $writer.WriteLine("RCPT TO:<$RecipientEmail>")
        $writer.Flush()
        $reader.ReadLine()
        
        $writer.WriteLine("DATA")
        $writer.Flush()
        $reader.ReadLine()
        
        # メール内容
        $writer.WriteLine("From: $SenderEmail")
        $writer.WriteLine("To: $RecipientEmail")
        $writer.WriteLine("Subject: $Subject")
        $writer.WriteLine("Date: $(Get-Date -Format 'r')")
        $writer.WriteLine("Content-Type: text/plain; charset=utf-8")
        $writer.WriteLine("")
        $writer.WriteLine("Robocopyバックアップの実行結果をお知らせします。")
        $writer.WriteLine("")
        $writer.WriteLine("実行日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $writer.WriteLine("結果: $Result")
        $writer.WriteLine("コピー元: {source_path}")
        $writer.WriteLine("コピー先: {dest_path}")
        $writer.WriteLine("")
        if ($LogFilePath -and (Test-Path $LogFilePath)) {{
            $writer.WriteLine("ログファイル: $LogFilePath")
            $writer.WriteLine("")
            $writer.WriteLine("=== ログ内容（全文） ===")
            $logLines = Get-Content $LogFilePath -ErrorAction SilentlyContinue
            if ($logLines) {{
                foreach ($line in $logLines) {{
                    $writer.WriteLine($line)
                }}
            }} else {{
                $writer.WriteLine("ログファイルの内容を読み取れませんでした")
            }}
        }} else {{
            $writer.WriteLine("ログファイルが見つかりません: $LogFilePath")
        }}
        $writer.WriteLine(".")
        $writer.Flush()
        
        $response = $reader.ReadLine()
        if (-not $response.StartsWith("250")) {{
            throw "送信失敗: $response"
        }}
        
        $writer.WriteLine("QUIT")
        $writer.Flush()
        
        Write-Output "メール送信完了"
        exit 0
        
    }} catch {{
        Write-Output "メール送信エラー: $($_.Exception.Message)"
        exit 1
    }} finally {{
        if ($stream) {{ $stream.Close() }}
        if ($tcpClient) {{ $tcpClient.Close() }}
    }}'''
            
            # ファイル保存
            with open(ps_path, 'w', encoding='cp932') as f:
                f.write(ps_content)
            
            self.log_message(f"バッチ用メールスクリプトを生成しました: {ps_path}")
            return ps_path
            
        except Exception as e:
            self.log_message(f"バッチ用メールスクリプト生成エラー: {str(e)}", "error")
            raise

    def generate_batch_script(self, task_name):
        """タスクスケジューラ用のバッチファイルを生成（フレキシブル対応版）"""
        try:
            batch_filename = f"{task_name}.bat"
            batch_path = os.path.abspath(batch_filename)
            
            # バッチファイルの内容を構築
            batch_content = self.build_batch_content()
            
            # バッチファイルをSJIS(CP932)で作成
            with open(batch_path, 'w', encoding='cp932') as f:
                f.write(batch_content)
            
            self.log_message(f"バッチファイルを生成しました (SJIS): {batch_path}")
            return batch_path
            
        except Exception as e:
            error_msg = f"バッチファイル生成エラー: {str(e)}"
            self.log_message(error_msg, "error")
            raise

    def build_batch_content(self):
        """バッチファイルの内容を構築（フレキシブルメール対応版）"""
        source = self.source_var.get()
        dest = self.dest_var.get()
        options = self.build_robocopy_options()  # ログオプションなしのバージョンを使用
        
        # ベースログファイルのパス（拡張子なし）
        base_log_file = self.log_file_var.get() if self.log_file_var.get() else "robocopy_schedule_log"
        if base_log_file.endswith('.txt'):
            base_log_file = base_log_file[:-4]  # .txt を除去
        base_log_file = os.path.abspath(base_log_file)
        
        batch_lines = [
            "@echo off",
            "chcp 932 >nul",  # SJIS(CP932)に設定してログファイルの文字化けを防ぐ
            "setlocal enabledelayedexpansion",
            "",
            ":: ===========================================",
            ":: Robocopy自動バックアップスクリプト",
            f":: 生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            ":: ===========================================",
            "",
            "set ERROR_OCCURRED=0",
            "",
            ":: 実行毎に一意なログファイル名を生成",
            "for /f \"tokens=1-3 delims=/ \" %%a in (\"%date%\") do (",
            "    set DATE_PART=%%c%%b%%a",
            ")",
            "for /f \"tokens=1-3 delims=:. \" %%a in (\"%time%\") do (",
            "    set TIME_PART=%%a%%b%%c",
            ")",
            ":: 時刻の先頭スペースを0に置換",
            "set TIME_PART=%TIME_PART: =0%",
            "",
            f"set LOG_FILE=\"{base_log_file}_%DATE_PART%_%TIME_PART%.txt\"",
            "set TEMP_LOG=\"" + os.path.dirname(base_log_file) + "\\robocopy_temp_%DATE_PART%_%TIME_PART%.log\"",
            "",
            ":: ログファイルにヘッダーを出力",
            "echo =========================================== > %LOG_FILE%",
            "echo Robocopy自動バックアップ実行ログ >> %LOG_FILE%",
            "echo 実行開始: %date% %time% >> %LOG_FILE%",
            "echo =========================================== >> %LOG_FILE%",
            f"echo コピー元: {source} >> %LOG_FILE%",
            f"echo コピー先: {dest} >> %LOG_FILE%",
            "echo. >> %LOG_FILE%",
            "",
        ]
        
        # ネットワーク認証（コピー元）
        if self.is_network_path(source) and self.source_username_var.get():
            server_path = "\\\\" + source.split("\\")[2]
            username = self.escape_batch_string(self.source_username_var.get())
            password = self.escape_batch_password(self.source_password_var.get())
            domain = self.escape_batch_string(self.source_domain_var.get())
            
            if domain:
                user_part = f"{domain}\\{username}"
            else:
                user_part = username
                
            batch_lines.extend([
                ":: コピー元ネットワーク認証",
                "echo [%date% %time%] コピー元ネットワーク認証中... >> %LOG_FILE%",
                f"echo コピー元サーバー: {server_path} >> %LOG_FILE%",
                f"echo ユーザー: {user_part} >> %LOG_FILE%",
                "",
                ":: 既存の接続を切断",
                f"net use \"{server_path}\" /delete /y >nul 2>&1",
                "",
                ":: 新しい接続を確立",
                f"net use \"{server_path}\" {password} /user:\"{user_part}\" /persistent:no",
                "if !errorlevel! neq 0 (",
                "    echo [%date% %time%] ERROR: コピー元ネットワーク認証に失敗 ^(エラーコード: !errorlevel!^) >> %LOG_FILE%",
                "    echo 詳細: net use コマンドが失敗しました >> %LOG_FILE%",
                "    set ERROR_OCCURRED=1",
                "    goto :CLEANUP",
                ")",
                "echo [%date% %time%] コピー元ネットワーク認証成功 >> %LOG_FILE%",
                "",
            ])
        
        # ネットワーク認証（コピー先）
        if self.is_network_path(dest) and self.dest_username_var.get():
            server_path = "\\\\" + dest.split("\\")[2]
            username = self.escape_batch_string(self.dest_username_var.get())
            password = self.escape_batch_password(self.dest_password_var.get())
            domain = self.escape_batch_string(self.dest_domain_var.get())
            
            if domain:
                user_part = f"{domain}\\{username}"
            else:
                user_part = username
                
            # コピー元と同じサーバーかチェック
            source_server = ""
            if self.is_network_path(source):
                source_server = "\\\\" + source.split("\\")[2]
            
            if server_path != source_server:  # 異なるサーバーの場合のみ認証
                batch_lines.extend([
                    ":: コピー先ネットワーク認証",
                    "echo [%date% %time%] コピー先ネットワーク認証中... >> %LOG_FILE%",
                    f"echo コピー先サーバー: {server_path} >> %LOG_FILE%",
                    f"echo ユーザー: {user_part} >> %LOG_FILE%",
                    "",
                    ":: 既存の接続を切断",
                    f"net use \"{server_path}\" /delete /y >nul 2>&1",
                    "",
                    ":: 新しい接続を確立",
                    f"net use \"{server_path}\" {password} /user:\"{user_part}\" /persistent:no",
                    "if !errorlevel! neq 0 (",
                    "    echo [%date% %time%] ERROR: コピー先ネットワーク認証に失敗 ^(エラーコード: !errorlevel!^) >> %LOG_FILE%",
                    "    echo 詳細: net use コマンドが失敗しました >> %LOG_FILE%",
                    "    set ERROR_OCCURRED=1",
                    "    goto :CLEANUP",
                    ")",
                    "echo [%date% %time%] コピー先ネットワーク認証成功 >> %LOG_FILE%",
                    "",
                ])
            else:
                batch_lines.extend([
                    "echo [%date% %time%] コピー先は同じサーバーのため認証をスキップ >> %LOG_FILE%",
                    "",
                ])
        
        # 接続確認とRobocopy実行
        batch_lines.extend([
            ":: フォルダアクセス確認",
            "echo [%date% %time%] フォルダアクセス確認中... >> %LOG_FILE%",
            "",
            ":: コピー元フォルダの確認",
            f"if not exist \"{source}\" (",
            f"    echo [%date% %time%] ERROR: コピー元フォルダにアクセスできません: {source} >> %LOG_FILE%",
            "    set ERROR_OCCURRED=1",
            "    goto :CLEANUP",
            ")",
            "echo [%date% %time%] コピー元フォルダアクセス確認OK >> %LOG_FILE%",
            "",
            ":: コピー先フォルダの確認（存在しない場合は作成を試行）",
            f"if not exist \"{dest}\" (",
            f"    echo [%date% %time%] コピー先フォルダが存在しないため作成を試行: {dest} >> %LOG_FILE%",
            f"    mkdir \"{dest}\" 2>>%LOG_FILE%",
            "    if !errorlevel! neq 0 (",
            f"        echo [%date% %time%] ERROR: コピー先フォルダを作成できません: {dest} >> %LOG_FILE%",
            "        set ERROR_OCCURRED=1",
            "        goto :CLEANUP",
            "    )",
            ")",
            "echo [%date% %time%] コピー先フォルダアクセス確認OK >> %LOG_FILE%",
            "",
            ":: Robocopy実行",
            "echo [%date% %time%] Robocopy実行開始 >> %LOG_FILE%",
            f"echo コマンド: robocopy \"{source}\" \"{dest}\" {options} >> %LOG_FILE%",
            "echo ---------------------------------------- >> %LOG_FILE%",
            "",
            ":: 既存の一時ログファイルを削除",
            "if exist %TEMP_LOG% del %TEMP_LOG%",
            "",
            ":: robocopyを実行（一時ログファイルに出力）",
            f"robocopy \"{source}\" \"{dest}\" {options} > %TEMP_LOG% 2>&1",
            "set ROBOCOPY_EXIT_CODE=!errorlevel!",
            "",
            ":: robocopyの出力をメインログファイルに統合",
            "if exist %TEMP_LOG% (",
            "    type %TEMP_LOG% >> %LOG_FILE%",
            "    del %TEMP_LOG%",
            ") else (",
            "    echo robocopyの出力ファイルが見つかりません >> %LOG_FILE%",
            ")",
            "",
            "echo ---------------------------------------- >> %LOG_FILE%",
            "",
            ":: Robocopy結果判定（0-7は正常、8以上はエラー）",
            "if !ROBOCOPY_EXIT_CODE! LSS 8 (",
            "    echo [%date% %time%] Robocopy実行成功 ^(戻り値: !ROBOCOPY_EXIT_CODE!^) >> %LOG_FILE%",
            "    set BACKUP_SUCCESS=1",
            ") else (",
            "    echo [%date% %time%] Robocopy実行失敗 ^(戻り値: !ROBOCOPY_EXIT_CODE!^) >> %LOG_FILE%",
            "    set BACKUP_SUCCESS=0",
            "    set ERROR_OCCURRED=1",
            ")",
            "",
        ])
        
        # メール送信（フレキシブル対応）
        if self.email_enabled_var.get():
            # 絶対パスでPowerShellスクリプトのパスを取得
            task_name = self.task_name_var.get()
            ps_script_name = f"{task_name}_batch_mail.ps1"
            ps_script_path = os.path.abspath(ps_script_name)
            
            batch_lines.extend([
                ":: メール送信（フレキシブル認証）",
                "echo [%date% %time%] メール送信準備中... >> %LOG_FILE%",
                f"echo PowerShellスクリプトパス: {ps_script_path} >> %LOG_FILE%",
                "",
                ":: PowerShellスクリプトの存在確認",
                f"if not exist \"{ps_script_path}\" (",
                "    echo [%date% %time%] WARNING: PowerShellスクリプトが見つかりません >> %LOG_FILE%",
                f"    echo ファイルパス: {ps_script_path} >> %LOG_FILE%",
                "    echo カレントディレクトリ: %CD% >> %LOG_FILE%",
                "    echo メール送信をスキップします >> %LOG_FILE%",
                "    goto :SKIP_MAIL",
                ")",
                "",
                "echo [%date% %time%] メール送信中... >> %LOG_FILE%",
                ":: PowerShellスクリプトにログファイル名を渡して実行",
                f"powershell -ExecutionPolicy Bypass -File \"{ps_script_path}\" !BACKUP_SUCCESS! %LOG_FILE% >> %LOG_FILE% 2>&1",
                "if !errorlevel! neq 0 (",
                "    echo [%date% %time%] WARNING: メール送信に失敗 >> %LOG_FILE%",
                ") else (",
                "    echo [%date% %time%] メール送信完了 >> %LOG_FILE%",
                ")",
                "",
                ":SKIP_MAIL",
                "",
            ])
        
        # クリーンアップ
        cleanup_lines = [
            ":CLEANUP",
            ":: クリーンアップ開始",
            "echo [%date% %time%] クリーンアップ開始 >> %LOG_FILE%",
            "",
            ":: 一時ファイルの削除",
            "if exist %TEMP_LOG% (",
            "    del %TEMP_LOG%",
            "    echo [%date% %time%] 一時ログファイルを削除 >> %LOG_FILE%",
            ")",
            "",
            ":: ネットワーク接続切断",
        ]
        
        # コピー元のネットワーク切断
        if self.is_network_path(source) and self.source_username_var.get():
            server_path = "\\\\" + source.split("\\")[2]
            cleanup_lines.extend([
                f"net use \"{server_path}\" /delete /y >nul 2>&1",
                f"echo [%date% %time%] コピー元ネットワーク接続切断: {server_path} >> %LOG_FILE%",
            ])
        
        # コピー先のネットワーク切断
        if self.is_network_path(dest) and self.dest_username_var.get():
            server_path = "\\\\" + dest.split("\\")[2]
            # 同じサーバーでない場合のみ切断
            source_server = "\\\\" + source.split("\\")[2] if self.is_network_path(source) else ""
            if server_path != source_server:
                cleanup_lines.extend([
                    f"net use \"{server_path}\" /delete /y >nul 2>&1",
                    f"echo [%date% %time%] コピー先ネットワーク接続切断: {server_path} >> %LOG_FILE%",
                ])
        
        cleanup_lines.extend([
            "",
            ":: 実行結果サマリー",
            "echo =========================================== >> %LOG_FILE%",
            "if %ERROR_OCCURRED% equ 0 (",
            "    echo [%date% %time%] バックアップ正常完了 >> %LOG_FILE%",
            "    echo バックアップが正常に完了しました。",
            "    echo ログファイル: %LOG_FILE%",
            ") else (",
            "    echo [%date% %time%] バックアップ異常終了 >> %LOG_FILE%",
            "    echo バックアップ中にエラーが発生しました。",
            "    echo ログファイルを確認してください: %LOG_FILE%",
            ")",
            "echo 実行終了: %date% %time% >> %LOG_FILE%",
            "echo =========================================== >> %LOG_FILE%",
            "",
            "exit /b %ERROR_OCCURRED%"
        ])
        
        batch_lines.extend(cleanup_lines)
        
        return "\n".join(batch_lines)

    def toggle_email_settings(self):
        """メール設定の有効/無効を切り替え"""
        if self.email_enabled_var.get():
            # メール設定を有効化
            for widget in self.email_settings_frame.winfo_children():
                if isinstance(widget, ttk.Frame):
                    # フレーム内のウィジェットも有効化
                    for child in widget.winfo_children():
                        if hasattr(child, 'configure'):
                            if isinstance(child, ttk.Combobox):
                                child.configure(state='readonly')  # プルダウンは読み取り専用で有効化
                            elif not isinstance(child, ttk.Label):
                                child.configure(state='normal')
                elif hasattr(widget, 'configure'):
                    if isinstance(widget, ttk.Combobox):
                        widget.configure(state='readonly')  # プルダウンは読み取り専用で有効化
                    elif not isinstance(widget, ttk.Label):
                        widget.configure(state='normal')
        else:
            # メール設定を無効化
            for widget in self.email_settings_frame.winfo_children():
                if isinstance(widget, ttk.Frame):
                    # フレーム内のウィジェットも無効化
                    for child in widget.winfo_children():
                        if hasattr(child, 'configure') and not isinstance(child, ttk.Label):
                            child.configure(state='disabled')
                elif hasattr(widget, 'configure') and not isinstance(widget, ttk.Label):
                    widget.configure(state='disabled')

def main():
    # GUIモードでのみ実行（--scheduledオプションは不要になった）
    root = tk.Tk()
    app = RobocopyScheduler(root)
    root.mainloop()

if __name__ == "__main__":
    main()