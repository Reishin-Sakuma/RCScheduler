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
        self.root.geometry("800x750")
        
        # 設定を保存するファイル名
        self.config_file = "robocopy_config.json"
        self.task_name = "RobocopyBackupTask"
        
        self.create_widgets()
        self.load_config()
        self.update_task_status()
    
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Robocopy設定セクション
        robocopy_frame = ttk.LabelFrame(main_frame, text="Robocopy設定", padding="10")
        robocopy_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # コピー元フォルダ
        ttk.Label(robocopy_frame, text="コピー元フォルダ:").grid(row=0, column=0, sticky=tk.W)
        self.source_var = tk.StringVar()
        ttk.Entry(robocopy_frame, textvariable=self.source_var, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(robocopy_frame, text="参照", 
                  command=self.browse_source).grid(row=0, column=2)
        
        # コピー先フォルダ
        ttk.Label(robocopy_frame, text="コピー先フォルダ:").grid(row=1, column=0, sticky=tk.W)
        self.dest_var = tk.StringVar()
        ttk.Entry(robocopy_frame, textvariable=self.dest_var, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(robocopy_frame, text="参照", 
                  command=self.browse_dest).grid(row=1, column=2)
        
        # Robocopyオプション
        ttk.Label(robocopy_frame, text="追加オプション:").grid(row=2, column=0, sticky=tk.W)
        self.options_var = tk.StringVar(value="/E /R:3 /W:10")
        ttk.Entry(robocopy_frame, textvariable=self.options_var, width=50).grid(row=2, column=1, padx=5)
        
        # スケジュール設定セクション
        schedule_frame = ttk.LabelFrame(main_frame, text="スケジュール設定", padding="10")
        schedule_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
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
        
        self.task_status_var = tk.StringVar(value="確認中...")
        ttk.Label(status_frame, textvariable=self.task_status_var).grid(row=0, column=0, sticky=tk.W)
        ttk.Button(status_frame, text="ステータス更新", 
                  command=self.update_task_status).grid(row=0, column=1, padx=10)
        
        # ログ表示エリア
        log_frame = ttk.LabelFrame(main_frame, text="実行ログ", padding="10")
        log_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.log_text = tk.Text(log_frame, height=12, width=80)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # ステータスバー
        self.status_var = tk.StringVar(value="準備完了")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
    
    def browse_source(self):
        """コピー元フォルダを選択"""
        folder = filedialog.askdirectory(title="コピー元フォルダを選択")
        if folder:
            self.source_var.set(folder)
    
    def browse_dest(self):
        """コピー先フォルダを選択"""
        folder = filedialog.askdirectory(title="コピー先フォルダを選択")
        if folder:
            self.dest_var.set(folder)
    
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
    
    def log_message(self, message):
        """ログにメッセージを追加"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def run_robocopy(self):
        """Robocopyを実行"""
        source = self.source_var.get()
        dest = self.dest_var.get()
        options = self.options_var.get()
        
        if not source or not dest:
            self.log_message("エラー: コピー元またはコピー先が指定されていません")
            return False, "コピー元またはコピー先が指定されていません"
        
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
                self.log_message("エラー出力:")
                for line in stderr_output.strip().split('\n'):
                    if line.strip():
                        self.log_message(f"  {line}")
            
            self.log_message("-" * 50)  # 区切り線
            
            # Robocopyの戻り値を確認（0-7は正常、8以上はエラー）
            if return_code < 8:
                self.log_message(f"Robocopy実行完了（戻り値: {return_code}）")
                success = True
                message = f"バックアップが正常に完了しました。\n戻り値: {return_code}\n\n" + "\n".join(output_lines)
            else:
                self.log_message(f"Robocopyでエラーが発生しました（戻り値: {return_code}）")
                success = False
                message = f"バックアップでエラーが発生しました。\n戻り値: {return_code}\n\n" + "\n".join(output_lines)
                if stderr_output:
                    message += f"\n\nエラー詳細:\n{stderr_output}"
            
            return success, message
            
        except Exception as e:
            error_msg = f"実行エラー: {str(e)}"
            self.log_message(error_msg)
            return False, error_msg
    
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
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            server.quit()
            
            self.log_message("メール送信完了")
            
        except Exception as e:
            self.log_message(f"メール送信エラー: {str(e)}")
    
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
        cmd = f'''schtasks /CREATE /TN "{self.task_name}" /TR "\\"{python_path}\\" \\"{script_path}\\" --scheduled" {schedule_type} /ST {start_time} /F'''
        
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            if result.returncode == 0:
                self.log_message(f"タスクを作成しました: {self.task_name}")
                messagebox.showinfo("成功", f"スケジュールタスク '{self.task_name}' を作成しました")
                self.update_task_status()
            else:
                error_msg = f"タスク作成エラー: {result.stderr}"
                self.log_message(error_msg)
                messagebox.showerror("エラー", error_msg)
                
        except Exception as e:
            error_msg = f"タスク作成エラー: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("エラー", error_msg)
    
    def delete_scheduled_task(self):
        """スケジュールされたタスクを削除"""
        cmd = f'schtasks /DELETE /TN "{self.task_name}" /F'
        
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp932')
            
            if result.returncode == 0:
                self.log_message(f"タスクを削除しました: {self.task_name}")
                messagebox.showinfo("成功", f"スケジュールタスク '{self.task_name}' を削除しました")
                self.update_task_status()
            else:
                error_msg = f"タスク削除エラー: {result.stderr}"
                self.log_message(error_msg)
                messagebox.showerror("エラー", error_msg)
                
        except Exception as e:
            error_msg = f"タスク削除エラー: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("エラー", error_msg)
    
    def update_task_status(self):
        """タスクのステータスを更新"""
        cmd = f'schtasks /QUERY /TN "{self.task_name}"'
        
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
            'options': self.options_var.get(),
            'frequency': self.frequency_var.get(),
            'weekday': self.weekday_var.get(),
            'hour': self.hour_var.get(),
            'minute': self.minute_var.get(),
            'email_enabled': self.email_enabled_var.get(),
            'smtp_server': self.smtp_server_var.get(),
            'smtp_port': self.smtp_port_var.get(),
            'sender_email': self.sender_email_var.get(),
            'sender_password': self.sender_password_var.get(),  # 注意: 実運用では暗号化が必要
            'recipient_email': self.recipient_email_var.get()
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            self.log_message("設定を保存しました")
            
        except Exception as e:
            error_msg = f"設定保存エラー: {str(e)}"
            self.log_message(error_msg)
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
            self.options_var.set(config.get('options', '/E /R:3 /W:10'))
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
            
            self.toggle_email_settings()
            self.log_message("設定を読み込みました")
            
        except Exception as e:
            self.log_message(f"設定読み込みエラー: {str(e)}")

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