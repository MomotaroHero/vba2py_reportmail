import os
import win32com.client as win32
import pythoncom
import datetime
import time
from pathlib import Path

class LoungeReport:
    def __init__(self):
        # Initialize configuration variables
        self.config = {
            'jp': {
                'in_folder': '',
                'out_folder': '',
                'sheet_name': '',
                'mail_to': '',
                'mail_bcc': '',
                'mail_sub': '',
                'mail_text': ''
            },
            'en': {
                'in_folder': '',
                'out_folder': '',
                'sheet_name': '',
                'mail_to': '',
                'mail_bcc': '',
                'mail_sub': '',
                'mail_text': ''
            },
            'md_folder': '',
            'md_files': {
                'kix': '',
                'itm': '',
                'kobe': ''
            },
            'pax': {
                'in_folder': '',
                'files': {
                    'kix_int': '',
                    'kix_dom': '',
                    'itm_dom': '',
                    'kobe_dom': ''
                }
            },
            'start_time': ''
        }
        
        # COM objects
        self.excel = None
        self.outlook = None
        self.mode = ''
        
        # Load configuration
        self.load_config()

    def load_config(self):
        """Load configuration from Excel"""
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(os.path.join(os.path.dirname(__file__), "VBA01_20240324_KIX_Common_Lounge_Report_SendMail_ver2.0.xlsm"))
            ws = wb.Sheets("開始ボタン")
            
            # Load Japanese config
            self.config['jp']['in_folder'] = ws.Cells(10, 4).Value
            self.config['jp']['out_folder'] = ws.Cells(12, 4).Value
            self.config['jp']['sheet_name'] = ws.Cells(15, 4).Value
            self.config['jp']['mail_to'] = ws.Cells(16, 4).Value
            self.config['jp']['mail_bcc'] = ws.Cells(17, 4).Value
            self.config['jp']['mail_sub'] = ws.Cells(18, 4).Value
            self.config['jp']['mail_text'] = ws.Cells(19, 4).Value
            
            # Load English config
            self.config['en']['in_folder'] = ws.Cells(11, 4).Value
            self.config['en']['out_folder'] = ws.Cells(13, 4).Value
            self.config['en']['sheet_name'] = ws.Cells(21, 4).Value
            self.config['en']['mail_to'] = ws.Cells(22, 4).Value
            self.config['en']['mail_bcc'] = ws.Cells(23, 4).Value
            self.config['en']['mail_sub'] = ws.Cells(24, 4).Value
            self.config['en']['mail_text'] = ws.Cells(25, 4).Value
            
            # Load other config
            self.config['md_folder'] = ws.Cells(27, 4).Value
            self.config['md_files']['kix'] = ws.Cells(28, 4).Value
            self.config['md_files']['itm'] = ws.Cells(29, 4).Value
            self.config['md_files']['kobe'] = ws.Cells(30, 4).Value
            self.config['pax']['in_folder'] = ws.Cells(32, 4).Value
            self.config['pax']['files']['kix_int'] = ws.Cells(33, 4).Value
            self.config['pax']['files']['kix_dom'] = ws.Cells(34, 4).Value
            self.config['pax']['files']['itm_dom'] = ws.Cells(35, 4).Value
            self.config['pax']['files']['kobe_dom'] = ws.Cells(36, 4).Value
            self.config['start_time'] = ws.Cells(38, 4).Value
            
            wb.Close(False)
            excel.Quit()
            
        except Exception as e:
            print(f"設定の読み込み中にエラーが発生しました: {str(e)}")
            raise

    def control_main(self, mode="solo"):
        """メイン制御関数"""
        self.mode = mode
        
        if mode == "solo":
            self.proc_main()
        elif mode == "sch":
            msg = f"日次レポートの自動送信を開始します。\n翌日 [{self.config['start_time']}] に自動実行します。"
            print(msg)
            self.time_schedule()

    def proc_main(self):
        """メイン処理"""
        try:
            # PDF生成
            pdf_path = self.create_pdf("en", "CommonLounge_SalesReport.xlsx", "2025年4月以降")
            
            # メール送信
            self.send_mail("en", pdf_path)
            
            # 次回実行スケジュール
            self.time_reschedule()
            
        except Exception as e:
            print(f"処理中にエラーが発生しました: {str(e)}")
            raise
        finally:
            self.cleanup()

    def create_pdf(self, lang, file_name, sheet_name):
        """PDF生成関数"""
        try:
            config = self.config[lang]
            file_path = os.path.join(config['in_folder'], file_name)
            
            # Excel起動
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # ワークブック開く
            wb = excel.Workbooks.Open(file_path)
            ws = wb.Sheets(sheet_name)
            
            # PDFファイル名生成
            pdf_name = f"{datetime.datetime.now().strftime('%Y.%m.%d')}_{file_name.replace('.xlsx', '.pdf')}"
            pdf_path = os.path.join(config['out_folder'], pdf_name)
            
            # PDF出力
            ws.ExportAsFixedFormat(
                Type=win32.constants.xlTypePDF,
                Filename=pdf_path,
                Quality=win32.constants.xlQualityStandard
            )
            
            wb.Close(False)
            excel.Quit()
            
            return pdf_path
            
        except Exception as e:
            print(f"PDF生成中にエラーが発生しました: {str(e)}")
            raise

    def send_mail(self, lang, pdf_path):
        """メール送信関数"""
        try:
            config = self.config[lang]
            
            # Outlook起動
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # olMailItem
            
            # メール設定
            mail.To = config['mail_to']
            mail.BCC = config['mail_bcc']
            mail.Subject = config['mail_sub']
            mail.Body = config['mail_text'] + "\n\n"
            
            # 添付ファイル
            if os.path.exists(pdf_path):
                mail.Attachments.Add(pdf_path)
            
            # メール送信
            mail.Send()
            
        except Exception as e:
            print(f"メール送信中にエラーが発生しました: {str(e)}")
            raise
        finally:
            if 'mail' in locals():
                del mail
            if 'outlook' in locals():
                del outlook

    def time_schedule(self):
        """スケジュール設定"""
        try:
            scheduled_time = datetime.datetime.strptime(self.config['start_time'], "%H:%M:%S").time()
            now = datetime.datetime.now()
            scheduled_datetime = datetime.datetime.combine(now.date(), scheduled_time)
            
            if scheduled_datetime < now:
                scheduled_datetime += datetime.timedelta(days=1)
                
            delta = (scheduled_datetime - now).total_seconds()
            print(f"次回実行予定: {scheduled_datetime}")
            time.sleep(delta)
            
            self.proc_main()
            
        except Exception as e:
            print(f"スケジュール設定中にエラーが発生しました: {str(e)}")
            raise

    def time_reschedule(self):
        """次回実行スケジュール再設定"""
        if self.mode == "sch":
            self.time_schedule()
        else:
            self.mode = "sch"

    def cleanup(self):
        """リソース解放"""
        try:
            if hasattr(self, 'excel') and self.excel:
                self.excel.Quit()
                del self.excel
        except:
            pass
            
        try:
            if hasattr(self, 'outlook') and self.outlook:
                del self.outlook
        except:
            pass
        
        pythoncom.CoUninitialize()

    def __del__(self):
        """デストラクタ"""
        self.cleanup()
