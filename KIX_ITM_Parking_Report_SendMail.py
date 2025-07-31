import os
import win32com.client as win32
import pythoncom
import datetime
from pathlib import Path
import time

class ParkingReport:
    def __init__(self):
        # Initialize configuration variables
        self.in_folder_name = ""
        self.out_folder_name = ""
        self.jp_in_folder_name = ""
        self.en_in_folder_name = ""
        self.jp_out_folder_name = ""
        self.en_out_folder_name = ""
        self.sheet_name = ""
        self.mail_to = ""
        self.mail_bcc = ""
        self.mail_sub = ""
        self.mail_text = ""
        self.jp_sheet_name = ""
        self.jp_mail_to = ""
        self.jp_mail_bcc = ""
        self.jp_mail_sub = ""
        self.jp_mail_text = ""
        self.en_sheet_name = ""
        self.en_mail_to = ""
        self.en_mail_bcc = ""
        self.en_mail_sub = ""
        self.en_mail_text = ""
        self.md_folder_name = ""
        self.md_file_name_kix = ""
        self.md_file_name_itm = ""
        self.md_file_name_kobe = ""
        self.in_folder_name_pax = ""
        self.in_file_name_pax_kix_int = ""
        self.in_file_name_pax_kix_dom = ""
        self.in_file_name_pax_itm_dom = ""
        self.in_file_name_pax_kobe_dom = ""
        self.start_time = ""
        self.mode = ""
        self.outfile_path = ""
        
        # Initialize COM objects
        self.excel = None
        self.outlook = None
        
        # Load configuration
        self.load_config()

    def load_config(self):
        """Load configuration from Excel"""
        try:
            config_path = os.path.join(os.path.dirname(__file__), "VBA00_20230626_KIX_ITM_Parking_Report_SendMail_ver2.0.xlsm")
            if not os.path.exists(config_path):
                raise FileNotFoundError("Configuration file not found")
            
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(config_path)
            ws = wb.Sheets("開始ボタン")
            
            # Load configuration values
            self.jp_in_folder_name = ws.Cells(10, 4).Value
            self.en_in_folder_name = ws.Cells(11, 4).Value
            self.jp_out_folder_name = ws.Cells(12, 4).Value
            self.en_out_folder_name = ws.Cells(13, 4).Value
            
            self.jp_sheet_name = ws.Cells(15, 4).Value
            self.jp_mail_to = ws.Cells(16, 4).Value
            self.jp_mail_bcc = ws.Cells(17, 4).Value
            self.jp_mail_sub = ws.Cells(18, 4).Value
            self.jp_mail_text = ws.Cells(19, 4).Value
            
            self.en_sheet_name = ws.Cells(21, 4).Value
            self.en_mail_to = ws.Cells(22, 4).Value
            self.en_mail_bcc = ws.Cells(23, 4).Value
            self.en_mail_sub = ws.Cells(24, 4).Value
            self.en_mail_text = ws.Cells(25, 4).Value
            
            self.md_folder_name = ws.Cells(27, 4).Value
            self.md_file_name_kix = ws.Cells(28, 4).Value
            self.md_file_name_itm = ws.Cells(29, 4).Value
            self.md_file_name_kobe = ws.Cells(30, 4).Value
            
            self.in_folder_name_pax = ws.Cells(32, 4).Value
            self.in_file_name_pax_kix_int = ws.Cells(33, 4).Value
            self.in_file_name_pax_kix_dom = ws.Cells(34, 4).Value
            self.in_file_name_pax_itm_dom = ws.Cells(35, 4).Value
            self.in_file_name_pax_kobe_dom = ws.Cells(36, 4).Value
            
            self.start_time = ws.Cells(38, 4).Value                
            
            wb.Close(False)
            excel.Quit()
            
        except Exception as e:
            print(f"Error loading configuration: {str(e)}")
            raise

    def control_main(self, mode="solo"):
        """Main control function"""
        self.mode = mode
        
        if mode == "solo":
            self.proc_main()
        elif mode == "sch":
            msg = f"日次レポートの自動送信を開始します。\n翌日 [{self.start_time}] に自動実行します。"
            print(msg)
            self.time_schedule()

    def proc_main(self):
        """Main processing function"""
        try:
            # Create PDF reports
            pdf_1 = self.create_pdf("JP", "【vs2024】駐車場・アクセス施設売上速報2025.xlsx", "提出用")
            pdf_2 = self.create_pdf("EN", "【KIX】【vs2024】駐車場売上速報2025.xlsx", "提出用")
            
            # Send emails
            self.send_mail_01("JP", pdf_1)
            self.send_mail_02("EN", pdf_2)
            
            # Schedule next run if needed
            self.time_reschedule()
            
        except Exception as e:
            print(f"Error in proc_main: {str(e)}")
            raise
        finally:
            self.cleanup()

    def create_pdf(self, lang_mode, out_file_name, sh_name):
        """Create PDF from Excel sheet"""
        try:
            self.set_lang_config(lang_mode)
            file_path = self.edit_file_path(self.in_folder_name, out_file_name)
            
            # Open Excel file
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(file_path)
            ws = wb.Sheets(sh_name)
            
            # Generate PDF path
            prefix = datetime.datetime.now().strftime("%Y.%m.%d") + "_" + out_file_name.replace(".xlsx", ".pdf")
            out_path = os.path.join(self.out_folder_name, prefix)
            
            # Export to PDF
            ws.ExportAsFixedFormat(
                Type=win32.constants.xlTypePDF,
                Filename=out_path,
                Quality=win32.constants.xlQualityStandard
            )
            
            wb.Close(False)
            excel.Quit()
            
            return out_path
            
        except Exception as e:
            print(f"Error creating PDF: {str(e)}")
            raise

    def send_mail_01(self, lang_mode, pdf_path):
        """Send email with attachment (Japanese version)"""
        self._send_mail(lang_mode, [pdf_path])

    def send_mail_02(self, lang_mode, pdf_path):
        """Send email with attachment (English version)"""
        self._send_mail(lang_mode, [pdf_path])

    def _send_mail(self, lang_mode, pdf_paths):
        """Internal email sending function"""
        try:
            self.set_lang_config(lang_mode)
            
            # Initialize Outlook
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # olMailItem
            
            # Set email properties
            mail.To = self.mail_to
            mail.BCC = self.mail_bcc
            mail.Subject = self.mail_sub
            mail.Body = self.mail_text + "\n\n"
            
            print("pdf_paths", pdf_paths)
            # Add attachments
            for pdf in pdf_paths:
                if os.path.exists(pdf):
                    mail.Attachments.Add(pdf) 
                else : 
                    raise Exception(f"File not found: {pdf}")           

            print("mail From", mail.From)
            print("mail.To", mail.To)
            print("mail.BCC", mail.BCC)
            print("mail.Subject", mail.Subject)
            print("mail.Body", mail.Body)
            print("mail:", mail)
            # Send email
            try:
                mail.Send()
                print("sent mail successfully")
            except Exception as e:
                print(f"Error sending email: {str(e)}")
                raise
            
        except Exception as e:
            print(f"Error sending email: {str(e)}")
            raise
        finally:
            if 'mail' in locals():
                del mail
            if 'outlook' in locals():
                del outlook

    def set_lang_config(self, lang_mode):
        """Set language-specific configuration"""
        if lang_mode == "JP":
            self.in_folder_name = self.jp_in_folder_name
            self.out_folder_name = self.jp_out_folder_name
            self.sheet_name = self.jp_sheet_name
            self.mail_to = self.jp_mail_to
            self.mail_bcc = self.jp_mail_bcc
            self.mail_sub = self.jp_mail_sub
            self.mail_text = self.jp_mail_text
        elif lang_mode == "EN":
            self.in_folder_name = self.en_in_folder_name
            self.out_folder_name = self.en_out_folder_name
            self.sheet_name = self.en_sheet_name
            self.mail_to = self.en_mail_to
            self.mail_bcc = self.en_mail_bcc
            self.mail_sub = self.en_mail_sub
            self.mail_text = self.en_mail_text

    def edit_file_path(self, folder_name, file_name):
        """Create proper file path"""
        return os.path.join(folder_name, file_name)

    def time_schedule(self):
        """Schedule the next run"""
        try:
            scheduled_time = datetime.datetime.strptime(self.start_time, "%H:%M:%S").time()
            now = datetime.datetime.now()
            scheduled_datetime = datetime.datetime.combine(now.date(), scheduled_time)
            
            if scheduled_datetime < now:
                scheduled_datetime += datetime.timedelta(days=1)
                
            delta = (scheduled_datetime - now).total_seconds()
            print(f"Next run scheduled for {scheduled_datetime}")
            time.sleep(delta)
            
            self.proc_main()
            
        except Exception as e:
            print(f"Error in scheduling: {str(e)}")
            raise

    def time_reschedule(self):
        """Reschedule if in scheduled mode"""
        if self.mode == "sch":
            self.time_schedule()
        else:
            self.mode = "sch"

    def cleanup(self):
        """Clean up COM objects"""
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
        """Destructor for cleanup"""
        self.cleanup()
