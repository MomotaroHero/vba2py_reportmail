
import os
import win32com.client as win32
import pythoncom
import datetime
import time
from pathlib import Path

class SalesReportCCOs:
    def __init__(self):
        # Initialize configuration variables
        self.jp_in_folder = ""
        self.en_in_folder = ""
        self.na_in_folder = ""
        self.jp_out_folder = ""
        self.jp_out_folder2 = ""
        self.en_out_folder = ""
        self.na_out_folder = ""
        
        # Email configuration
        self.jp_mail_to = ""
        self.jp_mail_bcc = ""
        self.jp_mail_sub = ""
        self.jp_mail_text = ""
        
        self.en_mail_to = ""
        self.en_mail_bcc = ""
        self.en_mail_sub = ""
        self.en_mail_text = ""
        
        self.na_mail_to = ""
        self.na_mail_bcc = ""
        self.na_mail_sub = ""
        self.na_mail_text = ""
        
        # Master data files
        self.md_folder = ""
        self.md_kix_file = ""
        self.md_itm_file = ""
        self.md_kobe_file = ""
        
        # PAX data files
        self.pax_folder = ""
        self.pax_kix_int = ""
        self.pax_kix_dom = ""
        self.pax_itm_dom = ""
        self.pax_kobe_dom = ""
        
        self.start_time = ""
        self.mode = ""
        
        # Initialize Excel application with proper COM handling
        self.excel = None
        self.init_excel()
        
        # Load configuration
        self.load_config()
    
    def init_excel(self):
        """Initialize Excel application with error handling"""
        try:
            pythoncom.CoInitialize()
            self.excel = win32.gencache.EnsureDispatch('Excel.Application')
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
        except Exception as e:
            print(f"Excel initialization failed: {str(e)}")
            self.cleanup()
            raise
    
    def load_config(self):
        """Load configuration from Excel file"""
        try:
            config_path = os.path.join(os.path.dirname(__file__), "VBA09_20231128_KIX_Sales_Report_SendMail_ver2.0_CCOs.xlsm")
            if not os.path.exists(config_path):
                raise FileNotFoundError("Configuration file not found")
            
            # Using win32com to read Excel (preserving VBA behavior)
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(config_path)
            ws = wb.Sheets("開始ボタン")
            
            # Load Japanese config
            self.jp_in_folder = ws.Cells(10, 4).Value
            self.jp_out_folder = ws.Cells(12, 4).Value
            self.jp_out_folder2 = ws.Cells(13, 4).Value
            self.jp_mail_to = ws.Cells(16, 4).Value
            self.jp_mail_bcc = ws.Cells(17, 4).Value
            self.jp_mail_sub = ws.Cells(18, 4).Value
            self.jp_mail_text = ws.Cells(19, 4).Value
            
            # Load English config
            self.en_in_folder = ws.Cells(11, 4).Value
            self.en_out_folder = ws.Cells(13, 4).Value
            self.en_mail_to = ws.Cells(22, 4).Value
            self.en_mail_bcc = ws.Cells(23, 4).Value
            self.en_mail_sub = ws.Cells(24, 4).Value
            self.en_mail_text = ws.Cells(25, 4).Value
            
            # Load master data config
            self.md_folder = ws.Cells(27, 4).Value
            self.md_kix_file = ws.Cells(28, 4).Value
            self.md_itm_file = ws.Cells(29, 4).Value
            self.md_kobe_file = ws.Cells(30, 4).Value
            
            # Load schedule time
            self.start_time = ws.Cells(38, 4).Value
            
            wb.Close(False)
            excel.Quit()
            
        except Exception as e:
            print(f"Config loading failed: {str(e)}")
            raise
    
    def control_main(self, mode="solo"):
        """Main control function"""
        self.mode = mode
        
        if mode == "solo":
            self.proc_main()
        elif mode == "sch":
            msg = f"日次レポートの自動送信を開始します。\n翌日 [{self.start_time}] に、自動実行します。"
            print(msg)
            self.time_schedule()
    
    def proc_main(self):
        """Main processing function"""
        try:
            # Create and send reports
            pdf_kix = self.create_pdf("JP", "Sales_Report_KIX_3.xlsx", "Sales")
            pdf_itm = self.create_pdf_itm("EN", "Sales_Report_ITM_3.xlsx", "Sales_Report_ITM_3.xlsx", "Sales")
            
            # Send emails
            self.send_mail("JP", pdf_kix)
            self.send_mail("EN", pdf_itm)
            
            # Schedule next run if needed
            self.time_reschedule()
            
        except Exception as e:
            print(f"Processing failed: {str(e)}")
            raise
    
    def create_pdf(self, lang_mode, file_name, sheet_name):
        """Create PDF from Excel report"""
        try:
            self.set_lang_config(lang_mode)
            file_path = self.edit_file_path(self.jp_in_folder, file_name)
            
            # Open workbook
            wb = self.excel.Workbooks.Open(file_path)
            sheet = wb.Sheets(sheet_name)
            
            # Generate output filename
            date_str = datetime.datetime.now().strftime("%Y.%m.%d")
            prefix = f"{date_str}_KIX_{sheet_name}_Report.pdf"
            out_path = os.path.join(self.jp_out_folder, prefix)
            
            # Export to PDF
            sheet.ExportAsFixedFormat(
                Type=win32.constants.xlTypePDF,
                Filename=out_path,
                Quality=win32.constants.xlQualityStandard
            )
            
            wb.Close(False)
            return out_path
            
        except Exception as e:
            print(f"PDF creation failed: {str(e)}")
            raise
    
    def create_pdf_itm(self, lang_mode, file_name, file_name2, sheet_name):
        """Create PDF for ITM reports"""
        try:
            self.set_lang_config(lang_mode)
            file_path = self.edit_file_path(self.jp_in_folder, file_name)
            
            # Open workbook
            wb = self.excel.Workbooks.Open(file_path)
            sheet = wb.Sheets(sheet_name)
            
            # Generate output filename
            date_str = datetime.datetime.now().strftime("%Y.%m.%d")
            prefix = f"{date_str}_ITM_{sheet_name}_Report.pdf"
            out_path = os.path.join(self.jp_out_folder, prefix)
            out_path2 = os.path.join(self.jp_out_folder2, prefix)
            
            # Export to PDF (two locations)
            sheet.ExportAsFixedFormat(
                Type=win32.constants.xlTypePDF,
                Filename=out_path
            )
            sheet.ExportAsFixedFormat(
                Type=win32.constants.xlTypePDF,
                Filename=out_path2
            )
            
            wb.Close(False)
            return out_path
            
        except Exception as e:
            print(f"ITM PDF creation failed: {str(e)}")
            raise
    
    def send_mail(self, lang_mode, pdf_path):
        """Send email with attachment"""
        try:
            self.set_lang_config(lang_mode)
            
            # Initialize Outlook
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            
            # Set email properties
            mail.To = self.mail_to
            mail.BCC = self.mail_bcc
            mail.Subject = self.mail_sub
            mail.Body = self.mail_text + "\n\n"
            
            # Add attachment
            if os.path.exists(pdf_path):
                mail.Attachments.Add(pdf_path)
            
            # Send email
            mail.Send()
            
        except Exception as e:
            print(f"Email sending failed: {str(e)}")
            raise
    
    def set_lang_config(self, lang_mode):
        """Set language-specific configuration"""
        if lang_mode == "JP":
            self.mail_to = self.jp_mail_to
            self.mail_bcc = self.jp_mail_bcc
            self.mail_sub = self.jp_mail_sub
            self.mail_text = self.jp_mail_text
        elif lang_mode == "EN":
            self.mail_to = self.en_mail_to
            self.mail_bcc = self.en_mail_bcc
            self.mail_sub = self.en_mail_sub
            self.mail_text = self.en_mail_text
    
    def edit_file_path(self, folder, file):
        """Create proper file path"""
        return os.path.join(folder, file)
    
    def time_schedule(self):
        """Schedule the next run"""
        try:
            run_time = datetime.datetime.strptime(self.start_time, "%H:%M:%S").time()
            now = datetime.datetime.now()
            scheduled = datetime.datetime.combine(now.date(), run_time)
            
            if scheduled < now:
                scheduled += datetime.timedelta(days=1)
                
            wait_seconds = (scheduled - now).total_seconds()
            print(f"Next run scheduled for {scheduled}")
            time.sleep(wait_seconds)
            self.proc_main()
            
        except Exception as e:
            print(f"Scheduling failed: {str(e)}")
            raise
    
    def cleanup(self):
        """Proper cleanup of COM objects"""
        try:
            if hasattr(self, 'excel') and self.excel:
                self.excel.DisplayAlerts = True
                self.excel.ScreenUpdating = True
                self.excel.Quit()
                del self.excel
        except:
            pass
        pythoncom.CoUninitialize()
    
    def __del__(self):
        """Destructor for cleanup"""
        self.cleanup()
