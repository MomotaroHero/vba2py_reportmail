import os
import win32com.client as win32
import pythoncom
import datetime
import time
from pathlib import Path

class SalesReport:
    def __init__(self):
        # Initialize configuration variables
        self.in_folder_name = ""
        self.in_file_name = ""
        self.out_folder_name = ""
        self.out_folder_name2 = ""
        self.out_file_name = ""
        
        # Language-specific configurations
        self.jp_config = {
            'in_folder': "",
            'out_folder': "",
            'out_folder2': "",
            'sheet_name': "",
            'mail_to': "",
            'mail_bcc': "",
            'mail_sub': "",
            'mail_text': ""
        }
        
        self.en_config = {
            'in_folder': "",
            'out_folder': "",
            'out_folder2': "",
            'sheet_name': "",
            'mail_to': "",
            'mail_bcc': "",
            'mail_sub': "",
            'mail_text': ""
        }
        
        self.na_config = {
            'in_folder': "",
            'out_folder': "",
            'out_folder2': "",
            'sheet_name': "",
            'mail_to': "",
            'mail_bcc': "",
            'mail_sub': "",
            'mail_text': ""
        }
        
        # Master data configuration
        self.md_folder_name = ""
        self.md_file_names = {
            'KIX': "",
            'ITM': "",
            'KOBE': ""
        }
        
        # PAX data configuration
        self.pax_config = {
            'in_folder': "",
            'files': {
                'KIX_Int': "",
                'KIX_Dom': "",
                'ITM_Dom': "",
                'KOBE_Dom': ""
            }
        }
        
        self.start_time = ""
        self.mode = ""
        
        # Initialize Excel with robust error handling
        self.excel = None
        self.init_excel()
        
        # Load configuration
        self.load_config()
    
    def init_excel(self):
        """Initialize Excel with COM object handling"""
        try:
            # Try to get existing instance first
            self.excel = win32.GetActiveObject("Excel.Application")
        except:
            try:
                # Create new instance
                self.excel = win32.Dispatch("Excel.Application")
            except Exception as e:
                print(f"Excel initialization failed: {str(e)}")
                self.cleanup()
                raise
        
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
    
    def cleanup(self):
        """Clean up COM objects"""
        if hasattr(self, 'excel') and self.excel:
            try:
                self.excel.Quit()
            except:
                pass
            finally:
                self.excel = None
        pythoncom.CoUninitialize()
    
    def load_config(self):
        """Load configuration from Excel file"""
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(os.path.join(os.path.dirname(__file__), "VBA09_20231128_KIX_Sales_Report_SendMail_ver2.0.xlsm"))
            ws = wb.Sheets("開始ボタン")

            # Calculate sheet
            ws.Activate()
            ws.Calculate()

            # Load Japanese config
            self.jp_config.update({
                'in_folder': ws.Cells(10, 4).Value,
                'out_folder': ws.Cells(12, 4).Value,
                'out_folder2': ws.Cells(13, 4).Value,
                'sheet_name': ws.Cells(15, 4).Value,
                'mail_to': ws.Cells(16, 4).Value,
                'mail_bcc': ws.Cells(17, 4).Value,
                'mail_sub': ws.Cells(18, 4).Value,
                'mail_text': ws.Cells(19, 4).Value
            })

            # Load English config  
            self.en_config.update({
                'in_folder': ws.Cells(10, 4).Value,
                'out_folder': ws.Cells(12, 4).Value,
                'out_folder2': ws.Cells(13, 4).Value,
                'sheet_name': ws.Cells(21, 4).Value,
                'mail_to': ws.Cells(22, 4).Value,
                'mail_bcc': ws.Cells(23, 4).Value,
                'mail_sub': ws.Cells(24, 4).Value,
                'mail_text': ws.Cells(25, 4).Value
            })

            # Load NA config
            self.na_config.update({
                'in_folder': ws.Cells(45, 4).Value,
                'out_folder': ws.Cells(47, 4).Value,
                'sheet_name': ws.Cells(50, 4).Value,
                'mail_to': ws.Cells(51, 4).Value,
                'mail_bcc': ws.Cells(52, 4).Value,
                'mail_sub': ws.Cells(53, 4).Value,
                'mail_text': ws.Cells(54, 4).Value
            })

            # Load master data config
            self.md_folder_name = ws.Cells(27, 4).Value
            self.md_file_names = {
                'KIX': ws.Cells(28, 4).Value,
                'ITM': ws.Cells(29, 4).Value,
                'KOBE': ws.Cells(30, 4).Value
            }

            # Load PAX config
            self.pax_config = {
                'in_folder': ws.Cells(32, 4).Value,
                'files': {
                    'KIX_Int': ws.Cells(33, 4).Value,
                    'KIX_Dom': ws.Cells(34, 4).Value,
                    'ITM_Dom': ws.Cells(35, 4).Value,
                    'KOBE_Dom': ws.Cells(36, 4).Value
                }
            }

            self.start_time = ws.Cells(38, 4).Value

            wb.Close(False)
            excel.Quit()

        except Exception as e:
            print(f"Error loading configuration: {str(e)}")
            raise
    
    def set_lang_config(self, lang_mode):
        """Set language configuration"""
        if lang_mode == "JP":
            config = self.jp_config
        elif lang_mode == "EN":
            config = self.en_config
        elif lang_mode == "NA":
            config = self.na_config
        else:
            raise ValueError(f"Invalid language mode: {lang_mode}")
        return config
    
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
            # Set language config
            self.set_lang_config('JP')
            
            # Create and send reports
            pdf_1 = self.create_pdf("JP", "Sales_Report_KIX_3.xlsx", "Sales")
            pdf_3 = self.create_pdf_itm("EN", "Sales_Report_ITM_3.xlsx", "Sales_Report_ITM_3.xlsx", "Sales")
            
            # Send emails
            self.send_mail_01("JP", pdf_1)
            self.send_mail_02("EN", pdf_3)
            
            # Schedule next run if needed
            self.time_reschedule()
            
        except Exception as e:
            print(f"Error in proc_main: {str(e)}")
            raise
        finally:
            self.cleanup()
    
    def create_pdf(self, lang_mode, out_file_name, sheet_name):
        """Create PDF from Excel report"""
        try:
            config = self.set_lang_config(lang_mode)
            
            file_path = self.edit_file_path(config['in_folder'], out_file_name)
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
            
            # Open workbook with retry mechanism
            wb = self.open_workbook(file_path)
            sheet = wb.Sheets(sheet_name)
            
            # Generate output filename
            date_str = datetime.datetime.now().strftime("%Y.%m.%d")
            prefix = f"{date_str}_KIX_{sheet_name}_Report.pdf"
            out_path = os.path.join(config['out_folder'], prefix)
            
            # Export to PDF
            sheet.ExportAsFixedFormat(
                Type=win32.constants.xlTypePDF,
                Filename=out_path,
                Quality=win32.constants.xlQualityStandard,
                IncludeDocProperties=True,
                IgnorePrintAreas=False
            )
            
            # Close workbook
            wb.Close(False)
            
            return out_path
            
        except Exception as e:
            print(f"Error creating PDF: {str(e)}")
            raise
    
    def create_pdf_itm(self, lang_mode, out_file_name, out_file_name2, sheet_name):
        """Special version for ITM reports with dual output"""
        try:
            config = self.set_lang_config(lang_mode)

            
            file_path = self.edit_file_path(config['in_folder'], out_file_name)
            
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
            
            wb = self.open_workbook(file_path)
            sheet = wb.Sheets(sheet_name)
            
            date_str = datetime.datetime.now().strftime("%Y.%m.%d")
            prefix = f"{date_str}_ITM_{sheet_name}_Report.pdf"

            # Primary output path
            out_path1 = os.path.join(config['out_folder'], prefix)
            
            sheet.ExportAsFixedFormat(
                Type=win32.constants.xlTypePDF,
                Filename=out_path1
            )
            
            # Secondary output path
            out_path2 = os.path.join(config['out_folder2'], prefix)
            
            sheet.ExportAsFixedFormat(
                Type=win32.constants.xlTypePDF,
                Filename=out_path2
            )
            
            wb.Close(False)
            return out_path1
            
        except Exception as e:
            print(f"Error creating ITM PDF: {str(e)}")
            raise
    
    def send_mail_01(self, lang_mode, pdf_path):
        """Send email with attachment"""
        try:
            config = self.set_lang_config(lang_mode)
            
            if not os.path.exists(pdf_path):
                raise FileNotFoundError(f"PDF not found: {pdf_path}")
            
            pythoncom.CoInitialize()
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # olMailItem
            
            mail.To = config['mail_to']
            mail.BCC = config['mail_bcc']
            mail.Subject = config['mail_sub']
            mail.Body = config['mail_text']
            
            mail.Attachments.Add(pdf_path)
            mail.Send()
            
        except Exception as e:
            print(f"Error sending mail: {str(e)}")
            raise
        finally:
            if 'mail' in locals():
                del mail
            if 'outlook' in locals():
                del outlook
            pythoncom.CoUninitialize()
    
    def open_workbook(self, file_path, max_retries=3):
        """Robust workbook opening with retries"""
        for attempt in range(max_retries):
            try:
                return self.excel.Workbooks.Open(file_path)
            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                time.sleep(1)
                self.cleanup()
                self.init_excel()
        return None
    
    def edit_file_path(self, folder_name, file_name):
        """Create proper file path"""
        path = Path(folder_name) / file_name
        if not path.exists():
            available = list(Path(folder_name).glob("*.*"))
            raise FileNotFoundError(
                f"File {file_name} not found in {folder_name}. "
                f"Available files: {[f.name for f in available]}"
            )
        return str(path)
    
    def time_schedule(self):
        """Schedule the next run"""
        now = datetime.datetime.now()
        scheduled_time = datetime.datetime.strptime(self.start_time, "%H:%M:%S").time()
        scheduled_datetime = datetime.datetime.combine(now.date(), scheduled_time)
        
        if scheduled_datetime < now:
            scheduled_datetime += datetime.timedelta(days=1)
        
        delta = scheduled_datetime - now
        seconds_until_run = delta.total_seconds()
        
        print(f"Next run scheduled for {scheduled_datetime}")
        time.sleep(seconds_until_run)
        self.proc_main()
    
    def __del__(self):
        """Destructor for cleanup"""
        self.cleanup()
