import os
import win32com.client as win32
import pythoncom
from datetime import datetime, timedelta
import time

class SDHCSVfilesSendMail:
    def __init__(self):
        # Initialize configuration variables
        self.in_folder_name = ""
        self.in_file_name = ""
        self.out_folder_name = ""
        self.out_file_name = ""
        
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
        self.pdf_path = []
        self.file_path = ""
        
        # Initialize Excel application
        self.excel = None
        self.initialize_excel()
        
        # Load configuration
        self.load_config()
    
    def initialize_excel(self):
        """Initialize Excel application with proper error handling"""
        try:
            self.excel = win32.gencache.EnsureDispatch('Excel.Application')
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
        except Exception as e:
            print(f"Excel initialization failed: {e}")
            self.cleanup()
            raise
    
    def load_config(self):
        """Load configuration from the Excel file"""
        try:
            config_sheet = self.excel.Workbooks.Open(os.path.join(os.path.dirname(__file__), "VBA11_20230619_SDH_CSVfiles_SendMail_ver2.0.xlsm")).Sheets("開始ボタン")
            
            # Load settings from the Excel file
            self.in_folder_name = config_sheet.Cells(10, 4).Value
            self.in_file_name = config_sheet.Cells(11, 4).Value
            self.out_folder_name = config_sheet.Cells(12, 4).Value
            self.out_file_name = config_sheet.Cells(13, 4).Value

            self.jp_sheet_name = config_sheet.Cells(15, 4).Value
            self.jp_mail_to = config_sheet.Cells(16, 4).Value
            self.jp_mail_bcc = config_sheet.Cells(17, 4).Value
            self.jp_mail_sub = config_sheet.Cells(18, 4).Value
            self.jp_mail_text = config_sheet.Cells(19, 4).Value

            self.en_sheet_name = config_sheet.Cells(21, 4).Value
            self.en_mail_to = config_sheet.Cells(22, 4).Value
            self.en_mail_bcc = config_sheet.Cells(23, 4).Value
            self.en_mail_sub = config_sheet.Cells(24, 4).Value
            self.en_mail_text = config_sheet.Cells(25, 4).Value
        
            self.md_folder_name = config_sheet.Cells(27, 4).Value
            self.md_file_name_kix = config_sheet.Cells(28, 4).Value
            self.md_file_name_itm = config_sheet.Cells(29, 4).Value
            self.md_file_name_kobe = config_sheet.Cells(30, 4).Value

            self.in_folder_name_pax = config_sheet.Cells(32, 4).Value
            self.in_file_name_pax_kix_int = config_sheet.Cells(33, 4).Value
            self.in_file_name_pax_kix_dom = config_sheet.Cells(34, 4).Value
            self.in_file_name_pax_itm_dom = config_sheet.Cells(35, 4).Value
            self.in_file_name_pax_kobe_dom = config_sheet.Cells(36, 4).Value

            self.start_time = config_sheet.Cells(38, 4).Value
            
            config_sheet.Parent.Close(False)
            
        except Exception as e:
            print(f"Error loading configuration: {e}")
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
        # Set variables
        self.set_lang_config("JP")
        
        # Generate file paths with yesterday's date
        yesterday = datetime.now() - timedelta(days=1)
        prefix = yesterday.strftime("%Y%m%d")
        prefix_year = yesterday.strftime("%Y")
        prefix_display = yesterday.strftime("%Y-%m-%d")
        
        # Define file paths
        files = {
            'car_park_kix' : 'F:\\Crowdworks\\ongoing\\vba_python\\second request\\20250723_car_park_kix.csv',
            'retail_transaction' : 'F:\\Crowdworks\\ongoing\\vba_python\\second request\\20250723_retail_transactional.csv',
        }
        # files = {
        #     'retail_transaction': f"C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\retail_transactional_data\\{prefix}_retail_transactional.csv",
        #     'retail_monthly': "C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\retail_monthly\\SDH_KIX commercial data non-consolidated.xlsx",
        #     'retail_shops': "C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\retail_master\\retail_master.csv",
        #     'krs_shops': "C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\retail_master\\KRS_shop_master.xlsx",
        #     'car_park_kix': f"C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\car_park_kix\\{prefix}_car_park_kix.csv",
        #     'car_park_itm': f"C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\car_park_itm\\{prefix}_car_park_itm.csv",
        #     'car_park_ukb': f"C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\car_park_ukb\\{prefix}_car_park_ukb.csv",
        #     'products_1': "C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\retail_products\\products_master_1.csv",
        #     'products_2': "C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\retail_products\\products_master_2.csv",
        #     'retail_daily': f"C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\retail_daily\\{prefix_year}_retail_daily.csv",
        #     'retail_daily_prev': "C:\\Users\\commercial-data-cent\\OneDrive - 関西エアポートオペレーションズ\\Non-aero_data\\データ管理業務\\SDH_KAP_NonAero\\retail_daily\\2024_retail_daily.csv"
        # }
        
        # Send emails with attachments
        self.send_mail_with_attachment(files['car_park_kix'], f"KAP-KIX_CARPARKS_PAYMENTS_{prefix_display}.csv", f"KAP-KIX_CARPARKS_PAYMENTS_{prefix_display}")
        self.send_mail_with_attachment(files['retail_transaction'], f"KAP_RETAIL_TRANSACTIONAL_{prefix_display}.csv", f"KAP_RETAIL_TRANSACTIONAL_{prefix_display}")
        # self.send_mail_with_attachment(files['retail_transaction'], f"KAP_RETAIL_TRANSACTIONAL_{prefix_display}.csv", f"KAP_RETAIL_TRANSACTIONAL_{prefix_display}")
        # self.send_mail_with_attachment(files['retail_monthly'], f"KAP_RETAIL_REVENUES_{prefix_display}.xlsx", f"KAP_RETAIL_REVENUES_{prefix_display}")
        # self.send_mail_with_attachment(files['retail_shops'], f"KAP_RETAIL_SHOPS_{prefix_display}.csv", f"KAP_RETAIL_SHOPS_{prefix_display}")
        # self.send_mail_with_attachment(files['krs_shops'], f"KAP-KIX_RETAIL_SHOPS_KRS_{prefix_display}.xlsx", f"KAP-KIX_RETAIL_SHOPS_KRS_{prefix_display}")
        # self.send_mail_with_attachment(files['car_park_kix'], f"KAP-KIX_CARPARKS_PAYMENTS_{prefix_display}.csv", f"KAP-KIX_CARPARKS_PAYMENTS_{prefix_display}")
        # self.send_mail_with_attachment(files['car_park_itm'], f"KAP-ITM_CARPARKS_PAYMENTS_{prefix_display}.csv", f"KAP-ITM_CARPARKS_PAYMENTS_{prefix_display}")
        # self.send_mail_with_attachment(files['car_park_ukb'], f"KAP-UKB_CARPARKS_PAYMENTS_{prefix_display}.csv", f"KAP-UKB_CARPARKS_PAYMENTS_{prefix_display}")
        # self.send_mail_with_attachment(files['products_1'], f"KAP_RETAIL_PRODUCTS_1_{prefix_display}.csv", f"KAP_RETAIL_PRODUCTS_1_{prefix_display}")
        # self.send_mail_with_attachment(files['products_2'], f"KAP_RETAIL_PRODUCTS_2_{prefix_display}.csv", f"KAP_RETAIL_PRODUCTS_2_{prefix_display}")
        # self.send_mail_with_attachment(files['retail_daily'], f"KAP_RETAIL_WEEKLY_CONSOLIDATED_{prefix_display}.csv", f"KAP_RETAIL_WEEKLY_CONSOLIDATED_{prefix_display}")
        
        # Schedule next run if needed
        self.time_reschedule()
    
    def send_mail_with_attachment(self, file_path, display_name, subject_suffix):
        """Send email with single attachment"""
        try:
            # Verify file exists
            if not os.path.exists(file_path):
                print(f"File not found: {file_path}")
                return

            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            
            mail.To = self.mail_to
            mail.BCC = self.mail_bcc
            mail.Subject = subject_suffix
            mail.Body = self.mail_text + "\n\n"
            
            print("mail_to", self.mail_to)
            print("mail_bcc", self.mail_bcc)
            print("mail_sub", self.mail_sub)
            print("mail_text", self.mail_text)
            # Add attachment with display name
            mail.Attachments.Add(file_path, 1, 1, display_name)
            
            try: 
                mail.Send()
            except Exception as e:
                print(f"Error sending email: {e}")
            
            print(f"Email sent with attachment: {display_name}")
            
        except Exception as e:
            print(f"Error sending email: {e}")
        finally:
            # Cleanup
            if 'mail' in locals():
                del mail
            if 'outlook' in locals():
                del outlook
    
    def set_lang_config(self, lang_mode):
        """Set language-specific configuration"""
        if lang_mode == "JP":
            self.sheet_name = self.jp_sheet_name
            self.mail_to = self.jp_mail_to
            self.mail_bcc = self.jp_mail_bcc
            self.mail_sub = self.jp_mail_sub
            self.mail_text = self.jp_mail_text
        elif lang_mode == "EN":
            self.sheet_name = self.en_sheet_name
            self.mail_to = self.en_mail_to
            self.mail_bcc = self.en_mail_bcc
            self.mail_sub = self.en_mail_sub
            self.mail_text = self.en_mail_text
    
    def edit_file_path(self, folder_name, file_name):
        """Create proper file path from folder and file names"""
        path = os.path.join(folder_name, file_name)
        if not os.path.exists(path):
            print(f"Warning: File path does not exist: {path}")
        return path
    
    def time_reschedule(self):
        """Schedule next run if needed"""
        if self.mode == "sch":
            self.time_schedule()
        else:
            self.mode = "sch"
    
    def time_schedule(self):
        """Schedule the next run"""
        now = datetime.now()
        scheduled_time = datetime.strptime(self.start_time, "%H:%M:%S").time()
        scheduled_datetime = datetime.combine(now.date(), scheduled_time)
        
        # If scheduled time is in the past, schedule for next day
        if scheduled_datetime < now:
            scheduled_datetime += timedelta(days=1)
        
        # Calculate seconds until next run
        delta = scheduled_datetime - now
        seconds_until_run = delta.total_seconds()
        
        print(f"Next run scheduled for {scheduled_datetime}")
        time.sleep(seconds_until_run)
        self.proc_main()
    
    def cleanup(self):
        """Cleanup resources"""
        try:
            if self.excel is not None:
                self.excel.DisplayAlerts = True
                self.excel.ScreenUpdating = True
                self.excel.Quit()
                del self.excel
        except Exception as e:
            print(f"Error during cleanup: {e}")
    
    def __del__(self):
        """Destructor"""
        self.cleanup()
