from SDH_CSVfiles_SendMail import SDHCSVfilesSendMail
from KIX_Sales_Report_SendMail import SalesReport
from KIX_Sales_Report_SendMail_CCOs import SalesReportCCOs
from KIX_ITM_Parking_Report_SendMail import ParkingReport
from KIX_common_Lounge_Report_SendMail import LoungeReport

if __name__ == "__main__":
    try:
        SDHCSVfilesSendMail = SDHCSVfilesSendMail()
        kix_sales_report = SalesReport()
        kix_sales_report_ccos = SalesReportCCOs()
        kix_parking_report = ParkingReport()
        kix_lounge_report = LoungeReport()

        # Run immediately
        SDHCSVfilesSendMail.control_main(mode="solo")  
        kix_sales_report.control_main(mode="solo")
        kix_sales_report_ccos.control_main(mode="solo")
        kix_parking_report.control_main(mode="solo") 
        kix_lounge_report.control_main(mode="solo")

        # Run on schedule
        # kix_sales_report.control_main(mode="sch")
        # kix_sales_report_ccos.control_main(mode="sch")
        # kix_parking_report.control_main(mode="sch")
        # kix_lounge_report.control_main(mode="sch")
        # SDHCSVfilesSendMail.control_main(mode="sch")

    except Exception as e:
        print(f"Error in main execution: {e}")