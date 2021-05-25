from KPISdata import generate_automated_KPIs
from RampUp import generate_ramp_up_report
from Report_scrap_Maria import generate_scrap_report_all
from AnualScrap import generate_anual_scrap_report
from HIldenKPI import automated_report
import pandas as pd

if __name__ == '__main__':

    selection = input('Select activity to perform:\n(1) KPIs Report\n(2) Ramp Up Report\n(3) Total Scrap'
                      '\n(4) Year in line Scrap\n(5) Hilden KPIS data\nSelection:\t')

    if selection == '1':
        # KPIs Report (Excel and Power)
        generate_automated_KPIs()
    elif selection == '2':
        # Ramp up report
        generate_ramp_up_report()
    elif selection == '3':
        # QA general scrap record
        generate_scrap_report_all()
    elif selection == '4':
        # Year in line scrap
        generate_anual_scrap_report()
    elif selection == '5':
        # Hilden KPIS data scrapper (requires to download the file)
        df = pd.read_excel('J:\\48 Documentation\\3.KPI\\Hilden\\HIL_summary 2020.xlsx',
                           sheet_name='Cartridges summary',
                           header=3)
        print('Data loaded from "48 Documentation\\3.KPI\\Hilden". Ensure the updated Excel has been downloaded')
        automated_report(df)
