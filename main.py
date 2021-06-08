from KPISdata import generate_automated_KPIs
from RampUp import generate_ramp_up_report
from Report_scrap_Maria import generate_scrap_report_all
from AnualScrap import generate_anual_scrap_report
from HIldenKPI import automated_report
import pandas as pd
from KPIS3 import generate_automated_KPIs_2


if __name__ == '__main__':

    while True:

        selection = input('Select activity to perform:\n'
                          '(1) KPIs Report\n'
                          '(2) Ramp Up Report\n'
                          '(3) Total Scrap\n'
                          '(4) Year in line Scrap\n'
                          '(5) Hilden KPIS data\n'
                          '(6) New matched KPIs\n'
                          'Selection:\t')

        if selection == '1':
            file = "M:\\48 Documentation\\3.KPI\\KPIS_Reports.xlsx"

            sheet = "Sheet1"

            # Open Excel sheet as Data Frame
            df = pd.read_excel(io=file, sheet_name=sheet, engine='opeenpyxl', header=1)
            # KPIs Report (Excel and Power)
            while True:
                generate_automated_KPIs(df)
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
            print('Opening the file "48 Documentation\\3.KPI\\Hilden". Ensure the updated Excel has been downloaded')
            # Hilden KPIS data scrapper (requires to download the file)
            df = pd.read_excel('J:\\48 Documentation\\3.KPI\\Hilden\\HIL_summary 2020.xlsx',
                               sheet_name='Cartridges summary',
                               header=5)
            print('OK => Data loaded')

            while True:
                automated_report(df)
                print('\n')
        elif selection == '6':

            # Hilden KPIS data scrapper (requires to download the file)
            path1 = 'J:\\48 Documentation\\3.KPI\\KPIS_Reports.xlsx'
            print('Opening the file\t' + path1)
            df1 = pd.read_excel(path1,
                                sheet_name='Sheet1',
                                engine='openpyxl')
            print('OK => Data loaded')
            df1.set_index(df1.columns[1], inplace=True)
            df1.drop(df1.columns[0], axis='columns', inplace=True)

            path2 = r'J:\48 Documentation\3.KPI\Cartridge Release Calendar BCN - Copia.xlsx'
            print('Opening the file\t' + path2)
            df2 = pd.read_excel(path2,
                                sheet_name='Calendar',
                                engine='openpyxl')
            print('OK => Data loaded')
            df2.set_index(df2.columns[3], inplace=True)

            df2.drop(df2.columns[:4], axis='columns', inplace=True)
            df2.drop(df2.columns[8:], axis='columns', inplace=True)

            df = pd.concat([df2, df1.reindex(df2.index)], axis=1)
            print('Data Frame matched\n')
            # df.to_excel(r'J:\48 Documentation\3.KPI\Matched_data.xlsx')

            print("Removing '-' from the 'Qty Com' column")
            df.replace({'Qty Com': '-'}, value=0, inplace=True)
            print("OK => Removed\n")

            while True:
                generate_automated_KPIs_2(df)
                print('\n')
