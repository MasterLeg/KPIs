import pandas as pd
from datetime import datetime
import calendar
from AnualScrap import generate_anual_scrap_report


def generate_automated_KPIs():
    file = "J:\\38 Production KPIs\\KPIs Emilio.xlsx"

    sheet = "Cartridge plan"

    # Open Excel sheet as Data Frame
    df = pd.read_excel(io=file, sheet_name=sheet)
    # Checking the data types
    # print(df.dtypes)

    # Ensuring that the Report Data ('U2 Date') and Release Date are both datetime types
    df['U2 Date'] = pd.to_datetime(df['U2 Date'], format='%d/%m/%Y')
    df['Release date'] = pd.to_datetime(df['Release date'], format='%Y-%m-%d')

    # Important methods used in this script
    # size() => Counts how many elements share a feature
    # t = df.groupby('Product').size()
    # sum() => Sums up all the selected

    week_number = int(input('Semana a calcular:\t'))
    year = int(input('Year:\t'))

    # Get the dates for the week beginning and end
    monday, sunday = get_dates_from_week_number(year, week_number)

    # Select the rows that are in the desired week
    week_lots = df.loc[(df['U2 Date'] >= monday) & (df['U2 Date'] <= sunday)]

    # Select only the commercial lots
    week_lots_commercial = week_lots.loc[(week_lots['Product'] == 'QIAstat Corona CEIVD') |
                                         (week_lots['Product'] == 'QIAstat Corona FDA') |
                                         (week_lots['Product'] == 'QIAstat Corona RUO') |
                                         (week_lots['Product'] == 'QIAstat Gastro CEIVD')]

    # All cartridges accountability
    manufactured = int(week_lots['Manufactured Qty'].sum())
    scrapped = int(week_lots['Total Scrap Qty'].sum())

    # Only commercial lots accountability
    manufactured_commercial = int(week_lots_commercial['Manufactured Qty'].sum())
    scrapped_commercial = int(week_lots_commercial['Total Scrap Qty'].sum())
    started_commercial = manufactured_commercial + scrapped_commercial

    print(f'============ REPORT SEMANA {week_number} ==============')
    # print('Cantidad de cartuchos empezados\t\t', started)
    # print('Cantidad de cartuchos fabricados\t', manufactured)
    # print('Cantidad de cartuchos scrapeados\t', scrapped)
    print('Started cartridges (commercial):\t', started_commercial)
    print('Finished cartridges (commercial):\t', manufactured_commercial)
    print('Scrapped cartridges (commercial):\t', scrapped_commercial)

    try:
        percentage_scrap = round(scrapped_commercial / started_commercial * 100, 1)
    except ZeroDivisionError:
        percentage_scrap = 0

    print('% Scrap:\t\t\t\t\t\t\t', percentage_scrap)

    # Monthly KPI
    month = monday.month
    month_name = calendar.month_name[month]
    # Get the month beginning and end date
    start_month, end_month = get_dates_from_month(month, year)

    # Select the rows that are in the desired week
    month_lots = df.loc[(df['U2 Date'] >= start_month) & (df['U2 Date'] <= end_month)]

    # Select only the commercial lots
    month_lots_commercial = month_lots.loc[(month_lots['Product'] == 'QIAstat Corona CEIVD') |
                                           (month_lots['Product'] == 'QIAstat Corona FDA') |
                                           (month_lots['Product'] == 'QIAstat Corona RUO') |
                                           (month_lots['Product'] == 'QIAstat Gastro CEIVD')]

    # Get all manufactured commercial cartridges in this month
    monthly_manufactured_commercial = int(month_lots_commercial['Manufactured Qty'].sum())

    # Select the non commercial lots
    all_month_cartridges = int(month_lots['Manufactured Qty'].sum())
    all_week_cartridges = int(week_lots['Manufactured Qty'].sum())

    month_no_commercial = all_month_cartridges - monthly_manufactured_commercial
    week_no_commercial = all_week_cartridges - manufactured_commercial

    print(f'Monthly finished cartridges {month_name}:\t\t', monthly_manufactured_commercial)
    print(f'R+D {month_name}:\t\t', month_no_commercial)
    print('\n')

    print(f'=========== QUALITY STATUS {year} =============')

    start_year, end_year = get_dates_from_year(year)
    # Select the lots in between the year
    lots_year = df.loc[(df['U2 Date'] >= start_year) & (df['U2 Date'] <= end_year)]
    # print(lots_year.groupby('Product').size())

    # Select only commercial lots
    lots_year_commercial = lots_year.loc[(lots_year['Product'] == 'QIAstat Corona CEIVD') |
                                         (lots_year['Product'] == 'QIAstat Corona FDA') |
                                         (lots_year['Product'] == 'QIAstat Corona RUO') |
                                         (lots_year['Product'] == 'QIAstat Gastro CEIVD') |
                                         (lots_year['Product'] == 'QIAstat Resp CEIVD')]

    # print(lots_year_commercial.groupby('Product').size())

    # Select the released lots
    lots_released = lots_year_commercial.loc[(lots_year_commercial['QC OK'] == 'Yes')]
    print('Lots Released:\t\t\t\t\t\t', lots_released.groupby('QC OK').size().sum())

    # Select the rejected lots
    lots_rejected = lots_year_commercial.loc[(lots_year_commercial['QC OK'] == 'No')]
    print('Lots Rejected:\t\t\t\t\t\t', lots_rejected.groupby('QC OK').size().sum())

    # Select the pending lots
    lots_pending = lots_year_commercial.loc[(lots_year_commercial['QC OK'] == 'q')]
    print('Lots Pending:\t\t\t\t\t\t', lots_pending.groupby('QC OK').size().sum())

    # Select the FTR lots
    lots_FTR = lots_year_commercial.loc[(lots_year_commercial['QC OK'] == 'Yes') &
                                        (lots_year_commercial['FTR'] == 'Yes')]
    print('Lots FTR:\t\t\t\t\t\t\t', lots_FTR.groupby('FTR').size().sum())

    # Get the total manufactured lots
    # print(lots_year_commercial.groupby('QC OK').size())
    print('TOTAL:\t\t\t\t\t\t\t\t', lots_year_commercial.groupby('QC OK').size().sum())
    print('\n')

    print(f"=========== LOTS BARCELONA {month} =============")
    print("Manufactured lots:\t\t\t\t\t", month_lots_commercial.groupby('Product').size().sum())

    month_lots_released = month_lots_commercial.loc[(month_lots_commercial['QC OK'] == 'Yes')]
    print("Released lots:\t\t\t\t\t\t", month_lots_released.groupby('QC OK').size().sum())

    month_lots_FTR = month_lots_commercial.loc[(month_lots_commercial['QC OK'] == 'Yes') &
                                               (month_lots_commercial['FTR'] == 'Yes')]
    print("FTR:\t\t\t\t\t\t\t\t", month_lots_FTR.groupby('FTR').size().sum())

    month_lots_rejected = month_lots_commercial.loc[(month_lots_commercial['QC OK'] == 'No')]
    print("Rejected lots:\t\t\t\t\t\t", month_lots_rejected.groupby('QC OK').size().sum())

    month_lots_pending = month_lots_commercial.loc[(month_lots_commercial['QC OK'] == 'q')]
    print("Pending lots:\t\t\t\t\t\t", month_lots_pending.groupby('QC OK').size().sum())

    print("Finished cartridges:\t\t\t\t", monthly_manufactured_commercial)
    print("Released cartridges:\t\t\t\t", int(month_lots_released['Manufactured Qty'].sum()))
    print("Pending cartridges:\t\t\t\t\t", int(month_lots_pending['Manufactured Qty'].sum()))
    print("Rejected cartridges:\t\t\t\t", int(month_lots_rejected['Manufactured Qty'].sum()))
    print('\n')

    print('===== GLOBAL COMMERCIAL SITUATION =====')

    # Get lots released in the desired month (Filter by Release date)
    # Select the rows that are in the desired month
    released_lots_month = df.loc[(df['Release date'] >= start_month) & (df['Release date'] <= end_month)]


    # Select only commercial lots
    released_lots_commercial = released_lots_month.loc[(released_lots_month['Product'] == 'QIAstat Corona CEIVD') |
                                                       (released_lots_month['Product'] == 'QIAstat Corona FDA') |
                                                       (released_lots_month['Product'] == 'QIAstat Corona RUO') |
                                                       (released_lots_month['Product'] == 'QIAstat Gastro CEIVD') |
                                                       (released_lots_month['Product'] == 'QIAstat Resp CEIVD')]

    # Select the released lots
    lots_released = released_lots_commercial.loc[(released_lots_commercial['QC OK'] == 'Yes')]

    released_kits = int(lots_released['Kits'].sum())
    print(f"Month:\t\t\t\t\t\t\t\t{month_name}")
    print(f'KITS released BCN:\t\t\t\t\t', released_kits)
    print('\n')

    print('========== POWERPOINT ============')
    print('======= Finished batches =========')
    print(f'========== MONTH {month_name} ============')
    print('Finished cartridges:\t', monthly_manufactured_commercial)
    print("Released:\t\t\t\t", int(month_lots_released['Manufactured Qty'].sum()), ' ', round(int(month_lots_released['Manufactured Qty'].sum())/monthly_manufactured_commercial*100, 1), '%')
    print("Pending release:\t\t", int(month_lots_pending['Manufactured Qty'].sum()), ' ', round(int(month_lots_pending['Manufactured Qty'].sum())/monthly_manufactured_commercial*100, 1), '%')
    print('Rejected:\t\t\t\t', int(month_lots_rejected['Manufactured Qty'].sum()), ' ', round(int(month_lots_rejected['Manufactured Qty'].sum())/monthly_manufactured_commercial*100, 1), '%')
    print('---------------------------------------')
    print('Annual in Line Scrap:')
    print(generate_anual_scrap_report())

    print(f"======== CALENDAR WEEK {week_number} ==========")
    print('Finished cartridges:\t', manufactured_commercial)

    week_lots_released = week_lots_commercial.loc[(week_lots_commercial['QC OK'] == 'Yes')]
    week_lots_pending = week_lots_commercial.loc[(week_lots_commercial['QC OK'] == 'q')]
    week_lots_rejected = week_lots_commercial.loc[(week_lots_commercial['QC OK'] == 'No')]

    released_cartridges = int(week_lots_released['Manufactured Qty'].sum())
    pending_cartridges = int(week_lots_pending['Manufactured Qty'].sum())
    rejected_cartridges = int(week_lots_rejected['Manufactured Qty'].sum())

    print("Released:\t\t\t\t", released_cartridges, ' ', round(released_cartridges/manufactured_commercial*100, 1), '%')
    print("Pending release:\t\t", pending_cartridges, ' ', round(pending_cartridges/manufactured_commercial*100, 1), '%')
    print("Rejected:\t\t\t\t", rejected_cartridges, ' ', round(rejected_cartridges/manufactured_commercial*100, 1), '%')

    print("=========== Non commercial ============")


    print(f'{month_name}:\t\t\t\t', month_no_commercial)
    print(f'Week {week_number}:\t\t\t\t', week_no_commercial)

    print('======== Released cartridges ===========')
    print(f'{month_name} 2020\t\t\t', released_kits*6, ' cart BCN')

    released_lots_week = df.loc[(df['Release date'] >= monday) & (df['Release date'] <= sunday)]

    # Select only commercial lots
    released_lots_commercial_week = released_lots_week.loc[(released_lots_week['Product'] == 'QIAstat Corona CEIVD') |
                                                       (released_lots_week['Product'] == 'QIAstat Corona FDA') |
                                                       (released_lots_week['Product'] == 'QIAstat Corona RUO') |
                                                       (released_lots_week['Product'] == 'QIAstat Gastro CEIVD') |
                                                       (released_lots_week['Product'] == 'QIAstat Resp CEIVD')]

    # Select the released lots
    lots_released = released_lots_commercial_week.loc[(released_lots_commercial_week['QC OK'] == 'Yes')]

    released_kits_week = int(lots_released['Kits'].sum())

    print(f'CALENDAR WEEK {week_number}\t\t', released_kits_week*6, ' cart BCN')



def get_dates_from_week_number(year, week_number):
    # Gets the datetime from year => 2020, week_number => 48 and day_number => 1 (monday)
    monday = datetime.fromisocalendar(year, week_number, 1)
    sunday = datetime.fromisocalendar(year, week_number, 7)
    return monday, sunday


def get_dates_from_month(month, year):
    # Gets the first and last days from a month
    start_month = datetime(year, month, 1)
    last_day_num = calendar.monthrange(year, month)[1]
    end_month = datetime(year, month, last_day_num)
    return start_month, end_month


def get_dates_from_year(year):
    # gets the first and last date from the selected year
    start_year = datetime(year, 1, 1)
    end_year = datetime(year, 12, 31)
    return start_year, end_year
