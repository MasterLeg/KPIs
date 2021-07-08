import pandas as pd
from datetime import datetime
import calendar

from typing import Tuple

from AnualScrap import generate_anual_scrap_report


def generate_automated_KPIs(df):
    # Checking the data types
    # print(df.dtypes)

    # DEPRECATED: Ensuring that the Report Data ('Reporting Date') and Release Date are both datetime types
    # df['Reporting Date'] = pd.to_datetime(df['Reporting Date'], format='%d/%m/%Y')
    # df['Release Date'] = pd.to_datetime(df['Release Date'], format='%Y-%m-%d')

    # Important methods used in this script
    # size() => Counts how many elements share a feature
    # t = df.groupby('Product').size()
    # sum() => Sums up all the selected

    week_number = int(input('Semana a calcular:\t'))
    year = int(input('Year:\t'))

    # Get the dates for the week beginning and end
    monday, sunday = get_dates_from_week_number(year, week_number)

    # Select the rows that are in the desired week
    week_lots = df.loc[(df['Reporting Date'] >= monday) & (df['Reporting Date'] <= sunday)]

    # Select only the commercial lots
    week_lots_commercial = week_lots.loc[(week_lots['Product'] == 'Corona CEIVD') |
                                         (week_lots['Product'] == 'Corona FDA') |
                                         (week_lots['Product'] == 'Corona RUO') |
                                         (week_lots['Product'] == 'Gastro CEIVD')]

    # All cartridges accountability
    manufactured = int(week_lots['Finished'].sum())
    scrapped = int(week_lots['Total Scrap Qty'].sum())

    # Only commercial lots accountability
    manufactured_commercial = int(week_lots_commercial['Finished'].sum())
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
    month_lots = df.loc[(df['Reporting Date'] >= start_month) & (df['Reporting Date'] <= end_month)]

    # Select only the commercial lots
    month_lots_commercial = month_lots.loc[(month_lots['Product'] == 'Corona CEIVD') |
                                           (month_lots['Product'] == 'Corona FDA') |
                                           (month_lots['Product'] == 'Corona RUO') |
                                           (month_lots['Product'] == 'Gastro CEIVD')]

    # Get all manufactured commercial cartridges in this month
    monthly_manufactured_commercial = int(month_lots_commercial['Finished'].sum())

    # Select the non commercial lots
    all_month_cartridges = int(month_lots['Finished'].sum())
    all_week_cartridges = int(week_lots['Finished'].sum())

    month_no_commercial = all_month_cartridges - monthly_manufactured_commercial
    week_no_commercial = all_week_cartridges - manufactured_commercial

    print(f'Monthly finished cartridges {month_name}:\t\t', monthly_manufactured_commercial)
    print(f'R+D {month_name}:\t\t', month_no_commercial)
    print('\n')

    print(f'=========== QUALITY STATUS {year} =============')

    start_year, end_year = get_dates_from_year(year)
    # Select the lots in between the year
    lots_year = df.loc[(df['Reporting Date'] >= start_year) & (df['Reporting Date'] <= end_year)]
    # print(lots_year.groupby('Product').size())

    # Select only commercial lots
    lots_year_commercial = lots_year.loc[(lots_year['Product'] == 'Corona CEIVD') |
                                         (lots_year['Product'] == 'Corona FDA') |
                                         (lots_year['Product'] == 'Corona RUO') |
                                         (lots_year['Product'] == 'Gastro CEIVD') |
                                         (lots_year['Product'] == 'Resp CEIVD')]

    # print(lots_year_commercial.groupby('Product').size())

    # Select the released lots
    lots_released = lots_year_commercial.loc[(lots_year_commercial['QA Status'] == 1)]
    print('Lots Released:\t\t\t\t\t\t', lots_released.groupby('QA Status').size().sum())

    # Select the rejected lots
    lots_rejected = lots_year_commercial.loc[(lots_year_commercial['QA Status'] == 0)]
    print('Lots Rejected:\t\t\t\t\t\t', lots_rejected.groupby('QA Status').size().sum())

    # Select the pending lots
    lots_pending = lots_year_commercial.loc[(lots_year_commercial['QA Status'] == 'q')]
    print('Lots Pending:\t\t\t\t\t\t', lots_pending.groupby('QA Status').size().sum())

    # Select the FTR lots
    lots_FTR = lots_year_commercial.loc[(lots_year_commercial['QA Status'] == 1) &
                                        (lots_year_commercial['FTR'] == 1)]
    print('Lots FTR:\t\t\t\t\t\t\t', lots_FTR.groupby('FTR').size().sum())

    # Get the total manufactured lots
    # print(lots_year_commercial.groupby('QA Status').size())
    print('TOTAL:\t\t\t\t\t\t\t\t', lots_year_commercial.groupby('QA Status').size().sum())
    print('\n')

    print(f"=========== LOTS BARCELONA {month} =============")
    print("Manufactured lots:\t\t\t\t\t", month_lots_commercial.groupby('Product').size().sum())

    month_lots_released = month_lots_commercial.loc[(month_lots_commercial['QA Status'] == 1)]
    print("Released lots:\t\t\t\t\t\t", month_lots_released.groupby('QA Status').size().sum())

    month_lots_FTR = month_lots_commercial.loc[(month_lots_commercial['QA Status'] == 1) &
                                               (month_lots_commercial['FTR'] == 1)]
    print("FTR:\t\t\t\t\t\t\t\t", month_lots_FTR.groupby('FTR').size().sum())

    month_lots_rejected = month_lots_commercial.loc[(month_lots_commercial['QA Status'] == 'No')]
    print("Rejected lots:\t\t\t\t\t\t", month_lots_rejected.groupby('QA Status').size().sum())

    month_lots_pending = month_lots_commercial.loc[(month_lots_commercial['QA Status'] == 'q')]
    print("Pending lots:\t\t\t\t\t\t", month_lots_pending.groupby('QA Status').size().sum())

    print("Finished cartridges:\t\t\t\t", monthly_manufactured_commercial)
    print("Released cartridges:\t\t\t\t", int(month_lots_released['Finished'].sum()))
    print("Pending cartridges:\t\t\t\t\t", int(month_lots_pending['Finished'].sum()))
    print("Rejected cartridges:\t\t\t\t", int(month_lots_rejected['Finished'].sum()))
    print('\n')

    print('===== GLOBAL COMMERCIAL SITUATION =====')

    # Get lots released in the desired month (Filter by Release Date)
    # Select the rows that are in the desired month
    released_lots_month = df.loc[(df['Release Date'] >= start_month) & (df['Release Date'] <= end_month)]

    # Select only commercial lots
    released_lots_commercial = released_lots_month.loc[(released_lots_month['Product'] == 'Corona CEIVD') |
                                                       (released_lots_month['Product'] == 'Corona FDA') |
                                                       (released_lots_month['Product'] == 'Corona RUO') |
                                                       (released_lots_month['Product'] == 'Gastro CEIVD') |
                                                       (released_lots_month['Product'] == 'Resp CEIVD')]

    # Select the released lots
    lots_released = released_lots_commercial.loc[(released_lots_commercial['QA Status'] == 1)]

    released_kits = int(lots_released['Kits'].sum())
    print(f"Month:\t\t\t\t\t\t\t\t{month_name}")
    print(f'KITS released BCN:\t\t\t\t\t', released_kits)
    print('\n')

    print('========== POWERPOINT ============')
    print('======= Finished batches =========')
    print(f'========== MONTH {month_name} ============')
    print('Finished cartridges:\t', monthly_manufactured_commercial)
    print("Released:\t\t\t\t", int(month_lots_released['Finished'].sum()), ' ',
          round(int(month_lots_released['Finished'].sum()) / monthly_manufactured_commercial * 100, 1), '%')
    print("Pending release:\t\t", int(month_lots_pending['Finished'].sum()), ' ',
          round(int(month_lots_pending['Finished'].sum()) / monthly_manufactured_commercial * 100, 1), '%')
    print('Rejected:\t\t\t\t', int(month_lots_rejected['Finished'].sum()), ' ',
          round(int(month_lots_rejected['Finished'].sum()) / monthly_manufactured_commercial * 100, 1), '%')
    print('---------------------------------------')
    print('Annual in Line Scrap:')
    print(generate_anual_scrap_report())

    print(f"======== CALENDAR WEEK {week_number} ==========")
    print('Finished cartridges:\t', manufactured_commercial)

    week_lots_released = week_lots_commercial.loc[(week_lots_commercial['QA Status'] == 1)]
    week_lots_pending = week_lots_commercial.loc[(week_lots_commercial['QA Status'] == 'q')]
    week_lots_rejected = week_lots_commercial.loc[(week_lots_commercial['QA Status'] == 'No')]

    released_cartridges = int(week_lots_released['Finished'].sum())
    pending_cartridges = int(week_lots_pending['Finished'].sum())
    rejected_cartridges = int(week_lots_rejected['Finished'].sum())

    print("Released:\t\t\t\t", released_cartridges, ' ', round(released_cartridges / manufactured_commercial * 100, 1),
          '%')
    print("Pending release:\t\t", pending_cartridges, ' ', round(pending_cartridges / manufactured_commercial * 100, 1),
          '%')
    print("Rejected:\t\t\t\t", rejected_cartridges, ' ', round(rejected_cartridges / manufactured_commercial * 100, 1),
          '%')

    print("=========== Non commercial ============")

    print(f'{month_name}:\t\t\t\t', month_no_commercial)
    print(f'Week {week_number}:\t\t\t\t', week_no_commercial)

    print('======== Released cartridges ===========')
    print(f'{month_name} 2020\t\t\t', released_kits * 6, ' cart BCN')

    released_lots_week = df.loc[(df['Release Date'] >= monday) & (df['Release Date'] <= sunday)]

    # Select only commercial lots
    released_lots_commercial_week = released_lots_week.loc[(released_lots_week['Product'] == 'Corona CEIVD') |
                                                           (released_lots_week['Product'] == 'Corona FDA') |
                                                           (released_lots_week['Product'] == 'Corona RUO') |
                                                           (released_lots_week['Product'] == 'Gastro CEIVD') |
                                                           (released_lots_week['Product'] == 'Resp CEIVD')]

    # Select the released lots
    lots_released = released_lots_commercial_week.loc[(released_lots_commercial_week['QA Status'] == 1)]

    released_kits_week = int(lots_released['Kits'].sum())

    print(f'CALENDAR WEEK {week_number}\t\t', released_kits_week * 6, ' cart BCN')


def get_dates_from_week_number(year: int, week_number: int) -> Tuple[datetime, datetime]:
    # Gets the datetime from year => 2020, week_number => 48 and day_number => 1 (monday)
    monday = datetime.fromisocalendar(year, week_number, 1)
    sunday = datetime.fromisocalendar(year, week_number, 7)
    return monday, sunday


def get_dates_from_month(month: int, year: int) -> Tuple[datetime, datetime]:
    # Gets the first and last days from a month
    start_month = datetime(year, month, 1)
    last_day_num = calendar.monthrange(year, month)[1]
    end_month = datetime(year, month, last_day_num)
    return start_month, end_month


def get_dates_from_year(year: int) -> Tuple[datetime, datetime]:
    # gets the first and last date from the selected year
    start_year = datetime(year, 1, 1)
    end_year = datetime(year, 12, 31)
    return start_year, end_year
