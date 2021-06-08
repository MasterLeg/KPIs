from KPISdata import get_dates_from_month, get_dates_from_week_number, get_dates_from_year
import calendar
from AnualScrap import generate_anual_scrap_report


def generate_automated_KPIs_2(df):
    # Checking the data types
    # print(df.dtypes)

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
    week_lots_commercial = week_lots.loc[(week_lots['Product'] == 'Corona CE-IVD') |
                                         (week_lots['Product'] == 'Corona USA') |
                                         (week_lots['Product'] == 'Corona RUO') |
                                         (week_lots['Product'] == 'Gastro CE-IVD')]

    # Only commercial lots accountability
    manufactured_commercial = week_lots_commercial['Finished'].sum()
    started_commercial = week_lots_commercial['Started'].sum()
    scrapped_commercial = started_commercial - manufactured_commercial
    try:
        percentage_scrap = round(scrapped_commercial / started_commercial * 100, 1)
    except ZeroDivisionError:
        percentage_scrap = 0

    print(f'============ REPORT SEMANA {week_number} ==============')
    print('Started cartridges (commercial):\t', int(started_commercial))
    print('Finished cartridges (commercial):\t', int(manufactured_commercial))
    print('Scrapped cartridges (commercial):\t', int(scrapped_commercial))
    print('% Scrap:\t\t\t\t\t\t\t', percentage_scrap)

    # Monthly KPI
    month = monday.month
    month_name = calendar.month_name[month]
    start_month, end_month = get_dates_from_month(month, year)

    # Select the rows that are in the desired week
    month_lots = df.loc[(df['Reporting Date'] >= start_month) & (df['Reporting Date'] <= end_month)]

    # Select only the commercial lots
    month_lots_commercial = month_lots.loc[(month_lots['Product'] == 'Corona CE-IVD') |
                                           (month_lots['Product'] == 'Corona USA') |
                                           (month_lots['Product'] == 'Corona RUO') |
                                           (month_lots['Product'] == 'Gastro CE-IVD')]

    # Get all manufactured commercial cartridges in this month
    monthly_manufactured_commercial = month_lots_commercial['Finished'].sum()

    # Select the non commercial lots
    all_month_cartridges = month_lots['Finished'].sum()
    all_week_cartridges = week_lots['Finished'].sum()
    month_no_commercial = all_month_cartridges - monthly_manufactured_commercial
    week_no_commercial = all_week_cartridges - manufactured_commercial

    print(f'Monthly finished cartridges {month_name}:\t', int(monthly_manufactured_commercial))
    print(f'R+D {month_name}:\t\t\t\t\t\t', int(month_no_commercial))
    print('\n')

    # ===========================================================================================================
    print(f'=========== QUALITY STATUS {year} =============')
    start_year, end_year = get_dates_from_year(year)
    # Select the lots in between the year
    lots_year = df.loc[(df['Reporting Date'] >= start_year) & (df['Reporting Date'] <= end_year)]

    # Select only commercial lots
    lots_year_commercial = lots_year.loc[(lots_year['Product'] == 'Corona CE-IVD') |
                                         (lots_year['Product'] == 'Corona USA') |
                                         (lots_year['Product'] == 'Corona RUO') |
                                         (lots_year['Product'] == 'Gastro CE-IVD') |
                                         (lots_year['Product'] == 'Resp CEIVD')]

    total_lots_year = lots_year_commercial.groupby('Product').size().sum()
    lots_released_or_rejected = lots_year_commercial.groupby('QA Status').size().sum()
    pending_lots = total_lots_year - lots_released_or_rejected

    # Select the released lots
    lots_released = lots_year_commercial.loc[(lots_year_commercial['QA Status'] == 1)]
    print('Lots Released:\t\t\t\t\t\t', lots_released.groupby('QA Status').size().sum())

    # Select the rejected lots
    lots_rejected = lots_year_commercial.loc[(lots_year_commercial['QA Status'] == 0)]
    print('Lots Rejected:\t\t\t\t\t\t', lots_rejected.groupby('QA Status').size().sum())

    # Select the pending lots
    print('Lots Pending:\t\t\t\t\t\t', pending_lots)

    # Select the FTR lots
    lots_FTR = lots_year_commercial.loc[(lots_year_commercial['QA Status'] == 1) &
                                        (lots_year_commercial['FTR'] == 1)]

    print('Lots FTR:\t\t\t\t\t\t\t', lots_FTR.groupby('FTR').size().sum())
    print('TOTAL:\t\t\t\t\t\t\t\t', total_lots_year, '\n')

    # ===========================================================================================================
    print(f"=========== LOTS BARCELONA {month_name} =============")
    total_lots_month = month_lots_commercial.groupby('Product').size().sum()
    print("Manufactured lots:\t\t\t\t\t", total_lots_month)

    month_lots_released = month_lots_commercial.loc[(month_lots_commercial['QA Status'] == 1)]
    print("Released lots:\t\t\t\t\t\t", month_lots_released.groupby('QA Status').size().sum())

    month_lots_FTR = month_lots_commercial.loc[(month_lots_commercial['QA Status'] == 1) &
                                               (month_lots_commercial['FTR'] == 1)]
    print("FTR:\t\t\t\t\t\t\t\t", month_lots_FTR.groupby('FTR').size().sum())

    month_lots_rejected = month_lots_commercial.loc[(month_lots_commercial['QA Status'] == 0)]
    print("Rejected lots:\t\t\t\t\t\t", month_lots_rejected.groupby('QA Status').size().sum())

    month_lots_pending = month_lots_commercial.loc[(month_lots_commercial['QA Status'] != 0) &
                                                   (month_lots_commercial['QA Status'] != 1)]
    print("Pending lots:\t\t\t\t\t\t", month_lots_pending.groupby('Product').size().sum())

    print("Finished cartridges:\t\t\t\t", int(monthly_manufactured_commercial))
    print("Released cartridges:\t\t\t\t", int(month_lots_released['Finished'].sum()))
    print("Pending cartridges:\t\t\t\t\t", int(month_lots_pending['Finished'].sum()))
    print("Rejected cartridges:\t\t\t\t", int(month_lots_rejected['Finished'].sum()), '\n')

    # ===========================================================================================================
    print('===== GLOBAL COMMERCIAL SITUATION =====')

    # Get lots released in the desired month (Filter by Release Date)
    # Select the rows that are in the desired month
    released_lots_month = df.loc[(df['Release Date'] >= start_month) & (df['Release Date'] <= end_month)]

    # Select only commercial lots
    released_lots_commercial_month = released_lots_month.loc[(released_lots_month['Product'] == 'Corona CE-IVD') |
                                                       (released_lots_month['Product'] == 'Corona USA') |
                                                       (released_lots_month['Product'] == 'Corona RUO') |
                                                       (released_lots_month['Product'] == 'Gastro CE-IVD') |
                                                       (released_lots_month['Product'] == 'Resp CEIVD')]

    released_kits = released_lots_commercial_month['Qty Com'].sum()
    print(f"Month:\t\t\t\t\t\t\t\t{month_name}")
    print(f'KITS released BCN:\t\t\t\t\t', released_kits)
    print('\n')

    # ===========================================================================================================
    print('========== POWERPOINT ============')
    print('======= Finished batches =========')
    released_cartridges_percentage = round(month_lots_released['Finished'].sum()/monthly_manufactured_commercial*100, 1)
    pending_cartridges_percentage = round(month_lots_pending['Finished'].sum()/monthly_manufactured_commercial*100, 1)
    rejected_cartridges_percentage = round(month_lots_rejected['Finished'].sum()/monthly_manufactured_commercial*100, 1)

    print(f'========== MONTH {month_name} ============')
    print('Finished cartridges:\t', int(monthly_manufactured_commercial))
    print("Released:\t\t\t\t", int(month_lots_released['Finished'].sum()), ' ', released_cartridges_percentage, '%')
    print("Pending release:\t\t", int(month_lots_pending['Finished'].sum()), ' ', pending_cartridges_percentage, '%')
    print('Rejected:\t\t\t\t', int(month_lots_rejected['Finished'].sum()), ' ', rejected_cartridges_percentage, '%')
    print('---------------------------------------')
    print('Annual in Line Scrap:')
    generate_anual_scrap_report(df)
    print('\n')

    # ===========================================================================================================
    print(f"======== CALENDAR WEEK {week_number} ==========")
    print('Finished cartridges:\t', int(manufactured_commercial))

    week_lots_released = week_lots_commercial.loc[(week_lots_commercial['QA Status'] == 1)]
    week_lots_pending = week_lots_commercial.loc[(week_lots_commercial['QA Status'] != 0) &
                                                 (week_lots_commercial['QA Status'] != 1)]
    week_lots_rejected = week_lots_commercial.loc[(week_lots_commercial['QA Status'] == 0)]

    released_cartridges = week_lots_released['Finished'].sum()
    pending_cartridges = week_lots_pending['Finished'].sum()
    rejected_cartridges = week_lots_rejected['Finished'].sum()

    released_cartridges_percentage = round(released_cartridges/manufactured_commercial*100, 1)
    pending_cartridges_percentage = round(pending_cartridges/manufactured_commercial*100, 1)
    rejected_cartridges_percentage = round(rejected_cartridges/manufactured_commercial*100, 1)

    print("Released:\t\t\t\t", int(released_cartridges), ' ', released_cartridges_percentage, '%')
    print("Pending release:\t\t", int(pending_cartridges), ' ', pending_cartridges_percentage, '%')
    print("Rejected:\t\t\t\t", int(rejected_cartridges), ' ', rejected_cartridges_percentage, '%')

    print("=========== Non commercial ============")
    print(f'{month_name}:\t\t\t\t', int(month_no_commercial))
    print(f'Week {week_number}:\t\t\t\t', int(week_no_commercial))

    print('======== Released cartridges ===========')
    print(f'{month_name} 2020\t\t\t', released_kits*6, ' cart BCN')

    released_lots_week = df.loc[(df['Release Date'] >= monday) & (df['Release Date'] <= sunday)]

    # Select only commercial lots
    released_lots_commercial_week = released_lots_week.loc[(released_lots_week['Product'] == 'Corona CE-IVD') |
                                                       (released_lots_week['Product'] == 'Corona USA') |
                                                       (released_lots_week['Product'] == 'Corona RUO') |
                                                       (released_lots_week['Product'] == 'Gastro CE-IVD') |
                                                       (released_lots_week['Product'] == 'Resp CEIVD')]

    # Select the released lots
    lots_released = released_lots_commercial_week.loc[(released_lots_commercial_week['QA Status'] == 1)]

    released_kits_week = lots_released['Qty Com'].sum()

    print(f'CALENDAR WEEK {week_number}\t\t', released_kits_week*6, ' cart BCN')