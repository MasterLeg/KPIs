import xlwings as xw
import pandas as pd
import numpy as np
from datetime import datetime
import calendar


class Cell:
    """  Class: Cells
         Attributes:
         - row: [str] Letter that represents the cell vertical location in Excel
         - column: [str] Number that represents the cell |izontal location in Excel
         - value: [str] Value contained in the cell
    """

    # Construc|
    def __init__(self):
        app = xw.apps.active
        location = app.selection
        # Location: "$A$1"

        # Get location value
        self.value = location.value

        # Remove first character (first "$")
        location = location.address[1:]

        # Split to get the attributes row and column
        # "A$1" => split("$") => ['A', '1']
        (self.column, self.row) = location.split('$')


class Selected_data:

    def __init__(self):
        app = xw.apps.active

        # Selecting manually the info from the opened file (recommend going from B2 to Z:)
        selection = app.selection
        # print(selection)

        # Saving the values
        values = selection.value

        # Convert to a Data Frame
        df = pd.DataFrame(values)
        # print(df.head(10))

        # Assigning the column the correct name
        df.columns = df.loc[1]
        # print(df.columns)
        self.df = df


def automated_report(df):

    # TODO: Change from the Excel file to the Datsa BAse to get the data as they are not performing the regular update
    week_number = int(input('Semana a calcular:\t'))
    year = int(input('Year:\t'))

    # Showing the columns data types
    # print(df.dtypes)
    # print(df['Finished Kits '])

    # Ensuring that the Read Data ('End MFG date') is a datetime type
    df['End MFG date'] = pd.to_datetime(df['End MFG date'], format='%d/%m/%Y')

    # Removing fails in the % Scrap column
    for index, entry in enumerate(df['% Scrap']):
        if isinstance(entry, str):
            print('Encontrado[{}]: {} .=> Reemplazando por nan'.format(index, entry))
            df.replace({f'{entry}': np.nan}, inplace=True)

    # Get the dates | the week beginning and end
    monday, sunday = get_dates_from_week_number(year, week_number)

    # Select the rows that are in the desired week
    week_lots = df.loc[(df['End MFG date'] >= monday) & (df['End MFG date'] <= sunday)]

    # All cartridges accountability
    manufactured = int(week_lots['Finished cartridges'].sum())
    started = int(week_lots['Started Cartridges'].sum())

    # Calculate scrapped cartridges
    scrapped = started - manufactured

    print(f'============ REPORT SEMANA {week_number} ==============')
    # print('Cantidad de cartuchos empezados\t\t', started)
    # print('Cantidad de cartuchos fabricados\t', manufactured)
    # print('Cantidad de cartuchos scrapeados\t', scrapped)
    print('Started cartridges (commercial):\t', started)
    print('Finished cartridges (commercial):\t', manufactured)
    print('Scrapped cartridges (commercial):\t', scrapped)

    try:
        percentage_scrap = round(scrapped / started * 100, 1)
    except ZeroDivisionError:
        percentage_scrap = 0

    print('% Scrap:\t\t\t\t\t\t\t', percentage_scrap)

    # Monthly KPI
    month = monday.month
    month_name = calendar.month_name[month]
    # Get the month beginning and end date
    start_month, end_month = get_dates_from_month(month, year)

    # Select the rows that are in the desired week
    month_lots = df.loc[(df['End MFG date'] >= start_month) & (df['End MFG date'] <= end_month)]

    # Get all manufactured commercial cartridges in this month
    monthly_manufactured_commercial = int(month_lots['Finished cartridges'].sum())

    print(f'Monthly finished cartridges {month}:\t\t', monthly_manufactured_commercial)
    print('\n')

    print(f'=========== QUALITY STATUS {year} =============')

    start_year, end_year = get_dates_from_year(year)
    # Select the lots in between the year
    lots_year = df.loc[(df['End MFG date'] >= start_year) & (df['End MFG date'] <= end_year)]
    # print(lots_year.groupby('Product').size())

    # Show all types gruped by its type. Like:
    # RELEASED: 24
    # REJECTED: 22
    # print(lots_year_commercial.groupby('Product').size())

    # Select the released lots
    lots_released = lots_year.loc[(lots_year['FINAL\nSTATUS'] == 'RELEASED')]
    print('Lots Released:\t\t\t\t\t\t', lots_released.groupby('FINAL\nSTATUS').size().sum())

    # Select the rejected lots
    lots_rejected = lots_year.loc[(lots_year['FINAL\nSTATUS'] == 'REJECTED')]
    print('Lots Rejected:\t\t\t\t\t\t', lots_rejected.groupby('FINAL\nSTATUS').size().sum())

    # Select the pending lots
    lots_pending = lots_year.loc[
        (lots_year['FINAL\nSTATUS'] == 'ON HOLD') | (lots_year['FINAL\nSTATUS'] == 'IN PROCESS')]
    print('Lots Pending:\t\t\t\t\t\t', lots_pending.groupby('FINAL\nSTATUS').size().sum())

    # Select the FTR lots
    lots_FTR = lots_year.loc[(lots_year['FINAL\nSTATUS'] == 'RELEASED') &
                             (lots_year['FTR'] == 'FTR')]
    print('Lots FTR:\t\t\t\t\t\t\t', lots_FTR.groupby('FTR').size().sum())

    # Get the total manufactured lots (except 'IN PROCESS')
    lots_manufactured = lots_year.loc[(lots_year['FINAL\nSTATUS'] == 'RELEASED') |
                                      (lots_year['FINAL\nSTATUS'] == 'REJECTED') |
                                      (lots_year['FINAL\nSTATUS'] == 'ON HOLD') |
                                      (lots_year['FINAL\nSTATUS'] == 'IN PROCESS')]
    print('TOTAL:\t\t\t\t\t\t\t\t', lots_manufactured.groupby('FINAL\nSTATUS').size().sum())
    print('\n')

    print(f"=========== LOTS HILDEN {month_name} =============")
    print("Manufactured lots:\t\t\t\t\t", month_lots.groupby('Panel').size().sum())

    month_lots_released = month_lots.loc[(month_lots['FINAL\nSTATUS'] == 'RELEASED')]
    print("Released lots:\t\t\t\t\t\t", month_lots_released.groupby('FINAL\nSTATUS').size().sum())

    month_lots_FTR = month_lots.loc[(month_lots['FINAL\nSTATUS'] == 'RELEASED') &
                                    (month_lots['FTR'] == 'FTR')]
    print("FTR:\t\t\t\t\t\t\t\t", month_lots_FTR.groupby('FTR').size().sum())

    month_lots_rejected = month_lots.loc[(month_lots['FINAL\nSTATUS'] == 'REJECTED')]
    print("Rejected lots:\t\t\t\t\t\t", month_lots_rejected.groupby('FINAL\nSTATUS').size().sum())

    month_lots_pending = month_lots.loc[
        (month_lots['FINAL\nSTATUS'] == 'ON HOLD') | (month_lots['FINAL\nSTATUS'] == 'IN PROCESS')]
    print("Pending lots:\t\t\t\t\t\t", month_lots_pending.groupby('FINAL\nSTATUS').size().sum())

    print("Finished cartridges:\t\t\t\t", monthly_manufactured_commercial)
    print("Released cartridges:\t\t\t\t", int(month_lots_released['Finished cartridges'].sum()))
    print("Pending cartridges:\t\t\t\t\t", int(month_lots_pending['Finished cartridges'].sum()))
    print("Rejected cartridges:\t\t\t\t", int(month_lots_rejected['Finished cartridges'].sum()))
    print('\n')

    print('===== GLOBAL COMMERCIAL SITUATION =====')

    # Get lots released in the desired month (Filter by Release date)
    # Select the rows that are in the desired month
    released_lots_month = df.loc[(df['Release/block date'] >= start_month) & (df['Release/block date'] <= end_month)]

    # Select the released lots
    lots_released = released_lots_month.loc[(released_lots_month['FINAL\nSTATUS'] == 'RELEASED')]

    # print('Testing: ', lots_released['Finished Kits '])

    released_kits = int(lots_released['Released/Blocked kits'].sum())

    print(f"Month:\t\t\t\t\t\t\t\t{month_name}")
    print(f'KITS released HIL:\t\t\t\t\t', released_kits)
    print('\n')

    print('========== POWERPOINT ============')
    print('======= Finished batches =========')
    print(f'========== MONTH {month_name} ============')
    print('Finished cartridges:\t', monthly_manufactured_commercial)
    print("Released:\t\t\t\t", int(month_lots_released['Finished cartridges'].sum()), ' ',
          round(int(month_lots_released['Finished cartridges'].sum()) / monthly_manufactured_commercial * 100, 1), '%')
    print("Pending release:\t\t", int(month_lots_pending['Finished cartridges'].sum()), ' ',
          round(int(month_lots_pending['Finished cartridges'].sum()) / monthly_manufactured_commercial * 100, 1), '%')
    print('Rejected:\t\t\t\t', int(month_lots_rejected['Finished cartridges'].sum()), ' ',
          round(int(month_lots_rejected['Finished cartridges'].sum()) / monthly_manufactured_commercial * 100, 1), '%')
    print('---------------------------------------')

    anual_scrap = round(lots_year['% Scrap'].mean() * 100, 1)
    print('Annual in Line Scrap: ', anual_scrap)

    print(f"======== CALENDAR WEEK {week_number} ==========")
    print('Finished cartridges:\t', manufactured)

    week_lots_released = week_lots.loc[(week_lots['FINAL\nSTATUS'] == 'RELEASED')]
    week_lots_pending = week_lots.loc[
        (week_lots['FINAL\nSTATUS'] == 'ON HOLD') | (week_lots['FINAL\nSTATUS'] == 'IN PROCESS')]
    week_lots_rejected = week_lots.loc[(week_lots['FINAL\nSTATUS'] == 'REJECTED')]

    released_cartridges = int(week_lots_released['Finished cartridges'].sum())
    pending_cartridges = int(week_lots_pending['Finished cartridges'].sum())
    rejected_cartridges = int(week_lots_rejected['Finished cartridges'].sum())

    print("Released:\t\t\t\t", released_cartridges, ' ', round(released_cartridges / manufactured * 100, 1),
          '%')
    print("Pending release:\t\t", pending_cartridges, ' ', round(pending_cartridges / manufactured * 100, 1),
          '%')
    print("Rejected:\t\t\t\t", rejected_cartridges, ' ', round(rejected_cartridges / manufactured * 100, 1),
          '%')

    print("=========== Non commercial ============")
    all_month_cartridges = int(month_lots['Finished cartridges'].sum())
    all_week_cartridges = int(week_lots['Finished cartridges'].sum())

    month_no_commercial = all_month_cartridges - monthly_manufactured_commercial
    week_no_commercial = all_week_cartridges - manufactured

    print(f'{month_name}:\t\t\t\t', month_no_commercial)
    print(f'Week {week_number}:\t\t\t\t', week_no_commercial)

    print('======== Released cartridges ===========')
    print(f'{month_name} {year}\t\t\t', released_kits * 6, ' cart HIL')

    released_lots_week = df.loc[(df['Release/block date'] >= monday) & (df['Release/block date'] <= sunday)]

    # Select the released lots
    lots_released = released_lots_week.loc[(released_lots_week['FINAL\nSTATUS'] == 'RELEASED')]

    released_kits_week = int(lots_released['Released/Blocked kits'].sum())

    print(f'CALENDAR WEEK {week_number}\t\t', released_kits_week * 6, ' cart HIL')


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


def delete_information(df, delete_rows=None, delete_columns=None):
    """
    Deletes from the DataFrame any non necessary columns or rows, selected by the user
    :param df: [DataFrame] raw Data Frame
    :param delete_rows: Number of the first rows to delete. If 'NA' does not apply
    F.Ex: delete_rows = 9 => Deletes in the DataFrame the rows: [1, 2, 3, 4, 5, 6, 7, 8, 9]
    Note: in the Excel file rows start from 1, so are the same way codified in the DataFrame (not start from 0)
    :param delete_columns: Number of first columns to delete. If 'NA' does not apply
    F.Ex: delete_columns = 1 => Deletes in the Data Frame only the first column
    :return df1: [DataFrame] DataFrame with the selected information deleted
    """

    # Check if any row or column must be deleted
    if delete_columns is not None and delete_rows is None:
        # Delete the indicated columns
        df1 = df.drop(df.columns[:delete_columns], axis=1)

    elif delete_columns is None and delete_rows is not None:
        # Delete the indicated rows
        df1 = pd.DataFrame(df[delete_rows:].values)

    elif delete_columns is not None and delete_rows is not None:
        # Delete the columns and then the rows
        tdf = df.drop(df.columns[delete_columns], axis=1)
        df1 = pd.DataFrame(tdf[delete_rows:].values)

    # If all == 'NA' do not delete anything
    else:
        df1 = df

    return df1
