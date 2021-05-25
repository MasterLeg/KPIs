import pandas as pd
from datetime import datetime


def generate_ramp_up_report():
    file = "J:\\38 Production KPIs\\KPIs Emilio.xlsx"

    sheet = "Cartridge plan"

    # Open Excel sheet as Data Frame
    df = pd.read_excel(io=file, sheet_name=sheet)

    # Ensuring that the Report Data ('U2 Date') and Release Date are both datetime types
    df['U2 Date'] = pd.to_datetime(df['U2 Date'], format='%d/%m/%Y')
    df['Release date'] = pd.to_datetime(df['Release date'], format='%Y-%m-%d')

    # Precreating the variables to avoid problems
    start_date = datetime(2020, 1, 1)
    end_date = datetime(2020, 1, 1)

    try:
        # Getting the desired dates from the console
        start_date = convert_string_to_datetime(input('Starting date (%d/%m/%Y):\t'))
        end_date = convert_string_to_datetime(input('Ending date (%d/%m/%Y):\t\t'))

    except:
        print('Ramp Up, Line 24, Error: Introducing wrong date format. Must be: %d/%m/%Y')

    # Select the rows that are in the desired week
    week_lots = df.loc[(df['U2 Date'] >= start_date) & (df['U2 Date'] <= end_date)]

    # Getting the metrics
    metrics(SSL=1, week_lots=week_lots)
    metrics(SSL=3, week_lots=week_lots)


def metrics(SSL, week_lots):
    """
    Generates the same report for the each line
    :param SSL: [int] The Line number. It can be: 1 or 3
    :param SSL: [DataFrame] The DataFrame with the selected rows that belong to the desired week
    :return: None (only reports on console)
    """
    # Select the lots from the desired SSL
    week_lots_SSL = week_lots.loc[(week_lots['SSL'] == SSL)]

    # Finished cartridges
    finished_cartridges = int(week_lots_SSL['Manufactured Qty'].sum())

    # Scrapped cartridges
    scrapped_cartridges = int(week_lots_SSL['Total Scrap Qty'].sum())

    # Started cartridges
    started_cartridges = finished_cartridges + scrapped_cartridges

    # Scrap percentage (%) (rounded to 2 decimals)
    scrap_percentage = round(scrapped_cartridges / started_cartridges * 100, 1)

    # Good cartridges (approved by QA)
    released_lots = week_lots_SSL.loc[(week_lots_SSL['QC OK'] == 'Yes')]
    released_cartridges = int(released_lots['Manufactured Qty'].sum())

    # Cartridges for R+D
    internal_use_cartridges = int(week_lots_SSL['Cartridges R+D'].sum())

    # Shipped cartridges to Roermond
    shipped = int(week_lots_SSL['Cartridges sent'].sum())

    # Rejected cartridges (batches)
    rejected_lots = week_lots_SSL.loc[(week_lots_SSL['QC OK'] == 'No')]
    rejected_cartridges = int(rejected_lots['Manufactured Qty'].sum())

    print(f'========== SSL{SSL} ============')
    print('Started cartridges: ', started_cartridges)
    print('Finished cartridges: ', finished_cartridges)
    print('Good cartridges SSL1: ', released_cartridges)
    print('Good cartridges for internal use: ', internal_use_cartridges)
    print('Shipped cartridges to ROE: ', shipped)
    print('Inline Scrap (%)', scrap_percentage)
    print('Rejected cartridges (batch scrap):', rejected_cartridges)

    print('\n')


def convert_string_to_datetime(string_date):
    """
    Converts a string into  a datetime object. F.Ex:
    '03-12-2020' => datetime(2020, 12, 03)

    :param string_date: [str] The date in string format
    :return converted_date: [datetime] The date in datetime format
    """
    converted_date = datetime.strptime(string_date, '%d/%m/%Y')
    return converted_date


