import pandas as pd
from KPISdata import get_dates_from_week_number


def generate_scrap_report_all():
    file = "J:\\38 Production KPIs\\KPIs Emilio.xlsx"

    sheet = "Cartridge plan"

    # Open Excel sheet as Data Frame
    df = pd.read_excel(io=file, sheet_name=sheet)
    # Checking the data types
    # print(df.dtypes)

    # Ensuring that the Report Data ('U2 Date') and Release Date are both datetime types
    df['U2 Date'] = pd.to_datetime(df['U2 Date'], format='%d/%m/%Y')
    df['Release date'] = pd.to_datetime(df['Release date'], format='%Y-%m-%d')

    week_number = int(input('Semana a calcular:\t'))
    year = 2020

    # Get the dates for the week beginning and end
    monday, sunday = get_dates_from_week_number(year, week_number)

    # Select the rows that are in the desired week
    week_lots = df.loc[(df['U2 Date'] >= monday) & (df['U2 Date'] <= sunday)]

    # Finished cartridges
    finished_cartridges = int(week_lots['Manufactured Qty'].sum())

    # Scrapped cartridges
    scrapped_cartridges = int(week_lots['Total Scrap Qty'].sum())

    # Started cartridges
    started_cartridges = finished_cartridges + scrapped_cartridges

    # Scrap percentage (%) (rounded to 2 decimals)
    scrap_percentage = round(scrapped_cartridges / started_cartridges * 100, 2)
    print(f'Week {week_number} scrap: ', scrap_percentage)
