import pandas as pd
from datetime import datetime, timedelta


def generate_anual_scrap_report():
    # TODO: Cambiar el nnual scrap: desde el 1 de enero hacer Monthly scrap de los lotes
    file = "J:\\38 Production KPIs\\KPIs Emilio.xlsx"

    sheet = "Cartridge plan"

    # Open Excel sheet as Data Frame
    df = pd.read_excel(io=file, sheet_name=sheet)
    # Checking the data types
    # print(df.dtypes)

    # Ensuring that the Report Data ('U2 Date') and Release Date are both datetime types
    df['U2 Date'] = pd.to_datetime(df['U2 Date'], format='%d/%m/%Y')
    df['Release date'] = pd.to_datetime(df['Release date'], format='%Y-%m-%d')

    # Select if mobile window 'y' (from current date to one year before) or in this year
    # mobile_window = input('¿Ventana móvil? y/n:\t') == 'y'  # boolean

    # Dates to select
    end_date = datetime.today()
    start_date = end_date - timedelta(days=365)

    # Not mobile window = select natural year starting and ending date
    # if not mobile_window:
    #   year = int(input('Año a calcular:\t'))
    #   start_date = datetime(year, 1, 1)
    #  end_date = datetime(year, 12, 31)

    # Select the rows that are in the desired week
    week_lots = df.loc[(df['U2 Date'] >= start_date) & (df['U2 Date'] <= end_date)]

    # Finished cartridges
    finished_cartridges = int(week_lots['Manufactured Qty'].sum())

    # Scrapped cartridges
    scrapped_cartridges = int(week_lots['Total Scrap Qty'].sum())

    # Started cartridges
    started_cartridges = finished_cartridges + scrapped_cartridges

    # Scrap percentage (%) (rounded to 2 decimals)
    scrap_percentage = round(scrapped_cartridges / started_cartridges * 100, 2)
    print(f'Year {end_date.year} scrap: ', scrap_percentage)
