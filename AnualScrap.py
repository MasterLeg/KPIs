from datetime import datetime, timedelta


def generate_anual_scrap_report(df):
    # TODO: Cambiar el annual scrap: desde el 1 de enero hacer Monthly scrap de los lotes
    # OPTION 1: Mobile window
    # Dates to select
    end_date = datetime.today()
    start_date = end_date - timedelta(days=365)

    # OPTION 2: Fixed window: select natural year starting and ending date
    #   year = int(input('Año a calcular:\t'))
    #   start_date = datetime(year, 1, 1)
    #  end_date = datetime(year, 12, 31)

    # Select the rows that are in the desired week
    week_lots = df.loc[(df['Reporting Date'] >= start_date) & (df['Reporting Date'] <= end_date)]

    finished_cartridges = int(week_lots['Finished'].sum())
    started_cartridges = int(week_lots['Started'].sum())
    scrapped_cartridges = started_cartridges - finished_cartridges
    scrap_percentage = round(scrapped_cartridges / started_cartridges * 100, 2)
    print(f'Year {end_date.year} scrap: ', scrap_percentage)
