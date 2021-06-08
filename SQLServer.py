import pyodbc
import logging
from tabulate import tabulate
from itertools import chain
from datetime import datetime
from KPISdata import get_dates_from_week_number
import pandas as pd


class SQLServer:

    # Columns in the SQL server
    """
    sn => Serial number
    time_lsn => Datetime Serial number printed
    LOT_lsn => Lot number
    GTIN_lsn => Lot GTIN
    timestamp_cve1 => Datetime first CVE (WS2)
    Pass_cve1 => Result first CVE
    decision_cve1
    decision_cve1_txt
    timestamp_cve2 => Datetime second CVE (WS7)
    Pass_cve2
    decision_cve2
    decision_cve2_txt
    timestamp_cve3 => Datetime third CVE (WS13)
    Pass_cve3
    decision_cve3
    decision_cve3_txt
    timestamp_lt
    Result_lt
    date_lp => Pouch label date (final step)
    date_lkp
    timetravel
    timetravel_txt
    last_step
    last_step_txt
    skipped_step
    status => Cartridge status: {1: 'PACKAGED', 2: 'FINISHED', 3: 'MISSING', 4: 'BAD', 5: 'NOT USED', 6: 'BAD AND FINISHED'}
    status_txt
    Text_mfm
    ID_mfm
    MFMGroup_mfm
    Name_mfm_g
    Active_mfm
    Name_prd
    in_qc
    lot_final_status_qc
    panel_qc
    lot_date =>
    line
    """

    def __init__(self, server, database, username, password, table):
        self.table = table

        setup_connection = ('DRIVER={};SERVER={};DATABASE={};UID={};PWD={}').format('{ODBC Driver 17 for SQL Server}',
                                                                                    server,
                                                                                    database,
                                                                                    username,
                                                                                    password)

        logging.warning('Estableciendo conexión: ' + setup_connection)

        self.connection = pyodbc.connect(setup_connection)
        self.cursor = self.connection.cursor()

        logging.warning('Establecida la conexión')

        self.columns_summary()

        self.lines = ['SSL2', 'SSL4']

    def __del__(self):
        self.cursor.commit()
        self.connection.close()

    def execute_query(self, query):
        self.cursor.execute(query)
        return self.cursor.fetchall()

    def columns_summary(self):
        query_info = """
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'qiastatdx_master_full'
            """

        info_columns_names = self.execute_query(query_info)

        # Convert from a list of tuples to a unique tuple to print horizontally
        info_column_names_one_tuple = [tuple(chain(*info_columns_names))]
        print(tabulate(info_column_names_one_tuple))

    def summary(self):
        superquery = ("""
        SELECT LOT_lsn, COUNT(DISTINCT sn) AS Started, MAX(lot_date) AS REPORTING_DATE, line AS LINE
        FROM {}
        WHERE line IN ('{}', '{}')
        GROUP BY LOT_lsn, line
        ORDER BY MAX(time_lsn) DESC
        """).format(self.table, self.lines[0], self.lines[1])

        print(tabulate(self.execute_query(superquery), headers=['Lot', 'Started', 'Report Date', 'Line']))

    def cartridges_split_by_line_ordered(self):
        """
        Information about: Line, Lot, Stsatus and Count.
        It tells how many cartridges are in each status for each lot, ordered by line and lot number
        """

        query = ("""
            SELECT line, LOT_lsn, status_txt, COUNT(DISTINCT(sn)) AS Counting_status
            FROM {0}
            GROUP BY LOT_lsn, line, status_txt, status
            HAVING line IN ('{1}', '{2}')
            ORDER BY CASE line WHEN '{1}' Then 0 WHEN '{2}' Then 1 ELSE 2 END ASC, LOT_lsn, status_txt
            """
        ).format(self.table, self.lines[0], self.lines[1])

        print(tabulate(self.execute_query(query), headers=['Line', 'Lot', 'Status', 'Count']))

    def started_cartridges_and_dates(self):

        query2 = (
            """
            SELECT LOT_lsn, COUNT(DISTINCT sn), line, max(timestamp_lt)
            FROM {} 
            GROUP BY LOT_lsn, line
            HAVING line IN ('{1}', '{2}')
            ORDER BY MAX(lot_date);
            """
        ).format(self.table, self.lines[0], self.lines[1])

        print(tabulate(self.execute_query(query2), headers=['LOT', 'STARTED', 'LINE', 'LOT_DATE']))


    def tabulate_data(self):

        query = ("""
        SELECT line, LOT_lsn, max(date_lp), 
        SUM(CASE WHEN status_txt NOT IN ('NOT USED') THEN 1 ELSE 0 END) AS Started,
        SUM(CASE WHEN status_txt IN ('FINISHED', 'PACKAGED') THEN 1 ELSE 0 END) AS Finished,
        lot_final_status_qc
        FROM {0}
        GROUP BY LOT_lsn, line, lot_final_status_qc
        HAVING Line IN ('{1}', '{2}')
        ORDER BY LOT_lsn;
        """).format(self.table, self.lines[0], self.lines[1])

        def get_scrap(row):
            try:
                scrap = round((row[3] - row[4]) / row[3] * 100, 2)
            except ZeroDivisionError:
                scrap = 0
            return row[0], row[1], row[2], scrap, row[3], row[4], row[5]

        fetched_data = self.execute_query(query)

        fetch_data = list(map(get_scrap, fetched_data))

        headers = ['Line', 'Lot', 'Report Date', 'Scrap', 'Started', 'Finished', 'QA Status']

        print(tabulate(fetch_data, headers=headers))

        df = pd.DataFrame(fetch_data, columns=headers)
        df.to_excel('Hilden_data.xlsx')

        # week_number =  int(input('Introduce week to plot'))
        # year = int(input('Introduce the year to be plotting'))
        week_number = 22
        year = 2021

        monday, sunday = get_dates_from_week_number(year, week_number)

        # Ensuring that the Report Data ('U2 Date') and Release Date are both datetime types
        df['Report Date'] = pd.to_datetime(df['Report Date'], format='%Y-%m-%d %H:%M:%S.%f')

        for entry in df['Report Date']:
            print(entry)

        lots_week = df.loc[(df['Report Date'] >= monday) & (df['Report Date'] <= sunday)]

        print(lots_week)

        started = lots_week['Started'].sum()
        finished = lots_week['Finished'].sum()

        print('Week {} Started: '.format(week_number), started)
        print('Week {} Finished: '.format(week_number), finished)
        # lots_week = df.loc([df['']])

        # while True:
        #     lote = int(input('información del lote: '))
        #     print(df.loc[df['Lot'] == lote])


if __name__ == '__main__':
    sql_server = SQLServer(server='prd-dl2-weu-rg.database.windows.net',
                           database='prd-dlsqlaccesslayer-weu-sqldb',
                           username='AccessManufacturingReportingReader',
                           password='KASKiajdjoqe^523jk',
                           table='manufacturing_reporting.qiastatdx_master'
                           )

    # Summary about the lots
    # sql_server.summary()

    # All kind of cartridges for each lot, ordered by line and then by lot number
    # sql_server.cartridges_split_by_line_ordered()

    sql_server.tabulate_data()

    # Selected dates
    date1 = datetime(2021, 5, 3)
    date2 = datetime(2021, 5, 9)

    # print('Selected date: ', date2.strftime('%d/%m/%Y'))

    # Finished by lots
    # query3 = ("""
    #     SELECT SUM(CASE WHEN status_txt IN ('FINISHED', 'PACKAGED') THEN 1 ELSE 0 END) AS Finished, line
    #     FROM {}
    #     WHERE date_lp BETWEEN '{}' AND '{}'
    #     GROUP BY line
    #     HAVING line IN ('SSL2', 'SSL4')
    #     ORDER BY line
    #     """).format(sql_server.table, date1, date2)

    query3 = ("""
        SELECT SUM(CASE WHEN status_txt IN ('FINISHED', 'PACKAGED') THEN 1 ELSE 0 END) AS Finished
        FROM {}
        WHERE date_lp BETWEEN '{}' AND '{}'
        AND line IN ('SSL2', 'SSL4')
        """).format(sql_server.table, date1, date2)

    print(tabulate(sql_server.execute_query(query3), headers=['Finished', 'Line']))

    query4 = ("""
        SELECT COUNT(*) - SUM(CASE WHEN status_txt IN ('NOT USED') THEN 1 ELSE 0 END) AS Started, line
        FROM {}
        WHERE time_lsn BETWEEN '{}' AND '{}'
        GROUP BY line
        HAVING line IN ('SSL2', 'SSL4')
        ORDER BY line
    """).format(sql_server.table, date1, date2)

    # print(tabulate(sql_server.execute_query(query4), headers=['Started', 'Line']))

    # Checking if the finished query has added the 'BAD AND FINISHED' too
    query5 = ("""
        SELECT SUM(CASE WHEN status_txt IN ('FINISHED', 'PACKAGED', 'BAD AND FINISHED') THEN 1 ELSE 0 END) AS Finished, line
        FROM {}
        WHERE date_lp BETWEEN '{}' AND '{}'
        GROUP BY line
        HAVING line IN ('SSL1', 'SSL3')
        ORDER BY line
        """).format(sql_server.table, date1, date2)

    # print(tabulate(sql_server.execute_query(query5), headers=['All Finished', 'Line']))

    query6 = ("""
        SELECT *
        FROM {}
        WHERE LOT_lsn = {};
    """).format(sql_server.table, 210111)

    # print(tabulate(sql_server.execute_query(query6)))

    # Started cartridges by lot
    query7 = ("""
        SELECT COUNT(*) - SUM (CASE WHEN status_txt = 'NOT USED' THEN 1 ELSE 0 END)
        FROM {}
        WHERE LOT_lsn = time_lsn;
    """).format(sql_server.table, 21219)
