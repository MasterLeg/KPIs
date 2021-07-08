import unittest
from SQLServer import SQLServer
from tabulate import tabulate
from HIldenKPI import Hilden_Dataframe
from time import time


class MyTestCase(unittest.TestCase):
    def test_something(self):
        server = SQLServer(server='prd-dl2-weu-rg.database.windows.net',
                           database='prd-dlsqlaccesslayer-weu-sqldb',
                           username='AccessManufacturingReportingReader',
                           password='KASKiajdjoqe^523jk',
                           table='manufacturing_reporting.qiastatdx_master'
                           )

        print(server.get_information_report_lots())
        print('\n\n\n')

        print(tabulate(
            server.get_information_report_lots(),
            headers=['Lot number',
                     'Reporting Date',
                     'Production Date',
                     'Release Date',
                     'Started',
                     'Finished',
                     'Status',
                     'Released/Blocked kits'])
        )

        self.assertEqual(True, True)

        tic = time()

        Hilden_Dataframe()

        print('Time required to execute the algorithm:\t', time() - tic)

        self.assertEqual(True, True)


if __name__ == '__main__':
    unittest.main()
