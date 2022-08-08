from openpyxl import load_workbook

from correct_hours.report_processors.xero.rate_processor import XeroRateProcessor


def test_process_rates():
    workbook = load_workbook("tests/data/rates.xlsx")
    processor = XeroRateProcessor(workbook)
    rates = processor.process()
    print(rates)
    # TODO: implement test
