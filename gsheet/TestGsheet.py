__author__ = "Igor Gomes"
__email__ = "igor.mlgomes@gmail.com"

from unittest import TestCase
from gsheet.Gsheet import Gsheet


class TestGsheet(TestCase):
    _gs = Gsheet(spreadsheet_id="")  # Insert the spreadsheet's ID

    def test_get_credentials(self):
        self.assertIsNotNone(self._gs.get_credentials(credential=""))  # Insert the credentials' file name

    def test_get_service(self):
        self.assertIsNotNone(self._gs.get_service())

    def test_sheet_append_values(self):
        values = [['']]  # Insert the values to insert to the spreadsheet

        self._gs.sheet_append_values(range_="", values=values)  # Insert the range to insert the values

    def test_sheet_clear_values(self):
        self._gs.sheet_clear_values(range_="")  # Insert the range to clear in the spreadsheet

    def test_sheet_get_values(self):
        values = self._gs.sheet_get_values(range_="", value_render_option="").get('values')  # Insert the range to
        # search for and the value render option
        print(values)

    def test_sheet_update_values(self):
        values = [['']]  # Insert the values to be updated to the spreadsheet

        self._gs.sheet_update_values(range_='', values=values)  # Insert the range to be updated in the
        # spreadsheet
