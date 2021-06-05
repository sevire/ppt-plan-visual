from unittest import TestCase
from ddt import ddt, data, unpack
import source.utilities as ut
from source.tests.testing_utilities import parse_date

first_day_of_month_test_data = [
    ('2021-03-21', '2021-03-01')
]


@ddt
class TestUtilities(TestCase):
    @data(*first_day_of_month_test_data)
    @unpack
    def test_first_day_of_month(self, string_date, from_string_date):
        date = parse_date(string_date)
        exp_result = parse_date(from_string_date)

        result = ut.first_day_of_month(date)

        self.assertEqual(exp_result, result)
