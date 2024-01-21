from helper_calculations import general_test_handler
from unittest import TestCase, TestLoader, TextTestRunner

# General tests ########################################################################################################
'''
General tests to fix behavior of general calculation methods.
Based on:
    - 2022 quarterly report "P 6-04 (a)"
    - Given trial balance
    - Given config file that contains only indicators needed for testing specific behavior. Not contains all indicators.
    - Given expected result that was manually tested and agreed
'''

# General test #1
'''
First test based on trial balance used in initial development
'''
result_1, expected_result_1 = general_test_handler(
    "calculations/general_2022_quarter_P_6_04_a/Helios - AZP - obratová předvaha.xls",
    "AZP - obratová předvaha", "A", "C", "D", "E",
    "calculations/general_2022_quarter_P_6_04_a/P_6-04_a.json",
    "calculations/general_2022_quarter_P_6_04_a/expected_result_1.json")

# General test #2
'''
Test after created method trial_balance_reader.account_col_modify (adding zeroes)
'''
result_2, expected_result_2 = general_test_handler(
    "calculations/general_2022_quarter_P_6_04_a/Forza Sole - obratová předvaha k 30062016.xlsx",
    "List1", "A", "D", "E", "G",
    "calculations/general_2022_quarter_P_6_04_a/P_6-04_a.json",
    "calculations/general_2022_quarter_P_6_04_a/expected_result_2.json")

# General test #3
'''
Test after created method trial_balance_reader.set_float
'''
result_3, expected_result_3 = general_test_handler(
    "calculations/general_2022_quarter_P_6_04_a/Helios - obratová předvaha 12_2017 po auditu.xlsx",
    "HELIOS Orange - Denní stavy účt", "A", "F", "G", "H",
    "calculations/general_2022_quarter_P_6_04_a/P_6-04_a.json",
    "calculations/general_2022_quarter_P_6_04_a/expected_result_3.json")

# General test #4
'''
Test for another trial balance to have as big sample as possible
'''
result_4, expected_result_4 = general_test_handler(
    "calculations/general_2022_quarter_P_6_04_a/Helios - TECHP předvaha 2021.xlsx",
    "Obratová předvaha (L-M)", "A", "D", "E", "J",
    "calculations/general_2022_quarter_P_6_04_a/P_6-04_a.json",
    "calculations/general_2022_quarter_P_6_04_a/expected_result_4.json")

# General test #5
'''
Test for another trial balance. This one is messy formated, but still read perfect by TrialBalance Reader :)
'''
result_5, expected_result_5 = general_test_handler(
    "calculations/general_2022_quarter_P_6_04_a/Obratova PREDVAHA 10_2015.xls",
    "Sheet1", "A", "E", "F", "G",
    "calculations/general_2022_quarter_P_6_04_a/P_6-04_a.json",
    "calculations/general_2022_quarter_P_6_04_a/expected_result_5.json")

# General test #6
'''
Test after method trial_balance_reader.account_col_modify upgraded (adding zeroes)
'''
result_6, expected_result_6 = general_test_handler(
    "calculations/general_2022_quarter_P_6_04_a/Obratová předvaha 2014.XLS",
    "Obratová předvaha 2014", "A", "G", "I", "M",
    "calculations/general_2022_quarter_P_6_04_a/P_6-04_a.json",
    "calculations/general_2022_quarter_P_6_04_a/expected_result_6.json")

# General test #7
'''
Test after method trial_balance_reader.account_col_modify upgraded (marking totals)
'''
result_7, expected_result_7 = general_test_handler(
    "calculations/general_2022_quarter_P_6_04_a/obratová předvaha 2014.xlsx",
    "List1", "B", "R", "V", "X",
    "calculations/general_2022_quarter_P_6_04_a/P_6-04_a.json",
    "calculations/general_2022_quarter_P_6_04_a/expected_result_7.json")

# General test #8
'''
Test after method trial_balance_reader.account_col_modify upgraded (marking totals)
'''
result_8, expected_result_8 = general_test_handler(
    "calculations/general_2022_quarter_P_6_04_a/Obratová předvaha po opr.rezerv a 425.xlsx",
    "List1", "A", "E", "G", "H",
    "calculations/general_2022_quarter_P_6_04_a/P_6-04_a.json",
    "calculations/general_2022_quarter_P_6_04_a/expected_result_8.json")

# Product tests ########################################################################################################
'''
Product tests to fix behavior of calculation methods used in particular products
Based on:
    - Template of report from bureau
    - Given trial balance
    - Given config file that contains indicators needed for testing product. Config file used directly in product.
    - Given expected result that was manually tested and agreed
'''

# 2023_quarter_P_6_04_a test #1
result_1_2023_quarter_P_6_04_a, expected_result_1_2023_quarter_P_6_04_a = general_test_handler (
    "calculations/2023_quarter_P_6_04_a/Helios - AZP - obratová předvaha.xls",
    "AZP - obratová předvaha", "A", "C", "D", "E",
    "../reports/2023/quarter/P_6-04_a.json",
    "calculations/2023_quarter_P_6_04_a/expected_result_1_2023_quarter_P_6_04_a.json")

# Assertion ############################################################################################################
class TestCalculations(TestCase):

    def test_1_general(self):
        self.assertEqual(result_1, expected_result_1)

    def test_2_general(self):
        self.assertEqual(result_2, expected_result_2)

    def test_3_general(self):
        self.assertEqual(result_3, expected_result_3)

    def test_4_general(self):
        self.assertEqual(result_4, expected_result_4)

    def test_5_general(self):
        self.assertEqual(result_5, expected_result_5)
    def test_6_general(self):
        self.assertEqual(result_6, expected_result_6)

    def test_7_general(self):
        self.assertEqual(result_7, expected_result_7)

    def test_8_general(self):
        self.assertEqual(result_8, expected_result_8)

    def test_1_2023_quarter_P_6_04_a(self):
        self.assertEqual(result_1_2023_quarter_P_6_04_a, expected_result_1_2023_quarter_P_6_04_a)

suite = TestLoader().loadTestsFromTestCase(TestCalculations)
TextTestRunner(verbosity=2).run(suite)
