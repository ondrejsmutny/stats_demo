from trial_balance_reader import TrialBalanceReader
from calculations import Calculator
from json import load

def general_test_handler(file_path, sheet, account_col, debit_turnover_col, credit_turnover_col, end_balance_col,
                         report_config_path, expected_result_path):
    '''
    Handling all instances and methods needed for generating results to be tested

    Args:
        file_path (str): For TrialBalanceReader
        sheet (str): For TrialBalanceReader
        account_col (str): For TrialBalanceReader
        debit_turnover_col (str): For TrialBalanceReader
        credit_turnover_col (str): For TrialBalanceReader
        end_balance_col (str): For TrialBalanceReader
        report_config_path (str): For Calculator
        expected_result_path (str): Path to file containing expected result that was manually tested and agreed

    Returns:
        result (dict): Generated result of running the tested methods
        expected_result (dict) Expected result for assertion with generated result
    '''
    reader = TrialBalanceReader(file_path, sheet, account_col, debit_turnover_col, credit_turnover_col, end_balance_col)

    final_df = reader.get_final_df()

    calculator = Calculator(report_config_path, final_df)

    result = calculator.calculation_handler()

    expected_result_file = open(expected_result_path)
    expected_result = load(expected_result_file)
    expected_result_file.close()

    return result, expected_result
