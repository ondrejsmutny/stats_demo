from sys import argv
from json import load
from pandas import ExcelFile
from string import ascii_uppercase
from trial_balance_reader import TrialBalanceReader
from calculations import Calculator
from report_writer import ReportWriter
from os import path
from PyQt6 import QtWidgets, QtGui
from secret_keys import encrypt_message_with_key
from requests import get

class App(QtWidgets.QApplication):
    '''
    Main application

    Args:
        logger (instance): Instance of logger provided by executive file
        timestamp (str): Stamp with time when executable file was run
    '''

    def __init__(self, logger, timestamp):
        super(App, self).__init__(argv)

        self.logger = logger
        self.timestamp = timestamp

    def build(self, report_config_path, report_template_path):
        '''
        Initiates necessary forms instances for application

        Args:
            report_config_path (str): Path to json file with configuration of selected report
            report_template_path (str): Path to file with report template
        '''
        # Loading report info
        self.report_config_path = report_config_path
        self.report_template_path = report_template_path

        with open(self.report_config_path, encoding="utf-8") as self.report_config_file:
            self.report_config = load(self.report_config_file)

        self.report_info = self.report_config.get("Info").get("01").get("1")
        self.report_config_file.close()

        # Forms initiating
        self.userAuthenticationForm = UserAuthenticationForm(self, self.logger, self.report_info)
        self.trialBalanceLoadForm = TrialBalanceLoadForm(self, self.logger, self.report_info)
        self.trialBalanceSetForm = TrialBalanceSetForm(self, self.logger, self.report_info, self.timestamp,
                                                       self.report_config_path, self.report_template_path)

        # Forms setting up
        self.userAuthenticationForm.setup()
        self.trialBalanceLoadForm.setup()
        self.trialBalanceSetForm.setup()

        self.exit(self.exec())

class UserAuthenticationForm(QtWidgets.QWidget):
    '''
    Processing authentication of user. Validation of users license.

     Args:
         root (instance): Instance of App class, the main root
         logger (instance): Instance of logger provided by executive file
         report_info (dict): Includes info part of configuration file
    '''

    def __init__(self, root, logger, report_info, **kwargs):
        super(UserAuthenticationForm, self).__init__(**kwargs)

        # Setting main instances
        self.root = root
        self.logger = logger

        # Getting report information
        self.report_info = report_info
        self.report_year = self.report_info.get('report_year')
        self.report_interval = self.report_info.get('report_interval')
        self.report_code = self.report_info.get('report_code')

        # Window settings
        self.setWindowTitle(f"StatS: {self.report_year} {self.report_interval} {self.report_code}")
        self.setWindowIcon(QtGui.QIcon("icons/form_2_64px.png"))
        self.setMinimumWidth(1100)

        # Main layout
        mainLayout = QtWidgets.QVBoxLayout()
        self.setLayout(mainLayout)
        mainLayout.addStretch()

        # Initial info
        self.infoLabel = QtWidgets.QLabel(f"Tato aplikace slouží pro zpracování výkazu: Rok {self.report_year},"
                                          f" {self.report_interval} výkaz {self.report_code}")
        infoLabelFont = QtGui.QFont()
        infoLabelFont.setBold(True)
        self.infoLabel.setFont(infoLabelFont)
        mainLayout.addWidget(self.infoLabel)

        # Authentication info
        authenticationLabel = QtWidgets.QLabel("Zadejte prosím uživatelské údaje, kterými jste se registrovali na webu"
                                               " dcba.cz, aby mohla být ověřena platnost licence.")
        mainLayout.addWidget(authenticationLabel)

        # Read last known existing username
        with open("logs/username_stamp.txt", "r") as username_stamp_file:
            username_stamp = username_stamp_file.read()
        username_stamp_file.close()

        # Username input
        userNameLayout = QtWidgets.QHBoxLayout()
        userNameLayout.addWidget(QtWidgets.QLabel("Email uživatele:"))
        userNameLayout.addStretch()
        self.username = QtWidgets.QLineEdit(username_stamp)
        userNameLayout.addWidget(self.username)
        mainLayout.addLayout(userNameLayout)

        # User password input
        userPasswordLayout = QtWidgets.QHBoxLayout()
        userPasswordLayout.addWidget(QtWidgets.QLabel("Heslo uživatele:"))
        userPasswordLayout.addStretch()
        self.password = QtWidgets.QLineEdit()
        self.password.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        userPasswordLayout.addWidget(self.password)
        mainLayout.addLayout(userPasswordLayout)

        # Password change info
        passwordChangeLabel = QtWidgets.QLabel("Zapomněli jste heslo nebo jej chcete změnit? Navštivte web dcba.cz")
        mainLayout.addWidget(passwordChangeLabel)

        # Login layout
        loginLayout = QtWidgets.QHBoxLayout()
        self.loginErrorLabel = (QtWidgets.QLabel(""))
        self.loginErrorLabel.setStyleSheet("color: red;")
        loginLayout.addWidget(self.loginErrorLabel)
        loginLayout.addStretch()
        self.loginButton = QtWidgets.QPushButton("Přihlásit")
        self.loginButton.clicked.connect(self.loginButtonClicked)
        loginLayout.addWidget(self.loginButton)
        mainLayout.addLayout(loginLayout)

        # End of form
        mainLayout.addStretch()
        self.show()

    def setup(self):
        '''
        Sets up references to other forms
        '''
        self.trialBalanceLoadForm = self.root.trialBalanceLoadForm

    def loginButtonClicked(self):
        '''
        Gets user credentials, encrypts them and sends them to web. Checks if user can continue to application based on
        information provided by response from web.
        '''
        self.loginErrorLabel.setText("")
        try:
            # Setting encryption
            KEY = b'AAA='
            KEY_FILENAME = "aaa.key"

            # Getting user credentials
            self.username_text = self.username.text()
            self.password_text = self.password.text()

            self.username_encrypted = encrypt_message_with_key(self.username_text, KEY, KEY_FILENAME)
            self.password_encrypted = encrypt_message_with_key(self.password_text, KEY, KEY_FILENAME)

            # Sending get request
            credentials = {"username_encrypted": self.username_encrypted, "username_key_filename": KEY_FILENAME,
                           "password_encrypted": self.password_encrypted, "password_key_filename": KEY_FILENAME,
                           "product_name": self.report_info.get("product_name")}
            response = get("https://dcba.cz/user_zone/external_request/", params=credentials)  # For production
            # response = get("http://127.0.0.1:8000/user_zone/external_request/", params=credentials)  # For mirror

            # Write last known existing username
            if response.headers.get("user_exists") == "True":
                with open("logs/username_stamp.txt", "w") as username_stamp_file:
                    username_stamp_file.write(self.username_text)
                username_stamp_file.close()

            # Checking if user can continue to application
            if response.headers.get("user_exists") != "True":
                self.loginErrorLabel.setText("Neznámý uživatel. Zkontrolujte správné zadání emailu nebo se registrujte"
                                             " na webu dcba.cz.")

            elif response.headers.get("user_is_authenticated") != "True":
                self.loginErrorLabel.setText("Neplatné heslo. Zkuste zadat znovu nebo navštivte web dcba.cz.")

            elif response.headers.get("licence_is_valid") != "True":
                self.loginErrorLabel.setText("Licence vypršela. Pro objednání nové navštivte web dcba.cz.")

            else:
                self.trialBalanceLoadForm.show()
                self.close()

        except:
            self.logger.exception("")
            self.loginErrorLabel.setText("Nepodařilo se spojit s webem dcba.cz.")

class TrialBalanceLoadForm(QtWidgets.QWidget):
    '''
    Form to load trial balance file by user selection

     Args:
         root (instance): Instance of App class, the main root
         logger (instance): Instance of logger provided by executive file
         report_info (dict): Includes info part of configuration file
    '''

    def __init__(self, root, logger, report_info, **kwargs):
        super(TrialBalanceLoadForm, self).__init__(**kwargs)

        # Setting main instances
        self.root = root
        self.logger = logger

        # Getting report information
        self.report_info = report_info
        self.report_year = self.report_info.get('report_year')
        self.report_interval = self.report_info.get('report_interval')
        self.report_code = self.report_info.get('report_code')

        self.trial_balance_path = ""

        # Window settings
        self.setWindowTitle(f"StatS: {self.report_year} {self.report_interval} {self.report_code}")
        self.setWindowIcon(QtGui.QIcon("icons/form_2_64px.png"))
        self.setMinimumWidth(1100)

        # Main layout
        mainLayout = QtWidgets.QVBoxLayout()
        self.setLayout(mainLayout)
        mainLayout.addStretch()

        # Initial info
        self.infoLabel = QtWidgets.QLabel(f"Tato aplikace slouží pro zpracování výkazu: Rok {self.report_year},"
                                     f" {self.report_interval} výkaz {self.report_code}")
        infoLabelFont = QtGui.QFont()
        infoLabelFont.setBold(True)
        self.infoLabel.setFont(infoLabelFont)
        mainLayout.addWidget(self.infoLabel)

        # Choose trial balance to load
        loadLayout = QtWidgets.QHBoxLayout()
        loadLayout.addWidget(QtWidgets.QLabel("Vyberte prosím excelový soubor, který obsahuje obratovou předvahu za"
                                               " sledované období:"))
        loadLayout.addStretch()
        self.loadButton = QtWidgets.QPushButton("Vybrat soubor")
        self.loadButton.clicked.connect(self.loadButtonClicked)
        loadLayout.addWidget(self.loadButton)
        mainLayout.addLayout(loadLayout)

        # Show loaded trial balance
        self.trial_balance_loaded = QtWidgets.QLabel(f"Vybraný soubor: {self.trial_balance_path}")
        mainLayout.addWidget(self.trial_balance_loaded)

        # Continue layout
        continueLayout = QtWidgets.QHBoxLayout()
        self.continueErrorLabel = (QtWidgets.QLabel(""))
        self.continueErrorLabel.setStyleSheet("color: red;")
        continueLayout.addWidget(self.continueErrorLabel)
        continueLayout.addStretch()
        self.continueButton = QtWidgets.QPushButton("Pokračovat")
        self.continueButton.clicked.connect(self.continueButtonClicked)
        continueLayout.addWidget(self.continueButton)
        mainLayout.addLayout(continueLayout)

        # End of form
        mainLayout.addStretch()

    def setup(self):
        '''
        Sets up references to other forms
        '''
        self.trialBalanceSetForm = self.root.trialBalanceSetForm

    def loadButtonClicked(self):
        """
        Loads trial balance by user selection
        """
        try:
            self.trial_balance_path, self.filter = QtWidgets.QFileDialog.getOpenFileName(
                parent=self, caption="Vybrat soubor", directory=".", filter="Excel (*.xlsx *.xls)")
            self.trial_balance_loaded.setText(f"Vybraný soubor: {self.trial_balance_path}")
            self.continueErrorLabel.setText("")
        except:
            self.logger.exception("")

    def continueButtonClicked(self):
        '''
        Checks if user selected trial balance and provides connection to next form
        '''
        if self.trial_balance_path == "":
            self.continueErrorLabel.setText("Nelze pokračovat bez vybrání souboru s obratovou předvahou!")
        else:
            try:
                self.trialBalanceSetForm.set_trial_balance_path(self.trial_balance_path)
                self.trialBalanceSetForm.setExcelSheetsBox()
                self.trialBalanceSetForm.show()
                self.close()
            except:
                self.logger.exception("")

class TrialBalanceSetForm(QtWidgets.QWidget):
    '''
    Allows the user to set trial balance columns needed for following calculation. Handles instances of all classes
    needed for application logic. Handles all methods needed for application logic.

    Args:
        root (instance): Instance of App class, the main root
        logger (instance): Instance of logger provided by executive file
        report_info (dict): Includes info part of configuration file
        timestamp (str): Stamp with time when executable file was run
        report_config_path (str): Path to json file with configuration of selected report
        report_template_path (str): Path to file with report template
    '''

    def __init__(self, root, logger, report_info, timestamp, report_config_path, report_template_path, **kwargs):
        super(TrialBalanceSetForm, self).__init__(**kwargs)

        # Setting main instances
        self.root = root
        self.logger = logger

        # Getting report information ###################################################################################
        self.report_info = report_info
        self.report_year = self.report_info.get('report_year')
        self.report_interval = self.report_info.get('report_interval')
        self.report_code = self.report_info.get('report_code')
        self.report_code_snake_case = self.report_info.get("report_code_snake_case")
        self.warning_tailor = self.report_info.get("warning_tailor")
        self.timestamp = timestamp
        self.report_config_path = report_config_path
        self.report_template_path = report_template_path

        self.trial_balance_path = ""

        # Creating bold font ###########################################################################################
        boldFont = QtGui.QFont()
        boldFont.setBold(True)

        # Window settings ##############################################################################################
        self.setWindowTitle(f"StatS: {self.report_year} {self.report_interval} {self.report_code}")
        self.setWindowIcon(QtGui.QIcon("icons/form_2_64px.png"))
        self.setMinimumWidth(1100)

        # Main layout ##################################################################################################
        mainLayout = QtWidgets.QVBoxLayout()
        self.setLayout(mainLayout)
        mainLayout.addStretch()

        # Initial info #################################################################################################
        self.trial_balance_loaded = QtWidgets.QLabel(f"Vybraný soubor: {self.trial_balance_path}")
        mainLayout.addWidget(self.trial_balance_loaded)
        initialLine = QtWidgets.QFrame()
        initialLine.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        mainLayout.addWidget(initialLine)

        # Setting sheet with trial balance from file ###################################################################
        sheetSetLayout = QtWidgets.QHBoxLayout()
        self.sheetSetLabel = QtWidgets.QLabel("Vyberte prosím list, na kterém se v excelovém souboru nachází"
                                              " obratová předvaha:")
        self.sheetSetLabel.setFont(boldFont)
        sheetSetLayout.addWidget(self.sheetSetLabel)
        sheetSetLayout.addStretch()
        self.excelSheetsBox = QtWidgets.QComboBox()
        self.excelSheetsBox.addItem("-")
        sheetSetLayout.addWidget(self.excelSheetsBox)
        mainLayout.addLayout(sheetSetLayout)
        sheetSetLine = QtWidgets.QFrame()
        sheetSetLine.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        mainLayout.addWidget(sheetSetLine)

        # Setting columns of the trial balance for report calculation ##################################################

        # Account column
        accountColLayout = QtWidgets.QHBoxLayout()
        self.accountColLabel = QtWidgets.QLabel("Sloupec s čísly účtů na listu:")
        self.accountColLabel.setFont(boldFont)
        accountColLayout.addWidget(self.accountColLabel)
        accountColLayout.addStretch()
        self.accountColBox = QtWidgets.QComboBox()
        self.accountColBox.addItem("-")
        self.accountColBox.addItems(list(ascii_uppercase))
        accountColLayout.addWidget(self.accountColBox)
        mainLayout.addLayout(accountColLayout)
        mainLayout.addWidget(QtWidgets.QLabel("(Musí obsahovat účty začínající prvním trojčíslím analytických účtů dle"
                                              " obecně uznávaných českých účetních standardů.)"))
        accountColLine = QtWidgets.QFrame()
        accountColLine.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        mainLayout.addWidget(accountColLine)

        # Debit turnover column
        debitTurnoverColLayout = QtWidgets.QHBoxLayout()
        self.debitTurnoverColLabel = QtWidgets.QLabel("Sloupec s obraty Má Dáti na listu:")
        self.debitTurnoverColLabel.setFont(boldFont)
        debitTurnoverColLayout.addWidget(self.debitTurnoverColLabel)
        debitTurnoverColLayout.addStretch()
        self.debitTurnoverColBox = QtWidgets.QComboBox()
        self.debitTurnoverColBox.addItem("-")
        self.debitTurnoverColBox.addItems(list(ascii_uppercase))
        debitTurnoverColLayout.addWidget(self.debitTurnoverColBox)
        mainLayout.addLayout(debitTurnoverColLayout)
        mainLayout.addWidget(QtWidgets.QLabel("(Musí obsahovat obraty od začátku vykazovaného období, nikoliv od"
                                              " začátku roku. Začátek roku ale může být někdy shodný se začátkem"
                                              " vykazovaného období.)"))
        debitTurnoverColLine = QtWidgets.QFrame()
        debitTurnoverColLine.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        mainLayout.addWidget(debitTurnoverColLine)

        # Credit turnover column
        creditTurnoverColLayout = QtWidgets.QHBoxLayout()
        self.creditTurnoverColLabel = QtWidgets.QLabel("Sloupec s obraty Dal na listu:")
        self.creditTurnoverColLabel.setFont(boldFont)
        creditTurnoverColLayout.addWidget(self.creditTurnoverColLabel)
        creditTurnoverColLayout.addStretch()
        self.creditTurnoverColBox = QtWidgets.QComboBox()
        self.creditTurnoverColBox.addItem("-")
        self.creditTurnoverColBox.addItems(list(ascii_uppercase))
        creditTurnoverColLayout.addWidget(self.creditTurnoverColBox)
        mainLayout.addLayout(creditTurnoverColLayout)
        mainLayout.addWidget(QtWidgets.QLabel("(Musí obsahovat obraty od začátku vykazovaného období, nikoliv od"
                                              " začátku roku. Začátek roku ale může být někdy shodný se začátkem"
                                              " vykazovaného období.)"))
        creditTurnoverColLine = QtWidgets.QFrame()
        creditTurnoverColLine.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        mainLayout.addWidget(creditTurnoverColLine)

        # End balance column
        endBalanceColLayout = QtWidgets.QHBoxLayout()
        self.endBalanceColLabel = QtWidgets.QLabel("Sloupec s konečnými zůstatky na listu:")
        self.endBalanceColLabel.setFont(boldFont)
        endBalanceColLayout.addWidget(self.endBalanceColLabel)
        endBalanceColLayout.addStretch()
        self.endBalanceColBox = QtWidgets.QComboBox()
        self.endBalanceColBox.addItem("-")
        self.endBalanceColBox.addItems(list(ascii_uppercase))
        endBalanceColLayout.addWidget(self.endBalanceColBox)
        mainLayout.addLayout(endBalanceColLayout)
        mainLayout.addWidget(QtWidgets.QLabel("(Musí obsahovat zůstatky účtů na konci vykazovaného období. Vzorec pro"
                                              " výpočet je KZ = Počáteční stav + Obrat MD - Obrat D)"))
        endBalanceColLine = QtWidgets.QFrame()
        endBalanceColLine.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        mainLayout.addWidget(endBalanceColLine)

        # Calculation ##################################################################################################
        self.analyticalAccWarnLabel = QtWidgets.QLabel(self.warning_tailor)
        self.analyticalAccWarnLabel.setStyleSheet("color: #FF8C00;")
        mainLayout.addWidget(self.analyticalAccWarnLabel)

        calculationLayout = QtWidgets.QHBoxLayout()
        self.calculationErrorLabel = (QtWidgets.QLabel(""))
        self.calculationErrorLabel.setStyleSheet("color: red;")
        calculationLayout.addWidget(self.calculationErrorLabel)
        calculationLayout.addStretch()
        self.calculationButton = QtWidgets.QPushButton("Vypočítat")
        self.calculationButton.clicked.connect(self.calculationButtonClicked)
        calculationLayout.addWidget(self.calculationButton)
        mainLayout.addLayout(calculationLayout)

        # End of form
        mainLayout.addStretch()

    def setup(self):
        pass

    def set_trial_balance_path(self, trial_balance_path):
        '''
        Takes over path of file with trial balance. Sets it to label in form.

        Args:
            trial_balance_path (str): Path to the excel file with trial balance
        '''
        try:
            self.trial_balance_path = trial_balance_path
            self.trial_balance_loaded.setText(f"Vybraný soubor: {self.trial_balance_path}")
        except:
            self.logger.exception("")

    def setExcelSheetsBox(self):
        '''
        Finds all sheets in excel file and sets them to the box in form.
        '''
        try:
            excel_sheets_file = ExcelFile(self.trial_balance_path)
            self.excel_sheets = excel_sheets_file.sheet_names
            self.excelSheetsBox.addItems(self.excel_sheets)
        except:
            self.logger.exception("")

    def calculationButtonClicked(self):
        '''
        Checks if selection by user was made in all boxes. Handles instances of all classes needed for application
        logic. Handles all methods needed for application logic.
        '''
        if self.excelSheetsBox.currentText() == "-" or self.accountColBox.currentText() == "-"\
                or self.debitTurnoverColBox.currentText() == "-" or self.creditTurnoverColBox.currentText() == "-"\
                or self.endBalanceColBox.currentText() == "-":
            self.calculationErrorLabel.setText("Nelze vypočítat! Nebyla provedena volba listu a všech slopců výše.")
        else:
            try:
                # Reading trial balance and forming it to shape needed in Calculator class
                reader = TrialBalanceReader(self.trial_balance_path,
                                                                 self.excelSheetsBox.currentText(),
                                                                 self.accountColBox.currentText(),
                                                                 self.debitTurnoverColBox.currentText(),
                                                                 self.creditTurnoverColBox.currentText(),
                                                                 self.endBalanceColBox.currentText())
                final_df = reader.get_final_df()

                # Calculating report
                calculator = Calculator(self.report_config_path, final_df)
                report_calculated = calculator.calculation_handler()

                # In case of need:
                # calculator.save_report_to_json(self.timestamp)

                # Creating output file name in the same folder where the trial balance is located
                report_output_dirname = path.dirname(self.trial_balance_path)
                self.report_output_path = f"{report_output_dirname}/{self.report_code_snake_case}_{self.timestamp}.pdf"

                # Writing calculated report into output file
                writer = ReportWriter(report_calculated, self.report_template_path, self.report_output_path)
                report_output = writer.write_report_by_watermark()

                # Changing error label
                self.calculationErrorLabel.setStyleSheet("color: green;")
                self.calculationErrorLabel.setText(f"Výkaz byl vypočítán a uložen do souboru: {self.report_output_path}")
            except:
                self.logger.exception("")

'''
Inspiration for running this application
'''
# from sys import executable
# from os import path, environ
# from datetime import datetime
# from time import time
# from logging import basicConfig, FileHandler, StreamHandler, DEBUG, getLogger
# from PyQt6 import QtWidgets
# from ui import App
#
# if __name__ == "__main__":
#     # Getting paths
#     application_path = path.dirname(executable)
#
#     timestamp = datetime.fromtimestamp(time()).strftime("%Y-%m-%d_%H-%M-%S")
#     log_path = f"logs/log_{timestamp}.txt"
#
#     pyqt_path = path.dirname(QtWidgets.__file__)
#     plugin_path = path.join(pyqt_path, "Qt6\plugins")
#     environ['QT_PLUGIN_PATH'] = plugin_path
#
#     # Creating logger
#     log_format = "[%(levelname)s] - %(asctime)s - %(name)s - : %(message)s in %(pathname)s:%(lineno)d"
#     basicConfig(handlers=[FileHandler(log_path), StreamHandler()], level=DEBUG,
#                         format=log_format)
#     logger = getLogger(__name__)
#     logger.info(f"Log file created with timestamp: {timestamp}\n")
#
#     # Running application
#     try:
#         root = App(logger, timestamp)
#         root.build("reports/2022/quarter/P_6-04_a.json", "reports/2022/quarter/P_6-04_a.pdf")
#
#         timestamp_end = datetime.fromtimestamp(time()).strftime("%Y-%m-%d_%H-%M-%S")
#         logger.info(f"Logging ended with timestamp end: {timestamp_end}\n")
#     except:
#         logger.exception("")
