from Config import NAZOV_APP, NAZOV_FIRMY,HEIGHT,WIDTH,X_POSITION,Y_POSITION
import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QPushButton, QComboBox, QTableWidget, QTableWidgetItem, 
                            QCalendarWidget, QCheckBox, QGroupBox, QMessageBox, QFileDialog, QFrame, 
                            QScrollArea, QSizePolicy,QCompleter,QAbstractItemView)
from PyQt6.QtCore import Qt, QDate, QSettings
from PyQt6.QtGui import QFont, QAction
import sqlite3
import openpyxl
from openpyxl import Workbook
import locale

class FrmGarden(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Nastavte locale pre formátovanie (môžete si vybrať podľa potreby)
        locale.setlocale(locale.LC_NUMERIC, 'en_US.UTF-8')

        self.setWindowTitle(NAZOV_APP)
        self.setGeometry(X_POSITION, Y_POSITION, WIDTH, HEIGHT)
        
        # Inicializácia nastavení
        self.settings = QSettings(NAZOV_FIRMY, NAZOV_APP)
        
        # Vytvorenie hlavného scroll area
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        
        # Hlavný widget, ktorý bude obsahovať všetky prvky
        self.main_widget = QWidget()
        self.main_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        # Hlavný layout pre main_widget
        self.main_layout = QVBoxLayout(self.main_widget)
        self.main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Inicializácia UI
        self.init_ui()
        
        # Nastavenie scroll area
        self.scroll.setWidget(self.main_widget)
        self.setCentralWidget(self.scroll)
        
        self.setup_database()
        self.load_initial_data()

        self.load_unique_clients()  # Načítanie klientov z databázy
        # completer = QCompleter(self.unique_clients, self)
        # completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)  # Ignorovanie veľkosti písmen
        # self.txtClients.setCompleter(completer)

    def init_ui(self):
        # Vytvorenie menu bar
        menubar = self.menuBar()
        
        # Vytvorenie menu "Zobrazenie"
        view_menu = menubar.addMenu("Font")
        
        # Vytvorenie akcií pre veľkosť písma
        male_pismo_action = QAction("Small", self)
        stredne_pismo_action = QAction("Medium", self)
        velke_pismo_action = QAction("Large", self)
        
        # Pripojenie akcií k handlerom
        male_pismo_action.triggered.connect(lambda: self.zmen_velkost_pisma(10))
        stredne_pismo_action.triggered.connect(lambda: self.zmen_velkost_pisma(12))
        velke_pismo_action.triggered.connect(lambda: self.zmen_velkost_pisma(14))
        
        # Pridanie akcií do menu
        view_menu.addAction(male_pismo_action)
        view_menu.addAction(stredne_pismo_action)
        view_menu.addAction(velke_pismo_action)
        
        # Načítanie uloženej veľkosti písma alebo použitie predvolenej
        ulozena_velkost = self.settings.value("velkost_pisma", 12, type=int)
        self.zmen_velkost_pisma(ulozena_velkost, ulozit_nastavenie=False)
        
        # Header frame
        header_frame = QFrame()
        header_frame.setFrameShape(QFrame.Shape.StyledPanel)
        header_layout = QHBoxLayout(header_frame)

        # Vytvorenie results_group (nahradzuje pôvodné gbResult)
        self.results_group = QGroupBox("RESULTS")
        results_layout = QVBoxLayout(self.results_group)
        
        # Add title label
        title_label = QLabel(NAZOV_APP)
        title_label.setFont(QFont('Arial', 16, QFont.Weight.Bold))
        header_layout.addWidget(title_label)
        
        # Add year/month selection
        self.cbRok = QComboBox()
        self.cbRok.addItems([str(year) for year in range(2020, 2050)])
        self.cbRok.currentIndexChanged.connect(self.cbRok_SelectedIndexChanged)
        
        self.cbMonths = QComboBox()
        self.cbMonths.addItems(["All", "January", "February", "March", "April", "May", "June", 
                              "July", "August", "September", "October", "November", "December"])
        self.cbMonths.currentIndexChanged.connect(self.cbMonths_SelectedIndexChanged)
        
        header_layout.addWidget(QLabel("Year:"))
        header_layout.addWidget(self.cbRok)
        header_layout.addWidget(QLabel("Month:"))
        header_layout.addWidget(self.cbMonths)
        
        self.main_layout.addWidget(header_frame)
        
        # Main content area
        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        
        # Left panel - Input section
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.Shape.StyledPanel)
        left_layout = QVBoxLayout(left_panel)
        
        # Calendar
        self.Calendar = QCalendarWidget()
        self.Calendar.setGridVisible(True)
        self.Calendar.clicked.connect(self.Calendar_DateSelected)
        left_layout.addWidget(self.Calendar)
        
        # Input fields
        input_group = QGroupBox("Transaction Details")
        input_layout = QVBoxLayout(input_group)
        
        # Date input
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Date:"))
        self.txtSelectDate = QLineEdit()
        self.txtSelectDate.setPlaceholderText("YYYY-MM-DD")
        date_layout.addWidget(self.txtSelectDate)
        input_layout.addLayout(date_layout)
        
        # Client input
        client_layout = QHBoxLayout()
        client_layout.addWidget(QLabel("Client:"))
        self.txtClients = QLineEdit()
        self.txtClients.setPlaceholderText("Client Name")
        client_layout.addWidget(self.txtClients)
        input_layout.addLayout(client_layout)
        
        # Income section
        income_group = QGroupBox("INCOME")
        income_layout = QVBoxLayout(income_group)
        
        self.txtCash = QLineEdit()
        self.txtCash.setPlaceholderText("Cash")
        income_layout.addWidget(QLabel("CASH:"))
        income_layout.addWidget(self.txtCash)
        
        self.txtCheck = QLineEdit()
        self.txtCheck.setPlaceholderText("Check")
        income_layout.addWidget(QLabel("CHECK:"))
        income_layout.addWidget(self.txtCheck)
        
        self.txtBank = QLineEdit()
        self.txtBank.setPlaceholderText("Bank Transfer")
        income_layout.addWidget(QLabel("BANK TRANSFER:"))
        income_layout.addWidget(self.txtBank)
        
        input_layout.addWidget(income_group)
      
        # Expenses section
        expenses_group = QGroupBox("EXPENSE")
        expenses_layout = QHBoxLayout(expenses_group)  # Zmena na horizontálny layout

        # Ľavá strana - EXPENSES a COST
        left_expenses = QVBoxLayout()
        self.txtExpenses = QLineEdit()
        self.txtExpenses.setPlaceholderText("Expenses Description")
        self.txtExpensesCost = QLineEdit()
        self.txtExpensesCost.setPlaceholderText("Cost of Expenses")

        left_expenses.addWidget(QLabel("EXPENSES:"))
        left_expenses.addWidget(self.txtExpenses)
        left_expenses.addWidget(QLabel("COST COMPLETE:"))
        left_expenses.addWidget(self.txtExpensesCost)

        # Pravá strana - EXPENSES CASH a CASH FOR STAFF
        right_expenses = QVBoxLayout()
        self.txtCashForStaff = QLineEdit()
        self.txtCashForStaff.setPlaceholderText("Cash for staff")
        # self.txtCashForStaff.setValidator(QIntValidator(0, 1))  # Povolí len 0 alebo 1

        self.txtCashForStaffName = QLineEdit()
        self.txtCashForStaffName.setPlaceholderText("Staff name")

        right_expenses.addWidget(QLabel("CASH FOR STAFF:"))
        right_expenses.addWidget(self.txtCashForStaffName)
        right_expenses.addWidget(QLabel("EXPENSES CASH:"))
        right_expenses.addWidget(self.txtCashForStaff)

        # Pridanie ľavej a pravej časti do hlavného layoutu
        expenses_layout.addLayout(left_expenses)
        expenses_layout.addLayout(right_expenses)
        input_layout.addWidget(expenses_group)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        self.btnPridaj = QPushButton("Add")
        self.btnPridaj.setEnabled(True)
        self.btnPridaj.clicked.connect(self.btnPridaj_Click)
        
        self.btnEdit = QPushButton("Edit")
        self.btnEdit.setEnabled(False)
        self.btnEdit.clicked.connect(self.btnEdit_Click)
        
        self.btnDelete = QPushButton("Delete")
        self.btnDelete.setEnabled(False)
        self.btnDelete.clicked.connect(self.btnDelete_Click)
        
        buttons_layout.addWidget(self.btnPridaj)
        buttons_layout.addWidget(self.btnEdit)
        buttons_layout.addWidget(self.btnDelete)
        
        input_layout.addLayout(buttons_layout)
        left_layout.addWidget(input_group)
        content_layout.addWidget(left_panel)
        
        # Right panel - Results section
        right_panel = QFrame()
        right_panel.setFrameShape(QFrame.Shape.StyledPanel)
        right_layout = QVBoxLayout(right_panel)


        
        # Results section
        results_group = QGroupBox("RESULTS")
        results_layout = QVBoxLayout(results_group)
        
        # Income results
        income_results = QGroupBox("INCOME RESULTS")
        income_results_layout = QVBoxLayout(income_results)
        
        self.lblCashResult = QLabel("CASH: 0.00 £")
        self.lblCheckResult = QLabel("CHECK: 0.00 £")
        self.lblBankResult = QLabel("BANK TRANSFER: 0.00 £")
        self.lblTotalIncome = QLabel("TOTAL INCOME: 0.00 £")
        
        income_results_layout.addWidget(self.lblCashResult)
        income_results_layout.addWidget(self.lblCheckResult)
        income_results_layout.addWidget(self.lblBankResult)
        income_results_layout.addWidget(self.lblTotalIncome)
        results_layout.addWidget(income_results)
        
        # Expense results
        expense_results = QGroupBox("EXPENSE RESULTS")
        expense_results_layout = QVBoxLayout(expense_results)
        
        self.lblExpensesResult = QLabel("EXPENSES: 0.00 £")
        self.lblExpensesCashResult = QLabel("EXPENSES CASH: 0.00 £")
        self.lblTotalExpenses = QLabel("TOTAL EXPENSES: 0.00 £")
        self.lblStaffCost = QLabel("STAFF COST: 0.00 £")  # Nový QLabel pre StaffCost
        
        expense_results_layout.addWidget(self.lblExpensesResult)
        expense_results_layout.addWidget(self.lblExpensesCashResult)
        expense_results_layout.addWidget(self.lblTotalExpenses)
        results_layout.addWidget(expense_results)
        expense_results_layout.addWidget(self.lblStaffCost)  # Pridanie StaffCost 
        
        # Profit results
        profit_results = QGroupBox("PROFIT ANALYSIS")
        profit_results_layout = QVBoxLayout(profit_results)
        
        self.lblGrossProfit = QLabel("GROSS PROFIT: 0.00 £")
        self.lblNetProfit = QLabel("NET PROFIT (after cash expenses): 0.00 £")
        self.lblProfitMargin = QLabel("PROFIT MARGIN: 0%")
        
        # Nastavenie tooltipov pre každý riadok
        self.lblGrossProfit.setToolTip("Celkový zisk pred odpočítaním nákladov.")
        self.lblNetProfit.setToolTip("Čistý zisk po odpočítaní všetkých nákladov.")
        self.lblProfitMargin.setToolTip("Percentuálny podiel zisku na celkových príjmoch.")

        profit_results_layout.addWidget(self.lblGrossProfit)
        profit_results_layout.addWidget(self.lblNetProfit)
        profit_results_layout.addWidget(self.lblProfitMargin)
        results_layout.addWidget(profit_results)
        
        # Data table
        self.DGZoznam = QTableWidget()
        self.DGZoznam.setColumnCount(9)
        self.DGZoznam.setHorizontalHeaderLabels(["Date", "Client", "Cash", "Check", "Bank", "ExpensesDesc", "ExpCost","StaffCost", "StaffName"])
        self.DGZoznam.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.DGZoznam.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.DGZoznam.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.DGZoznam.cellClicked.connect(self.DGZoznam_CellClick)
        self.DGZoznam.cellDoubleClicked.connect(self.DGZoznam_CellContentDoubleClick)

        self.last_sorted_column = None  # Uloženie posledného zoradeného stĺpca
        self.sort_order = Qt.SortOrder.AscendingOrder  # Predvolený
        # Pripojenie eventu na kliknutie na hlavičku stĺpca
        header = self.DGZoznam.horizontalHeader()
        header.sectionClicked.connect(self.handle_header_click)

        # self.DGZoznam.horizontalHeader().sectionClicked.connect(self.handle_header_click)
        
        results_layout.addWidget(self.DGZoznam)
        # results_group.setLayout(results_layout)
        right_layout.addWidget(results_group)
        # right_panel.setLayout(right_layout)
        content_layout.addWidget(right_panel)
        
        self.main_layout.addWidget(content_widget)

        # Označíme, že tabuľka je inicializovaná
        self.table_initialized = True

        # Nastavte tučné písmo pre výsledkové QLabel widgety
        self.nastav_tucne_pismo_pre_vysledky_textboxe()

    def handle_header_click(self, column):
        """Spracovanie kliknutia na hlavičku stĺpca."""
        if self.last_sorted_column == column:
            # Ak je to ten istý stĺpec, zmeň smer zoradenia
            self.sort_order = Qt.SortOrder.DescendingOrder if self.sort_order == Qt.SortOrder.AscendingOrder else Qt.SortOrder.AscendingOrder
        else:
            # Ak je to nový stĺpec, začni s vzostupným zoradením
            self.sort_order = Qt.SortOrder.AscendingOrder

        # Nastavenie číselných hodnôt pre zoradenie
        for row in range(self.DGZoznam.rowCount()):
            item = self.DGZoznam.item(row, column)
            if item:
                try:
                    # Konverzia textu na float pre číselné zoradenie
                    item.setData(Qt.ItemDataRole.EditRole, float(item.text()))
                except ValueError:
                    pass  # Ignorujeme nečíselné hodnoty

        # Zoradenie tabuľky podľa stĺpca a smeru
        self.DGZoznam.sortItems(column, self.sort_order)
        
        # Aktualizácia posledného zoradeného stĺpca
        self.last_sorted_column = column        

    # Pridajte túto metódu do triedy FrmGarden
    def nastav_tucne_pismo_pre_vysledky_textboxe(self):
        """Nastaví tučné písmo pre všetky výsledkové QLabel widgety."""
        bold_font = QFont()
        bold_font.setBold(True)
        
        # Nastavte tučné písmo pre výsledkové QLabel widgety
        self.lblCashResult.setFont(bold_font)
        self.lblCheckResult.setFont(bold_font)
        self.lblBankResult.setFont(bold_font)
        self.lblTotalIncome.setFont(bold_font)
        
        self.lblExpensesResult.setFont(bold_font)
        self.lblExpensesCashResult.setFont(bold_font)
        self.lblTotalExpenses.setFont(bold_font)
        
        self.lblGrossProfit.setFont(bold_font)
        self.lblNetProfit.setFont(bold_font)
        self.lblProfitMargin.setFont(bold_font) 
        
        # Nastavte tučné písmo pre všetky textové polia
        self.txtSelectDate.setFont(bold_font)
        self.txtClients.setFont(bold_font)
        self.txtCash.setFont(bold_font)
        self.txtCheck.setFont(bold_font)
        self.txtBank.setFont(bold_font)
        self.txtExpenses.setFont(bold_font)
        self.txtExpensesCost.setFont(bold_font)
        self.txtCashForStaff.setFont(bold_font)
        self.txtCashForStaffName.setFont(bold_font)       

    def zmen_velkost_pisma(self, velkost, ulozit_nastavenie=True):
        """Zmení veľkosť písma pre celú aplikáciu"""
        font = QFont()
        font.setPointSize(velkost)
        
        # Aplikovanie na všetky widgety
        self.setFont(font)
        
        # Špeciálna úprava pre názov aby bol väčší
        title_font = QFont()
        title_font.setPointSize(velkost + 4)  # Názov je o 4 body väčší
        for child in self.findChildren(QLabel):
            if child.text() == NAZOV_APP:  # Nájdeme názov
                child.setFont(title_font)
                break
        
        # Uloženie nastavenia ak je potrebné
        if ulozit_nastavenie:
            self.settings.setValue("velkost_pisma", velkost) 
        
    def setup_database(self):
        self.FullPath = "DbGarden.db"
        self.ConnectionString = f"Data Source={self.FullPath};"
        
        # Získanie aktuálneho dátumu
        current_date = QDate.currentDate()
        year = current_date.year()
        month = current_date.month()

        # Logika pre výber správneho roku
        self.current_year = str(year if month >= 4 else year - 1)

        # Vytvorenie spojenia, ktoré budeme používať v celom programe
        self.conn = sqlite3.connect(self.FullPath) 
        
        # Nastavenie aktuálneho roku v comboboxe
        self.cbRok.setCurrentText(self.current_year)

        # Vytvorenie tabuľky, ak neexistuje
        if not self.check_table_exists(self.current_year):
            self.create_table_if_not_exists(self.current_year)        
            
    def check_table_exists(self, year):
        cursor = self.conn.cursor()
        
        # Query to check if table exists
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='TableGarden{year}';")
        result = cursor.fetchone()
        
        # If fetchone() returns None, table does not exist
        return result is not None
    
    def create_table_if_not_exists(self, year):
        cursor = self.conn.cursor()
        
        cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS TableGarden{year} (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Date TEXT,
            [CLIENTS NAME] TEXT,
            Cash REAL,
            [CHECK] REAL,                 
            [BANK TRANSFER] REAL,
            Expenses TEXT,
            [EXPENSES COSTS] REAL,
            CashForStaff REAL,               
            CashForStaffName TEXT
        )
        """)
        
        self.conn.commit()
    
    def load_initial_data(self):
        self.cbMonths.setCurrentIndex(0)
        self.NacitajDatabazu()
    
    # ============= EVENT HANDLERS ==============
    def btnDelete_Click(self):
        reply = QMessageBox.question(self, 'Delete Record', 
                                    f"Delete this record , nr. {self.Riadok} ?", 
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self.VymazatZaznam()
            self.NacitajDatabazu()
    
    def DGZoznam_CellClick(self, row, column):
        # if row != -1:
        #     self.DGZoznam.selectRow(row)
        #     self.OznacenieRiadku = True
        #     self.Riadok = int(self.DGZoznam.item(row, 0).text())
        #     self.btnDelete.setEnabled(True)
        # else:
        #     self.btnDelete.setEnabled(False)
        #     self.OznacenieRiadku = False
        None
    
    def DGZoznam_CellContentDoubleClick(self, row, column):
        self.Riadok = self.row_ids[row]
        if row != -1:
            # self.Riadok = int(self.DGZoznam.item(row, 0).text())
            self.MazanietxtPopridaniDoSql()

            self.txtSelectDate.setText(self.DGZoznam.item(row, 0).text())
            self.txtClients.setText(self.DGZoznam.item(row, 1).text())
            self.txtCash.setText(self.DGZoznam.item(row, 2).text())
            self.txtCheck.setText(self.DGZoznam.item(row, 3).text())
            self.txtBank.setText(self.DGZoznam.item(row, 4).text())
            self.txtExpenses.setText(self.DGZoznam.item(row, 5).text())
            self.txtExpensesCost.setText(self.DGZoznam.item(row, 6).text())
            
            if self.current_year >= '2025':
                self.txtCashForStaff.setText(self.DGZoznam.item(row, 7).text())
                self.txtCashForStaffName.setText(self.DGZoznam.item(row, 8).text())
            
            self.btnEdit.setEnabled(True)
            self.btnDelete.setEnabled(True)
    
    def Calendar_DateSelected(self):
        selected_date = self.Calendar.selectedDate()
        self.txtSelectDate.setText(selected_date.toString("yyyy-MM-dd"))
    


    def reinit_table(self):
        """Reinicializácia tabuľky, ak bola zničená"""
        # Nájdeme starú tabuľku a odstránime ju
        for i in reversed(range(self.main_layout.count())):
            widget = self.main_layout.itemAt(i).widget()
            if widget and widget.objectName() == "DGZoznam":
                widget.deleteLater()
        
        # Vytvoríme novú tabuľku
        self.DGZoznam = QTableWidget()
        self.DGZoznam.setObjectName("DGZoznam")
        self.DGZoznam.setColumnCount(9)
        self.DGZoznam.setHorizontalHeaderLabels(["Date", "Client", "Cash", "Check", "Bank", "ExpensesDesc", "ExpCost","StaffCash", "StaffName"])
        self.DGZoznam.cellClicked.connect(self.DGZoznam_CellClick)
        self.DGZoznam.cellDoubleClicked.connect(self.DGZoznam_CellContentDoubleClick)
        self.last_sorted_column = None  # Uloženie posledného zoradeného stĺpca
        self.sort_order = Qt.SortOrder.AscendingOrder  # Predvolený
        # Pripojenie eventu na kliknutie na hlavičku stĺpca
        header = self.DGZoznam.horizontalHeader()
        header.sectionClicked.connect(self.handle_header_click)
        # Nájdeme results_layout a pridáme tabuľku späť
        for i in range(self.main_layout.count()):
            widget = self.main_layout.itemAt(i).widget()
            if isinstance(widget, QGroupBox) and widget.title() == "RESULTS":
                results_layout = widget.layout()
                if results_layout:
                    results_layout.addWidget(self.DGZoznam)
                    break
        
        self.table_initialized = True                
    
    def txtSelectDate_TextChanged(self, text):
        self.btnPridaj.setEnabled(bool(text))
    
    def btnPridaj_Click(self):

        if self.txtClients.text() and self.txtSelectDate.text():
            reply = QMessageBox.question(self, 'Add Record', 
                                        f"Add this record?", 
                                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            
            if reply == QMessageBox.StandardButton.Yes:
                self.ZalozitZaznam()
                self.NacitajDatabazu()
        else:
            QMessageBox.information(self,"Error","Write at least the Date and the Client",QMessageBox.StandardButton.Ok)
    
    def btnEdit_Click(self):
        reply = QMessageBox.question(self, 'Edit Record', 
                                    "Save changes?", 
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self.UpdateZaznam()
            self.NacitajDatabazu()
            self.MazanietxtPopridaniDoSql()
    
    # ============= DATABASE OPERATIONS ==============
    def NacitajDatabazu(self):
        try:
            cursor = self.conn.cursor()
            cursor.execute(f"SELECT * FROM TableGarden{self.current_year}")
            rows = cursor.fetchall()
            
            self.DGZoznam.setRowCount(len(rows))

            self.row_ids = []  # Uložíme ID pre každý riadok
            
            for row_idx, row in enumerate(rows):
                self.row_ids.append(row[0])  # Prvé pole (Id) uložíme do zoznamu
                for col_idx, value in enumerate(row[1:]):  # Skip ID column
                    numeric_columns = [2, 3, 4, 6, 7]  # Indexy stĺpcov s číselnými hodnotami
                    if col_idx in numeric_columns:
                    # Pre číselné hodnoty použite Qt.ItemDataRole.EditRole
                        item = QTableWidgetItem()
                        item.setData(Qt.ItemDataRole.EditRole, float(value) if value else 0.0)
                    else:
                        # Pre textové hodnoty použite štandardný zápis
                        item = QTableWidgetItem(str(value))

                    self.DGZoznam.setItem(row_idx, col_idx, item)
            
            # conn.close()
            self.StatistikaVypocet()
        except Exception as ex:
            QMessageBox.critical(self +  'Metoda: NacitajDatabazu', "Error", str(ex))

    def load_unique_clients(self):
        """Načíta unikátne hodnoty klientov z databázy, ak nie sú v aktuálnom roku, použije predchádzajúci rok."""
        try:
            cursor = self.conn.cursor()
            
            # 1. Skúsi načítať klientov z aktuálneho roka
            cursor.execute(f"SELECT DISTINCT [CLIENTS NAME] FROM TableGarden{self.current_year};")
            rows_current = cursor.fetchall()
            self.unique_clients = [row[0] for row in rows_current if row[0]]
            
            # 2. Ak nie sú žiadni klienti v aktuálnom roku
            if not self.unique_clients:
                previous_year = str(int(self.current_year) - 1)
                
                # Skontroluje existenciu tabuľky pre predchádzajúci rok
                cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='TableGarden{previous_year}';")
                if cursor.fetchone():
                    # Načíta klientov z predchádzajúceho roka
                    cursor.execute(f"SELECT DISTINCT [CLIENTS NAME] FROM TableGarden{previous_year};")
                    rows_previous = cursor.fetchall()
                    self.unique_clients = [row[0] for row in rows_previous if row[0]]
            
            # 3. Nastavenie autocomplete
            completer = QCompleter(self.unique_clients, self)
            completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            self.txtClients.setCompleter(completer)
            
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Chyba pri načítaní klientov: {str(ex)}")


    def NacitajDatabazuPodlaMesiac(self, month_number,year):
        try:
            cursor = self.conn.cursor()
            # SQL dotaz s filtrom podľa mesiaca
            cursor.execute(f"""
            SELECT * FROM TableGarden{year}
            WHERE strftime('%m', Date) = ? 
            AND strftime('%Y', Date) = ?
            """, (f"{month_number:02d}", year))
            
            rows = cursor.fetchall()
            
            self.DGZoznam.setRowCount(len(rows))
            
            for row_idx, row in enumerate(rows):
                for col_idx, value in enumerate(row[1:]):  # Skip ID column
                    item = QTableWidgetItem(str(value))
                    self.DGZoznam.setItem(row_idx, col_idx, item)
            
            self.StatistikaVypocet()
        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))            
    
    def ZalozitZaznam(self):
        try:
            conn = sqlite3.connect(self.FullPath)
            cursor = conn.cursor()
            
            sql = f"""
            INSERT INTO TableGarden{self.cbRok.currentText()} 
            (Date, [CLIENTS NAME], Cash, [CHECK], [BANK TRANSFER], Expenses, [EXPENSES COSTS], CashForStaff, CashForStaffName)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            
            values = (
                self.txtSelectDate.text(),
                self.txtClients.text(),
                float(self.txtCash.text()) if self.txtCash.text() else 0,
                float(self.txtCheck.text()) if self.txtCheck.text() else 0,
                float(self.txtBank.text()) if self.txtBank.text() else 0,
                self.txtExpenses.text(),
                float(self.txtExpensesCost.text()) if self.txtExpensesCost.text() else 0,
                float(self.txtCashForStaff.text()) if self.txtCashForStaff.text() else 0,
                self.txtCashForStaffName.text()
            )
            
            cursor.execute(sql, values)
            conn.commit()
            conn.close()
            
            QMessageBox.information(self, "Success", "Record added")
            self.MazanietxtPopridaniDoSql()
        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))
    
    def UpdateZaznam(self):
        try:
            conn = sqlite3.connect(self.FullPath)
            cursor = conn.cursor()
            
            sql = f"""
            UPDATE TableGarden{self.cbRok.currentText()} 
            SET Date=?, [CLIENTS NAME]=?, Cash=?, [CHECK]=?, [BANK TRANSFER]=?, 
            Expenses=?, [EXPENSES COSTS]=?, CashForStaff=?, CashForStaffName=?
            WHERE Id=?
            """
            
            values = (
                self.txtSelectDate.text(),
                self.txtClients.text(),
                float(self.txtCash.text()) if self.txtCash.text() else 0,
                float(self.txtCheck.text()) if self.txtCheck.text() else 0,
                float(self.txtBank.text()) if self.txtBank.text() else 0,
                self.txtExpenses.text(),
                float(self.txtExpensesCost.text()) if self.txtExpensesCost.text() else 0,
                float(self.txtCashForStaff.text()) if self.txtCashForStaff.text() else 0,
                self.txtCashForStaffName.text(),
                self.Riadok
            )
            
            cursor.execute(sql, values)
            conn.commit()
            # conn.close()
            
            QMessageBox.information(self, "Success", "Record updated")
        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))
    
    def VymazatZaznam(self):
        try:
            conn = sqlite3.connect(self.FullPath)
            cursor = conn.cursor()
            
            sql = f"DELETE FROM TableGarden{self.cbRok.currentText()} WHERE Id=?"
            cursor.execute(sql, (self.Riadok,))
            
            conn.commit()
            # conn.close()
            
            QMessageBox.information(self, "Success", "Record deleted")
            self.MazanietxtPopridaniDoSql()

        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))
    
    def MazanietxtPopridaniDoSql(self):
        self.txtSelectDate.clear()
        self.txtClients.clear()
        self.txtCash.clear()
        self.txtCheck.clear()
        self.txtBank.clear()
        self.txtExpenses.clear()
        self.txtExpensesCost.clear()
        self.txtCashForStaff.clear()
        self.txtCashForStaffName.clear()
        self.btnDelete.setEnabled(False)
        self.btnEdit.setEnabled(False)
        self.btnPridaj.setEnabled(True)
    
    def StatistikaVypocet(self):
        try:
                
            total_cash = 0
            total_check = 0
            total_bank = 0
            total_expenses = 0
            total_cash_expenses = 0
            staff_cost = 0  # Inicializácia StaffCost
            
            for row in range(self.DGZoznam.rowCount()):
                # Income
                cash_item = self.DGZoznam.item(row, 2)
                if cash_item and cash_item.text():
                    total_cash += float(cash_item.text())
                
                check_item = self.DGZoznam.item(row, 3)
                if check_item and check_item.text():
                    total_check += float(check_item.text())
                
                bank_item = self.DGZoznam.item(row, 4)
                if bank_item and bank_item.text():
                    total_bank += float(bank_item.text())
                
                # Expenses
                cost_item = self.DGZoznam.item(row, 6)
                if cost_item and cost_item.text():
                    total_expenses += float(cost_item.text())
                    
                    # # Check if this was a cash expense
                    # cash_exp_item = self.DGZoznam.item(row, 7)
                    # if cash_exp_item and cash_exp_item.text() == "1":
                    #     total_cash_expenses += float(cost_item.text())

                    # Staff Cost
                staff_cost_item = self.DGZoznam.item(row, 7)  # Stĺpec CashForStaff
                if staff_cost_item and staff_cost_item.text():
                    staff_cost += float(staff_cost_item.text())                    
            
            # Calculate totals
            total_income = total_cash + total_check + total_bank
            gross_profit = total_income - total_expenses
            net_profit = total_income - total_expenses  # Same as gross in this simple model
            profit_margin = (gross_profit / (total_cash + total_check + total_bank) * 100) if (total_cash + total_check + total_bank) > 0 else 0
            
            # Update UI with bold font for numeric values
            formatted_cash = self.format_number_with_spaces(total_cash)
            self.update_label_with_bold(self.lblCashResult, f"CASH: <b>{formatted_cash} £</b>")

            formatted_check = self.format_number_with_spaces(total_check)
            self.update_label_with_bold(self.lblCheckResult, f"CHECK: <b>{formatted_check} £</b>")

            formatted_bank = self.format_number_with_spaces(total_bank)
            self.update_label_with_bold(self.lblBankResult, f"BANK TRANSFER: <b>{formatted_bank} £</b>")

            formatted_income = self.format_number_with_spaces(total_income)
            self.update_label_with_bold(self.lblTotalIncome, f"TOTAL INCOME: <b>{formatted_income} £</b>")

            formatted_expenses = self.format_number_with_spaces(total_expenses)
            self.update_label_with_bold(self.lblExpensesResult, f"EXPENSES: <b>{formatted_expenses} £</b>")

            formatted_cash_expenses = self.format_number_with_spaces(total_cash_expenses)
            self.update_label_with_bold(self.lblExpensesCashResult, f"<s>EXPENSES CASH: <b>{formatted_cash_expenses} £</b></s>")

            formatted_staff_cost = self.format_number_with_spaces(staff_cost)
            self.lblStaffCost.setText(f"STAFF COST: <b>{formatted_staff_cost} £</b>")

            formatted_gross_profit = self.format_number_with_spaces(gross_profit)
            self.update_label_with_bold(self.lblGrossProfit, f"GROSS PROFIT: <b>{formatted_gross_profit} £</b>")

            formatted_net_profit = self.format_number_with_spaces(net_profit)
            self.update_label_with_bold(self.lblNetProfit, f"NET PROFIT: <b>{formatted_net_profit} £</b>")

            formatted_profit_margin = self.format_number_with_spaces(profit_margin)
            self.update_label_with_bold(self.lblProfitMargin, f"PROFIT MARGIN: <b>{formatted_profit_margin}%</b>")
            
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Chyba pri Statistike: {str(ex)}")
            
    def format_number_with_spaces(self,number):
        """Formátuje číslo s medzerami medzi tisíckami."""
        if number >= 1000 or number <= -1000:  # Kontrola, či je číslo tisíckové alebo vyššie
            return f"{number:,.2f}".replace(",", " ")
        else:
            return f"{number:.2f}"
        
    def update_label_with_bold(self, label, text):
        """Aktualizuje QLabel s tučným písmom iba pre číselné hodnoty."""
        parts = text.split(":") # Rozdelíme text na časti pred a po dvojbodke.

        if len(parts) == 2:
            bold_font = QFont()
            bold_font.setBold(True)

            normal_font = QFont()
            normal_font.setBold(False)

            label.setText(f"{parts[0]}: <b>{parts[1].strip()}</b>")
            label.setFont(normal_font) # Nastavíme normálne písmo pre celý text.

    def cbMonths_SelectedIndexChanged(self, index):
        try:
            selected_month = self.cbMonths.currentText()

            # Aktualizácia názvu skupiny výsledkov
            if selected_month == "All":
                self.results_group.setTitle("RESULTS")
            else:
                self.results_group.setTitle(f"RESULTS FOR: {selected_month}")
            
            if selected_month != "All":
                try:
                    #year = int(self.cbRok.currentText())
                    month_index = self.cbMonths.currentIndex()
                    
                    # Ak je vybraný konkrétny mesiac (nie "All")
                    if month_index > 0:  # "All" je index 0
                        month_number = month_index  # január = 1, február = 2, atď.
                        selected_year = self.cbRok.currentText()
                        # Načítanie dát pre konkrétny mesiac
                        self.NacitajDatabazuPodlaMesiac(month_number,selected_year)
                except Exception as ex:
                    QMessageBox.critical(self, "Error", str(ex))
            else:
                # Ak je vybrané "All", načítaj všetky dáta
                self.NacitajDatabazu()
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Chyba pri zmene mesiaca: {str(ex)}")

    def cbRok_SelectedIndexChanged(self, index):
      if hasattr(self, 'DGZoznam') and self.table_initialized:
        try:   
            # Nastavíme mesiac na "All"
            self.cbMonths.setCurrentIndex(0)

            # Aktualizujeme current_year podľa výberu v comboboxe
            self.current_year = self.cbRok.currentText()

            # Skontrolujeme existenciu tabuľky pre nový rok
            if not self.check_table_exists(self.current_year):
                self.create_table_if_not_exists(self.current_year)

            self.DGZoznam.clearContents()
            self.DGZoznam.setRowCount(0)
            self.NacitajDatabazu()
            
            # Aktualizujeme názov výsledkovej skupiny
            self.results_group.setTitle(f"RESULTS FOR YEAR: {self.current_year}") 
            # Aktualizácia autocomplete pre Client
            self.load_unique_clients()
        except RuntimeError:
            # Obnovíme tabuľku, ak bola zničená
            self.reinit_table()
            self.NacitajDatabazu()            

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FrmGarden()
    window.show()
    sys.exit(app.exec())