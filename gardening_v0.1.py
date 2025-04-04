import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QPushButton, QComboBox, QTableWidget, QTableWidgetItem, 
                            QCalendarWidget, QCheckBox, QGroupBox, QMessageBox, QFileDialog, QFrame)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QFont
import sqlite3
import openpyxl
from openpyxl import Workbook

class FrmGarden(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Gardening")
        self.setGeometry(100, 100, 1000, 700)
        
        self.init_ui()
        self.setup_database()
        self.load_initial_data()
        
    def init_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        
        # Create a frame for the header
        header_frame = QFrame()
        header_frame.setFrameShape(QFrame.Shape.StyledPanel)
        header_layout = QHBoxLayout()
        
        # Add title label
        title_label = QLabel("GARDENING")
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
        
        # self.setup_initial_data()
        
        header_frame.setLayout(header_layout)
        main_layout.addWidget(header_frame)
        
        # Create main content area
        content_widget = QWidget()
        content_layout = QHBoxLayout()
        
        # Left panel - Input section
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.Shape.StyledPanel)
        left_layout = QVBoxLayout()
        
        # Calendar
        self.Calendar = QCalendarWidget()
        self.Calendar.setGridVisible(True)
        self.Calendar.clicked.connect(self.Calendar_DateSelected)
        left_layout.addWidget(self.Calendar)
        
        # Input fields
        input_group = QGroupBox("Transaction Details")
        input_layout = QVBoxLayout()
        
        self.txtSelectDate = QLineEdit()
        self.txtSelectDate.setPlaceholderText("YYYY-MM-DD")
        input_layout.addWidget(QLabel("Date:"))
        input_layout.addWidget(self.txtSelectDate)
        
        self.txtClients = QLineEdit()
        self.txtClients.setPlaceholderText("Client Name")
        input_layout.addWidget(QLabel("Client:"))
        input_layout.addWidget(self.txtClients)
        
        # Income section
        income_group = QGroupBox("INCOME")
        income_layout = QVBoxLayout()
        
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
        
        income_group.setLayout(income_layout)
        input_layout.addWidget(income_group)
        
        # Expenses section
        expenses_group = QGroupBox("EXPENSE")
        expenses_layout = QVBoxLayout()
        
        self.txtExpenses = QLineEdit()
        self.txtExpenses.setPlaceholderText("Expenses Description")
        expenses_layout.addWidget(QLabel("EXPENSES:"))
        expenses_layout.addWidget(self.txtExpenses)
        
        self.txtExpensesCost = QLineEdit()
        self.txtExpensesCost.setPlaceholderText("Cost")
        expenses_layout.addWidget(QLabel("COST:"))
        expenses_layout.addWidget(self.txtExpensesCost)
        
        self.chckExpesiveCash = QCheckBox("EXPENSES CASH")
        expenses_layout.addWidget(self.chckExpesiveCash)
        
        expenses_group.setLayout(expenses_layout)
        input_layout.addWidget(expenses_group)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        self.btnPridaj = QPushButton("Add")
        self.btnPridaj.setEnabled(False)
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
        input_group.setLayout(input_layout)
        left_layout.addWidget(input_group)
        left_panel.setLayout(left_layout)
        content_layout.addWidget(left_panel)
        
        # Right panel - Results section
        right_panel = QFrame()
        right_panel.setFrameShape(QFrame.Shape.StyledPanel)
        right_layout = QVBoxLayout()
        
        # Results section
        results_group = QGroupBox("RESULTS")
        results_layout = QVBoxLayout()
        
        # Income results
        income_results = QGroupBox("INCOME RESULTS")
        income_results_layout = QVBoxLayout()
        
        self.lblCashResult = QLabel("CASH: 0.00 £")
        self.lblCheckResult = QLabel("CHECK: 0.00 £")
        self.lblBankResult = QLabel("BANK TRANSFER: 0.00 £")
        self.lblTotalIncome = QLabel("TOTAL INCOME: 0.00 £")
        
        income_results_layout.addWidget(self.lblCashResult)
        income_results_layout.addWidget(self.lblCheckResult)
        income_results_layout.addWidget(self.lblBankResult)
        income_results_layout.addWidget(self.lblTotalIncome)
        income_results.setLayout(income_results_layout)
        results_layout.addWidget(income_results)
        
        # Expense results
        expense_results = QGroupBox("EXPENSE RESULTS")
        expense_results_layout = QVBoxLayout()
        
        self.lblExpensesResult = QLabel("EXPENSES: 0.00 £")
        self.lblExpensesCashResult = QLabel("EXPENSES CASH: 0.00 £")
        self.lblTotalExpenses = QLabel("TOTAL EXPENSES: 0.00 £")
        
        expense_results_layout.addWidget(self.lblExpensesResult)
        expense_results_layout.addWidget(self.lblExpensesCashResult)
        expense_results_layout.addWidget(self.lblTotalExpenses)
        expense_results.setLayout(expense_results_layout)
        results_layout.addWidget(expense_results)
        
        # Profit results
        profit_results = QGroupBox("PROFIT ANALYSIS")
        profit_results_layout = QVBoxLayout()
        
        self.lblGrossProfit = QLabel("GROSS PROFIT: 0.00 £")
        self.lblNetProfit = QLabel("NET PROFIT (after cash expenses): 0.00 £")
        self.lblProfitMargin = QLabel("PROFIT MARGIN: 0%")
        
        profit_results_layout.addWidget(self.lblGrossProfit)
        profit_results_layout.addWidget(self.lblNetProfit)
        profit_results_layout.addWidget(self.lblProfitMargin)
        profit_results.setLayout(profit_results_layout)
        results_layout.addWidget(profit_results)
        
        # Data table
        self.DGZoznam = QTableWidget()
        self.DGZoznam.setColumnCount(8)
        self.DGZoznam.setHorizontalHeaderLabels(["Date", "Client", "Cash", "Check", "Bank", "Expenses", "Cost", "Cash Exp"])
        self.DGZoznam.cellClicked.connect(self.DGZoznam_CellClick)
        self.DGZoznam.cellDoubleClicked.connect(self.DGZoznam_CellContentDoubleClick)
        
        results_layout.addWidget(self.DGZoznam)
        results_group.setLayout(results_layout)
        right_layout.addWidget(results_group)
        right_panel.setLayout(right_layout)
        content_layout.addWidget(right_panel)
        
        content_widget.setLayout(content_layout)
        main_layout.addWidget(content_widget)
        
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
    # def setup_initial_data(self):
    #     # Determine correct year based on current month
    #     current_date = QDate.currentDate()
    #     year = current_date.year()
    #     month = current_date.month()

    #     correct_year = year if month > 4 else year - 1

    #     # Set the correct year and load database
    #     self.cbRok.setCurrentText(str(correct_year))
    #     self.NacitajDatabazu()    
        
    def setup_database(self):
        self.FullPath = "DbGarden.db"
        self.ConnectionString = f"Data Source={self.FullPath};"
        
        # Create table if not exists for current year
        current_year = self.cbRok.currentText()
        if not self.check_table_exists(current_year):
            self.create_table_if_not_exists(current_year)
            
    def check_table_exists(self, year):
        conn = sqlite3.connect(self.FullPath)
        cursor = conn.cursor()
        
        # Query to check if table exists
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='TableGarden{year}';")
        result = cursor.fetchone()
        
        conn.close()

        # If fetchone() returns None, table does not exist
        return result is not None
    
    def create_table_if_not_exists(self, year):
        conn = sqlite3.connect(self.FullPath)
        cursor = conn.cursor()
        
        cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS TableGarden{year} (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Date TEXT,
            Client TEXT,
            Cash REAL,
            CheckAmount REAL,
            BankTransfer REAL,
            Expenses TEXT,
            Cost REAL,
            CashExpense INTEGER
        )
        """)
        
        conn.commit()
        conn.close()
    
    def load_initial_data(self):
        current_date = QDate.currentDate()
        month = current_date.month()
        year = current_date.year()
        
        if month >= 4:
            self.cbRok.setCurrentText(str(year))
        else:
            self.cbRok.setCurrentText(str(year - 1))
        
        self.cbMonths.setCurrentIndex(0)
        self.NacitajDatabazu()
    
    # ============= EVENT HANDLERS ==============
    def btnDelete_Click(self):
        reply = QMessageBox.question(self, 'Delete Record', 
                                    "Delete this record?", 
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes and self.OznacenieRiadku:
            self.VymazatZaznam()
            self.NacitajDatabazu()
    
    def DGZoznam_CellClick(self, row, column):
        if row != -1:
            self.DGZoznam.selectRow(row)
            self.OznacenieRiadku = True
            self.Riadok = int(self.DGZoznam.item(row, 0).text())
            self.btnDelete.setEnabled(True)
        else:
            self.btnDelete.setEnabled(False)
            self.OznacenieRiadku = False
    
    def DGZoznam_CellContentDoubleClick(self, row, column):
        if row != -1:
            self.Riadok = int(self.DGZoznam.item(row, 0).text())
            self.MazanietxtPopridaniDoSql()
            
            self.txtSelectDate.setText(self.DGZoznam.item(row, 1).text())
            self.txtClients.setText(self.DGZoznam.item(row, 2).text())
            self.txtCash.setText(self.DGZoznam.item(row, 3).text())
            self.txtCheck.setText(self.DGZoznam.item(row, 4).text())
            self.txtBank.setText(self.DGZoznam.item(row, 5).text())
            self.txtExpenses.setText(self.DGZoznam.item(row, 6).text())
            self.txtExpensesCost.setText(self.DGZoznam.item(row, 7).text())
            self.chckExpesiveCash.setChecked(self.DGZoznam.item(row, 8).text() == "1")
            
            self.btnEdit.setEnabled(True)
    
    def Calendar_DateSelected(self):
        selected_date = self.Calendar.selectedDate()
        self.txtSelectDate.setText(selected_date.toString("yyyy-MM-dd"))
    
    def cbRok_SelectedIndexChanged(self, index):
        self.DGZoznam.clearContents()
        self.DGZoznam.setRowCount(0)
        self.NacitajDatabazu()
    
    def txtSelectDate_TextChanged(self, text):
        self.btnPridaj.setEnabled(bool(text))
    
    def btnPridaj_Click(self):
        reply = QMessageBox.question(self, 'Add Record', 
                                    f"Add this record?", 
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self.ZalozitZaznam()
            self.NacitajDatabazu()
    
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
            conn = sqlite3.connect(self.FullPath)
            cursor = conn.cursor()
            
            cursor.execute(f"SELECT * FROM TableGarden{self.cbRok.currentText()}")
            rows = cursor.fetchall()
            
            self.DGZoznam.setRowCount(len(rows))
            
            for row_idx, row in enumerate(rows):
                for col_idx, value in enumerate(row[1:]):  # Skip ID column
                    item = QTableWidgetItem(str(value))
                    self.DGZoznam.setItem(row_idx, col_idx, item)
            
            conn.close()
            self.StatistikaVypocet()
        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))
    
    def ZalozitZaznam(self):
        try:
            conn = sqlite3.connect(self.FullPath)
            cursor = conn.cursor()
            
            sql = f"""
            INSERT INTO TableGarden{self.cbRok.currentText()} 
            (Date, Client, Cash, CheckAmount, BankTransfer, Expenses, Cost, CashExpense) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """
            
            values = (
                self.txtSelectDate.text(),
                self.txtClients.text(),
                float(self.txtCash.text()) if self.txtCash.text() else 0,
                float(self.txtCheck.text()) if self.txtCheck.text() else 0,
                float(self.txtBank.text()) if self.txtBank.text() else 0,
                self.txtExpenses.text(),
                float(self.txtExpensesCost.text()) if self.txtExpensesCost.text() else 0,
                1 if self.chckExpesiveCash.isChecked() else 0
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
            SET Date=?, Client=?, Cash=?, CheckAmount=?, BankTransfer=?, 
            Expenses=?, Cost=?, CashExpense=?
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
                1 if self.chckExpesiveCash.isChecked() else 0,
                self.Riadok
            )
            
            cursor.execute(sql, values)
            conn.commit()
            conn.close()
            
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
            conn.close()
            
            QMessageBox.information(self, "Success", "Record deleted")
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
        self.chckExpesiveCash.setChecked(False)
        self.btnDelete.setEnabled(False)
        self.btnEdit.setEnabled(False)
        self.btnPridaj.setEnabled(False)
    
    def StatistikaVypocet(self):
        total_cash = 0
        total_check = 0
        total_bank = 0
        total_expenses = 0
        total_cash_expenses = 0
        
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
                
                # Check if this was a cash expense
                cash_exp_item = self.DGZoznam.item(row, 7)
                if cash_exp_item and cash_exp_item.text() == "1":
                    total_cash_expenses += float(cost_item.text())
        
        # Calculate totals
        total_income = total_cash + total_check + total_bank
        gross_profit = total_income - total_expenses
        net_profit = total_income - total_expenses  # Same as gross in this simple model
        profit_margin = (gross_profit / total_income * 100) if total_income > 0 else 0
        
        # Update UI
        self.lblCashResult.setText(f"CASH: {total_cash:.2f}")
        self.lblCheckResult.setText(f"CHECK: {total_check:.2f}")
        self.lblBankResult.setText(f"BANK TRANSFER: {total_bank:.2f}")
        self.lblTotalIncome.setText(f"TOTAL INCOME: {total_income:.2f}")
        
        self.lblExpensesResult.setText(f"EXPENSES: {total_expenses:.2f}")
        self.lblExpensesCashResult.setText(f"EXPENSES CASH: {total_cash_expenses:.2f}")
        self.lblTotalExpenses.setText(f"TOTAL EXPENSES: {total_expenses:.2f}")
        
        self.lblGrossProfit.setText(f"GROSS PROFIT: {gross_profit:.2f}")
        self.lblNetProfit.setText(f"NET PROFIT: {net_profit:.2f}")
        self.lblProfitMargin.setText(f"PROFIT MARGIN: {profit_margin:.1f}%")
        
    def cbMonths_SelectedIndexChanged(self, index):
        if not self.PrveOtvorenieMesiace:
            self.gbResult.setTitle(f"RESULT OF MONTH: {self.cbMonths.currentText()}")
            
            if self.cbMonths.currentText() != "All":
                try:
                    yYear = int(self.cbRok.currentText())
                    mMonth = self.Mesiace(self.cbMonths.currentText())
                    
                    # TODO: Implement filtering logic here
                    # This would need to be adapted to work with QTableWidget
                    
                    self.StatistikaVypocet()
                except Exception as ex:
                    QMessageBox.critical(self, "Error", str(ex))
            else:
                self.NacitajDatabazu()
                self.StatistikaVypocet()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FrmGarden()
    window.show()
    sys.exit(app.exec())