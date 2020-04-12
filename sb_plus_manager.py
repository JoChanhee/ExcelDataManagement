__author__ = "Chanhee Jo"
__copyright__ = "Copyright 2020, SB Plus Co.,Ltd"
__license__ = "GPL"
__version__ = "1.0.0"
__email__ = "teletovy@gmail.com, decision_1@naver.com"

import sys
import os
import openpyxl
import locale
import datetime

from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtCore import QStringListModel
from PyQt5.QtCore import Qt
from PyQt5.QtChart import QChart, QChartView, QBarSet, QBarSeries, QPercentBarSeries, QBarCategoryAxis
from PyQt5.QtGui import QPainter

from model.data_type import DataType, SupplierType, SheetType, ItemAttribute
from statistics_manager import StatisticsManager, REMAIN_COUNT

locale.setlocale(locale.LC_ALL, '')

# Data path
UI_PATH = "sb_plus_ui.ui"
DATA_PATH = "sb_plus_management_2020.xlsx"
BACKUP_DATA_PATH = "sb_plus_management_2020_backup.xlsx"

MPN_IDX = 0 #TODO::Need to get index from header

#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType(UI_PATH)[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        # Data
        self.excel_document = openpyxl.load_workbook(os.path.join(os.getcwd(), DATA_PATH))
        self.excel_document.save(BACKUP_DATA_PATH)

        self.management_sheet, self.management_header, self.management_contents = self.init_data_with_sheet_type(self.tableWidget, SheetType.MANAGEMENT)
        self.supplier_sheet, self.supplier_header, self.supplier_contents = self.init_data_with_sheet_type(
            self.supplierTableWidget, SheetType.SUPPLIER)
        self.item_sheet, self.item_header, self.item_contents = self.init_data_with_sheet_type(self.itemTableWidget, SheetType.ITEM)


        # Manager
        self.statistics_manager = StatisticsManager(header=self.management_header, contents=self.management_contents)

        self.init_data_with_dict(self.statisticsTableWidget, dict=self.statistics_manager.get_item_statistics_dict())

        # UI
        self.make_logo()
        self.make_comboBox()
        self.make_mpn_completer()
        self.make_bar_chart()

        # Event
        self.inputButton.clicked.connect(self.onClickInputButton)
        self.outputButton.clicked.connect(self.onClickOutputButton)
        self.supplierButton.clicked.connect(self.onClickSupplierButton)
        self.itemButton.clicked.connect(self.onClickItemButton)
        self.searchButton.clicked.connect(self.onClickSearchButton)

        self.editMpn.returnPressed.connect(self.onEnterMpnEdit)
        self.editMpn_2.returnPressed.connect(self.onEnterMpnEdit_2)
        self.searchMpn.returnPressed.connect(self.onClickSearchButton)



# Data management widget initialization
    def init_data_with_sheet_type(self, widget, eSheetType):
        sheet = self.excel_document.get_sheet_by_name(eSheetType.value)

        data = sheet.rows
        headers = []
        contents = []

        mpn_idx = -1

        for y, rows in enumerate(data):
            content = []
            for x, col in enumerate(rows):
                if y == 0:
                    headers.append(col.value)
                    if col.value == ItemAttribute.MPN.value:
                        mpn_idx = x
                else:
                    if x == mpn_idx:
                        content.append(str(col.value))
                    else:
                        content.append(col.value)

            if len(content) != 0:
                contents.append(content)

        widget.setColumnCount(len(headers))
        widget.setRowCount(len(contents))
        widget.setHorizontalHeaderLabels(headers)

        widget.resizeRowsToContents()

        for row_idx, content in enumerate(contents):
            for col_idx, item in enumerate(content):
                if col_idx != MPN_IDX and self.is_number(item):   # col_idx = mpn_idx
                    item = locale.format("%2.f", float(item), 1)
                else:
                    item = str(item)

                widget.setItem(row_idx, col_idx, QTableWidgetItem(item))

        return sheet, headers, contents

    def init_data_with_dict(self, widget, dict):

        headers = []
        contents = []
        is_initialized_header = False

        for key, value in dict.items():
            content = []

            if not is_initialized_header:
                headers.append(ItemAttribute.MPN.value)
            content.append(str(key))

            for value_key, value_value in value.items():
                if not is_initialized_header:
                    headers.append(str(value_key))
                content.append(value_value)

            is_initialized_header = True

            if len(content) != 0:
                contents.append(content)

        widget.setColumnCount(len(headers))
        widget.setRowCount(len(contents))
        widget.setHorizontalHeaderLabels(headers)

        widget.resizeRowsToContents()

        for row_idx, content in enumerate(contents):
            for col_idx, item in enumerate(content):
                if col_idx != MPN_IDX and self.is_number(item):   # col_idx = mpn_idx
                    item = locale.format("%2.f", float(item), 1)
                else:
                    item = str(item)

                widget.setItem(row_idx, col_idx, QTableWidgetItem(item))


# Management funcitons
    def add_content_into_management_table(self, content):
        last_row_idx = len(self.management_contents)
        self.tableWidget.setRowCount(last_row_idx + 1)

        self.management_contents.append(content)
        self.management_sheet.append(content)

        for col_idx, item in enumerate(content):
            if col_idx != MPN_IDX and self.is_number(item):  # col_idx = mpn_idx
                item = locale.format("%2.f", float(item), 1)

            self.tableWidget.setItem(last_row_idx, col_idx, QTableWidgetItem(item))

        self.excel_document.save(DATA_PATH)

        self.onChangedManagement()
        self.onChangedSupplier()

    def add_content_into_supplier_table(self, content): # mpn is not exist in supplier table
        last_row_idx = len(self.supplier_contents)
        self.supplierTableWidget.setRowCount(last_row_idx + 1)

        self.supplier_contents.append(content)
        self.supplier_sheet.append(content)

        for col_idx, item in enumerate(content):
            self.supplierTableWidget.setItem(last_row_idx, col_idx, QTableWidgetItem(item))

        self.excel_document.save(DATA_PATH)

        self.onChangedSupplier()

    def add_content_into_item_table(self, content):
        last_row_idx = len(self.item_contents)
        self.itemTableWidget.setRowCount(last_row_idx + 1)

        self.item_contents.append(content)
        self.item_sheet.append(content)

        for col_idx, item in enumerate(content):
            if col_idx != MPN_IDX and self.is_number(item):  # col_idx = mpn_idx
                item = locale.format("%2.f", float(item), 1)

            self.itemTableWidget.setItem(last_row_idx, col_idx, QTableWidgetItem(item))

        self.excel_document.save(DATA_PATH)

        self.onChangedItem()

    def add_content_into_search_tables(self, input_contents, output_contents):
        self.searchTableWidget.setColumnCount(len(self.management_header))
        self.searchTableWidget.setRowCount(len(input_contents))
        self.searchTableWidget.setHorizontalHeaderLabels(self.management_header)
        self.searchTableWidget.resizeColumnsToContents()
        self.searchTableWidget.horizontalHeader().setStretchLastSection(True)

        self.searchTableWidget_2.setColumnCount(len(self.management_header))
        self.searchTableWidget_2.setRowCount(len(output_contents))
        self.searchTableWidget_2.setHorizontalHeaderLabels(self.management_header)
        self.searchTableWidget_2.resizeColumnsToContents()
        self.searchTableWidget_2.horizontalHeader().setStretchLastSection(True)

        for row_idx, content in enumerate(input_contents):
            for col_idx, item in enumerate(content):
                if col_idx != MPN_IDX and self.is_number(item):  # col_idx = mpn_idx
                    item = locale.format("%2.f", float(item), 1)

                self.searchTableWidget.setItem(row_idx, col_idx, QTableWidgetItem(item))

        for row_idx, content in enumerate(output_contents):
            for col_idx, item in enumerate(content):
                if col_idx != MPN_IDX and self.is_number(item):  # col_idx = mpn_idx
                    item = locale.format("%2.f", float(item), 1)

                self.searchTableWidget_2.setItem(row_idx, col_idx, QTableWidgetItem(item))


# Check functions
    def is_empty_box(self, eDataType):
        editBoxes = None
        if eDataType == DataType.INPUT:
            editBoxes = [self.editMpn, self.editSupplier, self.editPrice, self.editCount]
        else:
            editBoxes = [self.editMpn_2, self.editSupplier_2, self.editPrice_2, self.editCount_2]

        for editBox in editBoxes:
            targetText = ""
            if editBox.__class__==QComboBox:
                targetText = editBox.currentText()
            else:
                targetText = editBox.text()

            if targetText == "":
                return True

        return False

    def is_empty_item_box(self):
        editBoxes = [self.editMpn_3, self.editSupplier_3, self.editPrice_3]

        for editBox in editBoxes:
            targetText = ""
            if editBox.__class__ == QComboBox:
                targetText = editBox.currentText()
            else:
                targetText = editBox.text()

            if targetText == "":
                return True

        return False

    def is_number(self, number):
        if number is None:
            return False
        if isinstance(number, datetime.datetime):
            return False
        try:
            float(number)
            return True
        except ValueError:
            return False

    def is_mpn_edit_box(self, editBox):
        if 'editMpn' in editBox.objectName():
            return True

        return False

    def is_empty_supplier_radio(self):
        bSupplier = self.supplierRadio.isChecked()
        bCustomer = self.customerRadio.isChecked()

        if not bSupplier and not bCustomer:
            return True

        return False

    def is_already_existed_item(self, targetContent):
        for content in self.item_contents:
            if content[0] == targetContent[0]:
                return True

        return False

# Get functions
    def get_edit_box_data(self, eDataType):
        if eDataType == DataType.INPUT:
            editBoxes = [self.editMpn, self.editPartName, self.editSupplier, self.editDate, self.editPrice, self.editCount]
        else:
            editBoxes = [self.editMpn_2, self.editPartName_2, self.editSupplier_2, self.editDate_2, self.editPrice_2, self.editCount_2]

        content = []
        for editBox in editBoxes:
            targetText = ""

            # get text with box type
            if editBox.__class__ == QComboBox:
                targetText = editBox.currentText()
            else:
                targetText = editBox.text()

            # set value type with info ( mpn? / number? )
            if self.is_mpn_edit_box(editBox):
                targetText = str(targetText)
            elif self.is_number(targetText):
                targetText = float(targetText)

            content.append(targetText)

        for dataTypeIdx, item in enumerate(self.management_header):
            if item == ItemAttribute.INPUT_OUTPUT.value:
                content.insert(dataTypeIdx, eDataType.value)

        return content

    def get_supplier_box_data(self):
        content = []

        bCustomer = self.customerRadio.isChecked()

        eSupplierType = SupplierType.SUPPLIER
        if bCustomer:
            eSupplierType = SupplierType.CUSTOMER

        content.append(self.addSupplier.text())
        content.append(eSupplierType.value)

        return content

    def get_item_box_data(self):
        editBoxes = [self.editMpn_3, self.editPartName_3, self.editSupplier_3, self.editPrice_3]

        content = []
        for editBox in editBoxes:
            targetText = ""
            if editBox.__class__ == QComboBox:
                targetText = editBox.currentText()
            else:
                targetText = editBox.text()

                if self.is_mpn_edit_box(editBox):
                    targetText = str(targetText)
                elif self.is_number(targetText):
                    targetText = float(targetText)

            content.append(targetText)

        return content

    def get_supplier_list_with_type(self, eSupplierType):
        supplier_list = []
        for supplier in self.supplier_contents:
            if supplier[1] == eSupplierType.value:
                supplier_list.append(supplier[0])

        return supplier_list

    def get_item_list_with_mpn(self, mpn, eDataType):
        mpn_idx = 0
        data_type_idx = 2
        for idx, attribute in enumerate(self.management_header):
            if attribute == ItemAttribute.INPUT_OUTPUT.value:
                data_type_idx = idx
            if attribute == ItemAttribute.MPN.value:
                mpn_idx = idx


        item_list = []
        for item in self.management_contents:
            if item[mpn_idx] == mpn and item[data_type_idx] == eDataType.value:
                item_list.append(item)

        return item_list

    def get_item_info_with_mpn(self, mpn):
        for item in self.item_contents:
            if item[0] == mpn:
                return item


# Callback functions
    def onClickInputButton(self):
        if self.is_empty_box(DataType.INPUT):
            print("please input data into empty slots.")
            return

        print("onClickInputButton")
        self.add_content_into_management_table(self.get_edit_box_data(DataType.INPUT))

    def onClickOutputButton(self):
        if self.is_empty_box(DataType.OUTPUT):
            print("please input data into empty slots.")
            return

        print("onClickOutputButton")
        self.add_content_into_management_table(self.get_edit_box_data(DataType.OUTPUT))

    def onClickSupplierButton(self):
        if self.is_empty_supplier_radio():
            print("Select supplier or customer")
            return

        self.add_content_into_supplier_table(self.get_supplier_box_data())

    def onClickItemButton(self):
        if self.is_empty_item_box():
            print("please input data into empty slots.")
            return

        content = self.get_item_box_data()
        if self.is_already_existed_item(content):
            print("item is already exist")
            return

        self.add_content_into_item_table(content)

    def onClickSearchButton(self):

        target_mpn = self.searchMpn.text()

        input_contents = self.get_item_list_with_mpn(target_mpn, DataType.INPUT)
        output_contents = self.get_item_list_with_mpn(target_mpn, DataType.OUTPUT)

        self.add_content_into_search_tables(input_contents, output_contents)

        self.init_data_with_dict(self.statisticsTableWidget_2, self.statistics_manager.get_item_statistics_dict(target_mpn))


    def onEnterMpnEdit(self):
        item_info = self.get_item_info_with_mpn(str(self.editMpn.text()))
        self.editPartName.setText(str(item_info[1]))
        index = self.editSupplier.findText(item_info[2], Qt.MatchFixedString)
        if index >= 0:
            self.editSupplier.setCurrentIndex(index)
        self.editPrice.setText(str(item_info[3]))

    def onEnterMpnEdit_2(self):
        item_info = self.get_item_info_with_mpn(str(self.editMpn_2.text()))
        self.editPartName_2.setText(str(item_info[1]))

        item_list = self.get_item_list_with_mpn(str(self.editMpn_2.text()), DataType.OUTPUT)
        supplier_text = item_info[2]
        if len(item_list) > 0:
            supplier_text = item_list[-1][3]

        index = self.editSupplier_2.findText(supplier_text, Qt.MatchFixedString)
        if index >= 0:
            self.editSupplier_2.setCurrentIndex(index)
        self.editPrice_2.setText(str(item_info[3]))


# User Interface
    def make_mpn_completer(self):
        mpn_list = []
        for content in self.item_contents:
            mpn_list.append(str(content[0]))

        model = QStringListModel()
        model.setStringList(mpn_list)

        completer = QCompleter()
        completer.setModel(model)

        self.editMpn.setCompleter(completer)
        self.editMpn_2.setCompleter(completer)
        self.searchMpn.setCompleter(completer)


    def make_logo(self):
        logoPixmap = QPixmap()
        logoPixmap.load("logo2.png")
        # logoPixmap = logoPixmap.scaledToWidth(350)
        logoPixmap = logoPixmap.scaledToHeight(100)
        self.logoLabel.setPixmap(logoPixmap)

        innovationPixmap = QPixmap()
        innovationPixmap.load("innovation.png")
        innovationPixmap = innovationPixmap.scaledToWidth(81)
        self.innovationLabel.setPixmap(innovationPixmap)

    def make_comboBox(self):
        supplier_list = self.get_supplier_list_with_type(SupplierType.SUPPLIER)
        for supplier_name in supplier_list:
            self.editSupplier_3.addItem(supplier_name)
            self.editSupplier.addItem(supplier_name)

        customer_list = self.get_supplier_list_with_type(SupplierType.CUSTOMER)
        for customer_name in customer_list:
            self.editSupplier_2.addItem(customer_name)

    def clear_comboBox(self):
        self.editSupplier.clear()
        self.editSupplier_2.clear()
        self.editSupplier_3.clear()

    def make_bar_chart(self):
        chart = QChart()

        chart.setTitle("현재 재고 현황")
        chart.setAnimationOptions(QChart.SeriesAnimations)

        item_statistics_dict = self.statistics_manager.get_item_statistics_dict()
        categories = ["2020년"]

        series = QBarSeries()
        for key, value in item_statistics_dict.items():
            set = QBarSet(key)
            set << value[REMAIN_COUNT]
            series.append(set)

        chart.addSeries(series)

        axis = QBarCategoryAxis()
        axis.append(categories)
        chart.createDefaultAxes()
        chart.setAxisX(axis, series)

        chart.legend().setVisible(True)
        chart.legend().setAlignment(Qt.AlignBottom)

        self.barChartView = QChartView(chart)
        self.barChartView.setRenderHint(QPainter.Antialiasing)
        self.statisticsWidget.setWidget(self.barChartView)

    def update_bar_chart(self):
        # self.barChartView.clear()
        None

    # onChangedData()
    def onChangedSupplier(self):
        self.clear_comboBox()
        self.make_comboBox()

    def onChangedItem(self):
        self.make_mpn_completer()
        self.update_bar_chart()

    def onChangedManagement(self):
        self.init_data_with_dict(self.statisticsTableWidget, dict=self.statistics_manager.get_item_statistics_dict())
        self.update_bar_chart()

if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()