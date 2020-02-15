__author__ = "Chanhee Jo"
__copyright__ = "Copyright 2020, SB Plus Co.,Ltd"
__license__ = "GPL"
__version__ = "1.0.0"
__email__ = "teletovy@gmail.com, decision_1@naver.com"

import enum

class DataType(enum.Enum):
    INPUT = "입고"   # 입고
    OUTPUT = "출고"   # 출고

class SheetType(enum.Enum):
    MANAGEMENT = "management_info"
    SUPPLIER = "supplier_info"
    ITEM = "item_info"

class SupplierType(enum.Enum):
    SUPPLIER = "Supplier"
    CUSTOMER = "Customer"

class ItemAttribute(enum.Enum):
    # Item attribute
    INPUT_OUTPUT = "input_output"
    MPN = "mpn"
    COUNT = "count"
    PRICE = "price"