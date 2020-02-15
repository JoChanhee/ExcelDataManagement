__author__ = "Chanhee Jo"
__copyright__ = "Copyright 2020, SB Plus Co.,Ltd"
__license__ = "GPL"
__version__ = "1.0.0"
__email__ = "teletovy@gmail.com, decision_1@naver.com"

from model.data_type import DataType, ItemAttribute

# Item statistics refinement attribute
INPUT_AVG_PRICE = "입고 단가"    # "input_avg_price"
INPUT_ITEM_NUM = "입고 수량"     # "input_item_num"
OUTPUT_AVG_PRICE = "출고 단가"   # "output_avg_price"
OUTPUT_ITEM_NUM = "출고 수량"    # "output_item_num"

# Item statistics output attribute
REMAIN_COUNT = "현재고"   # "remain_count"
PROFIT = "순이익"                # "profit"
TOTAL_PRICE = "총 재고 단가"  # "total_price"


class StatisticsManager(object):
    def __init__(self, header, contents):
        super().__init__()

        self.header = header
        self.contents = contents

    def get_item_statistics_dict(self, target_mpn=None):
        mpn_idx = 0
        data_type_idx = 2
        count_idx = 5
        price_idx = 6
        for idx, attribute in enumerate(self.header):
            if attribute == ItemAttribute.INPUT_OUTPUT.value:
                data_type_idx = idx
            elif attribute == ItemAttribute.MPN.value:
                mpn_idx = idx
            elif attribute == ItemAttribute.COUNT.value:
                count_idx = idx
            elif attribute == ItemAttribute.PRICE.value:
                price_idx = idx

        item_dict = {}
        for item in self.contents:
            # Input data
            key = item[mpn_idx]
            count = item[count_idx]
            price = item[price_idx]

            # construct only target mpn dict
            if target_mpn is not None and key != target_mpn:
                continue

            # Data Initialization
            if key not in item_dict:
                value = {}
                value[REMAIN_COUNT] = 0
                value[PROFIT] = 0.0
                value[INPUT_AVG_PRICE] = 0.0
                value[INPUT_ITEM_NUM] = 0
                value[OUTPUT_AVG_PRICE] = 0.0
                value[OUTPUT_ITEM_NUM] = 0
                value[TOTAL_PRICE] = 0.0

                item_dict[key] = value

            # Data refinement / Output data
            if item[data_type_idx] == DataType.INPUT.value:
                item_dict[key][INPUT_AVG_PRICE] += (count * price)
                item_dict[key][INPUT_ITEM_NUM] += count
            elif item[data_type_idx] == DataType.OUTPUT.value:
                item_dict[key][OUTPUT_AVG_PRICE] += (count * price)
                item_dict[key][OUTPUT_ITEM_NUM] += count

        for key, value in item_dict.items():
            if item_dict[key][INPUT_ITEM_NUM] != 0:
                item_dict[key][INPUT_AVG_PRICE] = item_dict[key][INPUT_AVG_PRICE] / float(item_dict[key][INPUT_ITEM_NUM])
            if item_dict[key][OUTPUT_ITEM_NUM] != 0:
                item_dict[key][OUTPUT_AVG_PRICE] = item_dict[key][OUTPUT_AVG_PRICE] / float(item_dict[key][OUTPUT_ITEM_NUM])

            item_dict[key][REMAIN_COUNT] = item_dict[key][INPUT_ITEM_NUM] - item_dict[key][OUTPUT_ITEM_NUM]
            item_dict[key][TOTAL_PRICE] = item_dict[key][INPUT_AVG_PRICE] * item_dict[key][REMAIN_COUNT]
            item_dict[key][PROFIT] = (item_dict[key][OUTPUT_AVG_PRICE] - item_dict[key][INPUT_AVG_PRICE]) * item_dict[key][OUTPUT_ITEM_NUM]

        return item_dict

    def update(self, header, contents):
        self.header = header
        self.contents = contents