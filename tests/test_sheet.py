import pytest
from tests import sheet

def test_get_dates(sheet):
    begin_date = "2021.09.02"
    days = 5
    dates = ["2021.09.02", "2021.09.03", "2021.09.04", "2021.09.05", "2021.09.06"]
    assert dates == sheet.get_dates(begin_date=begin_date, days=days)

def test_sheet_append(sheet):
    begin_date = "2021.09.02"
    value = "测试"
    dates = sheet.get_dates(begin_date=begin_date, days=1)
    names = ["周立君", "贾春宝"]
    sheet.sheet_append(name=names, dates=dates, value=value)

    value_dict = sheet.get_value()
    assert value_dict.get("周立君") == ("2021.09.02", "测试")
    assert value_dict.get("贾春宝") == ("2021.09.02", "测试")


