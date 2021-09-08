import pytest
from tests import sheet

import sys
import os
sys.path.append("..")
from app.Append_xlsx import get_name_value, append_xlsx

def test_get_append(sheet):
    name_value = get_name_value()
    assert name_value.get("不限") == ["周立君", "李春明"]

    filename = os.path.join("." + os.sep, "排班导入计划.xlsx")
    begin_date = "2021.09.08"
    days = 1
    t = append_xlsx(name_value=name_value, filename=filename, begin_date=begin_date, days=days)
    assert t is True