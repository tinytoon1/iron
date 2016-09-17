import pytest
import json
import os
from fixture.application import App

import clr

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=12.0.0.0, Culture=Neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

fixture = None
target = None


def load_config(file):
    global target
    if target is None:
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), file)
        with open(config_file) as config:
            target = json.load(config)
    return target


@pytest.fixture
def app(request):
    global fixture
    app_config = load_config(request.config.getoption("--target"))["app"]
    if fixture is None:
        fixture = App(path=app_config["where_is_app"], title=app_config["main_window_title"])
        fixture.run_application()
    return fixture


@pytest.fixture(scope="session", autouse=True)
def stop(request):
    def fin():
        fixture.destroy()
    request.addfinalizer(fin)
    return fixture


def pytest_generate_tests(metafunc):
    for fixture in metafunc.fixturenames:
        if fixture.startswith("xls_"):
            test_data = load_from_xls(fixture[4:])
            metafunc.parametrize(fixture, test_data, ids=[str(x) for x in test_data])


def load_from_xls(file):
    test_data = []
    excel_data = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data\%s.xlsx" % file)
    excel = Excel.ApplicationClass()
    excel.Visible = False
    if os.path.exists(excel_data):
        workbook = excel.Workbooks.Open(excel_data)
        worksheet = workbook.ActiveSheet
        for i in range(1, 100):
            cell = worksheet.Range['A%s' % i].Value2
            if cell is not None:
                test_data.append(cell)
        excel.Quit()
    return test_data
