import configparser


def test_config_loads():
    config = configparser.ConfigParser()
    config.read("config.ini")

    assert config.has_section("EMAIL_SETTINGS")
    assert config.has_section("EXCEL_REPORT_FORMATTING")