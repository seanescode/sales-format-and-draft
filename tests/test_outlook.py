from modules.outlook import is_outlook_running


def test_is_outlook_running_returns_boolean():
    result = is_outlook_running()

    assert isinstance(result, bool)