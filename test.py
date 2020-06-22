import ctypes
import logging
import time

import pyautogui
from pywinauto.application import Application

logging.basicConfig(level='INFO')
str_file_path = r"C:\Program Files\FileZilla FTP Client\filezilla.exe"


def wait_locate_center_on_screen(image_file_path):
    """
    为图像识别增加等待功能
    :param image_file_path: 图片文件路径
    :return:
    """
    find_count = 0
    while True:
        """
        confidence: 图片匹配度
        """
        check_update_box_pos = pyautogui.locateCenterOnScreen(image_file_path, confidence=0.8)
        if check_update_box_pos:
            return check_update_box_pos
        logging.info("识别中, 当前次数为" + str(find_count))
        find_count = find_count + 1
        time.sleep(0.5)


def fix_input_method_mode():
    logging.info("fix_input_method_mode?")
    pyautogui.write('a', interval=0.25)
    if condition_input_method_mode():
        pyautogui.press(['shift'])
    pyautogui.press(['backspace'])
    logging.info("fix_input_method_mode.")


def condition_input_method_mode():
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    lid_hex = hex(
        (user32.GetKeyboardLayout(user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), 0))) & (2 ** 16 - 1))

    is_chinese_input_method_mode = False
    if lid_hex == '0x804':
        is_chinese_input_method_mode = True
    return is_chinese_input_method_mode


test_data = {
    "host": u"10.10.50.21",
    "username": u"tristan",
    "password": u"test",
    "port": u"22",
}

pyautogui.wait_locate_center_on_screen = wait_locate_center_on_screen


def do_something(service_code, service_name):
    logging.info("%s框出现进行中" % service_name)
    something_box_check_pos = pyautogui.wait_locate_center_on_screen(
        r'image_location_resources/%s.png' % service_code)
    logging.info("%s框检查出现" % service_name)
    logging.info(something_box_check_pos)
    something_box_check_pos_x, something_box_check_pos_y = something_box_check_pos
    something_box_check_pos_x = something_box_check_pos_x + 60
    pyautogui.click(something_box_check_pos_x, something_box_check_pos_y)  # 点击输入框
    logging.info("%s输入框点击" % service_name)
    # 清理输入框中的文字
    logging.info("清理%s输入框中的文字进行中" % service_name)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press(['backspace'])
    logging.info("清理%s输入框中的文字" % service_name)
    pyautogui.write(test_data[service_code], interval=0.25)  # 输入
    logging.info("%s输入框输入" % service_name)


def test():
    fix_input_method_mode()
    do_something("port", "端口")
    do_something("password", "密码")
    do_something("username", "用户名")
    do_something("host", "主机")
    time.sleep(10)


if __name__ == '__main__':
    execute_file_path = str_file_path
    Application().start(execute_file_path)
    app = Application(backend='uia').connect(path=execute_file_path)
    win_filezilla = app.window(title=u'FileZilla')
    try:
        test()
        win_filezilla.close()
    except Exception as e:
        win_filezilla.close()
        raise e
