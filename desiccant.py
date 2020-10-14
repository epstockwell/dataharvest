import pytesseract
import pyautogui
import time
import json
import win32clipboard
import cv2
import os
import csv
import numpy as np
from scipy.interpolate import griddata
from PIL import Image


class DesiccantScraper:
    def __init__(self):
        self.win_loc = pyautogui.locateOnScreen("Templates/dri.png")
        self.getwintext_loc = pyautogui.locateOnScreen("Templates/getwintext.png")
        self.error_loc = pyautogui.locateOnScreen("Templates/error_ok.png")

        self.coord = {}
        with open("positions.json", "r") as in_file:
            self.coord = json.load(in_file)

    def update_interface(self):
        self.win_loc = pyautogui.locateOnScreen("Templates/dri.png")
        self.getwintext_loc = pyautogui.locateOnScreen("Templates/getwintext.png")
        self.error_loc = pyautogui.locateOnScreen("Templates/error_ok.png")

    @staticmethod
    def extract_text(screenshot, left, top, width, height):
        cp = screenshot.crop((left, top, left + width, top + height))
        return cp

    def click_shit(self, name):
        x = self.coord[name]['x'] + self.win_loc.left + self.coord[name]['Width'] // 2
        y = self.coord[name]['y'] + self.win_loc.top + self.coord[name]['Height'] // 2
        pyautogui.click(x=x, y=y)

    def click_error(self, name):
        x = self.coord[name]['x'] + self.error_loc.left + self.coord[name]['Width'] // 2
        y = self.coord[name]['y'] + self.error_loc.top + self.coord[name]['Height'] // 2
        pyautogui.click(x=x, y=y)

    def write_shit(self, name, text):
        x = self.coord[name]['x'] + self.win_loc.left + self.coord[name]['Width'] // 2
        y = self.coord[name]['y'] + self.win_loc.top + self.coord[name]['Height'] // 2
        pyautogui.click(x=x, y=y)
        time.sleep(0.1)
        pyautogui.click(x=x, y=y)
        pyautogui.write(str(text))

    def read_shit(self, name_list):
        result = []
        s = pyautogui.screenshot()
        for name in name_list:
            x = self.coord[name]['x'] + self.win_loc.left
            y = self.coord[name]['y'] + self.win_loc.top
            width = self.coord[name]['Width']
            height = self.coord[name]['Height']
            result.append(self.extract_text(s, x, y, width, height))
        return result

    def read_shit_copy(self, name, cast_to_float=True):
        x = self.coord[name]['x'] + self.win_loc.left + self.coord[name]['Width'] // 2
        y = self.coord[name]['y'] + self.win_loc.top + self.coord[name]['Height'] // 2
        pyautogui.click(x=x, y=y)
        time.sleep(0.1)
        pyautogui.click(x=x, y=y)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        if cast_to_float:
            try:
                data = float(data)
            except ValueError:
                data = "NaN"
        return data

    def read_many_shit_copy(self, name_list, cast_to_float=True):
        output = []
        for name in name_list:
            output.append(self.read_shit_copy(name, cast_to_float))
        return output

    def read_drag_shit(self, name, cast_to_float=True):
        x = self.coord[name]['x'] + self.win_loc.left + self.coord[name]['Width'] // 2
        y = self.coord[name]['y'] + self.win_loc.top + self.coord[name]['Height'] // 2
        pyautogui.moveTo(self.getwintext_loc.left + 30, self.getwintext_loc.top + 30)
        pyautogui.mouseDown()
        pyautogui.moveTo(x, y)
        pyautogui.mouseUp()
        pyautogui.moveTo(x=self.getwintext_loc.left + 100, y=self.getwintext_loc.top + 50)
        pyautogui.click()
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        if cast_to_float:
            data = float(data)
        return data

    def read_many_drag_shit(self, name_list, cast_to_float=True):
        output = []
        for name in name_list:
            output.append(self.read_drag_shit(name, cast_to_float))
        return output

    def read_regen(self, name_list):
        result = []
        for name in name_list:
            x = self.coord[name]['x'] + self.win_loc.left
            y = self.coord[name]['y'] + self.win_loc.top
            pyautogui.click(x=x, y=y, clicks=2)
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.1)
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()
            result.append(data)
        return result

    def adjust_energy(self, target):
        self.write_shit("Target Temp", "140")
        self.click_shit("Calculate")
        time.sleep(3)
        result = [float(self.read_drag_shit("React Energy")), float(self.read_regen(["Target Temp"])[0])]
        x = result[0]
        print(result)
        while x < target - 0.05 or x > target + 0.05:
            if x > target:
                self.write_shit("Target Temp", str(int(result[1] - 1)))
                error_exit()
            else:
                self.write_shit("Target Temp", str(int(result[1] + 1)))
                error_exit()
            self.click_shit("Calculate")
            t = time.process_time()
            while np.array(pyautogui.screenshot(
                    region=(self.win_loc.left + 160, self.win_loc.top + 200, 1, 1)
            ))[0, 0, 1] != 255:
                time.sleep(0.1)
                if time.process_time() - t > 3:
                    break
                error_exit()
            result = [float(self.read_drag_shit("React Energy")), float(self.read_regen(["Target Temp"])[0])]
            x = result[0]
        return result[0]

    def screen_setup(self, dia, depth, rph, by_process, process_cfm, react_cfm, react_temp, react_grains):
        # dia = 300
        # depth = 100
        # rph = 11
        # by_process = 63
        # process_cfm = 206
        # react_cfm = 38
        # react_temp = 75
        # react_grains = 64.7

        self.click_shit("IP")

        # writes needed
        self.write_shit("Diameter", str(dia))
        self.write_shit("Non-Standard Depth", str(depth))
        self.write_shit("RPH", str(rph))
        self.write_shit("Bypass Process", str(by_process))
        self.write_shit("Airflow scfm", str(process_cfm))
        self.write_shit("Reactivation CFM", str(react_cfm))
        self.write_shit("Reactivation Temp", str(react_temp))
        self.write_shit("Reactivation Grains", str(react_grains))

        # click adjustments
        self.click_shit("IP")
        self.click_shit("Non-standard")
        self.click_shit("Non-Standard")
        self.click_shit("DBT & RH")
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('up')
        pyautogui.press('enter')

    def process_temp_pull(self):
        text_box = convert_to_black(self.read_shit(["Process Out Temp"])[0])
        img_rgb = np.array(text_box)
        img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)

        result_text = ["" for _ in range(img_gray.shape[1])]

        file_names = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, "point", "Nan"]
        for cell in file_names:
            file_name = f"number images/{cell}_numbers.png"
            template = cv2.imread(file_name, cv2.IMREAD_UNCHANGED)[:, :, 2]

            res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
            threshold = 0.7
            loc = np.where(res >= threshold)

            for match in loc[1]:
                result_text[match] = str(cell) if cell != "point" else "."

        result_text = [elem for elem in result_text if elem != ""]
        result_num = "".join(result_text)
        return result_num


def get_grains():
    return np.genfromtxt("Grains.csv", delimiter=',', skip_header=1).reshape((20, 25, 3))


def grain_interpolation(temp, humidity):
    grain_input = get_grains()
    points = [temp, humidity]
    return griddata(grain_input[:, :, 0:2].reshape((-1, 2)), grain_input[:, :, 2].flatten(), points)


def convert_to_black(img, threshold=300):
    img_arr = np.array(img)
    out_img = 255 * np.ones(img_arr.shape, dtype=np.uint8)

    out_img[np.linalg.norm(img_arr, axis=2) < threshold] = [0, 0, 0]
    return Image.fromarray(out_img)


def error_exit():
    if pyautogui.locateCenterOnScreen('Templates/error_ok.png') is not None:
        x, y = pyautogui.locateCenterOnScreen('Templates/error_ok.png')
        pyautogui.click(x, y)
        return True
    return False


def read_csv(file_name="Output_at_conditions.csv"):
    result_ar = []
    if os.path.exists(file_name):
        with open(file_name, "r") as csvfile:
            result_ar = [row for row in csv.reader(csvfile)]
        result_ar = result_ar[1:] if len(result_ar) > 1 else []
    return result_ar


def get_data(column_name):
    row_names = ['Reactivation Outlet CFM', 'Reactivation Outlet Temp',
                 'Reactivation Outlet gr/lb', 'React Energy',
                 'Moisture Removal', 'Reactivation CFM', 'Reactivation Temp', 'Reactivation Grains',
                 'Airflow scfm', 'Temperature', 'Moisture', 'Target Temp', 'Process Temp', 'Error']

    assert column_name in row_names

    temp = []
    grains = []
    column_data = []
    temp_idx = row_names.index('Temperature')
    grains_idx = row_names.index('Moisture')
    col_idx = row_names.index(column_name)

    in_data = read_csv()

    def true_if_error(item):
        try:
            _ = float(item)
            return False
        except TypeError:
            return True
        except ValueError:
            return True

    for row_idx in range(len(in_data)):
        in_temp = in_data[row_idx][temp_idx]
        in_grains = in_data[row_idx][grains_idx]
        in_item = in_data[row_idx][col_idx]

        errors = true_if_error(in_temp) or true_if_error(in_grains) or true_if_error(in_item)

        if not errors:
            temp.append(float(in_temp))
            grains.append(float(in_grains))
            column_data.append(float(in_item))

    return grains, temp, column_data

