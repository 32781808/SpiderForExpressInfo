from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
import os
import time
from PIL import Image
import numpy as np
import pandas as pd
import win32api,win32con


class SpiderExpress(object):
    """
    快递信息自动查询类，完成快递信息滑块自动验证，并获取快递信息截图
    """

    def __init__(self, url):
        self.pwd_path = os.getcwd()
        self.serv = Service(os.path.join(self.pwd_path, "chromedriver.exe"))
        self.driver = None
        self.query_url = url
        self.__verification_picture_frame = None  # frame标签定位
        self.__verification_picture_location = None  # frame标签位置
        self.__verification_picture_info = None  # 验证图片在iframe标签中的位置
        self.__verification_slider = None  # 验证图片缺口滑块定位
        self.__drag_button_slider = None  # 验证滑块滑道位置
        self.__query_state = {}  # 记录快递单号查询状态

    def run_auto_slide(self, express_number):
        """
        自动滑块并截图
        :param express_number: 快递单号
        :return:
        """
        self.__get_driver(express_number)
        self.__drive_button()
        self.__find_express_info_and_save(express_number)

    def run_manual_slide(self, express_number):
        """
        手动滑块并截图
        :param express_number:快递单号
        :return:
        """
        self.__get_driver(express_number)
        self.__find_express_info_and_save(express_number)

    def get_query_state(self):
        return self.__query_state

    def __get_driver(self, express_number):
        """
        初始化浏览器驱动
        :param express_number:
        :return:
        """
        self.driver = webdriver.Chrome(service=self.serv)
        url = self.query_url + express_number
        self.driver.get(url)
        self.driver.maximize_window()
        time.sleep(1)

    def __find_express_info_and_save(self, express_number):

        success_element = (By.XPATH, "// *[ @ class = 's-line']")
        try:
            success_image = WebDriverWait(self.driver, 5, 0.5).until(EC.presence_of_element_located(success_element))
            if success_image:
                js = "window.scrollBy(0,350)"
                self.driver.execute_script(js)
                time.sleep(2)
                self.__save_express_info_screen(express_number)
                self.__query_state[express_number] = "已截图"
        except TimeoutException:
            self.__query_state[express_number] = "验证失败"
        except NoSuchElementException:
            self.__query_state[express_number] = "未查询到快递信息"
        finally:
            self.driver.close()

    def __save_express_info_screen(self, express_number):
        """
        自动保存快递信息截图
        :return:
        """
        map_locate = self.driver.find_element(By.XPATH, "// *[ @ id = 'bill-map-{}']".format(express_number))
        map_location = map_locate.location
        map_size = map_locate.size
        routes_locate = self.driver.find_element(By.XPATH, "//*[@class='routes-wrapper']")
        routes_location = routes_locate.location
        routes_size = routes_locate.size

        x = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)  # 获得屏幕分辨率X轴
        y = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)  # 获得屏幕分辨率Y轴

        rangle = (
            map_location['x'] * int(x/1920) + 100,
            map_location['y'] *int(y/925)- 290,
            map_location['x'] *int(x/1920)+100+ map_size['width']+300,
            map_location['y'] *int(y/925)- 290 + map_size['height'] + routes_size['height']
        )
        time.sleep(1)
        self.driver.save_screenshot("./express_info_picture.png")
        express_img = Image.open("./express_info_picture.png")
        express_img = express_img.convert('RGB')
        express_img = express_img.resize((1920, 925), Image.ANTIALIAS)
        picture_crop = express_img.crop(rangle)
        self.mkdir()
        picture_crop.save("./picture/{}.png".format(express_number))

    def __get_verification_picture(self):
        """
        获取带缺口的验证图片
        :return:
        """
        all_screen = self.driver.save_screenshot("./all_screen.png")
        img_verification = Image.open("./all_screen.png")
        img_verification = img_verification.convert("L")
        rangle = self.__verification_picture_rangle()
        img_verification = img_verification.resize((1920, 925), Image.ANTIALIAS)

        return img_verification.crop(rangle)

    def __get_slider_distance(self, img):
        """
        基于带缺口的验证图片，获取缺口距离
        :param img:
        :return:
        """
        picture_metrix = np.array(img)
        slider_picture_location = self.__verification_slider.location
        verification_location = self.__verification_picture_info.location
        location_start = int(slider_picture_location['y'] - verification_location['y'] + 34)

        i = 230

        while (i < 280):
            data1 = picture_metrix[location_start - 19:location_start + 19, i - 1]
            data2 = picture_metrix[location_start - 19:location_start + 19, i]
            pixel_sub1 = data1 - data2
            pixel1 = sum(pixel_sub1 > 70)
            pixel_sub2 = picture_metrix[location_start - 19:location_start + 19, i + 43] - picture_metrix[
                                                                                           location_start - 19:location_start + 19,
                                                                                           i + 42]
            pixel2 = sum(pixel_sub2 > 70)
            if (pixel1 > 24) & (pixel2 > 24):
                distance = i - 22
                break

            i += 1
        else:
            distance = 230

        return distance

    def __drive_button(self):
        """
        格局缺口距离进行滑块自动移动
        :return:
        """
        self.__locate_all_label()
        img = self.__get_verification_picture()
        distance = self.__get_slider_distance(img)

        tracks = self.__get_track(distance)
        self.__get_drag_button_slider()
        ActionChains(self.driver).click_and_hold(self.__drag_button_slider).perform()
        for track in tracks:
            ActionChains(self.driver).move_by_offset(track, 0).perform()
        ActionChains(self.driver).release().perform()
        self.driver.switch_to.default_content()

    def __get_track(self, distance: int):
        """
        根据滑块缺口距离构造滑动轨迹
        :param distance: 滑块缺口距离
        :return: 轨迹集合
        """
        # distance为传入的总距离
        # 移动轨迹
        track = []
        # 当前位移
        current = 0
        # 减速阈值
        mid = distance * 4 / 5
        track.append(int(mid))
        current = int(mid) + 3

        v = 10

        while current < distance:
            if current < mid:
                # 加速度为2
                a = 4
            else:
                # 加速度为-2
                a = -v / 10
            v0 = v
            # 当前速度
            v = v0 + a
            # 移动距离
            move = v0 + 1 / 2 * a
            # 当前位移
            current += move
            # 加入轨迹
            track.append(round(move))
        return track

    def __verification_picture_rangle(self):
        """
        获取验证图片裁剪位置
        :return:
        """
        verification_picture_location = self.__verification_picture_location
        verification_picture_in_frame_location = self.__verification_picture_info.location
        verification_picture_in_frame_sizes = self.__verification_picture_info.size

        rangle = (
            verification_picture_in_frame_location['x'] + verification_picture_location['x'],
            verification_picture_in_frame_location['y'] + verification_picture_location['y'],
            verification_picture_in_frame_location['x'] + verification_picture_location['x'] +
            verification_picture_in_frame_sizes['width'],
            verification_picture_in_frame_location['y'] + verification_picture_location['y'] +
            verification_picture_in_frame_sizes['height'],
        )
        return rangle

    # 网页元素定位
    def __get_verification_picture_iframe(self):
        """
        定位iframe标签
        :return:
        """
        self.__verification_picture_frame = self.driver.find_element(By.XPATH, "//iframe[@id='tcaptcha_popup']")

    def __locate_all_label(self):
        """
        获取验证图片信息（在整个网页的位置，大小）
        :return:
        """
        self.__get_verification_picture_iframe()
        self.__verification_picture_location = self.__verification_picture_frame.location
        # 获取iframe标签内的验证图片信息
        self.driver.switch_to.frame(self.driver.find_element(By.XPATH, "//iframe[@id='tcaptcha_popup']"))
        self.__get_verification_picture_to_frame()
        self.__get_verification_slider()
        self.__get_drag_button_slider()

    def __get_verification_picture_to_frame(self):
        """
        获取验证图片在iframe嵌套标签的位置
        :return:
        """
        self.__verification_picture_info = self.driver.find_element(By.XPATH, "//div[@class='tcaptcha-imgarea drag']")

    def __get_verification_slider(self):
        """
        获取缺口图片在iframe标签中的位置及大小
        :return:
        """
        self.__verification_slider = self.driver.find_element(By.XPATH, "//*[@id='slideBlock']")

    def __get_drag_button_slider(self):
        """
        获取滑块按钮的位置
        :return:
        """
        self.__drag_button_slider = self.driver.find_element(By.XPATH, "//*[@id='tcaptcha_drag_button']")

    def __get_verification_picture_info(self):
        """
        定位iframe
        :return:
        """
        return self.driver.find_element(By.XPATH, "//iframe[@id='tcaptcha_popup']")

    @staticmethod
    def get_express_number(file_path):
        """
        获取快递单号列表
        :param file_path:快递单号存储的exel表格
        :return:快递单号的dataframe
        """
        express_num_series = pd.read_excel(file_path, header=0)
        return express_num_series[express_num_series.columns[0]]

    @staticmethod
    def mkdir():
        """
        创建picture文件夹
        :return:
        """
        pwd_path = os.getcwd()
        express_picture_file = os.path.join(pwd_path, "picture")
        if not os.path.exists(express_picture_file):
            os.mkdir(express_picture_file)

        return express_picture_file


if __name__ == '__main__':
    spider_express = SpiderExpress("")
    expree_nums = spider_express.get_express_number('新建 Microsoft Excel 工作表(1).xlsx')
    picture_path = spider_express.mkdir()
    while True:
        exist_express_names = [item[:-4] for item in os.listdir(picture_path)]
        for item in expree_nums:
            if item in exist_express_names:
                continue
            spider_express.run_auto_slide(item)  # 自动滑块
            # sparse_express.run_manual_slide(item)     # 手动滑块
        if "验证失败" not in spider_express.get_query_state().values():
            break
        pass
