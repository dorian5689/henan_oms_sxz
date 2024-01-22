#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""
@Time ： 2024/1/16 17:22
@Auth ： Xq
@File ：run_henan_oms_sxz.py
@IDE ：PyCharm
"""

import os
import time

import schedule
from DrissionPage import ChromiumPage
from DrissionPage._configs.chromium_options import ChromiumOptions

from DataBaseInfo.MysqlInfo.MysqlTools import MysqlCurd
from FindSoft.Find_Exe import FindExeTools
from MacInfo.ChangeMAC import SetMac
from XpathConfig.HenanXpath import henan_ele_dict
import ddddocr
import re

import datetime
from DrissionPage.common import Keys

from Config.ConfigUkUsb import henan_wfname_dict_num
from ProcessInfo.ProcessTools import ProcessCure
from UkChange.run_ukchange import Change_Uk_Info
from DingInfo.DingBotMix import DingApiTools


class ReadyLogin(object):

    def __init__(self):
        pass

    def select_uk(self):
        CU = Change_Uk_Info()
        list_port = CU.select_comports()
        if list_port is None:
            return None, None
        exe_path = F'..{os.sep}ExeSoft{os.sep}HUB_Control通用版{os.sep}HUB_Control通用版.exe'
        process_name = F"HUB_Control通用版.exe"

        PT = ProcessCure()
        PT.admin_kill_process(process_name)

        import subprocess
        subprocess.Popen(exe_path)
        time.sleep(1)
        CU.select_use_device()

        import pygetwindow as gw
        res = gw.getWindowsWithTitle('HUB_Control通用版 示例程序')[0]
        time.sleep(3)
        return res, CU

    # def select_usb_id(self):
    #     from Config.ConfigUkUsb import henan_wfname_dict_num
    #     usb_ids = []
    #     for k, v in henan_wfname_dict_num.items():
    #         usb_ids.append(k)
    #     return usb_ids

    def change_usbid(self):
        res, CU = self.select_uk()
        from Config.ConfigUkUsb import henan_wfname_dict_num
        report_li = []
        for i, uuid in henan_wfname_dict_num.items():
            from datetime import datetime
            # 获取当前时间
            current_time = datetime.now()
            # 格式化当前时间
            start_run_time = current_time.strftime("%Y-%m-%d %H:%M:%S")
            new_nanfang = F'../DataBaseInfo/MysqlInfo/new_nanfang.yml'
            slect_zhuangtai_sql = F"select  usb序号,UK密钥MAC地址,场站,外网oms账号,外网oms密码,wfname_id  from data_oms_uk  where usb序号='{i}' and uuid ='{uuid}'  "

            data_info = MysqlCurd(new_nanfang).query_sql_return_header_and_data(slect_zhuangtai_sql).values.tolist()
            import datetime
            # 获取当前日期
            current_date = datetime.date.today()
            # print("当前日期为：", current_date)

            # 将当前日期减去1天
            previous_day = current_date - datetime.timedelta(days=1)
            # print("前一天的日期为：", previous_day)
            select_exit_true = F"SELECT 是否已完成 FROM data_oms where 电场名称='{data_info[0][2]}' AND 日期='{previous_day}'"
            res_exit_ture = MysqlCurd(new_nanfang).query_sql(select_exit_true)

            if res:
                time.sleep(2)
                res.maximize()
                time.sleep(3)
                CU.all_button()
                time.sleep(1)
                CU.radio_switch(f'{i}')
                time.sleep(3)
                res.minimize()

            for data in data_info:
                userid = int(data[0])
                mac_address = data[1]
                wfname = data[2]
                username = data[3]
                password = data[4]
                wfname_id = data[5]

                set_mac = SetMac()
                new_mac = mac_address
                set_mac.run(new_mac)
                time.sleep(1)
                try:
                    FT = FindExeTools()
                    FT.find_soft()
                except:
                    break
                time.sleep(3)

                try:
                    from datetime import datetime
                    # 获取当前时间
                    current_time = datetime.now()
                    # 格式化当前时间
                    start_run_time = current_time.strftime("%Y-%m-%d %H:%M:%S")
                    try:
                        RB = RunSxz(username, password, wfname, userid, wfname_id, start_run_time)
                        run_num = RB.run_sxz(userid)
                        print(F'运行一次的值:{run_num}')
                        if run_num == 1:
                            continue
                    except Exception as e:
                        pass
                        print(f'已经运行了一次{e}')
                except Exception as e:

                    print(F'主函数问题:{e}')
                    pass
            report_li.append(userid)
        return report_li


class RunSxz(object):
    def __init__(self, username=None, password=None, wfname=None, userid=None, wfname_id=None, start_run_time=None):
        """
        基于谷歌内核
        """
        self.username = username
        self.password = password
        self.wfname = wfname
        self.userid = userid
        self.wfname_id = wfname_id
        self.start_run_time = start_run_time
        # self.jf_token = F'c8eb8d7b8fe2a3c07843233bf225082126db09ab59506bd5631abef4304da29e'
        # 天润
        self.jf_token = F'c8eb8d7b8fe2a3c07843233bf225082126db09ab59506bd5631abef4304da29e'
        # 奈卢斯
        self.nls_token = F'acabcf918755694f2365051202cf3921a690594c1278e4b7fe960186dce58977'

        self.page = self.check_chrome()
        self.login = F"https://172.21.5.193:19070/app-portal/index.html"
        self.today = datetime.datetime.now().strftime("%Y-%m-%d")
        # 获取前一天的日期
        self.today_1 = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
        # 天润
        self.appkey = "dingtc3bwfk9fnqr4g7s"  # image测试
        self.appsecret = "C33oOe03_K5pitN_S2dUppBwgit2VnPW0yWnWYBM3GzogGKhdy2yFUGREl9fLICU"  # image测试
        self.chatid = "chatf3b32d9471c57b4a5a0979efdb06d087"  # image测试
        # 奈卢斯
        self.nls_appkey = "dingjk2duanbfvbywqzx"  # image测试
        self.nls_appsecret = "ICYb4-cvsvIk5DwuZY9zehc5UbpldqIClzS6uuIYFrhjU9z11guV6lold1qNqc2k"  # image测试
        self.nls_chatid = "chatf8ef1e955cf2c4e83a7e776e0011366c"  # image测试

        self.message_dl = {
            "msgtype": "markdown",
            "markdown": {
                "title": "OMS推送",
                "text":
                    F'第{self.wfname_id}个场站:{self.wfname}--已上报--电量--郑州集控'
            }
        }
        self.message_cn = {
            "msgtype": "markdown",
            "markdown": {
                "title": "OMS推送",
                "text":
                    F'第:{self.wfname_id}个场站:{self.wfname}--已上报--储能--郑州集控'
            }
        }

    def check_chrome(self):
        """
        有谷歌走谷歌,没有走edge
        暂时不打包谷歌,300M
        :return:
        """
        user_name = os.getlogin()

        try:
            co = ChromiumOptions()
            co.set_argument("--start-maximized")
            page = ChromiumPage(co)
            return page
        except Exception as e:
            try:
                browser_path = F'C:{os.sep}Users{os.sep}{user_name}{os.sep}AppData{os.sep}Local{os.sep}Google{os.sep}Chrome{os.sep}Application{os.sep}chrome.exe'
                co = ChromiumOptions().set_paths(browser_path=browser_path)
                co.set_argument("--start-maximized")
                page = ChromiumPage(co)
                return page

            except Exception as e:
                browser_path = F'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
                co = ChromiumOptions().set_paths(browser_path=browser_path)
                co.set_argument("--start-maximized")
                page = ChromiumPage(co)
                # print(F"异常值:{e}")
                print(F"找不到本机谷歌浏览器,使用的是edge!")
                return page

    def chrome_login_page(self, userid):

        try:
            # self.page.ele(F'{henan_ele_dict.get("again_post")}')
            # self.page.clear_cache(cookies=False)
            # time.sleep(2)
            time.sleep(1)
            self.page.get(self.login)
            time.sleep(3)
            # self.page.refresh()
            # self.page.wait
            # time.sleep(3)
            try:
                if "风险防控系统" in self.page.html:
                    self.page.ele('x://*[@id="app"]/section/header/div/div[2]/div[1]/div/span').click()
                    self.page.ele('x://html/body/ul/li[1]/span').click()
                    self.page.ele('x://html/body/div[2]/div/div[3]/button[2]/span').click()

            except:
                pass
            try:
                if "点击详情" in self.page.html:
                    self.page.ele(F'{henan_ele_dict.get("details-button")}').click()
                    self.page.ele(F'{henan_ele_dict.get("proceed-link")}').click()
            except:
                pass

            if int(userid) == 1:
                try:
                    self.exit_username_login()
                except Exception as e:
                    pass
            time.sleep(5)
            self.page(F'{henan_ele_dict.get("input_text")}').input(self.username)
            self.page(F'{henan_ele_dict.get("input_password")}').input(self.password)
            time.sleep(2)

            cap_text = self.send_code()

            return cap_text

        except Exception as e:
            pass
            # try:
            #     print(F'验证码error{e}')
            #     cap_text = self.chrome_login_page(userid)
            #     print(F'重新运行后验证：{cap_text}')
            #     return cap_text
            # except:
            #     if not cap_text:
            #         return

    def send_code(self):
        cap = self.page.ele(F'{henan_ele_dict.get("capture_img")}')
        cap.click()
        time.sleep(2)

        import os
        path = os.path.abspath(__file__)
        par_path = os.path.dirname(os.path.dirname(path))
        path = F"{par_path}{os.sep}Image{os.sep}CaptureImg"
        img_name = "验证码.png"
        img_path = F"{path}{os.sep}{img_name}"
        try:
            import shutil
            shutil.rmtree(path)
        except:
            pass
        cap.get_screenshot(path=img_path, name=img_name, )
        ocr = ddddocr.DdddOcr(beta=True)
        with open(img_path, 'rb') as f:
            img_bytes = f.read()
        cap_text = ocr.classification(img_bytes)
        print(f"验证码:{cap_text}")
        # 验证码长度不等于5或者包含中文字符或者不包含大写字母再次运行
        if len(cap_text) == 5 and cap_text.isalnum() and cap_text.islower():
            print(F'验证码有误:{cap_text}')
            return cap_text
        else:
            cap_text = self.send_code()
            print(F'再次验证:{cap_text}')
            return cap_text

    def welcome_user(self, ):
        if "欢迎登录" in self.page.html:
            cap_text = self.send_code()
            time.sleep(3)
            self.page.ele(F'{henan_ele_dict.get("capture_img_frame")}').input(cap_text)
            time.sleep(3)

            self.page.ele(F'{henan_ele_dict.get("login_button")}').click()
            time.sleep(3)
            self.page.wait
            return True
        else:
            return False

    def run_sxz(self, userid):
        cap_text = self.chrome_login_page(userid)
        print(F'已经识别的验证码:{cap_text}')
        time.sleep(3)
        self.page.ele(F'{henan_ele_dict.get("capture_img_frame")}').input(cap_text)
        time.sleep(3)
        self.page.ele(F'{henan_ele_dict.get("login_button")}').click()
        time.sleep(3)
        self.page.wait
        for _ in range(3):
            num_ = self.welcome_user()
            if num_ == True:
                continue

        if "验证码" in self.page.html:
            time.sleep(3)

            cap_text = self.send_code()
            time.sleep(3)

            self.page.ele(F'{henan_ele_dict.get("capture_img_frame")}').input(cap_text)
            self.page.ele(F'{henan_ele_dict.get("login_button")}').click()  # 登录按钮
            self.page.wait

        if "解析密码错误" in self.page.html:
            time.sleep(3)

            cap_text = self.send_code()
            self.page.ele(F'{henan_ele_dict.get("capture_img_frame")}').input(cap_text)
            self.page.ele(F'{henan_ele_dict.get("login_button")}').click()  # 登录按钮
            self.page.wait

        time.sleep(6)
        self.page.ele(F'{henan_ele_dict.get("sxz_button")}').click()
        self.page.wait
        time.sleep(3)

        #
        table0 = self.page.get_tab(0)
        try:
            self.run_3_8(table0)
        except:
            pass
        try:
            self.run_3_11(table0)
        except:
            pass
        try:
            self.run_4_1(table0)
        except:
            pass
        try:
            self.run_4_3(table0)
        except:
            pass
        try:
            self.run_4_4(table0)
        except:
            pass

        try:
            self.exit_username_sxz(table0)
            time.sleep(5)
            self.exit_username_oms(table0)
            try:
                self.page.quit()
                time.sleep(1)
                print("网页退出！")
            except:
                pass
            return 1
        except Exception as e:
            print(f'运行失败！{e}')
            return 0

    def run_3_8(self, table0):
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3")}').click()  # 调度考核管理 3
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_8")}').click()  # 风光功率预测考核日考核数据汇总 3-8
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_8_L111")}').click()  # 点击风场
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_8_L111_st")}').input(self.today_1)
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_8_L111_et")}').input(self.today)
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_8_L111")}').click()  # 点击风场
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_8_L111_search")}').click()
        # iframe_fg = self.page('风光功率预测考核日考核数据汇总')
        try:
            ddglkh_button3_8_L111_data = table0.ele(
                F'{henan_ele_dict.get("ddglkh_button3_8_L111_data")}').text.split('\n')
            ddglkh_button3_8_L111_data[0] = ddglkh_button3_8_L111_data[0].replace(" 00:00:00.0", "")
            for i in range(len(ddglkh_button3_8_L111_data)):
                ddglkh_button3_8_L111_data[i] = ddglkh_button3_8_L111_data[i].replace('\t', '').strip()
        except:
            ddglkh_button3_8_L111_data = [self.today_1, self.wfname,
                                          '', '', '', '', '', '', '', '']

        MC = MysqlCurd()
        table_name = F'data_sxz_ddgl_fgflyckh_rkhsjhz'
        field_name = ['check_data', 'check_wfname',
                      'today_reports_num', 'today_check_elect', 'today_check_acc', 'today_predicte_elect',
                      'now_reports_num', 'now_check_elect', 'now_check_acc', 'now_predicte_elect']
        delect_run_3_8_sql = F"delete from {table_name}  WHERE check_data ='{self.today_1}' and check_wfname='{self.wfname}'"
        MC.delete(delect_run_3_8_sql)

        MC.insert_list(table_name, field_name, ddglkh_button3_8_L111_data)
        res_4 = float(ddglkh_button3_8_L111_data[4])
        res__2 = float(ddglkh_button3_8_L111_data[-2])
        if res_4 < 80 and res__2 < 85:
            self.run_3_8_dingding(ddglkh_button3_8_L111_data)

    def run_3_8_dingding(self, table_list):
        DAT = DingApiTools(appkey_value=self.appkey, appsecret_value=self.appsecret, chatid_value=self.chatid)
        from DingDingMarkDown.HenanDingDingMarkDown import henan_run_3_8_dict_false

        from datetime import datetime, timedelta
        yesterday_time = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        img_path = F"..{os.sep}Image{os.sep}save_sxz{os.sep}{yesterday_time}{os.sep}"
        directory = os.path.dirname(img_path)
        if not os.path.exists(directory):
            os.makedirs(directory)

        from PIL import ImageGrab
        im = ImageGrab.grab()
        save_wind_wfname_sxq = F"{img_path}{self.wfname}_异常_风光功率预测考核日考核数据汇总.png"
        im.save(save_wind_wfname_sxq)
        wfname_append = henan_run_3_8_dict_false["markdown"]["text"] + F'{self.wfname}'
        henan_run_3_8_dict_false["markdown"]["text"] = wfname_append
        DAT.push_message(self.jf_token, henan_run_3_8_dict_false)
        DAT.send_file(save_wind_wfname_sxq, 0)
        print('上报了了！')

    def run_3_11(self, table0):
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_11")}').click()  # 有功功率变化日考核结果 3-11
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_11_L111_st")}').input(F'{self.today_1} \ue007')
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_11_L111_et")}').input(self.today)
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_11_L111")}').click()  # 选择电场
        table0.ele(F'{henan_ele_dict.get("ddglkh_button3_11_L111_st")}').input(F'{self.today_1}  \ue007')
        table0._wait
        try:
            ddglkh_button3_11_L111_data = table0.ele(
                F'{henan_ele_dict.get("ddglkh_button3_11_L111_data")}').text.split('\n')
            for i in range(len(ddglkh_button3_11_L111_data)):
                ddglkh_button3_11_L111_data[i] = ddglkh_button3_11_L111_data[i].replace('\t', '').strip()
        except:
            ddglkh_button3_11_L111_data = [self.today_1, self.wfname, '', '', ]
        MC = MysqlCurd()
        table_name = F'data_sxz_ddgl_ygglbh_rkhjg'
        field_name = ['check_data', 'check_wfname',
                      'one_check_elect', 'ten_check_elect']
        delect_run_3_8_sql = F"delete from {table_name}  WHERE check_data ='{self.today_1}' and check_wfname='{self.wfname}'"
        MC.delete(delect_run_3_8_sql)

        MC.insert_list(table_name, field_name, ddglkh_button3_11_L111_data)
        res_2 = float(ddglkh_button3_11_L111_data[2])
        res_3 = float(ddglkh_button3_11_L111_data[3])
        if res_2 > 0 and res_3 > 0:
            self.run_3_11_dingding(ddglkh_button3_11_L111_data)

    def run_3_11_dingding(self, table_list):
        DAT = DingApiTools(appkey_value=self.appkey, appsecret_value=self.appsecret, chatid_value=self.chatid)
        from DingDingMarkDown.HenanDingDingMarkDown import henan_run_3_8_dict_false

        from datetime import datetime, timedelta
        yesterday_time = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        img_path = F"..{os.sep}Image{os.sep}save_sxz{os.sep}{yesterday_time}{os.sep}"
        directory = os.path.dirname(img_path)
        if not os.path.exists(directory):
            os.makedirs(directory)

        from PIL import ImageGrab
        im = ImageGrab.grab()
        save_wind_wfname_sxq = F"{img_path}{self.wfname}_异常_有功功率变化日考核结果.png"
        im.save(save_wind_wfname_sxq)
        from DingDingMarkDown.HenanDingDingMarkDown import henan_run_3_11_dict_false
        wfname_append = henan_run_3_11_dict_false["markdown"]["text"] + F'{self.wfname}'
        henan_run_3_11_dict_false["markdown"]["text"] = wfname_append
        DAT.push_message(self.jf_token, henan_run_3_11_dict_false)
        DAT.send_file(save_wind_wfname_sxq, 0)
        print('异常了！')

    def run_4_1(self, table0):
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4")}').click()  # 技术考核管理 4
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_1")}').click()  # 动态无功装置补偿可用率 4-1
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_1_L111_st")}').input(F'{self.today_1}  \ue007')
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_1_L111_et")}').input(self.today)
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_1_L111")}').click()  # 点击风场
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_1_L111_search")}').click()
        try:
            jsglkh_button4_1_L111_data = table0.ele(
                F'{henan_ele_dict.get("jsglkh_button4_1_L111_data")}').text.split('\n')
            for i in range(len(jsglkh_button4_1_L111_data)):
                jsglkh_button4_1_L111_data[i] = jsglkh_button4_1_L111_data[i].replace('\t', '').strip()
        except:
            jsglkh_button4_1_L111_data = [self.today_1, self.wfname, '']
        MC = MysqlCurd()
        table_name = F'data_sxz_jsglkh_dtwgbczzkhjg'
        field_name = ['check_data', 'check_wfname', 'operation_rate']

        delect_run_3_8_sql = F"delete from {table_name}  WHERE check_data ='{self.today_1}' and check_wfname='{self.wfname}'"
        MC.delete(delect_run_3_8_sql)

        MC.insert_list(table_name, field_name, jsglkh_button4_1_L111_data)
        res_2 = float(jsglkh_button4_1_L111_data[-1])
        if res_2 < 98:
            self.run_4_1_dingding(jsglkh_button4_1_L111_data)

    def run_4_1_dingding(self, table_list):
        DAT = DingApiTools(appkey_value=self.appkey, appsecret_value=self.appsecret, chatid_value=self.chatid)

        from datetime import datetime, timedelta
        yesterday_time = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        img_path = F"..{os.sep}Image{os.sep}save_sxz{os.sep}{yesterday_time}{os.sep}"
        directory = os.path.dirname(img_path)
        if not os.path.exists(directory):
            os.makedirs(directory)

        from PIL import ImageGrab
        im = ImageGrab.grab()
        save_wind_wfname_sxq = F"{img_path}{self.wfname}_异常_AGC投运率.png"
        im.save(save_wind_wfname_sxq)
        from DingDingMarkDown.HenanDingDingMarkDown import henan_run_4_1_dict_false
        wfname_append = henan_run_4_1_dict_false["markdown"]["text"] + F'{self.wfname}'
        henan_run_4_1_dict_false["markdown"]["text"] = wfname_append
        DAT.push_message(self.jf_token, henan_run_4_1_dict_false)
        DAT.send_file(save_wind_wfname_sxq, 0)
        print('异常了！')

    def run_4_3(self, table0):
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_3")}').click()  # 有功调节能力考核 4-3
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_3_1")}').click()  # acg 投运率
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_3_1_L1")}').click()  # 点击风场
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_3_1_R2")}').click()  # 日投运率
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_3_1_R2_st")}').input(F'{self.today_1}  \ue007')
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_3_1_R2_et")}').input(self.today)
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_3_1_R2_search")}').click()
        try:
            jsglkh_button4_3_1_R2_data = table0.ele(
                F'{henan_ele_dict.get("jsglkh_button4_3_1_R2_data")}').text.split('\n')[1:]
            for i in range(len(jsglkh_button4_3_1_R2_data)):
                jsglkh_button4_3_1_R2_data[i] = jsglkh_button4_3_1_R2_data[i].replace('\t', '').strip()
        except:
            jsglkh_button4_3_1_R2_data = [self.today_1, self.wfname, '']
        MC = MysqlCurd()
        table_name = F'data_sxz_jsglkh_ygtjnlkh_agctyl'
        field_name = ['check_data', 'check_wfname', 'operation_rate']
        delect_run_3_8_sql = F"delete from {table_name}  WHERE check_data ='{self.today_1}' and check_wfname='{self.wfname}'"
        MC.delete(delect_run_3_8_sql)

        MC.insert_list(table_name, field_name, jsglkh_button4_3_1_R2_data)
        res_2 = float(jsglkh_button4_3_1_R2_data[-1])
        if res_2 < 96:
            self.run_4_3_dingding(jsglkh_button4_3_1_R2_data)

    def run_4_3_dingding(self, table_list):
        DAT = DingApiTools(appkey_value=self.appkey, appsecret_value=self.appsecret, chatid_value=self.chatid)

        from datetime import datetime, timedelta
        yesterday_time = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        img_path = F"..{os.sep}Image{os.sep}save_sxz{os.sep}{yesterday_time}{os.sep}"
        directory = os.path.dirname(img_path)
        if not os.path.exists(directory):
            os.makedirs(directory)

        from PIL import ImageGrab
        im = ImageGrab.grab()
        save_wind_wfname_sxq = F"{img_path}{self.wfname}_异常_acg日合格率.png"
        im.save(save_wind_wfname_sxq)
        from DingDingMarkDown.HenanDingDingMarkDown import henan_run_4_3_dict_false
        wfname_append = henan_run_4_3_dict_false["markdown"]["text"] + F'{self.wfname}'
        henan_run_4_3_dict_false["markdown"]["text"] = wfname_append
        DAT.push_message(self.jf_token, henan_run_4_3_dict_false)
        DAT.send_file(save_wind_wfname_sxq, 0)

    def run_4_4(self, table0):
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_4")}').click()  # 并网电压曲线考核结果 4-4
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_4_T3")}').click()  # 日合格率结果
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_4_T3_st")}').input(F'{self.today_1}  \ue007')
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_4_T3_et")}').input(self.today)
        # table0.ele(F'{henan_ele_dict.get("jsglkh_button41_L111")}').click()  # 点击风场
        table0.ele(F'{henan_ele_dict.get("jsglkh_button4_4_T3_search")}').click()
        # table0.close()
        try:
            jsglkh_button4_4_T3_data = table0.ele(F'{henan_ele_dict.get("jsglkh_button4_4_T3_data")}').text.split(
                '\n')[1:]
            for i in range(len(jsglkh_button4_4_T3_data)):
                jsglkh_button4_4_T3_data[i] = jsglkh_button4_4_T3_data[i].replace('\t', '').strip()
        except:
            jsglkh_button4_4_T3_data = [self.today_1, self.wfname,
                                        '', '', '', '', '']
        MC = MysqlCurd()
        table_name = F'data_sxz_ddgl_bwdyqxkhjg_rhgljg'
        field_name = ['check_data', 'check_wfname', 'generatrix_name', 'qualified_time',
                      'unqualified_time', 'count_time', 'rate']

        delect_run_3_8_sql = F"delete from {table_name}  WHERE check_data ='{self.today_1}' and check_wfname='{self.wfname}'"
        MC.delete(delect_run_3_8_sql)

        MC.insert_list(table_name, field_name, jsglkh_button4_4_T3_data)
        res_2 = float(jsglkh_button4_4_T3_data[-1])
        if res_2 < 99.9:
            self.run_4_4_dingding(jsglkh_button4_4_T3_data)

    def run_4_4_dingding(self, table_list):
        DAT = DingApiTools(appkey_value=self.appkey, appsecret_value=self.appsecret, chatid_value=self.chatid)

        from datetime import datetime, timedelta
        yesterday_time = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        img_path = F"..{os.sep}Image{os.sep}save_sxz{os.sep}{yesterday_time}{os.sep}"
        directory = os.path.dirname(img_path)
        if not os.path.exists(directory):
            os.makedirs(directory)

        from PIL import ImageGrab
        im = ImageGrab.grab()
        save_wind_wfname_sxq = F"{img_path}{self.wfname}_异常_并网电压曲线考核结果.png"
        im.save(save_wind_wfname_sxq)
        from DingDingMarkDown.HenanDingDingMarkDown import henan_run_4_4_dict_false
        wfname_append = henan_run_4_4_dict_false["markdown"]["text"] + F'{self.wfname}'
        henan_run_4_4_dict_false["markdown"]["text"] = wfname_append
        DAT.push_message(self.jf_token, henan_run_4_4_dict_false)
        DAT.send_file(save_wind_wfname_sxq, 0)
        print('异常了！')

    def exit_username_login(self):
        try:
            res = self.page.ele('x://*[@id="app"]/section/header/div/div[2]/div/div/span').click()
        except  Exception as e:
            print(F'没有退出选项---{e}')
            return

        if res:
            self.page.ele('x:/html/body/ul/li[1]/span').click()
            self.page.ele('x:/html/body/div[2]/div/div[3]/button[2]').click()

        else:
            return

    def exit_username_sxz(self, table0):
        """
        用于退出sxz系统
        :param table0:
        :return:
        """
        table0.ele('x://*[@id="app"]/div/div[2]/div/div[1]/div[3]/div[2]/div/span').click()
        table0.ele('x://html/body/ul/li/span').click()
        table0.ele('x://html/body/div[8]/div/div[3]/button[2]/span').click()
        table0.close()
        time.sleep(1)

    def exit_username_oms(self, table0):
        """
        用于退出oms系统
        并清除缓存
        :param table0:
        :return:
        """
        # tab_id = self.page.find_tabs(title='调度管理系统')
        # table0 = self.page.get_tab(tab_id)
        tab_id = self.page.tabs[0]
        table0 = self.page.get_tab(tab_id)

        table0.ele('x:/html/body/div/section/header/div/div[2]/div/div/span').click()
        table0.ele('x:/html/body/ul/li[1]/span').click()
        table0.ele('x:/html/body/div[2]/div/div[3]/button[2]/span').click()
        table0.close()
        time.sleep(1)


def run_henan_oms_sxz_demo():
    for i in range(3):
        close_chrome()
        try:
            report_li = ReadyLogin().change_usbid()
            print(F'上报场站:{report_li}\n')
        except Exception as e:
            print(f'异常,暂停30s-{e}')
            pass


def close_chrome():
    page = RunSxz().page
    page.get('https://www.baidu.com')

    try:
        page.quit()
    except Exception as e:
        print(f'退出百度失败!@{e}')
        pass


if __name__ == '__main__':
    # run_henan_oms_sxz_demo()

    print(F"自动化程序填报运行中,请勿关闭!")
    # print(F"保佑,保佑,正常运行!")
    schedule.every().day.at("03:20").do(run_henan_oms_sxz_demo)
    schedule.every().day.at("16:10").do(run_henan_oms_sxz_demo)
    while True:
        schedule.run_pending()

        time.sleep(1)
