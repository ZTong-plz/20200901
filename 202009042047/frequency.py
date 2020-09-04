import requests
import re
from config import infomation as ci
from handleexcel import HandlerExcel as he
import time
from pylab import *
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMainWindow, QApplication
from modeltwo import Ui_MainWindow
from ExportByRequests import ExportCSV
from multiprocessing import Pool

class MyDesiger(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyDesiger, self).__init__(parent)
        self.setupUi(self)
    def export_test(self):
        # 120181080727
        # 2349zzhZZH
        # zzh15910939170@163.com
        # zzh123456!
        #Exceeded 30 redirects.？
        export=ExportCSV()
        export.vpn_username=self.vu.text()
        export.vpn_password=self.vp.text()
        export.incite_username=self.iu.text()
        export.incite_passname=self.ip.text()
        #export.year_start=self.ys.currentText()
        #export.year_end=self.ye.currentText()
        export.year_start=self.ystart.text()
        export.year_end=self.yend.text()
        export.export()
        today = datetime.date.today().isoformat()
        self.textBrowser.setText(today)
        self.textBrowser.append(export.tip)
    def frequency_test(self):
        bbg.splitWoslist()
    def graph_polyline(self):
        #he().writeIn2(Dict)
        a=1
    def graph_columnar(self):
        b=2
class getfrequency1(object):

    vpn_username = "120181080118"
    vpn_password = "271517"
    ui=1
    handle=None
    def get_authenticity_token(self):
        try:
            s = requests.session()
            r = s.get('https://webvpn.ncepu.edu.cn')
        except Exception as e:
            print(e)
        else:
            reg = r'<input[\s]*type="hidden"[\s]*name="authenticity_token"[\s]*value="(.*)"[\s]*/[\s]*>'
            pattern = re.compile(reg)
            result = pattern.findall(r.content.decode('utf-8'))
            token = result[0]
            return token, s

    def getWebOfSciencePage(self):
        authen,ss = self.get_authenticity_token()
        url2 = 'https://www-webofknowledge-com.webvpn.ncepu.edu.cn'
        form_data = {
                'utf8': '✓',
                'authenticity_token': authen,
                'user[login]': '120181080118',  # 用户名
                'user[password]': '271517',  # 密码
                'user[dymatice_code]': 'unknown',
                'commit': '登录 Login'
            }
        #form_data['user[login]'] = self.vpn_username
        #form_data['user[password]'] = self.vpn_password

        try:
            res3 = ss.post(ci.signinvpn_url, data=form_data, headers=ci.signin_headers)
            res3.raise_for_status()
            res3.encoding = res3.apparent_encoding
            res = ss.get(url2, headers=ci.signin_headers)
            res.raise_for_status()
            res.encoding = res.apparent_encoding

            sid = re.search('SID=([^&]*)&?', res.url).group(1)
            final_url = 'https://apps-webofknowledge-com.webvpn.ncepu.edu.cn/WOS_GeneralSearch_input.do?product=WOS&SID='+sid+'&search_mode=GeneralSearch'
            res1 = ss.get(final_url, headers=ci.signin_headers)
            res1.raise_for_status()
            res1.encoding = res1.apparent_encoding
            # print(res1.url)


        except Exception as e:
            print(e)
        else:
            return ss, sid

    def splitWoslist(self):
            self.handle=he()
            pre_woslist = self.handle.getWOSList()
            final_woslist = []
            for i in range(0, len(pre_woslist), 20):
                final_woslist.append(pre_woslist[i:i+20])
            for a_list in final_woslist:
                self.loop(a_list)
            # print(len(final_woslist))

    def loop(self, a_list):
            ss, sid = self.getWebOfSciencePage()
            for wos in a_list:
                self.postwos(ss, sid, wos, 5)

    # def loop(self):
    #     po = Pool(3)
    #     self.handle = he()
    #     woslist =self.handle.getWOSList()
    #     # pre_woslist = self.handle.getWOSList()
    #     # final_woslist = []
    #     # for i in range(0, len(pre_woslist), 20):
    #     #     final_woslist.append(pre_woslist[i:i+20])
    #     # for a_list in final_woslist:
    #     #     self.loop(a_list)
    #     ss, sid = self.getWebOfSciencePage()
    #     for wos in woslist:
    #         po.apply_async(self.postwos, (ss, sid, wos, 5))
    #         # self.postwos(ss, sid, wos, 5)
    #     # self.postwos(ss, sid, '000339883200004', 5)
    #     po.close()
    #     po.join()
    #
    # def get_proxy_ip(self):
    #     requests.adapters.DEFAULT_RETRIES = 5
    #     requests_session = requests.session()
    #     requests_session.keep_alive = False
    #     try:
    #         proxy = requests_session.get(ci.PORXY_POOL).text
    #     except Exception as e:
    #         return {
    #             'code': 500,
    #             'message': '代理池异常',
    #             'error': str(e)
    #         }
    #     else:
    #         return {
    #             'code': 200,
    #             'message': 'OK',
    #             'proxy': 'http://{0}'.format(proxy)
    #         }

    def postwos(self, ss, sid, wos, count):
        Dict = {}
        Dict['wos'] = 'WOS:'+wos
        #deadline = '20200730'#???????????????????????????????????????
        deadline=self.ui.ddl.text()
        url = 'https://apps-webofknowledge-com.webvpn.ncepu.edu.cn/WOS_GeneralSearch.do'
        headers = {
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36',
            'Host': 'apps-webofknowledge-com.webvpn.ncepu.edu.cn',
            'Origin': 'https://apps-webofknowledge-com.webvpn.ncepu.edu.cn',
            'Referer': 'https://apps-webofknowledge-com.webvpn.ncepu.edu.cn/WOS_GeneralSearch_input.do?product=WOS&SID='+sid+'&search_mode=GeneralSearch',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'same-origin'
        }
        form_data = {
            'fieldCount': 1,
            'action': 'search',
            'product': 'WOS',
            'search_mode': 'GeneralSearch',
            'SID': sid,
            'max_field_count': 25,
            'formUpdated': 'true',
            'value(input1)': wos,
            'value(select1)': 'UT',
            'value(hidInput1)':'',
            'limitStatus': 'collapsed',
            'ss_lemmatization': 'On',
            'ss_spellchecking': 'Suggest',
            'SinceLastVisit_UTC': '',
            'SinceLastVisit_DATE': '',
            'range': 'ALL',
            'period': 'Range Selection',
            'startYear': '1985',
            'endYear': '2020',
            'update_back2search_link_param': 'yes',
            'ssStatus': 'display:none',
            'ss_showsuggestions': 'ON',
            'ss_numDefaultGeneralSearchFields': 1,
            'ss_query_language': 'auto',
            'rs_sort_by': 'PY.D;LD.D;SO.A;VL.D;PG.A;AU.A'
        }
        try:
            res = ss.post(url, data=form_data, headers=headers)
            res.raise_for_status()
            res.encoding = res.apparent_encoding
            articlename_reg = r'<value lang_id="">(.+)</value>'
            articlename_pattern = re.compile(articlename_reg)
            date_reg = r'<span[\s]*class=\"label\">(Published:|Early Access:)[\s\S]*?<value>(.+)</value>'
            date_pattern = re.compile(date_reg)
            articlename = articlename_pattern.search(res.content.decode('utf-8')).group(1)
            date = date_pattern.search(res.text).group(2)

            is_zero = re.search(r'class="search-results-data-cite">Times Cited: 0', res.text)
            if is_zero:
                pre_cites = '0'
            else:
                href = re.search(r'class="snowplow-times-cited-link" title="View all of the articles that cite this one" href="(.+?)"', res.text).group(1)
                cites_url1 = 'https://apps-webofknowledge-com-443.webvpn.ncepu.edu.cn'+href
                cites_url = cites_url1.replace('amp;', '')
                cites_res = ss.get(cites_url, headers=headers)
                cites_res.raise_for_status()
                # with open("D:/abc.txt", "wb") as f:
                #     f.write(cites_res.content)
                pre_cites = re.search(r'id="CAScorecard_count_WOSCLASSIC">[\s\S]*?>[\s]*([0-9]+)[\s]*</a>', cites_res.text).group(1)
                select_url1 = 'https://apps-webofknowledge-com.webvpn.ncepu.edu.cn/'+re.search(r'id="CAScorecard_count_WOSCLASSIC">[\s\S]*?<a href="([\s\S]+?)"', cites_res.text).group(1)
                select_url = select_url1.replace(' ', '')
                minus_res = ss.get(select_url)
                minus_res.raise_for_status()
        # print("pre_cites:", pre_cites)
        except Exception as e:
            if count > 0:
                self.postwos(ss, sid, wos, count-1)
            else:
                print(wos, e)
        else:
            #print(articlename)#输出信息1
            #self.ui.textBrowser.setText("articlename:"+articlename)
            self.ui.textBrowser.append("articlename:" + articlename)
            QApplication.processEvents()
            #print(date)#输出信息2
            self.ui.textBrowser.append("date:"+date)
            QApplication.processEvents()
            if pre_cites == '0':
                final_cites = 0
            else:
                final_cites = int(pre_cites)  # 还没剔除 时间不符的 总的被引频次
                cites = [int(pre_cites), 10][int(pre_cites) > 10]
                i = 1
                while i < cites+1:
                    reg = r'id="fetch_wos_subject_Span_'+str(i)+'"[\s\S]+?<span class="label">Published:[\s\S]+?<value>(.+?)</value>'
                    reg1 = r'<span id="early_access_month_'+str(i)+'" class="data_bold">(.+?)</span>'
                    if re.search(reg1, minus_res.text):
                        pre_published_time = re.search(reg1, minus_res.text).group(1)  # 得到待筛除文章的出版日期
                    else:
                        pre_published_time = re.search(reg, minus_res.text).group(1)
                    # 将published_time与统计截止日期进行比较
                    final_published_time = self.parseTime(pre_published_time)
                    if final_published_time == '0':
                        print("出错:", wos)
                        break
                    elif int(str(final_published_time)) - int(str(deadline)) > 0:
                        final_cites -= 1
                    else:
                        break
                    i += 1
                if i == 11 and int(pre_cites) > 10:
                    pagecount_top = re.search(r'<span id="pageCount.top">([0-9]+?)</span>', minus_res.text).group(1)
                    page = 2
                    next_page_url = re.search(r'<a class="paginationNext snowplow-navigation-nextpage-bottom"  href="([\s\S]+?)"', minus_res.text).group(1)
                    while page <= int(pagecount_top):
                        next_page_res = ss.get(next_page_url)
                        j = 1
                        while j < 11 and (page-1)*10+j <= int(pre_cites):
                            reg2 = r'id="fetch_wos_subject_Span_'+str((page-1)*10+j)+'"[\s\S]+?<span class="label">Published:[\s\S]+?<value>(.+?)</value>'
                            reg3 = r'<span id="early_access_month_'+str((page-1)*10+j)+'" class="data_bold">(.+?)</span>'
                            if re.search(reg3, next_page_res.text):
                                pre_published_time = re.search(reg3, next_page_res.text).group(1)  # 得到待筛除文章的出版日期
                            else:
                                pre_published_time = re.search(reg2, next_page_res.text).group(1)

                            # 将published_time与统计截止日期进行比较
                            final_published_time = self.parseTime(pre_published_time)
                            if final_published_time == '0':
                                print("出错:", wos)
                                break
                            elif int(str(final_published_time)) - int(str(deadline)) > 0:
                                final_cites -= 1
                            else:
                                break
                            j += 1
                        if j >= 11 and int(pre_cites) > page*10:
                            page += 1
                            next_page_url = re.search(r'<a class="paginationNext snowplow-navigation-nextpage-bottom"  href="([\s\S]+?)"', next_page_res.text).group(1)
                        else:
                            break
            #print("final_cites:", final_cites)#输出信息3
            self.ui.textBrowser.append("Cited frequency:"+(str)(final_cites))
            QApplication.processEvents()
            Dict['cites'] = final_cites
            Dict['date'] = date
            self.handle.writeInBySingle(Dict)
            #he().writeInBySingle(Dict)
    # def postwos(self, ss, sid, wos, count):
    #     # ss, sid = self.getWebOfSciencePage()
    #     proxies = {"http": self.get_proxy_ip().get('proxy')}
    #     # proxies = self.get_proxy()
    #     Dict = {}
    #     Dict['wos'] = 'WOS:' + wos
    #     #deadline = self.ui.ddl.text()
    #     deadline=20200831
    #     url = 'https://apps-webofknowledge-com.webvpn.ncepu.edu.cn/WOS_GeneralSearch.do'
    #     headers = {
    #         'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36',
    #         'Host': 'apps-webofknowledge-com.webvpn.ncepu.edu.cn',
    #         'Origin': 'https://apps-webofknowledge-com.webvpn.ncepu.edu.cn',
    #         'Referer': 'https://apps-webofknowledge-com.webvpn.ncepu.edu.cn/WOS_GeneralSearch_input.do?product=WOS&SID=' + sid + '&search_mode=GeneralSearch',
    #         'Sec-Fetch-Dest': 'document',
    #         'Sec-Fetch-Mode': 'navigate',
    #         'Sec-Fetch-Site': 'same-origin'
    #     }
    #     form_data = {
    #         'fieldCount': 1,
    #         'action': 'search',
    #         'product': 'WOS',
    #         'search_mode': 'GeneralSearch',
    #         'SID': sid,
    #         'max_field_count': 25,
    #         'formUpdated': 'true',
    #         'value(input1)': wos,
    #         'value(select1)': 'UT',
    #         'value(hidInput1)': '',
    #         'limitStatus': 'collapsed',
    #         'ss_lemmatization': 'On',
    #         'ss_spellchecking': 'Suggest',
    #         'SinceLastVisit_UTC': '',
    #         'SinceLastVisit_DATE': '',
    #         'range': 'ALL',
    #         'period': 'Range Selection',
    #         'startYear': '1985',
    #         'endYear': '2020',
    #         'update_back2search_link_param': 'yes',
    #         'ssStatus': 'display:none',
    #         'ss_showsuggestions': 'ON',
    #         'ss_numDefaultGeneralSearchFields': 1,
    #         'ss_query_language': 'auto',
    #         'rs_sort_by': 'PY.D;LD.D;SO.A;VL.D;PG.A;AU.A'
    #     }
    #     try:
    #         res = ss.post(url, data=form_data, headers=headers, proxies=proxies)
    #         res.raise_for_status()
    #         res.encoding = res.apparent_encoding
    #         articlename_reg = r'<value lang_id="">(.+)</value>'
    #         articlename_pattern = re.compile(articlename_reg)
    #         date_reg = r'<span[\s]*class=\"label\">(Published:|Early Access:)[\s\S]*?<value>(.+)</value>'
    #         date_pattern = re.compile(date_reg)
    #         articlename = articlename_pattern.search(res.content.decode('utf-8')).group(1)
    #         date = date_pattern.search(res.text).group(2)
    #
    #         is_zero = re.search(r'class="search-results-data-cite">Times Cited: 0', res.text)
    #         if is_zero:
    #             pre_cites = '0'
    #         else:
    #             href = re.search(r'class="snowplow-times-cited-link" title="View all of the articles that cite this one" href="(.+?)"',res.text).group(1)
    #             cites_url1 = 'https://apps-webofknowledge-com-443.webvpn.ncepu.edu.cn' + href
    #             cites_url = cites_url1.replace('amp;', '')
    #             cites_res = ss.get(cites_url, headers=headers, proxies=proxies)
    #             cites_res.raise_for_status()
    #
    #             pre_cites_iszero = re.search(r'id="CAScorecard_count_WOSCLASSIC">[\s]*0[\s]*</span>', cites_res.text)
    #             if pre_cites_iszero:
    #                 pre_cites = '0'
    #             else:
    #                 pre_cites = re.search(r'id="CAScorecard_count_WOSCLASSIC">[\s\S]*?>[\s]*([0-9]+)[\s]*</a>',cites_res.text).group(1)
    #
    #             print("pre_cites:", pre_cites)
    #     except Exception as e:
    #         if count > 0:
    #             self.postwos(ss, sid, wos, count - 1)
    #         else:
    #             print(wos, e)
    #     else:
    #         if pre_cites == '0':
    #             final_cites = 0
    #         else:
    #             select_url1 = 'https://apps-webofknowledge-com.webvpn.ncepu.edu.cn/' + re.search(r'id="CAScorecard_count_WOSCLASSIC">[\s\S]*?<a href="([\s\S]+?)"', cites_res.text).group(1)
    #             select_url = select_url1.replace(' ', '')
    #             minus_res = ss.get(select_url, proxies=proxies)
    #             minus_res.raise_for_status()
    #
    #             final_cites = int(pre_cites)  # 还没剔除 时间不符的 总的被引频次
    #             cites = [int(pre_cites), 10][int(pre_cites) > 10]
    #             i = 1
    #             while i < cites + 1:
    #                 reg = r'id="fetch_wos_subject_Span_' + str(
    #                     i) + '"[\s\S]+?<span class="label">Published:[\s\S]+?<value>(.+?)</value>'
    #                 reg1 = r'<span id="early_access_month_' + str(i) + '" class="data_bold">(.+?)</span>'
    #                 if re.search(reg1, minus_res.text):
    #                     pre_published_time = re.search(reg1, minus_res.text).group(1)  # 得到待筛除文章的出版日期
    #                 else:
    #                     pre_published_time = re.search(reg, minus_res.text).group(1)
    #                 # 将published_time与统计截止日期进行比较
    #                 final_published_time = self.parseTime(pre_published_time)
    #                 if final_published_time == '0':
    #                     print("出错:", wos)
    #                     break
    #                 elif int(str(final_published_time)) - int(str(deadline)) > 0:
    #                     final_cites -= 1
    #                 else:
    #                     break
    #                 i += 1
    #             if i == 11 and int(pre_cites) > 10:
    #                 pagecount_top = re.search(r'<span id="pageCount.top">([0-9]+?)</span>', minus_res.text).group(1)
    #                 page = 2
    #                 next_page_url = re.search(
    #                     r'<a class="paginationNext snowplow-navigation-nextpage-bottom"  href="([\s\S]+?)"',
    #                     minus_res.text).group(1)
    #                 while page <= int(pagecount_top):
    #                     next_page_res = ss.get(next_page_url)
    #                     j = 1
    #                     while j < 11 and (page - 1) * 10 + j <= int(pre_cites):
    #                         reg2 = r'id="fetch_wos_subject_Span_' + str((
    #                                                                             page - 1) * 10 + j) + '"[\s\S]+?<span class="label">Published:[\s\S]+?<value>(.+?)</value>'
    #                         reg3 = r'<span id="early_access_month_' + str(
    #                             (page - 1) * 10 + j) + '" class="data_bold">(.+?)</span>'
    #                         if re.search(reg3, next_page_res.text):
    #                             pre_published_time = re.search(reg3, next_page_res.text).group(1)  # 得到待筛除文章的出版日期
    #                         else:
    #                             pre_published_time = re.search(reg2, next_page_res.text).group(1)
    #
    #                         # 将published_time与统计截止日期进行比较
    #                         final_published_time = self.parseTime(pre_published_time)
    #                         if final_published_time == '0':
    #                             print("出错:", wos)
    #                             break
    #                         elif int(str(final_published_time)) - int(str(deadline)) > 0:
    #                             final_cites -= 1
    #                         else:
    #                             break
    #                         j += 1
    #                     if j >= 11 and int(pre_cites) > page * 10:
    #                         page += 1
    #                         next_page_url = re.search(
    #                             r'<a class="paginationNext snowplow-navigation-nextpage-bottom"  href="([\s\S]+?)"',
    #                             next_page_res.text).group(1)
    #                     else:
    #                         break
    #         self.write(articlename,date,final_cites)
            # self.ui.textBrowser.append("articlename:" + articlename)
            # QApplication.processEvents()
            # self.ui.textBrowser.append("date:" + date)
            # QApplication.processEvents()
            # self.ui.textBrowser.append("Cited frequency:" + (str)(final_cites))
            # QApplication.processEvents()
            # Dict['cites'] = final_cites
            # Dict['date'] = date
            # self.handle.writeInBySingle(Dict)
    # def write(self,a,b,c):
    #     self.ui.textBrowser.append("articlename:" + a)
    #     QApplication.processEvents()
    #     self.ui.textBrowser.append("date:" + b)
    #     QApplication.processEvents()
    #     self.ui.textBrowser.append("Cited frequency:" + (str)(c))
    #     QApplication.processEvents()
    def  bg(self):
                QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 保证qt designer预览界面和运行结果一致
                app = QApplication(sys.argv)
                self.ui = MyDesiger()
                self.ui.show()
                sys.exit(app.exec_())

    def parseTime(self, crawl_date):
        '''将日期格式化'''
        crawl_date = crawl_date.strip()
        #  对特殊日期进行处理
        if '.'in crawl_date:
            crawl_date = crawl_date.replace('.', '')
        if '-'in crawl_date:
            del_str = re.search(r'[a-zA-Z]{3}(.+?)[0-9\s]+', crawl_date).group(1)
            crawl_date = crawl_date.replace(del_str, '')

        num = len(crawl_date.split())-1
        try:
            if num == 0:  # 日期只有年:2020
                middle_time = time.mktime(time.strptime(crawl_date, "%Y"))
                final_time = time.strftime("%Y", time.localtime(middle_time))+'0000'
            elif num == 1:  # 日期有年月: Feb 2020
                middle_time = time.mktime(time.strptime(crawl_date, "%b %Y"))
                final_time = time.strftime("%Y%m", time.localtime(middle_time))+'00'
            elif num == 2:  # 日期有年月日：
                middle_time = time.mktime(time.strptime(crawl_date, "%b %d %Y"))
                final_time = time.strftime("%Y%m%d", time.localtime(middle_time))
            else:
                final_time = '0'
        except ValueError as ve:
            final_time = '0'
            print(ve)
        return final_time



if __name__ == '__main__':


    bbg=getfrequency1()
    bbg.bg()