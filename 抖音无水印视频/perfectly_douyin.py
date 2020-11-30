import re
import os
from requests_html import HTMLSession
#实例化请求对象
Request = HTMLSession()
from jsonpath import jsonpath


class douyin_all:
    def __init__(self,start_url)->None:
        """
        :param start_url:首先短视频路径获取
        """
        self.url = start_url
        self.json_url = 'https://www.iesdouyin.com/web/api/v2/aweme/iteminfo/?item_ids={}'
        self.headers = {
            'User-Agent':'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Mobile Safari/537.36'
        }

    def home_page(self)->str:
        """
        1，获取短视频页面路径
        2，视频id值
        :return: 路径 and id值
        """
        content_url = Request.get(url=self.url,headers=self.headers).url
        source_json_code = re.findall('video/(.*?)/',content_url)[0]

        return content_url,source_json_code


    def json_source(self,content_url,source_json_code)->dict:
        """
        :param content_url:反爬中的refer字段路径
        :param source_json_code: 拼接获取json数据的url
        :return: 视频的json数据
        """
        headers_1 = {
            'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Mobile Safari/537.36',
            'Refer': '{}'.format(content_url)
        }
        js_url = self.json_url.format(source_json_code)

        dy_json_data = Request.get(url=js_url,headers=headers_1).json()
        return dy_json_data

    def analysis_data(self,dy_json_data)->str:
        """
        :param dy_json_data:通过字典取值获取title and url
        :return:
        """
        mp4_title = dy_json_data['item_list'][0]['share_info']['share_title']
        mp4_url = dy_json_data['item_list'][0]['video']['play_addr']['url_list'][0]
        mp4_url_1,mp4_url_2 = mp4_url.split('wm')
        mp4_url = mp4_url_1+mp4_url_2
        print(mp4_title,mp4_url)
        return mp4_title,mp4_url

    def douyin_save(self,mp4_title,mp4_url):
        os_path = os.getcwd() + '/' + "抖音短视频" + '/'
        if not os.path.exists(os_path):
            os.mkdir(os_path)
        data = Request.get(url=mp4_url,headers=self.headers).content
        with open(os_path+mp4_title+'.mp4','wb') as f:
            f.write(data)
            print("{}---->已保存成功".format(mp4_title))

    def run(self):
        content_url, source_json_code = self.home_page()
        dy_json_data = self.json_source(content_url, source_json_code)
        mp4_title,mp4_url = self.analysis_data(dy_json_data)
        self.douyin_save(mp4_title,mp4_url)

if __name__ == '__main__':
    try:
        url = input('请输入抖音分享链接:')
        url_1 = re.search('https://.*\..*?\.com/.*?/', url)
        start_url = url_1.group()
        douyin = douyin_all(start_url)
        douyin.run()
    except:
        print('请确认你的分享链接是否正确')

