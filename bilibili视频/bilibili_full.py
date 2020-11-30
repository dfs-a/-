"""
思路:
1,首先获取分享链接
2,在浏览器中搜索链接
3,在html当中嵌套了js代码
4,仔细观察js代码,在playUrlInfo对象当中有源链接
"""

import os,re
from requests_html import HTMLSession
Requests = HTMLSession()
from jsonpath import jsonpath
#链接:https://www.bilibili.com/video/BV1uE411375k
class BiliBili:
    def __init__(self,url):
        self.start_url = url
        self.headers = {
            'user-agent':'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Mobile Safari/537.36 Edg/87.0.664.41'
        }

    def parse_url(self):
        page_html = Requests.get(self.start_url,headers=self.headers).html
        page_json = re.findall(r'window.__INITIAL_STATE__=(.+?)};', page_html.html)[0]
        source_title = page_html.xpath('//*[@id="app"]/div/div/div[4]/div[1]/div[1]/h1/text()')[0]
        source_url = re.findall(r'"url":"(.*?)"',page_json)[0]
        source_url = source_url.encode('latin-1').decode('unicode-escape')
        print(source_title,source_url)
        print("-----等待-----")
        return source_title,source_url

    def save_data(self,source_title,source_url):
        os_path = os.getcwd() + '/' + 'BiliBili' + '/'
        if not os.path.exists(os_path):
            os.mkdir(os_path)
        tv_data = Requests.get(url=source_url,headers=self.headers).content
        with open(os_path+source_title+'.mp4', 'wb') as f:
            f.write(tv_data)
        print("{}已保存".format(source_title))

    def run(self):
        source_title,source_url = self.parse_url()
        self.save_data(source_title,source_url)

if __name__ == '__main__':
    try:
        print("无水印并不是完全无水印,如果视频原作者自己添加了水印此脚本是无法去除的！")
        url = input("请输入b站分享链接:")
        bilbili = BiliBili(url)
        bilbili.run()
    except:
        print("检查是否为链接问题,链接没问题出错请联系开发者!")
