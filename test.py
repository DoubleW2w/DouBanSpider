# -*- coding: utf-8 -*-
# @Time     : 20:07
# @Author   : DoubleL2l
# @File     : test.py
# @Software : PyCharm
import urllib.request

url = 'http://www.baidu.com/'
request = urllib.request.Request(url)
source_code = urllib.request.urlopen(request).read().decode()
plain_text=str(source_code)
print(source_code)
print("-" * 100)
print(plain_text)