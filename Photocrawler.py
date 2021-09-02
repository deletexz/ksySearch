from bs4 import BeautifulSoup
import requests

if __name__ == '__main__':
    url = 'https://t.bilibili.com/?spm_id_from=333.788.b_696e7465726e6174696f6e616c486561646572.54'
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"}
    for i in range(0,100):
        req1 = requests.get(url, headers=headers)
        try:
            f = open(str(i)+"jpg", 'wb')
            f.write(req1.content)
            f.close()
        except:
            print("some error")