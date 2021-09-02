from bs4 import BeautifulSoup
import requests

if __name__ == '__main__':
    url = 'https://i0.hdslb.com/bfs/album/fcc3bf040d43fd46379d36c22b2a231ddf3aa9a9.jpg'
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"}
    for i in range(0,1):
        req1 = requests.get(url, headers=headers)
        try:
            f = open(str(i)+".jpg", 'wb')
            f.write(req1.content)
            f.close()
        except:
            print("some error")