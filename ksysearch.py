import requests        #导入requests包
import json
import xlwt

work_book = xlwt.Workbook(encoding='utf-8')
sheet = work_book.add_sheet('sheet表名')
i=0

url_1 = 'http://app.wftvqcm.com/pc/search/ajaxSearch?keyword='

keyword = "%E6%9C%8D%E5%8A%A1%E4%BC%81%E4%B8%9A"




url = url_1+keyword+'&page='

for index in range(1,50):
    urls = url+str(index)
    strhtml = requests.get(urls)        #Get方式获取网页数据
    jsonzd = json.loads(strhtml.text)
    ksydata = jsonzd["data"]
    print("")

    for index in ksydata:
        title = index["title"]
        link = index["link"]
        time = index["published_des"]

        sheet.write(i, 0, title)
        sheet.write(i, 1, link)
        sheet.write(i, 2, time)
        i=i+1


work_book.save('Excel表.xls')



# 志愿服务