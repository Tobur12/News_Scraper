import urllib.request
import pandas as pd
import sys
import os
import time
import shutil
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
import openpyxl

user_agent = 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.7) Gecko/2009021910 Firefox/3.0.7'
headers={'User-Agent':user_agent,}
Text = pd.read_excel(os.path.dirname(sys.argv[0]) + '/Search.xlsx', index_col=None, sheet_name='Text')

src_dir = os.path.dirname(sys.argv[0]) + '/Output.xlsx'
dst_dir = 'C:/Users/mtomas/Desktop/News.xlsx'
shutil.copy(src_dir,dst_dir)
OutputPath = dst_dir
OutputFile = pd.read_excel('' + OutputPath + '', index_col=None, sheet_name='Output')
OutputCounter= len(OutputFile.index)

for index1, row1 in Text.iterrows():
    try:
        print("Processing: " + row1[0])
        url = "https://www.google.com/search?q="+ row1[0].replace(" ", "+") +"&tbm=nws&source=lnt&tbs=sbd:1&sa=X&dpr=0.9"
        url = url.lower()
        request=urllib.request.Request(url,None,headers)
        response = urllib.request.urlopen(request).read().decode('utf-8')
      
        response = response[response.find('<div class="ezO2md">'):]
        response = response[:response.find('<table class="gNEi4d">')]
        ResponseSeries = pd.Series([response])
        ResponseSeries = ResponseSeries.str.split(r'<div class="ezO2md">', expand=True)
        ResponseSeries = ResponseSeries.T
        ResponseSeries = ResponseSeries.tail(-1)

        for index, row in ResponseSeries.iterrows():

            ResultText = row[0]
            ResultText = ResultText[ResultText.find('fuLhoc ZWRArf">'):]
            ResultText = ResultText[:ResultText.find('</span>')]
            ResultText = ResultText.replace('fuLhoc ZWRArf">','')

            DateOfNews = row[0]
            DateOfNews = DateOfNews[DateOfNews.find('<span class="fYyStc YVIcad">'):]
            DateOfNews = DateOfNews[:DateOfNews.find('</span>')]
            DateOfNews = DateOfNews.replace('<span class="fYyStc YVIcad">','')

            URLText = row[0]
            URLText = URLText[URLText.find('<div><a href="/url?q='):]
            URLText = URLText[:URLText.find('"><div>')]
            URLText = URLText.replace('<div><a href="/url?q=','')

            if (("hodina" in DateOfNews) or ("včera" in DateOfNews) or (" dny" in DateOfNews) or ("týdnem" in DateOfNews)) and ".cz" in URLText:
                LearnTemporary = {'Search Text': [row1[0]], 'Result Text' : ResultText, 'Date of News' : DateOfNews, 'URL' : URLText}
                Learn = pd.DataFrame(data=LearnTemporary)
                Learn = Learn.applymap(lambda x: ILLEGAL_CHARACTERS_RE.sub(r'', x) if isinstance(x, str) else x)

                with pd.ExcelWriter('' + OutputPath + '', mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                    Learn.to_excel(writer, sheet_name='Output', startrow=OutputCounter+1)
                path = '' + OutputPath + ''
                book = openpyxl.load_workbook(path)
                sheet = book['Output']
                sheet.delete_rows(OutputCounter+2, 1)  
                book.save(path)
                OutputCounter=OutputCounter+1
        time.sleep(2)
    except:
        pass
