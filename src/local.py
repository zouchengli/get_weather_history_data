#encoding=utf8
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import urllib2
from BeautifulSoup import BeautifulSoup
import xlsxwriter as xlw
cityCode = 'ZBAA' #This is Beijing China city code other city code please look www.wunderground.com docs
y = input('Please input year:')
m = input('Please input month:')
last_day = input('Please input month last day:')

if len(str(m)) < 2:  #format it
	mStamp = '0' + str(m)
else:
	mStamp = str(m)

timestamp = str(y) + '-' + mStamp

workbook = xlw.Workbook('beijing_weather_data_'+timestamp+'.xlsx')

for d in range(1,(last_day+1)):

    if len(str(d)) < 2: #format it
    	dStamp = '0' + str(d)
    else:
        dStamp = str(d)
    timestamp = str(y) + str('-') + str(m) + str('-') + str(d)

    sheet = workbook.add_worksheet(timestamp)
    title = ['Date(YMD)','Time (CST)', 'Temp.(°C)', 'Dew Point(°C)', 'Humidity', 'Pressure(hPa)', 'Visibility(km)', 'Wind Dir', 'Wind Speed(<km/h>/<m/s>)', 'Gust Speed(<km/h>/<m/s>)', 'Precip', 'Events', 'Conditions']
    for i in range(len(title)):
	sheet.write_string(0, i, title[i], workbook.add_format({'bold':True}))
    print "Getting data for " + timestamp
    url = "https://www.wunderground.com/history/airport/" + str(cityCode) + "/" + str(y) + "/" + str(m) + "/" + str(d) + "/DailyHistory.html"
    page = urllib2.urlopen(url)

    soup = BeautifulSoup(page)

    table = soup.findAll(attrs = {"class":"no-metars"});
    datas = [];
    for t in table:
	    tds = t.findAll('td');
	    data = [];
	    data.append(timestamp);
	    for td in tds:
		if (str(td.string) != str('None')):
			data.append(td.string);
		else:
			span = td.findAll('span');
			if (str(span) != str('[]')):
				high = span[1].string
				if (len(span) > 3):
					low = span[4].string
					if (str(low) != str('None')):
						high = high + '/' + low
				data.append(high);
	    datas.append(data);
	    print(data)
    row = 1;
    for ds in datas:
	 col = 0;
	 for d in ds:
	     if (str(d) == str('None')):
         	 d = '';
	     sheet.write_string(row, col, d)
	     col += 1
	 row += 1
workbook.close()