from builtins import print
import datetime

a = ""
a += "var res = document.getElementsByClassName('fr-trigger-texteditor');"
a += "for(var i>=3;i<=4;i++){"
a += "res[i].removeAttribute('readonly')"
a += "}"

today = datetime.datetime.now()
lastDay = datetime.date(datetime.date.today().year,datetime.date.today().month,1)-datetime.timedelta(1)
last = datetime.date(datetime.date.today().year,datetime.date.today().month,1)-datetime.timedelta(1)
thisMonth = datetime.datetime(today.year, today.month, 1).strftime("%Y-%m-%d")
print(lastDay)




