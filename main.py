# This is a sample Python script.
#in the code-change
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import datetime
import smtplib
import os
os.chdir(r"C:\Users\Sheeba\PycharmProjects\BrithdayWisher")
os.mkdir("testing")

# Press the green button in the gutter to run the script.
#enter the gmail id and password
GMAIL_ID= ''
GMAIL_PSWD= ''
def sendEmail(to,sub,msg):
    print(f"Email to {to} sent with subject: {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to ,f"Subjectt : {sub}\n\n{msg}")
    s.quit()

if __name__ == '__main__':
    df=pd.read_excel("birthday.xlsx")
    #print(df)
    today = datetime.datetime.now().strftime("%d-%m")#today is string format now
    yearNow=datetime.datetime.now().strftime("%Y")
    #print(yearNow)
    writeInd=[]
    #print(today)
    for index, item in df.iterrows():
        #print(index, item['Birthday'])
        #item['Birthday'] = datetime.datetime.strptime(item['Birthday'], '%Y-%m-%d %H:%M:S.%f')
        bday = item['Birthday'].strftime("%d-%m")
        #print(bday)
        #print(type(item['Year']))
        if(today == bday) and yearNow not in str(item['Year']):
            sendEmail(item['Email'], "Happy birtday", item['Dailogue'])
            writeInd.append(index)
    #print(writeInd)
    for i in writeInd:
        yr= df.loc[i,"Year"]
        df.loc[i, 'Year'] = str(yr) + ',' + yearNow
        #print(df.loc[i, 'Year'])
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
    #print(df)
    df.to_excel('birthday.xlsx',index=False)
