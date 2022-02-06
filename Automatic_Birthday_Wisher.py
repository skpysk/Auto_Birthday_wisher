import pandas as pd
import datetime
import smtplib

gmail_id  = "your_gmail.com"
gmail_psw = "your_password"

def sendemail(to,sub,msg):
    #print(f"email to {to} send subject : {sub} , or {msg}")
    s = smtplib.SMTP("smtp.gmail.com" , 587)
    s.starttls()
    s.login(gmail_id,gmail_psw)
    s.sendmail(gmail_id,to,f" subject : {sub}\n\n {msg}") # argument #1 from #2 to #3 msg
    s.quit

if __name__=="__main__":
    #sendemail(gmail_id,"happy birtday","test msg")
    data = pd.read_excel("data.xlsx") #Use pip or conda to install openpyxl
    today = datetime.datetime.now().strftime("%d-%m") # today date finder and use only date with month d day m month also Y year
    yearnow = datetime.datetime.now().strftime("%Y")
    #print(today)
    write = []
    for index, item in data.iterrows():    # data ka birtday wala index keform me dekhao
        #print(index, item['Birtday'])
        birtday = item['Birthday'].strftime("%d-%m")  # birtday ko yeh d-m ke format me output ker dega .str wala code
        #print(birtday)
        # ab ager excel ka koi topic lena he to usko itmem["topic"]  le sakte ho
        
        if ( today == birtday) and yearnow not in str(item["Year"]) :
            sendemail(item["Email"], "Happy Birtday" , item["Dialogue"])
            write.append(index)
    #print(write)
    for i in write:
        yr = data.loc[i, 'Year']
        data.loc[i, 'Year'] = str(yr) + ', ' + str(yearnow)
        #print(data.loc[i, 'Year'])
        #print(data)
        data.to_excel("data.xlsx",index=False)  # unames ayega bcz index =true he# data ko execel me (" this . xlsx") me daal do # permission denied matlove file pehle se open he khe per
        