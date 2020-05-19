import pandas as pd
import datetime
import smtplib


GMAIL_ID = 'Enter the Gmail Id that you wish to send the message from'
GMAIL_PASSWORD = 'Include your Gmail Password'


def send_email(to_email_id, subject, message):
    # print(f"Email to '{to_email_id}' sent with subject '{subject}' with message '{message}' is sent"
         # f" successfully")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PASSWORD)

    s.sendmail(GMAIL_ID, to_email_id, f"Subject: {subject} \n\n "
                                      f"This is an auto generated mail from Raunak Bhagwani." 
                                      f"He's just testing his Python application." 
                                      f"Sorry if it disturbed you. {message}")
    s.quit()


if __name__ == '__main__':

    df = pd.read_excel('Final_data.xlsx')
    today = datetime.datetime.now().strftime("%d-%m")
    year_now = datetime.datetime.now().strftime("%Y")
    write_index = []

    for index, item in df.iterrows():
        bday = item['Birthday'].strftime("%d-%m")
        if today == bday and year_now not in str(item['Year']):
            send_email(item['Email_Id'], 'Happy Birthday', item['Dialogue'])
            write_index.append(index)

    if len(write_index) > 0:
        for i in write_index:
            yr = df.loc[i, 'Year']
            df.loc[i, 'Year'] = str(yr) + ', ' + str(year_now)
        df.to_excel('Final_data.xlsx', index = False)

