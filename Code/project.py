from selenium import webdriver
import datetime
import math
import time
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart  # sending email
from email.mime.text import MIMEText  # constructing messages
from email.mime.base import MIMEBase
from email import encoders
today_run_date = datetime.datetime.now().strftime("%d-%m-%Y")
print("today date is", str(today_run_date))
pre_open = "pre_open_market"

driver = webdriver.Chrome(
    executable_path="/home/asuralts1/Temp/chromedriver")
driver.implicitly_wait(5)

driver.get(
    "https://www.nseindia.com/get-quotes/equity?symbol=QUESS#info-preopenmkt")
driver.implicitly_wait(10)
driver.maximize_window()
time.sleep(15)

PREV_CLOSE_QUESS = driver.find_element_by_xpath(
    '//*[@id="priceInfoTable"]/tbody/tr/td[1]').text
driver.implicitly_wait(5)
print("PREV CLOSE VALUE of quess==", PREV_CLOSE_QUESS.replace(",", ""))

OPEN_VALUE_QUESS = driver.find_element_by_xpath(
    '//*[@id="priceInfoTable"]/tbody/tr/td[2]').text
driver.implicitly_wait(5)
print("OPEN PRICE VALUE of quess==", OPEN_VALUE_QUESS.replace(",", ""))
BUY_QTY_QUESS = driver.find_element_by_xpath(
    '//*[@id="preOpenAto"]/tbody/tr[2]/td[1]').text
driver.implicitly_wait(5)
print("buy qty price of quess==", BUY_QTY_QUESS.replace(",", ""))
SELL_QTY_QUESS = driver.find_element_by_xpath(
    '//*[@id="preOpenAto"]/tbody/tr[2]/td[3]').text
driver.implicitly_wait(5)
print("sell qty value of quess==", SELL_QTY_QUESS.replace(",", ""))

Percentage_Change_Quess = ((float(OPEN_VALUE_QUESS.replace(",", ""))-float(
    PREV_CLOSE_QUESS.replace(",", "")))/float(PREV_CLOSE_QUESS.replace(",", "")))*100
pquess = float(Percentage_Change_Quess)
qqq = "{0:.2f}".format(pquess)
print(qqq)
time.sleep(10)
driver.close()

driver = webdriver.Chrome(
    executable_path="/home/asuralts1/Temp/chromedriver")
driver.implicitly_wait(5)

driver.get(
    "https://www.nseindia.com/get-quotes/equity?symbol=TEAMLEASE#info-preopenmkt")
driver.maximize_window()
time.sleep(15)

PREV_CLOSE_TEAM = driver.find_element_by_xpath(
    '//*[@id="priceInfoTable"]/tbody/tr/td[1]').text
driver.implicitly_wait(5)
print("PREV CLOSE VALUE of team is", PREV_CLOSE_TEAM.replace(",", ""))

OPEN_VALUE_TEAM = driver.find_element_by_xpath(
    '//*[@id="priceInfoTable"]/tbody/tr/td[2]').text
driver.implicitly_wait(5)
print("open value of team", OPEN_VALUE_TEAM.replace(",", ""))
BUY_QTY_TEAM = driver.find_element_by_xpath(
    '//*[@id="preOpenAto"]/tbody/tr[2]/td[1]').text
driver.implicitly_wait(6)
print("buy qty team is", BUY_QTY_TEAM.replace(",", ""))

SELL_QTY_TEAM = driver.find_element_by_xpath(
    '//*[@id="preOpenAto"]/tbody/tr[2]/td[3]').text
driver.implicitly_wait(5)
print("sell qty team is", SELL_QTY_TEAM.replace(",", ""))
Percentage_Change_team = ((float(OPEN_VALUE_TEAM.replace(",", ""))-float(
    PREV_CLOSE_TEAM.replace(",", "")))/float(PREV_CLOSE_TEAM.replace(",", "")))*100
pteam = float(Percentage_Change_team)
ttt = "{0:.2f}".format(pteam)
print(ttt)
time.sleep(10)
driver.close()

driver = webdriver.Chrome(
    executable_path="/home/asuralts1/Temp/chromedriver")
driver.implicitly_wait(5)

driver.get("https://www.nseindia.com/get-quotes/equity?symbol=SIS#info-preopenmkt")
driver.implicitly_wait(5)
driver.maximize_window()
time.sleep(15)

PREV_CLOSE_sis = driver.find_element_by_xpath(
    '//*[@id="priceInfoTable"]/tbody/tr/td[1]').text
driver.implicitly_wait(5)
print("PREV CLOSE VALUE of sis is", PREV_CLOSE_sis.replace(",", ""))

OPEN_VALUE_sis = driver.find_element_by_xpath(
    '//*[@id="priceInfoTable"]/tbody/tr/td[2]').text
driver.implicitly_wait(5)
print("open value of sis", OPEN_VALUE_sis.replace(",", ""))
BUY_QTY_sis = driver.find_element_by_xpath(
    '//*[@id="preOpenAto"]/tbody/tr[2]/td[1]').text
driver.implicitly_wait(10)
print("buy qty sis is", BUY_QTY_sis.replace(",", ""))
SELL_QTY_sis = driver.find_element_by_xpath(
    '//*[@id="preOpenAto"]/tbody/tr[2]/td[3]').text
driver.implicitly_wait(5)
print("sell qty sis is", SELL_QTY_sis.replace(",", ""))
Percentage_Change_sis = ((float(OPEN_VALUE_sis.replace(",", ""))-float(
    PREV_CLOSE_sis.replace(",", "")))/float(PREV_CLOSE_sis.replace(",", "")))*100
pteam = float(Percentage_Change_sis)
sss = "{0:.2f}".format(pteam)
print(sss)
driver.close()

file_path = open("/home/asuralts1/Temp/Missile/dashboard_template.txt")
data = file_path.read()
print(data)


def substitute(key, value):
    return data.replace(key, str(value))


def replace_cal_values(key, value):
    switcher = {
        '(?buy_quant_quess)': substitute(key, value),
        '(?sell_quantity_quess)': substitute(key, value),
        '(?previous_close_quess)': substitute(key, value),
        '(?open_price_quess)': substitute(key, value),
        '(?precentage_change_quess)': substitute(key, value),
        '(?buy_quant_teamlease)': substitute(key, value),
        '(?sell_quantity_teamlease)': substitute(key, value),
        '(?previous_close_teamlease)': substitute(key, value),
        '(?open_price_teamlease)': substitute(key, value),
        '(?precentage_change_teamlease)': substitute(key, value),
        '(?buy_quant_sis)': substitute(key, value),
        '(?sell_quantity_sis)': substitute(key, value),
        '(?previous_close_sis)': substitute(key, value),
        '(?open_price_sis)': substitute(key, value),
        '(?precentage_change_sis)': substitute(key, value),
        '(?todays_date)': substitute(key, value),
        '(?process_details)': substitute(key, value),

    }
    return switcher.get(key, "nothing")


data = replace_cal_values("(?buy_quant_quess)", BUY_QTY_QUESS)
data = replace_cal_values("(?previous_close_quess)", PREV_CLOSE_QUESS)
data = replace_cal_values("(?open_price_quess)", OPEN_VALUE_QUESS)
data = replace_cal_values("(?sell_quantity_quess)", SELL_QTY_QUESS)
data = replace_cal_values("(?precentage_change_quess)", qqq)
data = replace_cal_values("(?previous_close_teamlease)", PREV_CLOSE_TEAM)
data = replace_cal_values("(?open_price_teamlease)", OPEN_VALUE_TEAM)
data = replace_cal_values("(?buy_quant_teamlease)", BUY_QTY_TEAM)
data = replace_cal_values("(?sell_quantity_teamlease)", SELL_QTY_TEAM)
data = replace_cal_values("(?precentage_change_teamlease)", ttt)
data = replace_cal_values("(?previous_close_sis)", PREV_CLOSE_sis)
data = replace_cal_values("(?open_price_sis)", OPEN_VALUE_sis)
data = replace_cal_values("(?buy_quant_sis)", BUY_QTY_sis)
data = replace_cal_values("(?sell_quantity_sis)", SELL_QTY_sis)
data = replace_cal_values("(?precentage_change_sis)", sss)
data = replace_cal_values("(?todays_date)", today_run_date)
data = replace_cal_values("(?process_details)", pre_open)

print(data)

msg = MIMEMultipart()
html = data


subject = 'Pre - OpenMarket Details'
sender = 'mahinaidu141@gmail.com'
recipients = []
ex = pd.read_csv("/home/asuralts1/Temp/Missile/mails.csv")
head = ex.columns
for i, row in ex.iterrows():
    for j in head:
        recipients.append(row[j])
recipientscc = ['mahinaidu141@gmail.com']
BCC = []

msg['Subject'] = subject
msg['From'] = sender
msg['To'] = ", ".join(recipients)
msg['Cc'] = ", ".join(recipientscc)
# msg['Bcc'] = ", ".join(BCC)44444
# part1 = MIMEText(mailBody, 'plain')
part2 = MIMEText(html, 'html')
# msg.attach(part1)
msg.attach(part2)
s = smtplib.SMTP('smtp.office365.com', 587)
s.starttls()
s.login(sender, "Mahi141@")
recipients.extend(recipientscc)
recipients.extend(BCC)
s.sendmail(sender, recipients, msg.as_string())
print("Mail Sent")
s.quit()
