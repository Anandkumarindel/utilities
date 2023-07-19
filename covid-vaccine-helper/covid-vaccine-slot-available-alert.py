import argparse
import sys
import pprint
import requests
from datetime import date
import json
import os
import pandas
import collections
import smtplib
from email.mime.multipart import MIMEMultipart
from email.message import Message
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase

# xls_filename = 'covid-vaccine-slot-available-requests-sample.xlsx'
xls_filename = sys.argv[1]
pandas.set_option('display.max_colwidth', -1)


rest_base_url = 'https://cdn-api.co-vin.in/api/v2/appointment/sessions/public/'

header={
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36",
}


def get_vaccine_json(code=110093,type='pincode'):
    today = date.today().strftime("%d-%m-%Y")
    if type == 'pincode':
        rest_url = rest_base_url + 'calendarByPin?pincode=' + str(code) + '&date=' + str(today)
    #print(rest_url)

    data = requests.get(rest_url, headers=header, verify=False)
    if data.status_code != 200:
        raise Exception("Rest Call failed")
    return json.loads(data.text)


# def parse_args():
#     parser = argparse.ArgumentParser(description="Alert vaccine slot available to given Pincode or district.")
#     city_code = parser.add_mutually_exclusive_group()
#     # city_code.add_argument("-p", "--pincode", type=int, default=110093, help="pincode of the city")
#     # city_code.add_argument("-d", "--district_id", type=int, default=147, help="district_id of the city")
#
#     args = parser.parse_args()
#     return args

def get_requests_list(xls_filename):
    xl = pandas.read_excel(xls_filename, sheet_name='Vaccine_Slot_Register')
    return xl

def parse_data(data):
    all_stuff = []
    for center in data["centers"]:
        for session in center['sessions']:
            if int(session['available_capacity']) > 0:
                slots_time = session['slots']
                slots_time = "<br>".join(slots_time)
                min_age_limit = 'unknown via script'
                if session['min_age_limit'] == 18:
                    min_age_limit = '18+'
                elif session['min_age_limit'] == 45:
                    min_age_limit = '45+'
                stats = {
                    'Center Name' : center['name'],
                    'Center Address' : center['address'],
                    'pincode' : center['pincode'],
                    'date' : session['date'],
                    'vaccine' : session['vaccine'],
                    'available_capacity' : session['available_capacity'],
                    'slots' : slots_time,
                    'age': min_age_limit
                }
                all_stuff.append(stats)
    return all_stuff


def send_mail(mail_to, mail_subject, email_body_html):
    msg = MIMEMultipart()
    msg['From'] = '<email id to send email from>'
    msg['To'] = ', '.join(map(str,mail_to))
    msg['Subject'] = mail_subject

    send_from = '<email id to send email from>'

    username = '<email username>'
    password = 'email password'
    body = "<!DOCTYPE html> <html> <body> <h3>This is my message!</h3> </body>  </html> "
    mail_server = 'mail server ip'

    part1 = MIMEText(email_body_html, 'html')
    msg.attach(part1)


    smtp = smtplib.SMTP(mail_server)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, mail_to, msg.as_string())
    smtp.quit()



if __name__ == '__main__':
    # args = parse_args()
    print (xls_filename)
    xl = get_requests_list(xls_filename)
    for req in xl.itertuples():
        if str(req.Subscribed).lower() == 'yes':
            html = ""
            all_stuff = []
            for pincode in [s.strip() for s in str(req.PinCode_separated_by_commas_if_more_than_1).split(',')]:
                print (req)
                data = get_vaccine_json(pincode,type='pincode')
                all_stuff = all_stuff + parse_data(data)
            if len(all_stuff) > 0:
                df = pandas.DataFrame(data=all_stuff, columns=all_stuff[0].keys())
                df = df.sort_values(by=['date'], ascending=True)
                html = df.to_html(index=False, escape=False)

            if html:
                send_mail([req.Official_Email_ID],"Vaccine available slots for pincde locations", html)

