import cx_Oracle
import os
import requests


# create connection


try:
    conn = cx_Oracle.connect('phit/phit@//182.163.102.118:1521/phit')

except Exception as err:
    print('Exception occured while trying to create a connection', err)


else:
    try:
        cur = conn.cursor()
        sql = """ 
        
        
        select distinct t.invoiceno   INV_NO,
                        t.mrno        UPI,
                        t.mobileno    MOBILE_NO,
                        t.patientname PATIENT,
                        to_char(t.BILLED_AT,'dd Mon yyyy')   DATES
          from v_nps_SMS_to_be_alerted_inv t 
         """
        cur.execute(sql)
        rows = cur.fetchall()
        print(rows)
        api_key = "309f61194cfeee533c6ba70546e8d74378ce5"

        for index, record in enumerate(rows):
            #print('index is', index, ':', record[0])
            params = f"p=191319%26inv_no={record[0]}%26upi={record[1]}"
            url = f"http://nps.praava.health:2020/Feedback/combined/NPS_Combined_Pg_1.aspx?{params}"
            #api_url = f"https://cutt.ly/api/api.php?key={api_key}&short={url}"
            api_url = f"https://cutt.ly/api/api.php?key=309f61194cfeee533c6ba70546e8d74378ce5&short={url}"
            data = requests.get(api_url).json()["url"]
            if data["status"] == 7:
                shortened_url = data["shortLink"]
                #shortened_url = record[0]
                print(record[0], shortened_url)
                cur.execute(f"insert into p3r_short_url (short_url,org_url) values('{shortened_url}','{url}')")
                conn.commit()
                #sms_body = shortened_url
                sms_body = f"Thank you for choosing Praava. You have taken our service on {record[4]} with {record[1]}. Please take a minute to fill out this feedback survey to help us improve: {shortened_url}"
                sms_to = record[2]
                #sms_to = "8801844220485"
                sms_url = f"http://10.0.5.10/API/sms?key=4ebf50722d05a9e9f5d1818d55f6bf58&to={sms_to}&text={sms_body}"
                requests.get(sms_url)
            else:
                print("[!] Error Shortening URL:", data)


    except Exception as err:
        print('Exception occured while trying to create a connection', err)
    else:
       print('Completed.')
finally:
    cur.close()


