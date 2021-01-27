import cx_Oracle
import pdfkit
import os
import hashlib

import requests
from PyPDF2 import PdfFileWriter, PdfFileReader

# create connection


try:
    conn = cx_Oracle.connect('prhlivero/prhlivero421@//52.220.233.172:1521/hinaidbc')
    conn1 = cx_Oracle.connect('phit/phit@//182.163.102.118:1521/phit')

except Exception as err:
    print('Exception occured while trying to create a connection', err)


else:
    try:
        cur = conn.cursor()
        cur1 = conn1.cursor()
        sql = """ 

        select distinct o.MRNO as UPI,
        p.PATIENTNAME          as NAME,
        lo.sample_id            as Sample_No,
        lo.AGE                  as Age,
        lo.gender               as gender,
        lo.orderedby,
        (case when hd.DEPARTMENT_NAME is null then 'self' else hd.DEPARTMENT_NAME end)      as Location,
        p.MOBILENO             as Contact_No,
        to_char(lo.SampleGenerated_Date,'dd/mm/yyyy HH24:mi') as Collected_On,
        pa.ADDRESS             as Address,
        to_char(lo.SampleAccepted_Date,'dd/mm/yyyy HH24:mi')  as Received_On,
        (case when u.username is null then 'self' else u.username end)      as Doctor_name,
        
        
        to_char(lr.CERTIFIED_DATE,'dd/mm/yyyy HH24:mi')       as Reported_On,

        lr.Result_Value as Result,
        
        (case when p.EMAIL is null then 'self' else p.EMAIL end) as Email
       
        from LABRESULT lr

        left join LABORDER lo

        on lr.lab_order_id = lo.lab_order_id


        left join orders o

        on lr.patient_id = o.patient_id

        left join PRHLIVE.PATIENT p

        on lr.patient_id = p.patient_id

        left join (select patient_id,address from patient_address where active=1) pa

        on p.patient_id = pa.patient_id
        left join HISUSER u on u.id = o.attendingdoctor -- take attending_doctor from orders table instead t.orderedby


        --left join referraldoctor r

        --on lo.orderedby = r.id

        left join employee e

        on u.employee_id = e.employee_id

        left join hisdepartment hd
        on e.department_id = hd.department_id
         where lo.lab_service_id = 1327
         and trunc(lr.certified_date,'HH24') = trunc((sysdate-1/24),'HH24')
         --and lr.certified_date> = trunc(sysdate-1) + 16/24
         and o.mrno is not null





        """
        cur.execute(sql)
        rows = cur.fetchall()
        #print(rows)
        if 1==1:
            for index, record in enumerate(rows):
                break_address=record[9]
                if len(break_address ) >31:
                    break_address = break_address[0:30]+ (' ') +'<br>'+break_address[30:]
                    print(break_address)
                # print('index is', index, ':', record[0])
                fh = open('Covid_Report.html', 'w')
                # fh.write(record[0] + "\n")
                # fh.write("""<div style='position:absolute;left:128.55px;top:97.78px" class="cls_004"><span class="cls_004">"""+record[0]+"""</span></div>""" + "\n")
                # fh.write("""<div style="position:absolute;left:141.23px;top:72.90px" class="cls_003"><span class="cls_003">Laboratory Report : Molecular Infection Testing</span></div>"""+"\n")
                fh.writelines(["""<html>""" + "\n",
                               """<head><meta http-equiv=Content-Type content="text/html; charset=UTF-8">""" + "\n",
                               """<style type="text/css">""" + "\n",
                               """span.cls_003{font-family:Times,serif;font-size:15.1px;color:rgb(0,0,0);font-weight:bold;font-style:normal;text-decoration: none}""" + "\n",
                               """div.cls_003{font-family:Times,serif;font-size:15.1px;color:rgb(0,0,0);font-weight:bold;font-style:normal;text-decoration: none}""" + "\n",
                               """span.cls_002{font-family:Times,serif;font-size:10.1px;color:rgb(0,0,0);font-weight:bold;font-style:normal;text-decoration: none}""" + "\n",
                               """div.cls_002{font-family:Times,serif;font-size:10.1px;color:rgb(0,0,0);font-weight:bold;font-style:normal;text-decoration: none}""" + "\n",
                               """span.cls_004{font-family:Times,serif;font-size:10.1px;color:rgb(0,0,0);font-weight:normal;font-style:normal;text-decoration: none}""" + "\n",
                               """div.cls_004{font-family:Times,serif;font-size:10.1px;color:rgb(0,0,0);font-weight:normal;font-style:normal;text-decoration: none}""" + "\n",
                               """</style>""" + "\n",
                               """<script type="text/javascript" src="Covid Report_files/wz_jsgraphics.js"></script>""" + "\n",
                               """</head>""" + "\n",
                               """<body>""" + "\n",
                               """<div style="position:absolute;left:50%;margin-left:-297px;top:0px;width:595px;height:841px;border-style:outset;overflow:hidden">""" + "\n",
                               """<div style="position:absolute;left:0px;top:0px">""" + "\n",
                               """<img src="file:///C:/Bappy/Covid/Covid_Report_files/background.jpg" width=595 height=841></div>""" + "\n",
                               """<div style="position:absolute;left:141.23px;top:72.90px" class="cls_003"><span class="cls_003">Laboratory Report : Molecular Infection Testing</span></div>""" + "\n",
                               """<div style="position:absolute;left:20.55px;top:98.15px" class="cls_002"><span class="cls_002">UPI</span></div>""" + "\n",
                               """<div style="position:absolute;left:128.55px;top:97.78px" class="cls_004"><span class="cls_004">""" +
                               record[0] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:299.44px;top:98.15px" class="cls_002"><span class="cls_002">Sample Type</span></div>""" + "\n",
                               """<div style="position:absolute;left:407.44px;top:97.78px" class="cls_004"><span class="cls_004">Nasopharyngeal swab</span></div>""" + "\n",
                               """<div style="position:absolute;left:20.55px;top:112.68px" class="cls_002"><span class="cls_002">Name</span></div>""" + "\n",
                               """<div style="position:absolute;left:128.55px;top:112.31px" class="cls_004"><span class="cls_004">""" +
                               record[1] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:299.44px;top:112.68px" class="cls_002"><span class="cls_002">Sample No</span></div>""" + "\n",
                               """<div style="position:absolute;left:407.44px;top:112.31px" class="cls_004"><span class="cls_004">""" +
                               record[2] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:20.55px;top:127.21px" class="cls_002"><span class="cls_002">Age /Gender</span></div>""" + "\n",
                               """<div style="position:absolute;left:128.55px;top:126.84px" class="cls_004"><span class="cls_004">""" +
                               record[3] + """/""" + record[4] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:299.44px;top:127.21px" class="cls_002"><span class="cls_002">Location</span></div>""" + "\n",
                               """<div style="position:absolute;left:407.44px;top:126.84px" class="cls_004"><span class="cls_004">""" + record[6] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:20.55px;top:141.74px" class="cls_002"><span class="cls_002">Contact No</span></div>""" + "\n",
                               """<div style="position:absolute;left:128.55px;top:141.37px" class="cls_004"><span class="cls_004">""" +
                               record[7] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:299.44px;top:141.74px" class="cls_002"><span class="cls_002">Collected On</span></div>""" + "\n",
                               """<div style="position:absolute;left:407.44px;top:141.37px" class="cls_004"><span class="cls_004">""" +
                               record[8] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:20.55px;top:156.27px" class="cls_002"><span class="cls_002">Address</span></div>""" + "\n",
                               """<div style="position:absolute;left:128.55px;top:155.90px" class="cls_004"><span class="cls_004">""" +
                               break_address + """ ,</span></div>""" + "\n",
                               """<div style="position:absolute;left:299.44px;top:156.27px" class="cls_002"><span class="cls_002">Received On</span></div>""" + "\n",
                               """<div style="position:absolute;left:407.44px;top:155.90px" class="cls_004"><span class="cls_004">""" +
                               record[10] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:128.55px;top:167.06px" class="cls_004"><span class="cls_004"></span></div>""" + "\n",
                               """<div style="position:absolute;left:20.55px;top:181.59px" class="cls_002"><span class="cls_002">Referred By</span></div>""" + "\n",
                               """<div style="position:absolute;left:128.55px;top:181.22px" class="cls_004"><span class="cls_004">""" + record[11] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:299.44px;top:181.59px" class="cls_002"><span class="cls_002">Reported On</span></div>""" + "\n",
                               """<div style="position:absolute;left:407.44px;top:181.22px" class="cls_004"><span class="cls_004">""" +
                               record[12] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:22.05px;top:197.62px" class="cls_002"><span class="cls_002">Test Requested</span></div>""" + "\n",
                               """<div style="position:absolute;left:128.55px;top:197.25px" class="cls_004"><span class="cls_004">COVID-19</span></div>""" + "\n",
                               """<div style="position:absolute;left:61.67px;top:218.15px" class="cls_002"><span class="cls_002">Test Name</span></div>""" + "\n",
                               """<div style="position:absolute;left:164.64px;top:218.15px" class="cls_002"><span class="cls_002">Test Method</span></div>""" + "\n",
                               """<div style="position:absolute;left:277.93px;top:218.15px" class="cls_002"><span class="cls_002">Result</span></div>""" + "\n",
                               """<div style="position:absolute;left:388.48px;top:218.15px" class="cls_002"><span class="cls_002">Unit</span></div>""" + "\n",
                               """<div style="position:absolute;left:475.21px;top:218.15px" class="cls_002"><span class="cls_002">Reference Range</span></div>""" + "\n",
                               """<div style="position:absolute;left:22.05px;top:236.06px" class="cls_004"><span class="cls_004">SARS-CoV-2 PCR</span></div>""" + "\n",
                               """<div style="position:absolute;left:174.63px;top:236.06px" class="cls_004"><span class="cls_004">RT PCR</span></div>""" + "\n",
                               """<div style="position:absolute;left:275.43px;top:236.06px" class="cls_004"><span class="cls_004">""" +
                               record[13] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:509.78px;top:236.06px" class="cls_004"><span class="cls_004">-</span></div>""" + "\n",
                               """<div style="position:absolute;left:23.55px;top:270.00px" class="cls_002"><span class="cls_002">PrecautionaryComments :</span></div>""" + "\n",
                               """<div style="position:absolute;left:22.05px;top:287.91px" class="cls_004"><span class="cls_004">Please correlate with clinical conditions</span></div>""" + "\n",
                               """<div style="position:absolute;left:22.80px;top:309.94px" class="cls_002"><span class="cls_002">Limitations:</span></div>""" + "\n",
                               """<div style="position:absolute;left:21.30px;top:327.85px" class="cls_004"><span class="cls_004">Negative results do not rule out SARS-CoV-2 infection and should not be used as the only basis for patient management decisions.</span></div>""" + "\n",
                               """<div style="position:absolute;left:21.30px;top:339.01px" class="cls_004"><span class="cls_004">Negative results must be combined with clinical observations (symptoms), patient history, and epidemiological information. Please</span></div>""" + "\n",
                               """<div style="position:absolute;left:21.30px;top:350.17px" class="cls_004"><span class="cls_004">contact your doctor or book an appointment with a Praava doctor for further clarification.</span></div>""" + "\n",
                               """<div style="position:absolute;left:88.12px;top:666.38px" class="cls_002"><span class="cls_002">Prepared By</span></div>""" + "\n",
                               """<div style="position:absolute;left:274.33px;top:666.38px" class="cls_002"><span class="cls_002">Reviewed By</span></div>""" + "\n",
                               """<div style="position:absolute;left:453.46px;top:666.38px" class="cls_002"><span class="cls_002">Authorized By</span></div>""" + "\n",
                               """<div style="position:absolute;left:105.20px;top:680.54px" class="cls_004"><span class="cls_004">1384</span></div>""" + "\n",
                               """<div style="position:absolute;left:274.33px;top:680.54px" class="cls_004"><span class="cls_004">Shafiul Azam</span></div>""" + "\n",
                               """<div style="position:absolute;left:446.80px;top:680.54px" class="cls_004"><span class="cls_004">Dr. Zaheed Husain</span></div>""" + "\n",
                               """<div style="position:absolute;left:66.60px;top:694.70px" class="cls_004"><span class="cls_004">Junior Scientific Officer</span></div>""" + "\n",
                               """<div style="position:absolute;left:228.50px;top:694.70px" class="cls_004"><span class="cls_004">M.Sc. in Biochemistry &amp; Molecular</span></div>""" + "\n",
                               """<div style="position:absolute;left:414.03px;top:694.70px" class="cls_004"><span class="cls_004">Senior Director, Molecular Cancer</span></div>""" + "\n",
                               """<div style="position:absolute;left:231.28px;top:705.86px" class="cls_004"><span class="cls_004">Biology, Sr. Laboratory Supervisor</span></div>""" + "\n",
                               """<div style="position:absolute;left:461.10px;top:705.86px" class="cls_004"><span class="cls_004">Diagnostics</span></div>""" + "\n",
                               """<div style="position:absolute;left:21.30px;top:740.18px" class="cls_004"><span class="cls_004">""" +
                               record[1] + """/""" + record[0] + """</span></div>""" + "\n",
                               """<div style="position:absolute;left:257.06px;top:757.71px" class="cls_002"><span class="cls_002">Page</span></div>""" + "\n",
                               """<div style="position:absolute;left:291.48px;top:757.34px" class="cls_004"><span class="cls_004">1</span></div>""" + "\n",
                               """<div style="position:absolute;left:309.61px;top:757.34px" class="cls_004"><span class="cls_004">of</span></div>""" + "\n",
                               """<div style="position:absolute;left:331.08px;top:757.34px" class="cls_004"><span class="cls_004">1</span></div>""" + "\n",
                               """</div>""" + "\n",
                               """</body>""" + "\n",
                               """</html>"""])
                fh.close()
                # os.system('python main2.py') #multti variable
                # PDF Code started
                # path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
                # config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

                # pdfkit.from_file("Covid_Report.html","res1.pdf")

                # options = {
                # "enable-local-file-access": None
                # }

                # reportPdfName = os.path.join(r'C:\inetpub\wwwroot\P3R\Resources\PDF\Reports',record[0]+ ".pdf")
                # pdfkit.from_file("Covid_Report.html", reportPdfName, options=options, configuration=config)

                path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
                config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

                # pdfkit.from_file("Covid_Report.html","res1.pdf")

                options = {
                    'disable-smart-shrinking': "",
                    "enable-local-file-access": None
                }
                # patientid = r"\121.pdf"
                reportHash = (hashlib.md5(record[2].encode('UTF-8')).hexdigest())
                reportPdfName = os.path.join(r'C:\inetpub\wwwroot\P3R\Resources\PDF\Reports', reportHash + ".pdf")
                reportHashPdfName = (reportHash + ".pdf")
                # reportPdfName = r'C:\inetpub\wwwroot\P3R\Resources\PDF' + patientid
                # reportHash = (hashlib.md5(reportPdfName.encode('UTF-8')).hexdigest())
                pdfkit.from_file("Covid_Report.html", reportPdfName, options=options, configuration=config)

                # PDF Code Ended

                #pdf short url start
                api_key = "309f61194cfeee533c6ba70546e8d74378ce5"
                params = f"{reportHashPdfName}"
                url = f"http://partner.praava.com/Resources/PDF/Reports/{params}"
                # api_url = f"https://cutt.ly/api/api.php?key={api_key}&short={url}"
                api_url = f"https://cutt.ly/api/api.php?key=309f61194cfeee533c6ba70546e8d74378ce5&short={url}"
                data = requests.get(api_url).json()["url"]
                if data["status"] == 7:
                    shortened_url = data["shortLink"]

                    print(record[0], shortened_url)

                    insert_query = f"insert into p3r_covid_report (UPI,MOBILE_NO,CERTIFIED_AT,PATIENT_URL,ORGNL_URL) values('{record[0]}','{record[7]}',to_date('{record[12]}','dd/mm/yyyy HH24:mi'),'{shortened_url}','{url}')"

                    cur1.execute(insert_query)
                    conn1.commit()

                #pdf short url ended

                #sms part start

                sms_body = f"Your COVID-19 test result at Praava with UPI: {record[0]} and name: {record[1]} is {record[13]}.You will receive an email at {record[14]} with the report within the next 24 hours, please also check your spam folder if you can't find it in your inbox. If required, a report hard copy can be provided at our Praava Health clinic in Banani.You can also view and download your report from {shortened_url} using the last 4 digits of your mobile number. Thank you."
                sms_to = record[7]
                #sms_to = "8801682269024"
                sms_url = f"http://10.0.5.10/API/sms?key=4ebf50722d05a9e9f5d1818d55f6bf58&to={sms_to}&text={sms_body}"
                requests.get(sms_url)

                #sms part ended

                # password code.
                pdfwriter = PdfFileWriter()

                pdf = PdfFileReader(reportPdfName)

                for page_num in range(pdf.numPages):
                    pdfwriter.addPage(pdf.getPage(page_num))

                passw = record[7][-4:]
                #passw = '1'
                pdfwriter.encrypt(passw)

                with open(reportPdfName, 'wb') as f:
                    pdfwriter.write(f)
                    f.close()

    except Exception as err:
        print('Exception occured while trying to create a connection', err)
    else:
        print('Completed.')
    finally:
        cur.close()

finally:
    conn.close()

