import json
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import pandas as pd

#Constants
#---------------------------------------------------------

SWISSGRID_EIC = "10YCH-SWISSGRIDZ"
STATKRAFT_EIC = "11XSTATKRAFT001N"

AMPRION_EIC="10YDE-RWENET---I"

SUBSTRING_TO_FIND_SWISS="_TPS_11XSTATKRAFT001N_10XCH-SWISSGRIDC_"

NODE_SCHEDULE_MESSAGE = 'ScheduleMessage'
NODE_MESSAGEVERSION="MessageVersion"
NODE_SCHEDULE_TIMESERIES = 'ScheduleTimeSeries'
NODE_PERIOD = 'Period'
NODE_INTERVAL = 'Interval'
NODE_IN_AREA = 'InArea'
NODE_OUT_AREA = 'OutArea'
NODE_IN_PARTY = 'InParty'
NODE_OUT_PARTY = 'OutParty'
NODE_V = 'v'
NODE_QTY = 'Qty'
#---------------------------------------------------------


#Begin of function section
#---------------------------------------------------------


def send_mail(send_from, send_to, send_cc, send_bcc, subject, message, files=[],
              server="localhost", port=587, loginRequired=True, username='', password='',
              use_tls=True):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (list[str]): to name(s)
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
        use_tls (bool): use TLS mode
    """
    msg = MIMEMultipart()
    msg['From'] = send_from

    if len(send_to)==1:
        msg['To'] = send_to
    else:
        msg['To'] =", ".join(send_to)

    if len(send_cc)==1:
        msg['Cc'] = send_cc
    else:
        msg['Cc']=", ".join(send_cc)

    if len(send_cc)==1:
        msg['Bcc']=send_bcc
    else:
        msg['Bcc']=", ".join(send_bcc)

    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename={}'.format(Path(path).name))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    if use_tls:
        smtp.starttls()
    if loginRequired:
        smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()

    print("Mail sent successfully")

def SendMailPythonServer(send_to, send_cc, send_bcc,subject, message):
    msgBody = """<html><head></head>
        <style type = "text/css">
            table, td {height: 3px; font-size: 14px; padding: 5px; border: 1px solid black;}
            td {text-align: left;}
            body {font-size: 12px;font-family:Calibri}
            h2,h3 {font-family:Calibri}
            p {font-size: 14px;font-family:Calibri}
         </style>"""
                   
    msgBody += "<h2>" + subject + "</h2>"
    #msgBody += "<h3>" + message + "</h3>"
    msgBody += message
    
    strFrom = "no-reply-duswvpyt002p@statkraft.de"
    
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = subject
    msgRoot['From'] = strFrom
    if len(send_to)==1:
        msgRoot['To'] = send_to
    else:
         msgRoot['To'] =", ".join(send_to)

    if len(send_cc)==1:
         msgRoot['Cc'] = send_cc
    else:
         msgRoot['Cc']=", ".join(send_cc)

    if len(send_cc)==1:
         msgRoot['Bcc']=send_bcc
    else:
         msgRoot['Bcc']=", ".join(send_bcc)
    msgRoot.preamble = 'This is a multi-part message in MIME format.'

    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    msgText = MIMEText('Sorry this mail requires your mail client to allow for HTML e-mails.')
    msgAlternative.attach(msgText)

    msgText = MIMEText(msgBody, 'html')
    msgAlternative.attach(msgText)


    smtp = smtplib.SMTP('smtpdus.energycorp.com')
    smtp.sendmail(strFrom, recipients, msgRoot.as_string())
    smtp.quit()

    print("Mail sent successfully from " + strFrom)

def GetLatestSwissIntradaySchedule(myPath, substringToFind):
    from os import listdir
    from os.path import isfile, join
    from datetime import datetime

    onlyfiles = [f for f in listdir(myPath)
                 if isfile(join(myPath, f))
                 if f.__contains__(substringToFind)
                 if f.__contains__(datetime.now().strftime("%Y%m%d"))
                 ]

    latestSchedule=""
    if len(onlyfiles)>0:

        highestVersion=1
        latestSchedule=onlyfiles[0]

        for file in onlyfiles:

            version=float(file[-7:][:3])

            if version>highestVersion:
                highestVersion=version
                latestSchedule=file

    return latestSchedule


def HousekeepingIntradayScheduleVersion(myPath, substringToFind):
    from os import listdir
    from os import remove as FileDelete
    from os.path import isfile, join
    from datetime import datetime, timedelta

    onlyfiles = [f for f in listdir(myPath)
                 if isfile(join(myPath, f))
                 if substringToFind in f
                 if (datetime.now()+ timedelta(days = 2)).strftime("%Y%m%d") not in f
                 if (datetime.now()+ timedelta(days = 1)).strftime("%Y%m%d") not in f
                 if (datetime.now()+ timedelta(days = 0)).strftime("%Y%m%d") not in f
                 if (datetime.now()+ timedelta(days = -1)).strftime("%Y%m%d") not in f
                 if (datetime.now()+ timedelta(days = -2)).strftime("%Y%m%d") not in f
                 if (datetime.now()+ timedelta(days = -3)).strftime("%Y%m%d") not in f
                 if (datetime.now()+ timedelta(days = -4)).strftime("%Y%m%d") not in f
                 if (datetime.now()+ timedelta(days = -5)).strftime("%Y%m%d") not in f
                 if (datetime.now()+ timedelta(days = -6)).strftime("%Y%m%d") not in f
                 if (datetime.now()+ timedelta(days = -7)).strftime("%Y%m%d") not in f
                 ]

    for file in onlyfiles:
       FileDelete(join(myPath,file))

def GetLatestSwissIntradayScheduleVersion(myPath):
    from os import listdir
    from os.path import isfile, join
    from datetime import datetime

    onlyfiles = [f for f in listdir(myPath)
                 if isfile(join(myPath, f))
                 if f.__contains__(SUBSTRING_TO_FIND_SWISS)
                 if f.__contains__(datetime.now().strftime("%Y%m%d"))
                 ]

    highestVersion=1

    for file in onlyfiles:

        version=float(file[-7:][:3])

        if version>highestVersion:
            highestVersion=version


    return int(highestVersion)


def GetEmailBody(imbalancedPeriods, version):

    warnMessage="Hi," + "\n\n" + "please be advised that the current SWISS intraday schedule (Version: " + str(version) + ") is imbalanced by the following periods:" + "\n\n"

    for key in imbalancedPeriods.keys():

        print(key)

        warnMessage=warnMessage + key + "\n"


    warnMessage=warnMessage + "\n\n" + "Kind regards" + "\n\n\n" +"Stakraft Operations"

    return warnMessage

def GetEmailSubject(version):

    warnMessage="Attention: Swissgrid intraday schedule (V" + str(version) + ") is imbalanced"

    return warnMessage

def GetImportFlows(scheduleMessage, tsoEIC):
    externalFlowsInto = {}

    for ts in scheduleMessage[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES]:

        if str(ts[NODE_IN_AREA][NODE_V]) == tsoEIC and str(ts[NODE_OUT_AREA][NODE_V]) != tsoEIC and str(
                ts[NODE_IN_PARTY][NODE_V]) == STATKRAFT_EIC and str(ts[NODE_OUT_PARTY][NODE_V]) == STATKRAFT_EIC:
            externalFlowsInto[str(ts[NODE_OUT_AREA][NODE_V]) + " => " + str(ts[NODE_IN_AREA][NODE_V])] = ts

    return externalFlowsInto


def GetExportFlows(scheduleMessage, tsoEIC):
    externalFlowsOutOf = {}

    for ts in scheduleMessage[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES]:

        if str(ts[NODE_IN_AREA][NODE_V]) != tsoEIC and str(ts[NODE_OUT_AREA][NODE_V]) == tsoEIC and str(
                ts[NODE_IN_PARTY][NODE_V]) == STATKRAFT_EIC and str(ts[NODE_OUT_PARTY][NODE_V]) == STATKRAFT_EIC:
            externalFlowsOutOf[str(ts[NODE_OUT_AREA][NODE_V]) + " => " + str(ts[NODE_IN_AREA][NODE_V])] = ts

    return externalFlowsOutOf


def GetBuys(scheduleMessage, tsoEIC):
    internalBuyPositions = {}

    for ts in scheduleMessage[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES]:

        if str(ts[NODE_IN_AREA][NODE_V]) == tsoEIC and str(ts[NODE_OUT_AREA][NODE_V]) == tsoEIC and str(
                ts[NODE_IN_PARTY][NODE_V]) == STATKRAFT_EIC and str(ts[NODE_OUT_PARTY][NODE_V]) != STATKRAFT_EIC:
            internalBuyPositions[
                str(ts[NODE_OUT_PARTY][NODE_V]) + " => " + str(ts[NODE_IN_PARTY][NODE_V])] = ts

    return internalBuyPositions


def GetSells(scheduleMessage, tsoEIC):
    internalSellPositions = {}

    for ts in scheduleMessage[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES]:

        if str(ts[NODE_IN_AREA][NODE_V]) == tsoEIC and str(ts[NODE_OUT_AREA][NODE_V]) == tsoEIC and str(
                ts[NODE_IN_PARTY][NODE_V]) != STATKRAFT_EIC and str(ts[NODE_OUT_PARTY][NODE_V]) == STATKRAFT_EIC:
            internalSellPositions[
                str(ts[NODE_OUT_PARTY][NODE_V]) + " => " + str(ts[NODE_IN_PARTY][NODE_V])] = ts

    return internalSellPositions


def GetAggrPos(LstTimeseries):
    aggregatedPosition = {}

    if len(LstTimeseries) > 0:

        for x in range(len(list(LstTimeseries.values())[0][NODE_PERIOD][NODE_INTERVAL])):

            volume = 0

            for ts in LstTimeseries.values():
                thisVol = float(ts[NODE_PERIOD][NODE_INTERVAL][x][NODE_QTY][NODE_V])

                volume = volume + thisVol

            aggregatedPosition[x] = volume

    return aggregatedPosition


def GetJsonContent(uncPath):
    import xmltodict, json

    fileptr = open(uncPath.replace("\"", ""), "r")

    xml_content = fileptr.read()

    o = xmltodict.parse(xml_content)
    json1 = json.dumps(o)

    jsonContent = json1.replace('@', '')

    return jsonContent

def GetDirection(imbalanceVolume):
    if imbalanceVolume > 0:
        return "long"
    else:
        return "short"

def GetImbalancePeriods(exportFlows,importFlows,buyPositions,sellPositions,periods, leadTimeBufferQuarterPeriods):

    #Flows and positions are aggregated over balance groups

    imbalancedPeriods={}

    for x in range(periods):

        if GetCurrentQuarterPeriod()+leadTimeBufferQuarterPeriods <= x:

            if len(exportFlows)>0:
                flowsOut = exportFlows[x]
            else:
                flowsOut = 0

            if len(importFlows) > 0:
                flowsIn = importFlows[x]
            else:
                flowsIn = 0

            imbalanceVolume = flowsIn - flowsOut + buyPositions[x] - sellPositions[x]

            if imbalanceVolume != 0:

                key='Period ' + str( x+1).zfill(2) + " (" + GetTimestamp(x+1) + ")" + ": " + str(abs(imbalanceVolume)) + " MW " + GetDirection(imbalanceVolume)

                imbalancedPeriods[key] = key

    return imbalancedPeriods


def GetCurrentQuarterPeriod():
    import datetime
    now = datetime.datetime.now()

    currentHour= now.hour
    currentMinute=now.minute

    if currentMinute>0 and currentMinute<=15:
        minute=1
    elif currentMinute>15 and currentMinute <=30:
        minute=2
    elif currentMinute>30 and currentMinute<=45:
        minute=3
    else:
        minute=4

    return ((currentHour)*4)+minute

def GetTimestamp(Period):

    startHour =int ((Period-1) / 4);

    startMinute = ((startHour+1) * 4 - Period );

    if startMinute==0:
        #return (startHour).ToString("00") + ":45" + " - " + (startHour+1).ToString("00") + ":00"
        return str(startHour).zfill(2) + ":45" + " - " + str(startHour).zfill(2) + ":00"
    elif startMinute==1:
        #return (startHour).ToString("00") + ":30" + " - " + (startHour).ToString("00") + ":45"
        return str(startHour).zfill(2) + ":30" + " - " + str(startHour).zfill(2) + ":45"
    elif startMinute==2:
        #return (startHour).ToString("00") + ":15" + " - " + (startHour).ToString("00") + ":30"
        return str(startHour).zfill(2) + ":15" + " - " + str(startHour).zfill(2) + ":30"
    else:
        #return (startHour).ToString("00") + ":00" + " - " + (startHour).ToString("00") + ":15"
        return str(startHour).zfill(2) + ":00" + " - " + str(startHour).zfill(2) + ":15"


#End of function section
#---------------------------------------------------------

leadTimeBufferQuarterperiods=8

path="C:\\Test\\"

path=r"\\energycorp.com\\common\\DIVSEDE\\Operations\\DeltaXE\\Schedules_ManualUpload\\"

recipients=["lukas.dicke@web.de", "lukas.dicke@statkraft.de"]

substringToFind=SUBSTRING_TO_FIND_SWISS

#substringToFind="_TPS_11XSTATKRAFT001N_10XDE-RWENET---W"

tsoEic=SWISSGRID_EIC

#tsoEic=AMPRION_EIC

HousekeepingIntradayScheduleVersion(path,substringToFind)

file=GetLatestSwissIntradaySchedule(path,substringToFind)

time=GetCurrentQuarterPeriod()

if file!="":

    message = json.loads(GetJsonContent(path + file))

    messageVersion=message[NODE_SCHEDULE_MESSAGE][NODE_MESSAGEVERSION][NODE_V]


    imbalancedPeriods=GetImbalancePeriods(exportFlows= GetAggrPos(GetExportFlows(message,tsoEic)),
                                          importFlows= GetAggrPos(GetImportFlows(message,tsoEic)),
                                          buyPositions= GetAggrPos(GetBuys(message,tsoEic)),
                                          sellPositions= GetAggrPos(GetSells(message,tsoEic)),
                                          periods= len(message[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES][0][NODE_PERIOD][NODE_INTERVAL]),
                                          leadTimeBufferQuarterPeriods=leadTimeBufferQuarterperiods)

    if len(imbalancedPeriods)>0:

        print("Swissgrid schedule (V" + str(messageVersion) + ") is imbalanced. Email is triggered")

        #modified_dict = dict(d.values() for d in imbalancedPeriods)

       # df = pd.DataFrame(imbalancedPeriods.values(),columns=['Swiss intraday imbalances'])

        #print(df)

        # send_mail_statmark(send_from="nominations@statmark.de",
        #                    send_to=recipients,
        #                    send_cc= [],
        #                    send_bcc= [],
        #                    subject=GetEmailSubject(messageVersion),
        #                    message=GetEmailBody(imbalancedPeriods, messageVersion),
        #                    files=[],
        #                    server= "mail.123domain.eu",
        #                    port=587,
        #                    username="nominations@statmark.de",
        #                    password="fk390krvadf",
        #                    use_tls=True)

        SendMailPythonServer(send_to=recipients, 
                             send_cc=[], 
                             send_bcc=[],
                             subject= GetEmailSubject(messageVersion), 
                             message=GetEmailBody(imbalancedPeriods, messageVersion))

    else:
        print("Hooray: Swissgrid schedule (V" + str(messageVersion) + ") is perfectly balanced.")