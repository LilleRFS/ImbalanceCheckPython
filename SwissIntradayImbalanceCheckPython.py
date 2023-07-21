import json
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

#Constants
#---------------------------------------------------------

SWISSGRID_EIC = "10YCH-SWISSGRIDZ"
STATKRAFT_EIC = "11XSTATKRAFT001N"

NODE_SCHEDULE_MESSAGE = 'ScheduleMessage'
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


def send_mail(send_from, send_to,send_cc,send_bcc, subject, message, files=[],
              server="localhost", port=587, username='', password='',
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
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()

    print("Mail sent successfully")

def GetEmailBody(imbalancedPeriods):
    warnMessage="Hi," + "\n\n" + "please be advised that the current intraday schedule is imabalance in the following periods:" + "\n\n"
    for key in imbalancedPeriods.keys():
        print(key)
        warnMessage=warnMessage + key + "\n"

    warnMessage=warnMessage + "\n\n" + "Kind regards" + "\n\n\n" +"Stakraft Operations"

    return warnMessage

def GetExternalFlowsIntoSwitzerland(scheduleMessage):
    externalFlowsIntoSwitzerland = {}

    for ts in scheduleMessage[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES]:

        if str(ts[NODE_IN_AREA][NODE_V]) == SWISSGRID_EIC and str(ts[NODE_OUT_AREA][NODE_V]) != SWISSGRID_EIC and str(
                ts[NODE_IN_PARTY][NODE_V]) == STATKRAFT_EIC and str(ts[NODE_OUT_PARTY][NODE_V]) == STATKRAFT_EIC:
            externalFlowsIntoSwitzerland[str(ts[NODE_OUT_AREA][NODE_V]) + " => " + str(ts[NODE_IN_AREA][NODE_V])] = ts

    return externalFlowsIntoSwitzerland


def GetExternalFlowsOutOfSwitzerland(scheduleMessage):
    externalFlowsOutOfSwitzerland = {}

    for ts in scheduleMessage[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES]:

        if str(ts[NODE_IN_AREA][NODE_V]) != SWISSGRID_EIC and str(ts[NODE_OUT_AREA][NODE_V]) == SWISSGRID_EIC and str(
                ts[NODE_IN_PARTY][NODE_V]) == STATKRAFT_EIC and str(ts[NODE_OUT_PARTY][NODE_V]) == STATKRAFT_EIC:
            externalFlowsOutOfSwitzerland[str(ts[NODE_OUT_AREA][NODE_V]) + " => " + str(ts[NODE_IN_AREA][NODE_V])] = ts

    return externalFlowsOutOfSwitzerland


def GetInternalBuyPositionsSwitzerland(scheduleMessage):
    internalBuyPositionsSwitzerland = {}

    for ts in scheduleMessage[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES]:

        if str(ts[NODE_IN_AREA][NODE_V]) == SWISSGRID_EIC and str(ts[NODE_OUT_AREA][NODE_V]) == SWISSGRID_EIC and str(
                ts[NODE_IN_PARTY][NODE_V]) == STATKRAFT_EIC and str(ts[NODE_OUT_PARTY][NODE_V]) != STATKRAFT_EIC:
            internalBuyPositionsSwitzerland[
                str(ts[NODE_OUT_PARTY][NODE_V]) + " => " + str(ts[NODE_IN_PARTY][NODE_V])] = ts

    return internalBuyPositionsSwitzerland


def GetInternalSellPositionsSwitzerland(scheduleMessage):
    internalSellPositionsSwitzerland = {}

    for ts in scheduleMessage[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES]:

        if str(ts[NODE_IN_AREA][NODE_V]) == SWISSGRID_EIC and str(ts[NODE_OUT_AREA][NODE_V]) == SWISSGRID_EIC and str(
                ts[NODE_IN_PARTY][NODE_V]) != STATKRAFT_EIC and str(ts[NODE_OUT_PARTY][NODE_V]) == STATKRAFT_EIC:
            internalSellPositionsSwitzerland[
                str(ts[NODE_OUT_PARTY][NODE_V]) + " => " + str(ts[NODE_IN_PARTY][NODE_V])] = ts

    return internalSellPositionsSwitzerland


def AggregatePositions(LstTimeseries):
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

#End of function section
#---------------------------------------------------------

path="C:\\Test\\"
file="20230607_TPS_11XSTATKRAFT001N_10XCH-SWISSGRIDC_007.xml"

file="20230719_TPS_11XSTATKRAFT001N_10XCH-SWISSGRIDC_002.xml"

uncPath=path + file

uncPath = input("Please specify Swissgrid schedule path:")

jsonContent = GetJsonContent(uncPath)

message = json.loads(jsonContent)

periods = len(message[NODE_SCHEDULE_MESSAGE][NODE_SCHEDULE_TIMESERIES][0][NODE_PERIOD][NODE_INTERVAL])

externalFlowsIntoSwitzerland = GetExternalFlowsIntoSwitzerland(message)
externalFlowsOutOfSwitzerland = GetExternalFlowsOutOfSwitzerland(message)

internalBuyPositions = GetInternalBuyPositionsSwitzerland(message)
internalSellPositions = GetInternalSellPositionsSwitzerland(message)

aggrExternalFlowIntoSwitzerland = AggregatePositions(externalFlowsIntoSwitzerland)

aggrExternalFlowOutOfSwitzerland = AggregatePositions(externalFlowsOutOfSwitzerland)

aggrInternalBuyPositions = AggregatePositions(internalBuyPositions)

aggrInternalSellPositions = AggregatePositions(internalSellPositions)

imbalance=False

imbalancedPeriods={}

for x in range(periods):

    if len(aggrExternalFlowOutOfSwitzerland)>0:
        flowsOut = aggrExternalFlowOutOfSwitzerland[x]
    else:
        flowsOut = 0

    if len(aggrExternalFlowIntoSwitzerland) > 0:
        flowsIn = aggrExternalFlowIntoSwitzerland[x]
    else:
        flowsIn = 0

    imbalanceVolume = flowsIn- flowsOut + aggrInternalBuyPositions[x] - aggrInternalSellPositions[x]

    if imbalanceVolume != 0:
        
        direction = ""
        if imbalanceVolume > 0:
            direction = "long"
        else:
            direction = "short"

        key='Period ' + str( x+1).zfill(2) + ": " + str(abs(imbalanceVolume)) + " MW " + direction

        imbalancedPeriods[key] = key

        imbalance = True

if imbalance == True:

    print("Swissgrid schedule is imbalanced. Email is triggered")

    send_mail(send_from="nominations@statmark.de",
              send_to=["lukas.dicke@web.de","lukas.dicke@statkraft.de"],
              send_cc= [],
              send_bcc= [],
              subject="Attention: Swissgrid schedule is imbalanced",
              message=GetEmailBody(imbalancedPeriods),
              files=[],
              server= "mail.123domain.eu",
              port=587,
              username="nominations@statmark.de",
              password="fk390krvadf",
              use_tls=True)


    send_mail(send_from="nominations@statkraft.de",
              send_to=["lukas.dicke@web.de","lukas.dicke@statkraft.de"],
              send_cc= [],
              send_bcc= [],
              subject="Attention: Swissgrid schedule is imbalanced",
              message=GetEmailBody(imbalancedPeriods),
              files=[],
              server= "maildus.energycorp.com",
              port=25,
              username="nominations@statkraft.de",
              password="OlympicGames2018",
              use_tls=True)

else:
    print("Hooray: Swissgrid schedule is balanced.")