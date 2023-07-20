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

import json

#Begin of function section
#---------------------------------------------------------

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
    print("Swissgrid schedule is imbalanced in the following periods:")

    for key in imbalancedPeriods.keys():
        print(key)
else:
    print("Hooray: Swissgrid schedule is balanced.")
