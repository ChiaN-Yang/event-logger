import win32evtlog
import xmltodict
import csv
import configparser

def SearchEvents(LogName, EventId=112, count=20):
    # EventLog = win32evtlog.EvtOpenLog(LogName, 1, None)

    # TotalRecords = win32evtlog.EvtGetLogInfo(EventLog, win32evtlog.EvtLogNumberOfLogRecords)[0]
    ResultSet = win32evtlog.EvtQuery(LogName, win32evtlog.EvtQueryReverseDirection, "*[System[(EventID=%d)]]" % EventId, None)

    EventList = []
    for evt in win32evtlog.EvtNext(ResultSet, count):
        res = xmltodict.parse(win32evtlog.EvtRender(evt, 1))

        EventData = {}
        EventData['TimeCreated'] = res['Event']['System']['TimeCreated']['@SystemTime'].split('.')[0]
        for e in res['Event']['EventData']['Data']:
            if '#text' in e:
                EventData[e['@Name']] = e['#text']

        EventList.append(EventData)

    return EventList


def filterLogsByDate(log_data, start_date, end_date):
    filtered_log = list()
    for log in log_data:
        if end_date > log['TimeCreated'] > start_date:
            filtered_log.append(log)

    return filtered_log


def eventList2DeviceSet(event_list):
    device = set()
    for event in event_list:
        device.add(event['Prop_DeviceName'])
    return device


def writecsv(my_dict, csv_name):
    my_dict = list(my_dict)
    with open(csv_name, 'w', newline='') as f:
        w = csv.writer(f)
        for item in my_dict:
            w.writerow([item])


def main():
    config = configparser.ConfigParser()
    config.read('config.ini')
    number_logs = int(config.get('Settings','number_logs'))
    start_time = config.get('Settings','start_date')
    end_time = config.get('Settings','end_date')
    events = SearchEvents('Microsoft-Windows-DeviceSetupManager/Admin', count=number_logs)
    filted_events = filterLogsByDate(events, start_time, end_time)
    device_set = eventList2DeviceSet(filted_events)
    print(device_set)
    writecsv(device_set, 'device.csv')


if __name__ == '__main__':
    main()
