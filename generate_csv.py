# =================README=====================
# Start from here
# 1. Locate CX logbook directory and change the csv reader path below.
# e.g. if logbook is at root directory, enter "./Merged1_LogBook_2017 Mar-2019 Feb]"
logbook_paths = ["./kenneth/Merged1_LogBook_2017 Jan-2018 Dec.txt",
                 "./kenneth/Merged1_LogBook_2018 Dec-2020 Nov.txt",
                 "./kenneth/Merged1_LogBook_2022 Nov-2024 Apr.txt",]


# 2. Whose logbook is this?
name = "kenneth"

# 3. If P2X not needed, set p2x to False (automatically revert to P2/P1US)
p2x = True

# 4. (optional) go to daynight.py and adjust parameters if needed.
# =============================================



import datetime
import math
from daynight import caldaynight
from openpyxl import load_workbook
import csv
from tail_to_type import b773

log = []
# log = [{departure_date:2022/12/12, off_block_UTC: 12:22....}, {},....]

iata_to_icao = {}
with open("iata-icao.csv", encoding="utf8") as iataicaomap:
    reader = csv.reader(iataicaomap)
    for row in reader:
        if row[5][-1] not in {'0','1','2','3','4','5','6','7','8','9'}:
            continue
        iata_to_icao[row[2]] = row[3]


def logger(reader):
    for row in reader:
        if not row:
            continue
        if row[0][:2] != "20":
            continue
        flight_info = row[0].split()

        if len(flight_info) == 4:
            # sim duty
            if log and log[-1]['departure_date'] == flight_info[0] and (not log[-1]["isFlightDuty"] and log[-1]["duty_code"] == flight_info[3])  :
                continue
            log.append({"isFlightDuty": False,
                        "departure_date": flight_info[0],
                        "duty_code": flight_info[3],
                        })
        else:
            # flight duty
            if log and len(log[-1]) == 13:
                if (log and log[-1]['departure_date'] == flight_info[0] and
                        log[-1]['origin'] == iata_to_icao[flight_info[2]]):
                    print("DUPLICATE")
                    continue
            else:
                if (log and log[-1]['departure_date'] == flight_info[0] and
                        (log[-1]['duty_code'] == flight_info[1])):
                    print("DUPLICATE")
                    continue

            # Adjust CX logbook departure date(local) to UTC.
            # if off block time is +1 or -1 ,adjust departure date
            if len(flight_info[6]) > 5:
                if flight_info[6][-2:] == "+1":
                    flight_info[0] = datetime.datetime.strftime(
                        (datetime.datetime.strptime(flight_info[0], "%Y/%m/%d") + datetime.timedelta(days=1)),
                        "%Y/%m/%d"
                    )
                else:
                    flight_info[0] = datetime.datetime.strftime(
                        (datetime.datetime.strptime(flight_info[0], "%Y/%m/%d") - datetime.timedelta(days=1)),
                        "%Y/%m/%d"
                    )

            if flight_info[4] in b773:
                ac_type = 'B777-300'
            else:
                ac_type = 'B777-300ER'
            if len(flight_info) == 13:
                # missing airborne and landing time
                raise Exception("Missing airborne and landing time in log, please go and edit your log.")

            log.append({"isFlightDuty": True,

                        "departure_date":flight_info[0],

                        "reg" : flight_info[4],

                        "type" : ac_type,

                        "pic" : flight_info[-2] + " " + flight_info[-1],

                        "origin":iata_to_icao[flight_info[2]],

                        "dest":iata_to_icao[flight_info[3]],

                        "off_block_UTC":flight_info[6],

                        "airborne_UTC":flight_info[7],

                        "landing_UTC":flight_info[8],

                        "on_block_UTC":flight_info[9],

                        "takeoff": flight_info[10],

                        "landing": flight_info[11],
                        })


# READS CX LOGBOOK FORMAT HERE, CHANGE DIRECTORY, ADD CSV READER IF MULTIPLE LOGBOOKS EXIST.
for path in logbook_paths:
    with open(path, encoding="utf8") as logbook1:
        reader = csv.reader(logbook1)
        logger(reader)


for i in range(len(log)):
    flight = log[i]
    if flight['isFlightDuty']:
        departure_date = flight['departure_date']
        origin = flight['origin']
        dest = flight['dest']
        off_block_UTC = flight['off_block_UTC']
        airborne_UTC = flight['airborne_UTC']
        landing_UTC = flight['landing_UTC']
        on_block_UTC = flight['on_block_UTC']
        log[i]['day'], log[i]['night'] = caldaynight(origin,dest,airborne_UTC,landing_UTC,departure_date,off_block_UTC,on_block_UTC,p2x)

total_day = 0
total_night = 0
total_sectors = 0

for r in log:
    if "day" in r:
        total_day += r['day']
    if "night" in r:
        total_night += r['night']
    if r["isFlightDuty"]:
        total_sectors += 1
    print(r)
total = total_day + total_night

print(total,total_day,total_night,total_sectors)

# daynight hours, change resulting file name as required
with open(f"./results/daynighthours-{name}.csv", "w", newline='') as file:
    writer = csv.writer(file)
    field = ['isFlightDuty','departure date UTC','type','registration','pilot-in-command','origin','dest','off block UTC','airborne UTC','landing UTC','on block UTC','day','night','duty code']
    writer.writerow(field)
    for flight in log:
        if flight['isFlightDuty']:
            writer.writerow([True,flight['departure_date'],flight['type'],flight['reg'],flight['pic'],flight['origin'],
                             flight['dest'],flight['off_block_UTC'],flight['airborne_UTC'],flight['landing_UTC'],
                            flight['on_block_UTC'],flight['day'],flight['night'],''])
        else:
            writer.writerow([False,flight['departure_date'],"","","","","","","","","","","",flight['duty_code']])
    writer.writerow(["","","","","","","","","","","total day/night",total_day,total_night])
    writer.writerow(["","","","","","","","","","","","total hours:",total])
    writer.writerow(["", "", "", "","", "","","", "", "", "", "total sectors:", total_sectors])


# REPORT LEFT PAGE, change resulting file name as required
with open(f"./results/{name}_report_left.csv", "w", newline="") as file:
    writer = csv.writer(file)
    fields = ['Year/20XX','Month/Date','Type','Registration','Pilot-in-command','Co-pilot or student',
              "Holder's operating capacity",'  From   To ','Take-offs','Landings']
    writer.writerow(fields)
    for flight in log:
        if flight['isFlightDuty']:
            if p2x:
                writer.writerow([flight['departure_date'].split('/')[0],
                                 flight['departure_date'].split('/')[1] + '/' + flight['departure_date'].split('/')[2],
                                 flight['type'],
                                 flight['reg'],
                                 flight['pic'],
                                 'Self',
                                 'P2X',
                                 f'{flight["origin"]}    {flight["dest"]}',
                                 flight['takeoff'],
                                 flight['landing']
                                 ])
            elif flight['takeoff'] or flight['landing']:
                # PF sector
                writer.writerow([flight['departure_date'].split('/')[0],
                                 flight['departure_date'].split('/')[1] + '/' + flight['departure_date'].split('/')[2],
                                 flight['type'],
                                 flight['reg'],
                                 flight['pic'],
                                 'Self',
                                 'P1/US',
                                 f'{flight["origin"]}    {flight["dest"]}',
                                 flight['takeoff'],
                                 flight['landing']
                                 ])
            else:
                # PM sector
                writer.writerow([flight['departure_date'].split('/')[0],
                                 flight['departure_date'].split('/')[1] + '/' + flight['departure_date'].split('/')[2],
                                 flight['type'],
                                 flight['reg'],
                                 flight['pic'],
                                 'Self',
                                 'P2',
                                 f'{flight["origin"]}    {flight["dest"]}',
                                 flight['takeoff'],
                                 flight['landing']
                                 ])

        else:
            writer.writerow([flight['departure_date'].split('/')[0],
                             flight['departure_date'].split('/')[1] + '/' + flight['departure_date'].split('/')[2],
                             'B777',
                             'CPA',
                             "",
                             "Self",
                             'P/UT',
                             "",
                             "",
                             ""
            ])


# REPORT RIGHT PAGE, change resulting file name as required
with open(f"./results/{name}_report_right.csv", "w", newline="") as file:
    writer = csv.writer(file)
    fields = ['P1 day','P2 day', 'P2X day', 'Dual day','P1 night','P2 night', 'P2X night', 'Dual night',
              'Instrument Flying','Simulator Time','Remarks']
    writer.writerow(fields)
    for flight in log:
        if flight["isFlightDuty"]:
            if p2x:
                writer.writerow([
                    "",
                    "",
                    flight['day'],
                    "",
                    "",
                    "",
                    flight['night'],
                    "",
                    flight['day'] + flight['night'],
                    "",
                    ""
                ])
            elif flight['takeoff'] or flight['landing']:
                #PF sector, log P1/US (P1 column)
                writer.writerow([
                    flight['day'],
                    "",
                    "",
                    flight['night'],
                    "",
                    "",
                    "",
                    "",
                    flight['day'] + flight['night'],
                    "",
                    ""
                ])
            else:
                # PM sector, log P2
                writer.writerow([
                    "",
                    flight['day'],
                    "",
                    "",
                    flight['night'],
                    "",
                    "",
                    "",
                    flight['day'] + flight['night'],
                    "",
                    ""
                ])
        else:
            writer.writerow([
                "","","","","","","","",
                "2",
                "4",
                flight['duty_code']
            ])


# Generate CAD Format report
# calculate maximum pages needed for the report
first_flight_dept_year, final_flight_dept_year = int(log[0]['departure_date'][:4]), int(log[-1]['departure_date'][:4])
total_pages = math.ceil(len(log)/19) + final_flight_dept_year - first_flight_dept_year
wb = load_workbook(filename='./HKCAD_logbook_format.xlsx')
sheet = wb['Sheet1']
column_range = {
    1: "A",
    2: "B",
    3: "C",
    4: "D",
    5: "E",
    6: "F",
    7: "G",
    8: "H",
    9: "I",
    10: "J",
    11: "K",
    12: "L",
    13: "M",
    14: "N",
    15: "O",
    16: "P",
    17: "Q",
    18: "R",
    19: "S",
    20: "T",
}
current_year = int(log[0]['departure_date'][:4])
current_year_page = 1
current_flight = 0

# Create required pages
print(f'total pages{total_pages}')
for i in range(2, total_pages):
    ws = wb.copy_worksheet(wb['Sheet1'])

logbook_filled = False
for sheet in wb:
    if logbook_filled:
        break
    current_flight_dept_year = int(log[current_flight]['departure_date'][:4])
    if current_flight_dept_year != current_year:
        current_year += 1
        current_year_page = 1
        continue
    sheet.title = f"{current_year}-{current_year_page}"
    sheet_filled_up = False
    for r in sheet:
        if sheet_filled_up or logbook_filled:
            break
        if r[0].row >= 24:
            break
        for cell in r:
            if current_flight == len(log):
                logbook_filled = True
                break
            if cell.row >= 24:
                sheet_filled_up = True
                break
            if current_year != int(log[current_flight]['departure_date'][:4]):
                current_year += 1
                current_year_page = 0
                sheet_filled_up = True
                break
            if column_range[cell.column] == "A" and cell.row == 4:
                cell.value = current_year
            if 4 < cell.row < 24:
                if column_range[cell.column] == "A":
                    cell.value = log[current_flight]['departure_date'][5:]
                elif column_range[cell.column] == "B" and log[current_flight]['isFlightDuty']:
                    cell.value = log[current_flight]['type']
                elif column_range[cell.column] == "C" and log[current_flight]['isFlightDuty']:
                    cell.value = log[current_flight]['reg']
                elif column_range[cell.column] == "D" and log[current_flight]['isFlightDuty']:
                    cell.value = log[current_flight]['pic']
                elif column_range[cell.column] == "E" and log[current_flight]['isFlightDuty']:
                    cell.value = "Self"
                elif column_range[cell.column] == "F" and log[current_flight]['isFlightDuty']:
                    if p2x:
                        cell.value = "P2X"
                    elif log[current_flight]['takeoff'] or log[current_flight]['landing']:
                        cell.value = "P1/US"
                    else:
                        cell.value = "P2"
                elif column_range[cell.column] == "F" and not log[current_flight]['isFlightDuty']:
                    cell.value = "P/UT"
                elif column_range[cell.column] == "G" and log[current_flight]['isFlightDuty']:
                    cell.value = f'{log[current_flight]["origin"]}    {log[current_flight]["dest"]}'

                elif column_range[cell.column] == "I" and log[current_flight]['isFlightDuty']:
                    if not p2x and (log[current_flight]['takeoff'] or log[current_flight]['landing']):
                        cell.value = log[current_flight]['day']

                elif column_range[cell.column] == "J" and log[current_flight]['isFlightDuty']:
                    if not p2x and not log[current_flight]['takeoff'] and not log[current_flight]['landing']:
                        cell.value = log[current_flight]['day']
                elif column_range[cell.column] == "K" and log[current_flight]['isFlightDuty'] and p2x:
                    cell.value = log[current_flight]['day']

                elif column_range[cell.column] == "M" and log[current_flight]['isFlightDuty']:
                    if not p2x and (log[current_flight]['takeoff'] or log[current_flight]['landing']):
                        cell.value = log[current_flight]['night']
                elif column_range[cell.column] == "N" and log[current_flight]['isFlightDuty']:
                    if not p2x and not log[current_flight]['takeoff'] and not log[current_flight]['landing']:
                        cell.value = log[current_flight]['night']
                elif column_range[cell.column] == "O" and log[current_flight]['isFlightDuty'] and p2x:
                    cell.value = log[current_flight]['night']

                elif column_range[cell.column] == "Q" and log[current_flight]['isFlightDuty']:
                    cell.value = log[current_flight]['day'] + log[current_flight]['night']
                elif column_range[cell.column] == "Q" and not log[current_flight]['isFlightDuty']:
                    cell.value = 2
                elif column_range[cell.column] == "R" and not log[current_flight]['isFlightDuty']:
                    cell.value = 4
                elif column_range[cell.column] == "S":
                    if not log[current_flight]['isFlightDuty']:
                        cell.value = log[current_flight]['duty_code']
                    current_flight += 1

    current_year_page += 1

wb.save(f"././results/{name}-report.xlsx")