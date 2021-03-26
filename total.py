import xlrd
import xlsxwriter

inputWorkbook = xlrd.open_workbook("Tüm.xls")

attendees = []

for x in range(4):
    inputWorksheet = inputWorkbook.sheet_by_index(x)
    for t in range(inputWorksheet.nrows-1):
        t += 1
        attendee = []
        attendee.append(inputWorksheet.cell_value(t,0))
        attendee.append(inputWorksheet.cell_value(t,1))
        attendee.append(inputWorksheet.cell_value(t,3))
        attendees.append(attendee)
    

attendees.sort()

k = 0
attendees_vol2 = []
attendee = []

for x in range(len(attendees)):
    try:
        if attendees[x][1] == "info@hmaksu.com":
            if attendees[x][0].lower() == attendees[x+1][0].lower():
                k += 1
            else:
                if k == 0:
                    attendees[x].append(int(attendees[x][-1]))
                    attendees_vol2.append(attendees[x])
                else:
                    for q in range(len(attendees[x])):
                        if attendees[x] != attendees[x][-1]:
                            attendee.append(attendees[x][q])
                    total_time = 0
                    for t in range(k+1):
                        total_time += int(attendees[x-t][-1])
                    attendee.append(total_time)
                    
                    attendees_vol2.append(attendee)
                    attendee = []
                    k = 0
        else:
            if attendees[x][0].lower() == attendees[x+1][0].lower() or attendees[x][1] == attendees[x+1][1]:
                k += 1
            else:
                if k == 0:
                    attendees[x].append(int(attendees[x][-1]))
                    attendees_vol2.append(attendees[x])
                else:
                    for q in range(len(attendees[x])):
                        if attendees[x] != attendees[x][-1]:
                            attendee.append(attendees[x][q])
                    total_time = 0
                    for t in range(k+1):
                        total_time += int(attendees[x-t][-1])
                    attendee.append(total_time)
                    
                    attendees_vol2.append(attendee)
                    attendee = []
                    k = 0
    except:
        if k == 0:
            attendees[x].append(int(attendees[x][-1]))
            attendees_vol2.append(attendees[x])
        else:
                for q in range(len(attendees[x])):
                    if attendees[x] != attendees[x][-1]:
                        attendee.append(attendees[x][q])
                total_time = 0
                for t in range(k+1):
                    total_time += int(attendees[x-t][-1])
                attendee.append(total_time)
                
                attendees_vol2.append(attendee)
                attendee = []

outworkbook = xlsxwriter.Workbook("Tüm Edit.xlsx")
outworksheet = outworkbook.add_worksheet()

outworksheet.write(0, 0, "İsim Soyisim")
outworksheet.write(0, 1, "E-Posta Adresi")
outworksheet.write(0, 2, "Toplam Süre")

for x in range(len(attendees_vol2)):
    outworksheet.write(x+1, 0, attendees_vol2[x][0])
    outworksheet.write(x+1, 1, attendees_vol2[x][1])
    outworksheet.write(x+1, 2, attendees_vol2[x][3])
    

outworkbook.close()
