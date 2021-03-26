from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
import os
import xlrd
import xlsxwriter

root = Tk()
root.title("CivilCon")
root.iconbitmap("CC.ico")
root.geometry("500x500")


class CivilCon:

    def __init__(self, master): #First Page
        self.master = master
        Label(self.master, text = "Kaç oturum var?").grid(row = 0, column = 0)

        self.clicked = StringVar()
        OptionMenu(self.master, self.clicked, "1", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10").grid( row = 0, column = 1)

        Button(self.master, text = "Seç", command = self.session).grid(row = 0, column = 4)
        Button(self.master, text = "Excel Dosyası", command = self.Excel).grid(row = 0, column = 3)

    def Excel(self):
        self.attachment_file_directory = filedialog.askopenfilename(initialdir = os.path, title = "Excel")
    
    def session(self):
        try:
            if self.attachment_file_directory[-3:] == "xls":
                for widget in self.master.winfo_children():
                    widget.destroy()

                variables_for_dict = []
                key_of_dict = []
                        
                for x in range(int(self.clicked.get())):
                    variables_for_dict.append("self.clicked_version1"+str(x))
                    variables_for_dict.append("self.clicked_version2"+str(x))
                    variables_for_dict.append("self.clicked_version3"+str(x))
                    variables_for_dict.append("self.clicked_version4"+str(x))
                    key_of_dict.append(StringVar())
                    key_of_dict.append(StringVar())
                    key_of_dict.append(StringVar())
                    key_of_dict.append(StringVar())

                self.variable_dictionary = dict(zip(variables_for_dict, key_of_dict))
                        
                Label(self.master, text = "Başlangıç").grid(row = 0, column = 1)  
                Label(self.master, text = "|").grid(row = 0, column = 3)              
                Label(self.master, text = "Bitiş").grid(row = 0, column = 4)
                Label(self.master, text = "Saat").grid(row = 1, column = 1)                
                Label(self.master, text = "Dakika").grid(row = 1, column = 2)
                Label(self.master, text = "|").grid(row = 1, column = 3)
                Label(self.master, text = "Saat").grid(row = 1, column = 4)                
                Label(self.master, text = "Dakika").grid(row = 1, column = 5)
                
                for x in range(int(self.clicked.get())):
                    
                    Label(self.master, text = str(x+1) + ". Oturum").grid(row = x+2, column = 0)
                    OptionMenu(self.master, self.variable_dictionary["self.clicked_version1"+str(x)] , "01", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24").grid( row = x+2, column = 1)
                    OptionMenu(self.master, self.variable_dictionary["self.clicked_version2"+str(x)], "00", "00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59").grid( row = x+2, column = 2)
                    Label(self.master, text = "|").grid(row = x+2, column = 3)
                    OptionMenu(self.master, self.variable_dictionary["self.clicked_version3"+str(x)], "01", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24").grid( row = x+2, column = 4)
                    OptionMenu(self.master, self.variable_dictionary["self.clicked_version4"+str(x)], "00", "00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59").grid( row = x+2, column = 5)

                Button(self.master, text = "Başlat", command = self.start).grid(row = int(self.clicked.get())+10, column = 5)
            else:
                self.Excel()
        except:
            self.Excel()
            
    def start(self):
        sessions = []
        for k, v in self.variable_dictionary.items():
            sessions.append(v.get())
            
        sessions_vol2 = []
        for x in range(len(sessions)):
            if x%2 == 0:
                try:
                    sessions_vol2.append(sessions[x]+":"+sessions[x+1])
                except:
                    sessions_vol2.append(sessions[-2]+":"+sessions[-1])
        sessions = sessions_vol2

        try:
            path = self.attachment_file_directory
        except:
            self.Excel()

        for widget in self.master.winfo_children():
            widget.destroy()

        Label(self.master, text = "Kaç oturum var?").grid(row = 0, column = 0)

        self.clicked = StringVar()
        OptionMenu(self.master, self.clicked, "1", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10").grid( row = 0, column = 1)

        Button(self.master, text = "Seç", command = self.session).grid(row = 0, column = 4)
        Button(self.master, text = "Excel Dosyası", command = self.Excel).grid(row = 0, column = 3)

        attendees = []
        inputWorkbook = xlrd.open_workbook(path)
        inputWorksheet = inputWorkbook.sheet_by_index(0)

        for x in range(inputWorksheet.nrows-4):
            x += 4
            attendees.append(inputWorksheet.cell_value(x,0))

        attendees.sort()

        attendees_list_form = []

        for x in attendees:
            x = "CC | "+x
            attendees_list_form.append(x.split(","))

        for x in attendees_list_form:
            for k in range(len(x)):
                if x[k] == "":
                    x[k] = "info@hmaksu.com"
                    
        attendees_vol2 = []

        k = 0

        for x in range(len(attendees_list_form)):
            attendees_list_form[x].pop()
            attendees_list_form[x].pop()
            try:
                if attendees_list_form[x][0] == attendees_list_form[x+1][0] or attendees_list_form[x][1] == attendees_list_form[x+1][1]:
                    k += 1
                    continue

                else:
                    if k == 0:
                        attendee = attendees_list_form[x]
                        attendee.sort()
                        attendees_vol2.append(attendees_list_form[x])
                    else:
                        attendee = attendees_list_form[x]
                        
                        for t in range(k):
                            if k == t:
                                continue
                            else:
                                t += 1
                                attendee.append(attendees_list_form[x-t][-1])
                                attendee.append(attendees_list_form[x-t][-2])
                        
                        attendee.sort()
                        attendees_vol2.append(attendee)
                    k = 0
                    
            except:
                if k == 0:
                    attendee = attendees_list_form[x]
                    attendee.sort()
                    attendees_vol2.append(attendees_list_form[x])
                else:
                    attendee = attendees_list_form[x]
                    for t in range(k):
                        if k == t:
                            continue
                        else:
                            t += 1
                            attendee.append(attendees_list_form[x-t][-1])
                            attendee.append(attendees_list_form[x-t][-2])
                        
                    attendee.sort()
                    attendees_vol2.append(attendee)

        attendee = []
        attendees = []
        attendee_vol3 = []
        attendees_vol3 = []

        for x in attendees_vol2:
            attendee.append(x[-2])
            attendee.append(x[-1])
            attendee.append(x[0].split()[1])
            attendee.append(x[-3].split()[1])
            attendees.append(attendee)
            attendee = []

            attendee_vol3.append(x[-2])
            attendee_vol3.append(x[-1])
            for t in x:
                if x[-2] == t or x[-1] == t:
                    continue
                else:
                    attendee_vol3.append(t.split()[1])
            attendees_vol3.append(attendee_vol3)
            attendee_vol3 = []

        outworkbook = xlsxwriter.Workbook("Sheet.xlsx")
        outworksheet = outworkbook.add_worksheet()
        outworksheet.write(0, 0, "İsim-Soyisim")
        outworksheet.write(0, 1, "E-Posta Adresi")

        sessions_vol2 = []
        for x in range(len(sessions)):
            try:
                if x%2 == 0:
                    sessions_vol2.append(sessions[x]+" - "+sessions[x+1])
            except:
                sessions_vol2.append(sessions[-2]+" - "+sessions[-1])
                
        sessions = sessions_vol2
        
        for x in range(len(sessions)):
            outworksheet.write(0, x+2, str(x+1)+". Oturum "+sessions[x])

        for x in range(len(attendees)):
            for k in range(len(attendees[x])):
                if k < 2:
                    outworksheet.write(x+1, k, attendees[x][k])

                for t in range(len(sessions)):
                    #print("="*30)
                    #print(attendees[x][3])
                    #print(attendees[x][2])
                    #print(sessions[t])
                    #print("="*30)
                    if int(attendees[x][3].replace(":","")[:-2]) < int(sessions[t].replace(":","")[:-7]) or int(attendees[x][2].replace(":","")[:-2]) > int(sessions[t].replace(":","")[7:]):
                        outworksheet.write(x+1, t+2, "Katılmadı")
                    else:
                        outworksheet.write(x+1, t+2, "Katıldı")                        
                        
        outworksheet.write(0, len(sessions)+2, "Toplam Süre")
        for x in range(len(attendees_vol3)):
            total_time = 0
            for t in range(len(attendees_vol3[x])):
                if t == 0 or t == 1:
                    continue
                elif t%2 != 0:
                    total_time += int(attendees_vol3[x][t].replace(":","")[:2])*60+int(attendees_vol3[x][t].replace(":","")[2:4])-int(attendees_vol3[x][t-1].replace(":","")[:2])*60-int(attendees_vol3[x][t-1].replace(":","")[2:4])
                    
            outworksheet.write(x+1, len(sessions)+2, str(total_time))
        
        outworkbook.close()
            
e = CivilCon(root)

root.mainloop()
