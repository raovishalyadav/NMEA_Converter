'''
Copyright (C) 2021 Vishal Yadav
'''
import xlsxwriter

f1 = open("NMEA.txt","a+")
f2 = open("NMEA_converted.txt","a+")
with open("NMEA.txt", "r") as infile:
    lines = infile.readlines()

    for x in lines:
        if '$GNGGA' in x:
            f2.write(" "+str(x.split(',')[1])+" "+str(x.split(',')[2])+str(x.split(',')[3])+" "+str(x.split(',')[4])+str(x.split(',')[5])+" "+str(x.split(',')[9])+" "+str(x.split(',')[7])+" "+" "+" "+" "+" "+" "+" "+" "+str(x.split(',')[8])+"\n")

        if '$GNGLL' in x:
            f2.write(" "+str(x.split(',')[5])+" "+str(x.split(',')[1])+str(x.split(',')[2])+" "+str(x.split(',')[3])+str(x.split(',')[4])+"\n")

        if ('$GNGSA' in x) and (str(x.split(',')[-1][0:1]) == str('1')):
            f2.write(" "+" "+" "+" "+" "+" "+str('GPS')+" "+str(x.split(',')[2])+" "+str(x.split(',,,,,,')[0].split(',')[3:]).replace('[','').replace(']','').replace("'",'').replace(" ",'')+" "+" "+" "+" "+str(x.split(',,,,,,')[1].split(',')[0])+" "+str(x.split(',,,,,,')[1].split(',')[1])+" "+str(x.split(',,,,,,')[1].split(',')[2])+"\n")

        if ('$GNGSA' in x) and (str(x.split(',')[-1][0:1]) == str('4')):
            f2.write(" "+" "+" "+" "+" "+" "+str('NAVIC')+" "+str(x.split(',')[2])+" "+str(x.split(',,,,,,')[0].split(',')[3:]).replace('[','').replace(']','').replace("'",'').replace(" ",'')+" "+" "+" "+" "+str(x.split(',,,,,,,,,,,,')[1].split(',')[0])+" "+str(x.split(',,,,,,,,,,,,')[1].split(',')[1])+" "+str(x.split(',,,,,,,,,,,,')[1].split(',')[2])+"\n")

        if '$GPGSV' in x:
            s = int((x.count(',')-4)/4)
            for l in range(1,s+1):
                f2.write(" "+" "+" "+" "+" "+str(x.split(',')[3])+" "+str('GPS')+" "+" "+str(x.split(',')[4*l])+" "+str(x.split(',')[(4*l)+1])+" "+str(x.split(',')[(4*l)+2])+" "+str(x.split(',')[(4*l)+3])+" "+"\n")

        if '$GIGSV' in x:
            s = int((x.count(',')-4)/4)
            for l in range(1,s+1):
                f2.write(" "+" "+" "+" "+" "+str(x.split(',')[3])+" "+str('NAVIC')+" "+" "+str(x.split(',')[4*l])+" "+str(x.split(',')[(4*l)+1])+" "+str(x.split(',')[(4*l)+2])+" "+str(x.split(',')[(4*l)+3])+" "+"\n")

        if '$GNRMC' in x:
                f2.write(str(x.split(',')[9])+" "+str(x.split(',')[1])+" "+str(x.split(',')[3])+str(x.split(',')[4])+" "+str(x.split(',')[5])+str(x.split(',')[6])+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+str(x.split(',')[7])+" "+str(x.split(',')[8])+" "+"\n")

        if '$GNVTG' in x:
                f2.write(" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+str(x.split(',')[5])+" "+str(x.split(',')[1])+" "+"\n")

        if '$GNZDA' in x:
                f2.write(" "+str(x.split(',')[1])+" "+" "+"\n")

        if '$PIRNSF' in x:
                f2.write(" fffffffffffff"+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+" "+str(x.split(',')[1])+" "+str(x.split(',')[2])+" "+str(x.split(',')[3:-1]).replace('[','').replace(']','').replace("'",'').replace(" ",'')+" "+str(x.split(',')[-1].split('*')[0])+" "+"\n")

def Txt2Xlsx(self, data, row = 0):
    for colNum, value in enumerate(data):
            self.write(row, colNum, value)
xlsxwriter.worksheet.Worksheet.addRow = Txt2Xlsx
workbook = xlsxwriter.Workbook("NMEA_converted.xlsx")
worksheet = workbook.add_worksheet()
wrap = workbook.add_format()
format = workbook.add_format({'bold': 1,'align': 'center','valign': 'vcenter','text_wrap': True})
worksheet.set_column(0,0,8)
worksheet.set_column(1,1,10)
worksheet.set_column(2,2,12)
worksheet.set_column(3,3,12)
worksheet.set_column(4,4,8)
worksheet.set_column(5,5,11)
worksheet.set_column(6,6,11)
worksheet.set_column(7,7,8)
worksheet.set_column(8,8,18)
worksheet.set_column(9,9,10)
worksheet.set_column(10,10,10)
worksheet.set_column(11,11,8)
worksheet.set_column(12,12,8)
worksheet.set_column(13,13,8)
worksheet.set_column(14,14,8)
worksheet.set_column(19,19,19)
worksheet.set_column(20,20,19)
worksheet.set_column(21,21,19)

worksheet.write('A1', 'UTC Date', format)
worksheet.write('B1', 'UTC Time', format)
worksheet.write('C1', 'Latitude', format)
worksheet.write('D1', 'Longitude', format)
worksheet.write('E1', 'Altitude', format)
worksheet.write('F1', 'Satellites in View', format)
worksheet.write('G1', 'GPS/NAVIC', format)
worksheet.write('H1', 'Position \n Fixed', format)
worksheet.write('I1', 'Satellite ID', format)
worksheet.write('J1', 'Elevation', format)
worksheet.write('K1', 'Azimuth', format)
worksheet.write('L1', 'SNR \n (in dB)', format)
worksheet.write('M1', 'PDOP', format)
worksheet.write('N1', 'HDOP', format)
worksheet.write('O1', 'VDOP', format)
worksheet.write('P1', 'Speed Over Ground \n (in Kmph)', format)
worksheet.write('Q1', 'Course Over Speed (in Degrees)', format)
worksheet.write('R1', 'SVID NAVIC Satellite PRN', format)
worksheet.write('S1', 'SFID', format)
worksheet.write('T1', 'Sub-Frame ', format)
worksheet.write('U1', 'SF Data Decoded \n Sub-Frame data Tail Bits', format)
worksheet.freeze_panes(1, 1)
row = 1
with open('NMEA_converted.txt', 'rt+') as f6:
    lines = f6.readlines()
    for line in lines:
        worksheet.addRow(data = line.split(" "), row = row)
        row += 1

my_format = workbook.add_format({'align': 'center','valign': 'vcenter'})

worksheet.set_column('A:XFD', None, my_format)
workbook.close()
