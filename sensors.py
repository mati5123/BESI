import csv
import os
import openpyxl

excel_file = r'D:\CSV\Sensory.xlsx'

workbook = openpyxl.load_workbook(excel_file)
worksheet = workbook.active

names = [cell.value for cell in worksheet['A']][1:]
maxvolume = [cell.value for cell in worksheet['B']][1:]
Density = [cell.value for cell in worksheet['C']][1:]
SensorRange = [cell.value for cell in worksheet['D']][1:]
MountingHeight = [cell.value for cell in worksheet['E']][1:]
MaxHeigh = [cell.value for cell in worksheet['F']][1:]

base_folder = r'D:\CSV\Tanks'
os.makedirs(base_folder, exist_ok=True)

for sensors, name in enumerate(names):
    csv_name = f'TK{sensors + 1}.csv'
    csv_file = os.path.join(base_folder, csv_name)
    csv_file_obj = open(csv_file, 'w', newline='')
    writer = csv.writer(csv_file_obj)

    writer.writerow([name + ';'])                               #1 sName
    writer.writerow([str(maxvolume[sensors]) + ';'])            #2 rMaxVolume
    writer.writerow([str(Density[sensors]) + ';'])              #3 iDensity
    writer.writerow([str(SensorRange[sensors]) + ';'])          #4 iSensorRange
    writer.writerow([str(MountingHeight[sensors]) + ';'])       #5 iMountingHeight
    writer.writerow(["90;"])                                    #6 iHighLevePer
    writer.writerow(["5;"])                                     #7 iLowLevelPer
    writer.writerow(["0;"])                                     #8 iRangeCorrection
    writer.writerow([str(int(MaxHeigh[sensors])/100) + ';'])    #9 rMaxHeight
    writer.writerow(["95;"])                                    #10 iHighHighLevelPer
    writer.writerow(["10;"])                                    #11 iAlarmTimeDelay
    writer.writerow(["5;"])                                     #12 iTabStep
    writer.writerow(["80;"])                                    #13 rTemp

    csv_file_obj.close()

print(f'Pliki csv z sensorami są tutaj: {base_folder}')