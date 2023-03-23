from pathlib import Path

import PySimpleGUI as sg
import pandas as pd
from datetime import datetime
harini = datetime.now()
charini = datetime.strftime(harini,"%d-%m-%Y")

def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    if not filepath:
        sg.popup("Sila klik pilih untuk memilih file / folder")
        return False
    sg.popup("File atau Folder mengarut.")
    return False


ver = [
    [sg.Text("Ver 1.0 2023, www.resakse.com", justification="right")],
]
layout = [
    [sg.Text("File Statistik"), sg.Push(), sg.Input(key="-IN-"),
     sg.FileBrowse("Pilih", file_types=(("Excel Files", "*.xls*"),))],
    [sg.Text("Folder untuk di Save"), sg.Push(), sg.Input(key="-OUT-"), sg.FolderBrowse("Pilih")],
    [sg.Exit("Keluar", s=16, button_color="tomato"), sg.Button("Proses"),
     sg.Column(ver, element_justification="right", expand_x=True)],
]


def prosesfile(ct_input, ct_output):
    df = pd.read_excel(ct_input)
    filename = Path(ct_input).stem
    outputfile = Path(ct_output) / f"{filename}-convert-{charini}.xlsx"
    # convert datetime
    df['DATE_OF_BIRTH'] = df['DATE_OF_BIRTH'].dt.strftime('%d/%m/%Y')
    #df['ADDED_DATE'] = df['ADDED_DATE'].dt.strftime('%d/%m/%Y %H')
    df['NATIONAL_ID_NO'] = df['NATIONAL_ID_NO'].astype(str).str.replace('.0', '', regex=False)
    # buang column tak berkenaan
    del df['ORDER_TYPE_CODE']
    del df['EXAM_CODE']
    del df['DATE_OF_BIRTH']
    del df['STATUS']
    del df['RADIOLOGIST_ID']
    ### Problem sebab 10.58am - 11.01am
    unik = df[df['SHORT_DESC'] != 'CT Lungs (HRCT)']  # filter yang bukan hrct lung
    # akhir = unik.drop_duplicates(
    #     ['PATIENT_ID', 'ADDED_DATE'])  # buang duplicate rekod pt id yang sama dlm jam yang sama
    # akhir.to_excel(outputfile, index=False)
    df = unik.sort_values(by=['ADDED_DATE','PATIENT_ID']) #sort kemudian shift & mark duplicate < 1 jam
    mask = (df['PATIENT_ID'] == df['PATIENT_ID'].shift(1)) & (abs(df['ADDED_DATE'] - df['ADDED_DATE'].shift(1)) < pd.Timedelta('1 hour'))
    df.loc[mask, 'to_remove'] = True
    df = df[df.to_remove != True] # filter out yang kena mark
    del df['to_remove'] # buang column
    df.to_excel(outputfile, index=False)
    sg.popup("Siap! :)")


window = sg.Window("Convert Statistik CT Scan", layout, icon=r'cticon.ico')

while True:
    event, values = window.read()
    print(event, values)
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break
    if event == "Proses":
        if (is_valid_path(values["-IN-"])) and (is_valid_path(values["-OUT-"])):
            prosesfile(values["-IN-"], values["-OUT-"])
window.close()
