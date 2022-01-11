import openpyxl
from pathlib import Path

NUM_ROWS = 50
XML_TEMPLATE ="""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<TipsContents xmlns="http://www.avendasys.com/tipsapiDefs/1.0">
  <TipsHeader exportTime="Tue Jan 11 11:19:49 CET 2022" version="6.9"/>
  <LocalUsers>{}
  </LocalUsers>
</TipsContents>"""

LOCALUSER_TEMPLATE = """
    <LocalUser changePwdNextLogin="false" enabled="true" roleName="FARMTEC MAC Auth" password="{2}" userName="{2}" userId="{2}">
      <LocalUserTags tagName="IP" tagValue="{1}"/>
      <LocalUserTags tagName="Description" tagValue="{0}"/>
    </LocalUser>"""

NUM_CELL_DESCRIPTION = 0
NUM_CELL_IP = 2
NUM_CELL_MAC = 4


def normalize_mac(mac):
    mac = mac.strip().lower()
    return "{}{}{}{}{}{}".format(mac[0:2], mac[3:5], mac[6:8], mac[9:11], mac[12:14], mac[15:17],)

def read_excel():
    global NUM_ROWS
    devices = []

    xlsx_file = Path('data', 'printers.xlsx')
    wb_obj = openpyxl.load_workbook(xlsx_file)
    sheet = wb_obj.active

    count = 0
    for row in sheet.iter_rows(max_row=NUM_ROWS):
        if count > 0:
            cellDescription = row[NUM_CELL_DESCRIPTION].value
            cellIP = row[NUM_CELL_IP].value
            cellMAC = row[NUM_CELL_MAC].value
            if (cellIP is not None) and (cellMAC is not None):
                device = {}
                device['printer'] = cellDescription.strip()
                device['ip'] = cellIP.strip()
                device['mac'] = normalize_mac(cellMAC)
                devices.append(device)
        count += 1

    return devices

def main():
    devices = read_excel()
    localUsers = ""
    for device in devices:
        localUsers += LOCALUSER_TEMPLATE.format(device["printer"], device["ip"], device["mac"])

    print(XML_TEMPLATE.format(localUsers))

main()