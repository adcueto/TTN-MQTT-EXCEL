import time
import ttn
import openpyxl
import base64
from base64 import b64decode
global iRow, nGw;

app_id = "beacons"  #The app name
access_key = "ttn-account-v2.ErABTkHwaF7QCVuIY0ScHh9ezPnMCaUnwhcD8DqPIo8" # The app key

dest_filename = "LoRa.xlsx"  #The name of File Excel

doc = openpyxl.load_workbook(filename=dest_filename) #Open File Excel
uSheet = doc.worksheets[0]  #"Uplink" Sheet
iRow=3 #begin of row third

#================= Function uplink Callback ===========================================
def uplink_callback(msg, client):
    global iRow, nGw; #variable global
    print("Received uplink from ", msg.dev_id)
    iColGw=21 #gateways Column
    nGw = len(msg.metadata.gateways) # gateways number

    #===== General processing ==================================================
    uSheet.cell(row=iRow, column=1).value = msg.app_id
    uSheet.cell(row=iRow, column=2).value = msg.dev_id
    uSheet.cell(row=iRow, column=3).value = msg.hardware_serial
    uSheet.cell(row=iRow, column=4).value = msg.port
    uSheet.cell(row=iRow, column=5).value = msg.counter
    uSheet.cell(row=iRow, column=6).value = b64decode(msg.payload_raw).hex() #convert payload from base64 to hexadecimal
    #===== Payload processing ==================================================
    uSheet.cell(row=iRow, column=7).value = msg.payload_fields.batV
    uSheet.cell(row=iRow, column=8).value = msg.payload_fields.fixFailed
    uSheet.cell(row=iRow, column=9).value = msg.payload_fields.inTrip
    uSheet.cell(row=iRow, column=10).value = msg.payload_fields.headingDeg
    uSheet.cell(row=iRow, column=11).value = msg.payload_fields.latitudeDeg
    uSheet.cell(row=iRow, column=12).value = msg.payload_fields.longitudeDeg
    uSheet.cell(row=iRow, column=13).value = msg.payload_fields.speedKmph
    #====== Metadata processing ===============================================
    uSheet.cell(row=iRow, column=14).value = msg.metadata.time
    uSheet.cell(row=iRow, column=15).value = msg.metadata.frequency
    uSheet.cell(row=iRow, column=16).value = msg.metadata.modulation
    uSheet.cell(row=iRow, column=17).value = msg.metadata.data_rate
    uSheet.cell(row=iRow, column=18).value = msg.metadata.airtime
    uSheet.cell(row=iRow, column=19).value = msg.metadata.coding_rate
    uSheet.cell(row=iRow, column=20).value = nGw # gateways number
    #====== Gateways processing ================================================
    for iGw in range(nGw):
        uSheet.cell(row=iRow, column=iColGw+0).value = msg.metadata.gateways[iGw].gtw_id
        uSheet.cell(row=iRow, column=iColGw+1).value = msg.metadata.gateways[iGw].timestamp
        uSheet.cell(row=iRow, column=iColGw+2).value = msg.metadata.gateways[iGw].time
        uSheet.cell(row=iRow, column=iColGw+3).value = msg.metadata.gateways[iGw].channel
        uSheet.cell(row=iRow, column=iColGw+4).value = msg.metadata.gateways[iGw].rssi
        uSheet.cell(row=iRow, column=iColGw+5).value = msg.metadata.gateways[iGw].snr
        #gSheet.cell(row=iRow, column=iColGw+6).value = msg.metadata.gateways[iGw].latitude
        #gSheet.cell(row=iRow, column=iColGw+7).value = msg.metadata.gateways[iGw].longitude
        iColGw+=8; #column increment
    #====== End processing =====================================================
    doc.save(filename = dest_filename) #save file
    print("File updated")
    iRow+=1 #row increment

#================= End Function ================================================

handler = ttn.HandlerClient(app_id, access_key)
# using mqtt client
mqtt_client = handler.data()
mqtt_client.set_uplink_callback(uplink_callback)
mqtt_client.connect()
time.sleep(86400) # duration time in seconds (86400 = 24 hrs)
mqtt_client.close() #close mqtt client
