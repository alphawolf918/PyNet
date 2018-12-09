import speedtest
import pyodbc
import ctypes
import win32serviceutil as winServ
import subprocess as sp

ctypes.windll.kernel32.SetConsoleTitleW("PyNet")

sqlServiceStatus = winServ.QueryServiceStatus("MSSQL$SQLEXPRESS01")[1]

if sqlServiceStatus != 4:
    print("Starting SQL Server Service...")
    winServ.StartService(sqlServiceName)
    print("SQL Server Service started successfully.\n")

def getWiFiName():
    cmd = sp.check_output(["netsh", "wlan", "show", "interfaces"], shell=True)
    cmd = str.replace(str(cmd), "  ", "")
    cmd = str.replace(cmd, "\\r\\n", "$")
    netInfo = str.split(cmd, "$")
    wifi = "NULL"
    for s in netInfo:
        x = str.split(str.strip(s), ":")
        if len(x) > 1:
            key, value = x[0], x[1]
            if str.rstrip(key, " ") == "SSID":
                wifi = str.lstrip(value, " ")
                break
    return wifi

print("Connecting to SpeedTest...")
speedtester = speedtest.Speedtest()
    
print("Getting set of closest servers...")
speedtester.get_closest_servers()

print("Retrieving best possible server...")
speedtester.get_best_server()

megabits = 1048576

print("Calculating download speed...")
downloadSpeed = speedtester.download() / megabits

print("Calculating upload speed...")
uploadSpeed = speedtester.upload() / megabits

downloadSpeed = str(round(downloadSpeed, 2))
uploadSpeed = str(round(uploadSpeed, 2))
wifiNet = getWiFiName()

print("Connecting to database...")
try:
    sqlCon = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                            "Server=localhost\SQLEXPRESS01;"
                            "Database=PyNet;"
                            "Trusted_Connection=yes;")

    print("Connected! Inserting network speed values...")
    cursor = sqlCon.cursor()
    cursor.execute("INSERT INTO net_speeds (downloadSpeed, uploadSpeed, network_name) "
                   "VALUES (" + downloadSpeed + ", " + uploadSpeed + ", '" + wifiNet +"')")
    cursor.commit()
    print("Values inserted, closing connection.")
    cursor.close()
    sqlCon.close()
    print("Connection closed.\n")
except:
    print("There was an error when connecting to the database.")

print("--------------")
print("Download Speed: " + downloadSpeed + " Mbps")
print("Upload Speed: " + uploadSpeed + " Mbps")
print("Wi-Fi: " + wifiNet)
print("--------------\n")