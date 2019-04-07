import speedtest as spd
import pyodbc as db
import ctypes
import win32serviceutil as winServ
import subprocess as sp
import configparser
import io

SW_HIDE = 0
info = sp.STARTUPINFO()
info.dwFlags = sp.STARTF_USESHOWWINDOW
info.wShowWindow = SW_HIDE
SERVICE_ACTIVE = 4

ctypes.windll.kernel32.SetConsoleTitleW("PyNet")

config = configparser.ConfigParser()
config.readfp(open(r"PyConfig.ini"))
confSection = "DEFAULT"

sqlServiceName = config.get(confSection, "SQLServiceName")
sqlDriver = config.get(confSection, "Driver")
sqlServ = config.get(confSection, "Server")
sqlDb = config.get(confSection, "Database")

sqlServiceStatus = winServ.QueryServiceStatus(sqlServiceName)[1]

if sqlServiceStatus != SERVICE_ACTIVE:
    try:
        print("Starting SQL Server Service...")
        winServ.StartService(sqlServiceName)
        print("SQL Server Service started successfully.\n")
    except:
        print("There was an error with starting the SQL Server Service.")

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
speedtester = spd.Speedtest()
    
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
    sqlCon = db.connect("Driver=" + sqlDriver + ";"
                        "Server=" + sqlServ + ";"
                        "Database=" + sqlDb + ";"
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