import speedtest
import pyodbc
import ctypes
import win32serviceutil as winServ

ctypes.windll.kernel32.SetConsoleTitleW("PyNet")

sqlServiceName = "MSSQL$SQLEXPRESS01"
sqlServiceStatus = winServ.QueryServiceStatus(sqlServiceName)[1]

if sqlServiceStatus != 4:
    print("Starting SQL Server Service...")
    winServ.StartService(sqlServiceName)
    print("SQL Server Service started successfully.\n")

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

print("Connecting to database...")
try:
    sqlCon = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                            "Server=localhost\SQLEXPRESS01;"
                            "Database=PyNet;"
                            "Trusted_Connection=yes;")

    print("Connected! Inserting network speed values...")
    cursor = sqlCon.cursor()
    cursor.execute("INSERT INTO net_speeds (downloadSpeed, uploadSpeed) VALUES (" + downloadSpeed + ", " + uploadSpeed + ")")
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
print("--------------\n")