import speedtest
import pyodbc

speedtester = speedtest.Speedtest()

speedtester.get_best_server()

megabits = 1048576

downloadSpeed = speedtester.download() / megabits
uploadSpeed = speedtester.upload() / megabits

downloadSpeed = round(downloadSpeed, 2)
uploadSpeed = round(uploadSpeed, 2)

downloadSpeed = str(downloadSpeed)
uploadSpeed = str(uploadSpeed)

sqlServer = "localhost\SQLEXPRESS01"
sqlTable = "net_speeds"
sqlDb = "PyNet"

sqlCon = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=" + sqlServer + ";"
                      "Database=" + sqlDb + ";"
                      "Trusted_Connection=yes;")

cursor = sqlCon.cursor()
cursor.execute("INSERT INTO " + sqlTable + " (downloadSpeed, uploadSpeed) VALUES (" + downloadSpeed + ", " + uploadSpeed + ")")
cursor.commit()
cursor.close()

print("-------")
print("Download Speed: " + str(downloadSpeed) + " mb/s")
print("Upload Speed: " + str(uploadSpeed) + " mb/s")
print("-------")