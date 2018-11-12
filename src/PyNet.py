import speedtest
import pyodbc

print("Connecting to SpeedTest...")
speedtester = speedtest.Speedtest()

print("Retrieving best possible server...")
speedtester.get_best_server()

megabits = 1048576

print("Calculating download speed...")
downloadSpeed = speedtester.download() / megabits

print("Calculating upload speed...")
uploadSpeed = speedtester.upload() / megabits

downloadSpeed = round(downloadSpeed, 2)
uploadSpeed = round(uploadSpeed, 2)

downloadSpeed = str(downloadSpeed)
uploadSpeed = str(uploadSpeed)

sqlServer = "localhost\SQLEXPRESS01"
sqlTable = "net_speeds"
sqlDb = "PyNet"

print("Connecting to database...")

sqlCon = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=" + sqlServer + ";"
                      "Database=" + sqlDb + ";"
                      "Trusted_Connection=yes;")
print("Connected! Inserting network speed values...")
cursor = sqlCon.cursor()
cursor.execute("INSERT INTO " + sqlTable + " (downloadSpeed, uploadSpeed) VALUES (" + downloadSpeed + ", " + uploadSpeed + ")")
cursor.commit()
print("Values inserted, closing connection.")
cursor.close()
print("Connection closed.\n")

print("-------")
print("Download Speed: " + str(downloadSpeed) + " mb/s")
print("Upload Speed: " + str(uploadSpeed) + " mb/s")
print("-------\n")