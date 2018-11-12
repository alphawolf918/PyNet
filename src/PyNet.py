import speedtest

speedtester = speedtest.Speedtest()

speedtester.get_best_server()

megabits = 1048576

downloadSpeed = int(speedtester.download()) / megabits
uploadSpeed = int(speedtester.upload()) / megabits

downloadSpeed = round(downloadSpeed, 2)
uploadSpeed = round(uploadSpeed, 2)

print("-------")
print("Download Speed: " + str(downloadSpeed) + " mb/s")
print("Upload Speed: " + str(uploadSpeed) + " mb/s")
print("-------")