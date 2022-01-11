MsgBox "Welcome to Battery Monitor"& vbCrLf &"A simple battery monitor which keeps its eyes on whether your battery is fully charged or not"& vbCrLf &"~~~~~~~ CREATED BY DJ ~~~~~~~", "64", "BATTERY MONITOR"
set oLocator = CreateObject("WbemScripting.SWbemLocator")
set oServices = oLocator.ConnectServer(".","root\wmi")
set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")
for each oResult in oResults
   iFull = oResult.FullChargedCapacity
next

while (1)
  set oResults = oServices.ExecQuery("select * from batterystatus")
  for each oResult in oResults
    iRemaining = oResult.RemainingCapacity
    bCharging = oResult.Charging
  next
  iPercent = ((iRemaining / iFull) * 100) mod 100
  if bCharging and (iPercent > 98) Then MsgBox "Your battery has reached 98% !"& vbCrLf &"Please unplug your Laptop Charger...",48,"Battery Warning"
wend
