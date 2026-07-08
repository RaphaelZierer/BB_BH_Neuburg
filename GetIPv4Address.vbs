strComputer = "."
ON ERROR RESUME NEXT
SET objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
SET colSettings = objWMIService.ExecQuery ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True'")

FOR EACH objIP IN colSettings
 FOR i=LBound(objIP.IPAddress) TO UBound(objIP.IPAddress)
 If InStr(objIP.IPAddress(i),":") = 0 THEN ECHO objIP.IPAddress(i)
 NEXT
NEXT