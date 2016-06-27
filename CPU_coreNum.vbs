strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colCompSys = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objCS in colCompSys
 str = " 物理コア数: " & objCS.NumberOfCores & vbCrLf 
Next
Set colCompSys2 = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objCS in colCompSys2
 str = str & " ソケット数: " & objCS.NumberOfProcessors & vbCrLf 
 str = str & " 論理コア数: " & objCS.NumberOfLogicalProcessors & vbCtLf
Next
WScript.Echo str
