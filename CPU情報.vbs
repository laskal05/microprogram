Option Explicit

'WMIにて使用する各種オブジェクトを定義・生成する。
Dim oClassSet
Dim oClass
Dim oLocator
Dim oService
Dim sMesStr

'ローカルコンピュータに接続する。
Set oLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
Set oService = oLocator.ConnectServer
'クエリー条件をWQLにて指定する。
Set oClassSet = oService.ExecQuery("Select * From Win32_Processor")

'コレクションを解析する。
For Each oClass In oClassSet

sMesStr = sMesStr & "種類：" & oClass.Description & vbCrLf & _
"名前：" & oClass.Name & vbCrLf & _
"製造元：" & oClass.Manufacturer & vbCrLf & _
"現在の周波数：" & CStr(oClass.CurrentClockSpeed) & vbCrLf & _
"最大周波数：" & CStr(oClass.MaxClockSpeed) & vbCrLf & _
"L2キャッシュサイズ：" & CStr(oClass.L2CacheSize) & vbCrLf & vbCrLf

Next

MsgBox("Processorに関する情報です。" & vbCrLf & vbCrLf & sMesStr)

'使用した各種オブジェクトを後片付けする。
Set oClassSet = Nothing
Set oClass = Nothing
Set oService = Nothing
Set oLocator = Nothing
