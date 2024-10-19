' VBScript to continuously ping 8.8.8.8 and display graphical results
' Author: apenczolgolonabijorbiper
' Date: 2024-10-19
'
remoteHost = "8.8.8.8"
Set objShell = CreateObject("WScript.Shell")
Set objPing = objShell.Exec("cmd /c ping " & remoteHost & " -t")
'Set objPing = objShell.Exec("cmd.exe /c start /b cmd /c ping " & remoteHost & " -t > nul")
Set objIE = CreateObject("InternetExplorer.Application")

' Set up the browser window
objIE.Visible = True
objIE.FullScreen = False
objIE.Visible = True
objIE.AddressBar = False
objIE.MenuBar = False
objIE.ToolBar = False
objIE.Width = 100
objIE.Height = 100
objIE.Left = 100
objIE.Top = 100
objIE.navigate("about:blank")

Function GetGreenToOrangeShade(value)
    Dim red, green, blue, colorShade
    
    If value < 100 Then
        ' For values below 100, assign a much darker green
        red = 0
        green = Int(100 + (value / 100) * 155)  ' Green ranges from 100 to 255
        blue = 0
    ElseIf value < 300 Then
        ' For values between 100 and 300, keep it green
        red = 0
        green = 255
        blue = 0
    Else
        ' From 300 to 600, blend from green to reddish-orange
        Dim ratio
        ratio = (value - 300) / 300  ' Normalize between 0 and 1 for the range 300 to 1000

        ' Faster transition to reddish-orange:
        red = Int(255 * ratio)              ' Red increases from 0 to 255
        green = Int(255 - (180 * ratio))    ' Green decreases more rapidly, from 255 to 75
        blue = 100                            ' Blue remains constant at 0
    End If

    ' Combine into the RGB string
    colorShade = "rgb(" & red & "," & green & "," & blue & ")"

    GetGreenToOrangeShade = colorShade
End Function



' Initialize the HTML document with basic structure
Do While objIE.Busy
    WScript.Sleep 100
Loop
Set doc = objIE.Document
doc.Write "<html><head><title>Ping Results</title></head><body>"
doc.write "<table border=0 style='width: 100%; height: 100%;'>"
doc.Write "<tr><td colspan=5 id='pingOutput' style='color: black; font-size: 10; text-align: center'></td></tr>"
doc.Write "<tr style='height: 80%;'>"
doc.Write "<td id='status1' style='width: 20%; height: 100%; color: white; font-size: 14; text-align: center'></td>"
doc.Write "<td id='status2' style='width: 20%; height: 100%; color: white; font-size: 14; text-align: center'></td>"
doc.Write "<td id='status3' style='width: 20%; height: 100%; color: white; font-size: 14; text-align: center'></td>"
doc.Write "<td id='status4' style='width: 20%; height: 100%; color: white; font-size: 14; text-align: center'></td>"
doc.Write "<td id='status5' style='width: 20%; height: 100%; color: white; font-size: 14; text-align: center'></td>"
doc.Write "</tr></table>"
'doc.Write "<pre id='pingOutput'></pre>"
doc.Write "</body></html>"


' Variables to track ping success
Dim successCount
successCount = 0
successTime=""
Const MAX_SUCCESS_COUNT = 5

WScript.Sleep 2000  ' Wait for IE to load the window
objShell.Run "SetAlwaysOnTop.ahk", 0, False  ' Change path to your AutoHotkey script


Do While objIE.Visible
    ' Capture output from ping command
    If not isNull(objPing) and Not objPing.StdOut.AtEndOfStream Then
	
        strPingResult = objPing.StdOut.ReadLine()
        
	On Error Resume Next
        ' Update the ping results in the HTML window
'        doc.getElementById("pingOutput").innerHTML = doc.getElementById("pingOutput").innerHTML & strPingResult & vbCrLf
        
'MsgBox(doc.getElementById("status2").style.backgroundColor)
        doc.getElementById("status5").style.backgroundColor = doc.getElementById("status4").style.backgroundColor
        doc.getElementById("status5").innerHTML = doc.getElementById("status4").innerHTML
        doc.getElementById("status4").style.backgroundColor = doc.getElementById("status3").style.backgroundColor
        doc.getElementById("status4").innerHTML = doc.getElementById("status3").innerHTML
        doc.getElementById("status3").style.backgroundColor = doc.getElementById("status2").style.backgroundColor
        doc.getElementById("status3").innerHTML = doc.getElementById("status2").innerHTML
        doc.getElementById("status2").style.backgroundColor = doc.getElementById("status1").style.backgroundColor
        doc.getElementById("status2").innerHTML = doc.getElementById("status1").innerHTML

        ' Check if the result contains "Reply from" indicating a successful ping
        If InStr(strPingResult, "Reply from") > 0 and InStr(strPingResult, "unreachable") = 0 Then
            successCount = successCount + 1
	    successTime =  Right(strPingResult, len(strPingResult)-inStr(strPingResult,"time=")-len("time=")+1)
'            doc.getElementById("pingOutput").innerHTML = "+++" & successTime & "+++"
	    successTime =  Left(successTime, inStr(successTime, "ms TTL")-1)
	    doc.getElementById("status1").innerHTML = successTime
        Else
            successCount = 0  ' Reset count if ping fails
	    doc.getElementById("status1").innerHTML = "x"
	    doc.getElementById("pingOutput").innerHTML = strPingResult
        End If
        
        ' Update status rectangle color based on success count
        If successCount >= MAX_SUCCESS_COUNT Then
'            doc.getElementById("status1").style.backgroundColor = "green"
		doc.getElementById("status1").style.backgroundColor = GetGreenToOrangeShade(successTime)
            doc.getElementById("pingOutput").innerHtml = "Connection to " & remoteHost & " is stable."
        Else
            doc.getElementById("status1").style.backgroundColor = "red"
'            doc.getElementById("status1").innerHtml = "x"
        End If
        
        ' Scroll the window to show the latest ping result
'        doc.ParentWindow.scrollTo 0, doc.body.scrollHeight
    Else
	objPing = null
	doc.getElementById("pingOutput").innerHTML = "Ping crashed - need a restart."
    End If
    
    ' Pause for a short time to prevent freezing
    WScript.Sleep 100
Loop
objIE.Quit
Set objIE = Nothing
Set objPing = Nothing