' VBScript to continuously ping 8.8.8.8 and display graphical results
' Author: apenczolgolonabijorbiper
' Date: 2024-10-19
'
remoteHost = "8.8.8.8"
Set objShell = CreateObject("WScript.Shell")
Set objPing = objShell.Exec("cmd /c ping " & remoteHost & " -t")
'Set objPing = objShell.Exec("cmd /c mode con: cols=10 lines=10 && ping " & remoteHost & " -t")
'Set objPing = objShell.Exec("cmd.exe /c start /b cmd /c ping " & remoteHost & " -t > nul")
Set objIE = CreateObject("InternetExplorer.Application")

' Set up the browser window
objIE.Visible = True
objIE.FullScreen = False
objIE.Visible = True
objIE.AddressBar = False
objIE.MenuBar = False
objIE.ToolBar = False
objIE.StatusBar = False
objIE.Resizable = False
objIE.Silent = True
'objIE.TheaterMode = False

objIE.Width = 1 'minium IE width 250
objIE.Height = 1 'minium IE height 100

objIE.navigate("about:blank")
objIE.document.focus()

for each vid in getobject("winmgmts:").instancesof("Win32_VideoController")
	intHorizontal = vid.CurrentHorizontalResolution 
	intVertical = vid.currentVerticalResolution
next

objIE.Left = intHorizontal-objIE.Width
objIE.Top = intVertical-objIE.Height-80

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
doc.Write "<html><head><title>&#8584; Pingoji &#9637;</title></head>"
doc.write "<body>"
doc.write "<table border=0 style='width: 100%; height: 100%;' cellspacing=1>"
doc.Write "<tr><td colspan=6 id='pingOutput' style='color: black; font-size: 11; text-align: center'></td></tr>"
doc.Write "<tr style='height: 1%;'>"
doc.Write "<td id='status' style='width: 1%; height: 90%; color: white; font-size: 9; text-align: center'></td>"
doc.Write "<td id='status1' style='width: 10%; height: 90%; color: white; font-size: 9; text-align: center'></td>"
doc.Write "<td id='status2' style='width: 10%; height: 90%; color: white; font-size: 9; text-align: center'></td>"
doc.Write "<td id='status3' style='width: 10%; height: 90%; color: white; font-size: 9; text-align: center'></td>"
doc.Write "<td id='status4' style='width: 10%; height: 90%; color: white; font-size: 9; text-align: center'></td>"
doc.Write "<td id='status5' style='width: 10%; height: 90%; color: white; font-size: 9; text-align: center'></td>"
doc.Write "</tr>"
doc.Write "<tr><td colspan=6 id='info' style='width: 1%; color: black; font-size: 5; text-align: left'>&#8658;</td></tr>"
doc.Write "<tr><td colspan=6 id='info2' style='width: 1%; color: black; font-size: 3; text-align: left'>&#8658;</td></tr>"
doc.write "</table>"
doc.Write "</body></html>"

' Variables to track ping success
successCount = 0
successTime = ""
Const MAX_SUCCESS_COUNT = 3 ' when connection becomes stable
startTime = Time()
thisTime = Time()
barCount = 0
bar2Count = 0
bar2Sum = 0

WScript.Sleep 1000  ' Wait for IE to load the window
'objShell.Run "SetAlwaysOnTop.ahk", 0, False  ' Change path to your AutoHotkey script

Do While objIE.Visible
    ' Capture output from ping command
    If not isNull(objPing) and Not objPing.StdOut.AtEndOfStream Then
	
	On Error Resume Next
        ' Update the ping results in the HTML window
        
        if doc.getElementById("status").style.backgroundColor = "blue" then
		doc.getElementById("status").style.backgroundColor = "white"
	Else
		doc.getElementById("status").style.backgroundColor = "blue" 
	End if

        doc.getElementById("status5").style.backgroundColor = doc.getElementById("status4").style.backgroundColor
        doc.getElementById("status5").innerHTML = doc.getElementById("status4").innerHTML
        doc.getElementById("status4").style.backgroundColor = doc.getElementById("status3").style.backgroundColor
        doc.getElementById("status4").innerHTML = doc.getElementById("status3").innerHTML
        doc.getElementById("status3").style.backgroundColor = doc.getElementById("status2").style.backgroundColor
        doc.getElementById("status3").innerHTML = doc.getElementById("status2").innerHTML
        doc.getElementById("status2").style.backgroundColor = doc.getElementById("status1").style.backgroundColor
        doc.getElementById("status2").innerHTML = doc.getElementById("status1").innerHTML

        strPingResult = objPing.StdOut.ReadLine()
	thisTime = Time()
	barCount=barCount+1
	bar2Count=bar2Count+1

        if barCount>(objIE.Width/4.1) Then
            doc.getElementById("info").innerHTML = Left(doc.getElementById("info").innerHTML,InStrRev(doc.getElementById("info").innerHTML, "<font")-1)		
	end if

	if bar2Count mod 60 = 0 Then
	    if bar2Count > (objIE.Width/3.3)*80 then
              doc.getElementById("info2").innerHTML = Left(doc.getElementById("info2").innerHTML,InStrRev(doc.getElementById("info2").innerHTML, "<font")-1)		
	    end if
            doc.getElementById("info2").innerHTML = "<font color='" & GetGreenToOrangeShade(bar2Sum/60) & "'>&#9608;</font>" & doc.getElementById("info2").innerHTML
	    bar2Sum=0
	end if

        ' Check if the result contains remoteHost and TTL indicating a successful ping
        If InStr(strPingResult, remoteHost) > 0 and InStr(strPingResult, "TTL") > 0 Then
            successCount = successCount + 1
	    successTime =  Right(strPingResult, len(strPingResult)-inStr(strPingResult,"time=")-len("time=")+1)
	    successTime =  Left(successTime, inStr(successTime, "ms TTL")-1)
	    doc.getElementById("status1").innerHTML = successTime
	    bar2Sum = bar2Sum + successTime
        Else
            successCount = 0  ' Reset count if ping fails
	    doc.getElementById("status1").innerHTML = "x"
	    doc.getElementById("pingOutput").innerHTML = strPingResult
            doc.getElementById("status1").style.backgroundColor = "red"
            doc.getElementById("info").innerHTML = "<font color=red>&#9608;</font>" & doc.getElementById("info").innerHTML
	    startTime = Time()
	    bar2Sum = bar2Sum + 1000
        End If
        
        ' Update status rectangle color based on success count
        If successCount >= MAX_SUCCESS_COUNT Then
	    greenColor = GetGreenToOrangeShade(successTime)
            doc.getElementById("pingOutput").innerHtml = "Connection to " & remoteHost & " is stable (" & DateDiff("s", startTime, thisTime) & "s)."
	    doc.getElementById("status1").style.backgroundColor = greenColor
            doc.getElementById("info").innerHTML = "<font color='" & greenColor & "'>&#9608;</font>" & doc.getElementById("info").innerHTML 
        Else
	    if successCount >= 0 then
	        doc.getElementById("pingOutput").innerHtml = "Connection to " & remoteHost & " is unstable."
		if doc.getElementById("status1").innerHTML <> "x" then
			amberColor = "#FFBF00"
	          doc.getElementById("status1").style.backgroundColor = amberColor
                  doc.getElementById("info").innerHTML = "<font color='" & amberColor & "'>&#9608;</font>" & doc.getElementById("info").innerHTML
		end if
	    end if
        End If

    Else
	objPing = null
	doc.getElementById("pingOutput").innerHTML = "Ping crashed - need a restart."
    End If
    
    ' Pause for a short time to prevent freezing
    WScript.Sleep 10

If Err.Number <> 0 Then
'  WScript.Echo "Error in SomeCodeHere: " & Err.Number & ", " & Err.Source & ", " & Err.Description
  Err.Clear

MsgBox("Thanks for using Pingoji")
objIE.Terminate()
Set objIE = Nothing
Set objPing = Nothing
Set objShell = Nothing

  exit do
End If

Loop
