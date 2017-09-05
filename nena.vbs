URL="https://zkillboard.com/system/30045344/"

set xmlhttp = createobject ("msxml2.xmlhttp.3.0")

Function CheckDrop(Arg1)
	xmlhttp.open "get", Arg1, false
	xmlhttp.send
	dropText= xmlhttp.responseText
    droposSt = InStr(1, dropText,"Total Dropped:")+199
	droposEnd = InStr(droposSt, dropText,"h5>")-2
	dropa = Mid(dropText, droposSt, droposEnd-droposSt)
	'Msgbox dropa
    'CheckDrop = "need to check " & Arg1
	CheckDrop =  Replace(dropa,",","")
End Function


Function LocDrop(Arg1)
	xmlhttp.open "get", Arg1, false
	xmlhttp.send
	dropText= xmlhttp.responseText
    droposSt = InStr(1, dropText,"Location:")+49
	droposEnd = InStr(droposSt, dropText,"td>")-2
	dropa = Mid(dropText, droposSt, droposEnd-droposSt)
	LocDrop =  Replace(dropa,"</a>","   ")
	
End Function

htmlName="nena_fw_kill.html"

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile(htmlName,True)  


resFile.write ("<html><head><script src=sorttable.js></script></head><body bgcolor=#4d4d4d><table  border=1 class=sortable>" & vbCrLf)



'''''''   crazy shit...
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")    
dateTime.SetVarDate (now())
'MsgBox  "Local Time:  " & dateTime
'MsgBox  "UTC Time: " & dateTime.GetVarDate (false)
utcTime= dateTime.GetVarDate (false)
'MsgBox  "UTC Time: " & FormatDateTime(utcTime,4)
'''

'resFile.write (FormatDateTime(utcTime,4) & ",ship,bablo,locdrop" & vbCrLf)
resFile.write ("<tr> <th>"&FormatDateTime(utcTime,4)&"</th> <th>ship</th> <th>bablo</th> <th>location of drop</th>  </tr> " & vbCrLf)


xmlhttp.open "get", URL, false
xmlhttp.send
MyText= xmlhttp.responseText

startpos=1
for i =1 to  25 '50
    '  winwin killListRow
    curpos=InStr(startpos, MyText,"killListRow winwin")
	
	cur1st=InStr(curpos, MyText,"window.location=")
	kilaStr= Mid(MyText, cur1st+17,15)
	kilaNum= Mid(MyText, cur1st+23,8)
	kilaLink= "https://zkillboard.com"+Mid(MyText, cur1st+17,15)
	'MsgBox kilaLink
	'MsgBox kilaNum
	'MsgBox kilaStr
	
	timekill =Mid(MyText, cur1st+62,5)        

	cur2st=InStr(cur1st,MyText, "<a href="+chr(34)+kilaStr+chr(34))
	cur2end=InStr(cur2st+6, MyText, "<")
	utrata=Mid(MyText, cur2st+26, cur2end-cur2st-26)
	
	cur2ast=InStr(cur2st,MyText, "class="+chr(34)+"eveimage img-rounded"+chr(34)+" alt="  )+1
	cur2aend=InStr(cur2ast+6, MyText, "/")
	karabl=Mid(MyText, cur2ast+33, cur2aend-cur2ast-34)
	'MsgBox  karabl

	
	lo1111cdrop="X3"
	lo1111cdrop=LocDrop(kilaLink)
	
	resFile.write ("<tr> <td>"&timekill &"</td> <td>"&karabl &"</td> <td>"&utrata &"</td> <td>"&lo1111cdrop &"</td> </tr> " & vbCrLf)
	'resFile.write (timekill & "," &  karabl & "," & utrata & "," & lo1111cdrop & vbCrLf)
	
	'Msgbox "Kill:"+kilaNum+"   Karabl:"+karabl+"  Uron:"+utrata+"   system:"+system+"  region:"+region
	
	startpos=curpos+1
	
next

resFile.write ("</table></body></html>")
resFile.Close


set shell = WScript.CreateObject("WScript.Shell")
shell.Run "cmd /c  start " + htmlName