'URL="https://eve-central.com/home/tradefind_display.html?qtype=Regions&newsearch=1&fromt=10000002&to=10000002"
URL="https://eve-central.com/home/tradefind_display.html?qtype=SystemToRegion&newsearch=1&fromt=30000142&to=10000002"

htmlName="dashbo.html"

' https://eve-central.com/home/tradefind_display.html?qtype=Regions&newsearch=1&fromt=10000002&to=10000002
' https://eve-central.com/home/tradefind_display.html?qtype=Regions&fromt=10000002&to=10000002&age=24&minprofit=500000&size=8000&startat=50&sort=jprofit
' https://eve-central.com/home/tradefind_display.html?set=1&fromt=10000002&to=10000002&qtype=Regions&age=24&minprofit=500000&size=8000&limit=50&sort=jumps&prefer_sec=1
' https://eve-central.com/home/tradefind_display.html?set=1&fromt=10000002&to=10000002&qtype=Regions&age=24&minprofit=500000&size=8000&limit=100&sort=jumps&prefer_sec=1
'
'https://eve-central.com/home/quicklook.html?typeid=3687
 

set xmlhttp = createobject ("msxml2.xmlhttp.3.0")


Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile(htmlName,True) 

resFile.write ("<html><head><script src=sorttable.js></script></head><body bgcolor=#4d4d4d><table  border=1 class=sortable>" & vbCrLf)
resFile.write ("<tr><th>Login</th><th>Name</th><th>race</th><th>balance</th><th>location</th><th>training</th><th>Seka</th></tr>" & vbCrLf)


loginsStr="bluesteel241 paulmo122 anton.chertilo aryk38 5 6 iven.pipetka Mimoprohodil12"
apikeysStr="6014844 5715885 5651449 6015647 6017529 6017605 6021390 5982762"
verikodsStr="0awkJWlIm5vcTuU2cggga1Z9k0c96fX1CNGAy6rOKYR5W40FmBdb0gfUnv8Twnry rdEcYVofH2T7cexsgyjEBJ5KWR8tZRFi4hyCS4b23eDO4rVkfsIIh8LFrrOqW8Qt "&_
    "yLKPiZsCDO4wwy0nSB3696zXR0VhT92Ts5YTHmwu9k15gnXyKibFPr8vTVygihs9 Bucb3pLW8WqJdpBKdKkjTFAqHr2lMIdvoB3Wd4FPrujE0asZAwEJRX9RUN4r9QBB "&_
	"w2UXv2NcevsBNR0OuuN6oO5KPJll54rhPu4keen884YMmPIzaJCtpYBY0LpVPlHM "&_
	"XgurnDKIHiBGcebGckel3S9QGsQYJgQ1o5dnYQBfUvOLn6B7xtB9fKrLB6duUulR "&_
	"7bUyKvym6SNYV6CYdz5QZ1SI4ncW2mCghtiLf8ymTFUvtPgtQGdqCLH1U4KpisIf "&_
	"JoVCFvxRNxnUkIKi5I2l9T5I4qUsq0brfQhNXvXanHEcHLIf8RP1RKFzz6431Ial"
charidsStr="2112084011 96729321 96794174 96832975 96634566 96676530 96773060 2112365983"

Dim apikeys
apikeys = Split (apikeysStr)
Dim verikods
verikods = Split (verikodsStr)
Dim charids
charids = Split (charidsStr)
Dim logins
logins = Split (loginsStr)



'https://api.eveonline.com/eve/CharacterInfo.xml.aspx?keyID=5715885&vCode=rdEcYVofH2T7cexsgyjEBJ5KWR8tZRFi4hyCS4b23eDO4rVkfsIIh8LFrrOqW8Qt&characterID=96729321
'<characterID>96729321</characterID><characterName>Paulo Panala</characterName><race>Caldari</race>
'<accountBalance>6171506.07</accountBalance><skillPoints>2897005</skillPoints>
'<nextTrainingEnds>2017-02-28 02:04:02</nextTrainingEnds><shipName>Paulo Panala's Corax</shipName>
'<shipTypeID>32876</shipTypeID><shipTypeName>Corax</shipTypeName><corporationID>98487344</corporationID>
'<corporation>SOROKIN NORMA</corporation><corporationDate>2016-12-26 18:08:00</corporationDate>
'<lastKnownLocation>Jita IV - Moon 4 - Caldari Navy Assembly Plant</lastKnownLocation>


url1="https://api.eveonline.com/eve/CharacterInfo.xml.aspx?keyID="
url2="&vCode="
url3="&characterID="

for j = 0 to UBound(apikeys)
    URL=url1+apikeys(j)+url2+verikods(j)+url3+charids(j)

    xmlhttp.open "get", URL, false
    xmlhttp.send
    MyText=xmlhttp.responseText

    startpos=1
    NameSt=InStr(startpos, MyText,"<characterName>")+15
    NameEnd=InStr(NameSt, MyText,"</characterName>")
    NameStr=Mid(MyText, NameSt, NameEnd-NameSt)

    RaceSt=InStr(startpos, MyText,"<race>")+6
    RaceEnd=InStr(RaceSt, MyText,"</race>")
    RaceStr=Mid(MyText, RaceSt, RaceEnd-RaceSt)

    BalSt=InStr(startpos, MyText,"<accountBalance>")+16
    BalEnd=InStr(BalSt, MyText,"</accountBalance>")
    BalStr=Mid(MyText, BalSt, BalEnd-BalSt)
	
	SekaSt=InStr(startpos, MyText,"<securityStatus>")+16
    SekaEnd=InStr(SekaSt, MyText,"</securityStatus>")
    SekaStr=Mid(MyText, SekaSt, SekaEnd-SekaSt)

    CurSt=InStr(startpos, MyText,"<currentTime>")+13
    CurEnd=InStr(CurSt, MyText,"</currentTime>")
    CurStr=Mid(MyText, CurSt, CurEnd-CurSt)
	CurDate=CDate(CurStr)
	
    LocKnown = InStr(startpos, MyText,"<lastKnownLocation />")
    if LocKnown>0 then
       LocStr="X/3"
    else  
       LocSt=InStr(startpos, MyText,"<lastKnownLocation>")+19
       LocEnd=InStr(LocSt, MyText,"</lastKnownLocation>")
       LocStr=Mid(MyText, LocSt, LocEnd-LocSt)
    end if
	
    TraSt=InStr(startpos, MyText,"<nextTrainingEnds>")+18
	if TraSt>18 then
       TraEnd=InStr(TraSt, MyText,"</nextTrainingEnds>")
       TraStr= Mid(MyText, TraSt, TraEnd-TraSt)
	   TraDate=CDate(TraStr)
	   difa=DateDiff("h",CurDate,TraDate)
	   TraStr=CStr(difa)+" hours"
	else
       TraStr= " NETU!!! "
    end if	
	if difa<0 then TraStr= " NETU!!! "
   
    resFile.write ("<tr><td>"+logins(j)+"</td> <td>"&NameStr&"</td> <td>"&RaceStr&"</td> <td align='right'>"&FormatCurrency(BalStr)&_
	"</td> <td>"&LocStr&"</td><td>"&TraStr&"</td> <td>"+SekaStr+"</td><td><a href='"+URL+"'>link</a></td> </tr>" & vbCrLf )
   
next

resFile.write ("</table><br>")
resFile.write ("<a href=https://www.fuzzwork.co.uk/blueprint/?typeid=17672>NUKLEAR XL</a><br>")
resFile.write ("</body></html>")
resFile.Close


set shell = WScript.CreateObject("WScript.Shell")
shell.Run "cmd /c  start " + htmlName



'bluesteel241
'6014844
'0awkJWlIm5vcTuU2cggga1Z9k0c96fX1CNGAy6rOKYR5W40FmBdb0gfUnv8Twnry
'Julia Timoshenko24
'https://api.eveonline.com/eve/CharacterID.xml.aspx?names=Paulo%20Panala
'<row name="Julia Timoshenko24" characterID="2112084011"/>
'
'https://api.eveonline.com/eve/CharacterInfo.xml.aspx?keyID=6014844&vCode=0awkJWlIm5vcTuU2cggga1Z9k0c96fX1CNGAy6rOKYR5W40FmBdb0gfUnv8Twnry&characterID=2112084011