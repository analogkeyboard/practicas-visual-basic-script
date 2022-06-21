sumA = 30, sumB = 15 
wscript.echo "sumA: " & TypeName(sumA)
wscript.echo "sumB: " & TypeName(sumB)
Test = CBool(sumA = sumB) 
sumB= sumB * 2 
Test = CBool(sumA = sumB) 
wscript.echo "Test: " & TypeName(Test)
dateStr = "December 10, 2005" 
wscript.echo "dateStr: " & TypeName(dateStr)
theDate = CDate(dateStr) 
wscript.echo "theDate:"  & TypeName(theDate)
timeStr = "8:25:10 AM" 
theTime = CDate(timeStr) 
wscript.echo "timeStr:"  & TypeName(timeStr)
wscript.echo "theTime: " & TypeName(theTime)
aDouble = 715.255 
aString = CStr(aDouble) 
wscript.echo "aDouble: " & TypeName(aDouble)
wscript.echo "aString: " & TypeName(aString)