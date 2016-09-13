set objFSO = CreateObject("scripting.FileSystemObject")

strInfile = WScript.Arguments(0)
strOutFile = strInFile & "_tableau.csv"

Set objInfile 	= objFSO.OpenTextFile(strInfile,1)
set objOutfile 	= objFSO.CreateTextFile(strOutfile)

objOutfile.WriteLine "Date/Time,MetricFullName,Measure"

count = 0 
Do While not objInfile.AtEndOfstream
	strReadline = replace(objInfile.ReadLine, chr(34), "")
	arrReadline = Split(strReadLine, ",")
    
   If count = 0 Then
		arrHeaders = arrReadLine
   Else
	' American Date
   	strDateTime = left(arrReadLine(0),19)
	' European Date
'	strDateTime = mid(strDatetime,4,2) & "/" & left(strDatetime,2) & mid(strDateTime, 6)

	For a = 1 to Ubound(arrReadLine)
		If arrReadline(a) = " " Then
			objOutFile.Writeline Chr(34) & strDateTime & Chr(34) & "," & Chr(34) & arrHeaders(a) & Chr(34) & "," &  ""
		Else
			objOutFile.Writeline Chr(34) & strDateTime & Chr(34) & "," & Chr(34) & arrHeaders(a) & Chr(34) & "," &  arrReadLine(a) 
		End If 
	Next

   End If
   count = count +1
Loop

objInfile.close
objOutFile.Close