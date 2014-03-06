' Taken from the Stack Exchange Network. Based on http://stackoverflow.com/a/9541478 by Nathan Rice (https://stackoverflow.com/users/1159822/nathan-rice)
' Licensed under cc-by-sa 3.0 https://creativecommons.org/licenses/by-sa/3.0/
Set shell = WScript.CreateObject("WScript.Shell")
sCurrentDirectory = shell.CurrentDirectory & "\"

' Taken from the Stack Exchange Network. Based on http://serverfault.com/a/29720 by Maxwell (https://serverfault.com/users/8132/maxwell)
' Licensed under cc-by-sa 3.0 https://creativecommons.org/licenses/by-sa/3.0/
'Set your settings

Set args = Wscript.Arguments

strFileURL = args(0) & args(2)
strHDLocation = sCurrentDirectory & args(1) & args(2)

' Fetch the file

Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

objXMLHTTP.open "GET", strFileURL, false
objXMLHTTP.send()

If objXMLHTTP.Status = 200 Then
   Set objADOStream = CreateObject("ADODB.Stream")
   objADOStream.Open
   objADOStream.Type = 1 'adTypeBinary

   objADOStream.Write objXMLHTTP.ResponseBody
   objADOStream.Position = 0    'Set the stream position to the start

   Set objFSO = Createobject("Scripting.FileSystemObject")

   If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation

   Set objFSO = Nothing

   objADOStream.SaveToFile strHDLocation
   objADOStream.Close
   Set objADOStream = Nothing
End if

Set objXMLHTTP = Nothing