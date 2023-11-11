'Bing Images as Wallpaper.  The following code can be used to download the background images from Bing.com and setting them as your desktop wallpaper.  Just copy this code and paste into a text file which you save with the .vbs file etension.  Place the file in your startup folder to autorun the script.  It checks for new Bing background images every 8 minutes and will cause either WScript.exe or CScript.exe to appear in task manager when it is running.

dim html,startpos,endpos,length,startposfn,endposfn,lengthfn

Call Main()

Sub Wait()
'Edit this value to change how often the script checks Bing for a new background image, in milliseconds, set for around 8 minutes here.
WScript.Sleep 500000
Call Main()
End Sub

Sub Main

'Querying www.bing.com for html source code

Set WshShell = WScript.CreateObject("WScript.Shell")
Set http = CreateObject("Microsoft.XmlHttp")
http.open "GET", "http://www.bing.com" , FALSE
http.send ""

'WScript.Echo http.responseText 'Unrem to debug Bing HTML source code returned

'Parsing HTML source as a string
html = http.responseText

'Now to extract the background image details from the HTML source, if Microsoft recode the page, this is where to look.

'The start string we use to find the image URL in the HTML source code
startpos=InStr(html,"var g_img={url:'")

'The end string we use to find the image URL in the html source code
endpos=InStr(html,"',id:'bgDiv',d:200,cN:'_SS'")
length=endpos-(startpos+16)

'WScript.Echo (Mid(html,startpos+16,length)) 'Unrem to debug the image url appended to www.bing.com
imagefile =(Mid(html,startpos+16,length))

'Now we have the image url, some more manipulation to give us a nice filename, the same as we did with the image URL, this could all be done in one step as the path below Bing.com to the image does not appear to change, but oh well.

startposfn=InStr(imagefile,"hpk2\/")
endposfn=InStr(imagefile,"_")
lengthfn=endposfn-(startposfn+6)
'WScript.Echo (Mid(imagefile,startposfn+6,lengthfn)) 'Unrem to debug the image file name
imagefilename =(Mid(imagefile,startposfn+6,lengthfn))


'Now we have the image URL and a nice filename, time to download the image to our local machine

strFileURL = "http://www.bing.com" + imagefile

'Searches for the folder to save the images in and creates the BingWallpaper folder in the root of our profile if required.

Set objWshShell = WScript.CreateObject("WScript.Shell")
path = objWshShell.Environment("PROCESS")("UserProfile") & "\BingWallpaper"
set filesys=CreateObject("Scripting.FileSystemObject")
If Not filesys.FolderExists(path) Then
Set folder = filesys.CreateFolder(path)
End If

'Setting the location and file name to save the image as
strHDLocation = objWshShell.Environment("PROCESS")("UserProfile") & "\BingWallpaper\" & imagefilename & ".jpg"

'Getting the image file and saving to the local machine
Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
objXMLHTTP.open "GET", strFileURL, false
objXMLHTTP.send()
 
If objXMLHTTP.Status = 200 Then
Set objADOStream = CreateObject("ADODB.Stream")
objADOStream.Open
objADOStream.Type = 1 'adTypeBinary
objADOStream.Write objXMLHTTP.ResponseBody
objADOStream.Position = 0    
Set objFSO = Createobject("Scripting.FileSystemObject")
If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
Set objFSO = Nothing
objADOStream.SaveToFile strHDLocation
objADOStream.Close
Set objADOStream = Nothing
End if

'Setting the downloaded image as our desktop background, by writing the relevent registry keys.
Set wshShell = WScript.CreateObject("WScript.Shell")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
sWallPaper = objWshShell.Environment("PROCESS")("UserProfile") & "\BingWallpaper\" & imagefilename & ".jpg"
oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper

'This pause is here because the next bit is very ropey
WScript.Sleep 2000

'Refresh the desktop background, works well on XP, but is unsupported on Vista/7, does kind of work occasionally with them.
Set oShell = CreateObject("WScript.Shell")
oShell.Run _
"%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", _
1, True


'Close everything we used and go back into the waiting state.
set oShell = nothing 
set objXMLHTTP = nothing
set WshShell = nothing
set http = nothing
Call Wait()
End Sub
