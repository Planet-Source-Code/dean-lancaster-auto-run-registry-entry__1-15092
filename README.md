<div align="center">

## Auto Run Registry Entry


</div>

### Description

If you have ever needed to have your program start at the same time as windows starts without putting the program into the start up menu but have been unable to figure out how to do it then just follow the following text.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dean lancaster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dean-lancaster.md)
**Level**          |Beginner
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dean-lancaster-auto-run-registry-entry__1-15092/archive/master.zip)





### Source Code

What you have to do to use this code is:
In your project (for example, form_load())
insert the following text somewhere that would be useful to your project: <br><br>
<font color="red">
<pre>
Call SetStringValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "Currency", App.Path + "\" + App.EXEName + ".exe")<br><br>
</font></pre>
This code will put the current project path and the current project name in the windows auto run section in the windows registry.<br>
The benfit of this is so that you can have your program start as windows starts.<br>
Add to your project a module <br><br>
project -> add module<br><br>
and insert the following code<br><br>
<font color="red">
<pre>
Public Exist As Boolean<br><br>
Public Const HKEY_CLASSES_ROOT = &H80000000<br>
Public Const HKEY_CURRENT_USER = &H80000001<br>
Public Const HKEY_LOCAL_MACHINE = &H80000002<br>
Public Const HKEY_USERS = &H80000003<br>
Public Const HKEY_PERFORMANCE_DATA = &H80000004<br>
Public Const ERROR_SUCCESS = 0&<br>
Public Const REG_SZ = 1<br><br>
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long<br><br>
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long<br><br>
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long<br><br>
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long<br><br>
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long<br><br>
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long<br><br>
</font></pre>
After this has been added also add the following code to your module file<br><br>
<font color="red">
<pre>Public Sub SetStringValue(Hkey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim i As Long
 'Create the key
 i = RegCreateKey(Hkey, strPath, keyhand)
 'Set the value
 i = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
 'Close the key
 i = RegCloseKey(keyhand)
End Sub</pre>
</font><br><br>
This then should allow the setstringvalue to be accessed and the information you need be put into the windows auto run registry.<br>
I must thank <a href="http://www.okdeluxe.com"> T. L. Phillips </a> for the code for the module form, that is where i optained it.<br>
Hope you like the code! :)<br>
if you need help with it (although you shouldn't) mail me :) <br> (As for the Link to mr phillips's website, that is beyond my control and the link that is on his source submissions)

