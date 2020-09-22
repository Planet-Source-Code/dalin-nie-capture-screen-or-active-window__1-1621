<div align="center">

## Capture Screen or Active Window


</div>

### Description

This function capture the screen or the active window of your computer

Programmatically and save it to a .bmp file. This may allows you to get another machine's

screen through network!!! Fully tested in VB5.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dalin Nie](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dalin-nie.md)
**Level**          |Unknown
**User Rating**    |3.6 (18 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dalin-nie-capture-screen-or-active-window__1-1621/archive/master.zip)





### Source Code

```
'1: Declare
' This should be in the form's heneral declaration area.
' If you do it in a module, omit the word "Private"
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
 ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'
'2. The Function
' You can add this to your form's code
' or you can put it in a module if the declaration is in a module
Public Function fSaveGuiToFile(ByVal theFile As String) As Boolean
' Name: fSaveGuiToFile
' Author: Dalin Nie
' Written: 4/2/99
' Purpose:
' This procedure will Capture the Screen or the active window of your Computer and Save it as
' a .bmp file
' Input:
' theFile file Name with path, where you want the .bmp to be saved
'
' Output:
' True if successful
'
Dim lString As String
On Error goto Trap
'Check if the File Exist
 If Dir(theFile) <> "" Then Exit Function
 'To get the Entire Screen
 Call keybd_event(vbKeySnapshot, 1, 0, 0)
 'To get the Active Window
 'Call keybd_event(vbKeySnapshot, 0, 0, 0)
 SavePicture Clipboard.GetData(vbCFBitmap), theFile
fSaveGuiToFile = True
Exit Function
Trap:
'Error handling
MsgBox "Error Occured in fSaveGuiToFile. Error #: " & Err.Number & ", " & Err.Description
End Function
'
3. To call the function, add the code:
Call fSaveGuiToFile(yourFileNAme)
' Example: in a command1_click event add: call fSaveGuiToFile("C:\Scrn_pic.bmp")
'When you run your app, click command1, the screen will be saved in c:\scrn_pic.bmp.
```

