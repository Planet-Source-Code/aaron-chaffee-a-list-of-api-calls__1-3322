<div align="center">

## A list of Api calls


</div>

### Description

A list a usefull api calls that i thought people might be able to use.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aaron Chaffee](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aaron-chaffee.md)
**Level**          |Unknown
**User Rating**    |4.0 (32 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aaron-chaffee-a-list-of-api-calls__1-3322/archive/master.zip)





### Source Code

```
How do I change the Double click time of the mouse?
The double click time is the time between two consecutive mouse clicks that will cause a double click event. You can change the time from your VB Application by calling the SetDoubleClickTime API function. It has only one parameter. This is the new DoubleClick time delay in milliseconds.
Declare Function SetDoubleClickTime Lib "user32" Alias _
"SetDoubleClickTime" (ByVal wCount As Long) As Long
N.B. These changes affect the entire system.
----------------------------------------------------------------------
How can I hide the cursor?
You can use the API function Showcursor, that allows you to control the visibility of the cursor. The declaration for this function is:
Declare Function ShowCursor& Lib "user32" _
(ByVal bShow As Long)
The Parameter bShow is set to True (non-zero) to display the cursor, False to hide it.
----------------------------------------------------------------------
How do I swap the mouse buttons?
Use the API Function SwapMouseButton to swap the functions of the Left and Right mouse buttons. The declare for this function is:
Declare Function SwapMouseButton& Lib "user32" _
(ByVal bSwap as long)
To swap the mouse buttons, call this function with the variable bSwap = True. Set bSwap to False to restore normal operation.
----------------------------------------------------------------------
How can I move the mouse cursor?
You can use the SetCursorPos Api function. It accepts two parameters. These are the x position and the y position in screen pixel coordinates. You can get the size of the screen by calling GetSystemMetrics function with the correct constants. This example puts the mouse cursor in the top left hand corner.
t& = SetCursorPos(0,0)
This will only work if the formula has bee declared in the declarations section:
Declare Function SetCursorPosition& Lib "user32" _
(ByVal x as long, ByVal y as long)
----------------------------------------------------------------------
How do I find out how much disk space is occupied?
Use the function GetDiskFreeSpace. The declaration for this API function is:
Declare Function GetDiskFreeSpace Lib "kernel32" Alias _
"GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters _
As Long) As Long
Here is an example of how to find out how much free space a drive has:
Dim SectorsPerCluster&
Dim BytesPerSector&
Dim NumberOfFreeClusters&
Dim TotalNumberOfClusters&
Dim FreeBytes&
dummy& = GetDiskFreeSpace("c:\", SectorsPerCluster, _
BytesPerSector, NumberOfFreeClusters, TotalNumberOfClusters)
FreeBytes = NumberOfFreeClusters * SectorsPerCluster * _
BytesPerSector
The Long FreeBytes contains the number of free bytes on the drive.
----------------------------------------------------------------------
Changing the screen resolution
A big problem for many vb-programmers is how to change the screen resolution, also because in the Api-viewer the variable for EnumDisplaySettings and ChangeDisplaySettings is missing!
1. Code for the basic-module
Declare Function EnumDisplaySettings Lib "user32" _
Alias "EnumDisplaySettingsA" _
(ByVal lpszDeviceName As Long, _
ByVal iModeNum As Long, _
lpDevMode As Any) As BooleanDeclare Function ChangeDisplaySettings Lib "user32" _
Alias "ChangeDisplaySettingsA" _
(lpDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As LongPublic Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type
Example
Changes the resolution to 640x480 with the current colordepth.
Dim DevM As DEVMODE
'Get the info into DevM
erg& = EnumDisplaySettings(0&, 0&, DevM)
'We don't change the colordepth, because a
'rebot will be necessary
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
DevM.dmPelsWidth = 640 'ScreenWidth
DevM.dmPelsHeight = 480 'ScreenHeight
'DevM.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)
'Now change the display and check if possibleerg& = ChangeDisplaySettings(DevM, CDS_TEST)
'Check if succesfullSelect Case erg&
Case DISP_CHANGE_RESTART
  an = MsgBox("You've to reboot", vbYesNo + vbSystemModal, "Info")
  If an = vbYes Then
    erg& = ExitWindowsEx(EWX_REBOOT, 0&)
  End If
Case DISP_CHANGE_SUCCESSFUL
  erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
  MsgBox "Everything's ok", vbOKOnly + vbSystemModal, "It worked!"
Case Else
  MsgBox "Mode not supported", vbOKOnly + vbSystemModal, "Error"
End SelectEnd Sub
----------------------------------------------------------------------
How to display the item which the mouse is over in a list box
I have had many letters which have asked me how to you display in a tooltip or some other means, such as a text box, the current item's text in a list box which the mouse pointer is hovering over. I now have the answer which uses the SendMessage API.
Start A new Standard-EXE project, form1 is created by default.
Add a list box and a text box to form1.
Open up the code window for Form1 and type the following
Option Explicit
Private Declare Function SendMessage Lib _
"user32" Alias "SendMessageA" (ByVal hwnd _
As Long, ByVal wMsg As Long, ByVal wParam _
As Long, lParam As Any) As Long
Private Const LB_ITEMFROMPOINT = &H1A9
Private Sub Form_Load()
With List1
  .AddItem "Visit"
  .AddItem "Steve Anderson Web Site AT"
  .AddItem "http://www.microweird.demon.co.uk"
End With
End Sub
Private Sub List1_MouseMove(Button _
As Integer, Shift As Integer, X As _
Single, Y As Single)
Dim lXPoint As Long
Dim lYPoint As Long
Dim lIndex As Long
If Button = 0 Then ' if no button was pressed
  lXPoint = CLng(X / Screen.TwipsPerPixelX)
  lYPoint = CLng(Y / Screen.TwipsPerPixelY)
  With List1
    ' get selected item from list
    lIndex = SendMessage(.hwnd, _
    LB_ITEMFROMPOINT, 0, ByVal _
    ((lYPoint * 65536) + lXPoint))
    ' show tip or clear last one
    If (lIndex >= 0) And _
    (lIndex <= .ListCount) Then
      .ToolTipText = .List(lIndex)
      Text1.Text = .List(lIndex)
    Else
      .ToolTipText = ""
    End If
  End With
End If
End Sub
Run the project(F5) and hover your cursor over different items in the list box and they will be displayed in a tooltip and in Text1.
----------------------------------------------------------------------
Finding out the amount of free memory
It is easy to return the amount of free memory in windows, using the GlobalMemoryStatus API call. Insert the following into a module's declarations section:
Public Type MEMORYSTATUS
dwLength As Long
dwMemoryLoad As Long
dwTotalPhys As Long
dwAvailPhys As Long
dwTotalPageFile As Long
dwAvailPageFile As Long
dwTotalVirtual As Long
dwAvailVirtual As Long
End TypePublic Declare Sub GlobalMemoryStatus _
Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Now, add this code to get the values:
Dim MS As MEMORYSTATUS
MS.dwLength = Len(MS)
GlobalMemoryStatus MS
' MS.dwMemoryLoad contains percentage memory used
' MS.dwTotalPhys contains total amount of physical memory in bytes
' MS.dwAvailPhys contains available physical memory
' MS.dwTotalPageFile contains total amount of memory in the page file
' MS.dwAvailPageFile contains available amount of memory in the page file
' MS.dwTotalVirtual contains total amount of virtual memory
' MS.dwAvailVirtual contains available virtual memory
You could use this in about boxes or making a memory monitoring system
----------------------------------------------------------------------
```

