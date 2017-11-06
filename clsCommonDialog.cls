VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "clsCommonDialog"

'VGC 03/23/2017: CHANGES!

' ** This code is from the Microsoft Knowledge Base.
' ***************************************************
' ** VERSION 1.0 CLASS
' ** BEGIN
' **   MultiUse = -1            ' ** True
' **   Persistable = 0          ' ** NotPersistable
' **   DataBindingBehavior = 0  ' ** vbNone
' **   DataSourceBehavior = 0   ' ** vbNone
' **   MTSTransactionMode = 0   ' ** NotAnMTSObject
' ** End
' ** Attribute VB_Name = "clsCommonDialog"
' ** Attribute VB_GlobalNameSpace = False
' ** Attribute VB_Creatable = True
' ** Attribute VB_PredeclaredId = False
' ** Attribute VB_Exposed = True
' ***************************************************

' ** API function to retrieve extended error information.
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

' ** Internal property buffers.
Private intAction As Integer       ' ** Internal buffer for Action property.
Private blnCancelError As Boolean  ' ** Internal buffer for CancelError property.
Private lngColor As Long           ' ** Internal buffer for Color property.
'Private lngCopies As Long          ' ** Internal buffer for lCopies property.  NOT USED!
Private strDefaultExt As String    ' ** Internal buffer for sDefaultExt property.
Private strDialogTitle As String   ' ** Internal buffer for DialogTitle property.
Private strFileName As String      ' ** Internal buffer for FileName property.
Private strFileTitle As String     ' ** Internal buffer for FileTitle property.
Private strFilter As String        ' ** Internal buffer for Filter property.
Private intFilterIndex As Integer  ' ** Internal buffer for FilterIndex property.
Private lngFlags As Long           ' ** Internal buffer for Flags property.
Private lngHdc As Long             ' ** Internal buffer for hdc property.
Private strInitDir As String       ' ** Internal buffer for InitDir property.
Private lngMax As Long             ' ** Internal buffer for Max property.
Private lngMaxFileSize As Long     ' ** Internal buffer for MaxFileSize property.
Private lngMin As Long             ' ** Internal buffer for Min property.
Private objObject As Object        ' ** Internal buffer for Object property.
Private lngApiReturn As Long       ' ** Internal buffer for APIReturn property.
Private lngExtendedError As Long   ' ** Internal buffer for ExtendedError property.
' **

Private Sub Class_Initialize()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Class_Initialize"

110     With Me
120       .hDC = 0
130       .MaxFileSize = 256
140       .Max = 256
150       .FileTitle = vbNullString
160       .DialogTitle = "Please Select a File"
170       .InitDir = vbNullString
180       .DefaultExt = vbNullString
190     End With

EXITP:
200     Exit Sub

ERRH:
210     Select Case ERR.Number
        Case Else
220       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
230     End Select
240     Resume EXITP

End Sub

Public Sub ShowColor()
' ** Display the color dialog box.

300   On Error GoTo ERRH

        Const THIS_PROC As String = "ShowColor"

        Dim typCS As COLORSTRUC
        Dim arr_lngCustomColors(15) As Long
        Dim lngCustomColorSize As Long
        Dim lngCustomColorAddress As Long
        Dim lngMemHandle As Long
        Dim intX As Integer

        ' ** Initialize property buffers.
310     intAction = 3         ' ** Action property - ShowColor.
320     lngApiReturn = 0      ' ** APIReturn property.
330     lngExtendedError = 0  ' ** ExtendedError property.

        ' ** Prepare typCS data.

        ' ** lStructSize As Long.
340     typCS.lStructSize = Len(typCS)

        ' ** hwndOwner As Long.
350     typCS.hwnd = 0&  ' ** lngHdc

        ' ** hInstance As Long.

        ' ** rgblRetult As Long.
360     typCS.rgbResult = lngColor

        ' ** lpCustColors As Long.
        ' ** Fill custom colors array with all white.
370     For intX = 0 To UBound(arr_lngCustomColors)
380       arr_lngCustomColors(intX) = &HFFFFFF
390     Next

        ' ** Get size of memory needed for custom colors.
400     lngCustomColorSize = Len(arr_lngCustomColors(0)) * 16
        ' ** Get a global memory block to hold a copy of the custom colors.
410     lngMemHandle = GlobalAlloc(GMEM_HND, lngCustomColorSize)  ' ** API Function: modWindowFunctions.

420     If lngMemHandle <> 0 Then

          ' ** Lock the custom color's global memory block.
430       lngCustomColorAddress = GlobalLock(lngMemHandle)
440       If lngCustomColorAddress <> 0 Then

            ' ** Copy custom colors to the global memory block.
450         Call CopyMemory(ByVal lngCustomColorAddress, arr_lngCustomColors(0), lngCustomColorSize)  ' ** API Function: modWindowFunctions.

460         typCS.lpCustColors = lngCustomColorAddress

            ' ** flags As Long.
470         typCS.flags = lngFlags

            ' ** lCustData As Long
            ' ** lpfnHook As Long
            ' ** lpTemplateName As String

            ' ** Call the ChooseColor API function.
480         lngApiReturn = ChooseColor(typCS)

            ' ** Handle return from ChooseColor API function.
490         Select Case lngApiReturn
            Case 0  ' ** User canceled.
500           If blnCancelError = True Then
                ' ** Generate an error.
510   On Error GoTo 0
520             ERR.Raise Number:=vbObjectError + 894, description:="Cancel Pressed"
530           End If
540         Case 1  ' ** User selected a color.
              ' ** Update property buffer.
550           lngColor = typCS.rgbResult
560         Case Else   'an error occured
              ' ** Call CommDlgExtendedError.
570           lngExtendedError = CommDlgExtendedError
580         End Select

590       End If
600     End If

EXITP:
610     Exit Sub

ERRH:
620     Select Case ERR.Number
        Case Else
630       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
640     End Select
650     Resume EXITP

End Sub

Private Sub ShowFileDialog(ByVal intAct As Integer)
' ** Display the file dialog for ShowOpen or ShowSave.

700   On Error GoTo ERRH

        Const THIS_PROC As String = "ShowFileDialog"

        Dim typOpenFile As OPENFILENAME
        Dim lngMaxSize As Long
        Dim strFileNameBuff As String
        Dim strFileTitleBuff As String

        ' ** Init property buffers.
710     intAction = intAct    ' ** Action property.
720     lngApiReturn = 0      ' ** APIReturn property.
730     lngExtendedError = 0  ' ** ExtendedError property.

        ' ** Prepare typOpenFile data.

        ' ** typOpenFile.lStructSize As Long.
740     typOpenFile.lStructSize = Len(typOpenFile)

        ' ** typOpenFile.hWndOwner As Long - init from hdc property.
750     typOpenFile.hwndOwner = Application.hWndAccessApp  ' **  0& : Just use 0 !

        ' ** typOpenFile.lpstrFilter As String - init from Filter property.
760     typOpenFile.lpstrFilter = sAPIFilter(strFilter)

        ' ** typOpenFile.nFilterIndex As Long - init from FilterIndex property.
770     typOpenFile.nFilterIndex = intFilterIndex

        ' ** typOpenFile.lpstrFile As String.
        ' ** Determine size of buffer from MaxFileSize property.
780     If lngMaxFileSize > 0 Then
790       lngMaxSize = lngMaxFileSize
800     Else
810       lngMaxSize = 256
820     End If

        ' ** typOpenFile.lpstrFile As Long - init from FileName property.
        ' ** Prepare strFileNameBuff.
830     strFileNameBuff = strFileName
        ' ** Pad with spaces.
840     While Len(strFileNameBuff) < lngMaxSize - 1
850       strFileNameBuff = strFileNameBuff & " "
860     Wend
        ' ** Trim to length of lngMaxFileSize - 1.
870     strFileNameBuff = Mid(strFileNameBuff, 1, lngMaxFileSize - 1)
        ' ** Null terminate.
880     strFileNameBuff = strFileNameBuff & Chr$(0)
890     typOpenFile.lpstrFile = strFileNameBuff

        ' ** nMaxFile As Long - init from MaxFileSize property.
900     If lngMaxFileSize <> 255 Then  ' ** Default is 255.
910       typOpenFile.nMaxFile = lngMaxFileSize
920     End If

        ' ** lpstrFileTitle As String - init from FileTitle property.
        ' ** Prepare strFileTitleBuff.
930     strFileTitleBuff = strFileTitle
        ' ** Pad with spaces.
940     While Len(strFileTitleBuff) < lngMaxSize - 1
950       strFileTitleBuff = strFileTitleBuff & " "
960     Wend
        ' ** Trim to length of lngMaxFileSize - 1.
970     strFileTitleBuff = Mid(strFileTitleBuff, 1, lngMaxFileSize - 1)
        ' ** Null terminate.
980     strFileTitleBuff = strFileTitleBuff & Chr$(0)
990     typOpenFile.lpstrFileTitle = strFileTitleBuff

        ' ** typOpenFile.lpstrInitialDir As String - init from InitDir property.
1000    typOpenFile.lpstrInitialDir = strInitDir

        ' ** typOpenFile.lpstrTitle As String - init from DialogTitle property.
1010    typOpenFile.lpstrTitle = strDialogTitle

        ' ** typOpenFile.flags As Long - init from Flags property.
1020    typOpenFile.flags = lngFlags

        ' ** typOpenFile.lpstrDefExt As String - init from DefaultExt property.
1030    typOpenFile.lpstrDefExt = strDefaultExt

        ' ** Call the GetOpenFileName API function.
1040    Select Case intAction
        Case 1  ' ** ShowOpen.
1050      lngApiReturn = GetOpenFileName(typOpenFile)
1060    Case 2  ' ** ShowSave.
1070      lngApiReturn = GetSaveFileName(typOpenFile)
1080    Case Else  ' ** Unknown action.
1090      Exit Sub
1100    End Select

        ' ** Handle return from GetOpenFileName API function.
1110    Select Case lngApiReturn
        Case 0  ' ** User canceled.
1120      If blnCancelError = True Then
            ' ** Generate an error.
1130        ERR.Raise (2001)  ' ** You canceled the previous operation.
1140      End If
1150    Case 1  '** User selected or entered a file.
          ' ** strFileName gets part of typOpenFile.lpstrFile to the left of first Chr$(0).
1160      strFileName = sLeftOfNull(typOpenFile.lpstrFile)
1170      strFileTitle = sLeftOfNull(typOpenFile.lpstrFileTitle)
1180    Case Else  ' ** An error occured.
          ' ** Call CommDlgExtendedError.
1190      lngExtendedError = CommDlgExtendedError
1200    End Select

EXITP:
1210    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1240    End Select
1250    Resume EXITP

End Sub

Public Sub ShowOpen()
' ** Display the file open dialog box.

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowOpen"

1310    ShowFileDialog (1)  ' ** Action property - ShowOpen.

EXITP:
1320    Exit Sub

ERRH:
1330    Select Case ERR.Number
        Case Else
1340      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1350    End Select
1360    Resume EXITP

End Sub

Public Sub ShowSave()
' ** Display the file save dialog box.

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowSave"

1410    ShowFileDialog (2)  ' ** Action property - ShowSave.

EXITP:
1420    Exit Sub

ERRH:
1430    Select Case ERR.Number
        Case Else
1440      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1450    End Select
1460    Resume EXITP

End Sub

Public Property Get Filename() As String
' ** Return object's FileName property.

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "FileName"

1510    Filename = strFileName

EXITP:
1520    Exit Property

ERRH:
1530    Select Case ERR.Number
        Case Else
1540      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1550    End Select
1560    Resume EXITP

End Property

Public Property Let Filename(strNewValue As String)
' ** Assign object's FileName property.

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "FileName"

1610    strFileName = strNewValue

EXITP:
1620    Exit Property

ERRH:
1630    Select Case ERR.Number
        Case Else
1640      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1650    End Select
1660    Resume EXITP

End Property

Public Property Get Filter() As String
' ** Return object's Filter property.

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Filter"

1710    Filter = strFilter

EXITP:
1720    Exit Property

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1750    End Select
1760    Resume EXITP

End Property

Public Property Let Filter(strNewValue As String)
' ** Assign object's Filter property.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Filter"

1810    strFilter = strNewValue

EXITP:
1820    Exit Property

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1850    End Select
1860    Resume EXITP

End Property

Public Property Get Action() As Integer
' ** Return object's Action property.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Action"

1910    Action = intAction

EXITP:
1920    Exit Property

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1950    End Select
1960    Resume EXITP

End Property

Public Property Get FilterIndex() As Integer
' ** Return object's FilterIndex property.

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "FilterIndex"

2010    FilterIndex = intFilterIndex

EXITP:
2020    Exit Property

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2050    End Select
2060    Resume EXITP

End Property

Public Property Let FilterIndex(intNewValue As Integer)
' ** Assign object's FilterIndex property.

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "FilterIndex"

2110    intFilterIndex = intNewValue

EXITP:
2120    Exit Property

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2150    End Select
2160    Resume EXITP

End Property

Public Property Get CancelError() As Boolean
' ** Return object's CancelError property.

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "CancelError"

2210    CancelError = blnCancelError

EXITP:
2220    Exit Property

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2250    End Select
2260    Resume EXITP

End Property

Public Property Let CancelError(blnNewValue As Boolean)
' ** Assign object's CancelError property.

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "CancelError"

2310    blnCancelError = blnNewValue

EXITP:
2320    Exit Property

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2350    End Select
2360    Resume EXITP

End Property

Public Property Get Color() As Long
' ** Return object's Color property.

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Color"

2410    Color = lngColor

EXITP:
2420    Exit Property

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2450    End Select
2460    Resume EXITP

End Property

Public Property Let Color(lngNewValue As Long)
' ** Assign object's Color property.

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Color"

2510    lngColor = lngNewValue

EXITP:
2520    Exit Property

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2550    End Select
2560    Resume EXITP

End Property

Public Property Get DefaultExt() As String
' ** Return object's DefaultExt (file extention) property.

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "DefaultExt"

2610    DefaultExt = strDefaultExt

EXITP:
2620    Exit Property

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2650    End Select
2660    Resume EXITP

End Property

Public Property Let DefaultExt(strNewValue As String)
' ** Assign object's DefaultExt property.

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "DefaultExt"

2710    strDefaultExt = strNewValue

EXITP:
2720    Exit Property

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2750    End Select
2760    Resume EXITP

End Property

Public Property Get DialogTitle() As String
' ** Return object's DialogTitle property.

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "DialogTitle"

2810    DialogTitle = strDialogTitle

EXITP:
2820    Exit Property

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2850    End Select
2860    Resume EXITP

End Property

Public Property Let DialogTitle(strNewValue As String)
' ** Assign object's DialogTitle property.

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "DialogTitle"

2910    strDialogTitle = strNewValue

EXITP:
2920    Exit Property

ERRH:
2930    Select Case ERR.Number
        Case Else
2940      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2950    End Select
2960    Resume EXITP

End Property

Public Property Get FileTitle() As String
' ** Return object's FileTitle property

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "FileTitle"

3010    FileTitle = strFileTitle

EXITP:
3020    Exit Property

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3050    End Select
3060    Resume EXITP

End Property

Public Property Let FileTitle(strNewValue As String)
' ** Assign object's FileTitle property

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "FileTitle"

3110    strFileTitle = strNewValue

EXITP:
3120    Exit Property

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3150    End Select
3160    Resume EXITP

End Property

Public Property Get flags() As Long
' ** Return object's Flags property.

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "flags"

3210    flags = lngFlags

EXITP:
3220    Exit Property

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3250    End Select
3260    Resume EXITP

End Property

Public Property Let flags(lngNewValue As Long)
' ** Assign object's Flags property.

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "flags"

3310    lngFlags = lngNewValue

EXITP:
3320    Exit Property

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3350    End Select
3360    Resume EXITP

End Property

Public Property Get hDC() As Long
' ** Return object's hDC property.

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "hdc"

3410    hDC = lngHdc

EXITP:
3420    Exit Property

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3450    End Select
3460    Resume EXITP

End Property

Public Property Let hDC(lngNewValue As Long)
' ** Assign object's hDC property.

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "hdc"

3510    lngHdc = lngNewValue

EXITP:
3520    Exit Property

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3550    End Select
3560    Resume EXITP

End Property

Public Property Get InitDir() As String
' ** Return object's InitDir property.

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "InitDir"

3610    InitDir = strInitDir

EXITP:
3620    Exit Property

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3650    End Select
3660    Resume EXITP

End Property

Public Property Let InitDir(strNewValue As String)
' ** Assign object's InitDir property.

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "InitDir"

3710    strInitDir = strNewValue

EXITP:
3720    Exit Property

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3750    End Select
3760    Resume EXITP

End Property

Public Property Get Max() As Long
' ** Return object's Max property.

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Max"

3810    Max = lngMax

EXITP:
3820    Exit Property

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3850    End Select
3860    Resume EXITP

End Property

Public Property Let Max(lngNewValue As Long)
' ** Assign object's Max property.

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Max"

3910    lngMax = lngNewValue

EXITP:
3920    Exit Property

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3950    End Select
3960    Resume EXITP

End Property

Public Property Get min() As Long
' ** Return object's Min property.

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Min"

4010    min = lngMin

EXITP:
4020    Exit Property

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4050    End Select
4060    Resume EXITP

End Property

Public Property Let min(lngNewValue As Long)
' ** Assign object's Min property.

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Min"

4110    lngMin = lngNewValue

EXITP:
4120    Exit Property

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4150    End Select
4160    Resume EXITP

End Property

Public Property Get MaxFileSize() As Long
' ** Return object's MaxFileSize property.

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "MaxFileSize"

4210    MaxFileSize = lngMaxFileSize

EXITP:
4220    Exit Property

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4250    End Select
4260    Resume EXITP

End Property

Public Property Let MaxFileSize(lngNewValue As Long)
' ** Assign object's MaxFileSize property.

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "MaxFileSize"

4310    lngMaxFileSize = lngNewValue

EXITP:
4320    Exit Property

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4350    End Select
4360    Resume EXITP

End Property

Public Property Get Object() As Object
' ** Return object's Object property.

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Object"

4410    Object = objObject

EXITP:
4420    Exit Property

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4450    End Select
4460    Resume EXITP

End Property

Public Property Let Object(objNewValue As Object)
' ** Assign object's Object property.

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Object"

4510    objObject = objNewValue

EXITP:
4520    Exit Property

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4550    End Select
4560    Resume EXITP

End Property

Public Property Get APIReturn() As Long
' ** Return object's APIReturn property

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "APIReturn"

4610    APIReturn = lngApiReturn

EXITP:
4620    Exit Property

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4650    End Select
4660    Resume EXITP

End Property

Public Property Get ExtendedError() As Long
' ** Return object's ExtendedError property

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "ExtendedError"

4710    ExtendedError = lngExtendedError

EXITP:
4720    Exit Property

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4750    End Select
4760    Resume EXITP

End Property

Public Function BrowseFolder(strDialogTitle As String) As String
' ** This code was originally written by Terry Kreft.
' ** It is not to be altered or distributed,
' ** except as part of an application.
' ** You are free to use it in any application,
' ** provided the copyright notice is left unchanged.
' **
' ** Code courtesy of
' **   Terry Kreft

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "BrowseFolder"

        Dim lngX As Long, typBrowseInfo As BROWSEINFO, lngIList As Long
        Dim strPath As String, intPos01 As Integer
        Dim strRetVal As String

4810    strRetVal = vbNullString

4820    With typBrowseInfo
4830      .hwndOwner = hWndAccessApp
4840      .lpszTitle = strDialogTitle
4850      .ulFlags = BIF_RETURNONLYFSDIRS
4860    End With

4870    lngIList = SHBrowseForFolder(typBrowseInfo)  ' ** API Function: Above.
4880    strPath = Space$(512)
4890    lngX = SHGetPathFromIDList(ByVal lngIList, ByVal strPath)  ' ** API Function: Above.

4900    If lngX Then
4910      intPos01 = InStr(strPath, Chr(0))
4920      strRetVal = Left(strPath, intPos01 - 1)
4930    Else
4940      strRetVal = vbNullString
4950    End If

EXITP:
4960    BrowseFolder = strRetVal
4970    Exit Function

ERRH:
4980    strRetVal = RET_ERR
4990    Select Case ERR.Number
        Case Else
5000      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
5010    End Select
5020    Resume EXITP

End Function

Private Function sByteArrayToString(arr_bytBytes() As Byte) As String
' ** Return a string from a byte array.

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "sByteArrayToString"

        Dim lngBytePoint As Long, lngByteVal As Long
        Dim strRetVal As String

5110    strRetVal = vbNullString

        ' ** Init array pointer.
5120    lngBytePoint = LBound(arr_bytBytes)

        ' ** Fill strRetVal with characters in array.
5130    Do While lngBytePoint <= UBound(arr_bytBytes)

5140      lngByteVal = arr_bytBytes(lngBytePoint)

          ' ** Stop and return strRetVal as it is if Chr$(0) is encountered.
5150      If lngByteVal = 0& Then
5160        Exit Do
5170      Else
5180        strRetVal = strRetVal & Chr$(lngByteVal)
5190      End If

5200      lngBytePoint = lngBytePoint + 1

5210    Loop

EXITP:
5220    sByteArrayToString = strRetVal
5230    Exit Function

ERRH:
5240    strRetVal = RET_ERR
5250    Select Case ERR.Number
        Case Else
5260      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
5270    End Select
5280    Resume EXITP

End Function

Private Function sAPIFilter(strIn) As String
' ** Prepares strIn for use as a filter string in API common dialog functions.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "sAPIFilter"

        Dim lngChrNdx As Long
        Dim strOneChr As String
        Dim strRetVal As String

        ' ** Convert any | characters to nulls.
5310    For lngChrNdx = 1 To Len(strIn)
5320      strOneChr = Mid(strIn, lngChrNdx, 1)
5330      If strOneChr = "|" Then
5340        strRetVal = strRetVal & Chr$(0)
5350      Else
5360        strRetVal = strRetVal & strOneChr
5370      End If
5380    Next

        ' ** Add a null to the end.
5390    strRetVal = strRetVal & Chr$(0)

EXITP:
5400    sAPIFilter = strRetVal
5410    Exit Function

ERRH:
5420    strRetVal = RET_ERR
5430    Select Case ERR.Number
        Case Else
5440      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
5450    End Select
5460    Resume EXITP

End Function

Private Function sLeftOfNull(ByVal strIn As String) As String
' ** Returns the part of strIn preceding Chr$(0).

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "sLeftOfNull"

        Dim lngNullPos As Long
        Dim strRetVal As String

        ' ** Init output.
5510    strRetVal = strIn

        ' ** Get position of first Chr$(0) in strIn.
5520    lngNullPos = InStr(strIn, Chr$(0))

        ' ** Return part of strIn to left of first Chr$(0) if found.
5530    If lngNullPos > 0 Then
5540      strRetVal = Mid(strIn, 1, lngNullPos - 1)
5550    End If

EXITP:
5560    sLeftOfNull = strRetVal
5570    Exit Function

ERRH:
5580    strRetVal = RET_ERR
5590    Select Case ERR.Number
        Case Else
5600      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
5610    End Select
5620    Resume EXITP

End Function
