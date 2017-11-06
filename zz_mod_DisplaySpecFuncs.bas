Attribute VB_Name = "zz_mod_DisplaySpecFuncs"
Option Compare Database
Option Explicit

'VGC 09/29/2015: CHANGES!

Private Const THIS_NAME As String = "zz_mod_DisplaySpecFuncs"
' **

Public Sub DisplaySpecs_List()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "DisplaySpecs_List"

        Dim strComputer As String
        Dim objWMIService As Object
        Dim colItems As Object
        Dim objItem As Object

110   On Error Resume Next

120     strComputer = "."
130     Set objWMIService = GetObject("winmgmts:" & _
          "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

140     Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DisplayConfiguration")

150     For Each objItem In colItems
160       Debug.Print "'Bits Per Pel: " & objItem.BitsPerPel
170       Debug.Print "'Device Name: " & objItem.DeviceName
180       Debug.Print "'Display Flags: " & objItem.DisplayFlags
190       Debug.Print "'Display Frequency: " & objItem.DisplayFrequency
200       Debug.Print "'Driver Version: " & objItem.DriverVersion
210       Debug.Print "'Log Pixels: " & objItem.LogPixels
220       Debug.Print "'Pels Height: " & objItem.PelsHeight
230       Debug.Print "'Pels Width: " & objItem.PelsWidth
240       Debug.Print "'Setting ID: " & objItem.SettingID
250       Debug.Print "'Specification Version: " & objItem.SpecificationVersion
260       Debug.Print
270     Next

        'Bits Per Pel: 32
        'Device Name: Intel(R) HD Graphics Family
        'Display Flags: 0
        'Display Frequency: 60
        'Driver Version:
        'Log Pixels: 96
        'Pels Height: 768
        'Pels Width: 1024
        'Setting ID: Intel(R) HD Graphics Family
        'Specification Version: 1025

EXITP:
280     Exit Sub

ERRH:
290     Select Case ERR.Number
        Case Else
300       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
310     End Select
320     Resume EXITP

End Sub

Public Function Frm_Ctl_SetProp() As Boolean

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Frm_Ctl_SetProp"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
        Dim lngCtls As Long, arr_varCtl As Variant
        Dim strLastForm As String
        Dim lngFrmsChanged As Long, lngCtlsChanged As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCtl().
        Const C_DID  As Integer = 0
        Const C_DNAM As Integer = 1
        Const C_FID  As Integer = 2
        Const C_FNAM As Integer = 3
        Const C_CID  As Integer = 4
        Const C_CNAM As Integer = 5
        Const C_TYP  As Integer = 6
        Const C_CAP  As Integer = 7
        Const C_FCLR As Integer = 8
        Const C_FONT As Integer = 9
        Const C_FSIZ As Integer = 10
        Const C_BOLD As Integer = 11

410   On Error GoTo 0

420     Set dbs = CurrentDb
430     With dbs
          ' ** zzz_qry_Form_Control_03_01 (zzz_qry_Form_Control_01 (tblForm_Control,
          ' ** just 'acCommandButton'), just solid buttons), just ctlspec_forecolor = 0.
440       Set qdf = .QueryDefs("zzz_qry_Form_Control_03_03_02")
450       Set rst = qdf.OpenRecordset
460       With rst
470         .MoveLast
480         lngCtls = .RecordCount
490         .MoveFirst
500         arr_varCtl = .GetRows(lngCtls)
            ' ******************************************************
            ' ** Array: arr_varCtl()
            ' **
            ' **   Field  Element  Name                 Constant
            ' **   =====  =======  ===================  ==========
            ' **     1       0     dbs_id               C_DID
            ' **     2       1     dbs_name             C_DNAM
            ' **     3       2     frm_id               C_FID
            ' **     4       3     frm_name             C_FNAM
            ' **     5       4     ctl_id               C_CID
            ' **     6       5     ctl_name             C_CNAM
            ' **     7       6     ctltype_type         C_TYP
            ' **     8       7     ctl_caption          C_CAP
            ' **     9       8     ctlspec_forecolor    C_FCLR
            ' **    10       9     ctlspec_fontname     C_FONT
            ' **    11      10     ctlspec_fontsize     C_FSIZ
            ' **    12      11     ctlspec_fontbold     C_BOLD
            ' **
            ' ******************************************************
510         .Close
520       End With
530       Set rst = Nothing
540       Set qdf = Nothing
550       .Close
560     End With
570     Set dbs = Nothing

580     Debug.Print "'CTLS: " & CStr(lngCtls)
590     DoEvents

600     Do While Forms.Count > 0
610       DoCmd.Close acForm, Forms(0).Name
620     Loop

630     strLastForm = vbNullString
640     lngFrmsChanged = 0&: lngCtlsChanged = 0&

650     For lngX = 0& To (lngCtls - 1&)
660       If arr_varCtl(C_FNAM, lngX) <> strLastForm Then
670         If lngX > 0& Then
680           DoCmd.Close acForm, Forms(0).Name, acSaveYes
690         End If
700         DoCmd.OpenForm arr_varCtl(C_FNAM, lngX), acDesign, , , , acHidden
710         Set frm = Forms(0)
720         strLastForm = frm.Name
730         lngFrmsChanged = lngFrmsChanged + 1&
740       End If
750       With frm
760         Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
770         With ctl
780           .ForeColor = 3026478
790           lngCtlsChanged = lngCtlsChanged + 1&
800         End With
810       End With
820       Set ctl = Nothing
830     Next  ' ** lngX.
840     DoCmd.Close acForm, Forms(0).Name, acSaveYes
850     Set frm = Nothing

860     Debug.Print "'FRMS EDITED: " & CStr(lngFrmsChanged)
870     Debug.Print "'CTLS EDITED: " & CStr(lngCtlsChanged)
880     DoEvents

        'CTLS: 112
        'FRMS EDITED: 32
        'CTLS EDITED: 112
        'DONE!
890     Beep
900     Debug.Print "'DONE!"

EXITP:
910     Set ctl = Nothing
920     Set frm = Nothing
930     Set rst = Nothing
940     Set qdf = Nothing
950     Set dbs = Nothing
960     Frm_Ctl_SetProp = blnRetVal
970     Exit Function

ERRH:
980     blnRetVal = False
990     Select Case ERR.Number
        Case Else
1000      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1010    End Select
1020    Resume EXITP

End Function
