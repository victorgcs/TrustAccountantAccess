VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountComments"

'VGC 04/15/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   chkIncludeCurrency_AfterUpdate()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Include Archive:  H {chkIncludeArchive}
' **   Include Currency: L {chkIncludeCurrency}
' **   Update:           U {cmdUpdate}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAccountComments_Sub}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngClose_Left As Long, lngUpdate_Left As Long, lngSub_Width As Long, lngCurrID_Width As Long

Private blnAskedToAbandon As Boolean, blnViewOnly As Boolean, blnHasArchive As Boolean, blnThisArchive As Boolean
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim usr As DAO.User, grp As DAO.Group, frm As Access.Form
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           blnAskedToAbandon = False: blnHasArchive = False: blnThisArchive = False
170           gblnHasForExThis = False

180           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
190           SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Variables are fed empty, then populated ByRef.
210           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

220           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

230           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
240           If lngMonitorCnt > 1& Then
250             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
260             lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
270             If lngMonitorNum > 1& Then
280               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
290             End If
300           End If

310           .accountno = gstrAccountNo
320           .shortname = gstrAccountName

              ' ** Check user status, and lock all fields if ViewOnly.
330           blnViewOnly = False
340           For Each grp In DBEngine.Workspaces(0).Groups
350             If grp.Name = "ViewOnly" Then
360               For Each usr In grp.Users
370                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
380                   blnViewOnly = True
390                   Exit For
400                 End If
410               Next
420             End If
430           Next

440           Set frm = .frmAccountComments_Sub.Form

450           .form_width_line.Left = 0&  ' ** Just get it out of the way.

460           lngForm_Width = .Width
470           lngClose_Left = .cmdClose.Left
480           lngUpdate_Left = .cmdUpdate.Left
490           lngSub_Width = .frmAccountComments_Sub.Width
500           lngCurrID_Width = ((frm.curr_id.Left + frm.curr_id.Width) - (frm.Cost.Left + frm.Cost.Width))

510           SetArchiveOption  ' ** Procedure: Below.
520           DoEvents

530           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
540           DoEvents

550           chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
560           chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.

570           If blnViewOnly = True Then
580             frm.Comment.Locked = True
590           End If

600           .frmAccountComments_Sub.SetFocus
610           frm.FocusHolder.SetFocus
620           .cmdUpdate.Enabled = False

630           blnIsOpen = False

640         Else
650           Cancel = -1
660         End If
670       Else
680         Cancel = -1
690       End If
700     End With

710     If Cancel = -1 Then
720       Beep
730       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
740     End If

EXITP:
750     Set frm = Nothing
760     Set usr = Nothing
770     Set grp = Nothing
780     Exit Sub

ERRH:
790     DoCmd.Hourglass False
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub Form_Load()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

910     With Me

920       CLR_DISABLED_FG = CLR_DKGRY
930       CLR_DISABLED_BG = CLR_LTTEAL

940       .accountno.ForeColor = CLR_DISABLED_FG
950       .accountno.BackColor = CLR_DISABLED_BG
960       .shortname.ForeColor = CLR_DISABLED_FG
970       .shortname.BackColor = CLR_DISABLED_BG

980     End With

EXITP:
990     Exit Sub

ERRH:
1000    DoCmd.Hourglass False
1010    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Ignore.
1020    Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality.
        ' **   Save:             S {cmdSave on frmAccountComments_Sub}

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyEscape
1170        intRetVal = 0
1180        cmdClose_Click
1190      End Select
1200    End If

        ' ** Alt keys.
1210    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1220      Select Case intRetVal
          Case vbKeyX
1230        intRetVal = 0
1240        cmdClose_Click  ' ** Procedure: Below.
1250      End Select
1260    End If

        ' ** Ctrl keys.
1270    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1280      Select Case intRetVal
          Case vbKeyS
1290        With Me
1300          intRetVal = 0
1310          lngRecsCur = .frmAccountComments_Sub.Form.RecCnt  ' ** Form Function: frmAccountComments_Sub.
1320          If lngRecsCur > 0& Then
1330            .frmAccountComments_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountComments_Sub.
1340          Else
1350            Beep
1360          End If
1370        End With
1380      End Select
1390    End If

        ' ** Ctrl-Shift keys.
1400    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1410      Select Case intRetVal
          Case vbKeyF
1420        With Me
1430          intRetVal = 0
1440          .FocusHolder.SetFocus
1450        End With
1460      End Select
1470    End If

EXITP:
1480    KeyCode = intRetVal
1490    Exit Sub

ERRH:
1500    intRetVal = 0
1510    Select Case ERR.Number
        Case Else
1520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub Form_Timer()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1610    With Me
1620      .TimerInterval = 0&
1630      If gblnSetFocus = True Then
1640        gblnSetFocus = False
1650        strFormName = vbNullString
1660  On Error Resume Next
1670        Do Until strFormName = THIS_NAME
1680          DoCmd.SelectObject acForm, THIS_NAME, False
1690          strFormName = Screen.ActiveForm.Name
1700        Loop
1710  On Error GoTo ERRH
1720      End If
1730      gblnHasForExThis = HasForEx_Acct(gstrAccountNo, "L")  ' ** Module Function: modCurrencyFuncs.
1740      Select Case gblnHasForEx
          Case True
1750        Select Case gblnHasForExThis
            Case True
1760          If Pref_HasPref(THIS_NAME, "chkIncludeCurrency") = False Then  ' ** Module Procedure: modPreferenceFuncs.
1770            .chkIncludeCurrency = True
1780            chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
1790          End If
1800          .chkIncludeCurrency.Enabled = True
1810        Case False
1820          .chkIncludeCurrency.Enabled = True
1830        End Select
1840      Case False
1850        .chkIncludeCurrency = False
1860        chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
1870        .chkIncludeCurrency.Enabled = False
1880      End Select
1890      DoCmd.Hourglass False

1900    End With

EXITP:
1910    Exit Sub

ERRH:
1920    DoCmd.Hourglass False
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean
        Dim blnContinue As Boolean

2010    With Me

2020      DoCmd.Hourglass True
2030      DoEvents

2040      blnContinue = True

2050      If .chkUpdated = True And blnAskedToAbandon = False Then
2060        Beep
2070        DoCmd.Hourglass False
2080        If MsgBox("You have made changes, but have not confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") = vbYes Then
              ' ** Abandon changes.
2090        Else
2100          blnContinue = False
2110        End If
2120      End If

2130      If blnContinue = True Then

2140        DoCmd.Hourglass True
2150        DoEvents

2160        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2170        blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

            ' ** Make sure all reports are closed.
2180        If Reports.Count > 0 Then
2190          Do While Reports.Count > 0
2200            DoCmd.Close acReport, Reports(0).Name
2210          Loop
2220        End If

2230        gblnSetFocus = True
2240        gstrReturningForm = THIS_NAME
2250        strDocName = "frmMenu_Account"
2260        If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
2270          DoCmd.SelectObject acForm, strDocName, False
2280          Forms(strDocName).TimerInterval = 250&
2290        Else
2300          DoCmd.OpenForm strDocName, , , , , , "frmMenu_Main"
2310        End If

2320      End If
2330    End With

EXITP:
2340    Exit Sub

ERRH:
2350    DoCmd.Hourglass False
2360    Select Case ERR.Number
        Case Else
2370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2380    End Select
2390    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim blnContinue As Boolean

2410    blnContinue = True

2420    With Me

2430      .TimerInterval = 0&

2440      If .chkUpdated = True And blnViewOnly = False Then
2450        blnAskedToAbandon = False
2460        Beep
2470        If MsgBox("You have made changes, but have not confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") = vbYes Then
2480          blnAskedToAbandon = True
2490        Else
2500          blnContinue = False
2510        End If
2520      ElseIf blnViewOnly = True Then
2530        blnAskedToAbandon = True
2540      End If

2550      If blnContinue = True Then

2560        DoCmd.Close acForm, THIS_NAME

2570      End If

2580    End With

EXITP:
2590    Exit Sub

ERRH:
2600    Select Case ERR.Number
        Case 2501  ' ** The Close action was canceled.
          ' ** Ignore.
2610    Case Else
2620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2630    End Select
2640    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab
2770        With Me
2780          intRetVal = 0
2790          .frmAccountComments_Sub.SetFocus
2800          .frmAccountComments_Sub.Form.MoveRec acCmdRecordsGoToFirst
2810          .frmAccountComments_Sub.Form.Comment.SetFocus
2820        End With
2830      End Select
2840    End If

        ' ** Shift keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2860      Select Case intRetVal
          Case vbKeyTab
2870        With Me
2880          intRetVal = 0
2890          If .cmdUpdate.Enabled = True Then
2900            .cmdUpdate.SetFocus
2910          ElseIf .chkIncludeCurrency.Enabled = True Then
2920            .chkIncludeCurrency.SetFocus
2930          ElseIf .chkIncludeArchive.Enabled = True Then
2940            .chkIncludeArchive.SetFocus
2950          Else
2960            .frmAccountComments_Sub.SetFocus
2970            .frmAccountComments_Sub.Form.MoveRec acCmdRecordsGoToLast
2980            .frmAccountComments_Sub.Form.Comment.SetFocus
2990          End If
3000        End With
3010      End Select
3020    End If

EXITP:
3030    KeyCode = intRetVal
3040    Exit Sub

ERRH:
3050    intRetVal = 0
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        'Dim cnxn As ADODB.Connection, rsx1 As ADODB.Recordset, rsx2 As ADODB.Recordset  ' ** Early binding.
        Dim cnxn As Object, rsx1 As Object, rsx2 As Object, rsx3 As Object               ' ** Late binding.
        Dim lngRecs1 As Long, lngRecs2 As Long, lngRecs3 As Long
        Dim blnInTrans As Boolean
        Dim lngX As Long

3110    With Me
3120      If blnViewOnly = False Then

3130        blnInTrans = False

            'Set rsx1 = New ADODB.Recordset             ' ** Early binding.
            'Set rsx2 = New ADODB.Recordset             ' ** Early binding.
3140        Set rsx1 = CreateObject("ADODB.Recordset")  ' ** Late binding.
3150        Set rsx2 = CreateObject("ADODB.Recordset")  ' ** Late binding.
3160        Set rsx3 = CreateObject("ADODB.Recordset")  ' ** Late binding.
3170        Set cnxn = CurrentProject.Connection
3180        rsx1.Open "tmpEdit10", cnxn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
3190        With rsx1  ' ** adOpenDynamic doesn't want to give a RecordCount!
3200          If .BOF = True And .EOF = True Then
                ' ** If the data source returns no records, the provider sets both the BOF and
                ' ** EOF properties to True, and the current record position is undefined. You can
                ' ** still add new data to this empty Recordset object if the cursor type allows it.
3210            lngRecs1 = 0&
3220          Else
3230            .MoveLast
3240            lngRecs1 = .RecordCount
3250            .MoveFirst
3260          End If
3270        End With
3280        rsx2.Open "ledger", cnxn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
3290        With rsx2
3300          If .BOF = True And .EOF = True Then
3310            lngRecs2 = 0&
3320          Else
3330            .MoveLast
3340            lngRecs2 = .RecordCount
3350            .MoveFirst
3360          End If
3370        End With
3380        rsx3.Open "LedgerArchive", cnxn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
3390        With rsx3
3400          If .BOF = True And .EOF = True Then
3410            lngRecs3 = 0&
3420          Else
3430            .MoveLast
3440            lngRecs3 = .RecordCount
3450            .MoveFirst
3460          End If
3470        End With

            ' ** AdOpen enumeration (CursorTypeEnum):
            ' **   -1  adOpenUnspecified  Does not specify the type of cursor.
            ' **    0  adOpenForwardOnly  Uses a forward-only cursor. Identical to a static cursor, except that you
            ' **                          can only scroll forward through records. This improves performance when
            ' **                          you need to make only one pass through a Recordset. (Default)
            ' **    1  adOpenKeyset       Uses a keyset cursor. Like a dynamic cursor, except that you can't see records
            ' **                          that other users add, although records that other users delete are inaccessible
            ' **                          from your Recordset. Data changes by other users are still visible.
            ' **    2  adOpenDynamic      Uses a dynamic cursor. Additions, changes, and deletions by other users are
            ' **                          visible, and all types of movement through the Recordset are allowed, except
            ' **                          for bookmarks, if the provider doesn't support them.
            ' **    3  adOpenStatic       Uses a static cursor. A static copy of a set of records that you can
            ' **                          use to find data or generate reports. Additions, changes, or
            ' **                          deletions by other users are not visible.

            ' ** AdLock enumeration (LockTypeEnum):
            ' **   -1  adLockUnspecified      Does not specify a type of lock. For clones, the clone is created
            ' **                              with the same lock type as the original.
            ' **    1  adLockReadOnly         Indicates read-only records. You cannot alter the data.
            ' **    2  adLockPessimistic      Indicates pessimistic locking, record by record. The provider does what
            ' **                              is necessary to ensure successful editing of the records, usually by
            ' **                              locking records at the data source immediately after editing.
            ' **    3  adLockOptimistic       Indicates optimistic locking, record by record. The provider uses
            ' **                              optimistic locking, locking records only when you call the Update method.
            ' **    4  adLockBatchOptimistic  Indicates optimistic batch updates. Required for batch update mode.

            ' ** AdCmd enumeration (CommandTypeEnum):
            ' **    -1  adCmdUnspecified  Does not specify the command type argument.
            ' **     1  adCmdText         Evaluates CommandText as a textual definition of a command or stored procedure call.
            ' **     2  adCmdTable        Evaluates CommandText as a table name whose columns are all returned
            ' **                          by an internally generated SQL query.
            ' **     4  adCmdStoredProc   Evaluates CommandText as a stored procedure name.
            ' **     8  adCmdUnknown      Indicates that the type of command in the CommandText property is not known. (Default)
            ' **   256  adCmdFile         Evaluates CommandText as the file name of a persistently stored Recordset.
            ' **                          Used with Recordset.Open or Requery only.
            ' **   512  adCmdTableDirect  Evaluates CommandText as a table name whose columns are all returned. Used with
            ' **                          Recordset.Open or Requery only. To use the Seek method, the Recordset must be
            ' **                          opened with adCmdTableDirect.
            ' **                          This value cannot be combined with the ExecuteOptionEnum value adAsyncExecute.

            ' ** AdEdit enumeration (EditModeEnum):
            ' **   0  adEditNone        Indicates that no editing operation is in progress.
            ' **   1  adEditInProgress  Indicates that data in the current record has been modified but not saved.
            ' **   2  adEditAdd         Indicates that the AddNew method has been called, and the current record
            ' **                        in the copy buffer is a new record that has not been saved in the database.
            ' **   4  adEditDelete      Indicates that the current record has been deleted.

3480        If lngRecs1 > 0 And lngRecs2 > 0 Then

3490          cnxn.BeginTrans
3500          blnInTrans = True
3510          For lngX = 1& To lngRecs1
3520            rsx2.MoveFirst
3530            rsx2.Find "journalno = " & CStr(rsx1.Fields("journalno"))
3540            If rsx2.EOF = True Then
3550              If lngRecs3 > 0& Then
3560                rsx3.MoveFirst
3570                rsx3.Find "journalno = " & CStr(rsx1.Fields("journalno"))
3580                If rsx3.EOF = True Then
3590                  MsgBox "Journal entry not found in ledger.", vbCritical + vbOKOnly, "Ledger Error"
3600                Else
3610                  If IsNull(rsx1.Fields("Comment")) = False Then
3620                    If IsNull(rsx3.Fields("description")) = False Then
3630                      If rsx1.Fields("Comment") <> rsx3.Fields("description") Then
3640                        rsx3.Fields("description") = rsx1.Fields("Comment")
3650                        rsx3.Update
3660                      End If
3670                    Else
3680                      rsx3.Fields("description") = rsx1.Fields("Comment")
3690                      rsx3.Update
3700                    End If
3710                  Else
3720                    If IsNull(rsx3.Fields("description")) = False Then
3730                      rsx3.Fields("description") = Null
3740                      rsx3.Update
3750                    End If
3760                  End If
3770                End If
3780              Else
3790                MsgBox "Journal entry not found in ledger.", vbCritical + vbOKOnly, "Ledger Error"
3800              End If
3810            Else
3820              If IsNull(rsx1.Fields("Comment")) = False Then
3830                If IsNull(rsx2.Fields("description")) = False Then
3840                  If rsx1.Fields("Comment") <> rsx2.Fields("description") Then
3850                    rsx2.Fields("description") = rsx1.Fields("Comment")
3860                    rsx2.Update
3870                  End If
3880                Else
3890                  rsx2.Fields("description") = rsx1.Fields("Comment")
3900                  rsx2.Update
3910                End If
3920              Else
3930                If IsNull(rsx2.Fields("description")) = False Then
3940                  rsx2.Fields("description") = Null
3950                  rsx2.Update
3960                End If
3970              End If
3980            End If
3990            If lngX < lngRecs1 Then rsx1.MoveNext
4000          Next
4010          cnxn.CommitTrans
4020          blnInTrans = False

4030          rsx1.Close
4040          rsx2.Close
4050          rsx3.Close
4060          cnxn.Close

4070          .chkUpdated = False
4080          .cmdClose.SetFocus
4090          .cmdUpdate.Enabled = False

4100          MsgBox "Updates completed.", vbInformation + vbOKOnly, ("Update Successful" & Space(40))

4110          DoCmd.Close acForm, THIS_NAME

4120        Else
4130          MsgBox "Records not found.", vbCritical + vbOKOnly, ("Update Failed" & Space(40))
4140          rsx1.Close
4150          rsx2.Close
4160          rsx3.Close
4170          cnxn.Close
4180        End If

4190      End If
4200    End With

EXITP:
4210    Set rsx1 = Nothing
4220    Set rsx2 = Nothing
4230    Set rsx3 = Nothing
4240    Set cnxn = Nothing
4250    Exit Sub

ERRH:
4260    If blnInTrans Then
4270      cnxn.RollbackTrans
4280      blnInTrans = False
4290      rsx3.Close
4300      rsx2.Close
4310      rsx1.Close
4320      cnxn.Close
4330    End If
4340    Select Case ERR.Number
        Case Else
4350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4360    End Select
4370    Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4420    intShiftDown = (Shift And acShiftMask) > 0
4430    intAltDown = (Shift And acAltMask) > 0
4440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4460      Select Case intRetVal
          Case vbKeyTab
4470        With Me
4480          intRetVal = 0
4490          .cmdClose.SetFocus
4500        End With
4510      End Select
4520    End If

        ' ** Shift keys.
4530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4540      Select Case intRetVal
          Case vbKeyTab
4550        With Me
4560          intRetVal = 0
4570          If .chkIncludeCurrency.Enabled = True Then
4580            .chkIncludeCurrency.SetFocus
4590          ElseIf .chkIncludeArchive.Enabled = True Then
4600            .chkIncludeArchive.SetFocus
4610          Else
4620            .frmAccountComments_Sub.SetFocus
4630            .frmAccountComments_Sub.Form.MoveRec acCmdRecordsGoToLast
4640            .frmAccountComments_Sub.Form.Comment.SetFocus
4650          End If
4660        End With
4670      End Select
4680    End If

EXITP:
4690    KeyCode = intRetVal
4700    Exit Sub

ERRH:
4710    intRetVal = 0
4720    Select Case ERR.Number
        Case Else
4730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4740    End Select
4750    Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

        Dim strFilter As String

4810    With Me
4820      If blnThisArchive = True Then
4830        Select Case .chkIncludeArchive
            Case True
4840          .chkIncludeArchive_lbl.FontBold = True
4850          .archive_trans_lbl.Visible = True
4860          .archive_trans_lbl_box.Visible = True
4870          strFilter = vbNullString
4880          .frmAccountComments_Sub.Form.Filter = strFilter
4890          .frmAccountComments_Sub.Form.FilterOn = False
4900        Case False
4910          .chkIncludeArchive_lbl.FontBold = False
4920          .archive_trans_lbl.Visible = False
4930          .archive_trans_lbl_box.Visible = False
4940          strFilter = "[src] = 'L'"
4950          .frmAccountComments_Sub.Form.Filter = strFilter
4960          .frmAccountComments_Sub.Form.FilterOn = True
4970        End Select
4980        .frmAccountComments_Sub.Form.Requery
4990        .frmAccountComments_Sub.SetFocus
5000      End If
5010    End With

EXITP:
5020    Exit Sub

ERRH:
5030    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
5040    Case Else
5050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5060    End Select
5070    Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab
5170        With Me
5180          intRetVal = 0
5190          If .chkIncludeCurrency.Enabled = True Then
5200            .chkIncludeCurrency.SetFocus
5210          ElseIf .cmdUpdate.Enabled = True Then
5220            .cmdUpdate.SetFocus
5230          Else
5240            .cmdClose.SetFocus
5250          End If
5260        End With
5270      End Select
5280    End If

        ' ** Shift keys.
5290    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5300      Select Case intRetVal
          Case vbKeyTab
5310        With Me
5320          intRetVal = 0
5330          .frmAccountComments_Sub.SetFocus
5340          .frmAccountComments_Sub.Form.MoveRec acCmdRecordsGoToLast
5350          .frmAccountComments_Sub.Form.Comment.SetFocus
5360        End With
5370      End Select
5380    End If

EXITP:
5390    KeyCode = intRetVal
5400    Exit Sub

ERRH:
5410    intRetVal = 0
5420    Select Case ERR.Number
        Case Else
5430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5440    End Select
5450    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

5510    With Me

5520      Select Case chkIncludeCurrency
          Case True
5530        .chkIncludeCurrency_lbl.FontBold = True
5540        .frmAccountComments_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountComments_Sub.
5550        IncludeCurrency True  ' ** Procedure: Below.
5560      Case False
5570        .chkIncludeCurrency_lbl.FontBold = False
5580        .frmAccountComments_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountComments_Sub.
5590        IncludeCurrency False  ' ** Procedure: Below.
5600      End Select
5610      .frmAccountComments_Sub.SetFocus

5620    End With

EXITP:
5630    Exit Sub

ERRH:
5640    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
5650    Case Else
5660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5670    End Select
5680    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5720    intShiftDown = (Shift And acShiftMask) > 0
5730    intAltDown = (Shift And acAltMask) > 0
5740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab
5770        With Me
5780          intRetVal = 0
5790          If .cmdUpdate.Enabled = True Then
5800            .cmdUpdate.SetFocus
5810          Else
5820            .cmdClose.SetFocus
5830          End If
5840        End With
5850      End Select
5860    End If

        ' ** Shift keys.
5870    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5880      Select Case intRetVal
          Case vbKeyTab
5890        With Me
5900          intRetVal = 0
5910          If .chkIncludeArchive.Enabled = True Then
5920            .chkIncludeArchive.SetFocus
5930          Else
5940            .frmAccountComments_Sub.SetFocus
5950            .frmAccountComments_Sub.Form.MoveRec acCmdRecordsGoToLast
5960            .frmAccountComments_Sub.Form.Comment.SetFocus
5970          End If
5980        End With
5990      End Select
6000    End If

EXITP:
6010    KeyCode = intRetVal
6020    Exit Sub

ERRH:
6030    intRetVal = 0
6040    Select Case ERR.Number
        Case Else
6050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6060    End Select
6070    Resume EXITP

End Sub

Private Sub SetArchiveOption()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "SetArchiveOption"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

6110    With Me

6120      blnHasArchive = False: blnThisArchive = False

6130      Set dbs = CurrentDb
6140      With dbs

6150        Set rst = .OpenRecordset("LedgerArchive", dbOpenDynaset, dbReadOnly)
6160        With rst
6170          If .BOF = True And .EOF = True Then
                ' ** No archive whatsoever!
6180          Else
6190            .MoveLast
6200            If .RecordCount > 1 Then
6210              blnHasArchive = True
6220            End If
6230          End If
6240          .Close
6250        End With
6260        Set rst = Nothing

6270        If blnHasArchive = True Then
              ' ** LedgerArchive, with 'Liability (+)'/'Liability (-)' broken out,
              ' ** with totdescx, by specified GlobalVarGet('gstrAccountNo').
6280          Set qdf = .QueryDefs("qryAccountProfile_Comments_07")
6290          Set rst = qdf.OpenRecordset
6300          With rst
6310            If .BOF = True And .EOF = True Then
                  ' ** This account has no archive.
6320            Else
6330              .MoveLast
6340              If .RecordCount > 0 Then
6350                blnThisArchive = True
6360              End If
6370            End If
6380            .Close
6390          End With
6400          Set rst = Nothing
6410          Set qdf = Nothing
6420        End If  ' ** blnHasArchive.

6430        .Close
6440      End With
6450      Set dbs = Nothing

6460      .archive_trans_lbl.Visible = False
6470      .archive_trans_lbl_box.Visible = False

6480      Select Case blnHasArchive
          Case True
6490        Select Case blnThisArchive
            Case True
6500          .chkIncludeArchive.Enabled = True
6510          .chkIncludeArchive_lbl.Visible = True
6520          .chkIncludeArchive_lbl4_dim.Visible = False
6530          .chkIncludeArchive_lbl4_dim_hi.Visible = False
6540          .chkIncludeArchive_lbl5_dim.Visible = False
6550          .chkIncludeArchive_lbl5_dim_hi.Visible = False
6560        Case False
6570          .chkIncludeArchive.Enabled = False
6580          .chkIncludeArchive_lbl.Visible = False
6590          .chkIncludeArchive_lbl4_dim.Visible = True
6600          .chkIncludeArchive_lbl4_dim_hi.Visible = True
6610          .chkIncludeArchive_lbl5_dim.Visible = True
6620          .chkIncludeArchive_lbl5_dim_hi.Visible = True
6630        End Select
6640        .chkIncludeArchive_lbl2_dim.Visible = False
6650        .chkIncludeArchive_lbl2_dim_hi.Visible = False
6660        .chkIncludeArchive_lbl3_dim.Visible = False
6670        .chkIncludeArchive_lbl3_dim_hi.Visible = False
6680      Case False
6690        .chkIncludeArchive.Enabled = False
6700        .chkIncludeArchive_lbl.Visible = False
6710        .chkIncludeArchive_lbl4_dim.Visible = False
6720        .chkIncludeArchive_lbl4_dim_hi.Visible = False
6730        .chkIncludeArchive_lbl5_dim.Visible = False
6740        .chkIncludeArchive_lbl5_dim_hi.Visible = False
6750        .chkIncludeArchive_lbl2_dim.Visible = True
6760        .chkIncludeArchive_lbl2_dim_hi.Visible = True
6770        .chkIncludeArchive_lbl3_dim.Visible = True
6780        .chkIncludeArchive_lbl3_dim_hi.Visible = True
6790      End Select

6800    End With

EXITP:
6810    Set rst = Nothing
6820    Set qdf = Nothing
6830    Set dbs = Nothing
6840    Exit Sub

ERRH:
6850    Select Case ERR.Number
        Case Else
6860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6870    End Select
6880    Resume EXITP

End Sub

Private Sub IncludeCurrency(blnInclude As Boolean)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long

6910    With Me
6920      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6930        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6940      End If
6950      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
6960      lngMonitorNum = 1&: lngTmp01 = 0&
6970      EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
6980      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
6990      Select Case .chkIncludeCurrency
          Case True
7000        DoCmd.SelectObject acForm, THIS_NAME, False
7010        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
7020        DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, lngFrm_Height  'lngFrm_Top
7030        If lngMonitorNum > 1& Then
7040          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
7050        End If
7060        .Width = lngForm_Width
7070        .cmdClose.Left = lngClose_Left
7080        .cmdUpdate.Left = lngUpdate_Left
7090        .frmAccountComments_Sub.Width = lngSub_Width
7100        .frmAccountComments_Sub_box.Width = (.frmAccountComments_Sub.Width + (2& * lngTpp))
7110        .Nav_box01.Width = .frmAccountComments_Sub.Width
7120        .Header_vline01.Left = .Width
7130        .Header_vline02.Left = .Width
7140        .Detail_vline01.Left = .Width
7150        .Detail_vline02.Left = .Width
7160        .Footer_vline01.Left = .Width
7170        .Footer_vline02.Left = .Width
7180        .Header_hline01.Width = .Width
7190        .Header_hline02.Width = .Width
7200        .Detail_hline01.Width = .Width
7210        .Detail_hline02.Width = .Width
7220        .Footer_hline01.Width = .Width
7230        .Footer_hline02.Width = .Width
7240        .Repaint
7250        DoEvents
7260      Case False
7270        .cmdClose.Left = ((lngClose_Left - lngCurrID_Width) + lngTpp)
7280        .cmdUpdate.Left = ((lngUpdate_Left - lngCurrID_Width) + lngTpp)
7290        .frmAccountComments_Sub.Width = ((lngSub_Width - lngCurrID_Width) + lngTpp)
7300        .frmAccountComments_Sub_box.Width = (.frmAccountComments_Sub.Width + (2& * lngTpp))
7310        .Nav_box01.Width = .frmAccountComments_Sub.Width
7320        .Header_vline01.Left = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7330        .Header_vline02.Left = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7340        .Detail_vline01.Left = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7350        .Detail_vline02.Left = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7360        .Footer_vline01.Left = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7370        .Footer_vline02.Left = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7380        .Header_hline01.Width = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7390        .Header_hline02.Width = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7400        .Detail_hline01.Width = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7410        .Detail_hline02.Width = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7420        .Footer_hline01.Width = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7430        .Footer_hline02.Width = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7440        .Width = ((lngForm_Width - lngCurrID_Width) + lngTpp)
7450        DoCmd.SelectObject acForm, THIS_NAME, False
7460        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
7470        DoCmd.MoveSize (lngFrm_Left + (lngCurrID_Width / 2&)), lngTmp01, ((lngFrm_Width - lngCurrID_Width) + lngTpp), lngFrm_Height  'lngFrm_Top
7480        If lngMonitorNum > 1& Then
7490          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
7500        End If
7510      End Select
7520    End With

EXITP:
7530    Exit Sub

ERRH:
7540    Select Case ERR.Number
        Case Else
7550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7560    End Select
7570    Resume EXITP

End Sub
