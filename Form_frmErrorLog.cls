VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmErrorLog"

'VGC 04/17/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Excel:            E {cmdExcel}
' **   Print:            P {cmdPrint}
' **   Clear All:        R {cmdClear}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}
' **   1 Line:           1 {opgMsgLines_opt1}
' **   2 Lines:          2 {opgMsgLines_opt2}
' **   3 Lines:          3 {opgMsgLines_opt3}
' **   4 Lines:          4 {opgMsgLines_opt4}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

Private lngForm_Width As Long, lngDetailHLine01_Top As Long, lngPageOf_Top As Long, lngMsgLines_Top As Long
Private lngMsg_Height As Long, lngDetail_Height As Long, lngPix_Height As Long, lngSizable_Offset As Long
Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngSub_Left As Long, lngSub_Top As Long, lngSub_Width As Long, lngSub_Height As Long
Private lngClose_Offset As Long, lngClear_Offset As Long, lngLine_Offset As Long

Private blnIsOpen As Boolean
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnIsOpen = True
160           DoCmd.Restore

170           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           .Ver = AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.
200           .Rel = CDate(AppDate_Get(True))  ' ** Module Function: modAppVersionFuncs.

210           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

              ' ** Variables are fed empty, then populated ByRef.
220           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

230           .form_width_line.Left = 0&  ' ** Just get these out of the way.
240           .detail_height_line.Top = 0&

250           lngForm_Width = .Width
260           lngClose_Offset = (lngFrm_Width - .cmdClose.Left)
270           lngClear_Offset = (lngFrm_Width - .cmdClear.Left)
280           lngLine_Offset = (.Header_vline01.Left - .cmdClose.Left)
290           lngSizable_Offset = (.Sizable_lbl1.Left - .cmdClose.Left)
300           lngSub_Left = .frmErrorLog_Sub.Left
310           lngSub_Top = .frmErrorLog_Sub.Top
320           lngSub_Width = .frmErrorLog_Sub.Width
330           lngSub_Height = .frmErrorLog_Sub.Height
340           lngDetailHLine01_Top = .Detail_hline01.Top
350           lngPageOf_Top = .chkPageOf.Top
360           lngMsgLines_Top = .opgMsgLines.Top

370           lngMsg_Height = .frmErrorLog_Sub.Form.ErrLog_Message_Disp.Height  ' ** 4 lines.
380           lngDetail_Height = .frmErrorLog_Sub.Form.Detail.Height
390           lngPix_Height = (lngMsg_Height / 64&)

400           lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
410           If lngRecsCur > 0& Then
420             .cmdPreview.Enabled = True
430             .cmdPrint.Enabled = True
      #If NoExcel Then
440             .cmdExcel.Enabled = False
      #Else
450             .cmdExcel.Enabled = True
      #End If
460             .cmdClear.Enabled = True
470           Else
480             .cmdPreview.Enabled = False
490             .cmdPrint.Enabled = False
500             .cmdExcel.Enabled = False
510             .cmdClear.Enabled = False
520           End If

530           opgMsgLines_AfterUpdate  ' ** Procedure: Below.
540           chkPageOf_AfterUpdate  ' ** Procedure: Below.

550         Else
560           Cancel = -1
570         End If
580       Else
590         Cancel = -1
600       End If

610     End With

620     DoCmd.Hourglass False

630     If Cancel = -1 Then
640       Beep
650       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
660     End If

670     blnIsOpen = False

EXITP:
680     Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
720     End Select
730     Resume EXITP

End Sub

Private Sub Form_Load()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRptPath As String
        Dim lngRecs As Long

810     With Me
820       If IsNull(.UserReportPath) = True Then
            ' ** if the user has a UserReportPath preference from another screen, use that.
830         Set dbs = CurrentDb
840         With dbs
              ' ** qryReport_Path_02 (qryReport_Path_01 (tblPreference_Control, just 'UserReportPath'),
              ' ** linked to tblPreference_User, by specified [usr]), grouped by prefuser_text.
850           Set qdf = .QueryDefs("qryReport_Path_03")  '##dbs_id
860           With qdf.Parameters
870             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
880           End With
890           Set rst = qdf.OpenRecordset
900           With rst
910             If .BOF = True And .EOF = True Then
                  ' ** Nope.
920             Else
930               .MoveLast
940               lngRecs = .RecordCount
950               .MoveFirst
960               If lngRecs = 1& Then
                    ' ** If they only have one, or all of them are the same, use it!
970                 strRptPath = ![prefuser_text]
980               End If
990             End If
1000            .Close
1010          End With
1020          .Close
1030        End With
1040        If strRptPath <> vbNullString Then
1050          .UserReportPath = strRptPath
1060        End If
1070      End If
1080    End With

EXITP:
1090    Set rst = Nothing
1100    Set qdf = Nothing
1110    Set dbs = Nothing
1120    Exit Sub

ERRH:
1130    DoCmd.Hourglass False
1140    Select Case ERR.Number
        Case Else
1150      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1160    End Select
1170    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyEscape
1270        intRetVal = 0
1280        cmdClose_Click  ' ** Procedure: Below.
1290      End Select
1300    End If

        ' ** Alt keys.
1310    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1320      Select Case intRetVal
          Case vbKeyX
1330        intRetVal = 0
1340        cmdClose_Click  ' ** Procedure: Below.
1350      End Select
1360    End If

        ' ** Ctrl keys.
1370    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1380      Select Case intRetVal
          Case vbKeyG
1390        With Me
1400          intRetVal = 0
1410          If .chkPageOf.Enabled = True Then
1420            .chkPageOf.SetFocus
1430            .chkPageOf = (Not .chkPageOf)
1440            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1450          End If
1460        End With
1470      End Select
1480    End If

        ' ** Ctrl-Shift keys.
1490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1500      Select Case intRetVal
          Case vbKeyF
1510        With Me
1520          intRetVal = 0
1530          .FocusHolder.SetFocus
1540        End With
1550      End Select
1560    End If

EXITP:
1570    KeyCode = intRetVal
1580    Exit Sub

ERRH:
1590    intRetVal = 0
1600    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1610    Case Else
1620      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1630    End Select
1640    Resume EXITP

End Sub

Private Sub Form_Resize()
' ** This is a little messy, and not very generic!

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Resize"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngSubHeight As Long, lngDiff As Long
        Dim lngTmp01 As Long
        Dim lngX As Long

1710    With Me

1720      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1730        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1740      End If

1750      If blnIsOpen = False Then

1760        GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.

1770        If lngWidth <> lngFrm_Width Or lngHeight <> lngFrm_Height Then

1780          If lngWidth > lngFrm_Width Then
1790            lngDiff = (lngWidth - lngFrm_Width)
1800            .Width = (lngForm_Width + lngDiff)
1810            .cmdClose.Left = (lngWidth - lngClose_Offset)
1820            .cmdClear.Left = (lngWidth - lngClear_Offset)
1830            .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
1840            .Sizable_lbl2.Left = (.cmdClose.Left + lngSizable_Offset)
1850            .frmErrorLog_Sub.Width = (lngSub_Width + lngDiff)
1860            .frmErrorLog_Sub_box.Width = (.frmErrorLog_Sub.Width + (2& * lngTpp))
1870            .Nav_box01.Width = .frmErrorLog_Sub.Width
1880            .Header_vline01.Left = (.cmdClose.Left + lngLine_Offset)
1890            .Header_vline02.Left = .Header_vline01.Left
1900            .Footer_vline01.Left = (.cmdClose.Left + lngLine_Offset)
1910            .Footer_vline02.Left = .Footer_vline01.Left
1920            .Header_hline01.Width = (.cmdClose.Left + lngLine_Offset)
1930            .Header_hline02.Width = .Header_hline01.Width
1940            .Footer_hline01.Width = (.cmdClose.Left + lngLine_Offset)
1950            .Footer_hline02.Width = .Footer_hline01.Width
1960            .Detail_vline01.Left = .Header_vline01.Left
1970            .Detail_vline02.Left = .Detail_vline01.Left
1980            .Detail_hline01.Width = .Detail_vline01.Left
1990            .Detail_hline02.Width = .Detail_hline01.Width
2000          Else
                'lngTmp01 = ((lngClear_Offset - lngClose_Offset) - .cmdClear.Width)
                'lngTmp02 = (.opgMsgLines_box.Left + .opgMsgLines_box.Width)
                'lngTmp03 = (lngTmp02 + lngTmp01)
2010            lngDiff = (lngFrm_Width - lngWidth)
2020            If (lngWidth - lngClose_Offset) >= 9750& Then 'lngTmp03 Then  ' ** I can't seem to make sense of the
2030              .cmdClose.Left = (lngWidth - lngClose_Offset)               ' ** real values vs. my calculated ones!
2040              .cmdClear.Left = (lngWidth - lngClear_Offset)
2050            End If
2060            .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
2070            .Sizable_lbl2.Left = (.cmdClose.Left + lngSizable_Offset)
2080            .frmErrorLog_Sub.Width = (lngSub_Width - lngDiff)
2090            .frmErrorLog_Sub_box.Width = (.frmErrorLog_Sub.Width + (2& * lngTpp))
2100            .Nav_box01.Width = .frmErrorLog_Sub.Width
2110            .Header_vline01.Left = (.cmdClose.Left + lngLine_Offset)
2120            .Header_vline02.Left = .Header_vline01.Left
2130            .Footer_vline01.Left = (.cmdClose.Left + lngLine_Offset)
2140            .Footer_vline02.Left = .Footer_vline01.Left
2150            .Header_hline01.Width = (.cmdClose.Left + lngLine_Offset)
2160            .Header_hline02.Width = .Header_hline01.Width
2170            .Footer_hline01.Width = (.cmdClose.Left + lngLine_Offset)
2180            .Footer_hline02.Width = .Footer_hline01.Width
2190            .Detail_vline01.Left = .Header_vline01.Left
2200            .Detail_vline02.Left = .Detail_vline01.Left
2210            .Detail_hline01.Width = .Detail_vline01.Left
2220            .Detail_hline02.Width = .Detail_hline01.Width
2230          End If

2240          lngSubHeight = .frmErrorLog_Sub.Height
2250          If lngHeight > lngFrm_Height Then
2260            lngDiff = (lngHeight - lngFrm_Height)
2270            .Detail.Height = (lngDetail_Height + lngDiff)
2280            .frmErrorLog_Sub.Height = (lngSub_Height + lngDiff)
2290            .Detail.Height = .Detail.Height + 300&  ' ** Without this, it stops working! Why?
2300            .frmErrorLog_Sub_box.Height = (.frmErrorLog_Sub.Height + (6& * lngTpp))
2310            .ShortcutMenu_lbl.Top = ((.frmErrorLog_Sub_box.Top + .frmErrorLog_Sub_box.Height) + lngTpp)
2320            .ShortcutMenu_up_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
2330            .ShortcutMenu_down_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
2340            .Nav_vline01.Top = (.frmErrorLog_Sub.Top + .frmErrorLog_Sub.Height)
2350            .Nav_vline02.Top = .Nav_vline01.Top
2360            .Nav_vline03.Top = .Nav_vline01.Top
2370            .Nav_vline04.Top = .Nav_vline01.Top
2380            .Nav_hline01.Top = .Nav_vline01.Top
2390            .Nav_hline02.Top = (.Nav_vline01.Top + lngTpp)
2400            .Nav_box01.Top = (.frmErrorLog_Sub.Top + .frmErrorLog_Sub.Height)
2410            .Detail_vline01.Top = (lngDetailHLine01_Top + lngDiff)
2420            .Detail_vline02.Top = .Detail_vline01.Top
2430            .Detail_hline01.Top = .Detail_vline01.Top
2440            .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
2450            .chkPageOf.Top = (lngPageOf_Top + lngDiff)
2460            .chkPageOf_lbl.Top = (.chkPageOf.Top - (2& * lngTpp))
2470            .chkPageOf_box.Top = (.chkPageOf_lbl.Top - lngTpp)
2480            lngTmp01 = ((lngMsgLines_Top + lngDiff) - .opgMsgLines.Top)
2490            lngTmp01 = (lngTmp01 / lngTpp)  ' ** Intentionally letting it round.
2500            For lngX = 1& To lngTmp01
2510              .opgMsgLines.Top = (.opgMsgLines.Top + lngTpp)
2520              .opgMsgLines_lbl.Top = (.opgMsgLines_lbl.Top + lngTpp)
2530              .opgMsgLines_box.Top = (.opgMsgLines_box.Top + lngTpp)
2540              .opgMsgLines_box2.Top = (.opgMsgLines_box2.Top + lngTpp)
2550              .opgMsgLines_opt_box.Top = (.opgMsgLines_opt_box.Top + lngTpp)
2560              .opgMsgLines_opt1.Top = (.opgMsgLines_opt1.Top + lngTpp)
2570              .opgMsgLines_opt1_lbl.Top = (.opgMsgLines_opt1_lbl.Top + lngTpp)
2580              .opgMsgLines_opt2.Top = (.opgMsgLines_opt2.Top + lngTpp)
2590              .opgMsgLines_opt2_lbl.Top = (.opgMsgLines_opt2_lbl.Top + lngTpp)
2600              .opgMsgLines_opt3.Top = (.opgMsgLines_opt3.Top + lngTpp)
2610              .opgMsgLines_opt3_lbl.Top = (.opgMsgLines_opt3_lbl.Top + lngTpp)
2620              .opgMsgLines_opt4.Top = (.opgMsgLines_opt4.Top + lngTpp)
2630              .opgMsgLines_opt4_lbl.Top = (.opgMsgLines_opt4_lbl.Top + lngTpp)
2640            Next
                ' ** Just to make sure it's where it belongs.
2650            .opgMsgLines.Top = (lngMsgLines_Top + lngDiff)
2660            .opgMsgLines_box.Top = .opgMsgLines.Top
2670            .opgMsgLines_box2.Top = (.opgMsgLines_box.Top - (4& * lngTpp))
2680            .opgMsgLines_opt_box.Top = (.opgMsgLines.Top + lngTpp)
2690            .opgMsgLines_lbl.Top = (.opgMsgLines_opt_box.Top + (2& * lngTpp))
2700            .opgMsgLines_opt1_lbl.Top = .opgMsgLines_lbl.Top
2710            .opgMsgLines_opt1.Top = (.opgMsgLines_opt1_lbl.Top + lngTpp)
2720            .opgMsgLines_opt2_lbl.Top = .opgMsgLines_lbl.Top
2730            .opgMsgLines_opt2.Top = .opgMsgLines_opt1.Top
2740            .opgMsgLines_opt3_lbl.Top = .opgMsgLines_lbl.Top
2750            .opgMsgLines_opt3.Top = .opgMsgLines_opt1.Top
2760            .opgMsgLines_opt4_lbl.Top = .opgMsgLines_lbl.Top
2770            .opgMsgLines_opt4.Top = .opgMsgLines_opt1.Top
2780          Else
2790            If (lngSub_Height - (lngFrm_Height - lngHeight)) >= 1680& Then
2800              lngDiff = (lngFrm_Height - lngHeight)
2810              .frmErrorLog_Sub.Height = (lngSub_Height - lngDiff)
2820              .frmErrorLog_Sub_box.Height = (.frmErrorLog_Sub.Height + (6& * lngTpp))
2830              .ShortcutMenu_lbl.Top = ((.frmErrorLog_Sub_box.Top + .frmErrorLog_Sub_box.Height) + lngTpp)
2840              .ShortcutMenu_up_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
2850              .ShortcutMenu_down_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
2860              .Nav_vline01.Top = (.frmErrorLog_Sub.Top + .frmErrorLog_Sub.Height)
2870              .Nav_vline02.Top = .Nav_vline01.Top
2880              .Nav_vline03.Top = .Nav_vline01.Top
2890              .Nav_vline04.Top = .Nav_vline01.Top
2900              .Nav_hline01.Top = .Nav_vline01.Top
2910              .Nav_hline02.Top = (.Nav_vline01.Top + lngTpp)
2920              .Nav_box01.Top = (.frmErrorLog_Sub.Top + .frmErrorLog_Sub.Height)
2930              .Detail_vline01.Top = (lngDetailHLine01_Top - lngDiff)
2940              .Detail_vline02.Top = .Detail_vline01.Top
2950              .Detail_hline01.Top = .Detail_vline01.Top
2960              .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
2970              .chkPageOf.Top = (lngPageOf_Top - lngDiff)
2980              .chkPageOf_lbl.Top = (.chkPageOf.Top - (2& * lngTpp))
2990              .chkPageOf_box.Top = (.chkPageOf_lbl.Top - lngTpp)
3000              lngTmp01 = (.opgMsgLines.Top - (lngMsgLines_Top - lngDiff))
3010              lngTmp01 = (lngTmp01 / lngTpp)  ' ** Intentionally letting it round.
3020              For lngX = 1& To lngTmp01
3030                .opgMsgLines.Top = (.opgMsgLines.Top - lngTpp)
3040                .opgMsgLines_lbl.Top = (.opgMsgLines_lbl.Top - lngTpp)
3050                .opgMsgLines_box.Top = (.opgMsgLines_box.Top - lngTpp)
3060                .opgMsgLines_box2.Top = (.opgMsgLines_box2.Top - lngTpp)
3070                .opgMsgLines_opt_box.Top = (.opgMsgLines_opt_box.Top - lngTpp)
3080                .opgMsgLines_opt1.Top = (.opgMsgLines_opt1.Top - lngTpp)
3090                .opgMsgLines_opt1_lbl.Top = (.opgMsgLines_opt1_lbl.Top - lngTpp)
3100                .opgMsgLines_opt2.Top = (.opgMsgLines_opt2.Top - lngTpp)
3110                .opgMsgLines_opt2_lbl.Top = (.opgMsgLines_opt2_lbl.Top - lngTpp)
3120                .opgMsgLines_opt3.Top = (.opgMsgLines_opt3.Top - lngTpp)
3130                .opgMsgLines_opt3_lbl.Top = (.opgMsgLines_opt3_lbl.Top - lngTpp)
3140                .opgMsgLines_opt4.Top = (.opgMsgLines_opt4.Top - lngTpp)
3150                .opgMsgLines_opt4_lbl.Top = (.opgMsgLines_opt4_lbl.Top - lngTpp)
3160              Next
                  ' ** Just to make sure it's where it belongs.
3170              .opgMsgLines.Top = (lngMsgLines_Top - lngDiff)
3180              .opgMsgLines_box.Top = .opgMsgLines.Top
3190              .opgMsgLines_box2.Top = (.opgMsgLines_box.Top - (4& * lngTpp))
3200              .opgMsgLines_opt_box.Top = (.opgMsgLines.Top + lngTpp)
3210              .opgMsgLines_lbl.Top = (.opgMsgLines_opt_box.Top + (2& * lngTpp))
3220              .opgMsgLines_opt1_lbl.Top = .opgMsgLines_lbl.Top
3230              .opgMsgLines_opt1.Top = (.opgMsgLines_opt1_lbl.Top + lngTpp)
3240              .opgMsgLines_opt2_lbl.Top = .opgMsgLines_lbl.Top
3250              .opgMsgLines_opt2.Top = .opgMsgLines_opt1.Top
3260              .opgMsgLines_opt3_lbl.Top = .opgMsgLines_lbl.Top
3270              .opgMsgLines_opt3.Top = .opgMsgLines_opt1.Top
3280              .opgMsgLines_opt4_lbl.Top = .opgMsgLines_lbl.Top
3290              .opgMsgLines_opt4.Top = .opgMsgLines_opt1.Top
3300            End If
3310          End If

3320        End If

3330      End If

3340    End With

EXITP:
3350    Exit Sub

ERRH:
3360    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Ignore.
3370    Case Else
3380      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3390    End Select
3400    Resume EXITP

End Sub

Private Sub Form_Timer()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

3510    With Me
3520      .TimerInterval = 0&
3530      CmdBars_Prop  ' ** Module Procedure: modWindowsFuncs
3540      If gblnSetFocus = True Then
3550        gblnSetFocus = False
3560        strFormName = vbNullString
3570  On Error Resume Next
3580        Do Until strFormName = THIS_NAME
3590          DoCmd.SelectObject acForm, THIS_NAME, False
3600          strFormName = Screen.ActiveForm.Name
3610        Loop
3620  On Error GoTo ERRH
3630        DoCmd.Hourglass False
3640      End If
3650      If gblnGoToReport = True Then
3660        gblnGoToReport = False
3670        If .cmdPrint.Enabled = True Then
3680          .cmdPrint.SetFocus
3690        Else
3700          .cmdClose.SetFocus
3710        End If
3720        Beep
3730        DoCmd.Hourglass False
3740        DoEvents
3750      End If
3760    End With

EXITP:
3770    Exit Sub

ERRH:
3780    DoCmd.Hourglass False
3790    Select Case ERR.Number
        Case Else
3800      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3810    End Select
3820    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String

3910    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
3920    Do While Reports.Count > 0
3930      DoCmd.Close acReport, Reports(0).Name
3940    Loop

3950    Select Case strCallingForm
        Case "frmMenu_Maintenance"
3960      gblnSetFocus = True
3970      gstrReturningForm = THIS_NAME
3980      If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
3990        DoCmd.SelectObject acForm, strCallingForm, False
4000        Forms(strCallingForm).TimerInterval = 250&
4010      Else
4020        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Utility"
4030      End If
4040    Case vbNullString
4050      gblnSetFocus = True
4060      gstrReturningForm = THIS_NAME
4070      strDocName = "frmMenu_Maintenance"
4080      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
4090        DoCmd.SelectObject acForm, strDocName, False
4100        Forms(strDocName).TimerInterval = 250&
4110      Else
4120        DoCmd.OpenForm strDocName, , , , , , "frmMenu_Utility"
4130      End If
4140    Case Else
4150      DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
4160    End Select

EXITP:
4170    Exit Sub

ERRH:
4180    Select Case ERR.Number
        Case Else
4190      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4200    End Select
4210    Resume EXITP

End Sub

Public Sub cmdClose_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4310    DoCmd.Close acForm, THIS_NAME

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4350    End Select
4360    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
4490          lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
4500          If lngRecsCur > 0& Then
4510            .frmErrorLog_Sub.SetFocus
4520            .frmErrorLog_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmErrorLog_Sub.
4530            .frmErrorLog_Sub.Form.ErrLog_Date.SetFocus
4540          ElseIf .opgMsgLines.Enabled = True Then
4550            .opgMsgLines.SetFocus
4560          Else
4570            .chkPageOf.SetFocus
4580          End If
4590        End With
4600      End Select
4610    End If

        ' ** Shift keys.
4620    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4630      Select Case intRetVal
          Case vbKeyTab
4640        With Me
4650          intRetVal = 0
4660          If .cmdClear.Enabled = True Then
4670            .cmdClear.SetFocus
4680          Else
4690            .chkPageOf.SetFocus
4700          End If
4710        End With
4720      End Select
4730    End If

EXITP:
4740    KeyCode = intRetVal
4750    Exit Sub

ERRH:
4760    intRetVal = 0
4770    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4780    Case Else
4790      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4800    End Select
4810    Resume EXITP

End Sub

Private Sub cmdClear_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClear_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

4910    With Me
4920      lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
4930      If lngRecsCur > 0& Then
4940        Beep
4950        If MsgBox("Are you sure you want to clear out the Error Log?", _
                vbQuestion + vbYesNo + vbDefaultButton2, "Clear Error Log") = vbYes Then
4960          Set dbs = CurrentDb
4970          With dbs
                ' ** Empty tblErrorLog.
4980            Set qdf = .QueryDefs("qryErrLog_05")
4990            qdf.Execute
5000            .Close
5010          End With
5020          .frmErrorLog_Sub.Form.Requery
5030          .cmdClose.SetFocus
5040          DoEvents
5050          .cmdClear.Enabled = False
5060          .cmdPreview.Enabled = False
5070          .cmdPrint.Enabled = False
      #If NoExcel Then
5080          .cmdExcel.Enabled = False
      #Else
5090          .cmdExcel.Enabled = False
      #End If
5100        End If
5110      Else
5120        Beep
5130      End If
5140    End With

EXITP:
5150    Set qdf = Nothing
5160    Set dbs = Nothing
5170    Exit Sub

ERRH:
5180    Select Case ERR.Number
        Case Else
5190      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
5200    End Select
5210    Resume EXITP

End Sub

Private Sub cmdClear_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClear_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5320    intShiftDown = (Shift And acShiftMask) > 0
5330    intAltDown = (Shift And acAltMask) > 0
5340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab
5370        With Me
5380          intRetVal = 0
5390          .cmdClose.SetFocus
5400        End With
5410      End Select
5420    End If

        ' ** Shift keys.
5430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5440      Select Case intRetVal
          Case vbKeyTab
5450        With Me
5460          intRetVal = 0
5470          .chkPageOf.SetFocus
5480        End With
5490      End Select
5500    End If

EXITP:
5510    KeyCode = intRetVal
5520    Exit Sub

ERRH:
5530    intRetVal = 0
5540    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5550    Case Else
5560      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
5570    End Select
5580    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

5610    With Me
5620      lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
5630      If lngRecsCur > 0& Then
5640        strDocName = "rptErrorLog"
5650        gstrReportCallingForm = THIS_NAME  ' ** When not vbNullString, sets this form .Visible = False.
5660        DoCmd.OpenReport strDocName, acViewPreview
5670        DoCmd.Maximize
5680        DoCmd.RunCommand acCmdFitToWindow
5690      Else
5700        Beep
5710      End If
5720    End With

EXITP:
5730    Exit Sub

ERRH:
5740    If Forms(THIS_NAME).Visible = False Then
5750      Forms(THIS_NAME).Visible = True
5760    End If
5770    Select Case ERR.Number
        Case Else
5780      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
5790    End Select
5800    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5920    intShiftDown = (Shift And acShiftMask) > 0
5930    intAltDown = (Shift And acAltMask) > 0
5940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5960      Select Case intRetVal
          Case vbKeyTab
5970        With Me
5980          intRetVal = 0
5990          .cmdPrint.SetFocus
6000        End With
6010      End Select
6020    End If

        ' ** Shift keys.
6030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6040      Select Case intRetVal
          Case vbKeyTab
6050        With Me
6060          intRetVal = 0
6070          lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
6080          If lngRecsCur > 0& Then
6090            .frmErrorLog_Sub.SetFocus
6100            .frmErrorLog_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmErrorLog_Sub.
6110            .frmErrorLog_Sub.Form.ErrLog_FuncSub.SetFocus
6120          Else
6130            .cmdClose.SetFocus
6140          End If
6150        End With
6160      End Select
6170    End If

EXITP:
6180    KeyCode = intRetVal
6190    Exit Sub

ERRH:
6200    intRetVal = 0
6210    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6220    Case Else
6230      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
6240    End Select
6250    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

6310    With Me
6320      lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
6330      If lngRecsCur > 0& Then
6340        strDocName = "rptErrorLog"
6350        gstrReportCallingForm = vbNullString  ' ** When vbNullString, leaves this form .Visible = True.
6360        If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
6370          DoCmd.OpenReport strDocName, acViewPreview
6380          DoCmd.Maximize
6390          DoCmd.RunCommand acCmdFitToWindow
6400        Else
              '##GTR_Ref: rptErrorLog
6410          DoCmd.OpenReport strDocName, acViewNormal
6420        End If
6430      Else
6440        Beep
6450      End If
6460    End With

EXITP:
6470    Exit Sub

ERRH:
6480    If Forms(THIS_NAME).Visible = False Then
6490      Forms(THIS_NAME).Visible = True
6500    End If
6510    Select Case ERR.Number
        Case Else
6520      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
6530    End Select
6540    Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6620    intShiftDown = (Shift And acShiftMask) > 0
6630    intAltDown = (Shift And acAltMask) > 0
6640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6660      Select Case intRetVal
          Case vbKeyTab
6670        With Me
6680          intRetVal = 0
6690          If .cmdExcel.Enabled = True Then
6700            .cmdExcel.SetFocus
6710          Else
6720            .opgMsgLines.SetFocus
6730          End If
6740        End With
6750      End Select
6760    End If

        ' ** Shift keys.
6770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6780      Select Case intRetVal
          Case vbKeyTab
6790        With Me
6800          intRetVal = 0
6810          .cmdPreview.SetFocus
6820        End With
6830      End Select
6840    End If

EXITP:
6850    KeyCode = intRetVal
6860    Exit Sub

ERRH:
6870    intRetVal = 0
6880    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6890    Case Else
6900      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
6910    End Select
6920    Resume EXITP

End Sub

Private Sub cmdExcel_Click()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

7010    With Me
7020      lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
7030      If lngRecsCur > 0& Then

7040        If IsNull(.UserReportPath) = True Then
7050          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
7060        Else
7070          strRptPath = .UserReportPath
7080        End If
7090        strRptCap = "rptErrorLog_" & Format(Date, "yyyymmdd")

7100        strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

7110        If strRptPathFile <> vbNullString Then
7120          strQry = "qryErrLog_11"
7130          DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
7140          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
7150        End If

7160      Else
7170        Beep
7180      End If
7190    End With

EXITP:
7200    Exit Sub

ERRH:
7210    Select Case ERR.Number
        Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
7220      Beep
7230      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
7240    Case Else
7250      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
7260    End Select
7270    Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7320    intShiftDown = (Shift And acShiftMask) > 0
7330    intAltDown = (Shift And acAltMask) > 0
7340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7360      Select Case intRetVal
          Case vbKeyTab
7370        With Me
7380          intRetVal = 0
7390          .opgMsgLines.SetFocus
7400        End With
7410      End Select
7420    End If

        ' ** Shift keys.
7430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7440      Select Case intRetVal
          Case vbKeyTab
7450        With Me
7460          intRetVal = 0
7470          .cmdPrint.SetFocus
7480        End With
7490      End Select
7500    End If

EXITP:
7510    KeyCode = intRetVal
7520    Exit Sub

ERRH:
7530    intRetVal = 0
7540    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7550    Case Else
7560      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
7570    End Select
7580    Resume EXITP

End Sub

Private Sub opgMsgLines_AfterUpdate()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMsgLines_AfterUpdate"

        Dim frm As Access.Form
        Dim lngHeight As Long

7610    With Me
7620      .opgMsgLines_opt_box.Visible = False
7630      .opgMsgLines_opt_box.Top = (.opgMsgLines_opt1_lbl.Top - (2& * lngTpp))
7640      .opgMsgLines_opt1_lbl.FontBold = False
7650      .opgMsgLines_opt2_lbl.FontBold = False
7660      .opgMsgLines_opt3_lbl.FontBold = False
7670      .opgMsgLines_opt4_lbl.FontBold = False
7680      Set frm = .frmErrorLog_Sub.Form
7690      Select Case .opgMsgLines
          Case .opgMsgLines_opt1.OptionValue
7700        .opgMsgLines_opt1_lbl.FontBold = True
7710        lngHeight = ((lngPix_Height * 15&) + (lngPix_Height * 4&))
7720        frm.ErrLog_Message_Disp.Height = lngHeight  ' 285 '1 line
7730        frm.Detail.Height = lngHeight + (lngDetail_Height - lngMsg_Height)
7740        .opgMsgLines_opt_box.Left = (.opgMsgLines_opt1.Left - (3& * lngTpp))
7750        .opgMsgLines_opt_box.Visible = True
7760      Case .opgMsgLines_opt2.OptionValue
7770        .opgMsgLines_opt2_lbl.FontBold = True
7780        lngHeight = ((lngPix_Height * 30&) + (lngPix_Height * 4&))
7790        frm.ErrLog_Message_Disp.Height = lngHeight  ' 510 '2 lines
7800        frm.Detail.Height = lngHeight + (lngDetail_Height - lngMsg_Height)
7810        .opgMsgLines_opt_box.Left = (.opgMsgLines_opt2.Left - (3& * lngTpp))
7820        .opgMsgLines_opt_box.Visible = True
7830      Case .opgMsgLines_opt3.OptionValue
7840        .opgMsgLines_opt3_lbl.FontBold = True
7850        lngHeight = ((lngPix_Height * 45&) + (lngPix_Height * 4&))
7860        frm.ErrLog_Message_Disp.Height = lngHeight  ' 735 '3 lines
7870        frm.Detail.Height = lngHeight + (lngDetail_Height - lngMsg_Height)
7880        .opgMsgLines_opt_box.Left = (.opgMsgLines_opt3.Left - (3& * lngTpp))
7890        .opgMsgLines_opt_box.Visible = True
7900      Case .opgMsgLines_opt4.OptionValue
7910        .opgMsgLines_opt4_lbl.FontBold = True
7920        frm.ErrLog_Message_Disp.Height = lngMsg_Height  ' 960 '4 lines
7930        frm.Detail.Height = lngDetail_Height
7940        .opgMsgLines_opt_box.Left = (.opgMsgLines_opt4.Left - (3& * lngTpp))
7950        .opgMsgLines_opt_box.Visible = True
7960      End Select
7970    End With
7980    DoEvents

EXITP:
7990    Set frm = Nothing
8000    Exit Sub

ERRH:
8010    Select Case ERR.Number
        Case Else
8020      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
8030    End Select
8040    Resume EXITP

End Sub

Private Sub opgMsgLines_opt1_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMsgLines_opt1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8120    intShiftDown = (Shift And acShiftMask) > 0
8130    intAltDown = (Shift And acAltMask) > 0
8140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8170        With Me
8180          intRetVal = 0
8190          .chkPageOf.SetFocus
8200        End With
8210      End Select
8220    End If

        ' ** Shift keys.
8230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8250        With Me
8260          intRetVal = 0
8270          If .cmdExcel.Enabled = True Then
8280            .cmdExcel.SetFocus
8290          ElseIf .cmdPrint.Enabled = True Then
8300            .cmdPrint.SetFocus
8310          Else
8320            lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
8330            If lngRecsCur > 0& Then
8340              .frmErrorLog_Sub.SetFocus
8350              .frmErrorLog_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmErrorLog_Sub.
8360              .frmErrorLog_Sub.Form.ErrLog_FuncSub.SetFocus
8370            Else
8380              .cmdClose.SetFocus
8390            End If
8400          End If
8410        End With
8420      End Select
8430    End If

EXITP:
8440    KeyCode = intRetVal
8450    Exit Sub

ERRH:
8460    intRetVal = 0
8470    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8480    Case Else
8490      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
8500    End Select
8510    Resume EXITP

End Sub

Private Sub opgMsgLines_opt2_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMsgLines_opt2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8620    intShiftDown = (Shift And acShiftMask) > 0
8630    intAltDown = (Shift And acAltMask) > 0
8640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8670        With Me
8680          intRetVal = 0
8690          .chkPageOf.SetFocus
8700        End With
8710      End Select
8720    End If

        ' ** Shift keys.
8730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8750        With Me
8760          intRetVal = 0
8770          If .cmdExcel.Enabled = True Then
8780            .cmdExcel.SetFocus
8790          ElseIf .cmdPrint.Enabled = True Then
8800            .cmdPrint.SetFocus
8810          Else
8820            lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
8830            If lngRecsCur > 0& Then
8840              .frmErrorLog_Sub.SetFocus
8850              .frmErrorLog_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmErrorLog_Sub.
8860              .frmErrorLog_Sub.Form.ErrLog_FuncSub.SetFocus
8870            Else
8880              .cmdClose.SetFocus
8890            End If
8900          End If
8910        End With
8920      End Select
8930    End If

EXITP:
8940    KeyCode = intRetVal
8950    Exit Sub

ERRH:
8960    intRetVal = 0
8970    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8980    Case Else
8990      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
9000    End Select
9010    Resume EXITP

End Sub

Private Sub opgMsgLines_opt3_KeyDown(KeyCode As Integer, Shift As Integer)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMsgLines_opt3_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9120    intShiftDown = (Shift And acShiftMask) > 0
9130    intAltDown = (Shift And acAltMask) > 0
9140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9170        With Me
9180          intRetVal = 0
9190          .chkPageOf.SetFocus
9200        End With
9210      End Select
9220    End If

        ' ** Shift keys.
9230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9250        With Me
9260          intRetVal = 0
9270          If .cmdExcel.Enabled = True Then
9280            .cmdExcel.SetFocus
9290          ElseIf .cmdPrint.Enabled = True Then
9300            .cmdPrint.SetFocus
9310          Else
9320            lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
9330            If lngRecsCur > 0& Then
9340              .frmErrorLog_Sub.SetFocus
9350              .frmErrorLog_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmErrorLog_Sub.
9360              .frmErrorLog_Sub.Form.ErrLog_FuncSub.SetFocus
9370            Else
9380              .cmdClose.SetFocus
9390            End If
9400          End If
9410        End With
9420      End Select
9430    End If

EXITP:
9440    KeyCode = intRetVal
9450    Exit Sub

ERRH:
9460    intRetVal = 0
9470    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
9480    Case Else
9490      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
9500    End Select
9510    Resume EXITP

End Sub

Private Sub opgMsgLines_opt4_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMsgLines_opt4_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9670        With Me
9680          intRetVal = 0
9690          .chkPageOf.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9750        With Me
9760          intRetVal = 0
9770          If .cmdExcel.Enabled = True Then
9780            .cmdExcel.SetFocus
9790          ElseIf .cmdPrint.Enabled = True Then
9800            .cmdPrint.SetFocus
9810          Else
9820            lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
9830            If lngRecsCur > 0& Then
9840              .frmErrorLog_Sub.SetFocus
9850              .frmErrorLog_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmErrorLog_Sub.
9860              .frmErrorLog_Sub.Form.ErrLog_FuncSub.SetFocus
9870            Else
9880              .cmdClose.SetFocus
9890            End If
9900          End If
9910        End With
9920      End Select
9930    End If

EXITP:
9940    KeyCode = intRetVal
9950    Exit Sub

ERRH:
9960    intRetVal = 0
9970    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
9980    Case Else
9990      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
10000   End Select
10010   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

10110   With Me
10120     Select Case .chkPageOf
          Case True
10130       .chkPageOf_lbl.FontBold = True
10140     Case False
10150       .chkPageOf_lbl.FontBold = False
10160     End Select
10170   End With

EXITP:
10180   Exit Sub

ERRH:
10190   Select Case ERR.Number
        Case Else
10200     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
10210   End Select
10220   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10320   intShiftDown = (Shift And acShiftMask) > 0
10330   intAltDown = (Shift And acAltMask) > 0
10340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10370       With Me
10380         intRetVal = 0
10390         If .cmdClear.Enabled = True Then
10400           .cmdClear.SetFocus
10410         Else
10420           .cmdClose.SetFocus
10430         End If
10440       End With
10450     End Select
10460   End If

        ' ** Shift keys.
10470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10490       With Me
10500         intRetVal = 0
10510         If .opgMsgLines.Enabled = True Then
10520           .opgMsgLines.SetFocus
10530         ElseIf .cmdExcel.Enabled = True Then
10540           .cmdExcel.SetFocus
10550         ElseIf .cmdPrint.Enabled = True Then
10560           .cmdPrint.SetFocus
10570         Else
10580           lngRecsCur = .frmErrorLog_Sub.Form.RecCnt  ' ** Form Function: frmErrorLog_Sub.
10590           If lngRecsCur > 0& Then
10600             .frmErrorLog_Sub.SetFocus
10610             .frmErrorLog_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmErrorLog_Sub.
10620             .frmErrorLog_Sub.Form.ErrLog_FuncSub.SetFocus
10630           Else
10640             .cmdClose.SetFocus
10650           End If
10660         End If
10670       End With
10680     End Select
10690   End If

EXITP:
10700   KeyCode = intRetVal
10710   Exit Sub

ERRH:
10720   intRetVal = 0
10730   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
10740   Case Else
10750     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
10760   End Select
10770   Resume EXITP

End Sub
