VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptMasterBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptMasterBalance"

'VGC 03/10/2016: CHANGES!

Private Const strCallingForm As String = "frmMasterBalance"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         Select Case Forms(strCallingForm).chkPageOf
            Case True
150           .PageNumOf.Visible = True
160           .PageNum.Visible = False
170         Case False
180           .PageNum.Visible = True
190           .PageNumOf.Visible = False
200           .PageNumOf.ControlSource = vbNullString
210         End Select

220         .AsOf_lbl.Caption = "As of " & Format(Date, "mmmm dd, yyyy")
230         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

240         strOrderBy = Forms(strCallingForm).frmMasterBalance_Sub.Form.SortNow_Get  ' ** Form Function: frmMasterBalance_Sub.
            ' ** Potential sorts:
            ' **   [account_sort]
            ' **   [accountno]
            ' **   [shortname], [account_sort]
            ' **   [fix], [account_sort]
            ' **   [icash], [pcash], [cost], [account_sort]
            ' **   [icash_L], [pcash_L], [cost_L], [account_sort]
            ' **   [Ix], [account_sort]
            ' **   [pcash], [icash], [cost], [account_sort]
            ' **   [pcash_L], [icash_L], [cost_L], [account_sort]
            ' **   [Px], [account_sort]
            ' **   [cost], [icash], [pcash], [account_sort]
            ' **   [cost_L], [icash_L], [pcash_L], [account_sort]
            ' **   [Cx], [account_sort]
250         .OrderBy = strOrderBy
260         .OrderByOn = True

270         .Caption = "Master_Balance_" & Format(Date, "yymmdd")

280         Select Case Forms(strCallingForm).chkDiscrepancies
            Case True
290           .AsOf.Visible = True
300         Case False
310           .AsOf.Visible = False
320         End Select

330         Select Case Forms(strCallingForm).chkShowAcctMast
            Case True
              ' ** Leave as-is.
340         Case False
350           .ICash.Visible = False
360           .icash_lbl.Visible = False
370           .icash_lbl_line.Visible = False
380           .PCash.Visible = False
390           .pcash_lbl.Visible = False
400           .pcash_lbl_line.Visible = False
410           .Cost.Visible = False
420           .cost_lbl.Visible = False
430           .cost_lbl_line.Visible = False
440         End Select

450         Select Case Forms(strCallingForm).chkShowAcctSort
            Case True
              ' ** Not shown anyway.
460         Case False
              ' ** Not shown anyway.
470         End Select

480         If Forms(strCallingForm).frmMasterBalance_Sub.Form.RecsSum_Disc = 0 Then
490           .Ix.Visible = False
500           .Ix_lbl.Visible = False
510           .Ix_lbl_line.Visible = False
520           .Px.Visible = False
530           .Px_lbl.Visible = False
540           .Px_lbl_line.Visible = False
550           .Cx.Visible = False
560           .Cx_lbl.Visible = False
570           .Cx_lbl_line.Visible = False
580         End If

590       Else
600         Cancel = -1
610       End If
620     End With

630     If Cancel = -1 Then
640       Beep
650       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
660     End If

EXITP:
670     Exit Sub

ERRH:
680     Select Case ERR.Number
        Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub Report_Close()
' ** I can't figure out why this report is different!
' ** Most long-standing reports when opened in Preview do the
' ** Maximize, FitToWindow routine, then when they close,
' ** the underlying form is still normal size. This one
' ** leaves the form maximized! I've checked Form settings
' ** and can't see the reason. (Unless it's because this report
' ** comes from a Continuous Form, and the others don't.)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

        Dim blnIsMaximized As Boolean

810     DoCmd.Restore

820     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

830     blnIsMaximized = Forms(strCallingForm).FrmIsMax  ' ** Form Function: frmTransaction_Audit.
840     If blnIsMaximized = False Then
850       gblnSetFocus = True
860       Forms(strCallingForm).TimerInterval = 250&  ' ** 250 thousandths of a second (1/4th of a sec).
870     End If

EXITP:
880     Exit Sub

ERRH:
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub
