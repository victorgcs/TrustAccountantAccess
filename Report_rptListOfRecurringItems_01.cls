VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfRecurringItems_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfRecurringItems_01"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmRecurringItems"

Private lngTpp As Long
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form
        Dim strFilter As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
140         .PageNumOf.Visible = True
150         .PageNum.Visible = False

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         .detail_height_line.Top = 0&  ' ** Just get this out of the way.
180         .detail_height_line.Height = lngTpp

190         lngPageHeaderSection_Height = .Section(acPageHeader).Height
200         lngOneLine_Top = .Recur_Name_lbl.Top
210         lngTwoLine_Top = 0&
220         lngUnderline_Top = .Recur_Name_lbl_line.Top

230         Set frm = Forms(strCallingForm)
240         With frm
250           Select Case .opgType
              Case .opgType_optMisc.OptionValue
260             strFilter = "[Recur_Type]= 'Misc'"
270           Case .opgType_optPaid.OptionValue
280             strFilter = "[Recur_Type]= 'Payee'"
290           Case .opgType_optReceived.OptionValue
300             strFilter = "[Recur_Type]= 'Payor'"
310           Case .opgType_optAll.OptionValue
320             strFilter = vbNullString
330           End Select
340         End With
350         .Filter = strFilter
360         If strFilter = vbNullString Then
370           .FilterOn = False
380         Else
390           .FilterOn = True
400         End If

410         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
420         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
430         .AsOf2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

440         .Caption = "List_of_Recurring_Items_" & Format(Date, "yymmdd")

450       Else
460         Cancel = -1
470       End If
480     End With

490     If Cancel = -1 Then
500       Beep
510       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
520     End If

EXITP:
530     Set frm = Nothing
540     Exit Sub

ERRH:
550     Forms(strCallingForm).Visible = True
560     DoCmd.Restore
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

710     With Me
720       lngTmp01 = .AsOf2.Height
730       If .Page >= 2 Then
740         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
750         .Recur_Name_lbl.Top = (lngOneLine_Top + lngTmp01)
760         .Recur_Name_lbl_line.Top = (lngUnderline_Top + lngTmp01)
770         .AsOf2.Visible = True
780         .Header_lbl2.Visible = True
790         .Header_hline05.Visible = True
800       Else
810         .AsOf2.Visible = False
820         .Header_lbl2.Visible = False
830         .Header_hline05.Visible = False
840         .Recur_Name_lbl.Top = lngOneLine_Top
850         .Recur_Name_lbl_line.Top = lngUnderline_Top
860         .Section(acPageHeader).Height = lngPageHeaderSection_Height
870       End If
880     End With

EXITP:
890     Exit Sub

ERRH:
900     Forms(strCallingForm).Visible = True
910     DoCmd.Restore
920     Select Case ERR.Number
        Case Else
930       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
940     End Select
950     Resume EXITP

End Sub

Private Sub Report_Close()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1010    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1020    If Forms(strCallingForm).Visible = False Then
1030      DoCmd.Restore
1040      Forms(strCallingForm).Visible = True
1050      DoCmd.SelectObject acForm, strCallingForm, False
1060      DoCmd.Restore
1070    End If

EXITP:
1080    Exit Sub

ERRH:
1090    Forms(strCallingForm).Visible = True
1100    DoCmd.Restore
1110    Select Case ERR.Number
        Case Else
1120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1130    End Select
1140    Resume EXITP

End Sub
