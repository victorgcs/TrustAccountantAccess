VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfAssets_03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfAssets_03a"

'VGC 11/27/2016: CHANGES!

' ** Sort: [rate], [description_masterasset_sort], [assettype]

Private Const strCallingForm As String = "frmAssets"
Private lngTpp As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim lngTmp01 As Long

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
140         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

            ' ** Default to PageNumOf if calling form has no check box.
150         .PageNumOf.Visible = True
160         .PageNum.Visible = False

170         Set frm = Forms(strCallingForm)
180         Select Case frm.chkShowHidden
            Case True
190           .RecordSource = "qryRpt_ListOfAssets_02"
200         Case False
              ' ** As is: qryRpt_ListOfAssets_01.
210           If .RecordSource <> "qryRpt_ListOfAssets_01" Then
220             .RecordSource = "qryRpt_ListOfAssets_01"
230           End If
240         End Select

250         If Forms(strCallingForm).chkIncludeCurrency = False Then
260           .curr_code.Visible = False
270           .curr_code_lbl.Visible = False
280           .curr_code_lbl_line.Visible = False
290           lngTmp01 = ((.curr_code.Left + .curr_code.Width) - (.totdesc.Left + .totdesc.Width))
300           .totdesc.Width = (.totdesc.Width + lngTmp01)
310           .totdesc_lbl.Width = .totdesc.Width
320           .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
330         End If

340         Select Case frm.opgShow
            Case frm.opgShow_optAll.OptionValue
350           strFilter = vbNullString
360           .Filter = strFilter
370           .FilterOn = False
380           .Caption = "List_of_Assets_" & Format(Date, "yymmdd")
390         Case frm.opgShow_optHeldOnly.OptionValue
400           strFilter = "[shareface] <> 0"
410           .Filter = strFilter
420           .FilterOn = True
430           .Header_lbl.Caption = "List of Assets Held"
440           .Caption = "List_of_Assets_Held_" & Format(Date, "yymmdd")
450         End Select

460         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
470         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

480       Else
490         Cancel = -1
500       End If
510     End With

520     If Cancel = -1 Then
530       Beep
540       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
550     End If

EXITP:
560     Set frm = Nothing
570     Exit Sub

ERRH:
580     DoCmd.Restore
590     Forms(strCallingForm).Visible = True
600     Select Case ERR.Number
        Case Else
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End Select
630     Resume EXITP

End Sub

Private Sub Report_Close()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

710     DoCmd.Restore

720     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

730     If Forms(strCallingForm).Visible = False Then
740       Forms(strCallingForm).Visible = True
750       DoEvents
760     End If

EXITP:
770     Exit Sub

ERRH:
780     DoCmd.Restore
790     Forms(strCallingForm).Visible = True
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub
