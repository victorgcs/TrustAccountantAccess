VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfLocations"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmLocations"

Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

            ' ** Default to PageNumOf if calling form has no check box.
130         .PageNumOf.Visible = True
140         .PageNum.Visible = False

150         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

160         lngPageHeaderSection_Height = .Section(acPageHeader).Height
170         lngOneLine_Top = .Loc_Name_lbl.Top
180         lngTwoLine_Top = 0&
190         lngUnderline_Top = .Loc_Name_lbl_line.Top

200         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
210         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
220         .AsOf2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

230         strOrderBy = Forms(strCallingForm).frmLocations_Sub.Form.SortNow_Get  ' ** Form Function: frmLocations_Sub.
            ' ** Potential Sorts:
            ' **   [Loc_Name], [Loc_Address1]
            ' **   [Loc_Address1], [Loc_Name]
            ' **   [Loc_City], [Loc_Name], [Loc_Address1]
            ' **   [Loc_State], [Loc_Name], [Loc_Address1]
            ' **   [Loc_Zip], [Loc_Name], [Loc_Address1]
            ' **   [Loc_Country], [Loc_Name], [Loc_Address1]
            ' **   [Loc_PostalCode], [Loc_Name], [Loc_Address1]
            ' **   [Loc_Phone], [Loc_Name], [Loc_Address1]
240         .OrderBy = strOrderBy
250         .OrderByOn = True

260         .Caption = "List_of_Locations_" & Format(Date, "yymmdd")

270       Else
280         Cancel = -1
290       End If
300     End With

310     If Cancel = -1 Then
320       Beep
330       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
340     End If

EXITP:
350     Exit Sub

ERRH:
360     Forms(strCallingForm).Visible = True
370     DoCmd.Restore
380     Select Case ERR.Number
        Case Else
390       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
400     End Select
410     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

510     With Me
520       lngTmp01 = .AsOf2.Height
530       If .Page >= 2 Then
540         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
550         .Loc_Name_lbl.Top = (lngOneLine_Top + lngTmp01)
560         .Loc_Name_lbl_line.Top = (lngUnderline_Top + lngTmp01)
570         .Loc_PhoneExt_lbl.Top = (lngOneLine_Top + lngTmp01)
580         .Loc_PhoneExt_lbl_line.Top = (lngUnderline_Top + lngTmp01)
590         .AsOf2.Visible = True
600         .Header_lbl2.Visible = True
610         .Header_hline05.Visible = True
620       Else
630         .AsOf2.Visible = False
640         .Header_lbl2.Visible = False
650         .Header_hline05.Visible = False
660         .Loc_Name_lbl.Top = lngOneLine_Top
670         .Loc_Name_lbl_line.Top = lngUnderline_Top
680         .Loc_PhoneExt_lbl.Top = lngOneLine_Top
690         .Loc_PhoneExt_lbl_line.Top = lngUnderline_Top
700         .Section(acPageHeader).Height = lngPageHeaderSection_Height
710       End If
720     End With

EXITP:
730     Exit Sub

ERRH:
740     Forms(strCallingForm).Visible = True
750     DoCmd.Restore
760     Select Case ERR.Number
        Case Else
770       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
780     End Select
790     Resume EXITP

End Sub

Private Sub Report_Close()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

810     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

820     If Forms(strCallingForm).Visible = False Then
830       DoCmd.Restore
840       Forms(strCallingForm).Visible = True
850       DoCmd.SelectObject acForm, strCallingForm, False
860       DoCmd.Restore
870     End If

EXITP:
880     Exit Sub

ERRH:
890     Forms(strCallingForm).Visible = True
900     DoCmd.Restore
910     Select Case ERR.Number
        Case Else
920       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
930     End Select
940     Resume EXITP

End Sub
