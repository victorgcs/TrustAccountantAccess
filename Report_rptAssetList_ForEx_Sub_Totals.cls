VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAssetList_ForEx_Sub_Totals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAssetList_ForEx_Sub_Totals"

'VGC 01/25/2016: CHANGES!

Private Const strCallingForm As String = "frmStatementParameters"

Private frm As Access.Form
Private blnPriceHistory As Boolean
Private strCallingReport As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120   On Error Resume Next
130       strCallingReport = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingReport = "rptAssetList_ForEx" Then

160         blnPriceHistory = False

170         Set frm = Forms(strCallingForm)

180         blnPriceHistory = frm.UsePriceHistory

190         Select Case gblnMessage
            Case True
              ' ** Rollbacks were needed.
200           Select Case frm.chkStatements
              Case True
210             Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' ** qryStatementParameters_AssetList_81_29_01 (xx), rounded, with TotalCost_str, TotalMarket_str.
                  '.RecordSource = "qryStatementParameters_AssetList_81_29_02"
220             Case False
                  ' ** qryStatementParameters_AssetList_81_09_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                  '.RecordSource = "qryStatementParameters_AssetList_81_09_02"
230             End Select  ' ** blnPriceHistory.
240           Case False
250             Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' ** qryStatementParameters_AssetList_81_28_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                  '.RecordSource = "qryStatementParameters_AssetList_81_28_02"
260             Case False
                  ' ** qryStatementParameters_AssetList_81_08_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                  '.RecordSource = "qryStatementParameters_AssetList_81_08_02"
270             End Select  ' ** blnPriceHistory.
280           End Select  ' ** chkStatements.
290         Case False
              ' ** No Rollbacks needed.
300           Select Case frm.chkRelatedAccounts
              Case True
                ' ** With Related Accounts.
310             Select Case frm.chkStatements
                Case True
320               Select Case blnPriceHistory
                  Case True
                    ' ** PRICING HISTORY!
                    ' ** qryStatementParameters_AssetList_81_34_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                    '.RecordSource = "qryStatementParameters_AssetList_81_34_02"
330               Case False
                    ' ** qryStatementParameters_AssetList_81_14_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                    '.RecordSource = "qryStatementParameters_AssetList_81_14_02"
340               End Select  ' ** blnPriceHistory.
350             Case False
360               Select Case blnPriceHistory
                  Case True
                    ' ** PRICING HISTORY!
                    ' ** qryStatementParameters_AssetList_81_33_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                    '.RecordSource = "qryStatementParameters_AssetList_81_33_02"
370               Case False
                    ' ** qryStatementParameters_AssetList_81_13_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                    '.RecordSource = "qryStatementParameters_AssetList_81_13_02"
380               End Select  ' ** blnPriceHistory.
390             End Select  ' ** chkStatements.
400           Case False
                ' ** Without Related Accounts.
410             Select Case frm.opgAccountNumber
                Case frm.opgAccountNumber_optSpecified.OptionValue
                  ' ** One Account.
420               Select Case frm.chkNoAssets
                  Case True
430                 Select Case frm.chkStatements
                    Case True
440                   Select Case blnPriceHistory
                      Case True
                        ' ** PRICING HISTORY!
                        ' ** qryStatementParameters_AssetList_81_26_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                        '.RecordSource = "qryStatementParameters_AssetList_81_26_02"
450                   Case False
                        ' ** qryStatementParameters_AssetList_81_06_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                        '.RecordSource = "qryStatementParameters_AssetList_81_06_02"
460                   End Select  ' ** blnPriceHistory.
470                 Case False
480                   Select Case blnPriceHistory
                      Case True
                        ' ** PRICING HISTORY!
                        ' ** qryStatementParameters_AssetList_81_23_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                        '.RecordSource = "qryStatementParameters_AssetList_81_23_02"
490                   Case False
                        ' ** qryStatementParameters_AssetList_81_03_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                        '.RecordSource = "qryStatementParameters_AssetList_81_03_02"
500                   End Select  ' ** blnPriceHistory.
510                 End Select  ' ** chkStatements.
520               Case False
530                 Select Case frm.chkStatements
                    Case True
540                   Select Case blnPriceHistory
                      Case True
                        ' ** PRICING HISTORY!
                        ' ** qryStatementParameters_AssetList_81_25_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                        '.RecordSource = "qryStatementParameters_AssetList_81_25_02"
550                   Case False
                        ' ** qryStatementParameters_AssetList_81_05_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                        '.RecordSource = "qryStatementParameters_AssetList_81_05_02"
560                   End Select  ' ** blnPriceHistory.
570                 Case False
580                   Select Case blnPriceHistory
                      Case True
                        ' ** PRICING HISTORY!
                        ' ** qryStatementParameters_AssetList_81_22_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                        '.RecordSource = "qryStatementParameters_AssetList_81_22_02"
590                   Case False
                        ' ** qryStatementParameters_AssetList_81_02_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                        '.RecordSource = "qryStatementParameters_AssetList_81_02_02"
600                   End Select  ' ** blnPriceHistory.
610                 End Select  ' ** chkStatements.
620               End Select  ' ** chkNoAssets.
630             Case frm.opgAccountNumber_optAll.OptionValue
                  ' ** All Accounts.
640               Select Case frm.chkStatements
                  Case True
650                 Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** qryStatementParameters_AssetList_81_24_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                      '.RecordSource = "qryStatementParameters_AssetList_81_24_02"
660                 Case False
                      ' ** qryStatementParameters_AssetList_81_04_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                      '.RecordSource = "qryStatementParameters_AssetList_81_04_02"
670                 End Select  ' ** blnPriceHistory.
680               Case False
690                 Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** qryStatementParameters_AssetList_81_21_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                      '.RecordSource = "qryStatementParameters_AssetList_81_21_02"
700                 Case False
                      ' ** qryStatementParameters_AssetList_81_01_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
                      '.RecordSource = "qryStatementParameters_AssetList_81_01_02"
710                 End Select  ' ** blnPriceHistory.
720               End Select  ' ** chkStatements.
730             End Select  ' ** opgAccountNumber.
740           End Select  ' ** chkRelatedAccounts.
750         End Select  ' ** gblnMessage.

760       Else
770         Cancel = -1
780       End If
790     End With

800     If Cancel = -1 Then
810       Beep
820       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
830     End If

EXITP:
840     Exit Sub

ERRH:
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
870     End Select
880     Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

        Dim strFormat As String

910     With Me
920       Select Case .currsym_symbol
          Case "CHF"
940         strFormat = "\C\H\F#,##0.00;(\C\H\F#,##0.00);\C\H\F0.00;\C\H\F0.00"
950       Case Else
960         strFormat = .currsym_symbol & "#,##0.00;(" & .currsym_symbol & "#,##0.00);" & .currsym_symbol & "0.00;" & .currsym_symbol & "0.00"
970       End Select
980       .TotalCost.Format = strFormat
990       .TotalMarket.Format = strFormat
1000    End With

EXITP:
1010    Exit Sub

ERRH:
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Sub
