VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptStatement_Balance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptStatement_Balance"

'VGC 03/14/2013: CHANGES!

Private Const strCallingForm As String = "frmStatementBalance"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim varTmp00 As Variant

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
140         .PageNumOf.Visible = True
150         .PageNum.Visible = False

160         If .Filter = vbNullString Then
              ' ** Let it break on each account.
170         Else
180           .Section(acGroupLevel1Footer).ForceNewPage = acForceNewPageNone
190         End If

            ' ** AcForceNewPage enumeration:
            ' **   0  acForceNewPageNone    None            The current section (the section for which you're setting the property)
            ' **                                            is printed on the current page. (Default)
            ' **   1  acForceNewPageBefore  Before Section  The current section is printed at the top of a new page.
            ' **   2  acForceNewPageAfter   After Section   The section immediately following the current section
            ' **                                            is printed at the top of a new page.
            ' **   3  acForceNewPageBoth    Before & After  The current section is printed at the top of a new page,
            ' **                                            and the next section is printed at the top of a new page.

            ' ** KeepTogether section property:
            ' **   -1  Yes  True   Microsoft Access starts printing the section at the top of the
            ' **                   next page if it can't print the entire section on the current page.
            ' **    0  No   False  Microsoft Access prints as much of the section as possible on the
            ' **                   current page and prints the rest on the next page. (Default)

            ' ** HasContinued section property:
            ' **   -1  True   Part of the current section has been printed on the previous page.
            ' **    0  False  Part of the current section hasn't been printed on the previous page.

            ' ** WillContinue section property:
            ' **   -1  True   The current section continues on the following page.
            ' **    0  False  The current section doesn't continue on the following page.

            ' ** AcNewRowOrCol enumeration:
            ' **   0  acNewRowOrColNone    None            The row or column breaks are determined by the settings in the
            ' **                                           Page Setup dialog box (available by clicking Page Setup on the
            ' **                                           File menu) and the available space on the page. (Default)
            ' **   1  acNewRowOrColBefore  Before Section  Microsoft Access starts printing the current section (the section
            ' **                                           for which you're setting the property, such as a group header
            ' **                                           section) in a new row or column. It then prints the next section,
            ' **                                           such as a detail section, in that same row or column.
            ' **   2  acNewRowOrColAfter   After Section   Microsoft Access starts printing the current section, such as a group
            ' **                                           header section, in the current row or column. It starts printing the
            ' **                                           next section, such as a detail section, in the next row or column.
            ' **   3  acNewRowOrColBoth    Before & After  Microsoft Access starts printing the current section in a new row or column.
            ' **                                           It starts printing the following section in the next row or column.

200         .AsOf.Caption = "As of " & Format(Date, "mmmm dd, yyyy")
210         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

220         Select Case Forms(strCallingForm).chkLegalName
            Case True
230           Select Case IsNull(Forms(strCallingForm).legalname)
              Case True
240             .legalname.Visible = False
250             .legalname2.Visible = False
260             .shortname.Visible = True
270             .shortname2.Visible = True
280           Case False
290             .legalname.Visible = True
300             .legalname2.Visible = True
310             .shortname.Visible = False
320             .shortname2.Visible = False
330           End Select
340         Case False
350           .legalname.Visible = False
360           .legalname2.Visible = False
370           .shortname.Visible = True
380           .shortname2.Visible = True
390         End Select  ' ** chkLegalName.

400         varTmp00 = Forms(strCallingForm).accountno
410         If IsNull(varTmp00) = False Then
420           varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
430           varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
440           .Caption = "Statement_Balance_" & varTmp00 & "_" & Format(Date, "yymmdd")
450         End If

460       Else
470         Cancel = -1
480       End If

490     End With

500     If Cancel = -1 Then
510       Beep
520       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
530     End If

EXITP:
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

Private Sub Report_NoData(Cancel As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

710     MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
720     Cancel = -1

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
