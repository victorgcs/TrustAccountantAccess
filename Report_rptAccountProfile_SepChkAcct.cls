VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAccountProfile_SepChkAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAccountProfile_SepChkAcct"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm1 As String = "frmRpt_AccountProfile"
Private Const strCallingForm2 As String = "frmRpt_AccountReviews"
Private Const strCallingForm3 As String = "frmAccountProfile"
Private Const strCallingForm4 As String = "frmAccountProfile_Add"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strFilter As String
        Dim strAccountNo As String
        Dim blnPageOf As Boolean

110     TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

120     With Me
130       If IsLoaded(strCallingForm1, acForm) = True Or IsLoaded(strCallingForm2, acForm) = True Or _
              IsLoaded(strCallingForm3, acForm) = True Or IsLoaded(strCallingForm4, acForm) Then

140         strAccountNo = gstrAccountNo

150         blnPageOf = False

            ' ** Default to PageNumOf if calling form has no check box.
160         If IsLoaded(strCallingForm1, acForm) = True Then  ' ** Module Function: modFileUtilities.
170           blnPageOf = Forms(strCallingForm1).chkPageOf
180         ElseIf IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Function: modFileUtilities.
190           blnPageOf = Forms(strCallingForm2).chkPageOf
200         ElseIf IsLoaded(strCallingForm3, acForm) = True Then  ' ** Module Function: modFileUtilities.
210           blnPageOf = True
220         ElseIf IsLoaded(strCallingForm4, acForm) = True Then  ' ** Module Function: modFileUtilities.
230           blnPageOf = True
240         End If

250         Select Case blnPageOf
            Case True
260           .PageNumOf.Visible = True
270           .PageNum.Visible = False
280         Case False
290           .PageNum.Visible = True
300           .PageNumOf.Visible = False
310           .PageNumOf.ControlSource = vbNullString
320         End Select

330         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
340         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

350         If IsLoaded(strCallingForm1, acForm) = True Then  ' ** Module Function: modFileUtilities.
360           Select Case Forms(strCallingForm1).opgAccountNumber
              Case Forms(strCallingForm1).opgAccountNumber_optSpecified.OptionValue
370             Select Case gblnSeparateCheckingAccounts
                Case True
                  ' ** Account, w/add'l fields, for separate checking; by specified FormRef('accountno').
380               .RecordSource = "qryRpt_AccountProfile_03b"
390             Case False
                  ' ** Account, w/add'l fields, for single checking; by specified FormRef('accountno').
400               .RecordSource = "qryRpt_AccountProfile_03d"
410             End Select
420           Case Forms(strCallingForm1).opgAccountNumber_optAll.OptionValue
430             Select Case gblnSeparateCheckingAccounts
                Case True
                  ' ** Account, w/add'l fields, for separate checking; all accounts.
440               .RecordSource = "qryRpt_AccountProfile_03a"
450             Case False
                  ' ** Account, w/add'l fields, for single checking; all accounts.
460               .RecordSource = "qryRpt_AccountProfile_03c"
470             End Select
480           End Select
490         ElseIf IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Function: modFileUtilities.
500           Select Case gblnSeparateCheckingAccounts
              Case True
                ' ** Account, w/add'l fields, for separate checking; by specified FormRef('accountno').
510             .RecordSource = "qryRpt_AccountProfile_03b"
520           Case False
                ' ** Account, w/add'l fields, for single checking; by specified FormRef('accountno').
530             .RecordSource = "qryRpt_AccountProfile_03d"
540           End Select
550         ElseIf IsLoaded(strCallingForm3, acForm) = True Or IsLoaded(strCallingForm4, acForm) = True Then   ' ** Module Function: modFileUtilities.
560           Select Case gblnSeparateCheckingAccounts
              Case True
                ' ** Account, w/add'l fields, for separate checking; by specified FormRef('accountno').
570             .RecordSource = "qryRpt_AccountProfile_03b"
580           Case False
                ' ** Account, w/add'l fields, for single checking; by specified FormRef('accountno').
590             .RecordSource = "qryRpt_AccountProfile_03d"
600           End Select
610           If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
                ' ** It's in Preview.
620             If IsLoaded(strCallingForm3, acForm) = True Then   ' ** Module Function: modFileUtilities.
630               Forms(strCallingForm3).Visible = False
640             ElseIf IsLoaded(strCallingForm4, acForm) = True Then   ' ** Module Function: modFileUtilities.
650               Forms(strCallingForm4).Visible = False
660             End If
670           End If
680         End If

690         If strAccountNo = "All" Then
700           strFilter = vbNullString
710           .Filter = strFilter
720           .FilterOn = False
730         Else
740           strFilter = "[accountno] = '" & strAccountNo & "'"
750           .Filter = strFilter
760           .FilterOn = True
770         End If

780         .Caption = "Account_Profile_" & strAccountNo & "_" & Format(Date, "yymmdd")

790       Else
800         Cancel = -1
810       End If
820     End With

830     If Cancel = -1 Then
840       Beep
850       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
860     End If

EXITP:
870     Exit Sub

ERRH:
880     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
890       If IsLoaded(strCallingForm3, acForm) = True Then   ' ** Module Function: modFileUtilities.
900         Forms(strCallingForm3).Visible = True
910       ElseIf IsLoaded(strCallingForm4, acForm) = True Then   ' ** Module Function: modFileUtilities.
920         Forms(strCallingForm4).Visible = True
930       End If
940     End If
950     Select Case ERR.Number
        Case Else
960       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
970     End Select
980     Resume EXITP

End Sub

Private Sub Report_Close()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1010    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1020    If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1030      If IsLoaded(strCallingForm3, acForm) = True Then   ' ** Module Function: modFileUtilities.
1040        DoCmd.SelectObject acForm, "frmMenu_Title", True
1050        DoCmd.Restore
1060        DoCmd.SelectObject acForm, strCallingForm3, False
1070      ElseIf IsLoaded(strCallingForm4, acForm) = True Then   ' ** Module Function: modFileUtilities.
1080        Forms(strCallingForm4).Visible = True
1090        DoCmd.SelectObject acForm, "frmMenu_Title", True
1100        DoCmd.Restore
1110        DoCmd.SelectObject acForm, strCallingForm4, False
1120      End If
1130    End If

EXITP:
1140    Exit Sub

ERRH:
1150    If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
          ' ** It's in Preview.
1160      If IsLoaded(strCallingForm3, acForm) = True Then   ' ** Module Function: modFileUtilities.
1170        Forms(strCallingForm3).Visible = True
1180      ElseIf IsLoaded(strCallingForm4, acForm) = True Then   ' ** Module Function: modFileUtilities.
1190        Forms(strCallingForm4).Visible = True
1200      End If
1210    End If
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub
