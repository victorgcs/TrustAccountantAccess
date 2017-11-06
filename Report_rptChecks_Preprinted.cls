VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptChecks_Preprinted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptChecks_Preprinted"

'VGC 07/13/2017: CHANGES!

' **********************************************
' ** NOTE: Popup required for proper printing!
' **********************************************

Private Const strCallingForm As String = "frmRpt_Checks"

Private strAccountNo As String
Private intMode As Integer
Private strTmpCallingForm As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form
        Dim strSortOrd As String, blnDesc As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
130         If IsNull(.OpenArgs) = False Then
140           strTmp01 = Trim(.OpenArgs)
150           If strTmp01 <> vbNullString Then
160             intPos01 = InStr(strTmp01, "~")
170             If intPos01 > 0 Then

180               strSortOrd = Left(strTmp01, (intPos01 - 1))
190               strTmp01 = Mid(strTmp01, (intPos01 + 1))
200               intPos01 = InStr(strTmp01, "~")
210               blnDesc = CBool(Left(strTmp01, (intPos01 - 1)))
220               strTmp01 = Mid(strTmp01, (intPos01 + 1))
230               intPos01 = InStr(strTmp01, "~")
240               intMode = CInt(Left(strTmp01, (intPos01 - 1)))
250               strAccountNo = Mid(strTmp01, (intPos01 + 1))

260               If IsLoaded(strCallingForm, acForm) = True Then
270                 Set frm = Forms(strCallingForm)
280                 If frm.opgPayeeFont.Visible = True Then
290                   Select Case frm.opgPayeeFont
                      Case frm.opgPayeeFont_opt10Pt.OptionValue
300                     .RecurringItem1.Visible = True
310                     .RecurringItem1_12pt.Visible = False
320                   Case frm.opgPayeeFont_opt12Pt.OptionValue
330                     .RecurringItem1.Visible = False
340                     .RecurringItem1_12pt.Visible = True
350                   End Select
360                 Else
370                   .RecurringItem1.Visible = True
380                   .RecurringItem1_12pt.Visible = False
390                 End If
400                 Set frm = Nothing
410               Else
420                 .RecurringItem1.Visible = True
430                 .RecurringItem1_12pt.Visible = False
440               End If

450               TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

460               strTmpCallingForm = strCallingForm
470               Set frm = Forms(strCallingForm)

480               Select Case frm.opgPrint
                  Case frm.opgPrint_optAll.OptionValue
490                 .Caption = "Checks_" & Format(Date, "yyyymmdd") & "_All"
500               Case frm.opgPrint_optSpecified.OptionValue
510                 .Caption = "Checks_" & Format(Date, "yyyymmdd") & "_" & strAccountNo
520               Case frm.opgPrint_optBlank.OptionValue
                    ' ** Shouldn't be here!
530                 Cancel = -1
540               End Select

550               blnRetVal = Rpt_GroupLevel_Set(Me, strSortOrd, blnDesc)  ' ** Module Function: modReportFunctions.

560             Else
570               Cancel = -1
580             End If
590           Else
600             Cancel = -1
610           End If
620         Else
630           Cancel = -1
640         End If
650       Else
660         Cancel = -1
670       End If
680     End With

690     If Cancel = -1 Then
700       Beep
710       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
720     End If

EXITP:
730     Exit Sub

ERRH:
740     If Forms(strTmpCallingForm).Visible = False Then
750       Forms(strTmpCallingForm).Visible = True
760     End If
770     Select Case ERR.Number
        Case Else
780       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
790     End Select
800     Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

        Dim strPayeeCSZ As String

910     With Me

          ' ** AmountText1: =GetDollarString([amount]*-1)

920       If IsNull(.txtPayeeAddress) = False Then
930         If .txtPayeeAddress <> vbNullString Then
940           If IsNull(.RecurringItem_Countryx) = True And IsNull(.RecurringItem_PostalCodex) = True Then
950             strPayeeCSZ = Trim(IIf(IsNull(.RecurringItem_Cityx) = True, vbNullString, _
                  IIf(Trim(.RecurringItem_Cityx) = vbNullString, vbNullString, .RecurringItem_Cityx & ", ")) & _
                  Nz(.RecurringItem_Statex, vbNullString) & " " & FormatZip(Nz(.RecurringItem_Zipx, vbNullString)))
960           Else
970             If Trim(.RecurringItem_Countryx) = vbNullString And Trim(.RecurringItem_PostalCodex) = vbNullString Then
980               strPayeeCSZ = Trim(IIf(IsNull(.RecurringItem_Cityx) = True, vbNullString, _
                    IIf(Trim(.RecurringItem_Cityx) = vbNullString, vbNullString, .RecurringItem_Cityx & ", ")) & _
                    Nz(.RecurringItem_Statex, vbNullString) & " " & FormatZip(Nz(.RecurringItem_Zipx, vbNullString)))
990             Else
1000              strPayeeCSZ = Trim(IIf(IsNull(.RecurringItem_Cityx) = True, vbNullString, _
                    IIf(Trim(.RecurringItem_Cityx) = vbNullString, vbNullString, .RecurringItem_Cityx & ", ")) & _
                    Nz(.RecurringItem_Countryx, vbNullString) & " " & Nz(.RecurringItem_PostalCodex, vbNullString))
1010            End If
1020          End If
1030        Else
1040          strPayeeCSZ = vbNullString
1050        End If
1060      Else
1070        strPayeeCSZ = vbNullString
1080      End If

1090      If Trim(strPayeeCSZ) = "," Then
1100        .txtPayeeCSZ = vbNullString
1110      ElseIf Trim(strPayeeCSZ) = vbNullString Then
1120        .txtPayeeCSZ = vbNullString
1130      Else
1140        .txtPayeeCSZ = strPayeeCSZ
1150      End If

1160    End With

EXITP:
1170    Exit Sub

ERRH:
1180    If Forms(strTmpCallingForm).Visible = False Then
1190      Forms(strTmpCallingForm).Visible = True
1200    End If
1210    Select Case ERR.Number
        Case Else
1220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1230    End Select
1240    Resume EXITP

End Sub

Private Sub Report_Close()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1310    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1320    If Forms(strTmpCallingForm).Visible = False Then
1330      Forms(strTmpCallingForm).Visible = True
1340    End If

1350    If intMode = acViewPreview Then
1360      gblnSetFocus = True
1370      Forms(strTmpCallingForm).TimerInterval = 250&
1380    End If

EXITP:
1390    Exit Sub

ERRH:
1400    If Forms(strTmpCallingForm).Visible = False Then
1410      Forms(strTmpCallingForm).Visible = True
1420    End If
1430    Select Case ERR.Number
        Case Else
1440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1450    End Select
1460    Resume EXITP

End Sub
