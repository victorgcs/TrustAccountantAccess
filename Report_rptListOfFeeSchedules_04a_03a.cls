VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfFeeSchedules_04a_03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfFeeSchedules_04a_03a"

'VGC 07/02/2016: CHANGES!

Private Const strCallingForm As String = "frmFeeSchedules"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
140         .PageNumOf.Visible = True
150         .PageNum.Visible = False

160         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
170         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
180         .AsOf_lbl2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

            ' ** Current Sorting and Grouping:
            ' **   Schedule_Minimum
            ' **   Schedule_Name
            ' **   ScheduleDetail_Amount
            ' **   ScheduleDetail_Rate

            ' ** Current RecordSource:
            ' **   Schedule_ID
            ' **   Schedule_Name
            ' **   Schedule_Base
            ' **   Schedule_Minimum
            ' **   ScheduleDetail_ID
            ' **   ScheduleDetail_Rate
            ' **   ScheduleDetail_Rate_display
            ' **   ScheduleDetail_Amount

  ' ** Possible OrderBy's:
  ' **   Sortable Fields          Ascending                                         Descending
  ' **   ***********************  ************************************************  **********************************************************
  ' ** Sub1:
  ' **   Schedule_ID              [Schedule_ID]                                     [Schedule_ID] DESC
  ' **   Schedule_Name            [Schedule_Name]                                   [Schedule_Name] DESC
  ' **   Schedule_Base            [Schedule_Base], [Schedule_Name]                  [Schedule_Base] DESC, [Schedule_Name] DESC
  ' **   Schedule_Minimum         [Schedule_Minimum], [Schedule_Name]               [Schedule_Minimum] DESC, [Schedule_Name] DESC
  ' ** Sub2:
  ' **   ScheduleDetail_Rate      [ScheduleDetail_Rate], [ScheduleDetail_Amount]    [ScheduleDetail_Rate] DESC, [ScheduleDetail_Amount] DESC
  ' **   ScheduleDetail_Amount    [ScheduleDetail_Amount], [ScheduleDetail_Rate]    [ScheduleDetail_Amount] DESC, [ScheduleDetail_Rate] DESC

190         .Caption = "List_of_Fee_Schedules_" & Format(Date, "yymmdd")

200       Else
210         Cancel = -1
220       End If
230     End With

240     If Cancel = -1 Then
250       Beep
260       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
270     End If

EXITP:
280     Exit Sub

ERRH:
290     Forms(strCallingForm).Visible = True
300     DoCmd.Restore
310     Select Case ERR.Number
        Case Else
320       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
330     End Select
340     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

410     With Me
420       If .Page >= 2 Then
430         .Section(acPageHeader).Visible = True
440       Else
450         .Section(acPageHeader).Visible = False
460       End If
470     End With

EXITP:
480     Exit Sub

ERRH:
490     Forms(strCallingForm).Visible = True
500     DoCmd.Restore
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Sub

Private Sub Report_Close()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

610     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

620     If Forms(strCallingForm).Visible = False Then
630       DoCmd.Restore
640       Forms(strCallingForm).Visible = True
650       DoCmd.SelectObject acForm, strCallingForm, False
660       DoCmd.Restore
670     End If

EXITP:
680     Exit Sub

ERRH:
690     Forms(strCallingForm).Visible = True
700     DoCmd.Restore
710     Select Case ERR.Number
        Case Else
720       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
730     End Select
740     Resume EXITP

End Sub
