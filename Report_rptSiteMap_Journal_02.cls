VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptSiteMap_Journal_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptSiteMap_Journal_02"

'VGC 04/04/2016: CHANGES!

Private strCallingForm As String, strCallersCallingForm As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me

120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If strTmp01 <> vbNullString Then
150           intPos01 = InStr(strTmp01, "~")
160           If intPos01 > 0 Then
170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strCallersCallingForm = Mid(strTmp01, (intPos01 + 1))
190           End If
200         End If
210       End If

220       .Caption = "Trust_Accountant_Site_Map_Journal_Pt2"

230       TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

240     End With

EXITP:
250     Exit Sub

ERRH:
260     If strCallingForm <> vbNullString Then
270       If IsLoaded(strCallersCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
280         If IsLoaded("rptSiteMap", acReport) = False Then  ' ** Module Function: modFileUtilities.
290   On Error Resume Next
300           Forms(strCallersCallingForm).Visible = True
310         End If
320       End If
330       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
340         If IsLoaded("rptSiteMap", acReport) = False Then  ' ** Module Function: modFileUtilities.
350   On Error Resume Next
360           Forms(strCallingForm).Visible = True
370         End If
380       End If
390     End If
400     If ERR.Number <> 0 Then
410       Select Case ERR.Number
          Case Else
420         zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430       End Select
440     End If
450     Resume EXITP

End Sub

Private Sub Report_Close()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

510     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

520     If IsLoaded("rptSiteMap_01", acReport) = True Or IsLoaded("rptSiteMap_02", acReport) = True Or _
            IsLoaded("rptSiteMap_03", acReport) = True Or IsLoaded("rptSiteMap_04", acReport) = True Or _
            IsLoaded("rptSiteMap_Journal_01", acReport) = True Then  ' ** Module Function: modFileUtilities.
          ' ** Only if it's the last one.
530     Else
540       If strCallingForm <> vbNullString Then
550         If IsLoaded(strCallersCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
560           If Forms(strCallersCallingForm).Visible = False Then
570             If IsLoaded("rptSiteMap", acReport) = False Then  ' ** Module Function: modFileUtilities.
580               Forms(strCallersCallingForm).Visible = True
590             End If
600           End If
610         End If
620         If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
630           If Forms(strCallingForm).Visible = False Then
640             If IsLoaded("rptSiteMap", acReport) = False Then  ' ** Module Function: modFileUtilities.
650               Forms(strCallingForm).Visible = True
660             End If
670           End If
680         End If
690       End If
700     End If

EXITP:
710     Exit Sub

ERRH:
720     If strCallingForm <> vbNullString Then
730       If IsLoaded(strCallersCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
740         If IsLoaded("rptSiteMap", acReport) = False Then  ' ** Module Function: modFileUtilities.
750   On Error Resume Next
760           Forms(strCallersCallingForm).Visible = True
770         End If
780       End If
790       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
800         If IsLoaded("rptSiteMap", acReport) = False Then  ' ** Module Function: modFileUtilities.
810   On Error Resume Next
820           Forms(strCallingForm).Visible = True
830         End If
840       End If
850     End If
860     If ERR.Number <> 0 Then
870       Select Case ERR.Number
          Case Else
880         zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890       End Select
900     End If
910     Resume EXITP

End Sub
