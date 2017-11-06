VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptVersion_Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptVersion_Convert"

'VGC 08/11/2012: CHANGES!

Private Const strCallingForm1 As String = "frmVersion_Main"
Private Const strCallingForm2 As String = "frmVersion_Log"

Private lngVerCnvID As Long, intMode As Integer
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strFilter As String
        Dim strDocName As String
        Dim varTmp00 As Variant

110     With Me
120       If IsLoaded(strCallingForm1, acForm) = True Then  ' ** Module Function: modFileUtilities.
130         lngVerCnvID = Forms(strCallingForm1).vercnv_id
140       ElseIf IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Function: modFileUtilities.
150         lngVerCnvID = Forms(strCallingForm2).vercnv_id
160       Else
170         varTmp00 = InputBox("Enter Conversion ID (vercnv_id):", "Enter vercnv_id", 1)
180         Select Case IsNull(varTmp00)
            Case True
190           Cancel = -1
200         Case False
210           Select Case varTmp00
              Case vbNullString
220             Cancel = -1
230           Case Else
240             lngVerCnvID = Val(varTmp00)
250           End Select
260         End Select
270       End If
280       If lngVerCnvID > 0& Then

290         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
300         .PageNumOf.Visible = True
310         .PageNum.Visible = False

320         intMode = acViewNormal
330         If IsNull(.OpenArgs) = False Then
340           If Trim(.OpenArgs) <> vbNullString Then
350             If .OpenArgs = "Preview" Then
360               intMode = acViewPreview
370             ElseIf .OpenArgs = "Normal" Then
380               intMode = acViewNormal
390             End If
400           End If
410         End If

420         If intMode = acViewPreview Then
430           If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Function: modFileUtilities.
440             Forms(strCallingForm2).Visible = False
450             strDocName = "frmReportList"
460             If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
470               Forms(strDocName).Visible = False
480             End If
490             strDocName = "frmMenu_Report"
500             If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
510               Forms(strDocName).Visible = False
520             End If
530           End If
540         End If

550         strFilter = "[vercnv_id] = " & CStr(lngVerCnvID)
560         .Filter = strFilter
570         .FilterOn = True

580         .AsOf.Caption = Format(Now(), "mmmm dd, yyyy hh:nn:ss AM/PM")
590         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

600       Else
610         Cancel = -1
620       End If
630     End With

640     If Cancel = -1 And strCallingForm1 = vbNullString Then
          ' ** Just for the record.
650       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
660     End If

EXITP:
670     Exit Sub

ERRH:
680     DoCmd.Restore
690     strDocName = "frmMenu_Report"
700     If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
710       If Forms(strDocName).Visible = False Then
720         Forms(strDocName).Visible = True
730       End If
740     End If
750     strDocName = "frmReportList"
760     If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
770       If Forms(strDocName).Visible = False Then
780         Forms(strDocName).Visible = True
790       End If
800     End If
810     If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Function: modFileUtilities.
820       If Forms(strCallingForm2).Visible = False Then
830         Forms(strCallingForm2).Visible = True
840       End If
850     End If
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub Report_Close()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

        Dim strDocName As String

910     If intMode = acViewPreview Then
920       DoCmd.Restore
930       strDocName = "frmMenu_Report"
940       If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
950         If Forms(strDocName).Visible = False Then
960           Forms(strDocName).Visible = True
970         End If
980       End If
990       strDocName = "frmReportList"
1000      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
1010        If Forms(strDocName).Visible = False Then
1020          Forms(strDocName).Visible = True
1030        End If
1040      End If
1050      If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Function: modFileUtilities.
1060        If Forms(strCallingForm2).Visible = False Then
1070          Forms(strCallingForm2).Visible = True
1080        End If
1090      End If
1100    End If

1110    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1120    Exit Sub

ERRH:
1130    DoCmd.Restore
1140    strDocName = "frmMenu_Report"
1150    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
1160      If Forms(strDocName).Visible = False Then
1170        Forms(strDocName).Visible = True
1180      End If
1190    End If
1200    strDocName = "frmReportList"
1210    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
1220      If Forms(strDocName).Visible = False Then
1230        Forms(strDocName).Visible = True
1240      End If
1250    End If
1260    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Function: modFileUtilities.
1270      If Forms(strCallingForm2).Visible = False Then
1280        Forms(strCallingForm2).Visible = True
1290      End If
1300    End If
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Sub
