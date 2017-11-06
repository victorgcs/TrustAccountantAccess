VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptVBComponent_Monitor_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'VGC 03/01/2017: CHANGES!

Private Const THIS_NAME As String = "rptVBComponent_Monitor_02"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me

          ' ** Default to PageNumOf if calling form has no check box.
120       .PageNumOf.Visible = True
130       .PageNum.Visible = False

140       .detail_height_line.Height = 15&  ' ** Just get it out of the way.

150       .Caption = "Forms_Affected_By_Dual_Monitors_2_" & Format(Date, "yymmdd")

160     End With

EXITP:
170     Exit Sub

ERRH:
180     Select Case ERR.Number
        Case Else
190       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
200     End Select
210     Resume EXITP

End Sub
