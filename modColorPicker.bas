Attribute VB_Name = "modColorPicker"
Option Compare Database
Option Explicit

'VGC 04/09/2016: CHANGES!

' **  Original Code by Terry Kreft
' **  Modified by Stephen Lebans
' **  Contact Stephen@lebans.com

Private Const THIS_NAME As String = "modColorPicker"
'***********  Code Start  ***********
' **

Public Function aDialogColor(ByVal hwnd As Long) As Long

100   On Error GoTo ERRH

        Const THIS_PROC As String = "aDialogColor"

        Dim lngX As Long, typCS As COLORSTRUC ', arr_varCustColor(16) As Long
        Dim lngRetVal As Long

110     lngRetVal = 0&

120     typCS.lStructSize = Len(typCS)

130     If hwnd <> 0 Then
140       typCS.hwnd = hwnd
150     Else
160       typCS.hwnd = Application.hWndAccessApp
170     End If
180     typCS.flags = CC_SOLIDCOLOR
190     typCS.lpCustColors = String$(16 * 4, 0)
200     lngX = ChooseColor(typCS)  ' ** API Function: modWindowFunctions.
210     If lngX = 0& Then
          ' ** ERROR - use Default White.
          'prop = RGB(255, 255, 255)  ' ** White.
220       lngRetVal = -1&  ' ** False.
230     Else
          ' ** Normal processing.
240       lngRetVal = typCS.rgbResult
250     End If

EXITP:
260     aDialogColor = lngRetVal
270     Exit Function

ERRH:
280     lngRetVal = -1&
290     Select Case ERR.Number
        Case Else
300       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
310     End Select
320     Resume EXITP

End Function
' ***********  Code End   ***********
