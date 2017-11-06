Attribute VB_Name = "modScrollBarFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modScrollBarFuncs"

'VGC 03/22/2017: CHANGES!

' **** CODE START ****
' ** Place this code in a standard module.
' ** make sure you do not name the module
' ** to conflict with any of the functions below.
' **
' ** Author:    Stephen Lebans
' **            Stephen@lebans.com
' **            www.lebans.com
' **            Feb.20, 2000
' **
' ** Copyright: Lebans Holdings 1999 Ltd.
' **
' ** Functions: fGetScrollBarPos(frm As Access.Form) As Long
' **            fSetScrollBarPosVT(frm As Access.Form, lngIndex As Long) As Long
' **
' ** Credits:   Dev Ashish, Terry Kreft
' **            The Access Web
' **            http://www.mvps.org/access/
' **
' ** Why?:      Somebody asked for it!
' **
' ** BUGS:      Let me know!
' **            :-)

Private Type SCROLLINFO
  cbSize As Long
  fMask As Long
  nMin As Long
  nMax As Long
  nPage As Long
  nPos As Long
  nTrackPos As Long
End Type

Private Declare Function GetScrollInfo Lib "user32.dll" (ByVal hwnd As Long, ByVal N As Long, lpScrollInfo As SCROLLINFO) As Long

Private Declare Function SetScrollInfo Lib "user32.dll" (ByVal hwnd As Long, ByVal N As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long

'Private Const ACC_MAIN_CLASS = "OMain"
'Private Const ACC_FORM_CLASS = "OForm"

' ** Window Style Flags.
'Private Const WS_VISIBLE As Double = &H10000000
'Private Const WS_VSCROLL As Double = &H200000

' ** Scroll Bar Styles.
Private Const SBS_HORZ    As Integer = &H0&
Private Const SBS_VERT    As Integer = &H1&
Private Const SBS_SIZEBOX As Integer = &H8&

' ** ScrollBar Message.
'Private Const SBM_SETPOS         As Integer = &HE0
'Private Const SBM_GETPOS         As Integer = &HE1  ' ** /*not in win3.1 */
'Private Const SBM_SETRANGE       As Integer = &HE2  ' ** /*not in win3.1 */
'Private Const SBM_SETRANGEREDRAW As Integer = &HE6  ' ** /*not in win3.1 */
'Private Const SBM_GETRANGE       As Integer = &HE3  ' ** /*not in win3.1 */
'Private Const SBM_ENABLE_ARROWS  As Integer = &HE4  ' ** /*not in win3.1 */
'#if(WINVER >= 0x0400)
'Private Const SBM_SETSCROLLINFO  As Integer = &HE9
'Private Const SBM_GETSCROLLINFO  As Integer = &HEA

' * ScrollInfo fMask's.
Private Const SIF_RANGE           As Integer = &H1
Private Const SIF_PAGE            As Integer = &H2
Private Const SIF_POS             As Integer = &H4
'Private Const SIF_DISABLENOSCROLL As Integer = &H8
Private Const SIF_TRACKPOS        As Integer = &H10
Private Const SIF_ALL             As Integer = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
'(&H1 Or &H2 Or &H4 Or &H10)

'Private Const SB_HORZ          As Integer = 0  ' ** Scroll Bar Constants.
Private Const SB_CTL           As Integer = 2
'Private Const SB_VERT          As Integer = 1
'Private Const SB_LINEUP        As Integer = 0  ' ** Scroll Bar Commands.
'Private Const SB_LINELEFT      As Integer = 0
'Private Const SB_LINEDOWN      As Integer = 1
'Private Const SB_LINERIGHT     As Integer = 1
'Private Const SB_PAGEUP        As Integer = 2
'Private Const SB_PAGELEFT      As Integer = 2
'Private Const SB_PAGEDOWN      As Integer = 3
'Private Const SB_PAGERIGHT     As Integer = 3
Private Const SB_THUMBPOSITION As Integer = 4
'Private Const SB_THUMBTRACK    As Integer = 5
'Private Const SB_TOP           As Integer = 6
'Private Const SB_LEFT          As Integer = 6
'Private Const SB_BOTTOM        As Integer = 7
'Private Const SB_RIGHT         As Integer = 7
'Private Const SB_ENDSCROLL     As Integer = 8
' **

Public Function fGetScrollBarPosVT(frm As Access.Form) As Long
' ** Return ScrollBar Thumb position for the Vertical Scrollbar
' ** attached to the Form passed to this Function.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "fGetScrollBarPosVT"

        Dim sInfo As SCROLLINFO
        Dim lngHWndSB As Long
        Dim lngRet As Long
        Dim lngRetVal As Long

110     lngRetVal = -1&

        ' ** Init SCROLLINFO structure.
120     sInfo.fMask = SIF_ALL
130     sInfo.cbSize = Len(sInfo)
140     sInfo.nPos = 0
150     sInfo.nTrackPos = 0

        ' ** Call function to get handle to ScrollBar control if it is visible.
160     lngHWndSB = fIsScrollBarVT(frm)  ' ** Function: Below.
170     If lngHWndSB <> -1& Then
          ' ** Get the window's ScrollBar position.
180       lngRet = GetScrollInfo(lngHWndSB, SB_CTL, sInfo)  ' ** API Function: Above.
          'Debug.Print "nPos:" & sInfo.nPos & "  nPage:" & sInfo.nPage & "  nMax:" & sInfo.nMax
          'MsgBox "getscrollinfo returned " & sInfo.nPos & " , " & sInfo.nTrackPos
190       lngRetVal = sInfo.nPos + 1&
200     End If

EXITP:
210     fGetScrollBarPosVT = lngRetVal
220     Exit Function

ERRH:
230     lngRetVal = -1&
240     Select Case ERR.Number
        Case Else
250       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
260     End Select
270     Resume EXITP

End Function

Public Function fGetScrollBarPosHZ(frm As Access.Form) As Long
' ** Return ScrollBar Thumb position for the Horizontal Scrollbar
' ** attached to the Form passed to this Function.

300   On Error GoTo ERRH

        Const THIS_PROC As String = "fGetScrollBarPosHZ"

        Dim sInfo As SCROLLINFO
        Dim lngHWndSB As Long
        Dim lngRet As Long
        Dim lngRetVal As Long

310     lngRetVal = -1&

        ' ** Init SCROLLINFO structure.
320     sInfo.fMask = SIF_ALL
330     sInfo.cbSize = Len(sInfo)
340     sInfo.nPos = 0
350     sInfo.nTrackPos = 0

        ' ** Call function to get handle to ScrollBar control if it is visible.
360     lngHWndSB = fIsScrollBarHZ(frm)  ' ** Function: Below.
370     If lngHWndSB <> -1& Then
          ' ** Get the window's ScrollBar position.
380       lngRet = GetScrollInfo(lngHWndSB, SB_CTL, sInfo)  ' ** API Function: Above.
          'Debug.Print "nPos:" & sInfo.nPos & "  nPage:" & sInfo.nPage & "  nMax:" & sInfo.nMax
          'MsgBox "getscrollinfo returned " & sInfo.nPos & " , " & sInfo.nTrackPos
390       lngRetVal = sInfo.nPos + 1&
400     End If

EXITP:
410     fGetScrollBarPosHZ = lngRetVal
420     Exit Function

ERRH:
430     lngRetVal = -1&
440     Select Case ERR.Number
        Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Function

Public Function fSetScrollBarPosVT(frm As Access.Form, lngIndex As Long) As Long
' ** Set the Thumb Position for the
' ** Vertical ScrollBar of the Form passed to
' ** this Function.
' ** Remember that we must subtract 1 from the value
' ** passed to this Function for the desired
' ** Scrollbar position
' **
' *** LIMITED TO 32K ***
' ** Need to use ScrollInfo to overcome this limit
' ** Also need to figure out how Access
' ** calculates the ScrollBar page size!

500   On Error GoTo ERRH

        Const THIS_PROC As String = "fSetScrollBarPosVT"

        Dim sInfo As SCROLLINFO
        Dim lngHWndSB As Long
        Dim lngRet As Long
        Dim lngThumb As Long
        Dim lngRetVal As Long

510     lngRetVal = -1&

        ' ** Init SCROLLINFO structure.
520     sInfo.fMask = SIF_ALL
530     sInfo.cbSize = Len(sInfo)
540     sInfo.nPos = 0
550     sInfo.nTrackPos = 0

        ' ** Call function to get handle to
        ' ** ScrollBar control if it is visible.
560     lngHWndSB = fIsScrollBarVT(frm)  ' ** Function: Below.
570     If lngHWndSB <> -1& Then
          ' ** Set the value  for the ScrollBar.
          ' ** This corresponds to the top-most record
          ' ** that will be displayed in the Form.
580       If lngIndex = 999& Then
590         lngRet = GetScrollInfo(lngHWndSB, SB_CTL, sInfo)  ' ** API Function: Above.
600         lngIndex = sInfo.nMax
610       End If
620       lngThumb = MakeDWord(SB_THUMBPOSITION, CInt(lngIndex - 1&))  ' ** Function: Below.
630       lngRet = SendMessage(frm.hwnd, WM_VSCROLL, ByVal lngThumb, ByVal lngHWndSB)  ' ** API Function: modWindowFunctions.
          ' ** Return Success as our new ScrollBar Position.
640       lngRetVal = lngIndex
650     End If

EXITP:
660     fSetScrollBarPosVT = lngRetVal
670     Exit Function

ERRH:
680     lngRetVal = -1&
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Function

Public Function fSetScrollBarPosHZ(frm As Access.Form, lngIndex As Long) As Long
' ** Set the Thumb Position for the
' ** Horizontal ScrollBar of the Form passed to
' ** this Function.
' ** Remember that we must subtract 1 from the value
' ** passed to this Function for the desired
' ** Scrollbar position
' **
' *** LIMITED TO 32K ***
' ** Need to use ScrollInfo to overcome this limit
' ** Also need to figure out how Access
' ** calculates the ScrollBar page size!

800   On Error GoTo ERRH

        Const THIS_PROC As String = "fSetScrollBarPosHZ"

        Dim sInfo As SCROLLINFO
        Dim lngHWndSB As Long
        Dim lngRet As Long
        Dim lngThumb As Long
        Dim lngRetVal As Long

810     lngRetVal = -1&

        ' ** Init SCROLLINFO structure.
820     sInfo.fMask = SIF_ALL
830     sInfo.cbSize = Len(sInfo)
840     sInfo.nPos = 0
850     sInfo.nTrackPos = 0

        ' ** Call function to get handle to ScrollBar control if it is visible.
860     lngHWndSB = fIsScrollBarHZ(frm)  ' ** Function: Below.
870     If lngHWndSB <> -1& Then
          ' ** Set the value  for the ScrollBar.
          ' ** This corresponds to the top most record
          ' ** that will be displayed in the Form.
880       If lngIndex = 999& Then
890         lngRet = GetScrollInfo(lngHWndSB, SB_CTL, sInfo)  ' ** API Function: Above.
900         lngIndex = sInfo.nMax
910       End If
920       lngThumb = MakeDWord(SB_THUMBPOSITION, CInt(lngIndex - 1&))  ' ** Function: Below.
930       lngRet = SendMessage(frm.hwnd, WM_HSCROLL, ByVal lngThumb, ByVal lngHWndSB)  ' ** API Function: modWindowFunctions.
          ' ** Return Success as our new ScrollBar Position.
940       lngRetVal = lngIndex
950     End If

EXITP:
960     fSetScrollBarPosHZ = lngRetVal
970     Exit Function

ERRH:
980     lngRetVal = -1&
990     Select Case ERR.Number
        Case Else
1000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1010    End Select
1020    Resume EXITP

End Function

Private Function fIsScrollBarVT(frm As Access.Form) As Long
' ** Get ScrollBar's Vertical hWnd.

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "fIsScrollBarVT"

        Dim lngHWnd_VSB As Long
        Dim lngHWnd As Long
        Dim lngRetVal As Long

1110    lngRetVal = -1&

1120    lngHWnd = frm.hwnd

        ' ** Let's get the first Child Window of the Form.
1130    lngHWnd_VSB = GetWindow(lngHWnd, GW_CHILD)  ' ** API Function: Above.

        ' ** Let's walk through every Sibling Window of the Form.
1140    Do
          ' ** Thanks to Terry Kreft for explaining
          ' ** why the apiGetParent acll is not required.
          ' ** Terry is in a Class by himself! :-)
          'If apiGetParent(lngHWnd_VSB) <> lngHWnd Then Exit Do

1150      If fGetClassName(lngHWnd_VSB) = "scrollBar" Then  ' ** Function: Below.
1160        If GetWindowLong(lngHWnd_VSB, GWL_STYLE) And SBS_VERT Then  ' ** API Function: modWindowFunctions.
1170          lngRetVal = lngHWnd_VSB
1180        End If
1190      End If

1200      If lngRetVal = -1& Then
            ' ** Let's get the next Sibling Window.
1210        lngHWnd_VSB = GetWindow(lngHWnd_VSB, GW_HWNDNEXT)  ' ** API Function: Above.
            ' ** Let's start the process from the top again.
            ' ** Really just an error check.
1220      Else
1230        Exit Do
1240      End If

1250    Loop While lngHWnd_VSB <> 0&

1260    If lngRetVal = -1& Then
          ' ** Sorry - No Vertical ScrollBar control
          ' ** is currently visible for this Form.
1270    End If

EXITP:
1280    fIsScrollBarVT = lngRetVal
1290    Exit Function

ERRH:
1300    lngRetVal = -1&
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Function

Private Function fIsScrollBarHZ(frm As Access.Form) As Long
' ** Get ScrollBar's Horizontal hWnd.

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "fIsScrollBarHZ"

        Dim lngHWnd_VSB As Long
        Dim lngHWnd As Long
        Dim lngStyle As Long
        Dim lngRetVal As Long

1410    lngRetVal = -1&

1420    lngHWnd = frm.hwnd

        ' ** Let's get first Child Window of the Form.
1430    lngHWnd_VSB = GetWindow(lngHWnd, GW_CHILD)  ' ** API Function: Above.

        ' ** Let's walk through every Sibling Window of the Form.
1440    Do
          ' ** Thanks to Terry Kreft for explaining
          ' ** why the apiGetParent acll is not required.
          ' ** Terry is in a Class by himself! :-)
          'If apiGetParent(lngHWnd_VSB) <> lngHWnd Then Exit Do

1450      If fGetClassName(lngHWnd_VSB) = "scrollBar" Then  ' ** Function: Below.
1460        lngStyle = GetWindowLong(lngHWnd_VSB, GWL_STYLE)  ' ** API Function: modWindowFunctions.
1470        If (lngStyle And SBS_SIZEBOX) = False Then
              'If GetWindowLong(lngHWnd_VSB, GWL_STYLE) And SBS_HORZ Then  ' ** API Function: modWindowFunctions.
1480          If (lngStyle And 1) = SBS_HORZ Then
1490            lngRetVal = lngHWnd_VSB
1500          End If
1510        End If
1520      End If

1530      If lngRetVal = -1& Then
            ' ** Let's get the Next Sibling Window.
1540        lngHWnd_VSB = GetWindow(lngHWnd_VSB, GW_HWNDNEXT)  ' ** API Function: Above.
            ' ** Let's start the process from the top again.
            ' ** Really just an error check.
1550      Else
1560        Exit Do
1570      End If
1580    Loop While lngHWnd_VSB <> 0&

1590    If lngRetVal = -1& Then
          ' ** Sorry - No Vertical ScrollBar control
          ' ** is currently visible for this Form.
1600    End If

EXITP:
1610    fIsScrollBarHZ = lngRetVal
1620    Exit Function

ERRH:
1630    lngRetVal = -1&
1640    Select Case ERR.Number
        Case Else
1650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1660    End Select
1670    Resume EXITP

End Function

Private Function fGetClassName(lngHWnd As Long) As String
' ** From Dev Ashish's Site
' ** The Access Web
' ** http://www.mvps.org/access/

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "fGetClassName"

        Dim strBuffer As String
        Dim lngLen As Long
        Dim strRetVal As String

        Const MAX_LEN As Integer = 255

1710    strRetVal = vbNullString

1720    strBuffer = Space$(MAX_LEN)
1730    lngLen = GetClassName(lngHWnd, strBuffer, MAX_LEN)  ' ** API Function: modWindowFunctions.
1740    If lngLen > 0& Then strRetVal = Left(strBuffer, lngLen)

EXITP:
1750    fGetClassName = strRetVal
1760    Exit Function

ERRH:
1770    strRetVal = RET_ERR
1780    Select Case ERR.Number
        Case Else
1790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1800    End Select
1810    Resume EXITP

End Function

Private Function MakeDWord(intLoWord As Integer, intHiWord As Integer) As Long
' ** Here's the MakeDWord function from the MS KB.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "MakeDWord"

        Dim lngRetVal As Long

1910    lngRetVal = (intHiWord * &H10000) Or (intLoWord And &HFFFF&)

EXITP:
1920    MakeDWord = lngRetVal
1930    Exit Function

ERRH:
1940    lngRetVal = -1&
1950    Select Case ERR.Number
        Case Else
1960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1970    End Select
1980    Resume EXITP

End Function

Public Function fQuery(varInput As Variant) As Long
' ** Return ScrollBar Thumb position
' ** for the Vertical Scrollbar attached to the
' ** Form passed to this Function.

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "fQuery"

        Dim frm As Access.Form
        Dim lngHWndSB As Long
        Dim lngRet As Long
        Dim typSI As SCROLLINFO
        Dim lngRetVal As Long

2010    lngRetVal = -1&

2020    Set frm = Screen.ActiveForm

        ' ** Init SCROLLINFO structure.
2030    typSI.fMask = SIF_ALL
2040    typSI.cbSize = Len(typSI)
2050    typSI.nPos = 0
2060    typSI.nTrackPos = 0

        ' ** Call function to get handle to ScrollBar control if it is visible.
2070    lngHWndSB = fIsScrollBarVT(frm)  ' ** Function: Above.
2080    If lngHWndSB <> -1& Then

          'lngRet = SendMessage(frm.hWnd, WM_VSCROLL, SB_BOTTOM, 0&)  ' ** API Function: modWindowFunctions.

          ' ** Get the window's ScrollBar position.
2090      lngRet = GetScrollInfo(lngHWndSB, SB_CTL, typSI)  ' ** API Function: Above.
          'Debug.Print "nPos:" & typSI.nPos & "  nPage:" & typSI.nPage & "  nMax:" & typSI.nMax
          'MsgBox "getscrollinfo returned " & typSI.nPos & " , " & typSI.nTrackPos
          'lngRetVal = typSI.nPos + 1&
2100      lngRetVal = typSI.nMax
          'Debug.Print "MAX RECS:" & typSI.nMax

          ' ** Set the value for the ScrollBar.
          ' ** This corresponds to the top most record
          ' ** that will be displayed in the Form.
          'lngThumb = MakeDWord(SB_THUMBPOSITION, CInt(lngIndex - 1&))
          'lngRet = SendMessage(frm.hWnd, WM_VSCROLL, ByVal lngThumb, ByVal hWndSB)  ' ** API Function: modWindowFunctions.
          'lngRet = SendMessage(frm.hWnd, WM_VSCROLL, SB_BOTTOM, 0&)  ' ** API Function: modWindowFunctions.

2110    End If

EXITP:
2120    Set frm = Nothing
2130    fQuery = lngRetVal
2140    Exit Function

ERRH:
2150    lngRetVal = -1&
2160    Select Case ERR.Number
        Case Else
2170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2180    End Select
2190    Resume EXITP

End Function
