Attribute VB_Name = "zz_mod_SystemColorFuncs"
Option Compare Database
Option Explicit

'VGC 04/19/2016: CHANGES!

' ** VbSystemColors enumeration:
' **   -2147483648  vbScrollBars                0   COLOR_SCROLLBAR                0x80000000  &H80000000  Color of the gray area of a scroll bar.
' **   -2147483647  vbDesktop                   1   COLOR_BACKGROUND               0x80000001  &H80000001  Background color of the desktop window.
' **   -2147483646  vbActiveTitleBar            2   COLOR_ACTIVECAPTION            0x80000002  &H80000002  Color of the title bar for the active window.
' **   -2147483645  vbInactiveTitleBar          3   COLOR_INACTIVECAPTION          0x80000003  &H80000003  Color of the title bar for the inactive window.
' **   -2147483644  vbMenuBar                   4   COLOR_MENU                     0x80000004  &H80000004  Background color of a menu.
' **   -2147483643  vbWindowBackground          5   COLOR_WINDOW                   0x80000005  &H80000005  Background color of a window.
' **   -2147483642  vbWindowFrame               6   COLOR_WINDOWFRAME              0x80000006  &H80000005  Color of a window frame.
' **   -2147483641  vbMenuText                  7   COLOR_MENUTEXT                 0x80000007  &H80000006  Color of the text in a menu.
' **   -2147483640  vbWindowText                8   COLOR_WINDOWTEXT               0x80000008  &H80000007  Color of the text in a window.
' **   -2147483639  vbActiveTitleBarText        9   COLOR_CAPTIONTEXT              0x80000009  &H80000008  Color of the text in a title bar and of the size box and scroll bar arrow box. (mine, replaces vbTitleBarText)
' **   -2147483638  vbActiveBorder              10  COLOR_ACTIVEBORDER             0x8000000A  &H8000000A  Color of the border of an active window.
' **   -2147483637  vbInactiveBorder            11  COLOR_INACTIVEBORDER           0x8000000B  &H8000000B  Color of the border of an inactive window.
' **   -2147483636  vbApplicationWorkspace      12  COLOR_APPWORKSPACE             0x8000000C  &H8000000C  Background color of multiple document interface (MDI) applications.
' **   -2147483635  vbHighlight                 13  COLOR_HIGHLIGHT                0x8000000D  &H8000000D  Color of an item selected in a control.
' **   -2147483634  vbHighlightText             14  COLOR_HIGHLIGHTTEXT            0x8000000E  &H8000000E  Color of the text of an item selected in a control.
' **   -2147483633  vbButtonFace                15  COLOR_BTNFACE                  0x8000000F  &H8000000F  Color of shading on the face of command buttons. (also vbButtonFace)
' **   -2147483632  vbButtonShadow              16  COLOR_BTNSHADOW                0x80000010  &H80000010  Color of shading on the edge of command buttons. (also vbButtonShadow)
' **   -2147483631  vbGrayText                  17  COLOR_GRAYTEXT                 0x80000011  &H80000011  Color of shaded (disabled) text.
' **   -2147483630  vbButtonText                18  COLOR_BTNTEXT                  0x80000012  &H80000012  Color of the text on command buttons.
' **   -2147483629  vbInactiveTitleBarText      19  COLOR_INACTIVECAPTIONTEXT      0x80000013  &H80000013  Color of the text in the title bar of an inactive window. (mine, replaces vbInactiveCaptionText)
' **   -2147483628  vb3DHighlight               20  COLOR_BTNHIGHLIGHT             0x80000014  &H80000014  Highlight color of buttons for edges that face the light source.
' **   -2147483627  vb3DDKShadow                21  COLOR_3DDKSHADOW               0x80000015  &H80000015  Color of the dark shadow for three-dimensional display elements.
' **   -2147483626  vb3DLight                   22  COLOR_3DLIGHT                  0x80000016  &H80000016  Highlight color of three-dimensional display elements for edges that face the light source.
' **   -2147483625  vbInfoText                  23  COLOR_INFOTEXT                 0x80000017  &H80000017  Color of the text for ToolTip controls.
' **   -2147483624  vbInfoBackground            24  COLOR_INFOBK                   0x80000018  &H80000018  Background color for ToolTip controls.
' **   -2147483623  vbStaticBackground          25  COLOR_STATIC                   0x80000019  &H80000019  Background color for static controls and dialog boxes (mine)
' **   -2147483622  vbStaticText                26  COLOR_STATICTEXT               0x8000001A  &H8000001A  Color of the text for static controls (mine)
' **   -2147483621  vbActiveTitleBarGradient    27  COLOR_GRADIENTACTIVECAPTION    0x8000001B  &H8000001B  Color of the title bar of an active window that is filled with a color gradient (mine)
' **   -2147483620  vbInactiveTitleBarGradient  28  COLOR_GRADIENTINACTIVECAPTION  0x8000001C  &H8000001C  Color of the title bar of an inactive window that is filled with a color gradient (mine)
' **   -2147483619  vbMenuHighlightFlat         29  COLOR_MENUHILIGHT              0x8000001D  &H8000001D  Color used to highlight menu items when the menu appears as a flat menu (mine)
' **   -2147483618  vbMenuBackgroundFlat        30  COLOR_MENUBAR                  0x8000001E  &H8000001E  Background color for the menu bar when menus appear as flat menus (mine)

' ** Though the top 2 are in the list from the web, they come up blank
' ** in the Immediate Window, and I prefer them to vbTitleBarText and
' ** vbInactiveCaptionText. The rest are mine to fill out the list.

' ** Format color enumeration:
' **   0         vbBlack
' **   255       vbRed
' **   65280     vbGreen
' **   65535     vbYellow
' **   16711680  vbBlue
' **   16711935  vbMagenta
' **   16776960  vbCyan
' **   16777215  vbWhite

' ** VbSystemColors enumeration:  I'm leaving these Public!
'Public Const vbActiveTitleBarText       As Long = -2147483639
'Public Const vbInactiveTitleBarText     As Long = -2147483629
'Public Const vbStaticBackground         As Long = -2147483623
'Public Const vbStaticText               As Long = -2147483622
'Public Const vbActiveTitleBarGradient   As Long = -2147483621
'Public Const vbInactiveTitleBarGradient As Long = -2147483620
'Public Const vbMenuHighlightFlat        As Long = -2147483619
'Public Const vbMenuBackgroundFlat       As Long = -2147483618

' ** AcDisplay enumeration:  (my own)
'Public Const acDisplayAlways As Integer = 0
'Public Const acDisplayPrint  As Integer = 1
'Public Const acDisplayScreen As Integer = 2
' ** AcDisplay enumeration:  (my own)
' **   0  acDisplayAlways  Always       The object appears in Form view and when printed. (Default)
' **   1  acDisplayPrint   Print Only   The object is hidden in Form view but appears when printed.
' **   2  acDisplayScreen  Screen Only  The object appears in Form view but not when printed.
' **

Private Const THIS_NAME As String = "zz_mod_SystemColorFuncs"
' **

Public Function SystemColor_Move(frm As Access.Form) As Boolean
' ** Let's try to move my 'Nav_' lines to match Windows Standard.
' ** Called by:
' **   frmAccountComments.Form_Open()
' **   frmAccountHideTrans.Form_Open()
' **   frmAccountIncExpCodes.Form_Open()
' **   frmAccountTaxCodes.Form_Open()
' **   frmIncomeExpenseCodes.Form_Open()
' **   frmJournal_Columns.Form_Open()
' **   frmLocations.Form_Open()
' **   frmMasterBalance.Form_Open()
' **   frmMenu_Account.Form_Open()
' **   frmRecurringItems.Form_Open()
' **   frmStatementBalance.Form_Open()
' **   frmTransaction_Audit.Form_Open()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "SystemColor_Move"

        Dim ctl As Access.Control
        Dim lngSysColor As Long, lngRealColor As Long, lngTpp As Long
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     lngSysColor = frm.Detail.BackColor

130     If lngSysColor = vbButtonFace And IsAccess2007 = False Then  ' ** Module Function: modXAccess_07_10_Funcs.

          ' ** Get the real color number of the current system color.
140       lngRealColor = SystemColor_Get(COLOR_BTNFACE)  ' ** Function: Below.

150       Select Case lngRealColor
          Case MY_CLR_BGE
            ' ** My beige, Desert Theme.
            ' ** Leave them where they are.
160       Case 13160660
            ' ** Standard Windows gray.
170         lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         For Each ctl In frm.Controls
190           With ctl
200             If Left(.Name, 4) = "Nav_" And .ControlType = acLine Then
210               If .Name <> "Nav_hline03" Then  ' ** Doesn't apply to the Access 2007 blue line.
                    ' ** Move each of them 6 pixels left.
220                 .Left = .Left - (6& * lngTpp)
230               End If
240             End If
250           End With
260         Next
270       Case Else
            ' ** Let stand whatever's visible now, until I learn some more options.
280       End Select

290     End If

EXITP:
300     Set ctl = Nothing
310     SystemColor_Move = blnRetVal
320     Exit Function

ERRH:
330     blnRetVal = False
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Function

Public Function SystemColor_Match(frm As Access.Form, ctl1 As Access.Control, ctl2 As Access.Control) As Boolean
'VGC 05/30/2011: NO LONGER USED!
' ** Graphics meant to look like transparent GIFs actually take on the background
' ** color at the moment they're embedded, and do not change if the background
' ** changes, so I use regular BMPs and give them the background color of my machine.
' ** I have started embedding multiple copies of the image with different backgrounds,
' ** with this function determining which one matches the user's background.
' ** So far, I've only got 2. I should develop more options.
' ** The form background used throughout Trust Accountant is -2147483633,
' ** which is called the '3-D Face', or 'Button Face': vbButtonFace.
' **   On my machine:
' **     14215660 (Desert Theme Beige)
' **     R: 236  G: 233  B: 216
' **   On Rich's machine:
' **     13160660 (Windows Standard Gray)
' **     R: 212  G: 208  B: 200
' **   vbButtonFace = -2147483633
' **   SystemColor_Get(COLOR_BTNFACE) = 14215660 NO!

400   On Error GoTo ERRH

        Const THIS_PROC As String = "SystemColor_Match"

        Dim lngSysColor As Long, lngRealColor As Long
        Dim blnRetVal As Boolean

410     blnRetVal = True

420     lngSysColor = frm.Detail.BackColor

430     If lngSysColor = vbButtonFace Then
          ' ** The form is using our standard form BackColor: -2147483633.

          ' ** Get the real color number of the current system color.
440       lngRealColor = SystemColor_Get(COLOR_BTNFACE)  ' ** Function: Below.

450       Select Case lngRealColor
          Case MY_CLR_BGE
            ' ** My beige, Desert Theme.
460         Select Case frm.Name
            Case "frmAssets_Sub"

470         Case Else
480           If InStr(ctl1.Name, "_des") > 0 Then
490             ctl1.Visible = True
500             ctl2.Visible = False
510           ElseIf InStr(ctl2.Name, "_des") > 0 Then
520             ctl1.Visible = False
530             ctl2.Visible = True
540           End If
550         End Select
560       Case 13160660
            ' ** Standard Windows blue and gray.
570         Select Case frm.Name
            Case "frmAssets_Sub"

580         Case Else
590           If InStr(ctl1.Name, "_std") > 0 Then
600             ctl1.Visible = True
610             ctl2.Visible = False
620           ElseIf InStr(ctl2.Name, "_std") > 0 Then
630             ctl1.Visible = False
640             ctl2.Visible = True
650           End If
660         End Select
670       Case Else
            ' ** Let stand whatever's visible now, until I learn some more options.

680       End Select

690     End If

EXITP:
700     SystemColor_Match = blnRetVal
710     Exit Function

ERRH:
720     blnRetVal = False
730     Select Case ERR.Number
        Case Else
740       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
750     End Select
760     Resume EXITP

End Function

Public Function SystemColor_Suffix(frm As Access.Form) As String
'VGC 05/30/2011: NO LONGER USED!
' ** vbButtonFace: -2147483633
' **   My XP Style:
' **     14215660 (R: 236, G: 233, B: 216)
' **   Windows Default XP Style:
' **     13160660 (R: 212, G: 208, B: 200)
' **   Windows Classic Style:
' **                128 128 128

800   On Error GoTo ERRH

        Const THIS_PROC As String = "SystemColor_Suffix"

        Dim lngSysColor As Long, lngRealColor As Long
        Dim strRetVal As String

810     strRetVal = vbNullString

820     lngSysColor = frm.Detail.BackColor

830     If lngSysColor = vbButtonFace Then
          ' ** The form is using our standard form BackColor: -2147483633.

          ' ** Get the real color number of the current system color.
840       lngRealColor = SystemColor_Get(COLOR_BTNFACE)  ' ** Function: Below.

850       Select Case lngRealColor
          Case MY_CLR_BGE
            ' ** My beige, Desert Theme.
860         strRetVal = "_des"
870       Case 13160660
            ' ** Standard Windows gray.
880         strRetVal = "_std"
890       Case Else
            ' ** Let stand whatever's visible now, until I learn some more options.
900         strRetVal = "_des"
910       End Select

920     Else
930       strRetVal = "_std"
940     End If

EXITP:
950     SystemColor_Suffix = strRetVal
960     Exit Function

ERRH:
970     strRetVal = "_des"
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Function

Public Function SystemColor_Set(SystemElement As SYS_COLOR_VALUES, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Boolean
' *****************************************************************
' ** PURPOSE: Set system colors for a given system element
' ** PARAMETERS:
' **     SystemElement: One of values in SYS_COLOR VALUES.
' **     See Declarations
' **
' **     Red:           Red Value (0-255) of new color
' **     Green:         Green Value (0-255) of new color
' **     Blue:          Blue Value (0-255) of new color
' **
' ** RETURNS: True if successful, false otherwise
' **
' ** NOTE: Does not permanently change color;
' **       only changes the color for the session
' *****************************************************************

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "SystemColor_Set"

        Dim typCR As COLORREF
        Dim arr_typCR(0) As COLORREF
        Dim arr_lngElements(0) As Long
        'Dim lngElementsPtr As Long  NOT USED!
        'Dim lngColorRefPtr As Long  NOT USED!
        Dim lngRetVal As Long

1110    typCR.RED_VALUE = Red
1120    typCR.GREEN_VALUE = Green
1130    typCR.BLUE_VALUE = Blue
1140    arr_typCR(0) = typCR
1150    arr_lngElements(0) = SystemElement

        ' ** SetSysColors API Function requires pointers to arrays.

1160    lngRetVal = SetSysColors(1, arr_lngElements(0), ByVal (VarPtr(arr_typCR(0))))

EXITP:
1170    SystemColor_Set = lngRetVal <> 0
1180    Exit Function

ERRH:
1190    Select Case ERR.Number
        Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

End Function

Public Function SystemColor_Get(SystemElement As SYS_COLOR_VALUES) As OLE_COLOR
' *****************************************************************
' ** PURPOSE:   Get the system colors for a given system element
' ** PARAMETER:
' **            SystemElement: One of values in SYS_COLOR VALUES.
' **            See Declarations
' ** RETURNS:   Requested Color in OLE_COLOR (Long) format
' ** EXAMPLE:   Me.BackColor = SystemColor_Get(COLOR_MENUTEXT)
' *****************************************************************
1300  On Error GoTo ERRH

        Const THIS_PROC As String = "SystemColor_Get"

EXITP:
1310    SystemColor_Get = GetSysColor(SystemElement)
1320    Exit Function

ERRH:
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Function

Public Function RGBRev(varInput As Variant, blnHex As Boolean) As String
' ** RGB() function in reverse: Red, Green, Blue from color number.

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "RGBRev"

        Dim intLen As Integer
        Dim strTmp01 As String
        Dim strRetVal As String

1410    strRetVal = vbNullString

1420    If IsNull(varInput) = False Then
1430      strTmp01 = Trim(varInput)
1440      If Left(strTmp01, 1) = "#" Then strTmp01 = Mid(strTmp01, 2)
1450      intLen = Len(strTmp01)
1460      If intLen > 0 Then
            ' ** Office 2007 lists them in hex: 2-2-2, B-G-R.
1470        If blnHex = False Then strTmp01 = Hex(varInput)
1480        strTmp01 = Right("000000" & strTmp01, 6)
1490        intLen = 6
1500        If blnHex = True Then
1510          strTmp01 = Right(strTmp01, 2) & Mid(strTmp01, 3, 2) & Left(strTmp01, 2)
1520        End If
1530        strRetVal = HexX(Right(strTmp01, 2))  ' ** Module Function: mdl_Misc_Funcs.
1540        If intLen = 2 Then
1550          strRetVal = strRetVal & ",0,0"
1560        ElseIf intLen = 4 Then
1570          strRetVal = strRetVal & "," & HexX(Left(strTmp01, 2))  ' ** Module Function: mdl_Misc_Funcs.
1580          strRetVal = strRetVal & ",0"
1590        ElseIf intLen = 6 Then
1600          strRetVal = strRetVal & "," & HexX(Mid(strTmp01, 3, 2))  ' ** Module Function: mdl_Misc_Funcs.
1610          strRetVal = strRetVal & "," & HexX(Left(strTmp01, 2))  ' ** Module Function: mdl_Misc_Funcs.
1620        End If
1630      End If
1640    End If

EXITP:
1650    RGBRev = strRetVal
1660    Exit Function

ERRH:
1670    Select Case ERR.Number
        Case 13  ' ** Type mismatch.
1680      strRetVal = "#TYPE_MISMATCH"
1690    Case Else
1700      strRetVal = vbNullString
1710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1720    End Select
1730    Resume EXITP

End Function

Public Function RGB_Split(varInput As Variant, blnHex As Boolean) As String
' ** Converts color number to Red-Green-Blue constituents,
' ** given either a long integer or hex code.
' ** Parameters:
' **   Long Integer, False: 16777215 returns '255,255,255'.
' **   Hex,          True : #FFFFFF  returns '255,255,255'.
' ** See also RGBRev(), below, Hex(), internal, and Hexx(), modStringFuncs.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "RGB_Split"

        Dim intLen As Integer
        Dim strTmp01 As String
        Dim strRetVal As String

1810    strRetVal = vbNullString

1820    If IsNull(varInput) = False Then
1830      strTmp01 = Trim(varInput)
1840      If Left(strTmp01, 1) = "#" Then strTmp01 = Mid(strTmp01, 2)
1850      intLen = Len(strTmp01)
1860      If intLen > 0 Then
            ' ** Office 2007 lists them in hex: 2-2-2, B-G-R.
1870        If blnHex = False Then strTmp01 = Hex(varInput)
1880        strTmp01 = Right("000000" & strTmp01, 6)
1890        intLen = 6
1900        If blnHex = True Then
1910          strTmp01 = Right(strTmp01, 2) & Mid(strTmp01, 3, 2) & Left(strTmp01, 2)
1920        End If
1930        strRetVal = HexX(Right(strTmp01, 2))  ' ** Module Function: modStringFuncs.
1940        If intLen = 2 Then
1950          strRetVal = strRetVal & ",0,0"
1960        ElseIf intLen = 4 Then
1970          strRetVal = strRetVal & "," & HexX(Left(strTmp01, 2))  ' ** Module Function: modStringFuncs.
1980          strRetVal = strRetVal & ",0"
1990        ElseIf intLen = 6 Then
2000          strRetVal = strRetVal & "," & HexX(Mid(strTmp01, 3, 2))  ' ** Module Function: modStringFuncs.
2010          strRetVal = strRetVal & "," & HexX(Left(strTmp01, 2))  ' ** Module Function: modStringFuncs.
2020        End If
2030      End If
2040    End If

EXITP:
2050    RGB_Split = strRetVal
2060    Exit Function

ERRH:
2070    Select Case ERR.Number
        Case 13  ' ** Type mismatch.
2080      strRetVal = "#TYPE_MISMATCH"
2090    Case Else
2100      strRetVal = vbNullString
2110      Beep
2120      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()", _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2130    End Select
2140    Resume EXITP

End Function
