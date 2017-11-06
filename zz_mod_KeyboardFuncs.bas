Attribute VB_Name = "modKeyboardFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modKeyboardFuncs"

'VGC 08/21/2011: CHANGES!

Private Declare Sub KeyboardEvent Lib "user32.dll" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function GetKeyboardState Lib "user32.dll" (pbKeyState As Byte) As Long
' ** Because the SetKeyboardState function alters the input state of the calling thread and
' ** not the global input state of the system, an application cannot use SetKeyboardState
' ** to set the NUM LOCK, CAPS LOCK, or SCROLL LOCK (or the Japanese KANA) indicator lights
' ** on the keyboard. These can be set or cleared using SendInput to simulate keystrokes.
' **
' ** Windows NT/2000/XP: The keybd_event function can also toggle the NUM LOCK, CAPS LOCK,
' ** and SCROLL LOCK keys.
' **
' ** Windows 95/98/Me: The keybd_event function can toggle only the CAPS LOCK
' ** and SCROLL LOCK keys. It cannot toggle the NUM LOCK key.

Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Long

Private Declare Function SetKeyboardState Lib "user32.dll" (lppbKeyState As Byte) As Long

' ** Constant declarations:
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Private Const VK_CAPITAL = &H14  '20
Private Const VK_NUMLOCK = &H90  '144
Private Const VK_SCROLL = &H91   '145  HexX("&H91")

Private Const KY_CAPITAL    As Integer = 20
Private Const KY_NUMLOCK    As Integer = 144
Private Const KY_SCROLLLOCK As Integer = 145

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
' **

Public Function KeyMon_Fix(blnTurnOn As Boolean, lngPlatformID As Long) As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "KeyMon_Fix"

        Dim blnNumLockState As Boolean
        Dim arr_bytKeys(0 To 255) As Byte
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     GetKeyboardState arr_bytKeys(0)  ' ** API Function: Above.

        ' ** NumLock handling:
130     blnNumLockState = arr_bytKeys(VK_NUMLOCK)
140     If blnTurnOn <> blnNumLockState Then
150       If lngPlatformID = VER_PLATFORM_WIN32_WINDOWS Then  ' ** === Win95/98
160         arr_bytKeys(VK_NUMLOCK) = 1
170         SetKeyboardState arr_bytKeys(0)  ' ** API Function: Above.
180       ElseIf lngPlatformID = VER_PLATFORM_WIN32_NT Then   ' ** === WinNT  ####  THIS IS XP & VISTA  ####
            ' ** Simulate Key Press.
190         KeyboardEvent VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0  ' ** API Procedure: Above.
            ' ** Simulate Key Release.
200         KeyboardEvent VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0  ' ** API Procedure: Above.
210       End If
220     End If

EXITP:
230     KeyMon_Fix = blnRetVal
240     Exit Function

ERRH:
250     blnRetVal = False
260     Select Case ERR.Number
        Case Else
270       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
280     End Select
290     Resume EXITP

End Function

Public Function KeyMon_Write(arr_varKey As Variant, lngKeyCode As Long, strFormName As String) As Boolean

300   On Error GoTo ERRH

        Const THIS_PROC As String = "KeyMon_Write"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngKeys As Long
        Dim lngFrmID As Long, lngCtlID As Long, lngEvtID As Long
        Dim lngCtls As Long, arr_varCtl() As Variant
        Dim lngEvts As Long, arr_varEvt() As Variant
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim varTmp0 As Variant
        Dim blnRetVal As Boolean

        ' ** Array: arr_varKey().
        Const KEY_CTL As Integer = 0
        Const KEY_EVT As Integer = 1
        Const KEY_VAL As Integer = 2
        Const KEY_DAT As Integer = 3

        ' ** Array: arr_varCtl().
        Const CT_ELEMS As Integer = 1  ' ** Array's first-element UBound.
        Const CT_ID  As Integer = 0
        Const CT_NAM As Integer = 1

        ' ** Array: arr_varEvt().
        Const EV_ELEMS As Integer = 1  ' ** Array's first-element UBound.
        Const EV_ID  As Integer = 0
        Const EV_EXT As Integer = 1

310     blnRetVal = True

320     If IsEmpty(arr_varKey) = False Then
330       lngFrmID = DLookup("[frm_id]", "tblForm", "[frm_name] = '" & strFormName & "'")
340       lngCtls = 0&
350       ReDim arr_varCtl(CT_ELEMS, 0)
360       lngEvts = 0&
370       ReDim arr_varEvt(EV_ELEMS, 0)
380       Set dbs = CurrentDb
390       With dbs
            ' ********************************************************
            ' ** Array: arr_varKey()
            ' **
            ' **   Field  Element  Name                   Constant
            ' **   =====  =======  =====================  ==========
            ' **     1       0     keybrd_control         KEY_CTL
            ' **     2       1     keybrd_event           KEY_EVT
            ' **     3       2     keybrd_boolean         KEY_VAL
            ' **     4       3     keybrd_datemodified    KEY_DAT
            ' **
            ' ********************************************************
400         Set rst = .OpenRecordset("tblKeyboard_Monitor", dbOpenDynaset, dbAppendOnly)
410         With rst
420           lngKeys = UBound(arr_varKey, 2) + 1&
430           For lngX = 0& To (lngKeys - 1&)
440             .AddNew
450             ![keycode_value] = vbKeyNumlock
460             ![frm_id] = lngFrmID
                ' ** Use [NumLock_Monitor_chk] for 'Form' (coming from the Form_KeyDown() event).
470             If arr_varKey(KEY_CTL, lngX) = "Form" Then arr_varKey(KEY_CTL, lngX) = "NumLock_Monitor_chk"
480             lngCtlID = 0&
490             For lngY = 0& To (lngCtls - 1&)
500               If arr_varCtl(CT_NAM, lngY) = arr_varKey(KEY_CTL, lngX) Then
510                 lngCtlID = arr_varCtl(CT_ID, lngY)
520                 Exit For
530               End If
540             Next
550             If lngCtlID = 0& Then
560               varTmp0 = DLookup("[ctl_id]", "tblForm_Control", "[frm_id] = " & CStr(lngFrmID) & " And " & _
                    "[ctl_name] = '" & arr_varKey(KEY_CTL, lngX) & "'")
                  'If IsNull(varTmp0) = True Then
                  '  Stop
                  'End If
570               lngCtlID = varTmp0
580               lngCtls = lngCtls + 1&
590               lngE = lngCtls - 1&
600               ReDim Preserve arr_varCtl(CT_ELEMS, lngE)
610               arr_varCtl(CT_ID, lngE) = lngCtlID
620               arr_varCtl(CT_NAM, lngE) = arr_varKey(KEY_CTL, lngX)
630             End If
640             ![ctl_id] = lngCtlID
650             ![keybrd_control] = arr_varKey(KEY_CTL, lngX)
660             lngEvtID = 0&
670             For lngY = 0& To (lngEvts - 1&)
680               If arr_varCtl(EV_EXT, lngY) = arr_varKey(KEY_EVT, lngX) Then
690                 lngEvtID = arr_varCtl(EV_ID, lngY)
700                 Exit For
710               End If
720             Next
730             If lngEvtID = 0& Then
740               varTmp0 = DLookup("[vbcom_event_id]", "tblVBComponent_Event", "[vbcom_event_ext] = '" & arr_varKey(KEY_EVT, lngX) & "'")
750               lngEvtID = varTmp0
760               lngEvts = lngEvts + 1&
770               lngE = lngEvts - 1&
780               ReDim Preserve arr_varEvt(EV_ELEMS, lngE)
790               arr_varCtl(EV_ID, lngE) = lngEvtID
800               arr_varCtl(EV_EXT, lngE) = arr_varKey(KEY_EVT, lngX)
810             End If
820             ![vbcom_event_id] = lngEvtID
830             ![keybrd_event] = arr_varKey(KEY_EVT, lngX)
840             ![datatype_db_type] = dbBoolean
850             ![keybrd_boolean] = arr_varKey(KEY_VAL, lngX)
860             ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
870             ![keybrd_user] = GetUserName  ' ** Module Function: modFileUtilities.
880             ![keybrd_datemodified] = arr_varKey(KEY_DAT, lngX)
890             .Update
900           Next
910           .Close
920         End With
930         .Close
940       End With
950     End If

EXITP:
960     Set rst = Nothing
970     Set dbs = Nothing
980     KeyMon_Write = blnRetVal
990     Exit Function

ERRH:
1000    blnRetVal = False
1010    Select Case ERR.Number
        Case Else
1020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1030    End Select
1040    Resume EXITP

End Function

Public Function dwPlatformId_Get() As Long
' ** 0  VER_PLATFORM_WIN32s
' ** 1  VER_PLATFORM_WIN32_WINDOWS
' ** 2  VER_PLATFORM_WIN32_NT

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "dwPlatformId_Get"

        Dim o As OSVERSIONINFO
        Dim lngRetVal As Long

1110    o.dwOSVersionInfoSize = Len(o)
1120    GetVersionEx o  ' ** API Function: modWindowFunctions.
1130    lngRetVal = o.dwPlatformId

EXITP:
1140    dwPlatformId_Get = lngRetVal
1150    Exit Function

ERRH:
1160    Select Case ERR.Number
        Case Else
1170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1180    End Select
1190    Resume EXITP

End Function

Public Function IsCapLock(lngPlatformID As Long) As Boolean

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "IsCapLock"

        Dim blnRetVal As Boolean

1210    blnRetVal = KeyState_Get(KY_CAPITAL, lngPlatformID)  ' ** Function: Below.

EXITP:
1220    IsCapLock = blnRetVal
1230    Exit Function

ERRH:
1240    blnRetVal = False
1250    Select Case ERR.Number
        Case Else
1260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1270    End Select
1280    Resume EXITP

End Function

Public Function IsNumLock(lngPlatformID As Long) As Boolean

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "IsNumLock"

        Dim blnRetVal As Boolean

1310    blnRetVal = KeyState_Get(KY_NUMLOCK, lngPlatformID)  ' ** Function: Below.

EXITP:
1320    IsNumLock = blnRetVal
1330    Exit Function

ERRH:
1340    blnRetVal = False
1350    Select Case ERR.Number
        Case Else
1360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1370    End Select
1380    Resume EXITP

End Function

Public Function IsScrollLock(lngPlatformID As Long) As Boolean

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "IsScrollLock"

        Dim blnRetVal As Boolean

1410    blnRetVal = KeyState_Get(KY_SCROLLLOCK, lngPlatformID)  ' ** Function: Below.

EXITP:
1420    IsScrollLock = blnRetVal
1430    Exit Function

ERRH:
1440    blnRetVal = False
1450    Select Case ERR.Number
        Case Else
1460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1470    End Select
1480    Resume EXITP

End Function

Private Function KeyState_Get(lngKey As Long, lngPlatformID As Long) As Boolean

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "KeyState_Get"

        Dim arr_bytKeys(0 To 255) As Byte
        Dim blnRetVal As Boolean

1510    GetKeyboardState arr_bytKeys(0)  ' ** API Function: Above.

1520    If lngPlatformID = VER_PLATFORM_WIN32_NT Then           ' ** === WinNT     ***** XP HERE *****
1530      blnRetVal = CBool(arr_bytKeys(lngKey))
1540    ElseIf lngPlatformID = VER_PLATFORM_WIN32_WINDOWS Then  ' ** === Win95/98
1550      blnRetVal = GetKeyState(lngKey) And 1 = 1  ' ** API Function: Above.
1560    End If

EXITP:
1570    KeyState_Get = blnRetVal
1580    Exit Function

ERRH:
1590    blnRetVal = False
1600    Select Case ERR.Number
        Case Else
1610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1620    End Select
1630    Resume EXITP

End Function
