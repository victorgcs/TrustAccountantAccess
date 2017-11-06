VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDevices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "clsDevices"

'VGC 03/22/2017: CHANGES!

' ** From Access 2000 Developer's Handbook, Volume I
' ** by Getz, Litwin, and Gilbert. (Sybex)
' ** Copyright 1999. All rights reserved.

' ** Collection class, exposing a collection
' ** of Device objects -- all the installed
' ** printer devices.

' ** Requires:
' **   Device

' ==========================================
' == WARNING!
' ==
' == This module includes attributes that
' == cannot be set from within a VBA host.
' == In order to preserve these settings,
' == do not import this module into your
' == project from within the host.
' == ALWAYS move this from project
' == to project by exporting from VBA
' == to a text file, and them import
' == it into a new project. If you
' == simply import this into an
' == Access database, for example,
' == Accesss will strip out the import
' == information needed by this module.
' ==========================================

Private Const WM_WININICHANGE As Long = &H1A

Private Const SMTO_NORMAL As Integer = &H0

Private Const HWND_BROADCAST As Long = &HFFFF&

Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" _
  (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, _
  ByVal nSize As Long) As Long

Private Declare Function GetProfileSection Lib "kernel32.dll" Alias "GetProfileSectionA" _
  (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal lngSize As Long) As Long

Private Declare Function WriteProfileString Lib "kernel32.dll" Alias "WriteProfileStringA" _
  (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const adhcMaxSection  As Long = 10000
'Private Const adhcMaxPrinters As Long = 10000
Private Const adhcMaxSize     As Long = 2048

' ** The hidden collection of Device objects.
Private mcol As VBA.Collection
' **

Private Sub Class_Initialize()
' ** We instantiate and fill the collection
' ** when needed, not here. If someone
' ** just wants to find the default
' ** printer, there's no point
' ** doing this work.
' ** Set mcol = GetDevices()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Class_Initialize"

EXITP:
110     Exit Sub

ERRH:
120     Select Case ERR.Number
        Case Else
130       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
140     End Select
150     Resume EXITP

End Sub

Private Sub Class_Terminate()

200   On Error GoTo ERRH

        Const THIS_PROC As String = "Class_Terminate"

210     Set mcol = Nothing

EXITP:
220     Exit Sub

ERRH:
230     Select Case ERR.Number
        Case Else
240       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
250     End Select
260     Resume EXITP

End Sub

Public Function Enumerate() As IUnknown
' ** This function makes For Each...Next
' ** possible, and has a procedure
' ** attribute of -4, although you
' ** can't set that from within VBA.
' **
' ** We don't load the devices
' ** unless we need to, so the first
' ** time you call this method, you'll
' ** get the hit of reading from WIN.INI.

300   On Error GoTo ERRH

        Const THIS_PROC As String = "Enumerate"

310     If mcol Is Nothing Then
320       Set mcol = GetDevices  ' ** Function: Below.
330     End If

        ' ** If the previous call failed,
        ' ** then just get out.
340     If Not mcol Is Nothing Then
350       Set Enumerate = mcol.[_NewEnum]
360     End If

EXITP:
370     Exit Function

ERRH:
380     Select Case ERR.Number
        Case Else
390       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
400     End Select
410     Resume EXITP

End Function

Public Property Get Count() As Long
' ** Provide a normal collection's Count property.
' ** We don't load the devices
' ** unless we need to, so the first
' ** time you call this method, you'll
' ** get the hit of reading from WIN.INI.

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Count"

510     If mcol Is Nothing Then
520       Set mcol = GetDevices  ' ** Function: Below.
530     End If

        ' ** If the previous call failed,
        ' ** then return 0.
540     If mcol Is Nothing Then
550       Count = 0
560     Else
570       Count = mcol.Count
580     End If

EXITP:
590     Exit Property

ERRH:
600     Select Case ERR.Number
        Case Else
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
620     End Select
630     Resume EXITP

End Property

Public Property Get CurrentDevice() As clsDevice
' ** Retrieve a reference to the current Default printer.

700   On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentDevice"

710     Set CurrentDevice = GetCurrentDevice

EXITP:
720     Exit Property

ERRH:
730     Select Case ERR.Number
        Case Else
740       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
750     End Select
760     Resume EXITP

End Property

Public Property Set CurrentDevice(Value As clsDevice)
' ** Set the default printer, and tell
' ** the rest of the world about it.

800   On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentDevice"

810     Call SetDefaultPrinter(Value)

EXITP:
820     Exit Property

ERRH:
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
850     End Select
860     Resume EXITP

End Property

Public Property Get item(Key As Variant) As clsDevice
' ** Provide a collection's normal Item method.
' ** We don't load the devices
' ** unless we need to, so the first
' ** time you call this method, you'll
' ** get the hit of reading from WIN.INI.

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Item"

910     If mcol Is Nothing Then
920       Set mcol = GetDevices  ' ** Function: Below.
930     End If

        ' ** If the previous call failed,
        ' ** then return Nothing.
940     If mcol Is Nothing Then
950       Set item = Nothing
960     Else
970       Set item = mcol.item(Key)
980     End If

EXITP:
990     Exit Property

ERRH:
        ' ** The only error you're likely to get is
        ' ** a request for an item that doesn't exist
        ' ** (that's error 5), but it's still important.
1000    Select Case ERR.Number
        Case Else
1010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1020    End Select
1030    Resume EXITP

End Property

Private Function CurrentDeviceString() As String
' ** Get the current device name from WIN.INI. If you get back the
' ** Default string, then you have an error condition and the function
' ** returns an empty string.
' **
' ** From Access 2000 Developer's Handbook, Volume I
' ** by Getz, Litwin, and Gilbert (Sybex)
' ** Copyright 1999.  All rights reserved.

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentDeviceString"

        Dim strPrinter As String
        Dim intCount As Integer

1110    strPrinter = Space(adhcMaxSize)
1120    intCount = GetProfileString("Windows", "Device", "", strPrinter, adhcMaxSize - 1)
        ' ** "" is a 0-length string, so if
        ' ** no matches are found, return that zero-length string.
1130    CurrentDeviceString = Left(strPrinter, intCount)

EXITP:
1140    Exit Function

ERRH:
1150    Select Case ERR.Number
        Case Else
1160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1170    End Select
1180    Resume EXITP

End Function

Private Function GetCurrentDevice() As clsDevice
' ** Get the current device name from WIN.INI. Return a Device object
' ** with the default info if it exists.
' **
' ** From Access 2000 Developer's Handbook, Volume I
' ** by Getz, Litwin, and Gilbert (Sybex)
' ** Copyright 1999.  All rights reserved.

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "GetCurrentDevice"

        Dim clsDev As clsDevice
        Dim strPrinter As String
        Dim arr_strItems() As String

1210    strPrinter = CurrentDeviceString()
1220    If Len(strPrinter) > 0 Then
1230      Set clsDev = New clsDevice
1240      arr_strItems = Split(strPrinter, ",")

          ' ** If Split failed, or if there weren't three
          ' ** parts to the string, this code will
          ' ** raise a trappable error.
1250      clsDev.DeviceName = arr_strItems(0)
1260      clsDev.DriverName = arr_strItems(1)
1270      clsDev.Port = arr_strItems(2)
          ' ** If you're getting the default printer,
          ' ** then you should tell folks about it.
1280      clsDev.Default = True
1290    End If

EXITP:
1300    Set GetCurrentDevice = clsDev
1310    Exit Function

ERRH:
1320    Select Case ERR.Number
        Case Else
1330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1340    End Select
1350    Resume EXITP

End Function

Private Function GetDevices() As VBA.Collection
' ** Return a collection with all the devices found
' ** in the [devices] section in WIN.INI.
' **
' ** From Access 2000 Developer's Handbook, Volume I
' ** by Getz, Litwin, and Gilbert (Sybex)
' ** Copyright 1999.  All rights reserved.

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "GetDevices"

        Dim col As VBA.Collection
        Dim clsDev As clsDevice
        Dim strMatchstr As String
        Dim intCount As Integer
        Dim arr_strItems() As String
        Dim arr_strPrinters() As String
        Dim arr_strDetails() As String
        Dim intX As Integer

1410    Set col = New Collection
1420    strMatchstr = GetINIProfileSection("Devices")
1430    If Len(strMatchstr) = 0 Then
1440      intCount = 0
1450    Else
1460      arr_strPrinters = Split(strMatchstr, vbNullChar)
1470      ReDim aDevList(LBound(arr_strPrinters) To UBound(arr_strPrinters))
1480      For intX = LBound(arr_strPrinters) To UBound(arr_strPrinters)
1490        If Len(arr_strPrinters(intX)) > 0 Then
              ' ** varPrinters is full of strings like this:
              ' ** "HP DeskJet 890C Series=winspool,LPT1:"
              ' ** First split based on the "=", then on the ",".
1500          Set clsDev = New clsDevice
1510          arr_strItems = Split(arr_strPrinters(intX), "=")
1520          clsDev.DeviceName = arr_strItems(0)

1530          arr_strDetails = Split(arr_strItems(1), ",")
1540          clsDev.DriverName = arr_strDetails(0)
1550          clsDev.Port = arr_strDetails(1)
              ' ** I believe we're safe assuming that
              ' ** no two devices have the same name.
1560          col.Add clsDev, clsDev.DeviceName
1570        End If
1580      Next intX
1590    End If

EXITP:
1600    Set GetDevices = col
1610    Exit Function

ERRH:
1620    Select Case ERR.Number
        Case Else
1630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1640    End Select
1650    Resume EXITP

End Function

Private Function GetINIProfileSection(ByVal strGroup As String) As String
' ** Get a whole section from Win.INI. For example:
' ** strItems = GetINIProfileSection("Devices")
' ** will retrieve a list of all the devices in Win.INI,
' ** delimited with Null characters.
' **
' ** From Access 2000 Developer's Handbook, Volume I
' ** by Getz, Litwin, and Gilbert (Sybex)
' ** Copyright 1999.  All rights reserved.

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "GetINIProfileSection"

        Dim strBuffer As String
        Dim intCount As Integer

1710    strBuffer = Space(adhcMaxSection)
1720    intCount = GetProfileSection(strGroup, strBuffer, adhcMaxSection - 1)
1730    GetINIProfileSection = Left(strBuffer, intCount)

EXITP:
1740    Exit Function

ERRH:
1750    Select Case ERR.Number
        Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1770    End Select
1780    Resume EXITP

End Function

Private Sub SetDefaultPrinter(clsDev As clsDevice)
' ** Set the default printer device in Win.INI
' ** In:
' **   clsDev: a Device object to use as
' **        the source of information.
' ** Out:
' **   If successful, writes a string in the form:
' **     device=HP LaserJet 4,HPPCL5E,LPT1:
' **   to your Win.INI file.
' **
' ** From Access 2000 Developer's Handbook, Volume I
' ** by Getz, Litwin, and Gilbert (Sybex)
' ** Copyright 1999.  All rights reserved.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "SetDefaultPrinter"

        Dim strBuffer As String

        ' ** Build up the appropriate string.
1810    strBuffer = clsDev.DeviceName & "," & _
          clsDev.DriverName & "," & clsDev.Port

        ' ** Now write that string out to WIN.INI.
1820    Call WriteProfileString("Windows", "Device", strBuffer)
1830    Call SendMessageTimeoutStr(HWND_BROADCAST, WM_WININICHANGE, 0, "Windows", SMTO_NORMAL, 1000, 0)

EXITP:
1840    Exit Sub

ERRH:
1850    Select Case ERR.Number
        Case Else
1860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1870    End Select
1880    Resume EXITP

End Sub
