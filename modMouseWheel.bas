Attribute VB_Name = "modMouseWheel"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modMouseWheel"

'VGC 03/19/2017: CHANGES!

' ** NOTE:
' ** REACTIVATED 12/25/2008!
' ** {The functions in this module are no longer called as of v2.1.46
' ** because of problems some users have had. Since they really only
' ** apply to Access 97, well, it's about time they upgraded! VGC}

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

Private Declare Function StopMouseWheel Lib "MouseHook.dll" _
  (ByVal hwnd As Long, ByVal AccessThreadID As Long, _
  Optional ByVal bNoSubformScroll As Boolean = False, Optional ByVal blIsGlobal As Boolean = False) As Boolean

Private Declare Function StartMouseWheel Lib "MouseHook.dll" (ByVal hwnd As Long) As Boolean

Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long

' ** Instance returned from LoadLibrary call.
Private hLib As Long
' **

Public Function MouseWheelOFF(Optional varNoSubFormScroll As Variant = False, Optional varGlobalHook As Variant = False) As Boolean
' ** The Application.FileSearch was removed in Office 2007.
' ** If accessed, this property will return an error. To work around this issue, use
' ** the FileSystemObject to recursively search directories to find specific files.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "MouseWheelOFF"

        Dim fso As Scripting.FileSystemObject, fsfds As Scripting.Folders, fsfd As Scripting.Folder
        Dim strMsg As String, strCurAppPath As String, strCurDataPath As String
        Dim AccessThreadID As Long
        Dim strPath As String
        Dim blnFound As Boolean
        Dim blnRetVal As Boolean

110     blnRetVal = True

        ' ** Our error string.
120     strMsg = "Sorry... cannot find the MouseHook.dll file." & vbCrLf
130     strMsg = strMsg & "Please copy the MouseHook.dll file to your Windows System folder or into the same folder as this Access MDB."

        ' ** OK Try to load the DLL assuming it is in the Window System folder.
140     hLib = LoadLibrary("MouseHook.dll")  ' ** API Function: Above.
150     If hLib = 0 Then
          ' ** See if the DLL is in the same folder as this MDB.
          ' ** CurrentDB works with both A97 and A2K or higher.
160       hLib = LoadLibrary(CurrentAppPath & LNK_SEP & "MouseHook.dll")  ' ** API Function: Above; Module Function: modFileUtilities.
170       If hLib = 0 Then
180         blnFound = False: strPath = vbNullString
190         Set fso = CreateObject("Scripting.FileSystemObject")
            ' ** Look for a copy of MouseHook.dll in the current directory.
200         With fso
210           If Not .FileExists(CurrentAppPath & LNK_SEP & "MouseHook.dll") Then  ' ** Module Function: modFileUtilities.
                ' ** If not found, search for MouseHook.dll elsewhere.

220             strCurAppPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
230             strCurDataPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.

240             Set fsfd = .GetFolder(strCurAppPath)
250             Set fsfds = fsfd.SubFolders
260             blnRetVal = FSO_Folders(fsfds)  ' ** Function: Below.

270             strPath = gstrReportCallingForm  ' ** Borrowing this variable.
280             If blnRetVal = True And strPath <> vbNullString Then
290               blnFound = True
300               gstrTrustDataLocation = strPath  'OOPS! strPath ISN'T PUBLIC!  OH RATS, JUST BORROW ONE!
310             Else
320               blnRetVal = False
330             End If

340           End If
350         End With  ' ** fso.

360         If blnFound = True Then
370           hLib = LoadLibrary(strPath & LNK_SEP & "MouseHook.dll")  ' ** API Function: Above; Module Function: modFileUtilities.
380           If hLib = 0 Then
390             blnRetVal = False
400             MsgBox strMsg, vbExclamation + vbOKOnly, "Missing MouseHook.dll File"
410           End If
420         End If
430       End If
440     End If

        ' ** Get the ID for this thread.
450     If blnRetVal = True Then
460       AccessThreadID = GetCurrentThreadId  ' ** API Function: Above
          ' ** Call our MouseHook function in the MouseHook dll.
          ' ** Please not the Optional varGlobalHook BOOLEAN parameter
          ' ** Several developers asked for the MouseHook to be able to work with
          ' ** multiple instances of Access. In order to accomodate this request I
          ' ** have modified the function to allow the caller to
          ' ** specify a thread specific(this current instance of Access only) or
          ' ** a global(all applications) MouseWheel Hook.
          ' ** Only use the varGlobalHook if you will be running multiple instances of Access!
470   On Error Resume Next
480       blnRetVal = StopMouseWheel(Application.hWndAccessApp, AccessThreadID, varNoSubFormScroll, varGlobalHook)  ' ** API Function: Above
490       If ERR.Number <> 0 Then
500         Select Case ERR.Number
            Case 53  ' ** File not found.
              ' ** We should've already known this!
              ' ** Just let it go...
510         Case Else
520           zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530         End Select
540   On Error GoTo ERRH
550       Else
560   On Error GoTo ERRH
570       End If
580     End If

EXITP:
590     Set fsfds = Nothing
600     Set fsfd = Nothing
610     Set fso = Nothing
620     MouseWheelOFF = blnRetVal
630     Exit Function

ERRH:
640     blnRetVal = False
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Function

Public Function MouseWheelON() As Boolean

700   On Error GoTo ERRH

        Const THIS_PROC As String = "MouseWheelON"

        'Dim fso As Scripting.FileSystemObject
        'Dim strPath As String
        'Dim blnFound As Boolean
        Dim blnRetVal As Boolean

710   On Error Resume Next
720     blnRetVal = StartMouseWheel(Application.hWndAccessApp)  ' ** API Function: Above
730     If ERR.Number <> 0 Then
740       Select Case ERR.Number
          Case 53  ' ** File not found.
            ' ** Just let it go...
            'blnFound = False: strPath = vbNullString
            'If blnFound = False Then
            '  Set fso = CreateObject("Scripting.FileSystemObject")
            '  ' ** Look for a copy of MouseHook.dll in the current directory.
            '  If Not fso.FileExists(CurrentAppPath & LNK_SEP & "MouseHook.dll") Then  ' ** Module Function: modFileUtilities.
            '    ' ** If not found, search for MouseHook.dll elsewhere.
            '    With Application.FileSearch
            '      .newsearch
            '      '.LookIn = CurrentAppPath  ' ** Module Function: modFileUtilities.
            '      .SearchSubFolders = True
            '      .FileName = "MouseHook.dll"
            '      .MatchTextExactly = True
            '      ' ** If we find one or more, assume the first one is good.
            '      If .Execute() > 0 Then
            '        blnFound = True
            '        strPath = Parse_Path(.FoundFiles(1))  ' ** Module Function: modFileUtilities.
            '      End If
            '    End With
            '  End If
            'End If
            'If blnFound = True Then
            '  ' ** Register it?
            '  'regsvr32 <path & filename of dll or ocx>
            'Else
            '  ' **
            'End If
750       Case Else
760         zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770       End Select
780   On Error GoTo ERRH
790     Else
800   On Error GoTo ERRH
810     End If

820     If hLib <> 0 Then
830       hLib = FreeLibrary(hLib)
840     End If

EXITP:
        'Set fso = Nothing
850     MouseWheelON = blnRetVal
860     Exit Function

ERRH:
870     blnRetVal = False
880     Select Case ERR.Number
        Case Else
890       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
900     End Select
910     Resume EXITP

End Function
