Attribute VB_Name = "modBrowseFilesAndFolders"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modBrowseFilesAndFolders"

'VGC 03/23/2017: CHANGES!

' ** Acknowledgements to Candace Tripp and to KPD-Team (www.allapi.net) for portions of this module

' *******************************************************
' ** Declarations for Windows Common Dialog procedures.
' *******************************************************

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
' **

Public Function FileSaveDialog(strFileType As String, strFileNameInit As String, strFilePathInit As String, strTitle1 As String) As String
' ** Calls the API File Dialog Window.
' ** Returns full path of the File to be saved.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "FileSaveDialog"

        Dim clsDialog As Object
        Dim strRetVal As String, lngRetVal As Long

110     strRetVal = vbNullString

        ' ** Call the File Common Dialog Window.
120     Set clsDialog = New clsCommonDialog

        ' ** Fill in our properties.
130     Select Case strFileType
        Case "rtf"
140       clsDialog.Filter = "Rich Text Format (*.rtf)" & Chr$(0) & "*.rtf" & Chr$(0)
          'clsDialog.Filter = clsDialog.Filter & "MS-DOS Text (*.txt)" & Chr$(0) & "*.txt" & Chr$(0)
150     Case "xls"
160       clsDialog.Filter = "Microsoft Excel (*.xls)" & Chr$(0) & "*.xls" & Chr$(0)
170     Case Else
          'clsDialog.Filter = "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)
180       clsDialog.Filter = "MDB (*.MDB)" & Chr$(0) & "*.MDB" & Chr$(0)
190       clsDialog.Filter = clsDialog.Filter & "MDE (*.MDE)" & Chr$(0) & "*.MDE" & Chr$(0)
200     End Select

210     clsDialog.hDC = 0
220     clsDialog.MaxFileSize = 256
230     clsDialog.Max = 256
240     clsDialog.FileTitle = vbNullString    ' ** File name and extension (without path information) of the selected file.
250     clsDialog.Filename = strFileNameInit  ' ** File name used to initialize the File Name edit control.
260     clsDialog.DialogTitle = strTitle1
270     clsDialog.InitDir = strFilePathInit
280     clsDialog.DefaultExt = vbNullString

290     clsDialog.flags = OFN_OVERWRITEPROMPT

300     clsDialog.CancelError = False
        'clsDialog.CancelError = True  ' ** True raises: 2001  You canceled the previous operation.

        ' ** Display the File Dialog.
310     clsDialog.ShowSave

        ' ** See if user clicked Cancel.
320     lngRetVal = clsDialog.APIReturn
330     If lngRetVal = 0& Then     ' ** User canceled.
340       strRetVal = vbNullString
350     ElseIf lngRetVal = 1 Then  ' ** User selected or entered a file.
360       strRetVal = clsDialog.Filename
370     Else
          ' ** Raise the exception.
          'ERR.Raise vbObjectError + 513, "Form_frmRelationshipViews.FileSaveDialog", _
          '  "Please Select an Existing Access Database"
          'vbObjectError = -2147221504 : Invalid OLEVERB structure.
          'vbObjectError + 513 = -2147220991 : I have no idea what this is supposed to mean!
380     End If

EXITP:
390     Set clsDialog = Nothing
400     FileSaveDialog = strRetVal
410     Exit Function

ERRH:
420     strRetVal = vbNullString
430     Select Case ERR.Number
        Case Else
440       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
450     End Select
460     Resume EXITP

End Function

Public Function GetOpenFileSIS(Optional strFilter As String, Optional strInitialDir As String, Optional strTitle1 As String) As String
' ** Comments   : Returns filename and path from the Windows Open file Dialog.
' ** Parameters : strFilter - Windows filter string such as: "All Files (*.*)", "*.*", "Text Files (*.TXT)", "*.TXT".
' **              strInitialDir - The starting directory for the dialog.
' **              strTitle1 - The title for the dialog.
' ** Returns    : Full path and filename.

500   On Error GoTo ERRH

        Const THIS_PROC As String = "GetOpenFileSIS"

        Dim typOpenFile As OPENFILENAME
        Dim strRetVal As String

510     strRetVal = vbNullString

520     With typOpenFile

530       .lStructSize = Len(typOpenFile)
          ' ** Set the parent window.
540       .hwndOwner = Application.hWndAccessApp
          ' ** Set the application's instance.
550       .hInstance = 0

          ' ** Select a filter.
560       If strFilter <> vbNullString Then
570         .lpstrFilter = strFilter & vbNullChar
580       End If

          ' ** Set the default file type.
590       .nFilterIndex = 1
          ' ** Create a buffer for the file.
600       .lpstrFile = Space$(254)
          ' ** Set the maximum length of a returned file.
610       .nMaxFile = 255
          ' ** Create a buffer for the file title.
620       .lpstrFileTitle = Space$(254)
          ' ** Set the maximum length of a returned file title.
630       .nMaxFileTitle = 255

          ' ** Set the initial directory.
640       If strInitialDir <> vbNullString Then
650         .lpstrInitialDir = strInitialDir & vbNullChar
660       Else
670         .lpstrInitialDir = CurDir() & vbNullChar
680       End If

          ' ** Get the title.
690       If strTitle1 <> vbNullString Then
700         .lpstrTitle = strTitle1 & vbNullChar
710       End If

          ' ** File and path must exist.
720       .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST

730     End With

        'lStructSize As Long
        'hwndOwner As Long
        'hInstance As Long
        'lpstrFilter As String
        'lpstrCustomFilter As String
        'nMaxCustrFilter As Long
        'nFilterIndex As Long
        'lpstrFile As String
        'nMaxFile As Long
        'lpstrFileTitle As String
        'nMaxFileTitle As Long
        'lpstrInitialDir As String
        'lpstrTitle1 As String
        'flags As Long
        'nFileOffset As Integer
        'nFileExtension As Integer
        'lpstrDefExt As String
        'lCustrData As Long
        'lpfnHook As Long
        'lpTemplateName As String

740     If GetOpenFileName(typOpenFile) Then  ' ** API Function: Above.
750       strRetVal = Left(typOpenFile.lpstrFile, InStr(typOpenFile.lpstrFile, vbNullChar) - 1)
760     End If

EXITP:
770     GetOpenFileSIS = strRetVal
780     Exit Function

ERRH:
790     strRetVal = vbNullString
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
820     End Select
830     Resume EXITP

End Function

Public Function GetFolderPathSIS(Optional strTitle1 As String) As String
' ** Comments   : Returns folder path from the Windows Browse for Folder Dialog.
' ** Parameters : strTitle1 - The title for the dialog.
' ** Returns    : Full folder path.

900   On Error GoTo ERRH

        Const THIS_PROC As String = "GetFolderPathSIS"

        Dim typBrowseInfo As BROWSEINFO
        Dim lngIDList As Long
        Dim strRetVal As String

        ' ** Specify root dir for browse for folder by constants.
        ' ** You can also specify values by constants for searhcable folders and options.
        'Const dhcCSIdlDesktop              As Integer = &H0
        'Const dhcCSIdlPrograms             As Integer = &H2
        'Const dhcCSIdlControlPanel         As Integer = &H3
        'Const dhcCSIdlInstalledPrinters    As Integer = &H4
        'Const dhcCSIdlPersonal             As Integer = &H5
        'Const dhcCSIdlFavorites            As Integer = &H6
        'Const dhcCSIdlStartupPmGroup       As Integer = &H7
        'Const dhcCSIdlRecentDocDir         As Integer = &H8
        'Const dhcCSIdlSendToItemsDir       As Integer = &H9
        'Const dhcCSIdlRecycleBin           As Integer = &HA
        'Const dhcCSIdlStartMenu            As Integer = &HB
        'Const dhcCSIdlDesktopDirectory     As Integer = &H10
        Const dhcCSIdlMyComputer           As Integer = &H11
        'Const dhcCSIdlNetworkNeighborhood  As Integer = &H12
        'Const dhcCSIdlNetHoodFileSystemDir As Integer = &H13
        'Const dhcCSIdlFonts                As Integer = &H14
        'Const dhcCSIdlTemplates            As Integer = &H15

        ' ** Constants for limiting choices for BrowseForFolder Dialog.
        'Const dhcBifReturnAll                As Integer = &H0
        'Const dhcBifReturnOnlyFileSystemDirs As Integer = &H1
        'Const dhcBifDontGoBelowDomain        As Integer = &H2
        'Const dhcBifIncludeStatusText        As Integer = &H4
        'Const dhcBifSystemAncestors          As Integer = &H8
        'Const dhcBifBrowseForComputer        As Integer = &H1000
        'Const dhcBifBrowseForPrinter         As Integer = &H2000

        'THERE APPEARS TO BE NO WAY TO SPECIFY THE STARTING FOLDER,
        'OTHER THAN THE CONSTANTS LISTED ABOVE!
        'BROWSEINFO:
        '  hwndOwner As Long
        '  pIDLRoot As Long
        '  pszDisplayName As Long
        '  lpszTitle As Long
        '  ulFlags As Long
        '  lpfnCallback As Long
        '  lParam As Long
        '  iImage As Long

910     strRetVal = vbNullString

920     With typBrowseInfo
          ' ** Set the owner window.
930       .hwndOwner = Application.hWndAccessApp
          ' ** lstrcat appends the two strings and returns the memory address.
940       If Len(strTitle1) > 0 Then
950         .lpszTitle = lstrcat(strTitle1, "")  ' ** API Function: Above.
960       End If
          ' ** Return only if the user selected a directory.
970       .ulFlags = BIF_RETURNONLYFSDIRS
          ' ** Start in MyComputer.
980       .pidlRoot = dhcCSIdlMyComputer
990     End With

        ' ** Show the 'Browse for folder' dialog.
1000    lngIDList = SHBrowseForFolder(typBrowseInfo)  ' ** API Function: Above.
1010    If lngIDList Then
1020      strRetVal = String$(BIF_MAXPATH, 0)
          ' ** Get the path from the IDList.
1030      SHGetPathFromIDList lngIDList, strRetVal  ' ** API Function: Above.
          ' ** Free the block of memory.
1040      CoTaskMemFree lngIDList  ' ** API Function: Above.
1050      strRetVal = Left(strRetVal, InStr(strRetVal, vbNullChar) - 1)
1060    End If

EXITP:
1070    GetFolderPathSIS = strRetVal
1080    Exit Function

ERRH:
1090    strRetVal = vbNullString
1100    Select Case ERR.Number
        Case Else
1110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1120    End Select
1130    Resume EXITP

End Function

Public Function CreateWindowsFilterString(ParamArray varFilt() As Variant) As String
' ** Comments   : Builds a Windows formatted filter string for 'file type'.
' ** Parameters : varFilter - parameter array in the format: Text, Filter, Text, Filter ...
' **              Such as: "All Files (*.*)", "*.*", "Text Files (*.TXT)", "*.TXT"
' ** Returns    : Windows formatted filter string.

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "CreateWindowsFilterString"

        Dim intCount As Integer
        Dim intParamCount As Integer
        Dim strRetVal As String

        ' ** Get the count of paramaters passed to the function.
1210    intParamCount = UBound(varFilt)

1220    If (intParamCount <> -1) Then  ' ** If no elements in array.
          ' ** Count through each parameter.
1230      For intCount = 0 To intParamCount
1240        strRetVal = strRetVal & varFilt(intCount) & vbNullChar
1250      Next
          ' ** Check for an even number of parameters.
1260      If (intParamCount Mod 2) = 0 Then  ' ** Because it's a zero-based array, division by 2 means it's an odd number.
1270        strRetVal = strRetVal & "*.*" & vbNullChar
1280      End If
1290    End If

EXITP:
1300    CreateWindowsFilterString = strRetVal
1310    Exit Function

ERRH:
1320    strRetVal = vbNullString
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1350    End Select
1360    Resume EXITP

End Function
