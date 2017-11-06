Attribute VB_Name = "modWindowFunctions"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modWindowFunctions"

'VGC 03/23/2017: CHANGES!

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal wRevert As Long) As Long

Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

'Public Type SECURITY_ATTRIBUTES
'  nLength As Long
'  lpSecurityDescriptor As Long
'  bInheritHandle As Long
'End Type

Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

' ** ShowWindow() Commands.
'Public Const SW_HIDE            As Integer = 0
Public Const SW_SHOWNORMAL      As Integer = 1
'Public Const SW_NORMAL          As Integer = 1
Public Const SW_SHOWMINIMIZED   As Integer = 2
Public Const SW_SHOWMAXIMIZED   As Integer = 3
'Public Const SW_MAXIMIZE        As Integer = 3
'Public Const SW_SHOWNOACTIVATE  As Integer = 4
Public Const SW_SHOW            As Integer = 5
'Public Const SW_MINIMIZE        As Integer = 6
'Public Const SW_SHOWMINNOACTIVE As Integer = 7
'Public Const SW_SHOWNA          As Integer = 8
'Public Const SW_RESTORE         As Integer = 9
'Public Const SW_SHOWDEFAULT     As Integer = 10
'Public Const SW_MAX             As Integer = 10

' ** GetScreenRes() options.
'Public Const GSR_DRIVERVERSION  As Long = 0&     ' ** Device driver version
'Public Const GSR_TECHNOLOGY     As Long = 2&     ' ** Device classification
'Public Const GSR_HORZSIZE       As Long = 4&     ' ** Horizontal size in millimeters
'Public Const GSR_VERTSIZE       As Long = 6&     ' ** Vertical size in millimeters
Public Const GSR_HORZRES        As Long = 8&     ' ** Horizontal width in pixels
Public Const GSR_VERTRES        As Long = 10&    ' ** Vertical height in pixels
Public Const GSR_BITSPIXEL      As Long = 12&    ' ** Number of bits per pixel
'Public Const GSR_PLANES         As Long = 14&    ' ** Number of planes
'Public Const GSR_NUMBRUSHES     As Long = 16&    ' ** Number of brushes the device has
'Public Const GSR_NUMPENS        As Long = 18&    ' ** Number of pens the device has
'Public Const GSR_NUMMARKERS     As Long = 20&    ' ** Number of markers the device has
'Public Const GSR_NUMFONTS       As Long = 22&    ' ** Number of fonts the device has
'Public Const GSR_NUMCOLORS      As Long = 24&    ' ** Number of colors the device supports
'Public Const GSR_PDEVICESIZE    As Long = 26&    ' ** Size required for device descriptor
'Public Const GSR_CURVECAPS      As Long = 28&    ' ** Curve capabilities
'Public Const GSR_LINECAPS       As Long = 30&    ' ** Line capabilities
'Public Const GSR_POLYGONALCAPS  As Long = 32&    ' ** Polygonal capabilities
'Public Const GSR_TEXTCAPS       As Long = 34&    ' ** Text capabilities
'Public Const GSR_CLIPCAPS       As Long = 36&    ' ** Clipping capabilities
'Public Const GSR_RASTERCAPS     As Long = 38&    ' ** Bitblt capabilities
Public Const GSR_ASPECTX        As Long = 40&    ' ** Length of the X leg
Public Const GSR_ASPECTY        As Long = 42&    ' ** Length of the Y leg
Public Const GSR_ASPECTXY       As Long = 44&    ' ** Length of the hypotenuse
'Public Const GSR_SHADEBLENDCAPS As Long = 45&    ' ** Shading and blending caps (IE5)
Public Const GSR_LOGPIXELSX     As Long = 88&    ' ** Logical pixels/inch in X
Public Const GSR_LOGPIXELSY     As Long = 90&    ' ** Logical pixels/inch in Y
'Public Const GSR_SIZEPALETTE    As Long = 104&   ' ** Number of entries in physical palette
'Public Const GSR_NUMRESERVED    As Long = 106&   ' ** Number of reserved entries in palette
'Public Const GSR_COLORRES       As Long = 108&   ' ** Actual color resolution
Public Const GSR_VREFRESH       As Long = 116&   ' ** Current vertical refresh rate of the display device (for displays only) in Hz
Public Const GSR_DESKTOPVERTRES As Long = 117&   ' ** Horizontal width of entire desktop in pixels (NT5)
Public Const GSR_DESKTOPHORZRES As Long = 118&   ' ** Vertical height of entire desktop in pixels (NT5)
'Public Const GSR_BLTALIGNMENT   As Long = 119&   ' ** Preferred blt alignment

' ** SetWindowPos() Constants.
Public Const SWP_NOSIZE     As Long = &H1
Public Const SWP_NOMOVE     As Long = &H2
Public Const SWP_NOZORDER   As Long = &H4  ' ** Ignores the hWndInsertAfter.
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_SHOWWINDOW As Long = &H40
Public Const TOPMOST_FLAGS  As Long = (SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE)

' ** Window field offsets for SetWindowLong()/GetWindowLong().
'Public Const GWL_WNDPROC    As Long = (-4&)
Public Const GWL_HINSTANCE  As Long = (-6&)
'Public Const GWL_HWNDPARENT As Long = (-8&)
'Public Const GWL_ID         As Long = (-12&)
Public Const GWL_STYLE      As Long = (-16&)
'Public Const GWL_EXSTYLE    As Long = (-20&)
'Public Const GWL_USERDATA   As Long = (-21&)

' ** GetWindow() Constants.
Public Const GW_HWNDFIRST As Integer = 0
'Public Const GW_HWNDLAST  As Integer = 1
Public Const GW_HWNDNEXT  As Integer = 2
'Public Const GW_HWNDPREV  As Integer = 3
'Public Const GW_OWNER     As Integer = 4
Public Const GW_CHILD     As Integer = 5
'Public Const GW_MAX       As Integer = 5

' ** Constants for API memory functions.
Public Const GMEM_MOVEABLE As Long = &H2
Public Const GMEM_ZEROINIT As Long = &H40
Public Const GMEM_HND      As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

' ** Windows Message Constant.
Public Const WM_VSCROLL As Long = &H115
Public Const WM_HSCROLL As Long = &H114
Public Const WM_GETTEXT As Long = &HD

' ** GetWindowRect() Constants.
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type POINTAPI
  X As Long
  Y As Long
End Type

' ** The actual Date/Time values are stored this way.
Public Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Public Type MMTIME
  wType As Long
  Units As Long
  smpteVal As Long
  songPtrPos As Long
End Type

Public Const TIME_MS As Long = 1&

Public Type MSG
  hwnd As Long
  Message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type

'Public Type SMPTE
'  hour As Byte
'  min As Byte
'  Sec As Byte
'  frame As Byte
'  fps As Byte
'  dummy As Byte
'  Pad(2) As Byte
'End Type

' ** Used with ChooseColor() API.
Public Type COLORSTRUC
  lStructSize As Long
  hwnd As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Public Const CC_SOLIDCOLOR As Long = &H80

' ** Constants for color dialog.
' ** Right now, these can remain Private.
'Private Const CDERR_DIALOGFAILURE   As Long = &HFFFF
'Private Const CDERR_FINDRESFAILURE  As Long = &H6
'Private Const CDERR_GENERALCODES    As Long = &H0
'Private Const CDERR_INITIALIZATION  As Long = &H2
'Private Const CDERR_LOADRESFAILURE  As Long = &H7
'Private Const CDERR_LOADSTRFAILURE  As Long = &H5
'Private Const CDERR_LOCKRESFAILURE  As Long = &H8
'Private Const CDERR_MEMALLOCFAILURE As Long = &H9
'Private Const CDERR_MEMLOCKFAILURE  As Long = &HA
'Private Const CDERR_NOHINSTANCE     As Long = &H4
'Private Const CDERR_NOHOOK          As Long = &HB
'Private Const CDERR_NOTEMPLATE      As Long = &H3
'Private Const CDERR_REGISTERMSGFAIL As Long = &HC
'Private Const CDERR_STRUCTSIZE      As Long = &H1

'Public Type COLORREF
'  RED_VALUE As Byte
'  GREEN_VALUE As Byte
'  BLUE_VALUE As Byte
'End Type

' ** Color Types.
' ** Right now, these can remain Private.
'Private Const CTLCOLOR_MSGBOX    As Integer = 0
'Private Const CTLCOLOR_EDIT      As Integer = 1
'Private Const CTLCOLOR_LISTBOX   As Integer = 2
'Private Const CTLCOLOR_BTN       As Integer = 3
'Private Const CTLCOLOR_DLG       As Integer = 4
'Private Const CTLCOLOR_SCROLLBAR As Integer = 5
'Private Const CTLCOLOR_STATIC    As Integer = 6
'Private Const CTLCOLOR_MAX       As Integer = 8  ' ** Three bits max.

Public Enum SYS_COLOR_VALUES
  COLOR_SCROLLBAR = 0
  COLOR_BACKGROUND = 1
  COLOR_ACTIVECAPTION = 2
  COLOR_INACTIVECAPTION = 3
  COLOR_MENU = 4
  COLOR_WINDOW = 5
  COLOR_WINDOWFRAME = 6
  COLOR_MENUTEXT = 7
  COLOR_WINDOWTEXT = 8
  COLOR_CAPTIONTEXT = 9
  COLOR_ACTIVEBORDER = 10
  COLOR_INACTIVEBORDER = 11
  COLOR_APPWORKSPACE = 12
  COLOR_HIGHLIGHT = 13
  COLOR_HIGHLIGHTTEXT = 14
  COLOR_BTNFACE = 15
  COLOR_BTNSHADOW = 16
  COLOR_GRAYTEXT = 17
  COLOR_BTNTEXT = 18
  COLOR_INACTIVECAPTIONTEXT = 19
  COLOR_BTNHIGHLIGHT = 20
  COLOR_3DDKSHADOW = 21
  COLOR_3DLIGHT = 22
  COLOR_INFOTEXT = 23
  COLOR_INFOBK = 24
  COLOR_STATIC = 25
  COLOR_STATICTEXT = 26
  COLOR_GRADIENTACTIVECAPTION = 27
  COLOR_GRADIENTINACTIVECAPTION = 28
  COLOR_MENUHILIGHT = 29
  COLOR_MENUBAR = 30
End Enum

' ** Structure used by GetOpenFileName() and GetSaveFileName().
Public Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustrFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustrData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

' ** lStructSize
' **   Type: DWORD
' **   The length, in bytes, of the structure. Use sizeof (OPENFILENAME) for this parameter.
' ** hwndOwner
' **   Type: HWND
' **   A handle to the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner.
' ** hInstance
' **   Type: HINSTANCE
' **   If the OFN_ENABLETEMPLATEHANDLE flag is set in the Flags member, hInstance is a handle to a memory object containing a dialog box template. If the OFN_ENABLETEMPLATE flag is set, hInstance is a handle to a module that contains a dialog box template named by the lpTemplateName member. If neither flag is set, this member is ignored. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
' ** lpstrFilter
' **   Type: LPCTSTR
' **   A buffer containing pairs of null-terminated filter strings. The last string in the buffer must be terminated by two NULL characters.
' **   The first string in each pair is a display string that describes the filter (for example, "Text Files"), and the second string specifies the filter pattern (for example, "*.TXT"). To specify multiple filter patterns for a single display string, use a semicolon to separate the patterns (for example, "*.TXT;*.DOC;*.BAK"). A pattern string can be a combination of valid file name characters and the asterisk (*) wildcard character. Do not include spaces in the pattern string.
' **   The system does not change the order of the filters. It displays them in the File Types combo box in the order specified in lpstrFilter.
' **   If lpstrFilter is NULL, the dialog box does not display any filters.
' **   In the case of a shortcut, if no filter is set, GetOpenFileName and GetSaveFileName retrieve the name of the .lnk file, not its target. This behavior is the same as setting the OFN_NODEREFERENCELINKS flag in the Flags member. To retrieve a shortcut's target without filtering, use the string "All Files\0*.*\0\0".
' ** lpstrCustomFilter
' **   Type: LPTSTR
' **   A static buffer that contains a pair of null-terminated filter strings for preserving the filter pattern chosen by the user. The first string is your display string that describes the custom filter, and the second string is the filter pattern selected by the user. The first time your application creates the dialog box, you specify the first string, which can be any nonempty string. When the user selects a file, the dialog box copies the current filter pattern to the second string. The preserved filter pattern can be one of the patterns specified in the lpstrFilter buffer, or it can be a filter pattern typed by the user. The system uses the strings to initialize the user-defined file filter the next time the dialog box is created. If the nFilterIndex member is zero, the dialog box uses the custom filter.
' **   If this member is NULL, the dialog box does not preserve user-defined filter patterns.
' **   If this member is not NULL, the value of the nMaxCustFilter member must specify the size, in characters, of the lpstrCustomFilter buffer.
' ** nMaxCustFilter
' **   Type: DWORD
' **   The size, in characters, of the buffer identified by lpstrCustomFilter. This buffer should be at least 40 characters long. This member is ignored if lpstrCustomFilter is NULL or points to a NULL string.
' ** nFilterIndex
' **   Type: DWORD
' **   The index of the currently selected filter in the File Types control. The buffer pointed to by lpstrFilter contains pairs of strings that define the filters. The first pair of strings has an index value of 1, the second pair 2, and so on. An index of zero indicates the custom filter specified by lpstrCustomFilter. You can specify an index on input to indicate the initial filter description and filter pattern for the dialog box. When the user selects a file, nFilterIndex returns the index of the currently displayed filter. If nFilterIndex is zero and lpstrCustomFilter is NULL, the system uses the first filter in the lpstrFilter buffer. If all three members are zero or NULL, the system does not use any filters and does not show any files in the file list control of the dialog box.
' ** lpstrFile
' **   Type: LPTSTR
' **   The file name used to initialize the File Name edit control. The first character of this buffer must be NULL if initialization is not necessary. When the GetOpenFileName or GetSaveFileName function returns successfully, this buffer contains the drive designator, path, file name, and extension of the selected file.
' **   If the OFN_ALLOWMULTISELECT flag is set and the user selects multiple files, the buffer contains the current directory followed by the file names of the selected files. For Explorer-style dialog boxes, the directory and file name strings are NULL separated, with an extra NULL character after the last file name. For old-style dialog boxes, the strings are space separated and the function uses short file names for file names with spaces. You can use the FindFirstFile function to convert between long and short file names. If the user selects only one file, the lpstrFile string does not have a separator between the path and file name.
' **   If the buffer is too small, the function returns FALSE and the CommDlgExtendedError function returns FNERR_BUFFERTOOSMALL. In this case, the first two bytes of the lpstrFile buffer contain the required size, in bytes or characters.
' ** nMaxFile
' **   Type: DWORD
' **   The size, in characters, of the buffer pointed to by lpstrFile. The buffer must be large enough to store the path and file name string or strings, including the terminating NULL character. The GetOpenFileName and GetSaveFileName functions return FALSE if the buffer is too small to contain the file information. The buffer should be at least 256 characters long.
' ** lpstrFileTitle
' **   Type: LPTSTR
' **   The file name and extension (without path information) of the selected file. This member can be NULL.
' ** nMaxFileTitle
' **   Type: DWORD
' **   The size, in characters, of the buffer pointed to by lpstrFileTitle. This member is ignored if lpstrFileTitle is NULL.
' ** lpstrInitialDir
' **   Type: LPCTSTR
' **   The initial directory. The algorithm for selecting the initial directory varies on different platforms.
' **     Windows 7:
' **      1.If lpstrInitialDir has the same value as was passed the first time the application used an Open or Save As dialog box, the path most recently selected by the user is used as the initial directory.
' **      2.Otherwise, if lpstrFile contains a path, that path is the initial directory.
' **      3.Otherwise, if lpstrInitialDir is not NULL, it specifies the initial directory.
' **      4.If lpstrInitialDir is NULL and the current directory contains any files of the specified filter types, the initial directory is the current directory.
' **      5.Otherwise, the initial directory is the personal files directory of the current user.
' **      6.Otherwise, the initial directory is the Desktop folder.
' **     Windows 2000/XP/Vista:
' **      1.If lpstrFile contains a path, that path is the initial directory.
' **      2.Otherwise, lpstrInitialDir specifies the initial directory.
' **      3.Otherwise, if the application has used an Open or Save As dialog box in the past, the path most recently used is selected as the initial directory. However, if an application is not run for a long time, its saved selected path is discarded.
' **      4.If lpstrInitialDir is NULL and the current directory contains any files of the specified filter types, the initial directory is the current directory.
' **      5.Otherwise, the initial directory is the personal files directory of the current user.
' **      6.Otherwise, the initial directory is the Desktop folder.
' ** lpstrTitle
' **   Type: LPCTSTR
' **   A string to be placed in the title bar of the dialog box. If this member is NULL, the system uses the default title (that is, Save As or Open).
' ** Flags
' **   Type: DWORD
' **   A set of bit flags you can use to initialize the dialog box. When the dialog box returns, it sets these flags to indicate the user's input. This member can be a combination of the following flags.
' **     Value       Constant                  Meaning
' **     ==========  ========================  =========
' **     0x00000200  OFN_ALLOWMULTISELECT      The File Name list box allows multiple selections. If you also set the OFN_EXPLORER flag, the dialog box uses the Explorer-style user interface; otherwise, it uses the old-style user interface.
' **                                           If the user selects more than one file, the lpstrFile buffer returns the path to the current directory followed by the file names of the selected files. The nFileOffset member is the offset, in bytes or characters, to the first file name, and the nFileExtension member is not used. For Explorer-style dialog boxes, the directory and file name strings are NULL separated, with an extra NULL character after the last file name. This format enables the Explorer-style dialog boxes to return long file names that include spaces. For old-style dialog boxes, the directory and file name strings are separated by spaces and the function uses short file names for file names with spaces. You can use the FindFirstFile function to convert between long and short file names.
' **                                           If you specify a custom template for an old-style dialog box, the definition of the File Name list box must contain the LBS_EXTENDEDSEL value.
' **     0x00002000  OFN_CREATEPROMPT          If the user specifies a file that does not exist, this flag causes the dialog box to prompt the user for permission to create the file. If the user chooses to create the file, the dialog box closes and the function returns the specified name; otherwise, the dialog box remains open. If you use this flag with the OFN_ALLOWMULTISELECT flag, the dialog box allows the user to specify only one nonexistent file.
' **     0x02000000  OFN_DONTADDTORECENT       Prevents the system from adding a link to the selected file in the file system directory that contains the user's most recently used documents. To retrieve the location of this directory, call the SHGetSpecialFolderLocation function with the CSIDL_RECENT flag.
' **     0x00000020  OFN_ENABLEHOOK            Enables the hook function specified in the lpfnHook member.
' **     0x00400000  OFN_ENABLEINCLUDENOTIFY   Causes the dialog box to send CDN_INCLUDEITEM notification messages to your OFNHookProc hook procedure when the user opens a folder. The dialog box sends a notification for each item in the newly opened folder. These messages enable you to control which items the dialog box displays in the folder's item list.
' **     0x00800000  OFN_ENABLESIZING          Enables the Explorer-style dialog box to be resized using either the mouse or the keyboard. By default, the Explorer-style Open and Save As dialog boxes allow the dialog box to be resized regardless of whether this flag is set. This flag is necessary only if you provide a hook procedure or custom template. The old-style dialog box does not permit resizing.
' **     0x00000040  OFN_ENABLETEMPLATE        The lpTemplateName member is a pointer to the name of a dialog template resource in the module identified by the hInstance member. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
' **     0x00000080  OFN_ENABLETEMPLATEHANDLE  The hInstance member identifies a data block that contains a preloaded dialog box template. The system ignores lpTemplateName if this flag is specified. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
' **     0x00080000  OFN_EXPLORER              Indicates that any customizations made to the Open or Save As dialog box use the Explorer-style customization methods. For more information, see Explorer-Style Hook Procedures and Explorer-Style Custom Templates.
' **                                           By default, the Open and Save As dialog boxes use the Explorer-style user interface regardless of whether this flag is set. This flag is necessary only if you provide a hook procedure or custom template, or set the OFN_ALLOWMULTISELECT flag.
' **                                           If you want the old-style user interface, omit the OFN_EXPLORER flag and provide a replacement old-style template or hook procedure. If you want the old style but do not need a custom template or hook procedure, simply provide a hook procedure that always returns FALSE.
' **     0x00000400  OFN_EXTENSIONDIFFERENT    The user typed a file name extension that differs from the extension specified by lpstrDefExt. The function does not use this flag if lpstrDefExt is NULL.
' **     0x00001000  OFN_FILEMUSTEXIST         The user can type only names of existing files in the File Name entry field. If this flag is specified and the user enters an invalid name, the dialog box procedure displays a warning in a message box. If this flag is specified, the OFN_PATHMUSTEXIST flag is also used. This flag can be used in an Open dialog box. It cannot be used with a Save As dialog box.
' **     0x10000000  OFN_FORCESHOWHIDDEN       Forces the showing of system and hidden files, thus overriding the user setting to show or not show hidden files. However, a file that is marked both system and hidden is not shown.
' **     0x00000004  OFN_HIDEREADONLY          Hides the Read Only check box.
' **     0x00200000  OFN_LONGNAMES             For old-style dialog boxes, this flag causes the dialog box to use long file names. If this flag is not specified, or if the OFN_ALLOWMULTISELECT flag is also set, old-style dialog boxes use short file names (8.3 format) for file names with spaces. Explorer-style dialog boxes ignore this flag and always display long file names.
' **     0x00000008  OFN_NOCHANGEDIR           Restores the current directory to its original value if the user changed the directory while searching for files.
' **                                           This flag is ineffective for GetOpenFileName.
' **     0x00100000  OFN_NODEREFERENCELINKS    Directs the dialog box to return the path and file name of the selected shortcut (.LNK) file. If this value is not specified, the dialog box returns the path and file name of the file referenced by the shortcut.
' **     0x00040000  OFN_NOLONGNAMES           For old-style dialog boxes, this flag causes the dialog box to use short file names (8.3 format). Explorer-style dialog boxes ignore this flag and always display long file names.
' **     0x00020000  OFN_NONETWORKBUTTON       Hides and disables the Network button.
' **     0x00008000  OFN_NOREADONLYRETURN      The returned file does not have the Read Only check box selected and is not in a write-protected directory.
' **     0x00010000  OFN_NOTESTFILECREATE      The file is not created before the dialog box is closed. This flag should be specified if the application saves the file on a create-nonmodify network share. When an application specifies this flag, the library does not check for write protection, a full disk, an open drive door, or network protection. Applications using this flag must perform file operations carefully, because a file cannot be reopened once it is closed.
' **     0x00000100  OFN_NOVALIDATE            The common dialog boxes allow invalid characters in the returned file name. Typically, the calling application uses a hook procedure that checks the file name by using the FILEOKSTRING message. If the text box in the edit control is empty or contains nothing but spaces, the lists of files and directories are updated. If the text box in the edit control contains anything else, nFileOffset and nFileExtension are set to values generated by parsing the text. No default extension is added to the text, nor is text copied to the buffer specified by lpstrFileTitle. If the value specified by nFileOffset is less than zero, the file name is invalid. Otherwise, the file name is valid, and nFileExtension and nFileOffset can be used as if the OFN_NOVALIDATE flag had not been specified.
' **     0x00000002  OFN_OVERWRITEPROMPT       Causes the Save As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
' **     0x00000800  OFN_PATHMUSTEXIST         The user can type only valid paths and file names. If this flag is used and the user types an invalid path and file name in the File Name entry field, the dialog box function displays a warning in a message box.
' **     0x00000001  OFN_READONLY              Causes the Read Only check box to be selected initially when the dialog box is created. This flag indicates the state of the Read Only check box when the dialog box is closed.
' **     0x00004000  OFN_SHAREAWARE            Specifies that if a call to the OpenFile function fails because of a network sharing violation, the error is ignored and the dialog box returns the selected file name. If this flag is not set, the dialog box notifies your hook procedure when a network sharing violation occurs for the file name specified by the user. If you set the OFN_EXPLORER flag, the dialog box sends the CDN_SHAREVIOLATION message to the hook procedure. If you do not set OFN_EXPLORER, the dialog box sends the SHAREVISTRING registered message to the hook procedure.
' **     0x00000010  OFN_SHOWHELP              Causes the dialog box to display the Help button. The hwndOwner member must specify the window to receive the HELPMSGSTRING registered messages that the dialog box sends when the user clicks the Help button. An Explorer-style dialog box sends a CDN_HELP notification message to your hook procedure when the user clicks the Help button.
' ** nFileOffset
' **   Type: WORD
' **   The zero-based offset, in characters, from the beginning of the path to the file name in the string pointed to by lpstrFile. For the ANSI version, this is the number of bytes; for the Unicode version, this is the number of characters. For example, if lpstrFile points to the following string, "c:\dir1\dir2\file.ext", this member contains the value 13 to indicate the offset of the "file.ext" string. If the user selects more than one file, nFileOffset is the offset to the first file name.
' ** nFileExtension
' **   Type: WORD
' **   The zero-based offset, in characters, from the beginning of the path to the file name extension in the string pointed to by lpstrFile. For the ANSI version, this is the number of bytes; for the Unicode version, this is the number of characters. Usually the file name extension is the substring which follows the last occurrence of the dot (".") character. For example, txt is the extension of the filename readme.txt, html the extension of readme.txt.html. Therefore, if lpstrFile points to the string "c:\dir1\dir2\readme.txt", this member contains the value 20. If lpstrFile points to the string "c:\dir1\dir2\readme.txt.html", this member contains the value 24. If lpstrFile points to the string "c:\dir1\dir2\readme.txt.html.", this member contains the value 29. If lpstrFile points to a string that does not contain any "." character such as "c:\dir1\dir2\readme", this member contains zero.
' ** lpstrDefExt
' **   Type: LPCTSTR
' **   The default extension. GetOpenFileName and GetSaveFileName append this extension to the file name if the user fails to type an extension. This string can be any length, but only the first three characters are appended. The string should not contain a period (.). If this member is NULL and the user fails to type an extension, no extension is appended.
' ** lCustData
' **   Type: LPARAM
' **   Application-defined data that the system passes to the hook procedure identified by the lpfnHook member. When the system sends the WM_INITDIALOG message to the hook procedure, the message's lParam parameter is a pointer to the OPENFILENAME structure specified when the dialog box was created. The hook procedure can use this pointer to get the lCustData value.
' ** lpfnHook
' **   Type: LPOFNHOOKPROC
' **   A pointer to a hook procedure. This member is ignored unless the Flags member includes the OFN_ENABLEHOOK flag.
' **   If the OFN_EXPLORER flag is not set in the Flags member, lpfnHook is a pointer to an OFNHookProcOldStyle hook procedure that receives messages intended for the dialog box. The hook procedure returns FALSE to pass a message to the default dialog box procedure or TRUE to discard the message.
' **   If OFN_EXPLORER is set, lpfnHook is a pointer to an OFNHookProc hook procedure. The hook procedure receives notification messages sent from the dialog box. The hook procedure also receives messages for any additional controls that you defined by specifying a child dialog template. The hook procedure does not receive messages intended for the standard controls of the default dialog box.
' ** lpTemplateName
' **   Type: LPCTSTR
' **   The name of the dialog template resource in the module identified by the hInstance member. For numbered dialog box resources, this can be a value returned by the MAKEINTRESOURCE macro. This member is ignored unless the OFN_ENABLETEMPLATE flag is set in the Flags member. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
' ** pvReserved
' **   Type: void*
' **   This member is reserved.
' ** dwReserved
' **   Type: DWORD
' **   This member is reserved.
' ** FlagsEx
' **   Type: DWORD
' **   A set of bit flags you can use to initialize the dialog box. Currently, this member can be zero or the following flag.
' **     Value       Constant            Meaning
' **     ==========  ==================  =========
' **     0x00000001  OFN_EX_NOPLACESBAR  If this flag is set, the places bar is not displayed. If this flag is not set, Explorer-style dialog boxes include a places bar containing icons for commonly-used folders, such as Favorites and Desktop.

' ** Constants for OPENFILENAME.
'Public Const OFN_ALLOWMULTISELECT   As Long = &H200
'Public Const OFN_CREATEPROMPT       As Long = &H2000
'Public Const OFN_EXPLORER           As Long = &H80000
Public Const OFN_FILEMUSTEXIST      As Long = &H1000
'Public Const OFN_HIDEREADONLY       As Long = &H4
'Public Const OFN_NOCHANGEDIR        As Long = &H8
'Public Const OFN_NODEREFERENCELINKS As Long = &H100000
'Public Const OFN_NONETWORKBUTTON    As Long = &H20000
'Public Const OFN_NOREADONLYRETURN   As Long = &H8000
'Public Const OFN_NOVALIDATE         As Long = &H100
Public Const OFN_OVERWRITEPROMPT    As Long = &H2
Public Const OFN_PATHMUSTEXIST      As Long = &H800
'Public Const OFN_READONLY           As Long = &H1
'Public Const OFN_SHOWHELP           As Long = &H10

' ** Constants for file dialog.
' ** Right now, these can remain Private.
'Private Const FNERR_BUFFERTOOSMALL  As Integer = &H3003
'Private Const FNERR_FILENAMECODES   As Integer = &H3000
'Private Const FNERR_INVALIDFILENAME As Integer = &H3002
'Private Const FNERR_SUBCLASSFAILURE As Integer = &H3001

' ** Structure used by SHBrowseForFolder().
Public Type BROWSEINFO
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type

' ** Constants for BROWSEINFO.
Public Const BIF_RETURNONLYFSDIRS  As Long = &H1
Public Const BIF_MAXPATH           As Long = 260&
'Public Const BIF_DONTGOBELOWDOMAIN As Long = 2&

' ** Used with GetVersionEx() API.
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long          ' ** Size, in bytes, of this data structure
  dwMajorVersion      As Long          ' ** E.g., NT 3.51, dwMajorVersion = 3; NT 4.0, dwMajorVersion = 4.
  dwMinorVersion      As Long          ' ** E.g, NT 3.51, dwMinorVersion = 51; NT 4.0, dwVerMinor= 0.
  dwBuildNumber       As Long          ' ** NT: Build number of the OS.  ' ** Win9x: Build number of the OS in low-order word.  ' **        High-order word contains major & minor ver nos.
  dwPlatformId        As Long          ' ** Identifies the operating system platform.
  szCSDVersion        As String * 128  ' ** Maintenance string for PSS usage.   ' ** NT: String, such as "Service Pack 3".  ' ** Win9x: String providing arbitrary additional information.
End Type

' ** Used with GetVersionEx() API.
'Public Type OSVERSIONINFOEX
'  dwOSVersionInfoSize As Long
'  dwMajorVersion      As Long
'  dwMinorVersion      As Long
'  dwBuildNumber       As Long
'  dwPlatformId        As Long
'  szCSDVersion        As String * 128
'  wServicePackMajor   As Integer
'  wServicePackMinor   As Integer
'  wSuiteMask          As Integer
'  wProductType        As Byte
'  wReserved           As Byte
'End Type

' ** Used with GetSystemMetrics() API.  ' ** Not currently used!
'Public Const SM_CYCAPTION As Long = 4&  '** Height of caption or title.
'Public Const SM_TABLETPC As Long = 86&
'Public Const SM_MEDIACENTER As Long = 87&
'Public Const SM_STARTER As Long = 88&
'Public Const SM_SERVERR2 As Long = 89&

Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" _
  (ByVal dwExStyle As Long, ByVal lpClassname As String, ByVal lpWindowName As String, ByVal dwStyle As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, _
  ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

' ** API memory functions.
Public Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long

' ** Defined 'As Any' to support both OSVERSIONINFO and OSVERSIONINFOEX.
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As Any) As Long

' ** API function called by ChooseColor method.
Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As COLORSTRUC) As Long

Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

Public Declare Function SetSysColors Lib "user32.dll" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long

' ** API function called by ShowOpen method.
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

' ** API function called by ShowSave method.
' ** The GetSaveFileName function creates a Save common dialog box that
' ** lets the user specify the drive, directory, and name of a file to save.
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
' ** Library               :  Comdlg32
' ** Parameter Information :  ·lpofn
' **                       :  Pointer to an OPENFILENAME structure that contains information used to
' **                       :  initialize the dialog box. When GetSaveFileName returns, this structure
' **                       :  contains information about the user’s file selection.
' ** Return Values         :  If the user specifies a filename and clicks the OK button, the return value
' **                       :  is nonzero. The buffer pointed to by the lpstrFile member of the OPENFILENAME
' **                       :  structure contains the full path and filename specified by the user.
' **                       :  If the user cancels or closes the Save dialog box or an error occurs,
' **                       :  the return value is zero. To get extended error information, call the
' **                       :  CommDlgExtendedError function, which can return one of the following values:
' **                       :    CDERR_FINDRESFAILURE , CDERR_NOHINSTANCE, CDERR_INITIALIZATION, CDERR_NOHOOK,
' **                       :    CDERR_LOCKRESFAILURE, CDERR_NOTEMPLATE, CDERR_LOADRESFAILURE, CDERR_STRUCTSIZE,
' **                       :    CDERR_LOADSTRFAILURE, FNERR_BUFFERTOOSMALL, CDERR_MEMALLOCFAILURE,
' **                       :    FNERR_INVALIDFILENAME, CDERR_MEMLOCKFAILURE, FNERR_SUBCLASSFAILURE
' ** OPENFILENAME:
' **   lStructSize         :  Specifies the length, in bytes, of the structure.
' **   hwndOwner           :  Handle to the window that owns the dialog box. This member can be any
' **                       :  valid window handle, or it can be NULL if the dialog box has no owner.
' **   hInstance           :  Not supported.
' **   lpstrFilter         :  Long pointer to a buffer that contains pairs of null-terminated filter strings.
' **                       :  The last string in the buffer must be terminated by two NULL characters.
' **                       :  The first string in each pair is a display string that describes the filter
' **                       :  (for example, Text Files), and the second string specifies the filter pattern
' **                       :  (for example, *.TXT). To specify multiple filter patterns for a single display
' **                       :  string, use a semicolon to separate the patterns (for example, *.TXT;*.DOC;*.BAK).
' **                       :  A pattern string can be a combination of valid file name characters and the
' **                       :  asterisk (*) wildcard character. Do not include spaces in the pattern string.
' **                       :  The system does not change the order of the filters. It displays them in the
' **                       :  File Types combo box in the order specified in lpstrFilter.
' **                       :  If lpstrFilter is NULL, the dialog box does not display any filters.
' **   lpstrCustomFilter   :  Not supported.
' **   nMaxCustFilter      :  Not supported.
' **   nFilterIndex        :  Specifies the index of the currently selected filter in the File Types control.
' **                       :  The buffer pointed to by lpstrFilter contains pairs of strings that define the
' **                       :  filters. The first pair of strings has an index value of 1, the second pair 2,
' **                       :  and so on. You can specify an index on input to indicate the initial filter
' **                       :  description and filter pattern for the dialog box. When the user selects a file,
' **                       :  nFilterIndex returns the index of the currently displayed filter.
' **                       :  If nFilterIndex is zero, the system uses the first filter in the lpstrFilter buffer.
' **   lpstrFile           :  Long pointer to a buffer that contains a file name used to initialize the
' **                       :  File Name edit control. The first character of this buffer must be NULL if
' **                       :  initialization is not necessary. When the GetOpenFileName or GetSaveFileName
' **                       :  function returns successfully, this buffer contains the drive designator, path,
' **                       :  file name, and extension of the selected file.
' **                       :  If the buffer is too small, the function returns FALSE. In this case,
' **                       :  the first two bytes of the lpstrFile buffer contain the required size,
' **                       :  in bytes or characters.
' **   nMaxFile            :  Specifies the size, in bytes (ANSI version) or 16-bit characters (Unicode version),
' **                       :  of the buffer pointed to by lpstrFile. The GetOpenFileName and GetSaveFileName
' **                       :  functions return FALSE if the buffer is too small to contain the file information.
' **                       :  The buffer should be at least 256 characters long.
' **   lpstrFileTitle      :  Long pointer to a buffer that receives the file name and extension
' **                       :  (without path information) of the selected file. This member can be NULL.
' **   nMaxFileTitle       :  Specifies the size, in bytes (ANSI version) or 16-bit characters (Unicode version),
' **                       :  of the buffer pointed to by lpstrFileTitle. This member is ignored
' **                       :  if lpstrFileTitle is NULL.
' **   lpstrInitialDir     :  Long pointer to a string that specifies the initial file directory.
' **                       :  If this member is NULL, the system uses the root directory.
' **   lpstrTitle          :  Long pointer to a string to be placed in the title bar of the dialog box.
' **                       :  If this member is NULL, the system uses the default title (that is, SaveAs or Open).
' **   Flags               :  A bitmask of flags used to initialize the dialog box. When the dialog box returns,
' **                       :  it sets these flags to indicate the user's input. This member can be a combination
' **                       :  of the following flags.
' **                       :    Value                     Description
' **                       :    ========================  =========================================================================
' **                       :    OFN_ALLOWMULTISELECT      Not supported.
' **                       :    OFN_CREATEPROMPT          If the user specifies a file that does not exist, this flag causes the
' **                       :                              dialog box to prompt the user for permission to create the file. If the
' **                       :                              user chooses to create the file, the dialog box closes and the function
' **                       :                              returns the specified name; otherwise, the dialog box remains open.
' **                       :    OFN_ENABLEHOOK            Not supported.
' **                       :    OFN_ENABLESIZING          Not supported.
' **                       :    OFN_ENABLETEMPLATE        Not supported.
' **                       :    OFN_ENABLETEMPLATEHANDLE  Not supported.
' **                       :    OFN_EXPLORER              Ignored. The Explorer user interface is always used.
' **                       :    OFN_EXTENSIONDIFFERENT    Specifies that the user typed a file name extension that differs from
' **                       :                              the extension specified by lpstrDefExt. The function does not use this
' **                       :                              flag if lpstrDefExt is NULL.
' **                       :    OFN_FILEMUSTEXIST         Specifies that the user can type only names of existing files in the
' **                       :                              File Name entry field. If this flag is specified and the user enters an
' **                       :                              invalid name, the dialog box procedure displays a warning in a message
' **                       :                              box. If this flag is specified, the OFN_PATHMUSTEXIST flag is also used.
' **                       :    OFN_HIDEREADONLY          Hides the Read Only check box.
' **                       :    OFN_LONGNAMES             Ignored. The Explorer user interface is always used.
' **                       :    OFN_NOCHANGEDIR           Not supported.
' **                       :    OFN_NODEREFERENCELINKS    Directs the dialog box to return the path and file name of the selected
' **                       :                              shortcut (.LNK) file. If this value is not given, the dialog box returns
' **                       :                              the path and file name of the file referenced by the shortcut
' **                       :    OFN_NOLONGNAMES           Not supported.
' **                       :    OFN_NONETWORKBUTTON       Not supported.
' **                       :    OFN_NOREADONLYRETURN      Not supported.
' **                       :    OFN_NOTESTFILECREATE      Not supported.
' **                       :    OFN_NOVALIDATE            Ignored. A file name is always validated.
' **                       :    OFN_PROJECT               For version 2.0, causes the GetOpenFileName function to open the Project
' **                       :                              dialog box.
' **                       :    OFN_PROPERTY              For version 2.0, causes the GetSaveFileName function to open the Property
' **                       :                              dialog box.
' **                       :    OFN_OVERWRITEPROMPT       Causes the Save As dialog box to generate a message box if the selected
' **                       :                              file already exists. The user must confirm whether to overwrite the file.
' **                       :    OFN_PATHMUSTEXIST         Specifies that the user can type only valid paths and file names.
' **                       :                              If this flag is used and the user types an invalid path and file name in
' **                       :                              the File Name entry field, the dialog box function displays a warning in
' **                       :                              a message box.
' **                       :    OFN_READONLY              Not supported.
' **                       :    OFN_SHAREAWARE            Not supported.
' **                       :    OFN_SHOW_ALL              Specifies that if OFN_PROJECT is set, show the <All Folders> item.
' **                       :                              Note: This flag applies only to Windows Mobile devices.
' **                       :    OFN_SHOWHELP              Not supported.
' **   nFileOffset         :  Specifies the zero-based offset, in bytes (ANSI version) or 16-bit characters (Unicode version),
' **                       :  from the beginning of the path to the file name in the string pointed to by lpstrFile. For example,
' **                       :  if lpstrFile points to the following string, c:\dir1\dir2\file.ext, this member contains the
' **                       :  value 13 to indicate the offset of the file.ext string.
' **   nFileExtension      :  Specifies the zero-based offset, in bytes (ANSI version) or 16-bit characters (Unicode version),
' **                       :  from the beginning of the path to the file name extension in the string pointed to by lpstrFile.
' **                       :  For example, if lpstrFile points to the following string, c:\dir1\dir2\file.ext, this member
' **                       :  contains the value 18. If the user did not type an extension and lpstrDefExt is NULL, this member
' **                       :  specifies an offset to the terminating null character. If the user typed "." as the last character
' **                       :  in the file name, this member specifies zero.
' **   lpstrDefExt         :  Long pointer to a buffer that contains the default extension. GetOpenFileName and GetSaveFileName
' **                       :  append this extension to the file name if the user fails to type an extension. This string can be
' **                       :  any length, but only the first three characters are appended. The string should not contain a
' **                       :  period (.). If this member is NULL and the user fails to type an extension, no extension is appended.
' **   lCustData           :  Not supported.
' **   lpfnHook            :  Not supported.
' **   lpTemplateName      :  Not supported.

Public Declare Function GetActiveWindow Lib "user32.dll" () As Integer

'Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BROWSEINFO) As Long

Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
  (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef pidl As Long) As Long

'Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

' ** Function names ARE CASE-SENSITIVE!
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function timeGetSystemTime Lib "winmm.dll" (lpTime As MMTIME, ByVal uSize As Long) As Long

Public Declare Function GetPrivateProfileStringA Lib "kernel32.dll" _
  (ByVal lpApplicationName As String, ByVal lprofKeyName As Any, ByVal lprofDefString As String, ByVal lRetStringedString As String, _
  ByVal nSize As Integer, ByVal lpFileName As String) As Integer

Public Declare Function WritePrivateProfileStringA Lib "kernel32.dll" _
  (ByVal lpApplicationName As String, ByVal lprofKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Integer

Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
' ** EnumWindows() Function:
' **   Enumerates all top-level windows on the screen by passing the handle to each window, in turn,
' **   to an application-defined callback function. EnumWindows continues until the last top-level
' **   window is enumerated or the callback function returns FALSE.
' **
' ** Syntax:
' **   CopyBOOL WINAPI EnumWindows(
' **     __in  WNDENUMPROC lpEnumFunc,
' **     __in  LPARAM lParam
' **   );
' **
' ** Parameters:
' **   lpEnumFunc [in]
' **     Type: WNDENUMPROC
' **     A pointer to an application-defined callback function. For more information, see EnumWindowsProc.
' **   lParam [in]
' **     Type: LPARAM
' **     An application-defined value to be passed to the callback function.
' **
' ** Return Value:
' **   Type: BOOL
' **   If the function succeeds, the return value is nonzero.
' **   If the function fails, the return value is zero. To get extended error information, call GetLastError.
' **   If EnumWindowsProc returns zero, the return value is also zero. In this case, the callback function
' **     should call SetLastError to obtain a meaningful error code to be returned to the caller of EnumWindows.
' **
' ** Remarks:
' **   The EnumWindows function does not enumerate child windows, with the exception
' **     of a few top-level windows owned by the system that have the WS_CHILD style.
' **   This function is more reliable than calling the GetWindow function in a loop.
' **   An application that calls GetWindow to perform this task risks being caught in
' **     an infinite loop or referencing a handle to a window that has been destroyed.

Public Declare Function EnumChildWindows Lib "user32.dll" _
  (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
' ** EnumChildWindows() Function:
' **   Enumerates the child windows that belong to the specified parent window by passing the handle to each
' **   child window, in turn, to an application-defined callback function. EnumChildWindows continues until
' **   the last child window is enumerated or the callback function returns FALSE.
' **
' ** Syntax:
' **   CopyBOOL WINAPI EnumChildWindows(
' **     __in_opt  HWND hWndParent,
' **     __in      WNDENUMPROC lpEnumFunc,
' **     __in      LPARAM lParam
' **   );
' **
' ** Parameters:
' **   hWndParent [in, optional]
' **     Type: HWND
' **     A handle to the parent window whose child windows are to be enumerated.
' **     If this parameter is NULL, this function is equivalent to EnumWindows.
' **   lpEnumFunc [in]
' **     Type: WNDENUMPROC
' **     A pointer to an application-defined callback function. For more information, see EnumChildProc.
' **   lParam [in]
' **     Type: LPARAM
' **     An application-defined value to be passed to the callback function.
' **
' ** Return Value:
' **   Type: BOOL
' **   The return value is not used.
' **
' ** Remarks:
' **   If a child window has created child windows of its own, EnumChildWindows enumerates those windows as well.
' **   A child window that is moved or repositioned in the Z order during the enumeration process will be
' **     properly enumerated. The function does not enumerate a child window that is destroyed before being
' **     enumerated or that is created during the enumeration process.

Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" _
  (ByVal hwnd As Long, ByVal lpClassname As String, ByVal nMaxCount As Long) As Long

Public Declare Function SendMessageS Lib "user32.dll" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

' ** RegQueryInfoKey retrieves information about the specified
' ** key, such as the number of subkeys and values, the length
' ** of the longest value and key name, and the size of the
' ** longest data component among the key's values.
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
  (ByVal hCurKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, _
  lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
  lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
  lpftLastWriteTime As FILETIME) As Long
  'hKey                   : A handle to an open registry key. The key must have been opened
  '                       : with the KEY_QUERY_VALUE access right.
  'lpClass                : A pointer to a buffer that receives the key class. This parameter can be NULL.
  'lpcbClass              : A pointer to a variable that specifies the size of the buffer pointed to by
  '                       : the lpClass parameter, in characters.
  'lpReserved             : This parameter is reserved and must be NULL.
  'lpcSubKeys             : A pointer to a variable that receives the number of subkeys that are
  '                       : contained by the specified key. This parameter can be NULL.
  'lpcbMaxSubKeyLen       : A pointer to a variable that receives the size of the key's subkey with the longest name,
  '                       : in Unicode characters, not including the terminating null character. This parameter
  '                       : can be NULL. For Windows Me/98/95, the size includes the terminating null character.
  'lpcbMaxClassLen        : A pointer to a variable that receives the size of the longest string that specifies
  '                       : a subkey class, in Unicode characters. The count returned does not include the
  '                       : terminating null character. This parameter can be NULL.
  'lpcValues              : A pointer to a variable that receives the number of values that are associated
  '                       : with the key. This parameter can be NULL.
  'lpcbMaxValueNameLen    : A pointer to a variable that receives the size of the key's longest value name,
  '                       : in Unicode characters. The size does not include the terminating null character.
  '                       : This parameter can be NULL.
  'lpcbMaxValueLen        : A pointer to a variable that receives the size of the longest data component
  '                       : among the key's values, in bytes. This parameter can be NULL.
  'lpcbSecurityDescriptor : A pointer to a variable that receives the size of the key's security descriptor,
  '                       : in bytes. This parameter can be NULL.
  'lpftLastWriteTime      : A pointer to a FILETIME structure that receives the last write time.
  '                       : This parameter can be NULL.
  '                       : The function sets the members of the FILETIME structure to indicate the last time
  '                       : that the key or any of its value entries is modified.
  '                       : For Windows Me/98/95, the function sets the members of the FILETIME structure to 0 (zero),
  '                       : because the system does not keep track of registry key last write time information.

'Obsolete, replaced by RegEnumKeyEx().
'Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
'  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
'  'hKey     : A handle to an open registry key. The key must have been
'  '         : opened with the KEY_ENUMERATE_SUB_KEYS access right (as in, "Rights of the User").
'  'dwIndex  : The index of the subkey of hKey to be retrieved. This value should be zero for the
'  '         : first call to the RegEnumKey function and then incremented for subsequent calls.
'  'lpName   : A pointer to a buffer that receives the name of the subkey, including the terminating null character.
'  '         : This function copies only the name of the subkey, not the full key hierarchy, to the buffer.
'  'cbName   : The size of the buffer pointed to by the lpName parameter, in TCHARs. To determine the required
'  '         : buffer size, use the RegQueryInfoKey function to determine the size of the largest subkey for
'  '         : the key identified by the hKey parameter.

' ** RegEnumKeyEx enumerates subkeys of the specified open
' ** key. Retrieves the name (and its length) of each subkey.
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
  (ByVal hCurKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, _
   ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long  ' ** (All other module declarations of this are also Private.)
  'hKey              : Handle to a currently open key or one of the predefined reserved handle values.
  'dwIndex           : Index of the subkey to retrieve. This parameter should be zero for the first call to
  '                  : the RegEnumKeyEx function and then incremented for subsequent calls.
  '                  : Because subkeys are not ordered, any new subkey will have an arbitrary index.
  '                  : This means that the function may return subkeys in any order.
  'lpName            : Pointer to a buffer that receives the name of the subkey, including the
  '                  : terminating null character. The function copies only the name of the
  '                  : subkey, not the full key hierarchy, to the buffer.
  'lpcbName          : Pointer to a variable that specifies the size, in characters, of the buffer specified
  '                  : by the lpName parameter. This size should include the terminating null character.
  '                  : When the function returns, the variable pointed to by lpcName contains the number
  '                  : of characters stored in the buffer. The count returned does not include the
  '                  : terminating null character.
  'lpReserved        : Reserved; set to NULL.
  'lpClass           : Pointer to a buffer that contains the class of the enumerated subkey when the
  '                  : function returns. This parameter can be NULL if the class is not required.
  'lpcbClass         : Pointer to a variable that specifies the size, in characters, of the buffer specified
  '                  : by the lpClass parameter. The size should include the terminating null character.
  '                  : When the function returns, lpcbClass contains the number of characters stored
  '                  : in the buffer. The count returned does not include the terminating null character.
  '                  : This parameter can be NULL only if lpClass is NULL.
  'lpftLastWriteTime : Ignored; set to NULL.

' ** RegEnumValue enumerates the values for the specified open
' ** key. Retrieves the name (and its length) of each value,
' ** and the type, content and size of the data.
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hCurKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
  lpType As Long, lpData As Any, lpcbData As Long) As Long  ' ** (All other module declarations of this are also Private.)
  'hKey          : Handle to a currently open key or one of the predefined reserved handle values.
  'dwIndex       : Index of the value to retrieve. This parameter should be zero for the first call
  '              : to the RegEnumValue function and then be incremented for subsequent calls.
  '              : Because values are not ordered, any new value will have an arbitrary index.
  '              : This means that the function may return values in any order.
  'lpValueName   : Pointer to a buffer that receives the name of the value, including the terminating null character.
  'lpcbValueName : Pointer to a variable that specifies the size, in characters, of the buffer pointed to
  '              : by the lpValueName parameter. This size should include the terminating null character.
  '              : When the function returns, the variable pointed to by lpcchValueName contains
  '              : the number of characters stored in the buffer. The count returned does not
  '              : include the terminating null character.
  'lpReserved    : Reserved; set to NULL.
  'lpType        : Pointer to a variable that receives the type code for the value entry.
  'lpData        : Pointer to a buffer that receives the data for the value entry.
  '              : This parameter can be NULL if the data is not required.
  'lpcbData      : Pointer to a variable that specifies the size, in bytes, of the buffer pointed
  '              : to by the lpData parameter. When the function returns, the variable pointed to
  '              : by the lpcbData parameter contains the number of bytes stored in the buffer.
  '              : This parameter can be NULL, only if lpData is NULL.

' ** RegDeleteValue removes a named value from specified key.
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
  (ByVal hCurKey As Long, ByVal lpValueName As String) As Long  ' ** (All other module declarations of this are also Private.)
  'hKey         : A handle to an open registry key. The key must have been opened with the
  '             : KEY_SET_VALUE access right.
  'lpValueName  : The registry value to be removed. If this parameter is NULL or an empty string,
  '             : the value set by the RegSetValue function is removed.

' ** RegDeleteKey deletes a subkey. Under Win 95/98, also
' ** deletes all subkeys and values. Under Windows NT/2000,
' ** the subkey to be deleted must not have subkeys. The class
' ** attempts to use SHDeleteKey (see below) before using
' ** RegDeleteKey.
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long
  'hKey     : Handle of open key
  'lpSubKey : Path from handle to key

Public Declare Function SetWindowPos Lib "user32.dll" _
  (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
  ByVal Y As Long, ByVal Cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' ** Declarations for detecting mutiple TA instances.
Public Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" _
  (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" _
  (ByVal hwnd As Long) As Long

Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" _
  (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Declare Function GetWindowInfo Lib "user32.dll" (ByVal hwnd As Long, lpRec As WININFO_TYPE) As Boolean

Public Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' ** Both versions are used.
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
  (ByVal lpClassname As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" _
  (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
' ** FindWindowEx() Function:
' **   Retrieves a handle to a window whose class name and window name match the specified strings.
' **   The function searches child windows, beginning with the one following the specified child window.
' **   This function does not perform a case-sensitive search.
' **
' ** Parameters:
' **   hwndParent [in, optional]
' **     Type: HWND
' **     A handle to the parent window whose child windows are to be searched.
' **     If hwndParent is NULL, the function uses the desktop window as the parent window.
' **       The function searches among windows that are child windows of the desktop.
' **     If hwndParent is HWND_MESSAGE, the function searches all message-only windows.
' **   hwndChildAfter [in, optional]
' **     Type: HWND
' **     A handle to a child window. The search begins with the next child window in the Z order.
' **       The child window must be a direct child window of hwndParent, not just a descendant window.
' **     If hwndChildAfter is NULL, the search begins with the first child window of hwndParent.
' **     Note that if both hwndParent and hwndChildAfter are NULL, the function searches all top-level
' **       and message-only windows.
' **   lpszClass [in, optional]
' **     Type: LPCTSTR
' **     The class name or a class atom created by a previous call to the RegisterClass or RegisterClassEx function.
' **       The atom must be placed in the low-order word of lpszClass; the high-order word must be zero.
' **     If lpszClass is a string, it specifies the window class name.
' **       The class name can be any name registered with RegisterClass or RegisterClassEx,
' **       or any of the predefined control-class names, or it can be MAKEINTATOM(0x8000).
' **       In this latter case, 0x8000 is the atom for a menu class.
' **       For more information, see the Remarks section of this topic.
' **   lpszWindow [in, optional]
' **     Type: LPCTSTR
' **     The window name (the window's title). If this parameter is NULL, all window names match.
' **
' ** Return Value:
' **   Type: HWND
' **   If the function succeeds, the return value is a handle to the
' **     window that has the specified class and window names.
' **   If the function fails, the return value is NULL. To get extended error information, call GetLastError.
' **
' ** Remarks:
' **   If the lpszWindow parameter is not NULL, FindWindowEx calls the GetWindowText function
' **     to retrieve the window name for comparison. For a description of a potential problem
' **     that can arise, see the Remarks section of GetWindowText.
' **   An application can call this function in the following way.
' **     FindWindowEx( NULL, NULL, MAKEINTATOM(0x8000), NULL );
' **   Note that 0x8000 is the atom for a menu class. When an application calls this function,
' **     the function checks whether a context menu is being displayed that the application created.

Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' ** Alias for SendMessageTimeOut that allows a string as the lParam.
Public Declare Function SendMessageTimeoutStr Lib "user32.dll" Alias "SendMessageTimeoutA" _
  (ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As String, _
  ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" _
  (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long

Public Declare Function TranslateMessage Lib "user32.dll" (lpMsg As MSG) As Long

Public Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As MSG) As Long

Public Declare Sub PostQuitMessage Lib "user32.dll" (ByVal nExitCode As Long)

Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long

' ** Array: garr_varClass().
Public glngClasses As Long, garr_varClass() As Variant
Public Const CLS_ELEMS As Integer = 6  ' ** Array's first-element UBound().
Public Const CLS_CLASS   As Integer = 0
Public Const CLS_TITLE   As Integer = 1
Public Const CLS_HWND    As Integer = 2
Public Const CLS_VISIBLE As Integer = 3
Public Const CLS_PARENT  As Integer = 4
'Public Const CLS_CHILD   As Integer = 5
'Public Const CLS_PHWND   As Integer = 6

Private Type WININFO_TYPE
  cbSize As Long
  rcWindow As RECT
  rcClient As RECT
  dwStyle As Long
  dwExStyle As Long
  dwWindowStatus As Long
  cxWindowBorders As Long
  cyWindowBorders As Long
  atomWindowtype As Long
  wCreatorVersion As Long
End Type

' ** Developer's Twips per Pixel.
Private Const lngMyTpp As Long = 15&

Private blnFromCode As Boolean, blnNavPaneOpen As Boolean
Private blnWindowVisible As Boolean

' ** Moves MS Access window to top of Z-order.
Private Const HWND_TOP       As Long = &H0&
Private Const HWND_TOPMOST   As Long = &HFFFF&
Private Const HWND_NOTOPMOST As Long = &HFFFE&
'Private Const HWND_BROADCAST As Long = &HFFFF&
' **

Public Sub MakeTopMost(hwnd As Long)
' ** Doesn't really work!

100   On Error GoTo ERRH

        Const THIS_PROC As String = "MakeTopMost"

110     SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS  ' ** API Function: Above.

EXITP:
120     Exit Sub

ERRH:
130     Select Case ERR.Number
        Case Else
140       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
150     End Select
160     Resume EXITP

End Sub

Public Sub MakeNormal(hwnd As Long)

200   On Error GoTo ERRH

        Const THIS_PROC As String = "MakeNormal"

210     SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS  ' ** API Function: Above.

EXITP:
220     Exit Sub

ERRH:
230     Select Case ERR.Number
        Case Else
240       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
250     End Select
260     Resume EXITP

End Sub

Public Function Scr() As Boolean
' ** Reset the screen after a break (or anytime).
' ** Invoke from Debug/Immediate Window (or anywhere).

300   On Error GoTo ERRH

        Const THIS_PROC As String = "Scr"

        Dim blnRetVal As Boolean

310     blnRetVal = True

320   On Error Resume Next
330     DoCmd.Hourglass False
340     DoCmd.SetWarnings True  ' ** I normally want them on, but off when TA is running.
350     Application.Echo True
360     SysCmd acSysCmdClearStatus
        'Application.MenuBar = vbNullString

EXITP:
370     Scr = blnRetVal
380     Exit Function

ERRH:
390     blnRetVal = False
400     Select Case ERR.Number
        Case Else
410       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
420     End Select
430     Resume EXITP

End Function

Public Function IsMaximized(frm As Access.Form) As Boolean

500   On Error GoTo ERRH

        Const THIS_PROC As String = "IsMaximized"

        Dim blnRetVal As Boolean

510     blnRetVal = CBool(IsZoomed(frm.hwnd))

EXITP:
520     IsMaximized = blnRetVal
530     Exit Function

ERRH:
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

End Function

Public Function MaximizeApp() As Integer

600   On Error GoTo ERRH

        Const THIS_PROC As String = "MaximizeApp"

        Dim intRetVal As Integer

610     intRetVal = ShowWindow(hWndAccessApp, SW_SHOWMAXIMIZED)  ' ** API Function: Above.

EXITP:
620     MaximizeApp = intRetVal
630     Exit Function

ERRH:
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Function

Public Function MinimizeApp() As Integer

700   On Error GoTo ERRH

        Const THIS_PROC As String = "MinimizeApp"

        Dim intRetVal As Integer

710     intRetVal = ShowWindow(hWndAccessApp, SW_SHOWMINIMIZED)  ' ** API Function: Above.

EXITP:
720     MinimizeApp = intRetVal
730     Exit Function

ERRH:
740     Select Case ERR.Number
        Case Else
750       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
760     End Select
770     Resume EXITP

End Function

Public Function RestoreApp() As Integer

800   On Error GoTo ERRH

        Const THIS_PROC As String = "RestoreApp"

        Dim intRetVal As Integer

810     intRetVal = ShowWindow(hWndAccessApp, SW_SHOWNORMAL)  ' ** API Function: Above.

EXITP:
820     RestoreApp = intRetVal
830     Exit Function

ERRH:
840     Select Case ERR.Number
        Case Else
850       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
860     End Select
870     Resume EXITP

End Function

Public Function GetTPP() As Long

900   On Error GoTo ERRH

        Const THIS_PROC As String = "GetTPP"

        Dim lngRetVal As Long

910     lngRetVal = (GetScreenRes(GSR_HORZRES) / GetScreenRes(GSR_LOGPIXELSX))  ' ** Functions: Below.

EXITP:
920     GetTPP = lngRetVal
930     Exit Function

ERRH:
940     lngRetVal = 0&
950     Select Case ERR.Number
        Case Else
960       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
970     End Select
980     Resume EXITP

End Function

Public Function GetScreenRes(Optional varOption As Variant) As Long

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "GetScreenRes"

        Dim lngHdc As Long
        Dim lngCurrHRes As Long, lngCurrVRes As Long
        Dim lngAspectX As Long, lngAspectY As Long, lngHyp As Long
        Dim lngPpiX As Long, lngPpiY As Long
        Dim lngDeskWidth As Long, lngDeskHeight As Long
        Dim lngCurrBPP As Long, lngCurrVFreq As Long
        Dim strBPPtype As String, strFreqtype As String
        Dim lngRetVal As Long

1010    lngRetVal = 0&

1020    lngHdc = GetDC(hWndAccessApp)

        ' ** Screen resolutions available on Delta Data's Acer monitor:
        ' **  Horiz   Vert
        ' **  =====   ====
        ' **    800 x  600
        ' **    960 x  600
        ' **    960 x 1200
        ' **   1024 x  768
        ' **   1088 x  612
        ' **   1152 x  864
        ' **   1280 x  720
        ' **   1280 x  768
        ' **   1280 x  800
        ' **   1280 x  960
        ' **   1280 x 1024
        ' **   1360 x  768
        ' **   1440 x  900

1030    If IsMissing(varOption) = False Then
1040      lngRetVal = GetDeviceCaps(lngHdc, CLng(varOption))
1050    Else
          ' ** Only lists Debug when no option specified.

          ' ** Get the system settings.
1060      lngCurrHRes = GetDeviceCaps(lngHdc, GSR_HORZRES)
1070      lngCurrVRes = GetDeviceCaps(lngHdc, GSR_VERTRES)
1080      lngCurrBPP = GetDeviceCaps(lngHdc, GSR_BITSPIXEL)
1090      lngCurrVFreq = GetDeviceCaps(lngHdc, GSR_VREFRESH)
1100      lngAspectX = GetDeviceCaps(lngHdc, GSR_ASPECTX)
1110      lngAspectY = GetDeviceCaps(lngHdc, GSR_ASPECTY)
1120      lngHyp = GetDeviceCaps(lngHdc, GSR_ASPECTXY)
1130      lngPpiX = GetDeviceCaps(lngHdc, GSR_LOGPIXELSX)
1140      lngPpiY = GetDeviceCaps(lngHdc, GSR_LOGPIXELSY)
1150      lngDeskWidth = GetDeviceCaps(lngHdc, GSR_DESKTOPHORZRES)
1160      lngDeskHeight = GetDeviceCaps(lngHdc, GSR_DESKTOPVERTRES)

          ' ** Pretty up the descriptions a tad.
1170      Select Case lngCurrBPP
          Case 4
1180        strBPPtype = "(16 Color)"
1190      Case 8
1200        strBPPtype = "(256 Color)"
1210      Case 16
1220        strBPPtype = "(High Color)"
1230      Case 24, 32
1240        strBPPtype = "(True Color)"
1250      End Select

1260      Select Case lngCurrVFreq
          Case 0, 1
1270        strFreqtype = "(Hardware default)"
1280      Case Else
1290        strFreqtype = "(User-selected)"
1300      End Select

1310      Win_Mod_Restore  ' ** Procedure: Below.

1320      Debug.Print "'RES X : " & lngCurrHRes & " pixels"
1330      Debug.Print "'RES Y : " & lngCurrVRes & " pixels"
1340      Debug.Print "'BPP   : " & lngCurrBPP & " bits per pixel  " & strBPPtype
1350      Debug.Print "'FREQ  : " & lngCurrVFreq & " hz  " & strFreqtype
1360      Debug.Print "'ASP X : " & lngAspectX
1370      Debug.Print "'ASP Y : " & lngAspectY
1380      Debug.Print "'ASP XY: " & lngHyp
1390      Debug.Print "'PPI X : " & lngPpiX & " pixels per inch"
1400      Debug.Print "'PPI Y : " & lngPpiY & " pixels per inch"
1410      Debug.Print "'DESK X: " & lngDeskWidth & " pixels"
1420      Debug.Print "'DESK Y: " & lngDeskHeight & " pixels"

          ' ** Delta Data's Acer monitor.
          ' **   RES X : 1440 pixels
          ' **   RES Y : 900 pixels
          ' **   BPP   : 32 bits per pixel  (True Color)
          ' **   FREQ  : 60 hz  (User-selected)
          ' **   ASP X : 36
          ' **   ASP Y : 36
          ' **   ASP XY: 51
          ' **   PPI X : 96 pixels per inch
          ' **   PPI Y : 96 pixels per inch
          ' **   DESK X: 1440 pixels
          ' **   DESK Y: 900 pixels

1430    End If

EXITP:
1440    GetScreenRes = lngRetVal
1450    Exit Function

ERRH:
1460    lngRetVal = 0&
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Function

Public Function GetFrmDim(frm As Access.Form) As String

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFrmDim"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim strRetVal As String

1610    strRetVal = vbNullString

1620    If IsLoaded(frm.Name, acForm) = True Then  ' ** Module Function: modFileUtilities.
1630      GetFormDimensions frm, lngLeft, lngTop, lngWidth, lngHeight  ' ** Function: Below.
1640      Debug.Print "'LEFT: " & CStr(lngLeft) & "  TOP: " & CStr(lngTop) & "  WIDTH: " & CStr(lngWidth) & "  HEIGHT: " & CStr(lngHeight)
1650    Else
1660      Debug.Print "'FORM NOT FOUND! " & frm.Name
1670    End If

        ' ** 1440 Twips/Inch.

        ' ** The Twips/Pixel changes depending on screen resolution NOT font size.
        ' ** There are always 12000 X 9000 Twips on a screen.
        ' ** Therefore at a resolution of 800 X 600:
        ' **   By Width:  12000/800  = 15 Twips/Pixel
        ' **   By Height:  9000/600  = 15 Twips/Pixel
        ' ** NO, THIS DOESN'T SEEM TO WORK FOR MY SCREEN!
        ' ** I reliably work with the 15 Twips/Pixel value, but my
        ' ** screen resolution of 1440/900 would give a result of:
        ' **   By Width:  12000/1440 = 8.33 Twips/Pixel  !?
        ' **   By Height:  9000/900  = 10   Twips/Pixel  !?

        ' ** 12000/1440 = 8.333
        ' ** base/width = tpp
        ' ** 1440(12000/1440) = 1440(8.333)
        ' ** width(base/width) = width * tpp
        ' ** 12000 = 12000
        ' ** base = width * tpp

        ' ** 9000/900 = 10
        ' ** 900(9000/900) = 900(10)
        ' ** 9000 = 9000

EXITP:
1680    GetFrmDim = strRetVal
1690    Exit Function

ERRH:
1700    Select Case ERR.Number
        Case Else
1710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1720    End Select
1730    Resume EXITP

End Function

Public Sub GetFormDimensions(frm As Access.Form, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)
'*********************************************************************
' ** PURPOSE: Returns the left, top, width, and height
' **          measurements of a form window in twips.
' ** ARGUMENTS:
' **   frm: The form object whose measurements are to be determined.
' **   lngLeft, lngTop, lngWidth, lngHeight: Measurement variables
' **   that will return the dimensions of form frm 'by reference'.
' ** NOTE: The lngWidth and lngHeight values will be equivalent
' **   to those provided by the form WindowWidth and WindowHeight
' **   properties.
'*********************************************************************

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFormDimensions"

        Dim typRECT_Frm As RECT
        Dim typRECT_Cli As RECT
        Dim lngCliLeft As Long
        Dim lngCliTop  As Long

        ' ** Get the screen coordinates and window size of the form.
        ' ** The screen coordinates are returned in pixels measured
        ' ** from the upper-left corner of the screen.
1810    GetWindowRect frm.hwnd, typRECT_Frm  ' ** API Function: Above.
1820    lngLeft = typRECT_Frm.Left
1830    lngTop = typRECT_Frm.Top
1840    lngWidth = typRECT_Frm.Right - typRECT_Frm.Left
1850    lngHeight = typRECT_Frm.Bottom - typRECT_Frm.Top

        ' ** Convert the measurements from pixels to twips.
1860    ConvertPIXELSToTWIPS lngLeft, lngTop
1870    ConvertPIXELSToTWIPS lngWidth, lngHeight

        ' ** If the form is not a pop-up form, adjust the screen
        ' ** coordinates to measure from the top of the Microsoft
        ' ** Access typRECT_Cli window. Position 0,0 for a pop-up form
        ' ** is the upper-left corner of the screen, whereas position
        ' ** 0,0 for a normal window is the upper-left corner of the
        ' ** Microsoft Access client window below the menu bar.
1880    If GetWindowClass(frm.hwnd) <> "OFormPopup" Then

          ' ** Get the screen coordinates and window size of the
          ' ** typRECT_Cli window.
1890      GetWindowRect GetParent(frm.hwnd), typRECT_Cli  ' ** API Function: Above.
1900      lngCliLeft = typRECT_Cli.Left
1910      lngCliTop = typRECT_Cli.Top
1920      ConvertPIXELSToTWIPS lngCliLeft, lngCliTop

          ' ** Adjust the form dimensions from the typRECT_Cli
          ' ** measurements.
1930      lngLeft = lngLeft - lngCliLeft
1940      lngTop = lngTop - lngCliTop

1950    End If

        'Debug.Print "'lngLeft: " & CStr(lngLeft)
        'Debug.Print "'lngTop: " & CStr(lngTop)
        'Debug.Print "'lngWidth: " & CStr(lngWidth)
        'Debug.Print "'lngHeight: " & CStr(lngHeight)

EXITP:
1960    Exit Sub

ERRH:
1970    DoCmd.Hourglass False
1980    Select Case ERR.Number
        Case Else
1990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2000    End Select
2010    Resume EXITP

End Sub

Public Sub GetAppDimensions(lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "GetAppDimensions"

        Dim typRECT_Frm As RECT
        Dim typRECT_Cli As RECT
        Dim lngCliLeft As Long
        Dim lngCliTop  As Long

2110    GetWindowRect Application.hWndAccessApp, typRECT_Frm   ' ** API Function: Above.
2120    lngLeft = typRECT_Frm.Left
2130    lngTop = typRECT_Frm.Top
2140    lngWidth = typRECT_Frm.Right - typRECT_Frm.Left
2150    lngHeight = typRECT_Frm.Bottom - typRECT_Frm.Top

2160    ConvertPIXELSToTWIPS lngLeft, lngTop
2170    ConvertPIXELSToTWIPS lngWidth, lngHeight

2180    If GetWindowClass(Application.hWndAccessApp) <> "OFormPopup" Then

2190      GetWindowRect GetParent(Application.hWndAccessApp), typRECT_Cli  ' ** API Function: Above.
2200      lngCliLeft = typRECT_Cli.Left
2210      lngCliTop = typRECT_Cli.Top

2220      ConvertPIXELSToTWIPS lngCliLeft, lngCliTop

2230      lngLeft = lngLeft - lngCliLeft
2240      lngTop = lngTop - lngCliTop

2250    End If

EXITP:
2260    Exit Sub

ERRH:
2270    Select Case ERR.Number
        Case Else
2280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2290    End Select
2300    Resume EXITP

End Sub

Public Function GetAppDim_Test() As Boolean

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "GetAppDim_Test"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim blnRetVal As Boolean

2410    blnRetVal = True

2420    GetAppDimensions lngLeft, lngTop, lngWidth, lngHeight  ' ** Function: Above.

2430    Debug.Print "'lngTop = " & CStr(lngTop)
2440    Debug.Print "'lngLeft = " & CStr(lngLeft)
2450    Debug.Print "'lngWidth = " & CStr(lngWidth)
2460    Debug.Print "'lngHeight = " & CStr(lngHeight)

        'lngTop = -120
        'lngLeft = -120
        'lngWidth = 21840
        'lngHeight = 13290

EXITP:
2470    GetAppDim_Test = blnRetVal
2480    Exit Function

ERRH:
2490    Select Case ERR.Number
        Case Else
2500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2510    End Select
2520    Resume EXITP

End Function

Public Sub GetWindowBorders(frm As Access.Form, lngXBorders As Long, lngYBorders As Long)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "GetWindowBorders"

        Dim typWINFO As WININFO_TYPE
        Dim typRECT_Win As RECT, typRECT_Cli As RECT
        Dim lngSizeOf As Long
        Dim lngXBdr As Long, lngYBdr As Long
        Dim lngWinLeft As Long, lngWinRight As Long, lngWinTop As Long, lngWinBottom As Long
        Dim lngWinWidth As Long, lngWinHeight As Long
        Dim lngCliLeft As Long, lngCliRight As Long, lngCliTop As Long, lngCliBottom As Long
        Dim lngCliWidth As Long, lngCliHeight As Long
        Dim lngHeadHeight As Long

2610    lngSizeOf = LenB(typWINFO)
2620    typWINFO.cbSize = lngSizeOf

2630    GetWindowRect frm.hwnd, typRECT_Win  ' ** API Function: Above.
2640    GetClientRect frm.hwnd, typRECT_Cli  ' ** API Function: Above.

2650    GetWindowInfo Forms(0).hwnd, typWINFO  ' ** API Function: Above.

2660    lngWinLeft = typRECT_Win.Left
2670    lngWinRight = typRECT_Win.Right
2680    lngWinTop = typRECT_Win.Top
2690    lngWinBottom = typRECT_Win.Bottom
2700    lngWinWidth = lngWinRight - lngWinLeft
2710    lngWinHeight = lngWinBottom - lngWinTop

2720    ConvertPIXELSToTWIPS lngWinLeft, lngWinTop
2730    ConvertPIXELSToTWIPS lngWinWidth, lngWinHeight

2740    lngCliLeft = typRECT_Cli.Left
2750    lngCliRight = typRECT_Cli.Right
2760    lngCliTop = typRECT_Cli.Top
2770    lngCliBottom = typRECT_Cli.Bottom
2780    lngCliWidth = lngCliRight - lngCliLeft
2790    lngCliHeight = lngCliBottom - lngCliTop

2800    ConvertPIXELSToTWIPS lngCliLeft, lngCliTop
2810    ConvertPIXELSToTWIPS lngCliWidth, lngCliHeight

2820    lngXBdr = typWINFO.cxWindowBorders
2830    lngYBdr = typWINFO.cyWindowBorders

2840    ConvertPIXELSToTWIPS lngXBdr, lngYBdr

2850    lngHeadHeight = ((lngWinHeight - lngCliHeight) - (lngYBdr * 2&))
2860    lngXBorders = (lngXBdr * 2&)
2870    lngYBorders = (lngHeadHeight + (lngYBdr * 2&))

        'Debug.Print "'cbSize: " & typWINFO.cbSize
        'Debug.Print "'dwStyle: " & typWINFO.dwStyle
        'Debug.Print "'dwExStyle: " & typWINFO.dwExStyle
        'Debug.Print "'dwWindowStatus: " & typWINFO.dwWindowStatus
        'Debug.Print "'cxWindowBorders: " & lngXBdr
        'Debug.Print "'cyWindowBorders: " & lngYBdr
        'Debug.Print "'Win Left: " & lngWinLeft
        'Debug.Print "'Cli Left: " & lngCliLeft
        'Debug.Print "'Win Width: " & lngWinWidth
        'Debug.Print "'Cli Width: " & lngCliWidth
        'Debug.Print "'Win Top: " & lngWinTop
        'Debug.Print "'Cli Top: " & lngCliTop
        'Debug.Print "'Win Height: " & lngWinHeight
        'Debug.Print "'Cli Height: " & lngCliHeight
        'Debug.Print "'Win Head Height: " & lngHeadHeight  'YES!!!!!!!!!!!!!!

        'cbSize: 64
        'dwStyle: 1455423488
        'dwExStyle: 2368
        'dwWindowStatus: 0
        'cxWindowBorders: 45
        'cyWindowBorders: 45
        'Win Left: 6435
        'Cli Left: 0
        'Win Width: 8730
        'Cli Width: 8640
        'Win Top: 3495
        'Cli Top: 0
        'Win Height: 4545
        'Cli Height: 4065
        'Win Head Height: 390

        ' ** Type WINDOWINFO
        ' **   cbSize           : The size of the structure, in bytes. The caller must set this to sizeof(WINDOWINFO).
        ' **   rcWindow         : Pointer to a RECT structure that specifies the coordinates of the window.
        ' **   rcClient         : Pointer to a RECT structure that specifies the coordinates of the client area.
        ' **   dwStyle          : The window styles. For a table of window styles, see CreateWindowEx.
        ' **   dwExStyle        : The extended window styles. For a table of extended window styles, see CreateWindowEx.
        ' **   dwWindowStatus   : The window status. If this member is WS_ACTIVECAPTION, the window is active. Otherwise, this member is zero.
        ' **   cxWindowBorders  : The width of the window border, in pixels.
        ' **   cyWindowBorders  : The height of the window border, in pixels.
        ' **   atomWindowtype   : The window class atom (see RegisterClass).
        ' **   wCreatorVersion  : The Microsoft Windows version of the application that created the window.

EXITP:
2880    Exit Sub

ERRH:
2890    Select Case ERR.Number
        Case Else
2900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2910    End Select
2920    Resume EXITP

End Sub

Public Sub GetAppBorders(lngXBorders As Long, lngYBorders As Long)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "GetAppBorders"

        Dim typWINFO As WININFO_TYPE
        Dim typRECT_Win As RECT, typRECT_Cli As RECT
        Dim lngSizeOf As Long
        Dim lngXBdr As Long, lngYBdr As Long
        Dim lngWinLeft As Long, lngWinRight As Long, lngWinTop As Long, lngWinBottom As Long
        Dim lngWinWidth As Long, lngWinHeight As Long
        Dim lngCliLeft As Long, lngCliRight As Long, lngCliTop As Long, lngCliBottom As Long
        Dim lngCliWidth As Long, lngCliHeight As Long
        Dim lngHeadHeight As Long

3010    lngSizeOf = LenB(typWINFO)
3020    typWINFO.cbSize = lngSizeOf

3030    GetWindowRect Application.hWndAccessApp, typRECT_Win  ' ** API Function: Above.
3040    GetClientRect Application.hWndAccessApp, typRECT_Cli  ' ** API Function: Above.
3050    GetWindowInfo Application.hWndAccessApp, typWINFO  ' ** API Function: Above.

3060    lngWinLeft = typRECT_Win.Left
3070    lngWinRight = typRECT_Win.Right
3080    lngWinTop = typRECT_Win.Top
3090    lngWinBottom = typRECT_Win.Bottom
3100    lngWinWidth = lngWinRight - lngWinLeft
3110    lngWinHeight = lngWinBottom - lngWinTop

3120    ConvertPIXELSToTWIPS lngWinLeft, lngWinTop
3130    ConvertPIXELSToTWIPS lngWinWidth, lngWinHeight

3140    lngCliLeft = typRECT_Cli.Left
3150    lngCliRight = typRECT_Cli.Right
3160    lngCliTop = typRECT_Cli.Top
3170    lngCliBottom = typRECT_Cli.Bottom
3180    lngCliWidth = lngCliRight - lngCliLeft
3190    lngCliHeight = lngCliBottom - lngCliTop

3200    ConvertPIXELSToTWIPS lngCliLeft, lngCliTop
3210    ConvertPIXELSToTWIPS lngCliWidth, lngCliHeight

3220    lngXBdr = typWINFO.cxWindowBorders
3230    lngYBdr = typWINFO.cyWindowBorders

3240    ConvertPIXELSToTWIPS lngXBdr, lngYBdr

3250    lngHeadHeight = ((lngWinHeight - lngCliHeight) - (lngYBdr * 2&))
3260    lngXBorders = (lngXBdr * 2&)
3270    lngYBorders = (lngHeadHeight + (lngYBdr * 2&))

EXITP:
3280    Exit Sub

ERRH:
3290    Select Case ERR.Number
        Case Else
3300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3310    End Select
3320    Resume EXITP

End Sub

Public Function GetAppBrdrs_Test() As Boolean

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "GetAppBrdrs_Test"

        Dim lngXBorders As Long, lngYBorders As Long
        Dim blnRetVal As Boolean

3410    blnRetVal = True

3420    GetAppBorders lngXBorders, lngYBorders  ' ** Function: Above.

3430    Debug.Print "'lngXBorders = " & CStr(lngXBorders)
3440    Debug.Print "'lngYBorders = " & CStr(lngYBorders)

        'lngXBorders = 240
        'lngYBorders = 570

EXITP:
3450    GetAppBrdrs_Test = blnRetVal
3460    Exit Function

ERRH:
3470    Select Case ERR.Number
        Case Else
3480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3490    End Select
3500    Resume EXITP

End Function

Public Sub ConvertPIXELSToTWIPS(ByRef lngX As Long, ByRef lngY As Long)
'***************************************************************
' ** PURPOSE: Converts the two pixel measurements passed as
' **          arguments to twips.
' ** ARGUMENTS:
' **    X, Y: Measurement variables in pixels. These will be
' **          converted to twips and returned through the same
' **          variables 'by reference'.
'***************************************************************

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "ConvertPixelsToTwips"

        Dim lngHdc As Long, lngRetVal As Long
        Dim lngXPixelsPerInch As Long, lngYPixelsPerInch As Long

        ' ** Retrieve the current number of pixels per inch, which is
        ' ** resolution-dependent.
3610    lngHdc = GetDC(0)
3620    lngXPixelsPerInch = GetDeviceCaps(lngHdc, GSR_LOGPIXELSX)
3630    lngYPixelsPerInch = GetDeviceCaps(lngHdc, GSR_LOGPIXELSY)
3640    lngRetVal = ReleaseDC(0, lngHdc)

        ' ** Compute and return the measurements in twips.
3650    lngX = (lngX / lngXPixelsPerInch) * RZ_TWIPSPERINCH
3660    lngY = (lngY / lngYPixelsPerInch) * RZ_TWIPSPERINCH

EXITP:
3670    Exit Sub

ERRH:
3680    Select Case ERR.Number
        Case Else
3690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3700    End Select
3710    Resume EXITP

End Sub

Public Function GetAppName(lngHWnd As Long) As String
' ** This function returns the Caption Text of each window passed to it. If a window
' ** does not have a Caption bar, then this function returns a zero-length string ("").

        ' *******************************************************************************************************
        ' ** NOTE: DIRECTORY WINDOWS, LIKE EXPLORER, MAY HAVE THE SAME NAME AS AN APPLICATION IF IT HAPPENS
        ' **       TO BE ON A DIRECTORY OF THAT NAME. FOR Trust Accountant, USE IsTAOpen() IN modFileUtilities.
        ' *******************************************************************************************************

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "GetAppName"

        Dim lngResult As Long
        Dim strWinText As String * 255
        Dim strReturn As String

3810    lngResult = GetWindowText(lngHWnd, strWinText, 255)
3820    strReturn = Left(strWinText, lngResult)

EXITP:
3830    GetAppName = strReturn
3840    Exit Function

ERRH:
3850    strReturn = vbNullString
3860    Select Case ERR.Number
        Case Else
3870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3880    End Select
3890    Resume EXITP

End Function

Private Function GetWindowClass(lngHWnd As Long) As String
'*************************************************************
' PURPOSE: Retrieve the class of the passed window handle.
' ARGUMENTS:
'    lngHWnd: The window handle whose class is to be retrieved.
' RETURN:
'    The window class name.
'*************************************************************

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "GetWindowClass"

        Dim strBuff As String
        Dim intBuffSize As Integer

3910    strBuff = String$(255, " ")
3920    intBuffSize = GetClassName(lngHWnd, strBuff, 255)  ' ** API Function: Above.

EXITP:
3930    GetWindowClass = Left(strBuff, intBuffSize)
3940    Exit Function

ERRH:
3950    Select Case ERR.Number
        Case Else
3960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3970    End Select
3980    Resume EXITP

End Function

Public Function SizeAccess(lngLeft As Long, lngTop As Long, lngHeight As Long, lngWidth As Long) As Boolean

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "SizeAccess"

        Dim lngHWnd As Long
        Dim blnRetVal As Boolean

4010    blnRetVal = True

        ' ** Get handle to Microsoft Access.
4020    lngHWnd = Application.hWndAccessApp

        ' ** Position Microsoft Access.
4030    SetWindowPos lngHWnd, HWND_TOP, lngLeft, lngTop, lngWidth, lngHeight, SWP_NOZORDER  ' ** API Function: Above.

EXITP:
4040    SizeAccess = blnRetVal
4050    Exit Function

ERRH:
4060    Select Case ERR.Number
        Case Else
4070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4080    End Select
4090    Resume EXITP

End Function

Public Function GetWinPos_test() As Boolean

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "GetWinPos_test"

        Dim frmPost As Access.Form
        Dim lngLeft As Long
        Dim lngTop As Long
        Dim lngWidth As Long
        Dim lngHeight  As Long
        Dim blnRetVal As Boolean

4110    blnRetVal = True

4120    Set frmPost = Forms![frmMenu_Post]

4130    GetFormDimensions frmPost, lngLeft, lngTop, lngWidth, lngHeight

4140    MsgBox "frmMenu_Post is:" & vbCrLf & "Left: " & CStr(lngLeft) & vbCrLf & "Top: " & _
          CStr(lngTop) & vbCrLf & "Width: " & CStr(lngWidth) & vbCrLf & "Height: " & CStr(lngHeight), _
          vbInformation + vbOKOnly, "Window Information"

EXITP:
4150    Set frmPost = Nothing
4160    GetWinPos_test = blnRetVal
4170    Exit Function

ERRH:
4180    Select Case ERR.Number
        Case Else
4190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4200    End Select
4210    Resume EXITP

End Function

Public Function GetCountOfWindows(lngHWnd As Long, strAppCaption As String) As Long
' ** This function counts all instances of an application that
' ** are open, including any windows that are not visible.
' ** Arguments: LngHwnd        = Any valid window handle.
' **            StrAppCaption  = The window caption to search for.
' ** Example:   GetCountOfWindows(hWndAccessApp,"Microsoft Access")
' ** 07/04/2009: CURRENTLY NOT USED!

        ' *******************************************************************************************************
        ' ** NOTE: DIRECTORY WINDOWS, LIKE EXPLORER, MAY HAVE THE SAME NAME AS AN APPLICATION IF IT HAPPENS
        ' **       TO BE ON A DIRECTORY OF THAT NAME. FOR Trust Accountant, USE IsTAOpen() IN modFileUtilities.
        ' *******************************************************************************************************

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "GetCountOfWindows"

        Dim lngResult As Long
        Dim lngICount As Long
        Dim strAppName As String

4310    lngICount = 0&
4320    lngResult = GetWindow(lngHWnd, GW_HWNDFIRST)
4330    Do Until lngResult = 0
4340      If IsWindowVisible(lngResult) Then
4350        strAppName = GetAppName(lngResult)  ' ** Function: Above.
4360        If InStr(1, strAppName, strAppCaption) Then
4370          lngICount = lngICount + 1
4380        End If
4390      End If
4400      lngResult = GetWindow(lngResult, GW_HWNDNEXT)
4410    Loop

EXITP:
4420    GetCountOfWindows = lngICount
4430    Exit Function

ERRH:
4440    lngICount = 0&
4450    Select Case ERR.Number
        Case Else
4460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4470    End Select
4480    Resume EXITP

End Function

Public Sub Win_Mod_Restore()
' ** Restore the Module Window so that the Immediate Window is visible.

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Win_Mod_Restore"

4510    With Application.VBE.ActiveWindow  ' ** Needed this syntax!
4520      .WindowState = 0 'vbext_ws_Normal  ' ** Constant part of Microsoft Visual Basic for Applications Extensibility 5.3.
          ' ** Vbext_WindowState enumeration:
          ' **   0  vbext_ws_Normal    Normal (Default)
          ' **   1  vbext_ws_Minimize  Minimized (minimized to an icon)
          ' **   2  vbext_ws_Maximize  Maximized (enlarged to maximum size)
4530    End With

EXITP:
4540    Exit Sub

ERRH:
4550    Select Case ERR.Number
        Case 91  ' ** Object variable or With block variable not set.
          ' ** If there technically isn't an active VBA window, then it will error.
4560    Case Else
4570      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4580    End Select
4590    Resume EXITP

End Sub

Public Function Win_Resize(ByRef frm As Access.Form, ByRef lngLeft As Long, ByRef lngTop As Long, ByRef lngWidth As Long, ByRef lngHeight As Long, ByRef lngMyHeight As Long, ByRef lngMyWidth As Long) As Boolean
' ** Resize the calling form if necessary.

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Win_Resize"

        Dim lngTpp As Long, lngPpi As Long
        Dim blnRetVal As Boolean

4610    blnRetVal = False

4620    GetFormDimensions frm, lngLeft, lngTop, lngWidth, lngHeight  ' ** Function: Below.
        ' ** Mine shows 6 rows: lngHeight / lngTpp = 379 Pixels high, or 5685 Twips.
        'lngTpp = GetTPP  ' ** Function: Below.
4630    lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4640    lngPpi = GetScreenRes(GSR_LOGPIXELSX)  ' ** Function: Below.
4650    With frm
4660      If .Name = "frmMap_Div_Detail" Then
4670        .WinLeft.Visible = False
4680        .WinTop.Visible = False     ' **  frmMap_Div_Detail   frmIncomeExpenseCodes
4690        .WinWidth.Visible = False   ' ** ==================  ============
4700        .WinHeight.Visible = False  ' **  Rich's      My          My
4710        .WinTPP.Visible = False     ' ** computer  computer    computer
4720        .WinPPI.Visible = False     ' ** ========  ========  ============
4730        .WinLeft = lngLeft          ' **    900      5700        6570
4740        .WinTop = lngTop            ' **    555      1980        1200
4750        .WinWidth = lngWidth        ' **  10185     10200        8445
4760        .WinHeight = lngHeight      ' **   5580      5685        8070
4770        .WinTPP = lngTpp            ' **      8        15          15
4780        .WinPPI = lngPpi            ' **     96        96          96
4790      End If
4800    End With
4810    If lngTpp = lngMyTpp Then
          ' ** Their Tpp same as my Tpp.
4820      If (lngHeight <> lngMyHeight) Or (lngWidth <> lngMyWidth) Then
            ' ** On my screen, the window is 379 Pixels high, or 5685 Twips.
4830        blnRetVal = True
4840        lngHeight = lngMyHeight
4850        lngWidth = lngMyWidth
            'DoCmd.MoveSize lngLeft, lngTop, lngMyWidth, lngMyHeight
4860      End If
4870    Else
          ' ** Their Tpp different than my Tpp.
4880      If (lngHeight <> (lngMyHeight * (lngMyTpp / lngTpp))) Or (lngWidth <> (lngMyWidth * (lngMyTpp / lngTpp))) Then
            ' ** AHA!
            ' ** (MyTpp / lngTpp)     : 1.875
            ' ** (lngHeight * 1.875)  : 10462.5
4890        blnRetVal = True
4900        lngHeight = (lngMyHeight * (lngMyTpp / lngTpp))
4910        lngWidth = (lngMyWidth * (lngMyTpp / lngTpp))
            'DoCmd.MoveSize lngLeft, lngTop, (lngMyWidth * (lngMyTpp / lngTpp)), (lngMyHeight * (lngMyTpp / lngTpp))
4920      End If
4930    End If

EXITP:
4940    Win_Resize = blnRetVal
4950    Exit Function

ERRH:
4960    blnRetVal = False
4970    Select Case ERR.Number
        Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Function

Public Function Win_List_Open() As Boolean
' ** Not called; just for Immediate Window.

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "Win_List_Open"

        Dim lngParam As Long
        Dim blnDoChild As Boolean
        Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean, lngRetVal As Long

5110    blnRetVal = True

5120    blnDoChild = True

5130    lngParam = 0&
5140    glngClasses = 0&
5150    ReDim garr_varClass(CLS_ELEMS, 0)

5160    blnWindowVisible = True  ' ** Only check visible windows.

        ' ** Open Windows:
        'Afx:00400000:8:00010003:00000000:0006084B : Jasc Paint Shop Pro - Image1
        'BasicWindow : Weather
        'GestureFeedbackAnimationWindow :
        'Internet Explorer_Hidden :
        'Internet Explorer_Hidden :
        'OMain : Trust Accountant™
        'rctrl_renwnd32 : Inbox - victor.campbell@q.com - Microsoft Outlook
        'Solitaire : Solitaire
        'TextPad4 : TextPad - [Document1 *]
        'tooltips_class32 :
        'wndclass_desked_gsk : Microsoft Visual Basic for Applications - Trust [running]
        'CNT: 11

        ' ** Enumerate the list.
5170    lngRetVal = EnumWindows(AddressOf Win_Class_Load, lngParam)  ' ** API Function: Above.

        ' ** Binary Sort garr_varClass() array.
5180    For lngX = UBound(garr_varClass, 2) To 1& Step -1&
5190      For lngY = 0& To (lngX - 1&)
5200        If garr_varClass(CLS_CLASS, lngY) > garr_varClass(CLS_CLASS, (lngY + 1&)) Then
5210          strTmp01 = garr_varClass(CLS_CLASS, lngY)
5220          strTmp02 = garr_varClass(CLS_TITLE, lngY)
5230          garr_varClass(CLS_CLASS, lngY) = garr_varClass(CLS_CLASS, (lngY + 1&))
5240          garr_varClass(CLS_TITLE, lngY) = garr_varClass(CLS_TITLE, (lngY + 1&))
5250          garr_varClass(CLS_CLASS, (lngY + 1&)) = strTmp01
5260          garr_varClass(CLS_TITLE, (lngY + 1&)) = strTmp02
5270        End If
5280      Next
5290    Next

5300    Select Case blnDoChild
        Case True

5310      strTmp01 = "OMain"
5320      For lngX = 0& To (glngClasses - 1&)
5330        If garr_varClass(CLS_CLASS, lngX) = strTmp01 Then
5340          lngTmp03 = garr_varClass(CLS_HWND, lngX)
5350          Exit For
5360        End If
5370      Next

5380      lngParam = 0&
5390      glngClasses = 0&
5400      ReDim garr_varClass(CLS_ELEMS, 0)

5410      lngRetVal = EnumChildWindows(lngTmp03, AddressOf Win_Class_Load, lngParam)  ' ** API Function: Above.
5420      lngTmp03 = 0&
5430      For lngX = 0& To (glngClasses - 1&)
5440        lngTmp03 = lngTmp03 + 1&
5450        Debug.Print "'" & garr_varClass(CLS_CLASS, lngX) & " : " & garr_varClass(CLS_TITLE, lngX)
5460      Next

5470    Case False
5480      lngTmp03 = 0&
5490      For lngX = 0& To (glngClasses - 1&)
5500        If InStr(garr_varClass(CLS_CLASS, lngX), "G2M_") > 0 Or _
                InStr(garr_varClass(CLS_CLASS, lngX), "Shell_TrayWnd") > 0 Or _
                InStr(garr_varClass(CLS_CLASS, lngX), "Progman") > 0 Or _
                InStr(garr_varClass(CLS_CLASS, lngX), "LogoffMonitorThread") > 0 Or _
                InStr(garr_varClass(CLS_CLASS, lngX), "xxx") > 0 Or _
                InStr(garr_varClass(CLS_CLASS, lngX), "xxx") > 0 Then
              ' ** Skip.
5510        ElseIf InStr(garr_varClass(CLS_TITLE, lngX), "YYY") > 0 Then
              ' ** Skip.
5520        Else
5530          lngTmp03 = lngTmp03 + 1&
5540          Debug.Print "'" & garr_varClass(CLS_CLASS, lngX) & " : " & garr_varClass(CLS_TITLE, lngX)
5550        End If
5560      Next
5570    End Select

5580    Debug.Print "'CNT: " & CStr(lngTmp03)

EXITP:
5590    Win_List_Open = blnRetVal
5600    Exit Function

ERRH:
5610    blnRetVal = False
5620    Select Case ERR.Number
        Case Else
5630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5640    End Select
5650    Resume EXITP

End Function

Public Function Win_List_Open_Child(Optional varFromCode As Variant) As Boolean
' ** Not called; just for Immediate Window.

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "Win_List_Open_Child"

        Dim lngHWnd As Long
        Dim blnRetVal As Boolean

5710    blnRetVal = True
5720    blnNavPaneOpen = False

5730    lngHWnd = Application.hWndAccessApp

5740    Select Case IsMissing(varFromCode)
        Case True
5750      blnFromCode = False
5760    Case False
5770      blnFromCode = CBool(varFromCode)
5780    End Select

        ' ** One call directly to list parent.
5790    EnumChildWindow lngHWnd, 0  ' ** Function: Below.
        ' ** Then list the children.
5800    EnumChildWindows lngHWnd, AddressOf EnumChildWindow, 0  ' ** API Function: Above.

5810    If blnFromCode = True Then
5820      blnRetVal = blnNavPaneOpen
5830    End If

        'Enum 1836138, True  OMain, 'Trust Accountant™'
        'Enum 525688,  False MsoCommandBarDock, 'MsoDockLeft'
        'Enum 1442756, False MsoCommandBarDock, 'MsoDockRight'
        'Enum 6292994, False MsoCommandBar, 'Property Sheet'
        'Enum 1311874, True  MsoCommandBarDock, 'MsoDockTop'
        'Enum 394508,  True  MsoCommandBar, 'Ribbon'
        'Enum 1902500, True  MsoWorkPane, 'Ribbon'
        'Enum 2557536, True  NUIPane
        'Enum 656796,  True  NetUIHWND
        'Enum 1902242, False NetUICtrlNotifySink
        'Enum 2753908, False NetUICtrlNotifySink
        'Enum 1507918, True  MsoCommandBarDock, 'MsoDockBottom'
        'Enum 853170,  True  MsoCommandBar, 'Status Bar'
        'Enum 1181070, True  MsoWorkPane, 'Status Bar'
        'Enum 525854,  True  NUIPane
        'Enum 591256,  True  NetUIHWND
        'Enum 1050014, True  MDIClient
        'Enum 2032758, False MsoWorkPane, 'MsoWorkPane'
        'Enum 1442632, False MsoWorkPane, 'MsoWorkPane'
        'Enum 2164282, True  NetUINativeHWNDHost, 'Navigation Pane Host'
        'Enum 1049550, True  NetUIHWND
        'Enum 2163938, False NetUICtrlNotifySink
        'Enum 2229328, False RICHEDIT60W, 'Search...'
        'Enum 2819114, True  NetUINativeHWNDHost, 'ODocTabs'
        'Enum 1836314, True  NetUIHWND

EXITP:
5840    Win_List_Open_Child = blnRetVal
5850    Exit Function

ERRH:
5860    blnRetVal = False
5870    Select Case ERR.Number
        Case Else
5880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5890    End Select
5900    Resume EXITP

End Function

Public Function EnumChildWindow(ByVal lngChild As Long, ByVal lngParam As Long) As Long

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "EnumChildWindow"

        Dim strClass As String, strText As String
        Dim strMsg As String, strIsVis As String
        Dim lngRetVal As Long, intRetVal As Integer

6010    lngRetVal = 1&  ' ** Continue enumeration.

6020    strClass = Space(64)
6030    intRetVal = GetClassName(lngChild, strClass, 63)  ' ** API Function: Above.
6040    strClass = Left(strClass, intRetVal)

6050    strText = Space(256)
6060    intRetVal = SendMessageS(lngChild, WM_GETTEXT, 255, strText)  ' ** API Function: Above.
6070    strText = Left(strText, intRetVal)

6080    strIsVis = Format(IsWindowVisible(lngChild), "True/False")

6090    strMsg = "'Enum " & Left(CStr(lngChild) & "," & Space(9), 9) & Left(strIsVis & Space(6), 6) & strClass

6100    If Len(strText) > 0 Then
6110      strMsg = strMsg & ", '" & strText & "'"
6120    End If

6130    Select Case blnFromCode
        Case True
6140      If strClass = "NetUINativeHWNDHost" And strText = "Navigation Pane Host" Then
6150        Select Case strIsVis
            Case "True"
6160          blnNavPaneOpen = True
6170        Case "False"
6180          blnNavPaneOpen = False
6190        End Select
6200      End If
6210    Case False
6220      Debug.Print strMsg
6230    End Select

EXITP:
6240    EnumChildWindow = lngRetVal
6250    Exit Function

ERRH:
6260    lngRetVal = 0
6270    Select Case ERR.Number
        Case Else
6280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6290    End Select
6300    Resume EXITP

End Function

Private Function Win_Class_Load(ByVal lngHWnd As Long, ByVal lngParam As Long) As Boolean
' ** Called by:
' **   Win_List_Open(), Above.

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "Win_Class_Load"

        Dim strClass As String, strTitle1 As String
        Dim strClassBuf As String * 255, strTitle1Buf As String * 255
        Dim blnFound As Boolean
        Dim intVis As Integer
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

6410    blnRetVal = True

6420    strClass = GetClassName(lngHWnd, strClassBuf, 255)  ' ** API Function: Above.
6430    strClass = StripNulls(strClassBuf)  ' remove extra Nulls & spaces
6440    strTitle1 = GetWindowText(lngHWnd, strTitle1Buf, 255)
6450    strTitle1 = StripNulls(strTitle1Buf)

6460    If blnWindowVisible = False Then
          ' ** Check both visible and hidden windows.
6470      intVis = 1
6480    Else
          ' ** Check only visible windows.
6490      intVis = IsWindowVisible(lngHWnd)
6500    End If

        ' ** Check if Window is a parent and visible.
6510    If GetParent(lngHWnd) = 0 And intVis = 1 Then
6520      blnFound = False
6530      For lngX = 0& To (glngClasses - 1)
6540        If garr_varClass(CLS_CLASS, lngX) = strClass And _
                garr_varClass(CLS_TITLE, lngX) = strTitle1 And _
                garr_varClass(CLS_HWND, lngX) = lngHWnd Then
6550          blnFound = True
6560          Exit For
6570        End If
6580      Next
6590      If blnFound = False Then
6600        glngClasses = glngClasses + 1&
6610        lngE = glngClasses - 1&
6620        ReDim Preserve garr_varClass(CLS_ELEMS, lngE)
6630        garr_varClass(CLS_CLASS, lngE) = strClass
6640        garr_varClass(CLS_TITLE, lngE) = strTitle1
6650        garr_varClass(CLS_HWND, lngE) = lngHWnd
6660        garr_varClass(CLS_PARENT, lngE) = GetParent(lngHWnd)  ' ** API Function: modWindowFunctions.
6670        garr_varClass(CLS_VISIBLE, lngE) = IsWindowVisible(lngHWnd)  ' ** API Function: modWindowFunctions.
6680      End If
6690    End If

EXITP:
6700    Win_Class_Load = blnRetVal
6710    Exit Function

ERRH:
6720    blnRetVal = False
6730    Select Case ERR.Number
        Case Else
6740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6750    End Select
6760    Resume EXITP

End Function

Public Function StripNulls(strOriginal As String) As String
' ** Remove extra Nulls so String comparisons will work.

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "StripNulls"

        Dim intPos01 As Integer
        Dim strRetVal As String

6810    strRetVal = vbNullString

6820    intPos01 = InStr(strOriginal, Chr(0))
6830    Do While intPos01 > 0
6840      If intPos01 = Len(strOriginal) Then
6850        strOriginal = Left(strOriginal, (intPos01 - 1))
6860      Else
6870        strOriginal = Left(strOriginal, (intPos01 - 1)) & Mid(strOriginal, (intPos01 + 1))
6880      End If
6890      intPos01 = InStr(strOriginal, Chr(0))
6900    Loop
6910    strRetVal = strOriginal

EXITP:
6920    StripNulls = strRetVal
6930    Exit Function

ERRH:
6940    strRetVal = vbNullString
6950    Select Case ERR.Number
        Case Else
6960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6970    End Select
6980    Resume EXITP

End Function

Public Function CmdBars_Design() As Boolean
' ** Called from AutoKeys, using Ctrl+Shift+D.

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBars_Design"

        Dim blnRetVal As Boolean

7010    blnRetVal = True

7020    If CommandBars("Formatting (Form/Report)").Visible = False Then
7030      CommandBars("Formatting (Form/Report)").Visible = True
7040    End If

EXITP:
7050    CmdBars_Design = blnRetVal
7060    Exit Function

ERRH:
7070    Select Case ERR.Number
        Case Else
7080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7090    End Select
7100    Resume EXITP

End Function

Public Function CmdBars_Database(blnHide As Boolean) As Boolean

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBars_Database"

        Dim blnSkip As Boolean
        Dim blnRetVal As Boolean

7210    blnRetVal = True
7220    blnSkip = True

7230    If blnSkip = False Then
7240      Select Case blnHide
          Case True
7250        If CommandBars("Database").Visible = True Then
7260          CommandBars("Database").Visible = False
7270        End If
7280      Case False
7290        If CommandBars("Database").Visible = False Then
7300          CommandBars("Database").Visible = True
7310        End If
7320      End Select
7330    End If

EXITP:
7340    CmdBars_Database = blnRetVal
7350    Exit Function

ERRH:
7360    Select Case ERR.Number
        Case Else
7370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7380    End Select
7390    Resume EXITP

End Function

Public Sub CmdBars_Hide(blnHide As Boolean)
' ** Currently called from Form_Timer() by:
' **   frmJournal_Columns
' **   frmMasterBalance
' **   frmMenu_Title
' **   frmRecurringItems
' **   frmRpt_TransactionByType
' **   frmTransaction_Audit
' **   frmXAdmin_FileInfo
' **   frmXAdmin_Graphics
' **   frmXAdmin_Misc
' **   frmXAdmin_Registry
' **   frmXAdmin_Version

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBars_Hide"

        Dim blnSkip As Boolean

        ' ** When this is distributed, the Options for 'AllowFullMenus', 'AllowShortcutMenus',
        ' ** and 'AllowBuiltInToolbars' have all been set to False, so users don't see those.
        ' ** When I'm running it, the 'Menu Bar' and various other CommandBars
        ' ** do show up, but they won't when the user opens those forms.

7410    blnSkip = False

7420    Select Case blnHide
        Case True
7430      If blnSkip = False Then  'Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
7440        gblnCBarVis = False
            'If CommandBars("Form View").Controls("Properties").state = msoButtonDown Then
            '  ' ** Close the Properties window, if open.
            '  CommandBars("Form View").Controls("Properties").Execute
            'End If
            'If CommandBars("Database").Visible = True Then CommandBars("Database").Visible = False
            'If CommandBars("Form View").Visible = True Then CommandBars("Form View").Visible = False
7450        If CommandBars("Formatting (Form/Report)").Visible = True Then
7460          CommandBars("Formatting (Form/Report)").Visible = False
7470          gblnCBarVis = True
7480        End If
7490        If CommandBars("Formatting (Datasheet)").Visible = True Then
7500          CommandBars("Formatting (Datasheet)").Visible = False
7510          gblnCBarVis = True
7520        End If
7530      End If

7540    Case False
7550      If blnSkip = False Then
            'If gblnCBarVis = True Then
            '  CommandBars("Formatting (Form/Report)").Visible = True
            '  gblnCBarVis = False
            'End If
7560      End If
7570    End Select
        ' ** MsoButtonState enumeration:
        ' **   -1 msoButtonDown
        ' **    0 msoButtonUp
        ' **    2 msoButtonMixed

EXITP:
7580    Exit Sub

ERRH:
7590    Select Case ERR.Number
        Case Else
7600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7610    End Select
7620    Resume EXITP

End Sub

Public Sub CmdBars_Prop()
' ** Close the Properties popup if it's open.
' ** I don't yet know why some forms leave it showing and others don't.
' ** It may have to do with forms created from scratch vs. those
' ** created much earlier, perhaps with an earlier version of
' ** Access or even one created without the latest service pack.

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBars_Prop"

        Dim cbar As Office.CommandBar, cbctl As Office.CommandBarControl
        Dim blnSkip As Boolean

7710    blnSkip = True

7720    If blnSkip = False Then
7730      Set cbar = CommandBars![Form Design]
7740      With cbar
7750        For Each cbctl In cbar.Controls
7760          With cbctl
7770            If .Caption = "&Properties" Then
7780              If .state = msoButtonDown Then
                    ' ** MsoButtonState enumeration:
                    ' **    0  msoButtonUp
                    ' **   -1  msoButtonDown
                    ' **    2  msoButtonMixed
7790                .Execute
7800                Exit For
7810              End If
7820            End If
7830          End With
7840        Next
7850      End With
7860    End If

EXITP:
7870    Set cbctl = Nothing
7880    Set cbar = Nothing
7890    Exit Sub

ERRH:
7900    Select Case ERR.Number
        Case Else
7910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7920    End Select
7930    Resume EXITP

End Sub

Public Function CmdBars_Visible() As Boolean
' ** Show which command bars are currently visible.

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBars_Visible"

        Dim cbar As Office.CommandBar
        Dim blnRetVal As Boolean

8010    blnRetVal = True

8020    For Each cbar In CommandBars
8030      With cbar
8040        If .Visible = True Then
8050          Debug.Print "'CBAR: " & .Name
8060        End If
8070      End With
8080    Next

EXITP:
8090    CmdBars_Visible = blnRetVal
8100    Exit Function

ERRH:
8110    Select Case ERR.Number
        Case Else
8120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8130    End Select
8140    Resume EXITP

End Function

Public Function CmdBars_Create() As Boolean

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBars_Create"

        Dim cbar As Office.CommandBar, cbctls As Office.CommandBarControls, cbctl As Office.CommandBarControl, cbbtn As Office.CommandBarButton
        Dim lngCtls As Long
        Dim strTmp01 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

8210    blnRetVal = True

8220    Set cbar = CommandBars("TAReports")
8230    With cbar
8240      Set cbctls = .Controls
8250      With cbctls
            'Set cbbtn = .Add(msoControlButton)
            '' ** Add([Type], [Id], [Parameter], [Before], [Temporary]) As Office.CommandBarControl
8260        lngCtls = .Count
8270        For lngX = 1& To lngCtls
8280          Set cbctl = .item(lngX)
8290          With cbctl
8300            Debug.Print "'" & CStr(lngX) & "  " & Left(.Caption & Space(16), 16) & _
                  "  TYPE: " & Left(CmdBars_CtlType(.Type) & Space(18), 18) & "  BUILTIN: " & .BuiltIn  ' ** Function: Below.
8310            If .Type = msoControlButton Then strTmp01 = ("  SC: '" & .ShortcutText & "'") Else strTmp01 = vbNullString
8320            Debug.Print "'   TT: " & Left("'" & .TooltipText & "'" & Space(20), 20) & strTmp01
8330            If .Caption = "Pri&ntZeroDialog" Then
                  'strTmp01 = .DescriptionText
                  'strTmp02 = .TooltipText
                  'strTmp03 = .ShortcutText
                  '.ShortcutText = "Ctrl+P"
                  '.TooltipText = "Pri&nt..."
                  'Exit For
8340            ElseIf .Caption = "&PrintZero" And .HelpContextId = 3843 Then
                  '.TooltipText = "&Print..."
                  'strTmp02 = .TooltipText
                  'strTmp03 = .ShortcutText
8350            ElseIf .Caption = "Pri&nt..." And .HelpContextId = 3106 Then
                  'strTmp01 = .DescriptionText
                  'strTmp02 = .TooltipText
                  'strTmp03 = .ShortcutText  ' ** 'Ctrl+P'
                  'Exit For
8360            End If
                'If .DescriptionText <> vbNullString Then  ' ** None have DescriptionText.
                '  Debug.Print "'" & CStr(lngX) & "  " & .DescriptionText
                'End If
8370          End With
8380        Next
            'With cbbtn
            '  .Caption = "Pri&ntZeroDialog"
            '  .Enabled = True
            '  .OnAction = "mcrPrint_CA_0_Dialog"
            '  .ShortcutText = "Pri&ntZeroDialog"
            '  .TooltipText = "Pri&ntZeroDialog"
            '  .style = msoButtonIcon
            '  .Visible = True
            'End With
8390      End With
8400    End With

8410    Beep

        ' ** 1  &Print...         TYPE: msoControlButton    BUILTIN: True
        ' **    TT: '&Print...'           SC: ''
        ' ** 2  Pri&nt...         TYPE: msoControlButton    BUILTIN: True
        ' **    TT: 'Pri&nt... (Ctrl+P)'  SC: 'Ctrl+P'
        ' ** 3  &PrintZero        TYPE: msoControlButton    BUILTIN: False
        ' **    TT: '&Print...'           SC: ''
        ' ** 4  Pri&ntZeroDialog  TYPE: msoControlButton    BUILTIN: False
        ' **    TT: 'Pri&nt... (Ctrl+P)'  SC: 'Ctrl+P'
        ' ** 5  &Zoom:            TYPE: msoControlComboBox  BUILTIN: True
        ' **    TT: '&Zoom:'
        ' ** 6  &Close            TYPE: msoControlButton    BUILTIN: True
        ' **    TT: '&Close'              SC: ''

        ' ** MsoButton enumeration:
        ' **    0  msoButtonAutomatic
        ' **    1  msoButtonIcon
        ' **    2  msoButtonCaption
        ' **    3  msoButtonIcanAndCaption
        ' **    7  msoButtonIconAndWrapCaption
        ' **   11  msoButtonIconAndCaptionBelow
        ' **   14  msoButtonWrapCaption
        ' **   15  msoButtonIconAndWrapCaptionBelow

        ' ** MsoControlType enumeration:
        ' **    0  msoControlCustom
        ' **    1  msoControlButton
        ' **    2  msoControlEdit
        ' **    3  msoControlDropdown
        ' **    4  msoControlComboBox
        ' **    5  msoControlButtonDropdown
        ' **    6  msoControlSplitDropdown
        ' **    7  msoControlOCXDropdown
        ' **    8  msoControlGenericDropdown
        ' **    9  msoControlGraphicDropdown
        ' **   10  msoControlPopup
        ' **   11  msoControlGraphicPopup
        ' **   12  msoControlButtonPopup
        ' **   13  msoControlSplitButtonPopup
        ' **   14  msoControlSplitButtonMRUPopup
        ' **   15  msoControlLabel
        ' **   16  msoControlExpandingGrind      (&H10)
        ' **   17  msoControlSplitExpandingGrid  (&H11)
        ' **   18  msoControlGrid                (&H12)
        ' **   19  msoControlGauge               (&H13)
        ' **   20  msoControlGraphicCombo        (&H14)
        ' **   21  msoControlPane                (&H15)
        ' **   22  msoControlActiveX             (&H16)

EXITP:
8420    Set cbbtn = Nothing
8430    Set cbctl = Nothing
8440    Set cbctls = Nothing
8450    Set cbar = Nothing
8460    CmdBars_Create = blnRetVal
8470    Exit Function

ERRH:
8480    blnRetVal = False
8490    Select Case ERR.Number
        Case Else
8500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8510    End Select
8520    Resume EXITP

End Function

Public Function CmdBars_CtlType(intType As Integer) As String

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBars_CtlType"

        Dim strRetVal As String

8610    strRetVal = vbNullString

8620    Select Case intType
        Case msoControlCustom
8630      strRetVal = "msoControlCustom"
8640    Case msoControlButton
8650      strRetVal = "msoControlButton"
8660    Case msoControlEdit
8670      strRetVal = "msoControlEdit"
8680    Case msoControlDropdown
8690      strRetVal = "msoControlDropdown"
8700    Case msoControlComboBox
8710      strRetVal = "msoControlComboBox"
8720    Case msoControlButtonDropdown
8730      strRetVal = "msoControlButtonDropdown"
8740    Case msoControlSplitDropdown
8750      strRetVal = "msoControlSplitDropdown"
8760    Case msoControlOCXDropdown
8770      strRetVal = "msoControlOCXDropdown"
8780    Case msoControlGenericDropdown
8790      strRetVal = "msoControlGenericDropdown"
8800    Case msoControlGraphicDropdown
8810      strRetVal = "msoControlGraphicDropdown"
8820    Case msoControlPopup
8830      strRetVal = "msoControlPopup"
8840    Case msoControlGraphicPopup
8850      strRetVal = "msoControlGraphicPopup"
8860    Case msoControlButtonPopup
8870      strRetVal = "msoControlButtonPopup"
8880    Case msoControlSplitButtonPopup
8890      strRetVal = "msoControlSplitButtonPopup"
8900    Case msoControlSplitButtonMRUPopup
8910      strRetVal = "msoControlSplitButtonMRUPopup"
8920    Case msoControlLabel
8930      strRetVal = "msoControlLabel"
8940    Case msoControlExpandingGrid
8950      strRetVal = "msoControlExpandingGrid"
8960    Case msoControlSplitExpandingGrid
8970      strRetVal = "msoControlSplitExpandingGrid"
8980    Case msoControlGrid
8990      strRetVal = "msoControlGrid"
9000    Case msoControlGauge
9010      strRetVal = "msoControlGauge"
9020    Case msoControlGraphicCombo
9030      strRetVal = "msoControlGraphicCombo"
9040    Case msoControlPane
9050      strRetVal = "msoControlPane"
9060    Case msoControlActiveX
9070      strRetVal = "msoControlActiveX"
9080    End Select

EXITP:
9090    CmdBars_CtlType = strRetVal
9100    Exit Function

ERRH:
9110    strRetVal = vbNullString
9120    Select Case ERR.Number
        Case Else
9130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9140    End Select
9150    Resume EXITP

End Function

Public Function CmdBars_List() As Boolean

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBars_List"

        Dim cbar As Office.CommandBar
        Dim cbctl1 As Office.CommandBarControl, cbctl2 As Office.CommandBarControl
        Dim cbctl3 As Office.CommandBarControl, cbctl4 As Office.CommandBarControl
        Dim prp As Object
        Dim cbbtn As Office.CommandBarButton
        Dim blnRetVal As Boolean

9210    blnRetVal = True

        'JUST DOESN'T WORK!@#$%^&*

        '? Application.VBE.CommandBars("Menu Bar").Controls(10).Controls(2).OnAction
        'VBA_DAOHelp()

        'CBAR BTN: Microsoft &DAO Help
        '  BeginGroup        = False
        '  BuiltIn           = False
        '  Caption           = Microsoft &DAO Help
        '  Creator           = 1398031698
        '  DescriptionText   =
        '  Enabled           = True
        '  Height            = 19
        '  HelpContextId     = 0
        '  HelpFile          =
        '  Id                = 1
        '  Index             = 2
        '  IsPriorityDropped = False
        '  Left              = n/a
        '  OLEUsage          = 1
        '  OnAction          = VBA_DAOHelp
        '  Parameter         = DAO Help
        '  Parent            = Help
        '  Priority          = 3
        '  Tag               =
        '  ToolTipText       = Microsoft &DAO Help
        '  Top               = n/a
        '  Type              = 1
        '  Visible           = True
        '  Width             = 194

9220    Set cbar = Application.VBE.CommandBars("Menu Bar")
9230    With cbar
9240      For Each cbctl1 In .Controls
9250        With cbctl1
9260          If .Caption = "&Help" Then
                '&Help Index = 10
                'Debug.Print "'&Help Index = " & .index
9270            For Each cbctl2 In .Controls
9280              With cbctl2
9290                If InStr(.Caption, "DAO") > 0 Then
                      ' ** Since there is no Properties Collection,
                      ' ** they have to be individually called.
9300                  Debug.Print "'CBAR BTN: " & .Caption
9310                  Debug.Print "'  BeginGroup        = " & .BeginGroup
9320                  Debug.Print "'  BuiltIn           = " & .BuiltIn
9330                  Debug.Print "'  Caption           = " & .Caption
9340                  Debug.Print "'  Creator           = " & .Creator
9350                  Debug.Print "'  DescriptionText   = " & .DescriptionText
9360                  Debug.Print "'  Enabled           = " & .Enabled
9370                  Debug.Print "'  Height            = " & .Height
9380                  Debug.Print "'  HelpContextId     = " & .HelpContextId
9390                  Debug.Print "'  HelpFile          = " & .HelpFile
9400                  Debug.Print "'  Id                = " & .ID
9410                  Debug.Print "'  Index             = " & .index
9420                  Debug.Print "'  IsPriorityDropped = " & .IsPriorityDropped
9430                  Debug.Print "'  Left              = " & "n/a"
9440                  Debug.Print "'  OLEUsage          = " & .OLEUsage
9450                  Debug.Print "'  OnAction          = " & .OnAction
9460                  Debug.Print "'  Parameter         = " & .Parameter
9470                  Debug.Print "'  Parent            = " & .Parent.Name
9480                  Debug.Print "'  Priority          = " & .Priority
9490                  Debug.Print "'  Tag               = " & .Tag
9500                  Debug.Print "'  ToolTipText       = " & .TooltipText
9510                  Debug.Print "'  Top               = " & "n/a"
9520                  Debug.Print "'  Type              = " & .Type
9530                  Debug.Print "'  Visible           = " & .Visible
9540                  Debug.Print "'  Width             = " & .Width
9550                  Exit For
9560                End If
9570              End With
9580            Next
9590            Exit For
9600          End If
9610        End With
9620      Next
9630    End With

        ' ** This looks for 'Properties' entries to see if I can find the one present
        ' ** in the Customize menu on the Access side not present in the VBE side.
        ' ** NOT FOUND!
        'For Each cbar In Application.CommandBars  'Application.VBE.CommandBars
        '  With cbar
        '    For Each cbctl1 In .Controls
        '      With cbctl1
        '        If InStr(.Caption, "Properties") > 0 Then
        '          Debug.Print "'CBAR: " & cbar.Name & " CTL1: " & .Caption
        '        End If
        '        If .Type = msoControlPopup Then
        '          If .Controls.Count > 0& Then
        '            For Each cbctl2 In .Controls
        '              With cbctl2
        '                If InStr(.Caption, "Properties") > 0 Then
        '                  Debug.Print "'CBAR: " & cbar.Name & " CTL1: " & cbctl1.Caption & " CTL2: " & .Caption
        '                End If
        '                If .Type = msoControlPopup Then
        '                  If .Controls.Count > 0& Then
        '                    For Each cbctl3 In .Controls
        '                      With cbctl3
        '                        If InStr(.Caption, "Properties") > 0 Then
        '                          Debug.Print "'CBAR: " & cbar.Name & " CTL1: " & cbctl1.Caption & _
        '                            " CTL2: " & cbctl2.Caption & " CTL3: " & .Caption
        '                        End If
        '                        If .Type = msoControlPopup Then
        '                          If .Controls.Count > 0& Then
        '                            For Each cbctl4 In .Controls
        '                              With cbctl4
        '                                If InStr(.Caption, "Properties") > 0 Then
        '                                  Debug.Print "'CBAR: " & cbar.Name & " CTL1: " & cbctl1.Caption & _
        '                                    " CTL2: " & cbctl2.Caption & " CTL3: " & cbctl3.Caption & " CTL4: " & .Caption
        '                                End If
        '                                'Debug.Print "'CBAR: " & cbar.Name & " CTL1: " & cbctl1.Caption & _
        '                                '  " CTL2: " & cbctl2.Caption & " CTL3: " & .Caption & " CNT: " & CStr(.Controls.Count)
        '                              End With
        '                            Next
        '                          End If
        '                        End If
        '                      End With
        '                    Next
        '                  End If
        '                End If
        '              End With
        '            Next
        '          End If
        '        End If
        '      End With
        '    Next
        '    'Debug.Print "'" & .Name
        '  End With
        'Next

        'CBAR: Database CTL1: &Properties
        'CBAR: Table Design CTL1: &Properties
        'CBAR: Query Design CTL1: &Properties
        'CBAR: Form Design CTL1: &Properties
        'CBAR: Form View CTL1: &Properties
        'CBAR: Report Design CTL1: &Properties
        'CBAR: Page Design CTL1: &Properties
        'CBAR: View Design CTL1: &Properties
        'CBAR: Diagram Design CTL1: &Properties
        'CBAR: Page View CTL1: &Properties
        'CBAR: Menu Bar CTL1: &Edit CTL2: List Properties/Met&hods
        'CBAR: Menu Bar CTL1: &View CTL2: T&able CTL3: &Column Properties
        'CBAR: Menu Bar CTL1: &View CTL2: &Properties
        'CBAR: Menu Bar CTL1: &View CTL2: &Properties
        'CBAR: Menu Bar CTL1: &View CTL2: &Join Properties
        'CBAR: Menu Bar CTL1: &Tools CTL2: &Database Utilities CTL3: Conver&t Database CNT: 2
        'CBAR: Menu Bar CTL1: &Tools CTL2: Re&plication CTL3: Su&bscription Properties...
        'CBAR: Menu Bar CTL1: &Tools CTL2: Re&plication CTL3: Pub&lisher Properties...
        'CBAR: Menu Bar CTL1: &Tools CTL2: Sourc&e Code Control CTL3: SourceSafe &Properties...
        'CBAR: Database Table/Query CTL1: &Properties
        'CBAR: Database Form CTL1: &Properties
        'CBAR: Database Report CTL1: &Properties
        'CBAR: Database Macro CTL1: &Properties
        'CBAR: Database Module CTL1: &Properties
        'CBAR: Table DesignTitleBar CTL1: &Properties
        'CBAR: Table Design Upper Pane CTL1: &Properties
        'CBAR: Table Design Lower Pane CTL1: &Properties
        'CBAR: Query CTL1: &Properties...
        'CBAR: Query DesignFieldList CTL1: &Properties...
        'CBAR: Query Design Join CTL1: &Join Properties
        'CBAR: Query DesignGrid CTL1: &Properties...
        'CBAR: Query SQLTitleBar CTL1: &Properties...
        'CBAR: Form DesignTitleBar CTL1: &Properties
        'CBAR: Form Design Form CTL1: &Properties
        'CBAR: Form Design Section CTL1: &Properties
        'CBAR: Form Design Control CTL1: &Properties
        'CBAR: Form Design Control OLE CTL1: &Properties
        'CBAR: Report DesignTitleBar CTL1: &Properties
        'CBAR: Report DesignReport CTL1: &Properties
        'CBAR: Report Design Section CTL1: &Properties
        'CBAR: Report Design Control CTL1: &Properties
        'CBAR: Form View Popup CTL1: &Properties
        'CBAR: Form View Subform CTL1: &Properties
        'CBAR: Form View Control CTL1: &Properties
        'CBAR: Form View Subform Control CTL1: &Properties
        'CBAR: Form View Record CTL1: &Properties
        'CBAR: ModuleUncompiled CTL1: List Properties/Met&hods
        'CBAR: Database Page Popup CTL1: &Properties
        'CBAR: Tab Control CTL1: &Properties
        'CBAR: Tab Control on Report Design CTL1: &Properties
        'CBAR: Page Popup CTL1: &Properties
        'CBAR: View Table Design Popup CTL1: &Properties
        'CBAR: View Design Diagram Pane Popup CTL1: &Properties
        'CBAR: View Design Grid Pane Popup CTL1: &Properties
        'CBAR: View Design SQL Pane Popup CTL1: &Properties
        'CBAR: Diagram Design Popup CTL1: &Column Properties
        'CBAR: Diagram Design Popup CTL1: &Properties
        'CBAR: Diagram Popup CTL1: &Properties
        'CBAR: View Design Join Line Popup CTL1: &Properties
        'CBAR: View Table View Mode Submenu CTL1: &Column Properties
        'CBAR: Join Line Popup CTL1: &Properties
        'CBAR: Database View CTL1: &Properties
        'CBAR: Database Stored Procedure CTL1: &Properties
        'CBAR: Database Diagram CTL1: &Properties

        'Set cbar = Application.CommandBars("Menu Bar")
        ''Set cbar = Application.VBE.CommandBars("Menu Bar")
        'With cbar
        '  For Each cbctl1 In .Controls
        '    With cbctl1
        '      'If .Caption = "&View" Then
        '      '  lngIdx = .index
        '      '  'Debug.Print "'IDX: " & .index
        '      '  'Debug.Print "'" & Application.VBE.CommandBars("Menu Bar").Controls(.index).Caption
        '      '  Exit For
        '      'End If
        '      Debug.Print "'" & .Caption
        '    End With
        '  Next
        'Set cbctl2 = .Controls(lngIdx)
        'With cbctl2
        '  For Each cbctl1 In .Controls
        '    With cbctl1
        '      If .Caption = "&Toolbars" Then
        '        lngIdx = .index
        '        Exit For
        '      End If
        '      'Debug.Print "'" & .Caption
        '    End With
        '  Next
        '  Set cbctl1 = Nothing
        '  Set cbctl1 = .Controls(lngIdx)
        '  With cbctl1
        '    For Each cbctl3 In .Controls
        '      With cbctl3
        '        Debug.Print "'" & .Caption
        '      End With
        '    Next
        '  End With
        'End With

        ''With cbctl2
        '  Set cbbtn = cbar.Controls.Add(msoControlButton, , "DAO Help", 2, False)
        '  With cbbtn
        '    .style = msoButtonCaption
        '    .Caption = "Microsoft &DAO Help"
        '    .OnAction = "VBA_DAOHelp"
        '    '.ShortcutText = "x"
        '    '.HyperlinkType = msoCommandBarButtonHyperlinkOpen
        '    '.TooltipText = "www.microsoft.com"
        '    '.DescriptionText = "x"
        '    '.HelpFile = "C:\Documents and Settings\VictorC\Desktop\ADO-DAO\DAO360.CHM"
        '    '.HelpContextId = 1
        '  End With
        '  'expression.Add([Type], [Id], [Parameter], [Before], [Temporary])
        '  'For Each cbctl1 In .Controls
        '  '  With cbctl1
        '  '    Debug.Print "'" & .Caption
        '  '  End With
        '  'Next
        ''End With
        'End With

        'Types:
        'msoControlButton
        'msoControlEdit
        'msoControlDropdown
        'msoControlComboBox
        'msoControlPopup

        'Toolbars:
        'Debug
        'Edit
        'Standard
        'UserForm
        '&Customize...

        'View:
        '&Code
        'O&bject
        '&Definition
        'Last Positio&n
        '&Object Browser
        '&Immediate Window
        'Local&s Window
        'Watc&h Window
        'Call Stac&k...
        '&Project Explorer
        'Properties &Window
        'Toolbo&x
        '&Toolbars
        'Microsoft Access

        'Help:
        'Microsoft Visual Basic &Help
        'MSDN on the &Web
        '&About Microsoft Visual Basic...

        'Menu Bar:
        '&File
        '&Edit
        '&View
        '&Insert
        '&Debug
        '&Run
        '&Tools
        '&Add-Ins
        '&Window
        '&Help

        'Application.VBE.CommandBars:
        'Menu Bar
        'Standard
        'Edit
        'Debug
        'UserForm
        'Document
        'Project Window Insert
        'Toggle
        'Code Window
        'Code Window (Break)
        'Watch Window
        'Immediate Window
        'Locals Window
        'Project Window
        'Project Window (Break)
        'Object Browser
        'MSForms
        'MSForms Control
        'MSForms Control Group
        'MSForms Palette
        'MSForms Toolbox
        'MSForms MPC
        'MSForms DragDrop
        'Toolbox
        'Toolbox Group
        'Property Browser
        'Property Browser
        'Docked Window

        'Application.CommandBars:
        'Font/Fore Color
        'Fill/Back Color
        'Line/Border Style
        'Line/Border Width
        'Line/Border Color
        'Special Effect
        'Fill/Back Color
        'Font/Fore Color
        'Line/Border Color
        'Datasheet Special Effect
        'Gridlines
        'Appearance
        'Database
        'Relationship
        'Table Design
        'Table Datasheet
        'Query Design
        'Query Datasheet
        'Form Design
        'Form View
        'Filter/Sort
        'Report Design
        'Print Preview
        'Toolbox
        'Formatting (Form/Report)
        'Formatting (Datasheet)
        'Macro Design
        'Utility 1
        'Utility 2
        'Web
        'Source Code Control
        'Page Design
        'Formatting (Page)
        'View Design
        'Diagram Design
        'Stored Procedure Design
        'Trigger Design
        'Alignment and Sizing
        'Page View
        'Menu Bar
        'Database Table/Query
        'Database Form
        'Database Report
        'Database Macro
        'Database Module
        'Clipboard
        'TAReports
        'Database TitleBar
        'Table DesignTitleBar
        'Table Design Upper Pane
        'Table Design Lower Pane
        'Table Design Properties
        'Index TitleBar
        'Index Upper Pane
        'Index Lower Pane
        'Index Properties
        'Actions
        'Relationship TableFieldList
        'Relationship QueryFieldList
        'Relationship Join
        'Table Design Datasheet
        'Table Design Datasheet Column
        'Table Design Datasheet Row
        'Table Design Datasheet Cell
        'Query
        'Query DesignFieldList
        'Query Design Join
        'Query DesignGrid
        'Query SQLTitleBar
        'Query SQLText
        'Query Design Properties
        'Query Design Datasheet
        'Query Design Datasheet Column
        'Query Design Datasheet Row
        'Query Design Datasheet Cell
        'Filter General Context Menu
        'Filter Field
        'Filter FilterByForm
        'Form DesignTitleBar
        'Form Design Form
        'Form Design Section
        'Form Design Control
        'Form Design Control OLE
        'Report DesignTitleBar
        'Report DesignReport
        'Report Design Section
        'Report Design Control
        'Form/Report Properties
        'Form View Popup
        'Form View Subform
        'Form View Control
        'Form View Subform Control
        'Form View Record
        'Form Datasheet
        'Form Datasheet Column
        'Form Datasheet Subcolumn
        'Form Datasheet Row
        'Form Datasheet Cell
        'Print Preview Popup
        'Macro TitleBar
        'Macro UpperPane
        'Macro Condition
        'Macro Argument
        'ModuleUncompiled
        'ModuleCompiled
        'Module Immediate
        'Module Watch
        'Database Page Popup
        'Database Background
        'OLE Shared
        'Global
        'Object Browser
        'Module LocalsPane
        'Tab Control
        'Tab Control on Report Design
        'Page Popup
        'View Table Design Popup
        'View Design Background Popup
        'View Design Diagram Pane Popup
        'View Design Field Popup
        'View Design Grid Pane Popup
        'View Design SQL Pane Popup
        'Diagram Design Popup
        'Diagram Popup
        'View Design Multiple Select Popup
        'View Design Join Line Popup
        'View Table View Mode Submenu
        'View Show Panes Submenu
        'Web Page Layout
        'Join Line Popup
        'Stored Procedure Design Datasheet
        'View Design Datasheet
        'Diagram Design Label Popup
        'Stored Procedure Design Popup
        'Trigger Design Popup
        'Database Shortcut Popup
        'Database Custom Group Popup
        'Database View
        'Database Stored Procedure
        'Database Diagram

9640    Beep

EXITP:
9650    Set prp = Nothing
9660    Set cbbtn = Nothing
9670    Set cbctl1 = Nothing
9680    Set cbctl2 = Nothing
9690    Set cbctl3 = Nothing
9700    Set cbctl4 = Nothing
9710    Set cbar = Nothing
9720    CmdBars_List = blnRetVal
9730    Exit Function

ERRH:
9740    blnRetVal = False
9750    Select Case ERR.Number
        Case Else
9760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9770    End Select
9780    Resume EXITP

End Function

Public Function CmdBars_Clipboard(Optional varCalled As Variant) As Boolean
' ** Turn off that annoying multi-clipboard popup!

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBars_Clipboard"

        Dim cbar As Office.CommandBar
        Dim blnRetVal As Boolean

9810    blnRetVal = True

9820    Set cbar = Application.CommandBars("Clipboard")
9830    With cbar
9840      .Enabled = False
9850      .Visible = False
          '.Protection
          ' ** MsoBarProtection enumeration:
          ' **    0  msoBarNoProtection
          ' **    1  msoBarNoCustomize
          ' **    2  msoBarNoResize
          ' **    4  msoBarNoMove
          ' **    8  msoBarNoChangeVisible
          ' **   16  msoBarNoChangeDock
          ' **   32  msoBarNoVerticalDock
          ' **   64  msoBarNoHorizontalDock
9860    End With

9870    If IsMissing(varCalled) = True Then
9880      Beep
9890    End If

EXITP:
9900    Set cbar = Nothing
9910    CmdBars_Clipboard = blnRetVal
9920    Exit Function

ERRH:
9930    blnRetVal = False
9940    Select Case ERR.Number
        Case Else
9950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9960    End Select
9970    Resume EXITP

End Function

Public Function DoBeeps(Optional varTimBeeps As Variant, Optional varTimDurMS As Variant) As Boolean
' ** 600 ms or 650 ms works nicely.
' ** ms = milliseconds, same unit as TimerInterval.

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "DoBeeps"

        Dim MMT As MMTIME
        Dim lngTimStart As Long, lngTimFinish As Long, lngTimLast As Long, lngTimCnt As Long
        Dim lngTimBeeps As Long, lngTimDurMS As Long, dblTimDurSEC As Double
        Dim blnRetVal As Boolean

10010   blnRetVal = True

10020   MMT.wType = TIME_MS

10030   If IsMissing(varTimBeeps) = True Then lngTimBeeps = 4& Else lngTimBeeps = CLng(varTimBeeps)
10040   If IsMissing(varTimDurMS) = True Then lngTimDurMS = 650& Else lngTimDurMS = CLng(varTimDurMS)
10050   dblTimDurSEC = (lngTimDurMS / 1000)

10060   lngTimStart = GetTimeUnits  ' ** Function: Below.
10070   lngTimLast = lngTimStart
10080   lngTimCnt = 0&

10090   Do While lngTimCnt < lngTimBeeps  ' ** Zero-based, so one less than beeps desired.
10100     lngTimFinish = GetTimeUnits  ' ** Function: Below.
10110     If ((lngTimFinish - lngTimLast) / 1000) >= dblTimDurSEC Then  ' ** Wait till X Secs from the last beep.
10120       lngTimCnt = lngTimCnt + 1
10130       lngTimLast = lngTimFinish
10140       Beep
10150       DoEvents
10160     End If
10170   Loop

EXITP:
10180   DoBeeps = blnRetVal
10190   Exit Function

ERRH:
10200   blnRetVal = False
10210   Select Case ERR.Number
        Case Else
10220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10230   End Select
10240   Resume EXITP

End Function

Public Function GetTimeUnits() As Long

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "GetTimeUnits"

        Dim MMT As MMTIME
        Dim lngResponse As Long
        Dim lngRetVal As Long

10310   lngResponse = timeGetSystemTime(MMT, Len(MMT))  ' ** API Function: Above.
10320   lngRetVal = MMT.Units

EXITP:
10330   GetTimeUnits = lngRetVal
10340   Exit Function

ERRH:
10350   Select Case ERR.Number
        Case Else
10360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10370   End Select
10380   Resume EXITP

End Function

Public Sub AccessCloseButtonEnabled(blnEnabled As Boolean)
' ** Comments: Control the Access close button.
' **           Disabling it forces the user to exit within the application.
' ** Params  : blnEnabled       TRUE enables the close button, FALSE disabled it
' ** Owner   : Copyright (c) 2008-2011 from FMS, Inc.
' ** Source  : Total Visual SourceBook
' ** Usage   : Permission granted to subscribers of the FMS Newsletter
'DOESN'T WORK!!

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "AccessCloseButtonEnabled"

        Dim lngWindow As Long
        Dim lngMenu As Long
        Dim lngFlags As Long

        Const clngMF_ByCommand As Long = &H0&
        Const clngMF_Grayed    As Long = &H1&
        Const clngSC_Close     As Long = &HF060&

10410 On Error Resume Next

10420   lngWindow = Application.hWndAccessApp
10430   lngMenu = GetSystemMenu(lngWindow, 0)
10440   If blnEnabled Then
10450     lngFlags = clngMF_ByCommand And Not clngMF_Grayed
10460   Else
10470     lngFlags = clngMF_ByCommand Or clngMF_Grayed
10480   End If

10490   EnableMenuItem lngMenu, clngSC_Close, lngFlags  ' ** API Function: Above.

EXITP:
10500   Exit Sub

ERRH:
10510   Select Case ERR.Number
        Case Else
10520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10530   End Select
10540   Resume EXITP

End Sub

Public Function HelpBar_Close() As Boolean

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "HelpBar_Close"

        Dim blnRetVal As Boolean

10610   blnRetVal = True

        ' ** Remove the "Type a question for help" on the default menu bar in Access 2002 or 2003
10620   Application.CommandBars.DisableAskAQuestionDropdown = True

EXITP:
10630   HelpBar_Close = blnRetVal
10640   Exit Function

ERRH:
10650   Select Case ERR.Number
        Case Else
10660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10670   End Select
10680   Resume EXITP

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
10700 On Error GoTo ERRH

        Const THIS_PROC As String = "SystemColor_Get"

EXITP:
10710   SystemColor_Get = GetSysColor(SystemElement)
10720   Exit Function

ERRH:
10730   Select Case ERR.Number
        Case Else
10740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10750   End Select
10760   Resume EXITP

End Function
