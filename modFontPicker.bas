Attribute VB_Name = "modFontPicker"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modFontPicker"

'VGC 03/22/2017: CHANGES!

' ** Original Code by Terry Kreft
' ** Modified by Stephen Lebans
' ** Contact Stephen@lebans.com

'************  Code Start  ***********
Private Const LF_FACESIZE As Long = 32

'Private Const FW_BOLD As Long = 700

'Private Const CF_APPLY                As Long = &H200&
Private Const CF_ANSIONLY             As Long = &H400&
'Private Const CF_TTONLY               As Long = &H40000
Private Const CF_EFFECTS              As Long = &H100&
'Private Const CF_ENABLETEMPLATE       As Long = &H10&
'Private Const CF_ENABLETEMPLATEHANDLE As Long = &H20&
'Private Const CF_FIXEDPITCHONLY       As Long = &H4000&
'Private Const CF_FORCEFONTEXIST       As Long = &H10000
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40&
'Private Const CF_LIMITSIZE            As Long = &H2000&
'Private Const CF_NOFACESEL            As Long = &H80000
'Private Const CF_NOSCRIPTSEL          As Long = &H800000
'Private Const CF_NOSTYLESEL           As Long = &H100000
'Private Const CF_NOSIZESEL            As Long = &H200000
'Private Const CF_NOSIMULATIONS        As Long = &H1000&
Private Const CF_NOVECTORFONTS        As Long = &H800&
'Private Const CF_NOVERTFONTS          As Long = &H1000000
'Private Const CF_OEMTEXT              As Long = &H7
Private Const CF_PRINTERFONTS         As Long = &H2
'Private Const CF_SCALABLEONLY         As Long = &H20000
Private Const CF_SCREENFONTS          As Long = &H1
'Private Const CF_SCRIPTSONLY          As Long = CF_ANSIONLY
'Private Const CF_SELECTSCRIPT         As Long = &H400000
'Private Const CF_SHOWHELP             As Long = &H4&
'Private Const CF_USESTYLE             As Long = &H80&
'Private Const CF_WYSIWYG              As Long = &H8000
'Private Const CF_BOTH                 As Long = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'Private Const CF_NOOEMFONTS           As Long = CF_NOVECTORFONTS

Public Type FORMFONTINFO
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  Color As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
  lStructSize As Long
  hwnd As Long
  hDC As Long
  lpLogFont As Long
  iPointSize As Long
  flags As Long
  rgbColors As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  hInstance As Long
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type

Private Declare Function EnumFonts Lib "gdi32.dll" Alias "EnumFontsA" (ByVal hDC As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
'Parameter Information
' hDC
'Identifies the device context.
'·lpFaceName
'Points to a null-terminated character string that specifies the typeface name of the desired fonts. If lpFaceName is NULL, EnumFonts randomly selects and enumerates one font of each available typeface.
'·lpFontFunc
'Points to the application-defined callback function. For more information about the callback function, see the EnumFontsProc function.
'·lParam
'Points to any application-defined data. The data is passed to the callback function along with the font information.
'Return Values
'If the function succeeds, the return value is the last value returned by the callback function. Its meaning is defined by the application.

Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long

' ** Array: arr_varFont().
Private lngFonts As Long, arr_varFont() As Variant
Private Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const F_NAM  As Integer = 0
Private Const F_SHOW As Integer = 1
Private Const F_MICR As Integer = 2
' **

Public Function FindMICRFont() As String

100   On Error GoTo ERRH

        Const THIS_PROC As String = "FindMICRFont"

        Dim lngHdc As Long
        Dim intPos01 As Integer
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long
        Dim strRetVal As String

110     strRetVal = vbNullString

120     lngFonts = 0&
130     ReDim arr_varFont(F_ELEMS, 0)

140     lngHdc = GetDC(hWndAccessApp)  ' ** API Function: modWindowFunctions.
150     EnumFonts lngHdc, vbNullString, AddressOf EnumFontProc, 0  ' ** Function: Below.

        ' ** Binary Sort arr_varFont() array.
160     For lngX = UBound(arr_varFont, 2) To 1 Step -1
170       For lngY = 0 To (lngX - 1)
180         If arr_varFont(F_NAM, lngY) > arr_varFont(F_NAM, (lngY + 1)) Then
190           For lngZ = 0& To F_ELEMS
200             varTmp00 = arr_varFont(lngZ, lngY)
210             arr_varFont(lngZ, lngY) = arr_varFont(lngZ, (lngY + 1&))
220             arr_varFont(lngZ, (lngY + 1&)) = varTmp00
230           Next
240         End If
250       Next
260     Next

270     For lngX = 0& To (lngFonts - 1&)

          ' ** Check for a MICR font.
280       intPos01 = InStr(arr_varFont(F_NAM, lngX), "micr")
290       If intPos01 > 0 Then
300         If Mid(arr_varFont(F_NAM, lngX), intPos01, 9) <> "Microsoft" Then
310           arr_varFont(F_MICR, lngX) = CBool(True)
320         End If
330       End If

          ' ** @Batang, don't know what these are.
340       If Left(arr_varFont(F_NAM, lngX), 1) = "@" Then
350         arr_varFont(F_SHOW, lngX) = CBool(False)
360       End If

          ' ** WP ArabicScript Sihafa, don't know what these are.
370       If Left(arr_varFont(F_NAM, lngX), 3) = "WP " Then
380         arr_varFont(F_SHOW, lngX) = CBool(False)
390       End If

          ' ** WST_Czec, don't know what these are.
400       If Left(arr_varFont(F_NAM, lngX), 4) = "WST_" Then
410         arr_varFont(F_SHOW, lngX) = CBool(False)
420       End If

          ' ** ZWAdobeF, don't need to see these.
430       If Left(arr_varFont(F_NAM, lngX), 7) = "ZWAdobe" Then
440         arr_varFont(F_SHOW, lngX) = CBool(False)
450       End If

460     Next

470     For lngX = 0& To (lngFonts - 1&)
480       If arr_varFont(F_MICR, lngX) = True Then
490         If strRetVal <> vbNullString Then strRetVal = strRetVal & "; "
500         strRetVal = strRetVal & arr_varFont(F_NAM, lngX)
510       End If
520     Next

        'Debug.Print "'FONTS: " & CStr(lngFonts)
        'FONTS: 254

        'For lngX = 0& To (lngFonts - 1&)
        '  If arr_varFont(F_SHOW, lngX) = True Then
        '    Select Case arr_varFont(F_MICR, lngX)
        '    Case True
        '      Debug.Print "'***  " & arr_varFont(F_NAM, lngX) & "  ***"
        '    Case False
        '      Debug.Print "'" & arr_varFont(F_NAM, lngX)
        '    End Select
        '  End If
        'Next

        'Bodoni MT Condensed
        'Bodoni MT Poster Compressed
        'Book Antiqua
        'Bookman Old Style
        'Bookshelf Symbol 7
        'Boulevard BQ
        'Bradley Hand ITC
        'Briem Akademi Std Cond
        'Britannic Bold
        'Broadway
        'Brush Script MT
        'Calibri
        'Californian FB
        'Calisto MT
        'Cambria
        'Cambria Math
        'Candara
        'Castellar
        'Centaur
        'Century
        'Century Gothic
        'Century Schoolbook
        'Chiller
        'Clearface Gothic MT DemiBold
        'Colonna MT
        'Comic Sans MS
        'Consolas
        'Constantia
        'Cooper Black
        'Copperplate Gothic Bold
        'Copperplate Gothic Light
        'Corbel
        'Courier
        'Courier New
        'Courier New Baltic
        'Courier New CE
        'Courier New CYR
        'Courier New Greek
        'Courier New TUR
        'Curlz MT
        'Df Daves Raves TwoTT ITC
        'DragonWick
        'Dragonwick FG
        'Edwardian Script ITC
        'Elephant
        'Engravers MT
        'Eras Bold ITC
        'Eras Demi ITC
        'Eras Light ITC
        'Eras Medium ITC
        'Estrangelo Edessa
        'Euclid
        'Euclid Extra
        'Euclid Fraktur
        'Euclid Math One
        'Euclid Math Two
        'Euclid Symbol
        'Felix Titling
        'Fixedsys
        'Footlight MT Light
        'Forte
        'Franklin Gothic Book
        'Franklin Gothic Demi
        'Franklin Gothic Demi Cond
        'Franklin Gothic Heavy
        'Franklin Gothic Medium
        'Franklin Gothic Medium Cond
        'Freehand521 BT
        'Freestyle Script
        'French Script MT
        'Futura Md BT
        'Garamond
        'Gautami
        'Georgia
        'Gigi
        'Gill Sans MT
        'Gill Sans MT Condensed
        'Gill Sans MT Ext Condensed Bold
        'Gill Sans Ultra Bold
        'Gill Sans Ultra Bold Condensed
        'Gloucester MT Extra Condensed
        'Goudy Old Style
        'Goudy Stout
        'Haettenschweiler
        'Harlow Solid Italic
        'Harrington
        'Helvetica LT Std
        'High Tower Text
        'Impact
        'Imprint MT Shadow
        'Informal Roman
        'Jokerman
        'Juice ITC
        'Kartika
        'Kristen ITC
        'Kunstler Script
        'Latha
        'Lucida Bright
        'Lucida Calligraphy
        'Lucida Console
        'Lucida Fax
        'Lucida Handwriting
        'Lucida Sans
        'Lucida Sans Typewriter
        'Lucida Sans Unicode
        'Magik
        'Magneto
        'Maiandra GD
        'Malabar LT Pro Heavy
        'Mangal
        'Marlett
        'Matura MT Script Capitals
        '***  MICR  ***
        'Microsoft Sans Serif
        'Mistral
        'Modern
        'Modern No. 20
        'ModulaSansRegular
        'Monotype Corsiva
        'MS Mincho
        'MS Outlook
        'MS Reference Sans Serif
        'MS Reference Serif
        'MS Reference Specialty
        'MS Sans Serif
        'MS Serif
        'MT Extra
        'MV Boli
        'Niagara Engraved
        'Niagara Solid
        'OCR A Extended
        'OCR-A BT
        'OCR-B 10 BT
        'Old English Text MT
        'Onyx
        'Palace Script MT
        'Palatino Linotype
        'Papyrus
        'Parchment
        'Perpetua
        'Perpetua Titling MT
        'Playbill
        'PMingLiU
        'Poor Richard
        'Pristina
        'Raavi
        'Rage Italic
        'Ravie
        'Rockwell
        'Rockwell Condensed
        'Rockwell Extra Bold
        'Roman
        'Script
        'Script MT Bold
        'Segoe UI
        'Showcard Gothic
        'Shruti
        'SimSun
        'Sinaloa LET
        'Small Fonts
        'Snap ITC
        'Stencil
        'Stop
        'Swis721 Blk BT
        'Swis721 BT
        'Swis721 Hv BT
        'Swis721 Lt BT
        'Sylfaen
        'Symbol
        'System
        'Tahoma
        'Tempus Sans ITC
        'Terminal
        'Times New Roman
        'Times New Roman Baltic
        'Times New Roman CE
        'Times New Roman CYR
        'Times New Roman Greek
        'Times New Roman TUR
        'Trebuchet MS
        'Tunga
        'Tw Cen MT
        'Tw Cen MT Condensed
        'Tw Cen MT Condensed Extra Bold
        'Twentieth Century Medium
        'Univers LT 45 Light
        'Univers LT 55
        'UpsideDownJJ
        'Verdana
        'Viner Hand ITC
        'Vivaldi
        'Vladimir Script
        'Vrinda
        'Webdings
        'Wide Latin
        'Wingdings
        'Wingdings 2
        'Wingdings 3

EXITP:
530     FindMICRFont = strRetVal
540     Exit Function

ERRH:
550     strRetVal = RET_ERR
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Function

Private Function EnumFontProc(ByVal lplf As Long, ByVal lptm As Long, ByVal dwType As Long, ByVal lpData As Long) As Long

600   On Error GoTo ERRH

        Const THIS_PROC As String = "EnumFontProc"

        Dim LF As LOGFONT, strFontName As String, lngZeroPos As Long
        Dim lngE As Long
        Dim lngRetVal As Long

610     lngRetVal = 0&

620     CopyMemory LF, ByVal lplf, LenB(LF)  ' ** API Function: modWindowFunctions.

630     strFontName = StrConv(LF.lfFaceName, vbUnicode)
640     lngZeroPos = InStr(1, strFontName, Chr$(0))
650     If lngZeroPos > 0 Then strFontName = Left(strFontName, lngZeroPos - 1)

        'Form1.Print strFontName
        'Debug.Print "'" & strFontName
660     lngFonts = lngFonts + 1&
670     lngE = lngFonts - 1&
680     ReDim Preserve arr_varFont(F_ELEMS, lngE)
690     arr_varFont(F_NAM, lngE) = strFontName
700     arr_varFont(F_SHOW, lngE) = CBool(True)
710     arr_varFont(F_MICR, lngE) = CBool(False)

720     lngRetVal = 1&

EXITP:
730     EnumFontProc = lngRetVal
740     Exit Function

ERRH:
750     lngRetVal = 0&
760     Select Case ERR.Number
        Case Else
770       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
780     End Select
790     Resume EXITP

End Function

Public Function DialogFont(ByRef typFFI As FORMFONTINFO, Optional hwnd As Long = 0) As Boolean

800   On Error GoTo ERRH

        Const THIS_PROC As String = "DialogFont"

        Dim LF As LOGFONT, fs As FONTSTRUC
        Dim lngLogFontAddress As Long, lngMemHandle As Long
        Dim blnRetVal As Boolean

810     blnRetVal = True

820     LF.lfWeight = typFFI.Weight
830     LF.lfItalic = typFFI.Italic * -1
840     LF.lfUnderline = typFFI.UnderLine * -1
850     LF.lfHeight = -MultDiv(CLng(typFFI.Height), GetDeviceCaps(GetDC(hWndAccessApp), GSR_LOGPIXELSY), 72)
860     StringToByte typFFI.Name, LF.lfFaceName()  ' ** Procedure: Below.
870     fs.rgbColors = typFFI.Color
880     fs.lStructSize = Len(fs)

890     fs.hwnd = hwnd  ' April 2 Font Dialog is BEHIND CALENDAR!!! Application.hWndAccessApp

900     lngMemHandle = GlobalAlloc(GMEM_HND, Len(LF))  ' ** API Function: modWindowFunctions.
910     If lngMemHandle = 0 Then
920       blnRetVal = False
930     Else

940       lngLogFontAddress = GlobalLock(lngMemHandle)
950       If lngLogFontAddress = 0 Then
960         blnRetVal = False
970       Else

980         CopyMemory ByVal lngLogFontAddress, LF, Len(LF)  ' ** API Function: modWindowFunctions.
990         fs.lpLogFont = lngLogFontAddress
1000        fs.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
1010        If ChooseFont(fs) = 1 Then
1020          CopyMemory LF, ByVal lngLogFontAddress, Len(LF)  ' ** API Function: modWindowFunctions.
1030          typFFI.Weight = LF.lfWeight
1040          typFFI.Italic = CBool(LF.lfItalic)
1050          typFFI.UnderLine = CBool(LF.lfUnderline)
1060          typFFI.Name = ByteToString(LF.lfFaceName())
1070          typFFI.Height = CLng(fs.iPointSize / 10)
1080          typFFI.Color = fs.rgbColors
1090          blnRetVal = True
1100        Else
1110          blnRetVal = False
1120        End If

1130      End If
1140    End If

EXITP:
1150    DialogFont = blnRetVal
1160    Exit Function

ERRH:
1170    blnRetVal = False
1180    Select Case ERR.Number
        Case Else
1190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1200    End Select
1210    Resume EXITP

End Function

Public Function test_DialogFont() As String

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "test_DialogFont"

        Dim typFFI As FORMFONTINFO
        Dim strRetVal As String

1310    strRetVal = vbNullString

1320    With typFFI
1330      .Color = 255
1340      .Height = 12
1350      .Weight = 700
1360      .Italic = False
1370      .UnderLine = True
1380      .Name = "Tahoma"
1390    End With

1400    DialogFont typFFI  ' ** Function: Above.

1410    With typFFI
          ''debug.print "Font Name: "; .name
          ''debug.print "Font Size: "; .Height
          ''debug.print "Font Weight: "; .Weight
          ''debug.print "Font Italics: "; .Italic
          ''debug.print "Font Underline: "; .UnderLine
          ''debug.print "Font COlor: "; .Color
1420      strRetVal = .Name
1430    End With

EXITP:
1440    test_DialogFont = strRetVal
1450    Exit Function

ERRH:
1460    strRetVal = vbNullString
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Function

'************  Code End  ***********

Private Function ByteToString(aBytes() As Byte) As String

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "ByteToString"

        Dim lngDWBytePoint As Long, lngDWByteVal As Long, strOut As String

1610    lngDWBytePoint = LBound(aBytes)
1620    Do While lngDWBytePoint <= UBound(aBytes)
1630      lngDWByteVal = aBytes(lngDWBytePoint)
1640      If lngDWByteVal = 0 Then
1650        Exit Do
1660      Else
1670        strOut = strOut & Chr$(lngDWByteVal)
1680      End If
1690      lngDWBytePoint = lngDWBytePoint + 1
1700    Loop

EXITP:
1710    ByteToString = strOut
1720    Exit Function

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Function

Private Function MultDiv(In1 As Long, In2 As Long, In3 As Long) As Long
' ** Not to be confused with the API function MulDiv() in clsMonthCal.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "MultDiv"

        Dim lngTmp01 As Long

1810    If In3 <> 0 Then
1820      lngTmp01 = In1 * In2
1830      lngTmp01 = lngTmp01 / In3
1840    Else
1850      lngTmp01 = -1&
1860    End If

EXITP:
1870    MultDiv = lngTmp01
1880    Exit Function

ERRH:
1890    lngTmp01 = -1&
1900    Select Case ERR.Number
        Case Else
1910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1920    End Select
1930    Resume EXITP

End Function

Private Sub StringToByte(strInString As String, bytArray() As Byte)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "StringToByte"

        Dim intLbound As Integer
        Dim intUbound As Integer
        Dim intLen As Integer
        Dim intX As Integer

2010    intLbound = LBound(bytArray)
2020    intUbound = UBound(bytArray)
2030    intLen = Len(strInString)
2040    If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
2050    For intX = 1 To intLen
2060      bytArray(intX - 1 + intLbound) = Asc(Mid(strInString, intX, 1))
2070    Next

EXITP:
2080    Exit Sub

ERRH:
2090    Select Case ERR.Number
        Case Else
2100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2110    End Select
2120    Resume EXITP

End Sub

Public Function GetFontArray() As Variant

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFontArray"

        Dim lngHdc As Long
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long
        Dim arr_varRetVal As Variant

2210    lngFonts = 0&
2220    ReDim arr_varFont(F_ELEMS, 0)

2230    lngHdc = GetDC(hWndAccessApp)  ' ** API Function: modWindowFunctions.
2240    EnumFonts lngHdc, vbNullString, AddressOf EnumFontProc, 0  ' ** Function: Below.

        ' ** Binary Sort arr_varFont() array.
2250    For lngX = UBound(arr_varFont, 2) To 1 Step -1
2260      For lngY = 0 To (lngX - 1)
2270        If arr_varFont(F_NAM, lngY) > arr_varFont(F_NAM, (lngY + 1)) Then
2280          For lngZ = 0& To F_ELEMS
2290            varTmp00 = arr_varFont(lngZ, lngY)
2300            arr_varFont(lngZ, lngY) = arr_varFont(lngZ, (lngY + 1&))
2310            arr_varFont(lngZ, (lngY + 1&)) = varTmp00
2320          Next
2330        End If
2340      Next
2350    Next

2360    arr_varRetVal = arr_varFont

EXITP:
2370    GetFontArray = arr_varRetVal
2380    Exit Function

ERRH:
2390    arr_varRetVal = RET_ERR
2400    Select Case ERR.Number
        Case Else
2410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2420    End Select
2430    Resume EXITP

End Function

Public Function test_GetFontArray() As Boolean

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "test_GetFontArray"

        Dim lngTmp01 As Long, arr_varTmp02 As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

2510  On Error GoTo 0

2520    blnRetVal = True

2530    arr_varTmp02 = GetFontArray  ' ** Function: Above.
2540    lngTmp01 = UBound(arr_varTmp02, 2) + 1&

2550    Debug.Print "'FONTS: " & CStr(lngTmp01)
2560    DoEvents

2570    For lngX = 0& To (lngTmp01 - 1&)
2580      Debug.Print "'" & Left(CStr(lngX + 1) & "." & String(5, " "), 5) & arr_varTmp02(F_NAM, lngX)
2590      DoEvents
2600      If (lngX + 1&) Mod 100 = 0 Then
2610        Stop
2620      End If
2630    Next  ' ** lngX.

2640    Beep

        'FONTS: 346
'1.   @Arial Unicode MS
'2.   @Batang
'3.   @BatangChe
'4.   @Code2000
'5.   @DFKai-SB
'6.   @Dotum
'7.   @DotumChe
'8.   @FangSong
'9.   @Gulim
'10.  @GulimChe
'11.  @Gungsuh
'12.  @GungsuhChe
'13.  @KaiTi
'14.  @Malgun Gothic
'15.  @Meiryo
'16.  @Meiryo UI
'17.  @Microsoft JhengHei
'18.  @Microsoft YaHei
'19.  @MingLiU
'20.  @MingLiU_HKSCS
'21.  @MingLiU_HKSCS-ExtB
'22.  @MingLiU-ExtB
'23.  @MS Gothic
'24.  @MS Mincho
'25.  @MS PGothic
'26.  @MS PMincho
'27.  @MS UI Gothic
'28.  @NSimSun
'29.  @PMingLiU
'30.  @PMingLiU-ExtB
'31.  @SimHei
'32.  @SimSun
'33.  @SimSun-ExtB
'34.  Aeris B Std
'35.  AGaramond RegularSC
'36.  Agency FB
'37.  Aharoni
'38.  Algerian
'39.  Andalus
'40.  Angsana New
'41.  AngsanaUPC
'42.  Aparajita
'43.  Arabic Transparent
'44.  Arabic Typesetting
'45.  Arial
'46.  Arial Baltic
'47.  Arial Black
'48.  Arial CE
'49.  Arial CYR
'50.  Arial Greek
'51.  Arial Narrow
'52.  Arial Rounded MT Bold
'53.  Arial TUR
'54.  Arial Unicode MS
'55.  Avenida Alts LET
'56.  Avenida LET
'57.  Baskerville Old Face
'58.  Batang
'59.  BatangChe
'60.  Bauhaus 93
'61.  Bell MT
'62.  Bellevue BQ
'63.  Berlin Sans FB
'64.  Berlin Sans FB Demi
'65.  Bernard MT Condensed
'66.  Berthold-Script BQ Medium
'67.  Berthold-Script BQ Regular
'68.  Blackadder ITC
'69.  Bodoni MT
'70.  Bodoni MT Black
'71.  Bodoni MT Condensed
'72.  Bodoni MT Poster Compressed
'73.  Book Antiqua
'74.  Bookman Old Style
'75.  Bookshelf Symbol 7
'76.  Boulevard BQ
'77.  Bradley Hand ITC
'78.  Briem Akademi Std Cond
'79.  Britannic Bold
'80.  Broadway
'81.  Browallia New
'82.  BrowalliaUPC
'83.  Brush Script MT
'84.  Calibri
'85.  Calibri Light
'86.  Californian FB
'87.  Calisto MT
'88.  Cambria
'89.  Cambria Math
'90.  Candara
'91.  Castellar
'92.  Centaur
'93.  Century
'94.  Century Gothic
'95.  Century Schoolbook
'96.  Chiller
'97.  Clearface Gothic MT DemiBold
'98.  Code2000
'99.  Colonna MT
'100. Comic Sans MS
'101. Consolas
'102. Constantia
'103. Cooper Black
'104. Copperplate Gothic Bold
'105. Copperplate Gothic Light
'106. Corbel
'107. Cordia New
'108. CordiaUPC
'109. Coronet
'110. Courier
'111. Courier New
'112. Courier New Baltic
'113. Courier New CE
'114. Courier New CYR
'115. Courier New Greek
'116. Courier New TUR
'117. Curlz MT
'118. DaunPenh
'119. David
'120. Df Daves Raves TwoTT ITC
'121. DFKai-SB
'122. DilleniaUPC
'123. DokChampa
'124. Dotum
'125. DotumChe
'126. DragonWick
'127. Ebrima
'128. Edwardian Script ITC
'129. Elephant
'130. Engravers MT
'131. Eras Bold ITC
'132. Eras Demi ITC
'133. Eras Light ITC
'134. Eras Medium ITC
'135. Estrangelo Edessa
'136. EucrosiaUPC
'137. Euphemia
'138. FangSong
'139. Felix Titling
'140. Fixedsys
'141. Footlight MT Light
'142. Forte
'143. Franklin Gothic Book
'144. Franklin Gothic Demi
'145. Franklin Gothic Demi Cond
'146. Franklin Gothic Heavy
'147. Franklin Gothic Medium
'148. Franklin Gothic Medium Cond
'149. FrankRuehl
'150. FreesiaUPC
'151. Freestyle Script
'152. French Script MT
'153. Gabriola
'154. Garamond
'155. Gautami
'156. Georgia
'157. Gigi
'158. Gill Sans MT
'159. Gill Sans MT Condensed
'160. Gill Sans MT Ext Condensed Bold
'161. Gill Sans Ultra Bold
'162. Gill Sans Ultra Bold Condensed
'163. Gisha
'164. Gloucester MT Extra Condensed
'165. Goudy Old Style
'166. Goudy Stout
'167. Gulim
'168. GulimChe
'169. Gungsuh
'170. GungsuhChe
'171. Haettenschweiler
'172. Harlow Solid Italic
'173. Harrington
'174. Helvetica LT Std
'175. High Tower Text
'176. Hypatia Sans Pro
'177. Hypatia Sans Pro Black
'178. Hypatia Sans Pro ExtraLight
'179. Hypatia Sans Pro Light
'180. Hypatia Sans Pro Semibold
'181. Impact
'182. Imprint MT Shadow
'183. Informal Roman
'184. IrisUPC
'185. Iskoola Pota
'186. ITC Zapf Chancery
'187. ITC Zapf Dingbats
'188. JasmineUPC
'189. Jokerman
'190. Juice ITC
'191. KaiTi
'192. Kalinga
'193. Kartika
'194. Khmer UI
'195. KodchiangUPC
'196. Kokila
'197. Kristen ITC
'198. Kunstler Script
'199. Lao UI
'200. Latha
'201. Leelawadee
'202. Levenim MT
'203. LilyUPC
'204. Lucida Bright
'205. Lucida Calligraphy
'206. Lucida Console
'207. Lucida Fax
'208. Lucida Handwriting
'209. Lucida Sans
'210. Lucida Sans Typewriter
'211. Lucida Sans Unicode
'212. Magik
'213. Magneto
'214. Maiandra GD
'215. Malabar LT Pro Heavy
'216. Malgun Gothic
'217. Mangal
'218. Marigold
'219. Marlett
'220. Matura MT Script Capitals
'221. Meiryo
'222. Meiryo UI
'223. MICR
'224. Microsoft Himalaya
'225. Microsoft JhengHei
'226. Microsoft New Tai Lue
'227. Microsoft PhagsPa
'228. Microsoft Sans Serif
'229. Microsoft Tai Le
'230. Microsoft Uighur
'231. Microsoft YaHei
'232. Microsoft Yi Baiti
'233. MingLiU
'234. MingLiU_HKSCS
'235. MingLiU_HKSCS-ExtB
'236. MingLiU-ExtB
'237. Miriam
'238. Miriam Fixed
'239. Mistral
'240. Modern
'241. Modern No. 20
'242. ModulaSansRegular
'243. Mongolian Baiti
'244. Monotype Corsiva
'245. Monotype Sorts
'246. MoolBoran
'247. MS Gothic
'248. MS Mincho
'249. MS Outlook
'250. MS PGothic
'251. MS PMincho
'252. MS Reference Sans Serif
'253. MS Reference Specialty
'254. MS Sans Serif
'255. MS Serif
'256. MS UI Gothic
'257. MT Extra
'258. MV Boli
'259. Narkisim
'260. Niagara Engraved
'261. Niagara Solid
'262. NSimSun
'263. Nyala
'264. OCR A Extended
'265. Old English Text MT
'266. Onyx
'267. Palace Script MT
'268. Palatino Linotype
'269. Papyrus
'270. Parchment
'271. Perpetua
'272. Perpetua Titling MT
'273. Petemoss
'274. Plantagenet Cherokee
'275. Playbill
'276. PMingLiU
'277. PMingLiU-ExtB
'278. Poor Richard
'279. Pristina
'280. Raavi
'281. Rage Italic
'282. Ravie
'283. Rockwell
'284. Rockwell Condensed
'285. Rockwell Extra Bold
'286. Rod
'287. Roman
'288. Sakkal Majalla
'289. Script
'290. Script MT Bold
'291. Segoe Print
'292. Segoe Script
'293. Segoe UI
'294. Segoe UI Light
'295. Segoe UI Semibold
'296. Segoe UI Symbol
'297. Shonar Bangla
'298. Showcard Gothic
'299. Shruti
'300. SimHei
'301. Simplified Arabic
'302. Simplified Arabic Fixed
'303. SimSun
'304. SimSun-ExtB
'305. Sinaloa LET
'306. Small Fonts
'307. Snap ITC
'308. Stencil
'309. Stop
'310. SWGamekeys MT
'311. Sylfaen
'312. Symbol
'313. SymbolPS
'314. System
'315. Tahoma
'316. Tempus Sans ITC
'317. Terminal
'318. Times New Roman
'319. Times New Roman Baltic
'320. Times New Roman CE
'321. Times New Roman CYR
'322. Times New Roman Greek
'323. Times New Roman TUR
'324. Traditional Arabic
'325. Trebuchet MS
'326. Tunga
'327. Tw Cen MT
'328. Tw Cen MT Condensed
'329. Tw Cen MT Condensed Extra Bold
'330. Univers LT 45 Light
'331. Univers LT 55
'332. UpsideDownJJ
'333. Utsaah
'334. Vani
'335. Verdana
'336. Vijaya
'337. Viner Hand ITC
'338. Vivaldi
'339. Vladimir Script
'340. Vrinda
'341. Webdings
'342. Wide Latin
'343. Wingdings
'344. Wingdings 2
'345. Wingdings 3
'346. ZWAdobeF

EXITP:
2650    test_GetFontArray = blnRetVal
2660    Exit Function

ERRH:
2670    blnRetVal = False
2680    Select Case ERR.Number
        Case Else
2690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2700    End Select
2710    Resume EXITP

End Function
