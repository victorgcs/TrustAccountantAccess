Attribute VB_Name = "zz_mod_MDEPrepFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_MDEPrepFuncs"

'VGC 09/23/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const IsDev = 0  ' ** 0 = release; -1 = development.
' ** Also in:
' **   frmXAdmin_Misc
' **   modAutonumberFieldFuncs
' **   modExcelFuncs
' **   modVersionDocFuncs

#Const IsDemo = 0  ' ** 0 = new/upgrade; -1 = demo.
' ** Also in:
' **   modGlobConst
' **   modSecurityFunctions
' **   zz_mod_DatabaseDocFuncs

' ** #Const NoExcel = -1, NoExcel = 0
' ** C:\Program Files (x86)\Microsoft Office\OFFICE11\EXCEL.EXE

' ** Delete AutoKeys, rename AutoKeys_run to AutoKeys!

' ******************************************************
' ** See zz_mod_NewTestFuncs for checking TA.lic,
' ** and IniFile_Set_InstExp(), xIniFile_Set(), below.
' ******************************************************

' ** FORM OVER REPORT MEANS FORM'S PopUp SET True, SO SET False!

' ** 1.  For the Demo only, set IsDemo Compiler Constant to -1; see list of modules, above.
' ** 2.  Relink to the appropriate backends:
' **       Demo:     \NewDemo\DemoDatabase
' **       Upgrade:  \NewUpgrade\EmptyDatabase
' **       New:      \NewWorking\EmptyDatabase
' ** 3.  Check for event code where form property is empty:  VBA_Chk_Events(), below.
' ** 4.  Check for remarked "On Errors":                     VBA_Chk_OnErrorRem(), below.
' ** 5.  Update the VBAModules directory using:              VBA_ExportAll()
' ** 6.  Empty all temporary tables via mcrEmptyTmpTables.
' **     NOTE: The macro now also runs Setup_Demo(), below.
' ** 7.  Run zz_mcr_Clean_Templates.
' ** 8.  Run Setup_AutoNumber(), Below.
' ** 9.  Delete all "zz_qry_.." queries via zz_mcr_Delete_Query. They'll still remain in \StaticWork.
' ** 10. Delete any "zz_tbl_.." tables or links via zz_mcr_Delete_ZZTables. They'll still remain in \StaticWork.
' ** 11. Delete any "zz_.." forms or "zz_.." reports.
' ** 12. Delete the "zz_mcr_" macros, and AutoKeys.
' ** 13. Delete all "zz_mod_" modules except for:
' **       zz_mod_MDEPrepFuncs         (this)
' **       zz_mod_NewTestFuncs         (where TA.lic info can be checked and set)
' **       zz_mod_Display_Specs_Funcs  (for listing display specifications)
' ** 14. Change the Compiler Constant, above, to ISDEV = 0.
' ** 15. Delete any unneeded references, currently only 1. See VBA_Refs_OK() below.
' **       Microsoft Visual Basic for Applications Extensibility 5.3
' **       LEAVE! Microsoft Office 9.0 Object Library          NEEDED FOR REPORT COMMAND BAR STUFF!
' **       LEAVE! Microsoft ADO Ext. 6.0 for DDL and Security  NEEDED FOR RESETTING AUTONUMBER FIELDS!
' ** 16. Make sure the Demo hidden tables, all around, are empty!
' ** 17. Arrange Relationship diagram neatly.
' ** 18. Set ShowHidden Option back to unchecked (so demo license tables don't show).
' ** 19. Make sure the expired date in the \Untouched directory is set to 1/1/2008. See the functions in zz_mod_NewTestFuncs.
' ** 20. Move MDB to C:\Program Files\Delta Data\Trust Accountant\.
' **     This directory must already have been set up manually, with appropriate
' **     backend MDBs and TrustSec.mdw, as well as TA.lic, DDITrust.ini, and Trust.ico.
' *********************************
' ** FROM PROGRAM FILES LOCATION:
' *********************************
' ** 21. Relink all data to C:\Program Files\Delta Data\Trust Accountant\Database\.
' ** 22. Set Startup Options back to release settings:
' **       Display Form/Page : frmMenu_Title
' **       Application Icon  : C:\Program Files\Delta Data\Trust Accountant\
' **       All switches OFF, except for Display Status Bar
' ** 23. In the Database Window, move to the Pages tab!
' **     We have none, so someone opening with the Shift key held down will see nothing.
' ** 24. Compact & Repair WITH SHIFT KEY HELD DOWN!!!!
' ** 25. Create MDE WITH SHIFT KEY HELD DOWN!!!

' ** TA Programmers:
' **   gb ?
' **   Ed Perkins
' **   Mike ?
' **   Steve Shaw
' **   Victor Campbell

'###############################################################################
'### SRSEncrypt problems:
'###
'### Error 429: ActiveX component can't create object.
'### Re-register SRSEncrypt.dll
'###   C:\> RegSvr32 C:\Windows\System32\SRSEncrypt.dll
'###
'### SRSEncrypt.dll was written and compiled by Duane Johnson, Minneapolis, MN
'###
'###############################################################################

' ** CurrentUser      : superuser
' ** GetUserName()    : VictorC
' ** GetComputerName(): DELTADATA1

' ** References on an Access 2000 installation, IN PROPER ORDER:
' **   1. Visual Basic For Applications
' **        C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6.DLL
' **   2. Microsoft Access 9.0 Object Library
' **        C:\Program Files\Microsoft Office\Office\MSACC9.OLB
' **   3. Encode and Decode strings
' **        C:\Windows\system32\SRSENC~1.DLL
' **          Encrypt, Encode, Decrypt, Decode functions:
' **            CodeUtilities, GlobConst, modUtilities, frmDDate
' **   4. OLE Automation
' **        C:\Windows\system32\stdole2.tlb
' **          IUnknown object:
' **            Devices
' **   5. Microsoft Office 9.0 Object Library
' **        C:\Program Files\Microsoft Office\Office\MSO9.DLL
' **          CommandBar objects:
' **            zz_mod_MiscDevFuncs
' **   6. Microsoft DAO 3.6 Object Library
' **        C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll
' **          DAO objects:
' **            Numerous throughout
' **   7. Microsoft ActiveX Data Objects 2.1 Library
' **        C:\Program Files\Common Files\System\ado\msado21.tlb
' **          ADODB objects:
' **            frmRpt_Checks, rptChecks_Blank, clsCompanyInfo, clsEmployee, modCourtReports,
' **            modCourtReportsCA, modCourtReportsNY, zz_mod_Autonumber_Field_Funcs
' **   8. Microsoft Scripting Runtime
' **        C:\Windows\system32\scrrun.dll
' **          FileSystemObject object:
' **            modUtilities, frmBackup, frmRestoreOptions
' **   9. Microsoft ADO Ext. 6.0 for DDL and Security
' **        C:\Program Files\Common Files\System\ado\msadox.dll
' **          ADOX objects:
' **            zz_mod_Autonumber_Field_Funcs
' **  10. Microsoft Visual Basic for Applications Extensibility 5.3
' **        C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
' **          VBProject, VBComponent, CodeModule objects:
' **            zz_mod_Module_Format_Funcs

' ** Array: arr_varEvent().
Private lngEvents As Long, arr_varEvent As Variant
Private Const E_NAM   As Integer = 0
'Private Const E_ISFRM As Integer = 1
'Private Const E_ISRPT As Integer = 2
'Private Const E_ISCTL As Integer = 3
' **

Public Function FileAttr_Set(Optional varPathFile As Variant, Optional varAttribute As Variant, Optional varValue As Variant) As Boolean
' ** Hex Editor Neo permits changing all the dates!

100   On Error GoTo ERRH

        Const THIS_PROC As String = "FileAttr_Set"

        Dim fso As Scripting.FileSystemObject, fsfd As Scripting.Folder
        Dim fsfl As Scripting.File
        Dim strPathFile As String, strAttribute As String, varValueNew As Variant
        Dim strPath As String, strFile As String
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     If IsMissing(varPathFile) = True Then
130       strPathFile = "xx"
140       strAttribute = "yy"
150       varValueNew = Null
160     Else
170       If IsNull(varPathFile) = True Then
180         strPathFile = vbNullString
190       Else
200         If Trim(varPathFile) = vbNullString Then
210           strPathFile = vbNullString
220         Else
230           strPathFile = Trim(varPathFile)
240         End If
250       End If
260       If strPathFile <> vbNullString Then
270         strAttribute = Trim(varAttribute)
280         varValueNew = varValue
290       End If
300     End If

310     If strPathFile <> vbNullString Then

320       strPath = Parse_Path(strPathFile)  ' ** Module Function: modFileUtilities.
330       strFile = Parse_File(strPathFile)  ' ** Module Function: modFileUtilities.

340       Set fso = CreateObject("Scripting.FileSystemObject")
350       With fso

360         Set fsfd = .GetFolder(strPath)
370         Set fsfl = .GetFile(strFile)
380         With fsfl
390           Select Case strAttribute
              Case "DateCreated"
                ' ** READ-ONLY!!
                '.DateCreated = #8/1/209 7:01:01 PM#
400           Case "DateLastModified"
                ' ** READ-ONLY!!
                '.DateLastModified = #8/1/2009 7:05:01 PM#
410           End Select
420         End With

430       End With

440     Else
450       blnRetVal = False
460     End If

EXITP:
470     Set fsfl = Nothing
480     Set fsfd = Nothing
490     Set fso = Nothing
500     FileAttr_Set = blnRetVal
510     Exit Function

ERRH:
520     blnRetVal = False
530     Select Case ERR.Number
        Case Else
540       Beep
550       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
560     End Select
570     Resume EXITP

End Function

Public Function VBA_ExportAll() As Boolean
' ** Called by:
' **   QuikVBADoc() In zz_mod_ModuleDocFuncs

600   On Error GoTo ERRH

        Const THIS_PROC As String = "VBA_ExportAll"

        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent
        Dim fso As Scripting.FileSystemObject, fsfd As Scripting.Folder
        Dim fsfls As Scripting.FILES, fsfl As Scripting.File
        Dim lngMods As Long, arr_varMod() As Variant
        Dim lngDels As Long, arr_varDel() As Variant
        Dim strModName As String
        Dim strPath As String, strFile As String, strFileExt As String, strPathFile As String
        Dim blnFound As Boolean, blnDelete As Boolean
        Dim intPos01 As Integer
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varMod().
        Const M_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const M_MNAM As Integer = 0
        Const M_TYP  As Integer = 1
        Const M_PATH As Integer = 2
        Const M_FILE As Integer = 3

        ' ** Array: arr_varDel().
        Const D_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const D_NAM  As Integer = 0
        Const D_PATH As Integer = 1

610   On Error GoTo 0  ' ** I'd like to see the errors.

620     blnRetVal = True

630     strPath = gstrDir_Dev & LNK_SEP & "StaticRelease_2_2_24" & LNK_SEP & "VBAModules"

640     lngMods = 0&
650     ReDim arr_varMod(M_ELEMS, 0)

        ' ** Walk through every module.
660     Set vbp = Application.VBE.ActiveVBProject
670     With vbp
680       For Each vbc In .VBComponents
690         With vbc
700           strModName = .Name
710           strFileExt = vbNullString
720           Select Case .Type
              Case vbext_ct_ActiveXDesigner
730             strFileExt = "dsr"
740           Case vbext_ct_ClassModule
750             strFileExt = "cls"
760           Case vbext_ct_Document
770             strFileExt = "cls"  '"cld"  ' ** Access exports Form/Report modules as '.cls'.
780           Case vbext_ct_MSForm
790             strFileExt = "frm"
800           Case vbext_ct_StdModule
810             strFileExt = "bas"
820           Case Else
830             strFileExt = "txt"
840           End Select
850         End With
860         strFile = strModName & "." & strFileExt
870         lngMods = lngMods + 1&
880         lngE = lngMods - 1&
890         ReDim Preserve arr_varMod(M_ELEMS, lngE)
900         arr_varMod(M_MNAM, lngE) = strModName
910         arr_varMod(M_TYP, lngE) = .Type
920         arr_varMod(M_PATH, lngE) = strPath
930         arr_varMod(M_FILE, lngE) = strFile
940       Next
950     End With
960     Set vbc = Nothing
970     Set vbp = Nothing

        ' ** Export a text copy of every module to the \VBAModules directory.
980     For lngX = 0& To (lngMods - 1&)
990       strPathFile = arr_varMod(M_PATH, lngX) & LNK_SEP & arr_varMod(M_FILE, lngX)
1000      If FileExists(strPathFile) = True Then  ' ** Module Function: modFileUtilities.
1010        Kill strPathFile
1020      End If
1030      DoCmd.OutputTo acOutputModule, arr_varMod(M_MNAM, lngX), acFormatTXT, strPathFile
1040    Next

        ' ** Check for old module copies we can delete.
1050    Set fso = CreateObject("Scripting.FileSystemObject")
1060    With fso
1070      Set fsfd = .GetFolder(strPath)
1080      Set fsfls = fsfd.FILES
1090      For Each fsfl In fsfls
1100        With fsfl
1110          strFile = .Name
1120          intPos01 = InStr(strFile, ".")
1130          strFileExt = Mid(strFile, (intPos01 + 1))
1140          Select Case strFileExt
              Case "bas", "cls", "cld"
1150            blnFound = False
1160            For lngX = 0& To (lngMods - 1&)
1170              If arr_varMod(M_FILE, lngX) = strFile Then
1180                blnFound = True
1190                Exit For
1200              End If
1210            Next
1220            If blnFound = False Then
1230              lngDels = lngDels + 1&
1240              lngE = lngDels - 1&
1250              ReDim Preserve arr_varDel(D_ELEMS, lngE)
1260              arr_varDel(D_NAM, lngE) = strFile
1270              arr_varDel(D_PATH, lngE) = .Path  ' ** Includes path and file
1280            End If
1290          Case Else
                ' ** Don't care.
1300          End Select
1310        End With
1320      Next
1330    End With

1340    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
1350    DoEvents

1360    If lngDels > 0& Then
1370      For lngX = 0& To (lngDels - 1&)
1380        If InStr(arr_varDel(D_NAM, lngX), "zz_") = 0 Then
1390          blnDelete = True
1400          Debug.Print "'DEL? " & arr_varDel(D_NAM, lngX)
1410          Stop
1420          If blnDelete = True Then
1430            Kill arr_varDel(D_PATH, lngX)
1440          End If
1450        Else
1460          Debug.Print "'NOT DELETED: " & arr_varDel(D_NAM, lngX)
1470        End If
1480      Next
1490    End If

1500    Debug.Print "'DONE!  " & THIS_PROC & "()"

1510    Beep

        ' ** vbext_ComponentType enumeration:
        ' **     1  vbext_ct_StdModule        The specified module is a standard module
        ' **     2  vbext_ct_ClassModule      The specified module is a class module
        ' **     3  vbext_ct_MSForm           The specified module is behind a MSForms UserForm
        ' **     4  vbext_ct_ResFile          The specified module is for a Resource File
        ' **     5  vbext_ct_VBForm           The specified module is behind a Visual Basic Form
        ' **     6  vbext_ct_VBMDIForm        The specified module is behind a Visual Basic MDI Form
        ' **     7  vbext_ct_PropPage         The specified module is for a Property Page
        ' **     8  vbext_ct_UserControl      The specified module is behind a User-Defined Control
        ' **     9  vbext_ct_DocObject        The specified module is behind a Document Object
        ' **    10  vbext_ct_RelatedDocument  The specified module is for a Related Document
        ' **    11  vbext_ct_ActiveXDesigner  The specified module is behind an ActiveX Form
        ' **   100  vbext_ct_Document         The specified module is behind a Form or Report

EXITP:
1520    Set fsfl = Nothing
1530    Set fsfls = Nothing
1540    Set fsfd = Nothing
1550    Set fso = Nothing
1560    Set vbc = Nothing
1570    Set vbp = Nothing
1580    VBA_ExportAll = blnRetVal
1590    Exit Function

ERRH:
1600    blnRetVal = False
1610    Select Case ERR.Number
        Case Else
1620      Beep
1630      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1640    End Select
1650    Resume EXITP

End Function

Public Function IniFile_Set_InstExp() As Boolean
' ** See also:
' **   IniFile_Set() in modStartupFuncs.
' **   LicenseInfo_Get(), LicenseInfo_Set() in zz_mod_NewTestFuncs.

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "IniFile_Set_InstExp"

        Dim blnRetVal As Boolean

1710    If Len(TA_SEC) > Len(TA_SEC2) Then
1720      blnRetVal = xIniFile_Set("License", "Expires", EncodeString("01/01/2000"), _
            "C:\VictorGCS_Clients\TrustAccountant\NewDemo\DemoDatabase\" & gstrFile_LIC)
1730    Else
1740      blnRetVal = xIniFile_Set("License", "Expires", EncodeString("01/01/2000"), _
            "C:\VictorGCS_Clients\TrustAccountant\NewWorking\EmptyDatabase\" & gstrFile_LIC)
1750    End If

EXITP:
1760    IniFile_Set_InstExp = blnRetVal
1770    Exit Function

ERRH:
1780    blnRetVal = False
1790    Select Case ERR.Number
        Case Else
1800      Beep
1810      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1820    End Select
1830    Resume EXITP

End Function

Private Function xIniFile_Set(strSection As String, strSubSection As String, strValue As String, strFile As String) As Boolean
' ** Example:
' **   strSection = "License"
' **   strSubSection = "Firm"
' **   strValue = EncodeString(Me.txtLicensedTo)
' **   strFile = gstrTrustDataLocation & gstrFile_LIC
' ** See also:
' **   IniFile_Set() in modStartupFuncs.
' **   LicenseInfo_Get(), LicenseInfo_Set() in zz_mod_NewTestFuncs.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "xIniFile_Set"

        Dim lngRetVal As Long

1910    lngRetVal = WritePrivateProfileStringA(strSection, strSubSection, strValue, strFile)
1920    If lngRetVal = 0 Then
1930      xIniFile_Set = False
1940    Else
1950      xIniFile_Set = True
1960    End If

EXITP:
1970    Exit Function

ERRH:
1980    Select Case ERR.Number
        Case Else
1990      Beep
2000      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2010    End Select
2020    Resume EXITP

End Function

Public Function VBA_Refs_OK() As Boolean
' ** When the developer first opens the new MDE, this will
' ** let her/him know whether there are still development-only
' ** references attached. For release, it should always return True!

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "VBA_Refs_OK"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, ref As Access.Reference
        Dim strRefName As String
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

2110    blnRetVal = True

2120    If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
2130      If Right(Parse_File(CurrentDb.Name), 3) = gstrExt_AppRun Then
2140        Set dbs = CurrentDb
2150        With dbs
              ' ** tblReference, just ref_dev_only = True.
2160          Set qdf = .QueryDefs("qryReferences_01")
2170          Set rst = qdf.OpenRecordset()
2180          With rst
2190            If .BOF = True And .EOF = True Then
                  ' ** I know I've set 1 of them as Dev_Only!
2200            Else
2210              .MoveLast
2220              lngRecs = .RecordCount
2230              .MoveFirst
2240              For lngX = 1& To lngRecs
2250                strRefName = ![ref_name]
2260                For Each ref In Application.References
2270                  With ref
2280                    If .Name = strRefName Then
2290                      blnRetVal = False
2300                      Exit For
2310                    End If
2320                  End With
2330                Next
2340                If blnRetVal = False Then Exit For
2350                If lngX < lngRecs Then .MoveNext
2360              Next
2370            End If
2380            .Close
2390          End With
2400          .Close
2410        End With
2420      End If
2430    End If

EXITP:
2440    Set ref = Nothing
2450    Set rst = Nothing
2460    Set qdf = Nothing
2470    Set dbs = Nothing
2480    VBA_Refs_OK = blnRetVal
2490    Exit Function

ERRH:
2500    blnRetVal = False
2510    Select Case ERR.Number
        Case Else
2520      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2530    End Select
2540    Resume EXITP

End Function

Public Function VBA_Chk_OnErrorRem() As Boolean
' ** Look for 'On Error' statements that may have been left remarked-out for debugging.
' ** Requires Microsoft Visual Basic for Applications Extensibility 5.3.

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "VBA_Chk_OnErrorRem"

        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim strModName As String, strLastModName As String, strProcName As String, strLastProcName As String
        Dim lngLines As Long
        Dim strLine As String
        Dim blnFoundInMod As Boolean, blnFoundInProc As Boolean
        Dim blnNewProc As Boolean
        Dim lngOKs As Long, arr_varOK() As Variant
        Dim blnFound As Boolean
        Dim lngRems As Long, arr_varRem() As Variant
        Dim intPos01 As Integer, intPos02 As Integer
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varOK().
        Const O_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const O_MOD  As Integer = 0
        Const O_PROC As Integer = 1
        Const O_LIN  As Integer = 2

        ' ** Array: arr_varRem().
        Const R_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const R_MOD  As Integer = 0
        Const R_PROC As Integer = 1
        Const R_LIN  As Integer = 2

2610    blnRetVal = True

2620    lngOKs = 0&
2630    ReDim arr_varOK(O_ELEMS, 0)

        ' ** These remarked On Error's are OK; mostly within entire remarked-out procedures.
        lngOKs = lngOKs + 1&
        lngE = lngOKs - 1&
        ReDim Preserve arr_varOK(O_ELEMS, lngE)
        arr_varOK(O_MOD, lngE) = "Forms_frmJournal_Columns_Sub"
        arr_varOK(O_PROC, lngE) = "Detail_DblClick"
        arr_varOK(O_LIN, lngE) = 1326&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmAssets"
        'arr_varOK(O_PROC, lngE) = "Form_Load"
        'arr_varOK(O_LIN, lngE) = 660&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmArchiveTransactions"
        'arr_varOK(O_PROC, lngE) = "DateEnd_AfterUpdate"
        'arr_varOK(O_LIN, lngE) = 942&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmLicense"
        'arr_varOK(O_PROC, lngE) = "cmdOk_Click"
        'arr_varOK(O_LIN, lngE) = 86&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmLicense_Edit"
        'arr_varOK(O_PROC, lngE) = "cmdCopyCodes_Click"
        'arr_varOK(O_LIN, lngE) = 253&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmMenu_Title"
        'arr_varOK(O_PROC, lngE) = "Form_Load"
        'arr_varOK(O_LIN, lngE) = 46&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmMenu_Title"
        'arr_varOK(O_PROC, lngE) = "Form_Load"
        'arr_varOK(O_LIN, lngE) = 60&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmMenu_Utility"
        'arr_varOK(O_PROC, lngE) = "cmdRevCodes_Click"
        'arr_varOK(O_LIN, lngE) = 214&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmMenu_Utility"
        'arr_varOK(O_PROC, lngE) = "cmdRevCodes_Click"
        'arr_varOK(O_LIN, lngE) = 230&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmMenu_Utility"
        'arr_varOK(O_PROC, lngE) = "cmdUserMaintenance_Click"
        'arr_varOK(O_LIN, lngE) = 398&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "Form_frmMenu_Utility"
        'arr_varOK(O_PROC, lngE) = "Form_KeyDown"
        'arr_varOK(O_LIN, lngE) = 478&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "modUtilities"
        'arr_varOK(O_PROC, lngE) = "InitializeTables"
        'arr_varOK(O_LIN, lngE) = 1067&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "clsMonthCal"
        'arr_varOK(O_PROC, lngE) = "SetProperty"
        'arr_varOK(O_LIN, lngE) = 1913&

        'lngOKs = lngOKs + 1&
        'lngE = lngOKs - 1&
        'ReDim Preserve arr_varOK(O_ELEMS, lngE)
        'arr_varOK(O_MOD, lngE) = "clsMonthCal"
        'arr_varOK(O_PROC, lngE) = "SetProperty"
        'arr_varOK(O_LIN, lngE) = 1937&

2640    lngRems = 0&
2650    ReDim arr_varRem(R_ELEMS, 0)

2660    Set vbp = Application.VBE.ActiveVBProject
2670    With vbp
2680      strModName = vbNullString: strLastModName = vbNullString
2690      For Each vbc In .VBComponents
2700        With vbc

2710          If Left(.Name, 2) <> "z_" Then
                ' ** Ignore development modules.

2720            blnFoundInMod = False: blnFoundInProc = False
2730            blnNewProc = False
2740            strProcName = vbNullString: strLastProcName = vbNullString

2750            If .Name <> strLastModName Then strLastModName = strModName
2760            strModName = .Name

2770            Set cod = .CodeModule
2780            With cod

2790              lngLines = .CountOfLines

2800              For lngX = 1& To lngLines
                    ' ** .CodeModule.ProcOfLine(Line As Long, ProcKind As vbext_ProcKind) As String
                    ' **   Returns name of procedure that specified line is in.
                    ' **   Doesn't care if type of procedure is incorrect.

2810                strLine = vbNullString
2820                intPos01 = 0: intPos02 = 0

2830                If .ProcOfLine(lngX, vbext_pk_Proc) <> vbNullString Then
                      ' ** On Error will only be in a procedure.

2840                  strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
2850                  If strProcName <> strLastProcName Then
2860                    blnFoundInProc = False
2870                    strLastProcName = strProcName
2880                  End If

2890                  If blnFoundInProc = False Then
2900                    strLine = Trim(.Lines(lngX, 1))
2910                    If strLine <> vbNullString Then
2920                      If Left(strLine, 1) = "'" Then
2930                        intPos01 = InStr(strLine, "On Error ")
2940                        If intPos01 > 0 Then
2950                          intPos02 = InStr(intPos01, strLine, " GoTo ")
2960                          If intPos02 > 0 Then
2970                            If InStr(intPos02, strLine, "GoTo 0") = 0 Then
2980                              blnFound = False
2990                              For lngY = 0& To (lngOKs - 1&)
3000                                If arr_varOK(O_MOD, lngY) = strModName And _
                                        arr_varOK(O_PROC, lngY) = strProcName And _
                                        arr_varOK(O_LIN, lngY) = lngX Then
3010                                  blnFound = True
3020                                  Exit For
3030                                End If
3040                              Next
3050                              If blnFound = False Then
3060                                blnFoundInMod = True
                                    'blnFoundInProc = True  ' ** Let it list them all!
3070                                lngRems = lngRems + 1&
3080                                lngE = lngRems - 1&
3090                                ReDim Preserve arr_varRem(R_ELEMS, lngE)
3100                                arr_varRem(R_MOD, lngE) = strModName
3110                                arr_varRem(R_PROC, lngE) = strProcName
3120                                arr_varRem(R_LIN, lngE) = lngX
3130                              End If
3140                            End If
3150                          End If
3160                        End If
3170                      End If
3180                    End If
3190                  End If  ' ** blnFoundInProc.
3200                End If  ' ** Within procedure.
3210              Next  ' ** This Line: lngX.
3220            End With  ' ** This CodeModule: cod.
3230          End If  ' ** Not a zz_mod.. module.
3240        End With  ' ** This VBComponent: vbc.
3250      Next  ' ** For each VBComponent: vbc.
3260    End With  ' ** This VBProject: vbp.

3270    strLastModName = vbNullString
3280    If lngRems > 0& Then
3290      Debug.Print "'FOUND: " & lngRems
3300      For lngX = 0& To (lngRems - 1&)
3310        If arr_varRem(R_MOD, lngX) <> strLastModName Then
3320          Debug.Print "'" & arr_varRem(R_MOD, lngX)
3330          strLastModName = arr_varRem(R_MOD, lngX)
3340        End If
3350        Debug.Print "'  " & arr_varRem(R_LIN, lngX) & " " & arr_varRem(R_PROC, lngX)
3360      Next
3370    Else
3380      blnRetVal = False
3390      Debug.Print "'NONE FOUND!"
3400    End If

EXITP:
3410    Beep
3420    Set cod = Nothing
3430    Set vbc = Nothing
3440    Set vbp = Nothing
3450    VBA_Chk_OnErrorRem = blnRetVal
3460    Exit Function

ERRH:
3470    blnRetVal = False
3480    Select Case ERR.Number
        Case Else
3490      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3500    End Select
3510    Resume EXITP

End Function

Public Function VBA_Chk_Events() As Boolean
' ** Check to see if a form or report's module code is connected to the
' ** form/report; that is, there's an '[Event Procedure]' showing in the
' ** form/report's design view that matches what's in the module code.
' ** For some reason, these sometimes get disconnected!!

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "VBA_Chk_Events"

        'TRY THIS WITH SOME DESIGNATION ON THE 'As Reference'!
        Dim prj As Object, vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim frm As Access.Form, frm_ao As Access.AccessObject, rpt As Access.Report, rpt_ao As Access.AccessObject, ctl As Access.Control, Sec As Object, prp As Property
        Dim obj As Object
        Dim lngFrms As Long, arr_varFrm() As Variant
        Dim lngRpts As Long, arr_varRpt() As Variant
        Dim strModName As String, strFormName As String, strReportName As String, strObjType As String
        Dim strProcName As String, strEvent As String, strCtl As String, strCtlBase As String, strProp As String
        Dim lngLines As Long, lngLen As Long
        Dim lngDecLines As Long
        Dim strLine As String
        Dim blnProc As Boolean, blnFound As Boolean, blnChanged As Boolean
        Dim lngProcs As Long, arr_varProc() As Variant
        Dim lngChanges As Long, arr_varChange() As Variant
        Dim lngThisFormProcs As Long, lngThisFormChanges As Long, lngThisReportProcs As Long, lngThisReportChanges As Long
        Dim lngObjs As Long
        Dim lngProcsNotFound As Long
        Dim lngElemF As Long, lngElemP As Long, lngElemC As Long, lngElemR As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long
        Dim lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varFrm().
        Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const F_FNAM  As Integer = 0
        Const F_PROCS As Integer = 1
        Const F_CHGS  As Integer = 2

        ' ** Array: arr_varRpt().
        Const R_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const R_RNAM  As Integer = 0
        Const R_PROCS As Integer = 1
        Const R_CHGS  As Integer = 2

        ' ** Array: arr_varProc().
        Const P_ELEMS As Integer = 8  ' ** Array's first-element UBound().
        Const P_LINE  As Integer = 0
        Const P_NUM   As Integer = 1
        Const P_END   As Integer = 2
        Const P_PNAM  As Integer = 3
        Const P_CTL   As Integer = 4
        Const P_EVNT  As Integer = 5
        Const P_FELEM As Integer = 6
        Const P_RELEM As Integer = 7
        Const P_XEVNT As Integer = 8

        ' ** Array: arr_varChange().
        Const C_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const C_FELEM As Integer = 0
        Const C_FNAM  As Integer = 1
        Const C_RELEM As Integer = 2
        Const C_RNAM  As Integer = 3
        Const C_PROC  As Integer = 4

        'TRY THIS WITH SOME DESIGNATION ON THE 'As Reference'!
3610    blnRetVal = True

3620    lngProcsNotFound = 0&

3630    lngFrms = 0&
3640    ReDim arr_varFrm(F_ELEMS, 0)

3650    lngProcs = 0&
3660    ReDim arr_varProc(P_ELEMS, 0)

3670    lngChanges = 0&
3680    ReDim arr_varChange(C_ELEMS, 0)

        ' ** Get a list of all the forms with a class module.
3690    Set prj = Application.CurrentProject
3700    With prj
3710      lngFrms = .AllForms.Count
3720      ReDim arr_varFrm(F_ELEMS, (lngFrms - 1&))
3730      lngElemF = -1&
3740      For Each frm_ao In .AllForms
3750        lngElemF = lngElemF + 1&
3760        arr_varFrm(F_FNAM, lngElemF) = frm_ao.Name
3770        arr_varFrm(F_PROCS, lngElemF) = 0&
3780        arr_varFrm(F_CHGS, lngElemF) = 0&
3790      Next
3800    End With
3810    Set prj = Nothing

        ' ** Binary Sort arr_varFrm() array.
        ' ** Only the name needs to be moved, since all the other elements are empty.
3820    For lngX = UBound(arr_varFrm, 2) To 1& Step -1&
3830      For lngY = 0& To (lngX - 1&)
3840        If arr_varFrm(F_FNAM, lngY) > arr_varFrm(F_FNAM, (lngY + 1)) Then
3850          strTmp01 = arr_varFrm(F_FNAM, lngY)
3860          arr_varFrm(F_FNAM, lngY) = arr_varFrm(F_FNAM, (lngY + 1))
3870          arr_varFrm(F_FNAM, (lngY + 1)) = strTmp01
3880        End If
3890      Next
3900    Next

        ' ** Get a list of all the reports with a class module.
3910    Set prj = Application.CurrentProject
3920    With prj
3930      lngRpts = .AllReports.Count
3940      ReDim arr_varRpt(R_ELEMS, (lngRpts - 1&))
3950      lngElemF = -1&
3960      For Each rpt_ao In .AllReports
3970        lngElemF = lngElemF + 1&
3980        arr_varRpt(R_RNAM, lngElemF) = rpt_ao.Name
3990        arr_varRpt(R_PROCS, lngElemF) = 0&
4000        arr_varRpt(R_CHGS, lngElemF) = 0&
4010      Next
4020    End With
4030    Set prj = Nothing

        ' ** Binary Sort arr_varRpt() array.
        ' ** Only the name needs to be moved, since all the other elements are empty.
4040    For lngX = UBound(arr_varRpt, 2) To 1& Step -1&
4050      For lngY = 0& To (lngX - 1&)
4060        If arr_varRpt(R_RNAM, lngY) > arr_varRpt(R_RNAM, (lngY + 1)) Then
4070          strTmp01 = arr_varRpt(R_RNAM, lngY)
4080          arr_varRpt(R_RNAM, lngY) = arr_varRpt(R_RNAM, (lngY + 1))
4090          arr_varRpt(R_RNAM, (lngY + 1)) = strTmp01
4100        End If
4110      Next
4120    Next

        ' ** Load arr_varEvent() array.
4130    VBA_Event_Load  ' ** Function: Below.

        ' ** Walk through every module and collect array of procedures.
4140    Set vbp = Application.VBE.ActiveVBProject
4150    With vbp
4160      For Each vbc In .VBComponents
4170        With vbc
4180          strModName = .Name
4190          strFormName = vbNullString: strReportName = vbNullString
4200          intPos01 = InStr(strModName, "_")
4210          If intPos01 > 0 Then
4220            strObjType = Left(strModName, (intPos01 - 1))
4230            Select Case strObjType
                Case "Form"
4240              strFormName = Mid(strModName, 6)
4250            Case "Report"
4260              strReportName = Mid(strModName, 8)
4270            Case Else
                  ' ** No interested.
4280            End Select
4290          End If
4300          If strFormName <> vbNullString Or strReportName <> vbNullString Then
4310            blnFound = False
4320            lngElemF = -1&: lngElemR = -1&
4330            Select Case strObjType
                Case "Form"
4340              For lngX = 0& To (lngFrms - 1&)
4350                If arr_varFrm(F_FNAM, lngX) = strFormName Then
4360                  blnFound = True
4370                  lngElemF = lngX
4380                End If
4390              Next
4400            Case "Report"
4410              For lngX = 0& To (lngRpts - 1&)
4420                If arr_varRpt(R_RNAM, lngX) = strReportName Then
4430                  blnFound = True
4440                  lngElemR = lngX
4450                End If
4460              Next
4470            End Select
4480            If blnFound = False Then
4490              Beep
4500              Debug.Print "'" & UCase$(strObjType) & " NOT FOUND!"
4510            End If
4520            Set cod = .CodeModule
4530            With cod
4540              lngLines = .CountOfLines
4550              lngDecLines = .CountOfDeclarationLines
4560              blnProc = False
4570              lngThisFormProcs = 0&: lngThisReportProcs = 0&
4580              For lngX = (lngDecLines + 1&) To lngLines
4590                strLine = Trim(.Lines(lngX, 1))
4600                If strLine <> vbNullString Then
4610                  If Left(strLine, 1) <> "'" Then
4620                    intPos01 = InStr(strLine, " ")
4630                    If intPos01 > 0 Then
4640                      strTmp01 = Trim(Left(strLine, (intPos01 - 1)))
4650                      If IsNumeric(strTmp01) = False Then
4660                        Select Case strTmp01
                            Case "Private"
4670                          blnProc = True
4680                        Case "Public"
4690                          blnProc = True
4700                        Case "Sub"
4710                          blnProc = True
4720                        Case "Function"
4730                          blnProc = True
4740                        Case Else
4750                          ' ** We don't care.
4760                        End Select
4770                        If blnProc = True Then
4780                          Select Case strObjType
                              Case "Form"
4790                            lngThisFormProcs = lngThisFormProcs + 1&
4800                          Case "Report"
4810                            lngThisReportProcs = lngThisReportProcs + 1&
4820                          End Select
4830                          lngProcs = lngProcs + 1&
4840                          lngE = lngProcs - 1&
4850                          ReDim Preserve arr_varProc(P_ELEMS, lngE)
4860                          arr_varProc(P_LINE, lngE) = strLine
4870                          arr_varProc(P_NUM, lngE) = lngX
4880                          arr_varProc(P_END, lngE) = 0
4890                          Select Case strObjType
                              Case "Form"
4900                            arr_varProc(P_FELEM, lngE) = lngElemF
4910                          Case "Report"
4920                            arr_varProc(P_RELEM, lngE) = lngElemR
4930                          End Select
4940                          arr_varProc(P_XEVNT, lngE) = True
4950                          arr_varProc(P_PNAM, lngE) = vbNullString
4960                          blnProc = False
4970                        Else
4980                          If strTmp01 = "End" Then
4990                            If lngProcs > 0& Then
5000                              If arr_varProc(P_END, (lngProcs - 1&)) = 0& Then
5010                                arr_varProc(P_END, (lngProcs - 1&)) = lngX
5020                              End If
5030                            End If
5040                          End If
5050                        End If
5060                      End If
5070                    End If
5080                  End If
5090                End If
5100              Next  ' ** For each Line: lngX
5110            End With  ' ** This CodeModule: cod.
5120            Select Case strObjType
                Case "Form"
5130              arr_varFrm(F_PROCS, lngElemF) = lngThisFormProcs
5140            Case "Report"
5150              arr_varRpt(R_PROCS, lngElemR) = lngThisReportProcs
5160            End Select
5170          End If  ' ** Form or Report.
5180        End With  ' ** This VBComponent: vbc
5190      Next  ' ** For each VBComponent: vbc
5200    End With  ' ** This VBProject: vbp.

5210    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
5220    DoEvents

5230    If lngProcs > 0& Then

          ' ** Update array with potential events.
5240      For lngX = 0& To (lngProcs - 1&)
5250        lngElemP = lngX
5260        intPos01 = InStr(arr_varProc(P_LINE, lngElemP), "Sub ")  ' ** Ignore Functions and Properties.
5270        If intPos01 > 0 Then
5280          strTmp01 = Trim(Mid(arr_varProc(P_LINE, lngElemP), (intPos01 + 4)))
5290          intPos01 = InStr(strTmp01, "(")
5300          If intPos01 > 0 Then
5310            strProcName = Left(strTmp01, (intPos01 - 1))
5320            arr_varProc(P_PNAM, lngElemP) = strProcName
5330            intPos01 = InStr(strProcName, "_")
5340            If intPos01 > 0 Then
5350              lngLen = Len(strProcName)
                  ' ** Look for last underscore, not first.
5360              For lngY = lngLen To 1& Step -1&
5370                If Mid(strProcName, lngY, 1) = "_" Then
5380                  strEvent = Mid(strProcName, (lngY + 1))
5390                  strCtl = Left(strProcName, (lngY - 1))
5400                  arr_varProc(P_EVNT, lngElemP) = strEvent
5410                  arr_varProc(P_CTL, lngElemP) = strCtl
5420                  Exit For
5430                End If
5440              Next
5450            Else
                  ' ** Not a potential event, since they all have an underscore.
5460              arr_varProc(P_XEVNT, lngElemP) = False
5470            End If
5480          End If
5490        Else
              ' ** Events can only be Sub's; can't be Functions or Properties.
5500        End If
5510        If IsEmpty(arr_varProc(P_FELEM, lngElemP)) = True Then
5520          arr_varProc(P_FELEM, lngElemP) = -1&
5530        End If
5540        If IsEmpty(arr_varProc(P_RELEM, lngElemP)) = True Then
5550          arr_varProc(P_RELEM, lngElemP) = -1&
5560        End If
5570      Next

          ' ** Now cross-check potential events.
5580      For lngW = 1& To 2&

5590        Select Case lngW
            Case 1&  ' ** Forms.
5600          lngObjs = lngFrms
5610        Case 2&  ' ** Reports.
5620          lngObjs = lngRpts
5630        End Select
5640        For lngX = 0& To (lngObjs - 1&)

5650          Select Case lngW
              Case 1&  ' ** Forms.
5660            lngElemF = lngX
5670            lngTmp03 = arr_varFrm(F_PROCS, lngElemF)
5680          Case 2&  ' ** Reports.
5690            lngElemR = lngX
5700            lngTmp03 = arr_varRpt(R_PROCS, lngElemR)
5710          End Select
5720          If lngTmp03 > 0& Then
                ' ** This object has procedures.

5730            Select Case lngW
                Case 1&  ' ** Forms.
5740              strFormName = arr_varFrm(F_FNAM, lngElemF)
5750              strReportName = vbNullString
5760            Case 2&  ' ** Reports.
5770              strFormName = vbNullString
5780              strReportName = arr_varRpt(R_RNAM, lngElemR)
5790            End Select
5800            lngThisFormChanges = 0&: lngThisReportChanges = 0&
5810            blnChanged = False

                ' ** Open the object associated with this module.
5820            Select Case lngW
                Case 1&  ' ** Form.
5830              DoCmd.OpenForm strFormName, acDesign, , , , acHidden
5840              Set obj = Forms(strFormName)
5850            Case 2&  ' ** Report.
5860              DoCmd.OpenReport strReportName, acViewDesign, , , acHidden
5870              Set obj = Reports(strReportName)
5880            End Select
5890            With obj
5900              For lngY = 0& To (lngProcs - 1&)
5910                lngElemP = lngY
5920                If arr_varProc(P_XEVNT, lngElemP) = True Then
5930                  blnFound = False
5940                  Select Case lngW
                      Case 1&  ' ** Form.
5950                    If arr_varProc(P_FELEM, lngElemP) = lngElemF Then
5960                      blnFound = True
5970                    End If
5980                  Case 2&  ' ** Report.
5990                    If arr_varProc(P_RELEM, lngElemP) = lngElemR Then
6000                      blnFound = True
6010                    End If
6020                  End Select
6030                  If blnFound = True Then
6040                    If arr_varProc(P_PNAM, lngElemP) <> vbNullString Then

6050                      strProcName = arr_varProc(P_PNAM, lngElemP)
6060                      strCtl = arr_varProc(P_CTL, lngElemP)
6070                      strEvent = arr_varProc(P_EVNT, lngElemP)
6080                      strProp = vbNullString
6090                      Select Case strEvent
                          Case "BeforeInsert", "AfterInsert", "BeforeUpdate", "AfterUpdate", "BeforeDelConfirm", "AfterDelConfirm"
                            ' ** These are the only known events that don't begin with 'On'.
6100                        strProp = strEvent
6110                      Case Else
                            ' *****************************************************
                            ' ** Array: arr_varEvent()
                            ' **
                            ' **   Field  Element  Name                Constant
                            ' **   =====  =======  ==================  ==========
                            ' **     1       0     vbcom_event_name    E_NAM
                            ' **     2       1     vbcom_frm           E_ISFRM
                            ' **     3       2     vbcom_rpt           E_ISRPT
                            ' **     4       3     vbcom_ctl           E_ISCTL
                            ' **
                            ' *****************************************************
6120                        For lngZ = 0& To (lngEvents - 1&)
6130                          If arr_varEvent(E_NAM, lngZ) = ("On" & strEvent) Then
6140                            strProp = "On" & strEvent
6150                            Exit For
6160                          End If
6170                        Next
6180                      End Select

6190                      If strProp <> vbNullString Then

6200                        If IsNumeric(Right(strCtl, 1)) = True Then
6210                          strCtlBase = Left(strCtl, (Len(strCtl) - 1))
6220                        Else
6230                          strCtlBase = strCtl
6240                        End If

6250                        Select Case strCtlBase
                            Case "Form", "Report"  ' ** Form or Report property.
6260                          blnFound = False
6270                          For Each prp In .Properties
6280                            With prp
6290                              If .Name = strProp Then
6300                                blnFound = True
6310                                Exit For
6320                              End If
6330                            End With
6340                          Next
6350                          If blnFound = False Then
6360                            If strFormName = "frmXAdmin_Registry" And _
                                    (strProp = "OnNodeClick" Or strProp = "OnExpand" Or strProp = "OnCollapse" Or strProp = "OnMouseMove") Then
                                  ' ** These won't be found.
6370                            Else
6380                              Debug.Print "'1. NOT FOUND!  " & IIf(lngW = 1&, "FRM", "RPT") & ": " & _
                                    IIf(lngW = 1&, strFormName, strReportName) & "  CTL: " & .Name & "  EVT: " & strProp & "()"
6390                            End If
                                'NOT FOUND!  FRM: frmXAdmin_Registry  CTL: Registry_tvw  EVT: OnNodeClick()
                                'NOT FOUND!  FRM: frmXAdmin_Registry  CTL: Registry_tvw  EVT: OnExpand()
                                'NOT FOUND!  FRM: frmXAdmin_Registry  CTL: Registry_tvw  EVT: OnCollapse()
                                'NOT FOUND!  FRM: frmXAdmin_Registry  CTL: Registry_tvw  EVT: OnMouseMove()
6400                          Else
6410                            If .Properties(strProp) <> "[Event Procedure]" Then
6420                              .Properties(strProp) = "[Event Procedure]"
6430                              lngChanges = lngChanges + 1&
6440                              lngElemC = lngChanges - 1&
6450                              ReDim Preserve arr_varChange(C_ELEMS, lngElemC)
6460                              Select Case lngW
                                  Case 1&  ' ** Form.
6470                                arr_varChange(C_FELEM, lngElemC) = lngElemF
6480                                arr_varChange(C_FNAM, lngElemC) = strFormName
6490                                lngThisFormChanges = lngThisFormChanges + 1&
6500                              Case 2&  ' ** Report.
6510                                arr_varChange(C_RELEM, lngElemC) = lngElemR
6520                                arr_varChange(C_RNAM, lngElemC) = strReportName
6530                                lngThisReportChanges = lngThisReportChanges + 1&
6540                              End Select
6550                              arr_varChange(C_PROC, lngElemC) = strProcName
6560                              blnChanged = True
6570                            End If
6580                          End If
6590                        Case "Detail", "FormHeader", "ReportHeader", "FormFooter", "ReportFooter", _
                                "PageHeader", "PageHeaderSection", "PageFooter", "PageFooterSection", "GroupHeader", "GroupFooter", _
                                acPageHeader, "PageFooterSection"  ' ** Form/Report Section property.
6600  On Error Resume Next
6610                          Select Case strCtl
                              Case "Detail":            Set Sec = .Detail
6620                          Case "Detail0":           Set Sec = .Detail0
6630                          Case "Detail1":           Set Sec = .Detail1
6640                          Case "FormHeader":        Set Sec = .FormHeader
6650                          Case "FormHeader0":       Set Sec = .FormHeader0
6660                          Case "FormHeader1":       Set Sec = .FormHeader1
6670                          Case "ReportHeader":      Set Sec = .ReportHeader
6680                          Case "ReportHeader0":     Set Sec = .ReportHeader0
6690                          Case "ReportHeader1":     Set Sec = .ReportHeader1
6700                          Case "FormFooter":        Set Sec = .FormFooter
6710                          Case "FormFooter0":       Set Sec = .FormFooter0
6720                          Case "FormFooter1":       Set Sec = .FormFooter1
6730                          Case "FormFooter2":       Set Sec = .FormFooter2
6740                          Case "ReportFooter":      Set Sec = .ReportFooter
6750                          Case "ReportFooter0":     Set Sec = .ReportFooter0
6760                          Case "ReportFooter1":     Set Sec = .ReportFooter1
6770                          Case "ReportFooter2":     Set Sec = .ReportFooter2
6780                          Case "ReportFooter3":     Set Sec = .ReportFooter3
6790                          Case "ReportFooter4":     Set Sec = .ReportFooter4
6800                          Case "PageHeader":        Set Sec = .PageHeader
6810                          Case "PageHeader0":       Set Sec = .PageHeader0
6820                          Case "PageFooter":        Set Sec = .PageFooter
6830                          Case "PageFooter0":       Set Sec = .PageFooter0
6840                          Case "PageHeaderSection": Set Sec = .PageHeaderSection
6850                          Case "PageFooterSection": Set Sec = .PageFooterSection
6860                          Case "GroupHeader":       Set Sec = .GroupHeader
6870                          Case "GroupHeader0":      Set Sec = .GroupHeader0
6880                          Case "GroupHeader1":      Set Sec = .GroupHeader1
6890                          Case "GroupHeader2":      Set Sec = .GroupHeader2
6900                          Case "GroupHeader3":      Set Sec = .GroupHeader3
6910                          Case "GroupHeader4":      Set Sec = .GroupHeader4
6920                          Case "GroupHeader5":      Set Sec = .GroupHeader5
6930                          Case "GroupHeader6":      Set Sec = .GroupHeader6
6940                          Case "GroupHeader7":      Set Sec = .GroupHeader7
6950                          Case "GroupHeader8":      Set Sec = .GroupHeader8
6960                          Case "GroupFooter":       Set Sec = .GroupFooter
6970                          Case "GroupFooter0":      Set Sec = .GroupFooter0
6980                          Case "GroupFooter1":      Set Sec = .GroupFooter1
6990                          Case "GroupFooter2":      Set Sec = .GroupFooter2
7000                          Case "GroupFooter3":      Set Sec = .GroupFooter3
7010                          Case "GroupFooter4":      Set Sec = .GroupFooter4
7020                          Case "GroupFooter5":      Set Sec = .GroupFooter5
7030                          Case "GroupFooter6":      Set Sec = .GroupFooter6
7040                          Case "GroupFooter7":      Set Sec = .GroupFooter7
7050                          Case "GroupFooter8":      Set Sec = .GroupFooter8
7060                          End Select
7070                          If ERR = 0 Then
7080  On Error GoTo 0
7090                            With Sec
7100                              blnFound = False
7110                              For Each prp In .Properties
7120                                With prp
7130                                  If .Name = strProp Then
7140                                    blnFound = True
7150                                    Exit For
7160                                  End If
7170                                End With
7180                              Next
7190                              If blnFound = False Then
7200                                If strFormName = "frmXAdmin_Registry" And _
                                        (strProp = "OnNodeClick" Or strProp = "OnExpand" Or strProp = "OnCollapse" Or strProp = "OnMouseMove") Then
                                      ' ** These won't be found.
7210                                Else
7220                                  Debug.Print "'2. NOT FOUND!  " & IIf(lngW = 1&, "FRM", "RPT") & ": " & _
                                        IIf(lngW = 1&, strFormName, strReportName) & "  CTL: " & .Name & "  EVT: " & strProp & "()"
7230                                End If
7240                              Else
7250                                If .Properties(strProp) <> "[Event Procedure]" Then
7260                                  .Properties(strProp) = "[Event Procedure]"
7270                                  lngChanges = lngChanges + 1&
7280                                  lngElemC = lngChanges - 1&
7290                                  ReDim Preserve arr_varChange(C_ELEMS, lngElemC)
7300                                  Select Case lngW
                                      Case 1&  ' ** Form.
7310                                    arr_varChange(C_FELEM, lngElemC) = lngElemF
7320                                    arr_varChange(C_FNAM, lngElemC) = strFormName
7330                                    lngThisFormChanges = lngThisFormChanges + 1&
7340                                  Case 2&  ' ** Report.
7350                                    arr_varChange(C_RELEM, lngElemC) = lngElemR
7360                                    arr_varChange(C_RNAM, lngElemC) = strReportName
7370                                    lngThisReportChanges = lngThisReportChanges + 1&
7380                                  End Select
7390                                  arr_varChange(C_PROC, lngElemC) = strProcName
7400                                  blnChanged = True
7410                                End If
7420                              End If
7430                            End With
7440                          Else
                                ' ** A section that's been deleted!
7450  On Error GoTo 0
7460                          End If
7470                          Set Sec = Nothing
7480                        Case Else
7490                          blnFound = False
7500                          For Each ctl In .Controls
7510                            With ctl
7520                              If .Name = strCtl Then
7530                                blnFound = True
7540                                Exit For
7550                              End If
7560                            End With
7570                          Next
7580                          If blnFound = False Then
                                ' ** Try it with spaces instead of underscores.
7590                            strTmp02 = strCtl
7600                            intPos01 = InStr(strTmp02, "_")
7610                            If intPos01 > 0 Then
7620                              Do While intPos01 > 0
7630                                strTmp02 = Left(strTmp02, (intPos01 - 1)) & " " & Mid(strTmp02, (intPos01 + 1))
7640                                intPos01 = InStr(strTmp02, "_")
7650                              Loop
7660                              strTmp02 = Trim(strTmp02)
7670                              For Each ctl In .Controls
7680                                With ctl
7690                                  If .Name = strTmp02 Then
7700                                    blnFound = True
7710                                    Exit For
7720                                  End If
7730                                End With
7740                              Next
7750                              If blnFound = True Then
7760                                arr_varProc(P_CTL, lngElemP) = strTmp02
7770                                strCtl = strTmp02
7780                              End If
7790                            End If
7800                          End If
7810                          If blnFound = False Then
                                ' ** Try it with a slash instead of underscore.
                                'SUB NOT EVENT? Map 'share_face_Enter'  share/face
7820                            strTmp02 = strCtl
7830                            intPos01 = InStr(strTmp02, "_")
7840                            If intPos01 > 0 Then
7850                              Do While intPos01 > 0
7860                                strTmp02 = Left(strTmp02, (intPos01 - 1)) & "/" & Mid(strTmp02, (intPos01 + 1))
7870                                intPos01 = InStr(strTmp02, "_")
7880                              Loop
7890                              strTmp02 = Trim(strTmp02)
7900                              For Each ctl In .Controls
7910                                With ctl
7920                                  If .Name = strTmp02 Then
7930                                    blnFound = True
7940                                    Exit For
7950                                  End If
7960                                End With
7970                              Next
7980                              If blnFound = True Then
7990                                arr_varProc(P_CTL, lngElemP) = strTmp02
8000                                strCtl = strTmp02
8010                              End If
8020                            End If
8030                          End If
8040                          If blnFound = True Then
8050                            Set ctl = .Controls(strCtl)
8060                            With ctl
8070                              blnFound = False
8080                              For Each prp In .Properties
8090                                With prp
8100                                  If .Name = strProp Then
8110                                    blnFound = True
8120                                    Exit For
8130                                  End If
8140                                End With
8150                              Next
8160                              If blnFound = False Then
8170                                If strFormName = "frmXAdmin_Registry" And _
                                        (strProp = "OnNodeClick" Or strProp = "OnExpand" Or strProp = "OnCollapse" Or strProp = "OnMouseMove") Then
                                      ' ** These won't be found.
                                    ElseIf strFormName = "frmTransaction_Audit" And (strProp = "OnLoad") Then
                                      ' ** frmTransaction_Audit happens to have a control name that matches several Subs and Functions.
                                      'X 3. NOT FOUND!  FRM: frmTransaction_Audit  CTL: FilterRecs  EVT: OnLoad()
8180                                Else
8190                                  Debug.Print "'3. NOT FOUND!  " & IIf(lngW = 1&, "FRM", "RPT") & ": " & _
                                        IIf(lngW = 1&, strFormName, strReportName) & "  CTL: " & .Name & "  EVT: " & strProp & "()"
8200                                End If
8210                              Else
8220                                If .Properties(strProp) <> "[Event Procedure]" Then
8230                                  .Properties(strProp) = "[Event Procedure]"
8240                                  lngChanges = lngChanges + 1&
8250                                  lngElemC = lngChanges - 1&
8260                                  ReDim Preserve arr_varChange(C_ELEMS, lngElemC)
8270                                  Select Case lngW
                                      Case 1&  ' ** Form.
8280                                    arr_varChange(C_FELEM, lngElemC) = lngElemF
8290                                    arr_varChange(C_FNAM, lngElemC) = strFormName
8300                                    lngThisFormChanges = lngThisFormChanges + 1&
8310                                  Case 2&  ' ** Report.
8320                                    arr_varChange(C_RELEM, lngElemC) = lngElemR
8330                                    arr_varChange(C_RNAM, lngElemC) = strReportName
8340                                    lngThisReportChanges = lngThisReportChanges + 1&
8350                                  End Select
8360                                  arr_varChange(C_PROC, lngElemC) = strProcName
8370                                  blnChanged = True
8380                                End If
8390                              End If
8400                            End With
8410                          Else  ' ** blnFound = False.
8420                            If (strFormName = "frmIncomeExpenseCodes" And (strProcName = "cmdSave_Click" Or strProcName = "RevO_Load")) Or _
                                    ((strFormName = "frmAccountProfile" Or strFormName = "frmCompanyInfo" Or _
                                    strFormName = "frmLocation" Or strFormName = "frmLocationEditAdd" Or strFormName = "frmMasterBalance") And _
                                    strProcName = "cmdSave_Click") Or (strFormName = "frmAccountSearch" And _
                                    strProcName = "cmdClose_Click") Or _
                                    (strFormName = "frmAccountHideTrans_Match" And strProcName = "cmdSave_Click") Or _
                                    (strFormName = "frmTaxLot" And strProcName = "DoMultiLots_Sort") Or _
                                    (strFormName = "frmVersion_Main" And strProcName = "ConversionCheck_Response") Then
                                  ' ** I know about these.
8430                            ElseIf (strFormName = "frmPostingDate" And strProcName = "JrnlCol_Load") Or _
                                    (strFormName = "frmJournal_Columns" And strProcName = "StatusBar_Load") Or _
                                    (strFormName = "frmJournal_Columns_Sub" And strProcName = "StatusBar_Load") Then
                                  ' ** And these.
8440                            ElseIf ((strFormName = "frmXAdmin_Query_Sub" Or strFormName = "frmXAdmin_Table_Sub" Or _
                                    strFormName = "frmXAdmin_Misc_Pref_Sub") And strProcName = "cmdDelete_Click") Or _
                                    (strFormName = "frmAccountHideTrans_Hidden" And strProcName = "cmdRegen_Click") Or _
                                    (strFormName = "frmMenu_Utility" And strProcName = "cmdArchiveReports_Click") Then
                                  ' ** And these.
8450                            ElseIf ((strProcName = "cmdSave_Click") Or ((strFormName = "frmJournal" Or strFormName = "frmLoadTimer") And _
                                    (strProcName = "cmdClose_Click"))) Then
                                  ' ** Every one of them!
8460                            ElseIf (strFormName = "frmXAdmin_Registry" And _
                                    (strProcName = "NodeClick" Or strProcName = "Expand" Or _
                                    strProcName = "Collapse" Or strProcName = "RegTreeView_Load")) Then
                                  ' ** TreeView control events.
8470                            ElseIf (((strFormName = "frmXAdmin_Form_Graphics") And (strProcName = "Staging_Load")) Or _
                                    ((strFormName = "frmXAdmin_Form_Graphics_Sub") And (strProcName = "RecSel_Load")) Or _
                                    ((strFormName = "frmXAdmin_Graphics_Sub") And (strProcName = "GfxFormat_Load"))) Then
                                  ' ** These are fine.
8480                            ElseIf ((((strFormName = "frmTaxLot" Or strFormName = "frmJournal_Columns_TaxLot") And _
                                    (strProcName = "RecArray_Load")) Or (strProcName = "NoChar_Load")) Or _
                                    ((strFormName = "frmMenu_Account") And (strProcName = "CtlArray_Load"))) Then
                                  ' ** OK!
8490                            ElseIf (strFormName = "frmTransaction_Audit" Or strFormName = "frmTransaction_Audit_Sub" Or _
                                    strFormName = "frmTransaction_Audit_Sub_ds") And (strProcName = "SubFrmGfx_Load" Or _
                                    strProcName = "PrintTgls_Load" Or strProcName = "PrintTgls_Click" Or strProcName = "ColArray_Load") Then
                                  ' ** OK!
8500                            ElseIf ((strFormName = "frmRpt_CourtReports_NY") And (strProcName = "CapArray_Load") Or _
                                    ((strFormName = "frmCheckPOSPay") And (strProcName = "Sub1_Load")) Or _
                                    (strFormName = "frmCheckPOSPay_Sub1" And (strProcName = "CtlArray1_Load" Or _
                                    strProcName = "CtlArray2_Load"))) Then
                                  ' ** OK!
8510                            Else
                                  'X 1. SUB NOT EVENT? frmLoadTimer                 'cmdClose_Click'
                                  'X 1. SUB NOT EVENT? frmMenu_Account              'CtlArray_Load'
                                  'X 1. SUB NOT EVENT? frmTransaction_Audit_Sub_ds  'ColArray_Load'
                                  'X 1. SUB NOT EVENT? frmJournal                   'cmdClose_Click'
                                  'X 1. SUB NOT EVENT? frmCheckPOSPay               'Sub1_Load'
                                  'X 1. SUB NOT EVENT? frmCheckPOSPay_Sub1          'CtlArray1_Load'
                                  'X 1. SUB NOT EVENT? frmCheckPOSPay_Sub1          'CtlArray2_Load'
                                  'X 1. SUB NOT EVENT? frmRpt_CourtReports_NY       'CapArray_Load'
                                  'X 1. SUB NOT EVENT? frmTransaction_Audit         'SubFrmGfx_Load'
                                  'X 1. SUB NOT EVENT? frmTransaction_Audit_Sub     'PrintTgls_Load'
                                  'X 1. SUB NOT EVENT? frmTransaction_Audit_Sub     'PrintTgls_Click'
                                  'X 1. SUB NOT EVENT? frmMap_Div                   'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Div_Detail            'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Int                   'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Int_Detail            'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Misc_LTCL             'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Misc_LTCL_Detail      'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Misc_STCGL            'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Misc_STCGL_Detail     'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Rec                   'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Rec_Detail            'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmMap_Split                 'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmReinvest_Dividend         'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmReinvest_Interest         'NoChar_Load'
                                  'X 1. SUB NOT EVENT? frmTaxLot                    'RecArray_Load'
                                  'X 1. SUB NOT EVENT? frmJournal_Columns_TaxLot    'RecArray_Load'
                                  'X 1. SUB NOT EVENT? frmXAdmin_Form_Graphics      'Staging_Load'
                                  'X 1. SUB NOT EVENT? frmXAdmin_Form_Graphics_Sub  'RecSel_Load'
                                  'X 1. SUB NOT EVENT? frmXAdmin_Graphics_Sub       'GfxFormat_Load'
                                  'X 1. SUB NOT EVENT? frmJournal_Columns           'StatusBar_Load'
                                  'X 1. SUB NOT EVENT? frmJournal_Columns_Sub       'StatusBar_Load'
                                  'X 1. SUB NOT EVENT? frmXAdmin_Misc_Pref_Sub      'cmdDelete_Click'
8520                              lngProcsNotFound = lngProcsNotFound + 1&
8530                              Debug.Print "'1. SUB NOT EVENT? " & IIf(lngW = 1&, strFormName, strReportName) & " '" & strProcName & "'"
8540                            End If
8550                          End If
8560                        End Select  ' ** Case control type: strCtl.
8570                      Else
                            ' ** Not a recognized event.
8580                        If strProcName = "cmdSave_Click" Then
                              ' ** I don't think any of these are connected to a button.
8590                        ElseIf ((strFormName = "frmJournal" Or strFormName = "frmMenu_Post") And strEvent = "OptLoad") Or _
                                (strFormName = "frmAccountHideTrans_Hidden" And strEvent = "cmdRegen_Click") Or _
                                (strFormName = "frmXAdmin_Query_Sub" And strEvent = "cmdDelete_Click") Then
                              ' ** I know about these.
8600                        ElseIf (((strFormName = "frmRpt_AssetHistory" Or strFormName = "frmRpt_CourtReports_FL") And _
                                (strProcName = "GetBal_Beg" Or strProcName = "GetBal_End")) Or _
                                ((strFormName = "frmRpt_TransactionsByType") And (strProcName = "JType_Chk"))) Then
                              ' ** I know about these, too.
8610                        ElseIf ((strFormName = "frmRpt_CourtReports_NS" Or strFormName = "frmRpt_CourtReports_NY") And _
                                ((strProcName = "AssetList_PreviewPrint" Or strEvent = "PreviewPrint") Or _
                                (strProcName = "AssetList_Word" Or strEvent = "Word") Or _
                                (strProcName = "AssetList_Excel" Or strEvent = "Excel"))) Then
                              ' ** And these.
8620                        ElseIf (strFormName = "frmTaxLot" And strProcName = "DoMultiLots_Sort") Or _
                                (strFormName = "frmTaxLot_Sub" And (strProcName = "LocUpdate_TL" Or strEvent = "TL" Or _
                                strProcName = "SortNow_Set" Or strEvent = "Set")) Or _
                                (strFormName = "frmVersion_Main" And strProcName = "ConversionCheck_Response") Then
                              ' ** And these.
8630                        ElseIf ((strFormName = "frmTransaction_Audit" Or strFormName = "frmTransaction_Audit_Sub" Or _
                                strFormName = "frmTransaction_Audit_Sub_ds") And _
                                (strProcName = "FilterRecs_Rem" Or strProcName = "FilterRecs_Clr" Or strProcName = "Print_Chk" Or _
                                strEvent = "Rem" Or strEvent = "Clr" Or strEvent = "Chk" Or strEvent = "Set" Or _
                                strProcName = "PrintTgls_Focus" Or strProcName = "PrintTgls_Mouse" Or strProcName = "PrintTgls_Move" Or _
                                strProcName = "ColArray_Width")) Then
                              ' ** And these.
8640                        ElseIf (((strFormName = "frmAssets") And (strProcName = "IAHasChanged_Set")) Or _
                                ((strFormName = "frmAssets_Sub") And (strProcName = "Form_Width_Set" Or _
                                strProcName = "IAHasChanged_Set" Or strProcName = "IsHid_Width_Set")) Or _
                                ((strFormName = "frmPortfolioModeling") And (strProcName = "RecalcTots_NotReady"))) Then
                              ' ** And these.
8650                        ElseIf ((Left(strFormName, 14) = "frmJournal_Sub") And (strProcName = "Map_NewRec")) Or _
                                ((strFormName = "frmMenu_Account") And (strProcName = "JustClose_Set")) Or _
                                (strFormName = "frmAccountExport" And (strProcName = "Tier1_Enable_AE" Or _
                                strProcName = "Tier2_Enable_AE" Or strProcName = "Tier3_Enable_AE")) Then
                              ' ** Ummm... Yah, these too.
8660                        ElseIf (strFormName = "frmJournal_Columns" And ((strProcName = "IsMaximized_Set" Or strEvent = "Set") Or _
                                (strProcName = "JrnlCol_FocusSet" Or strEvent = "OptLoad"))) Or _
                                (strFormName = "frmJournal_Columns_TaxLot" And (strProcName = "DoMultiLots_Sort" Or strEvent = "Sort")) Or _
                                (strFormName = "frmJournal_Columns_TaxLot_Sub" And (strProcName = "LocUpdate_TL" Or _
                                strProcName = "LocUpdate_JCTL" Or strEvent = "JCTL" Or strProcName = "SortNow_Set" Or strEvent = "Set")) Or _
                                (strFormName = "frmTransaction_Audit" And (strProcName = "FormFields_Sort" Or strEvent = "Sort")) Then
                              ' ** And these.
8670                        ElseIf (strFormName = "frmJournal_Columns_Sub" And ((strProcName = "PostDate_Set" Or strProcName = "StatusBar_Set") Or _
                                strEvent = "Set") Or (strProcName = "JrnlCol_Sub_Clear" Or strEvent = "Clear") Or _
                                (strProcName = "TaxLot_Form" Or strEvent = "Form") Or _
                                (strProcName = "AddRec_Send" Or strProcName = "DelRec_Send")) Then
                              ' ** And these.
8680                        ElseIf ((strFormName = "frmXAdmin_ExportTbl_Sub") And (strProcName = "AllowOpen_Set")) Or _
                                ((strFormName = "frmXAdmin_Misc") And (strProcName = "PrefSub_Set")) Or _
                                ((strFormName = "frmCheckReconcile") And (strProcName = "GetBal_End" Or _
                                strProcName = "AccountsRowSource_Set" Or strProcName = "CheckingType_Set")) Or _
                                ((strFormName = "frmPortfolioModeling_Sub") And (strProcName = "RecalcTots_Set")) Then
                              ' ** And these.
8690                        ElseIf ((strFormName = "frmAccountProfile" Or strFormName = "frmAccountProfile_Sub" Or _
                                strFormName = "frmAccountProfile_Add" Or strFormName = "frmAccountProfile_Add_Sub") And _
                                (strProcName = "AcctChanged_Set" Or strProcName = "EscPressed_Set" Or strProcName = "HasSaved_Set" Or _
                                strProcName = "SysAcct_Set" Or strProcName = "Exit_Set" Or strProcName = "NewRec_Get" Or _
                                strProcName = "RelAccts_Set" Or strProcName = "ViewOnly_Set")) Then
                              ' ** These new ones.
8700                        ElseIf (strFormName = "frmXAdmin_Registry" And _
                                (strProcName = "Registry_tvw_NodeClick" Or strEvent = "NodeClick" Or _
                                strProcName = "Registry_tvw_Expand" Or strEvent = "Expand" Or _
                                strProcName = "Registry_tvw_Collapse" Or strEvent = "Collapse" Or _
                                strProcName = "RegTreeView_Load" Or strProcName = "RegTreeView_Get" Or _
                                strProcName = "RegTreeView_Set" Or strProcName = "RegTreeView_Icon")) Then
                              ' ** TreeView control events.
8710                        ElseIf (((strFormName = "frmStatementParameters") And (strProcName = "AcctsSchedRpt_Set")) Or _
                                ((strFormName = "frmXAdmin_Form_Graphics") And (strProcName = "CheckIDs_Set")) Or _
                                ((strFormName = "frmXAdmin_Form_Graphics_Sub") And (strProcName = "RecSel_Focus")) Or _
                                ((strFormName = "frmMasterBalance") And (strProcName = "ShowAcctMast_Win" Or strProcName = "ShowAcctSort_Win"))) Then
                              ' ** New ones.
8720                        ElseIf (((strFormName = "frmAccountContacts") And (strProcName = "AcctNoShort_Set" Or _
                                strProcName = "EnableCountry_SetFrmWidth" Or strProcName = "AcctNoShort_Move")) Or _
                                ((strFormName = "frmAccountProfile_RelAccts_Sub") And (strProcName = "CRelArray_Set" Or _
                                strProcName = "ORelArray_Set" Or strProcName = "RelArray_Set"))) Then
                              ' ** Newer ones.
8730                        ElseIf (((strFormName = "frmRpt_TaxLot") And (strProcName = "ChkGroup_Default" Or _
                                strProcName = "ChkGroup_Select" Or strProcName = "ChkGroup_Show")) Or _
                                ((strFormName = "frmCheckPOSPay_Sub1") And (strProcName = "Sub1_Disable")) Or _
                                ((strFormName = "frmCheckPOSPay_Sub2") And (strProcName = "Sub2_Disable")) Or _
                                ((strFormName = "frmCheckPOSPay_Sub3") And (strProcName = "Sub3_Disable"))) Then
                              ' ** Newest ones.
8740                        ElseIf ((strFormName = "frmAccruedIncome") And (strProcName = "HideIncExp_Purch" Or _
                                strProcName = "HideIncExp_Int")) Or ((strFormName = "frmFeeCalculations") And _
                                (strProcName = "HideIncExp_Rec" Or strProcName = "HideIncExp_Paid")) Or _
                                (strFormName = "frmLocation_Asset_Sub" And strProcName = "RecalcTots_Set") Then
                              ' ** Newester ones.
8750                        ElseIf (((strFormName = "frmLocations") And (strProcName = "EnableCountry_SetFrmWidth")) Or _
                                ((strFormName = "frmCurrency" Or strFormName = "frmCurrency_Rate") And _
                                (strProcName = "ResetFilter_Set")) Or ((strFormName = "frmAssetPricing_Import") And _
                                (strProcName = "IncludeCurrency_Update"))) Then
                              ' ** Newestest ones.
8760                        ElseIf (strProcName = "GTREmblem_Set" Or strProcName = "GTREmblem_Off" Or _
                                strProcName = "GTREmblem_Move" Or strProcName = "GTRSite_Off") Or _
                                (strFormName = "frmRecurringItems" And strProcName = "EnableCountry_SetFrmWidth") Or _
                                (strFormName = "frmRpt_TaxLot" And (strProcName = "AcctCtls_Move" Or strProcName = "cmdReset_Check")) Then
                              ' ** Newestest ones.
8770                        ElseIf (strFormName = "frmAccountHideTrans2" And strProcName = "UpdateTotals_Set") Or _
                                (strFormName = "frmAssets_Sub" And strProcName = "ParForm_Width_Set") Or _
                                ((strFormName = "frmAssetPricing" Or strFormName = "frmAssetPricing_Sub") And _
                                (strProcName = "IsOpen_Set" Or strProcName = "ForExTots_Set")) Then
                              ' ** OK.
8780                        ElseIf ((strProcName = "IncludeCurrency_Sub" Or strProcName = "ShowJournalNo_Sub" Or _
                                strProcName = "ShowAssetNo_Sub" Or strProcName = "ShowAssetTypeDesc_Sub" Or _
                                strProcName = "ShowCost_Sub" Or strProcName = "HasForEx_Set" Or _
                                strProcName = "FirstDate_Set" Or strProcName = "FirstDateMsg_Set") Or _
                                ((strFormName = "frmMenu_Account_Sub_List") And (strProcName = "SubResize_List")) Or ((strFormName = "frmMenu_Account_Sub_One") And _
                                (strProcName = "SubResize_One")) Or ((strFormName = "frmTransaction_Audit_Sub_Criteria") And (strProcName = "Calendar_Handler"))) Then
                              ' ** A-OK!
8790                        ElseIf ((strFormName = "frmRpt_CourtReports_CA") And (strProcName = "SendToFile_CA")) Or _
                                ((strFormName = "frmRpt_CourtReports_FL") And (strProcName = "SendToFile_FL")) Or _
                                ((strFormName = "frmRpt_CourtReports_NS") And (strProcName = "SendToFile_NS" Or _
                                strProcName = "AssetList_Excel_NS")) Or ((strFormName = "frmRpt_CourtReports_NY") And _
                                (strProcName = "AssetList_Excel_NY")) Then
                              ' ** Fine.
8800                        ElseIf (((strFormName = "frmAccountProfile_ReviewFreq" Or strFormName = "frmAccountProfile_ReviewFreq_Sub" Or _
                                strFormName = "frmAccountProfile_StatementFreq" Or strFormName = "frmAccountProfile_StatementFreq_Sub") And _
                                (strProcName = "MonthVals_Chk" Or strProcName = "MonthVals_Set")) Or (strFormName = "frmAccountHideTrans2" And _
                                strProcName = "shareface_chk") Or (strFormName = "frmPortfolioModeling_Sub" And _
                                strProcName = "FromEnter_Set")) Then
                              ' ** Fine.
8810                        ElseIf ((strFormName = "frmAdminOfficer" Or strFormName = "frmInvestmentObjective") And _
                                strProcName = "SortNow_Set") Or (strFormName = "frmAccountContacts" And strProcName = "Sub_View") Or _
                                (strFormName = "frmTransaction_Audit_Sub" And (strProcName = "ShowFields_Sub" Or strProcName = "ChkArray_Pop")) Or _
                                (strFormName = "frmTransaction_Audit_Sub_Criteria" And (strProcName = "FilterRec_GetArr" Or _
                                strProcName = "FilterRecs_Set")) Then
                              ' ** Equally fine.
8820                        ElseIf (strFormName = "frmXAdmin_Shortcut_Sub" And (strProcName = "AfterUpdate_Set" Or _
                                strProcName = "ShowFormName_Sub")) Or (strEvent = "NS") Or _
                                (strFormName = "frmAssetPricing_Import" And strProcName = "ProgBar_Width_Pric") Or _
                                (strFormName = "frmLinkData" And strProcName = "ProgBar_Width_Link") Or _
                                (strFormName = "frmMenu_Maintenance" And strProcName = "ProgBar_Width_Maint") Or _
                                (strFormName = "frmRpt_CourtReports_CA" And strProcName = "AllCancelSet1_CA") Or _
                                (strFormName = "frmRpt_CourtReports_FL" And strProcName = "AllCancelSet1_FL") Or _
                                (strFormName = "frmRpt_CourtReports_NS" And strProcName = "AllCancelSet1_NS") Or _
                                (strFormName = "frmRpt_CourtReports_NY" And strProcName = "AllCancelSet1_NY") Then
                              ' ** Unimaginably fine.
8830                        Else
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub                   'ChkArray_Pop'
                              'X 2. SUB NOT EVENT? frmMasterBalance                           'ShowAcctMast_Win'
                              'X 2. SUB NOT EVENT? frmMasterBalance                           'ShowAcctSort_Win'
                              'X 2. SUB NOT EVENT? frmMenu_Account_Sub_List                   'SubResize_List'
                              'X 2. SUB NOT EVENT? frmMenu_Account_Sub_One                    'SubResize_One'
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_Criteria          'Calendar_Handler'
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_ds 'Width'        'ColArray_Width
                              'X 2. SUB NOT EVENT? frmRpt_TaxLot                              'cmdReset_Check
                              'X 2. SUB NOT EVENT? frmJournal_Columns_Sub 'TL'                '
                              'X 2. SUB NOT EVENT? frmAssetPricing_Import 'Pric'              'ProgBar_Width_Pric
                              'X 2. SUB NOT EVENT? frmLinkData 'Link'                         'ProgBar_Width_Link
                              'X 2. SUB NOT EVENT? frmMenu_Maintenance 'Maint'                'ProgBar_Width_Maint
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_CA 'CA'                'AllCancelSet1_CA
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_FL 'FL'                'AllCancelSet1_FL
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NY 'NY'                'AllCancelSet1_NY
                              'X 2. SUB NOT EVENT? frmXAdmin_Shortcut_Sub                     'AfterUpdate_Set'
                              'X 2. SUB NOT EVENT? frmXAdmin_Shortcut_Sub                     'ShowFormName_Sub'
                              'X 2. SUB NOT EVENT? frmCheckPOSPay_Sub1                        'Sub1_Disable'
                              'X 2. SUB NOT EVENT? frmCheckPOSPay_Sub2                        'Sub2_Disable'
                              'X 2. SUB NOT EVENT? frmCheckPOSPay_Sub3                        'Sub3_Disable'
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub 'Sub'             'ShowFields_Sub
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_Criteria 'GetArr' 'FilterRec_GetArr
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_Criteria 'Set'    'FilterRecs_Set
                              'X 2. SUB NOT EVENT? frmAccountContacts 'View'                  'Sub_View
                              'X 2. SUB NOT EVENT? frmAdminOfficer                            'SortNow_Set
                              'X 2. SUB NOT EVENT? frmInvestmentObjective                     'SortNow_Set
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2 'chk'                 'shareface_chk
                              'X 2. SUB NOT EVENT? frmAccountProfile_ReviewFreq 'Chk'         'MonthVals_Chk
                              'X 2. SUB NOT EVENT? frmAccountProfile_ReviewFreq_Sub 'Set'     'MonthVals_Set
                              'X 2. SUB NOT EVENT? frmAccountProfile_StatementFreq 'Chk'      'MonthVals_Chk
                              'X 2. SUB NOT EVENT? frmAccountProfile_StatementFreq_Sub 'Set'  'MonthVals_Set
                              'X 2. SUB NOT EVENT? frmPortfolioModeling_Sub 'Set'             'FromEnter_Set
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_CA 'CA'                'SendToFile_CA
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_FL 'FL'                'SendToFile_FL
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NS 'NS'                'SendToFile_NS
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NS 'NS'                'AssetList_Excel_NS
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NY 'NY'                'AssetList_Excel_NY
                              'X 2. SUB NOT EVENT? frmAssetPricing_Import 'Update'            'IncludeCurrency_Update
                              'X 2. SUB NOT EVENT? frmAccountContacts 'Move'                  'AcctNoShort_Move
                              'X 2. SUB NOT EVENT? frmAssetPricing 'Set'                      'ForExTots_Set
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub 'Focus'           'PrintTgls_Focus
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub 'Mouse'           'PrintTgls_Mouse
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub 'Move'            'PrintTgls_Move
                              'X 2. SUB NOT EVENT? frmJournal_Columns_Sub 'Send'              'AddRec_Send
                              'X 2. SUB NOT EVENT? frmJournal_Columns_Sub 'Send'              'DelRec_Send
                              'X 2. SUB NOT EVENT? frmAccountAssets_Sub 'Sub'                 'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAccountComments_Sub 'Sub'               'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_Sub_List 'Sub'        'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_Sub_Pick 'Sub'        'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_One_Sub_List 'Sub'    'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_One_Sub_Pick 'Sub'    'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAccountIncExpCodes_Sub 'Sub'            'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAccountTaxCodes_Sub 'Sub'               'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAccountTransactions_Sub 'Sub'           'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAssetPricing_Sub 'Sub'                  'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAssetPricing_History_Sub 'Sub'          'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmTaxLot_Sub 'Sub'                        'IncludeCurrency_Sub
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_One_Sub_List 'Sub'    'ShowJournalNo_Sub
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_One_Sub_Pick 'Sub'    'ShowJournalNo_Sub
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_Sub_List 'Sub'        'ShowJournalNo_Sub
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_Sub_Pick 'Sub'        'ShowJournalNo_Sub
                              'X 2. SUB NOT EVENT? frmAssetPricing_History_Sub 'Sub'          'ShowAssetTypeDesc_Sub
                              'X 2. SUB NOT EVENT? frmAssetPricing_History_Sub 'Sub'          'ShowAssetNo_Sub
                              'X 2. SUB NOT EVENT? frmAccountIncExpCodes_Sub 'Sub'            'ShowCost_Sub
                              'X 2. SUB NOT EVENT? frmAccountTaxCodes_Sub 'Sub'               'ShowCost_Sub
                              'X 2. SUB NOT EVENT? frmAccountAssets_Sub 'Set'                 'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_One_Sub_List 'Set'    'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_One_Sub_Pick 'Set'    'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_Sub_List 'Set'        'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2_Sub_Pick 'Set'        'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmAccountIncExpCodes_Sub 'Set'            'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmAccountTaxCodes_Sub 'Set'               'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmAssetPricing_History_Sub 'Set'          'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmAssetPricing_Sub 'Set'                  'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmAssets_Sub 'Set'                        'HasForEx_Set
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NY 'Set'               'FirstDate_Set
                              'X 2. SUB NOT EVENT? frmStatementParameters 'Set'               'FirstDate_Set
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NY 'Set'               'FirstDateMsg_Set
                              'X 2. SUB NOT EVENT? frmStatementParameters 'Set'               'FirstDateMsg_Set
                              'X 2. SUB NOT EVENT? frmAccountHideTrans2 'Set'                 'UpdateTotals_Set
                              'X 2. SUB NOT EVENT? frmAssetPricing_Sub 'Set'                  'IsOpen_Set
                              'X 2. SUB NOT EVENT? frmAssets_Sub 'Set'                        'ParForm_Width_Set
                              'X 2. SUB NOT EVENT? frmAccountContacts 'SetFrmWidth'           'EnableCountry_SetFrmWidth
                              'X 2. SUB NOT EVENT? frmAccountProfile_Sub 'Set'                'ViewOnly_Set
                              'X 2. SUB NOT EVENT? frmCurrency 'Set'                          'ResetFilter_Set
                              'X 2. SUB NOT EVENT? frmCurrency_Rate 'Set'                     'ResetFilter_Set
                              'X 2. SUB NOT EVENT? frmLocations 'SetFrmWidth'                 'EnableCountry_SetFrmWidth
                              'X 2. SUB NOT EVENT? frmRecurringItems 'SetFrmWidth'            'EnableCountry_SetFrmWidth
                              'X 2. SUB NOT EVENT? frmRpt_TaxLot 'Move'                       'AcctCtls_Move
                              'X 2. SUB NOT EVENT? frmLocation_Asset_Sub 'Set'                'RecalcTots_Set
                              'X 2. SUB NOT EVENT? frmAccountHideTrans 'Off'                  'GTREmblem_Off
                              'X 2. SUB NOT EVENT? frmAccountProfile 'Off'                    'GTREmblem_Off
                              'X 2. SUB NOT EVENT? frmJournal_Columns 'Move'                  'GTREmblem_Move
                              'X 2. SUB NOT EVENT? frmJournal_Columns 'Off'                   'GTREmblem_Off
                              'X 2. SUB NOT EVENT? frmStatementParameters 'Off'               'GTREmblem_Off
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NY                     'PreviewPrint'
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NY                     'Word'
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NY                     'Excel'
                              'X 2. SUB NOT EVENT? frmAccountHideTrans                        'Set'
                              'X 2. SUB NOT EVENT? frmAccountProfile                          'Set'
                              'X 2. SUB NOT EVENT? frmJournal                                 'Set'
                              'X 2. SUB NOT EVENT? frmAccruedIncome 'Purch'                   'HideIncExp_Purch
                              'X 2. SUB NOT EVENT? frmAccruedIncome 'Int'                     'HideIncExp_Int
                              'X 2. SUB NOT EVENT? frmFeeCalculations 'Rec'                   'HideIncExp_Rec
                              'X 2. SUB NOT EVENT? frmFeeCalculations 'Paid'                  'HideIncExp_Paid
                              'X 2. SUB NOT EVENT? frmRpt_TaxLot 'Default'                    'ChkGroup_Default
                              'X 2. SUB NOT EVENT? frmRpt_TaxLot 'Select'                     'ChkGroup_Select
                              'X 2. SUB NOT EVENT? frmRpt_TaxLot 'Show'                       'ChkGroup_Show
                              'X 2. SUB NOT EVENT? frmAccountContacts 'Set'                   'AcctNoShort_Set
                              'X 2. SUB NOT EVENT? frmAccountProfile 'Set'                    'RelAccts_Set
                              'X 2. SUB NOT EVENT? frmAccountProfile_RelAccts_Sub 'Set'       'CRelArray_Set
                              'X 2. SUB NOT EVENT? frmAccountProfile_RelAccts_Sub 'Set'       'ORelArray_Set
                              'X 2. SUB NOT EVENT? frmAccountProfile_RelAccts_Sub 'Set'       'RelArray_Set
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_ds 'Rem'          'FilterRecs_Rem
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_ds 'Clr'          'FilterRecs_Clr
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_ds 'Chk'          'Print_Chk
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_ds 'Set'          'FilterRecs_Set
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_ds 'Set'          'SortNow_Set
                              'X 2. SUB NOT EVENT? frmTransaction_Audit_Sub_ds 'Set'          'TotalRecs_Set
                              'X 2. SUB NOT EVENT? frmStatementParameters 'Set'               'AcctsSchedRpt_Set()
                              'X 2. SUB NOT EVENT? frmXAdmin_Form_Graphics 'Set'              'CheckIDs_Set()
                              'X 2. SUB NOT EVENT? frmXAdmin_Form_Graphics_Sub 'Focus'        'RecSel_Focus()
                              'X 2. SUB NOT EVENT? frmCheckReconcile 'Set'                    'AccountsRowSource_Set()
                              'X 2. SUB NOT EVENT? frmCheckReconcile 'Set'                    'CheckingType_Set()
                              'X 2. SUB NOT EVENT? frmRpt_TransactionsByType 'Chk'            'JType_Chk()
                              'X 2. SUB NOT EVENT? frmJournal_Sub1_Dividend 'NewRec'          'Map_NewRec()
                              'X 2. SUB NOT EVENT? frmJournal_Sub2_Interest 'NewRec'          'Map_NewRec()
                              'X 2. SUB NOT EVENT? frmJournal_Sub3_Purchase 'NewRec'          'Map_NewRec()
                              'X 2. SUB NOT EVENT? frmJournal_Sub5_Misc 'NewRec'              'Map_NewRec()
                              'X 2. SUB NOT EVENT? frmXAdmin_Misc 'Set'                       'PrefSub_Set()
                              'X 2. SUB NOT EVENT? frmXAdmin_Registry 'Icon'                  'RegTreeView_Icon()
                              'X 2. SUB NOT EVENT? frmJournal_Columns 'FocusSet'              'JrnlCol_FocusSet()
                              'X 2. SUB NOT EVENT? frmMenu_Post 'OptLoad'                     'Journal_OptLoad()
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NS 'PreviewPrint'      'AssetList_PreviewPrint()
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NS 'Word'              'AssetList_Word()
                              'X 2. SUB NOT EVENT? frmRpt_CourtReports_NS 'Excel'             'AssetList_Excel()
                              'X 2. SUB NOT EVENT? frmAccountProfile 'Set'                    'AcctChanged_Set()
                              'X 2. SUB NOT EVENT? frmAccountProfile 'Set'                    'EscPressed_Set()
                              'X 2. SUB NOT EVENT? frmAccountProfile 'Set'                    'HasSaved_Set()
                              'X 2. SUB NOT EVENT? frmAccountProfile 'Set'                    'SysAcct_Set()
                              'X 2. SUB NOT EVENT? frmAccountProfile_Add 'Set'                'EscPressed_Set()
                              'X 2. SUB NOT EVENT? frmAccountProfile_Add_Sub 'Set'            'EscPressed_Set()
                              'X 2. SUB NOT EVENT? frmAccountProfile_Sub 'Set'                'EscPressed_Set()
                              'X 2. SUB NOT EVENT? frmAccountProfile_Sub 'Set'                'Exit_Set()
                              'X 2. SUB NOT EVENT? frmAccountProfile_Sub 'Get'                'NewRec_Get()
                              'X 2. SUB NOT EVENT? frmAccountProfile_Sub 'Set'                'SysAcct_Set()
                              'X 2. SUB NOT EVENT? frmPortfolioModeling 'NotReady'            'RecalcTots_NotReady()
                              'X 2. SUB NOT EVENT? frmCheckReconcile 'End'                    'GetBal_End()
                              'X 2. SUB NOT EVENT? frmJournal_Columns_Sub 'Form'              'TaxLot_Form()
                              'X 2. SUB NOT EVENT? frmPortfolioModeling_Sub 'Set'             'RecalcTots_Set()
                              'X 2. SUB NOT EVENT? frmJournal_Columns_TaxLot_Sub 'JCTL'       'LocUpdate_JCTL
                              'X 2. SUB NOT EVENT? frmJournal_Columns_TaxLot_Sub 'Set'        'SortNow_Set
                              'X 2. SUB NOT EVENT? frmTaxLot_Sub 'TL'                         'LocUpdate_TL
                              'X 2. SUB NOT EVENT? frmTaxLot_Sub 'Set'                        'SortNow_Set
                              'X 2. SUB NOT EVENT? frmJournal_Columns_Sub 'Set'               'PostDate_Set
                              'X 2. SUB NOT EVENT? frmJournal_Columns_Sub 'Set'               'StatusBar_Set
                              'X 2. SUB NOT EVENT? frmJournal_Columns_Sub 'Clear'             'JrnlCol_Sub_Clear
8840                          lngProcsNotFound = lngProcsNotFound + 1&
8850                          Debug.Print "'2. SUB NOT EVENT? " & IIf(lngW = 1&, strFormName, strReportName) & " '" & strEvent & "'"
8860                        End If
8870                      End If  ' ** It's a known property event: strProc <> vbNullString.

8880                    End If  ' ** The procedure has a name.
8890                  End If  ' ** It's a procedure in this form's module.
8900                End If  ' ** It's a potential event: P_XEVNT = True.
8910              Next  ' ** For each procedure: lngY, lngElemP.
8920            End With  ' ** This form or report: obj.

8930            If blnChanged = True Then
8940              Select Case lngW
                  Case 1&  ' ** Form.
8950                arr_varFrm(F_CHGS, lngElemF) = lngThisFormChanges
8960                DoCmd.Close acForm, strFormName, acSaveYes
8970              Case 2&  ' ** Report.
8980                arr_varRpt(R_CHGS, lngElemR) = lngThisReportChanges
8990                DoCmd.Close acReport, strReportName, acSaveYes
9000              End Select
9010            Else
9020              Select Case lngW
                  Case 1&  ' ** Form.
9030                DoCmd.Close acForm, strFormName, acSaveNo
9040              Case 2&  ' ** Report.
9050                DoCmd.Close acReport, strReportName, acSaveNo
9060              End Select
9070            End If

9080          End If  ' ** A form or report with a module (many don't).
9090        Next  ' ** For each form or report: lngX, lngElemF or lngElemR.
9100      Next  ' ** For Forms or Reports: lngW.

9110      For lngX = 0& To (lngFrms - 1&)
9120        lngElemF = lngX
9130        If arr_varFrm(F_CHGS, lngX) > 0& Then
9140          Debug.Print "'FRM PROC ADDED: " & arr_varFrm(F_FNAM, lngElemF)
9150          For lngY = 0& To (lngChanges - 1&)
9160            lngElemC = lngY
9170            If arr_varChange(C_FELEM, lngElemC) = lngElemF Then
9180              Debug.Print "'  " & arr_varChange(C_PROC, lngElemC)
9190            End If
9200          Next
9210        End If
9220      Next
9230      For lngX = 0& To (lngRpts - 1&)
9240        lngElemR = lngX
9250        If arr_varRpt(R_CHGS, lngX) > 0& Then
9260          Debug.Print "'RPT PROC ADDED: " & arr_varRpt(R_RNAM, lngElemF)
9270          For lngY = 0& To (lngChanges - 1&)
9280            lngElemC = lngY
9290            If arr_varChange(C_RELEM, lngElemC) = lngElemR Then
9300              Debug.Print "'  " & arr_varChange(C_PROC, lngElemC)
9310            End If
9320          Next
9330        End If
9340      Next

9350    End If  ' ** lngProcs > 0.

9360    If lngProcsNotFound = 0& Then
9370      Debug.Print "'NONE FOUND!"
9380    Else
9390      Debug.Print "'DONE"
9400    End If

9410    Beep

EXITP:
9420    Set Sec = Nothing
9430    Set frm = Nothing
9440    Set frm_ao = Nothing
9450    Set rpt = Nothing
9460    Set rpt_ao = Nothing
9470    Set ctl = Nothing
9480    Set cod = Nothing
9490    Set vbc = Nothing
9500    Set vbp = Nothing
9510    VBA_Chk_Events = blnRetVal
9520    Exit Function

ERRH:
9530    blnRetVal = False
9540    Select Case ERR.Number
        Case Else
9550      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
9560    End Select
9570    Resume EXITP

End Function

Public Function VBA_Chk_LineNums() As Boolean
' ** Check code line numbers, looking for un-numbered
' ** code lines, numbered remarks, and numbered blanks.

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "VBA_Chk_LineNums"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim strModName As String, strLine As String, strProcName As String
        Dim lngLines As Long, lngDecLines As Long
        Dim lngNums As Long, arr_varNum() As Variant
        Dim blnFound As Boolean
        Dim intPos01 As Integer
        Dim lngTmp01 As Long
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varNum().
        Const N_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const N_MOD As Integer = 0
        Const N_NUM As Integer = 1
        Const N_TYP As Integer = 2
        Const N_LIN As Integer = 3

9610    blnRetVal = True

9620    lngNums = 0&
9630    ReDim arr_varNum(N_ELEMS, 0)

9640    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

        ' ** 1441 in DoMultiLots_Round(), frmLotInformation, is 1st Case of Select Case.

        ' ** Walk through every module.
9650    Set vbp = Application.VBE.ActiveVBProject
9660    With vbp
9670      For Each vbc In .VBComponents
9680        With vbc
9690          strModName = .Name
9700          If Left(strModName, 7) <> "zz_mod_" And Left(strModName, 11) <> "Form_zz_frm" Then
9710            Set cod = .CodeModule
9720            With cod
9730              lngLines = .CountOfLines
9740              lngDecLines = .CountOfDeclarationLines
9750              If lngLines > 0& And lngDecLines > 0& And lngDecLines <> lngLines Then
9760                For lngX = (lngDecLines + 1&) To lngLines
9770                  strLine = .Lines(lngX, 1)
9780                  strProcName = vbNullString
9790                  If Trim(strLine) <> vbNullString Then
9800                    If Left(Trim(strLine), 1) = "'" Then
                          ' ** It's a remark. Skip.
9810                    ElseIf Left(Trim(strLine), 1) = "#" Then
                          ' ** It's a compiler directive. Skip.
9820                    Else
9830                      intPos01 = InStr(strLine, " ")
9840                      If intPos01 = 0 Then
                            ' ** One word line.
9850                        If Right(strLine, 1) = ":" Then
                              ' ** It's a label. Skip.
9860                        Else
9870                          If IsNumeric(Trim(strLine)) = True Then
                                ' ** A line number alone!
                                'Debug.Print "'NUM ALONE: " & CStr(lngX) & "  MOD: " & strModName & "  ~" & strLine & "~"
9880                            lngNums = lngNums + 1&
9890                            lngE = lngNums - 1&
9900                            ReDim Preserve arr_varNum(N_ELEMS, lngE)
9910                            arr_varNum(N_MOD, lngE) = strModName
9920                            arr_varNum(N_NUM, lngE) = lngX
9930                            arr_varNum(N_TYP, lngE) = "NUM ALONE"
9940                            arr_varNum(N_LIN, lngE) = strLine
9950                          Else
9960                            If Right(Trim(.Lines(lngX - 1&, 1)), 1) = "_" Then
                                  ' ** It's a line continuation. Skip.
9970                            Else
                                  'Debug.Print "'NO LINE NUM: " & CStr(lngX) & "  MOD: " & strModName & "  ~" & strLine & "~"
9980                              lngNums = lngNums + 1&
9990                              lngE = lngNums - 1&
10000                             ReDim Preserve arr_varNum(N_ELEMS, lngE)
10010                             arr_varNum(N_MOD, lngE) = strModName
10020                             arr_varNum(N_NUM, lngE) = lngX
10030                             arr_varNum(N_TYP, lngE) = "NO LINE NUM"
10040                             arr_varNum(N_LIN, lngE) = strLine
10050                           End If
10060                         End If
10070                       End If
10080                     Else
                            ' ** More than one word.
10090                       If Right(Trim(.Lines(lngX - 1&, 1)), 1) = "_" Then
                              ' ** It's a line continuation. Skip.
10100                       Else
10110                         If IsNumeric(Trim(Left(strLine, intPos01))) = True Then
10120                           If Left(Trim(Mid(strLine, intPos01)), 1) = "'" Then
10130                             strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
10140                             If strModName = "modQueryFunctions1" And strProcName = "Qry_CheckBox" Then
                                    ' ** These will come and go.
10150                             Else
                                    ' ** A numbered remark!
                                    'Debug.Print "'NUM REM: " & CStr(lngX) & "  MOD: " & strModName & "  ~" & strLine & "~"
10160                               lngNums = lngNums + 1&
10170                               lngE = lngNums - 1&
10180                               ReDim Preserve arr_varNum(N_ELEMS, lngE)
10190                               arr_varNum(N_MOD, lngE) = strModName
10200                               arr_varNum(N_NUM, lngE) = lngX
10210                               arr_varNum(N_TYP, lngE) = "NUM REM"
10220                               arr_varNum(N_LIN, lngE) = strLine
10230                             End If
10240                           Else
                                  ' ** It's got a line number. Skip.
10250                           End If
10260                         Else
                                ' ** Now check for start and end of procedure, Dim's, Const's, Static's, etc.
10270                           If Left(Trim(strLine), 7) = "End Sub" Or _
                                    Left(Trim(strLine), 12) = "End Function" Or _
                                    Left(Trim(strLine), 12) = "End Property" Then
                                  ' ** Skip.
10280                           ElseIf Left(Trim(strLine), 6) = "Public" Or _
                                    Left(Trim(strLine), 7) = "Private" Or _
                                    Left(Trim(strLine), 8) = "Function" Or _
                                    Left(Trim(strLine), 3) = "Sub" Or _
                                    Left(Trim(strLine), 8) = "Property" Then
                                  ' ** Skip.
10290                           ElseIf Left(Trim(strLine), 4) = "Dim " Or _
                                    Left(Trim(strLine), 6) = "Const " Or _
                                    Left(Trim(strLine), 7) = "Static " Then
                                  ' ** Skip.
10300                           ElseIf Left(Trim(strLine), 5) = "Case " Then
10310                             If Trim(.Lines(lngX - 1, 1)) <> vbNullString Then
10320                               If Left(Trim(.Lines(lngX - 1, 1)), 1) <> "'" Then
10330                                 If InStr(.Lines(lngX - 1, 1), "Select Case") > 0 Then
                                        ' ** First Case in Select block. Skip.
10340                                 Else
                                        'Debug.Print "'NO LINE NUM: " & CStr(lngX) & "  MOD: " & strModName & "  ~" & strLine & "~"
10350                                   lngNums = lngNums + 1&
10360                                   lngE = lngNums - 1&
10370                                   ReDim Preserve arr_varNum(N_ELEMS, lngE)
10380                                   arr_varNum(N_MOD, lngE) = strModName
10390                                   arr_varNum(N_NUM, lngE) = lngX
10400                                   arr_varNum(N_TYP, lngE) = "NO LINE NUM"
10410                                   arr_varNum(N_LIN, lngE) = strLine
10420                                 End If
10430                               Else
                                      ' ** Previous line was a remark.
10440                                 If InStr(.Lines(lngX - 2, 1), "Select Case") > 0 Then
                                        ' ** First Case in Select block. Skip.
10450                                 Else
                                        'Debug.Print "'NO LINE NUM: " & CStr(lngX) & "  MOD: " & strModName & "  ~" & strLine & "~"
10460                                   lngNums = lngNums + 1&
10470                                   lngE = lngNums - 1&
10480                                   ReDim Preserve arr_varNum(N_ELEMS, lngE)
10490                                   arr_varNum(N_MOD, lngE) = strModName
10500                                   arr_varNum(N_NUM, lngE) = lngX
10510                                   arr_varNum(N_TYP, lngE) = "NO LINE NUM"
10520                                   arr_varNum(N_LIN, lngE) = strLine
10530                                 End If
10540                               End If
10550                             Else
                                    ' ** Previous line was blank.
10560                               If InStr(.Lines(lngX - 2, 1), "Select Case") > 0 Then
                                      ' ** First Case in Select block. Skip.
10570                               Else
                                      'Debug.Print "'NO LINE NUM: " & CStr(lngX) & "  MOD: " & strModName & "  ~" & strLine & "~"
10580                                 lngNums = lngNums + 1&
10590                                 lngE = lngNums - 1&
10600                                 ReDim Preserve arr_varNum(N_ELEMS, lngE)
10610                                 arr_varNum(N_MOD, lngE) = strModName
10620                                 arr_varNum(N_NUM, lngE) = lngX
10630                                 arr_varNum(N_TYP, lngE) = "NO LINE NUM"
10640                                 arr_varNum(N_LIN, lngE) = strLine
10650                               End If
10660                             End If
10670                           Else
                                  ' ** What's left?
                                  'Debug.Print "'NO LINE NUM: " & CStr(lngX) & "  MOD: " & strModName & "  ~" & strLine & "~"
10680                             lngNums = lngNums + 1&
10690                             lngE = lngNums - 1&
10700                             ReDim Preserve arr_varNum(N_ELEMS, lngE)
10710                             arr_varNum(N_MOD, lngE) = strModName
10720                             arr_varNum(N_NUM, lngE) = lngX
10730                             arr_varNum(N_TYP, lngE) = "NO LINE NUM"
10740                             arr_varNum(N_LIN, lngE) = strLine
10750                           End If
10760                         End If  ' ** No line number.
10770                       End If  ' ** Not a continuation.
10780                     End If  ' ** Space: intPos01.
10790                   End If  ' ** Not a remark.
10800                 End If  ' ** Not an empty line.
10810               Next  ' ** lngX
10820             End If  ' ** Has procedures.
10830           End With  ' ** cod.
10840         End If  ' ** Not a zz_mod_.
10850       End With  ' ** vbc
10860     Next  ' ** vbc.
10870   End With  ' ** vbp.

10880   If lngNums > 0& Then
10890     blnFound = False

10900     Set dbs = CurrentDb
10910     With dbs
            ' ** Empty tblVBComponent_LineNumErrs.
10920       Set qdf = .QueryDefs("qryVBComponent_Event_02")
10930       qdf.Execute
10940       .Close
10950     End With

          ' ** Reset the Autonumber field to 1.
10960     ChangeSeed_Ext "tblVBComponent_LineNumErrs"  ' ** Module Function: modAutonumberFieldFuncs.

10970     Set dbs = CurrentDb
10980     With dbs
10990       Set rst = .OpenRecordset("tblVBComponent_LineNumErrs", dbOpenDynaset, dbConsistent)
11000       With rst
11010         lngTmp01 = lngNums
11020         For lngX = 0& To (lngNums - 1&)
                ' ** These are the first Case statement, but multiple
                ' ** remarked lines separate them from the Select Case.
                ' **   vbcom_name              vbcomline_linenum
                ' **   ======================  =================
                ' **   Form_frmLotInformation  1373
                ' **   modCalendar             1259
                ' **   modCourtReportsNS       111
                ' **   modMouseWheel           161
11030           If (arr_varNum(N_MOD, lngX) = "Form_frmLotInformation" And arr_varNum(N_NUM, lngX) = 1441&) Or _
                    (arr_varNum(N_MOD, lngX) = "modCalendar" And arr_varNum(N_NUM, lngX) = 1259&) Or _
                    (arr_varNum(N_MOD, lngX) = "modCourtReportsNS" And arr_varNum(N_NUM, lngX) = 111&) Or _
                    (arr_varNum(N_MOD, lngX) = "modMouseWheel" And arr_varNum(N_NUM, lngX) = 161&) Then
                  ' ** I know about these.
11040             lngTmp01 = lngTmp01 - 1&
11050           ElseIf (arr_varNum(N_MOD, lngX) = "modOperSysInfoFuncs1" And _
                    (arr_varNum(N_NUM, lngX) = 146& Or arr_varNum(N_NUM, lngX) = 292&)) Or _
                    (arr_varNum(N_MOD, lngX) = "modVersionDocFuncs" And _
                    (arr_varNum(N_NUM, lngX) = 32& Or arr_varNum(N_NUM, lngX) = 70& Or _
                    arr_varNum(N_NUM, lngX) = 76& Or arr_varNum(N_NUM, lngX) = 288&)) Then
                  ' ** Conditional Compilation structure can't take line numbers.
11060             lngTmp01 = lngTmp01 - 1&
11070           Else
11080             blnFound = True
11090             .AddNew
                  '![vbcom_id] =
11100             ![vbcom_name] = arr_varNum(N_MOD, lngX)
11110             ![vbcomline_linenum] = arr_varNum(N_NUM, lngX)
11120             ![vbcomline_type] = arr_varNum(N_TYP, lngX)
11130             ![vbcomline_line] = arr_varNum(N_LIN, lngX)
11140             .Update
11150           End If
11160         Next
11170         .Close
11180       End With
11190       .Close
11200     End With

11210     If blnFound = True Then
11220       Debug.Print "'LINENUM ERRS: " & CStr(lngTmp01)
11230     Else
11240       Debug.Print "'NONE FOUND!"
11250     End If
11260   Else
11270     Debug.Print "'NONE FOUND!"
11280   End If

11290   Beep

EXITP:
11300   Set cod = Nothing
11310   Set vbc = Nothing
11320   Set vbp = Nothing
11330   Set rst = Nothing
11340   Set qdf = Nothing
11350   Set dbs = Nothing
11360   VBA_Chk_LineNums = blnRetVal
11370   Exit Function

ERRH:
11380   blnRetVal = False
11390   Select Case ERR.Number
        Case Else
11400     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
11410   End Select
11420   Resume EXITP

End Function

Public Function VBA_Event_Load() As Boolean

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "VBA_Event_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

11510   blnRetVal = True

        ' ** Report events:
        'OnActivate
        'OnClose
        'OnDeactivate
        'OnError
        'OnFormat
        'OnNoData
        'OnOpen
        'OnPage
        'OnPrint
        'OnRetreat

11520   If lngEvents = 0& Then
11530     Set dbs = CurrentDb
11540     With dbs
            ' ** tblVBComponent_Event, sorted.
11550       Set qdf = .QueryDefs("qryVBComponent_Event_01")
11560       Set rst = qdf.OpenRecordset
11570       With rst
11580         .MoveLast
11590         lngEvents = .RecordCount
11600         .MoveFirst
11610         arr_varEvent = .GetRows(lngEvents)
              ' *****************************************************
              ' ** Array: arr_varEvent()
              ' **
              ' **   Field  Element  Name                Constant
              ' **   =====  =======  ==================  ==========
              ' **     1       0     vbcom_event_name    E_NAM
              ' **     2       1     vbcom_frm           E_ISFRM
              ' **     3       2     vbcom_rpt           E_ISRPT
              ' **     4       3     vbcom_ctl           E_ISCTL
              ' **
              ' *****************************************************
11620         .Close
11630       End With
11640       .Close
11650     End With
11660   End If

EXITP:
11670   Set rst = Nothing
11680   Set qdf = Nothing
11690   Set dbs = Nothing
11700   VBA_Event_Load = blnRetVal
11710   Exit Function

ERRH:
11720   blnRetVal = False
11730   Select Case ERR.Number
        Case Else
11740     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
11750   End Select
11760   Resume EXITP

End Function

Public Function List_Refs_MDE() As Boolean
' ** For more extensive reference info, see References_Doc() function in modXAdminFuncs.

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "List_Refs_MDE"

        Dim ref As Access.Reference
        Dim lngRefs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

11810   blnRetVal = True

11820   With Application
11830     lngRefs = .References.Count
11840     For lngX = 1& To lngRefs
11850       Set ref = .References(lngX)
11860       With ref
11870         Debug.Print "'" & .Name
11880       End With
11890     Next
11900   End With

EXITP:
11910   Beep
11920   Set ref = Nothing
11930   List_Refs_MDE = blnRetVal
11940   Exit Function

ERRH:
11950   blnRetVal = False
11960   Select Case ERR.Number
        Case Else
11970     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
11980   End Select
11990   Resume EXITP

End Function

Public Function Security_PermsGet() As Boolean
' ** Requires Microsoft ADO Ext. 6.0 for DDL and Security.
' ** List Permissions of specified file in specified directory.

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_PermsGet"

        'Dim cnxn As New ADODB.Connection, catx As New ADOX.Catalog  ' ** Early binding.
        Dim cnxn As Object, catx As Object                           ' ** Late binding.
        Dim lngRightsNow As Long
        Dim strMDBname As String, strMDWName As String
        Dim strUser As String, strPass As String
        Dim strObject As String
        Dim blnRetVal As Boolean

12010   blnRetVal = True

        'gstrDir_Dev As String = "C:\VictorGCS_Clients\TrustAccountant\NewWorking"  '## OK

12020   strMDBname = gstrFile_DataName
        'strMDBname = gstrFile_ArchDataName
12030   strObject = "ledger"
12040   strMDWName = gstrFile_SecurityName
12050   strUser = "superuser"
12060   strPass = TA_SEC

12070   Set cnxn = CreateObject("ADODB.Connection")  ' ** Late binding.
12080   With cnxn
12090     .Provider = "Microsoft.Jet.OLEDB.4.0"
12100     .Open "Data Source='" & gstrDir_Dev & LNK_SEP & gstrDir_DevDemo & LNK_SEP & strMDBname & "';" & "jet oledb:system database=" & _
            "'" & gstrDir_Dev & LNK_SEP & gstrDir_DevDemo & LNK_SEP & strMDWName & "'", strUser, strPass
12110   End With

12120   Set catx = CreateObject("ADOX.Catalog")
12130   catx.ActiveConnection = cnxn

12140   With catx

          ' ** Retrieve original permissions as a single long integer.
12150     lngRightsNow = catx.Users(strUser).GetPermissions(strObject, adPermObjTable)

          ' ** List the individual permissions by using ADOX constants.
12160     Debug.Print "'PERM: " & strMDBname & " " & strObject & " " & lngRightsNow
12170     Debug.Print Security_PermsDecode(lngRightsNow)  ' ** Function: Below.
          'PERM: TrustDta.mdb ledger -1072853760
          '  adRightDrop
          '  adRightReadDesign
          '  adRightWriteDesign
          '  adRightInsert
          '  adRightDelete
          '  adRightWritePermissions
          '  adRightWriteOwner
          '  adRightUpdate
          '  adRightRead
          'PERM: TrstArch.mdb ledger -1072853760
          '  adRightDrop
          '  adRightReadDesign
          '  adRightWriteDesign
          '  adRightInsert
          '  adRightDelete
          '  adRightWritePermissions
          '  adRightWriteOwner
          '  adRightUpdate
          '  adRightRead

          ' *********************************************************************************************
          ' ** GetPermissions Method(ADOX):
          ' ** Returns the permissions for a group or user on an object or object container.
          ' **
          ' ** ReturnValue=GroupOrUser.GetPermissions(Name, ObjectType [,ObjectTypeId])
          ' **
          ' ** Parameters:
          ' **   Name          A Variant value that specifies the name of the object for which to set
          ' **                 permissions. Set Name to a null value if you want to get the permissions
          ' **                 for the object container.
          ' **   ObjectType    A Long value which can be one of the ObjectTypeEnum constants, that
          ' **                 specifies the type of the object for which to get permissions.
          ' **                 AdObjectType enumeration:
          ' **                    2  adPermObjColumn            The object is a column.
          ' **                    3  adPermObjDatabase          The object is a database.
          ' **                    4  adPermObjProcedure         The object is a procedure.
          ' **                   -1  adPermObjProviderSpecific  The object is a type defined by
          ' **                                                  the provider. An error will occur
          ' **                                                  if the ObjectType parameter is
          ' **                                                  adPermObjProviderSpecific and an
          ' **                                                  ObjectTypeId is not supplied.
          ' **                    1  adPermObjTable             The object is a table.
          ' **                    5  adPermObjView              The object is a view.
          ' **   ObjectTypeId  Optional. A Variant value that specifies the GUID for a provider object
          ' **                 type not defined by the OLE DB specification. This parameter is required
          ' **                 if ObjectType is set to adPermObjProviderSpecific; otherwise, it is not
          ' **                 used.
          ' **   Return Value  Returns a Long value that specifies a bitmask containing the permissions
          ' **                 that the group or user has on the object. This value can be one or more
          ' **                 of the RightsEnum constants.
          ' *********************************************************************************************

          ' *********************************************************************************************
          ' ** SetPermissions Method(ADOX)
          ' ** Specifies the permissions for a group or user on an object.
          ' **
          ' ** GroupOrUser.SetPermissions Name, ObjectType, Action, Rights [, Inherit] [, ObjectTypeId]
          ' **
          ' ** Parameters
          ' **   Name          A String value that specifies the name of the object for which to set
          ' **                 permissions.
          ' **   ObjectType    A Long value which can be one of the ObjectTypeEnum constants, that
          ' **                 specifies the type of the object for which to get permissions.
          ' **   Action        A Long value which can be one of the ActionEnum constants that specifies
          ' **                 the type of action to perform when setting permissions.
          ' **   Rights        A Long value which can be a bitmask of one or more of the RightsEnum
          ' **                 constants, that indicates the rights to set.
          ' **   Inherit       Optional. A Long value which can be one of the InheritTypeEnum constants,
          ' **                 that specifies how objects will inherit these permissions. The default
          ' **                 value is adInheritNone.
          ' **   ObjectTypeId  Optional. A Variant value that specifies the GUID for a provider object
          ' **                 type that is not defined by the OLE DB specification. This parameter is
          ' **                 required if ObjectType is set to adPermObjProviderSpecific; otherwise,
          ' **                 it is not used.
          ' *********************************************************************************************

12180   End With

12190   cnxn.Close

EXITP:
12200   Set catx = Nothing
12210   Set cnxn = Nothing
12220   Security_PermsGet = blnRetVal
12230   Exit Function

ERRH:
12240   blnRetVal = False
12250   Select Case ERR.Number
        Case Else
12260     Beep
12270     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
12280   End Select
12290   Resume EXITP

End Function

Private Function Security_PermsDecode(lngPerm As Long) As String
' ** Requires Microsoft ADO Ext. 6.0 for DDL and Security.

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_PermsDecode"

        Dim strRetVal As String

12310   strRetVal = vbNullString

        ' ** Enumerate the permissions.
12320   If lngPerm And adRightNone Then strRetVal = strRetVal & "'  " & "adRightNone" & vbCrLf
12330   If lngPerm And adRightDrop Then strRetVal = strRetVal & "'  " & "adRightDrop" & vbCrLf
12340   If lngPerm And adRightExclusive Then strRetVal = strRetVal & "'  " & "adRightExclusive" & vbCrLf
12350   If lngPerm And adRightReadDesign Then strRetVal = strRetVal & "'  " & "adRightReadDesign" & vbCrLf
12360   If lngPerm And adRightWriteDesign Then strRetVal = strRetVal & "'  " & "adRightWriteDesign" & vbCrLf
12370   If lngPerm And adRightWithGrant Then strRetVal = strRetVal & "'  " & "adRightWithGrant" & vbCrLf
12380   If lngPerm And adRightReference Then strRetVal = strRetVal & "'  " & "adRightReference" & vbCrLf
12390   If lngPerm And adRightCreate Then strRetVal = strRetVal & "'  " & "adRightCreate" & vbCrLf
12400   If lngPerm And adRightInsert Then strRetVal = strRetVal & "'  " & "adRightInsert" & vbCrLf
12410   If lngPerm And adRightDelete Then strRetVal = strRetVal & "'  " & "adRightDelete" & vbCrLf
12420   If lngPerm And adRightReadPermissions Then strRetVal = strRetVal & "'  " & "adRightReadPermissions" & vbCrLf
12430   If lngPerm And adRightWritePermissions Then strRetVal = strRetVal & "'  " & "adRightWritePermissions" & vbCrLf
12440   If lngPerm And adRightWriteOwner Then strRetVal = strRetVal & "'  " & "adRightWriteOwner" & vbCrLf
12450   If lngPerm And adRightMaximumAllowed Then strRetVal = strRetVal & "'  " & "adRightMaximumAllowed" & vbCrLf
12460   If lngPerm And adRightFull Then strRetVal = strRetVal & "'  " & "adRightFull" & vbCrLf
12470   If lngPerm And adRightExecute Then strRetVal = strRetVal & "'  " & "adRightExecute" & vbCrLf
12480   If lngPerm And adRightUpdate Then strRetVal = strRetVal & "'  " & "adRightUpdate" & vbCrLf
12490   If lngPerm And adRightRead Then strRetVal = strRetVal & "'  " & "adRightRead" & vbCrLf

12500   If Right(strRetVal, 2) = vbCrLf Then strRetVal = Left(strRetVal, (Len(strRetVal) - 2))

        ' ** AdRights enumeration:
        ' **            0               adRightNone              The user or group has no permissions for the object.
        ' **          256 (&H100)       adRightDrop              The user or group has permission to remove objects
        ' **                                                     from the catalog. For example, Tables can be deleted
        ' **                                                     by a DROP TABLE SQL command.
        ' **          512 (&H200)       adRightExclusive         The user or group has permission to access the object
        ' **                                                     exclusively.
        ' **         1024 (&H400)       adRightReadDesign        The user or group has permission to read the design
        ' **                                                     for the object.
        ' **         2048 (&H800)       adRightWriteDesign       The user or group has permission to modify the design
        ' **                                                     for the object.
        ' **         4096 (&H1000)      adRightWithGrant         The user or group has permission to grant permissions
        ' **                                                     on the object.
        ' **         8192 (&H2000)      adRightReference         The user or group has permission to reference the object.
        ' **        16384 (&H4000)      adRightCreate            The user or group has permission to create new objects
        ' **                                                     of this type.
        ' **        32768 (&H8000)      adRightInsert            The user or group has permission to insert the object.
        ' **                                                     For objects such as Tables, the user has permission to
        ' **                                                     insert data into the table.
        ' **        65536 (&H10000)     adRightDelete            The user or group has permission to delete data from an
        ' **                                                     object. For objects such as Tables, the user has
        ' **                                                     permission to delete data values from records.
        ' **       131072 (&H20000)     adRightReadPermissions   The user or group can view, but not change, the specific
        ' **                                                     permissions for an object in the catalog.
        ' **       262144 (&H40000)     adRightWritePermissions  The user or group can modify the specific permissions
        ' **                                                     for an object in the catalog.
        ' **       524288 (&H80000)     adRightWriteOwner        The user or group has permission to modify the owner of
        ' **                                                     the object.
        ' **     33554432 (&H2000000)   adRightMaximumAllowed    The user or group has the maximum number of permissions
        ' **                                                     allowed by the provider. Specific permissions are
        ' **                                                     provider-dependent.
        ' **    268435456 (&H10000000)  adRightFull              The user or group has all permissions on the object.
        ' **    536870912 (&H20000000)  adRightExecute           The user or group has permission to execute the object.
        ' **   1073741824 (&H40000000)  adRightUpdate            The user or group has permission to update the object.
        ' **                                                     For objects such as Tables, the user has permission to
        ' **                                                     update the data in the table.
        ' **  -2147483648 (&H80000000)  adRightRead              The user or group has permission to read the object.
        ' **                                                     For objects such as Tables, the user has permission to
        ' **                                                     read the data in the table.

EXITP:
12510   Security_PermsDecode = strRetVal
12520   Exit Function

ERRH:
12530   strRetVal = RET_ERR
12540   Select Case ERR.Number
        Case Else
12550     Beep
12560     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
12570   End Select
12580   Resume EXITP

End Function

Public Function DeleteTable_ZZs() As Boolean
' ** Delete all my 'zz_' tables, as well as a few others, prior to release.
' ** Called by the Macro:
' **   zz_mcr_Delete_ZZTables

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "DeleteTable_ZZs"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, Rel As DAO.Relation
        Dim lngTbls As Long, arr_varTbl() As Variant
        Dim lngRels As Long, arr_varRel() As Variant
        Dim blnFound As Boolean
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varTbl().
        Const T_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const T_TNAM  As Integer = 0

        ' ** Array: arr_varRel().
        Const R_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const R_RNAM As Integer = 0
        Const R_FRGN As Integer = 1

12610   blnRetVal = True

12620   lngTbls = 0&
12630   ReDim arr_varTbl(T_ELEMS, 0)

12640   lngRels = 0&
12650   ReDim arr_varRel(R_ELEMS, 0)

12660   Set dbs = CurrentDb
12670   With dbs

          ' ** Collect the temp tables to be deleted.
12680     For Each tdf In .TableDefs
12690       With tdf
12700         If Left(.Name, 6) = "zz_tbl" Or Left(.Name, 6) = "zz_tmp" Or Left(.Name, 4) = "zz__" Then
12710           lngTbls = lngTbls + 1&
12720           lngE = lngTbls - 1&
12730           ReDim Preserve arr_varTbl(T_ELEMS, lngE)
12740           arr_varTbl(T_TNAM, lngE) = .Name
12750         End If
12760       End With
12770     Next

          ' ** Other temporary tables that may have only been used in query documentation.
12780     lngTbls = lngTbls + 1&
12790     lngE = lngTbls - 1&
12800     ReDim Preserve arr_varTbl(T_ELEMS, lngE)
12810     arr_varTbl(T_TNAM, lngE) = "account1"

12820     lngTbls = lngTbls + 1&
12830     lngE = lngTbls - 1&
12840     ReDim Preserve arr_varTbl(T_ELEMS, lngE)
12850     arr_varTbl(T_TNAM, lngE) = "ActiveAssets1"

12860     lngTbls = lngTbls + 1&
12870     lngE = lngTbls - 1&
12880     ReDim Preserve arr_varTbl(T_ELEMS, lngE)
12890     arr_varTbl(T_TNAM, lngE) = "ledger1"

12900     lngTbls = lngTbls + 1&
12910     lngE = lngTbls - 1&
12920     ReDim Preserve arr_varTbl(T_ELEMS, lngE)
12930     arr_varTbl(T_TNAM, lngE) = "masterasset1"

          'lngTbls = lngTbls + 1&
          'lngE = lngTbls - 1&
          'ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          'arr_varTbl(T_TNAM , lngE) = "tmpIncomeExpenseReports"

          'lngTbls = lngTbls + 1&
          'lngE = lngTbls - 1&
          'ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          'arr_varTbl(T_TNAM , lngE) = "tmp_RecurringItems"

          'lngTbls = lngTbls + 1&
          'lngE = lngTbls - 1&
          'ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          'arr_varTbl(T_TNAM , lngE) = "tmpRevCodeEdit"

          'lngTbls = lngTbls + 1&
          'lngE = lngTbls - 1&
          'ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          'arr_varTbl(T_TNAM , lngE) = "tmp_m_REVCODE"

          'lngTbls = lngTbls + 1&
          'lngE = lngTbls - 1&
          'ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          'arr_varTbl(T_TNAM , lngE) = "tmp_ActiveAssets"

          'lngTbls = lngTbls + 1&
          'lngE = lngTbls - 1&
          'ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          'arr_varTbl(T_TNAM , lngE) = "tmp_Journal"

          'lngTbls = lngTbls + 1&
          'lngE = lngTbls - 1&
          'ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          'arr_varTbl(T_TNAM , lngE) = "tmp_Ledger"

          'lngTbls = lngTbls + 1&
          'lngE = lngTbls - 1&
          'ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          'arr_varTbl(T_TNAM , lngE) = "USysRibbons"

12940     lngTbls = lngTbls + 1&
12950     lngE = lngTbls - 1&
12960     ReDim Preserve arr_varTbl(T_ELEMS, lngE)
12970     arr_varTbl(T_TNAM, lngE) = "tblVersion_Conversion_bak"   ' ** My history of version conversions.

12980     lngTbls = lngTbls + 1&
12990     lngE = lngTbls - 1&
13000     ReDim Preserve arr_varTbl(T_ELEMS, lngE)
13010     arr_varTbl(T_TNAM, lngE) = "zz_tmpUpdatedValues"

13020     For lngX = 0& To (lngTbls - 1&)
13030       For lngY = 0 To (lngTbls - 1&)
13040         If lngY <> lngX Then
13050           If arr_varTbl(T_TNAM, lngY) = arr_varTbl(T_TNAM, lngX) Then
13060             Debug.Print "'DUPE!  " & arr_varTbl(T_TNAM, lngY)
13070             Exit For
13080           End If
13090         End If
13100       Next
13110     Next

          'For lngX = 0& To (lngTbls - 1&)
          '  Debug.Print "'" & Left(CStr(lngX + 1&) & Space(4), 4) & arr_varTbl(T_TNAM, lngX)
          '  DoEvents
          'Next

'1   zz_tbl_DataComp_01
'2   zz_tbl_DataComp_02
'3   zz_tbl_DataComp_03
'4   zz_tbl_DataComp_04
'5   zz_tbl_DataComp_05
'6   zz_tbl_DataComp_06
'7   zz_tbl_DataComp_07
'8   zz_tbl_DataComp_08
'9   zz_tbl_Demonym
'10  zz_tbl_Demonym2
'11  zz_tbl_Form_Control_01
'12  zz_tbl_Form_Control_02
'13  zz_tbl_Form_Doc
'14  zz_tbl_Form_Graphics_01
'15  zz_tbl_Form_Property_Value
'16  zz_tbl_Geographical_Location
'17  zz_tbl_Language_Identifier
'18  zz_tbl_Report_List_01
'19  zz_tbl_Report_List_02
'20  zz_tbl_Statement_AssetList
'21  zz_tbl_Statement_Transaction
'22  zz_tbl_TAVersion_01
'23  zz_tbl_TAVersion_02
'24  zz_tbl_TransAudit_01
'25  zz_tbl_TransAudit_02
'26  zz_tbl_TransAudit_03
'27  zz_tbl_TransAudit_04
'28  zz_tbl_VBComponent_CompDir1
'29  zz_tbl_VBComponent_Local2
'30  zz_tbl_VBComponent_Procedure
'31  zz_tbl_VBComponent_Shortcut
'32  zz_tbl_VBComponent_Shortcut_02
'33  zz_tbl_VBComponent_Shortcut_03
'34  zz_tbl_VBComponent_Shortcut_04
'35  zz_tbl_VBComponent_Shortcut_05
'36  zz_tbl_VBComponent_Shortcut_06
'37  zz_tbl_VBComponent_Shortcut_07
'38  zz_tbl_VBComponent_Shortcut_08
'39  zz_tbl_VBComponent_Shortcut_09
'40  zz_tbl_VBComponent_Shortcut_10
'41  zz_tbl_VBComponent_Shortcut_11
'42  zz_tbl_VBComponent_Shortcut_12
'43  zz_tbl_VBComponent_Shortcut_13
'44  zz_tbl_VBComponent_Shortcut_14
'45  zz_tbl_VBComponent_Unknown
'46  account1
'47  ActiveAssets1
'48  ledger1
'49  masterasset1
'50  tblVersion_Conversion_bak
'51  zz_tmpUpdatedValues

          ' ** Collect any relationships they may be involved in.
13120     For lngX = 0& To (lngTbls - 1&)
13130       For Each Rel In .Relations
13140         With Rel
13150           If .Table = arr_varTbl(T_TNAM, lngX) Or .ForeignTable = arr_varTbl(T_TNAM, lngX) Then
13160             blnFound = False
13170             For lngY = 0& To (lngRels - 1&)
13180               If arr_varRel(R_RNAM, lngY) = .Name Then
13190                 blnFound = True
13200                 Exit For
13210               End If
13220             Next
13230             If blnFound = False Then
13240               lngRels = lngRels + 1&
13250               lngE = lngRels - 1&
13260               ReDim Preserve arr_varRel(R_ELEMS, lngE)
13270               arr_varRel(R_RNAM, lngE) = .Name
13280               arr_varRel(R_FRGN, lngE) = .ForeignTable
13290             End If
13300           End If
13310         End With
13320       Next
13330     Next

          ' ** Delete the relationships.
13340     For lngX = (lngRels - 1&) To 0& Step -1&
13350       .Relations.Delete arr_varRel(R_RNAM, lngX)
13360     Next

13370     .Relations.Refresh

          ' ** Delete the tables.
13380     For lngX = (lngTbls - 1&) To 0& Step -1&
13390       TableDelete CStr(arr_varTbl(T_TNAM, lngX))   ' ** Module Function: modFileUtilities.
13400     Next

13410     .TableDefs.Refresh

13420     .Close
13430   End With

13440   Beep

EXITP:
13450   Set Rel = Nothing
13460   Set tdf = Nothing
13470   Set dbs = Nothing
13480   Exit Function

ERRH:
13490   blnRetVal = False
13500   Select Case ERR.Number
        Case Else
13510     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
13520   End Select
13530   Resume EXITP

End Function

Public Function EmptyTable_AllTmp() As Boolean

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTable_AllTmp"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, strTmp05 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

        Const TBL_EMPTY As String = "qryTmp_Table_Empty_"

        ' ** mcrEmptyTmpTables:
        ' **   Turn warnings off
        ' **   Turn hourglass on
        ' **   zz_mod_MDEPrepFuncs.EmptyTable_AllTmp()
        ' **   zz_mod_MDEPrepFuncs.EmptyTable_TempZZ()
        ' **   zz_mod_MDEPrepFuncs.Setup_Demo()
        ' **   modAppVersionFuncs.AppIcon_Let()
        ' **   Turn hourglass off
        ' **   Turn warnings on
        ' **   modWindowFunctions.DoBeeps()

13610   blnRetVal = False
13620   lngX = 0&

13630   Set dbs = CurrentDb
13640   With dbs

          ' ** For the Demo, we need to make sure the EULA tables in TrstArch.mdb and TrustDta.mdb are empty as well.
          ' ** Since they're not ordinarily linked, they'll have to be in order to clear them out.
          ' ** The query qryTmp_Table_Empty_00d__~rmcp takes care of the EULA table here.
13650     If Len(TA_SEC) > Len(TA_SEC2) Then
            ' ** Demo version.
13660       If gstrTrustDataLocation = vbNullString Then
13670         IniFile_GetDataLoc  ' ** Module Function: modStartupFuncs.
13680       End If
13690       If TableExists("_~rmca") = False Then  ' ** Module Function: modFileUtilities.
13700         DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_ArchDataName), acTable, "_~rmca", "_~rmca"
13710       End If
13720       If TableExists("_~rmcd") = False Then  ' ** Module Function: modFileUtilities.
13730         DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), acTable, "_~rmcd", "_~rmcd"
13740       End If
13750       strTmp01 = "DELETE [_~rmca].* FROM [_~rmca];"
13760       Set qdf = .CreateQueryDef("", strTmp01)
13770       qdf.Execute
13780       strTmp01 = "DELETE [_~rmcd].* FROM [_~rmcd];"
13790       Set qdf = .CreateQueryDef("", strTmp01)
13800       qdf.Execute
13810       TableDelete "_~rmca"  ' ** Module Function: modFileUtilities.
13820       TableDelete "_~rmcd"  ' ** Module Function: modFileUtilities.
13830       strTmp01 = vbNullString
13840       Set qdf = Nothing
13850       gstrTrustDataLocation = vbNullString  ' ** Remember, this is run when we're cleaning everything up!
13860     End If

13870     For Each qdf In .QueryDefs
13880       strTmp01 = vbNullString: strTmp05 = vbNullString: strTmp03 = vbNullString
13890       With qdf
13900         If Left(.Name, Len(TBL_EMPTY)) = TBL_EMPTY Then
                ' ** Example: qryTmp_Table_Empty_030_tblTemplate_MasterAsset
13910           strTmp01 = Mid(.Name, (Len(TBL_EMPTY) + 1))  ' ** Starts with the number.
13920           intPos01 = InStr(strTmp01, "_")
13930           strTmp01 = Mid(strTmp01, (intPos01 + 1))  ' ** Strip the number, leaving only the table name.
13940           If strTmp01 = "Journal_Map" Then strTmp01 = "Journal Map"  ' ** qryTmp_Table_Empty_107_Journal_Map
13950           If strTmp01 = "License_Name" Then strTmp01 = "License Name"  ' ** qryTmp_Table_Empty_108_License_Name
13960           strTmp02 = "tblCheckReconcile_"  ' ** Handled by EmptyTable_CheckReconcile(), modCheckReconcile.
13970           strTmp03 = "tblPricing_"         ' ** Handled by EmptyTable_PricingTmp(), below.
13980           strTmp04 = "tmp_"                ' ** Handled by EmptyTable_Tmp(), below; without underscore, here.
13990           strTmp05 = "zz_tbl_RePost_"      ' ** Handled by EmptyTable_RePost(), below.
14000           If Left(strTmp01, Len(strTmp02)) <> strTmp02 And Left(strTmp01, Len(strTmp03)) <> strTmp03 And _
                    Left(strTmp01, Len(strTmp04)) <> strTmp04 And Left(strTmp01, Len(strTmp05)) <> strTmp05 Then
14010             If strTmp01 = "tblDatabase" Then
14020               If CurrentAppName = "Trust.mde" Then  ' ** Module Function: modFileUtilities.
                      ' ** Yes, absolutely run it.
14030                 qdf.Execute
14040               Else
      #If Not IsDev Then
                      ' ** Only run if release version.
14050                 qdf.Execute
      #End If
14060               End If
14070             Else
14080               If TableExists(strTmp01) = True Then  ' ** Module Function: modFileUtilities.
14090                 lngX = lngX + 1&
14100                 qdf.Execute
14110               End If
14120             End If
14130           End If
14140         End If
14150       End With
14160     Next

14170     For lngX = 1& To 2&
14180       Select Case lngX
            Case 1&
              ' ** Update qrySecurity_License_06a (tblDatabase, with dbs_path_ta_new, by specified [trustdir], [datadir]).
14190         Set qdf = .QueryDefs("qrySecurity_License_07a")
14200       Case 2&
              ' ** Update qrySecurity_License_06b (tblTemplate_Database, with dbs_path_ta_new, by specified [trustdir], [datadir]).
14210         Set qdf = .QueryDefs("qrySecurity_License_07b")
14220       End Select
14230       With qdf.Parameters
14240         If InStr(dbs.Name, gstrDir_Dev) = 0 Then
14250           If InStr(dbs.Name, gstrDir_Def) > 0 Then
14260             ![trustdir] = gstrDir_Def
14270             ![datadir] = gstrDir_Def & LNK_SEP & "Database"
14280           ElseIf InStr(dbs.Name, gstrDir_Def64) > 0 Then
14290             ![trustdir] = gstrDir_Def64
14300             ![datadir] = gstrDir_Def64 & LNK_SEP & "Database"
14310           Else
14320             ![trustdir] = CurrentAppPath  ' ** Module Function: modFileUtilities.
14330             ![datadir] = CurrentAppPath & LNK_SEP & "Database"  ' ** Module Function: modFileUtilities.
14340           End If
14350         Else
14360           Beep
14370           ![trustdir] = gstrDir_Dev
14380           ![datadir] = gstrDir_Dev & LNK_SEP & gstrDir_DevEmpty
14390         End If
14400       End With
14410       qdf.Execute
14420     Next

14430     .Close
14440   End With

14450   EmptyTable_CheckReconcile  ' ** Function: Below.
14460   EmptyTable_PricingTmp  ' ** Function: Below.
14470   EmptyTable_Tmp  ' ** Function: Below.
14480   EmptyTable_RePost  ' ** Function: Below.

EXITP:
14490   Set qdf = Nothing
14500   Set dbs = Nothing
14510   EmptyTable_AllTmp = blnRetVal
14520   Exit Function

ERRH:
14530   blnRetVal = False
14540   Select Case ERR.Number
        Case Else
14550     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
14560   End Select
14570   Resume EXITP

End Function

Public Function EmptyTable_CheckReconcile() As Boolean
' ** Empty all the Check Reconcile tables, both temporary and permanent.

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTable_CheckReconcile"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngQrys As Long, arr_varQry As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        'Const Q_QID As Integer = 0
        Const Q_NAM As Integer = 1
        'Const Q_TYP As Integer = 2
        'Const Q_DSC As Integer = 3
        'Const Q_NUM As Integer = 4

14610   blnRetVal = True

14620   Set dbs = CurrentDb
14630   With dbs
          ' ** tblQuery, just empty tblCheckReconcile_.. queries.  '## OK
14640     Set qdf = .QueryDefs("qryCheckReconcile_17")
14650     Set rst = qdf.OpenRecordset
14660     With rst
14670       If .BOF = True And .EOF = True Then
              ' ** Oops! I must've changed something!
14680         lngQrys = 0&
14690       Else
14700         .MoveLast
14710         lngQrys = .RecordCount
14720         .MoveFirst
14730         arr_varQry = .GetRows(lngQrys)
              ' ****************************************************
              ' ** Array: arr_varQry()
              ' **
              ' **   Field  Element  Name               Constant
              ' **   =====  =======  =================  ==========
              ' **     1       0     qry_id             Q_QID
              ' **     2       1     qry_name           Q_NAM
              ' **     3       2     qrytype_type       Q_TYP
              ' **     4       3     qry_description    Q_DSC
              ' **     5       4     qry_emptynum       Q_NUM
              ' **
              ' ****************************************************
14740       End If
14750       .Close
14760     End With
14770     For lngX = 0& To (lngQrys - 1&)
14780       Set qdf = .QueryDefs(arr_varQry(Q_NAM, lngX))
14790       qdf.Execute
14800     Next
14810     .Close
14820   End With

        'qryTmp_Table_Empty_181_tblCheckReconcile_Account
        'qryTmp_Table_Empty_182_tblCheckReconcile_Check
        'qryTmp_Table_Empty_183_tblCheckReconcile_Item
        'qryTmp_Table_Empty_184_tblCheckReconcile_Staging

EXITP:
14830   Set rst = Nothing
14840   Set qdf = Nothing
14850   Set dbs = Nothing
14860   EmptyTable_CheckReconcile = blnRetVal
14870   Exit Function

ERRH:
14880   blnRetVal = False
14890   Select Case ERR.Number
        Case Else
14900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14910   End Select
14920   Resume EXITP

End Function

Public Function EmptyTable_PricingTmp() As Boolean
' ** Empty all the temporary Pricing tables.

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTable_PricingTmp"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnSkipHistory As Boolean
        Dim blnRetVal As Boolean

        Const TBL_EMPTY As String = "qryTmp_Table_Empty_"
        Const TBL_PRICE As String = "tblPricing_"

15010   blnRetVal = True

15020   blnSkipHistory = True  ' ** True: Don't empty tblPricing_MasterAsset_History; False: Empty it.

15030   Set dbs = CurrentDb
15040   With dbs
15050     For Each qdf In .QueryDefs
15060       With qdf
15070         If Left(.Name, Len(TBL_EMPTY)) = TBL_EMPTY Then
15080           If Mid(.Name, Len(TBL_EMPTY) + 5, Len(TBL_PRICE)) = TBL_PRICE Then  ' ** 3-digit numbers now!
15090             If blnSkipHistory = True And Right(.Name, 8) = "_History" Then
                    ' ** Skip it.
15100             Else
15110               .Execute
15120             End If
15130           End If
15140         End If
15150       End With
15160     Next
15170     .Close
15180   End With

15190   Beep

EXITP:
15200   EmptyTable_PricingTmp = blnRetVal
15210   Exit Function

ERRH:
15220   blnRetVal = False
15230   Select Case ERR.Number
        Case Else
15240     Beep
15250     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
15260   End Select
15270   Resume EXITP

End Function

Public Function EmptyTable_Tmp() As Boolean
' ** If the tables prefixed tmp_.. exist, empty them.
' ** I think these are only used by the procedures in modBackendUpdate.
' **   tmp_ActiveAssets
' **   tmp_Journal
' **   tmp_Ledger
' **   tmp_m_REVCODE
' **   tmp_RecurringItems

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTable_Tmp"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngTmpTbls As Long, arr_varTmpTbl As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

15310   blnRetVal = True

15320   Set dbs = CurrentDb
15330   With dbs
15340     arr_varTmpTbl = TmpTblList  ' ** Module Function: modBackendUpdate.
15350     lngTmpTbls = UBound(arr_varTmpTbl) + 1&
15360     For lngX = 0& To (lngTmpTbls - 1&)
15370       Select Case arr_varTmpTbl(lngX)
            Case "tmp_ActiveAssets"
              ' ** Empty tmp_ActiveAssets.
15380         TableEmpty "tmp_ActiveAssets"  ' ** Module Function: modFileUtilities.
15390       Case "tmp_Journal"
              ' ** Empty tmp_Journal.
15400         TableEmpty "tmp_Journal"  ' ** Module Function: modFileUtilities.
15410       Case "tmp_Ledger"
              ' ** Empty tmp_Ledger.
15420         TableEmpty "tmp_Ledger"  ' ** Module Function: modFileUtilities.
15430       Case "tmp_m_REVCODE"
              ' ** Empty tmp_m_REVCODE.
15440         TableEmpty "tmp_m_REVCODE"  ' ** Module Function: modFileUtilities.
15450       Case "tmp_RecurringItems"
              ' ** Empty tmp_RecurringItems.
15460         TableEmpty "tmp_RecurringItems"  ' ** Module Function: modFileUtilities.
15470       End Select
15480     Next
15490     .Close
15500   End With

15510   Beep

EXITP:
15520   Set qdf = Nothing
15530   Set dbs = Nothing
15540   EmptyTable_Tmp = blnRetVal
15550   Exit Function

ERRH:
15560   blnRetVal = False
15570   Select Case ERR.Number
        Case Else
15580     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
15590   End Select
15600   Resume EXITP

End Function

Public Function EmptyTable_RePost() As Boolean
' ** If the tables prefixed 'zz_tbl_RePost_' exists, empty them.
' ** These are used by the procedures in modRePostFuncs.
' **   zz_tbl_Repost_Account
' **   zz_tbl_RePost_ActiveAssets
' **   zz_tbl_RePost_Debug
' **   zz_tbl_RePost_Journal
' **   zz_tbl_RePost_Ledger
' **   zz_tbl_RePost_MasterAsset
' **   zz_tbl_RePost_Posting
' ** Don't empty zz_tbl_RePost_Statement_Date!

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTable_RePost"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRePostTbls As Long, arr_varRePostTbl As Variant
        Dim blnSkip As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

15710   blnRetVal = True

        ' ** VGC 09/20/2015: REMOVED modRePostFuncs FROM RELEASE.
15720   blnSkip = True
15730   If blnSkip = False Then

15740     Set dbs = CurrentDb
15750     With dbs
            'arr_varRePostTbl = RePost_TblList  ' ** Module Function: modRePostFuncs.
15760       lngRePostTbls = UBound(arr_varRePostTbl) + 1&
15770       For lngX = 0& To (lngRePostTbls - 1&)
15780         Select Case arr_varRePostTbl(lngX)
              Case "zz_tbl_Repost_Account"
15790           If TableExists("zz_tbl_Repost_Account") = True Then  ' ** Module Function: modFileUtilities.
                  ' ** Empty zz_tbl_Repost_Account.
15800             Set qdf = .QueryDefs("zz_qry_Tmp_Table_Empty_96a_zz_tbl_Repost_Account")
15810             qdf.Execute
15820           End If
15830         Case "zz_tbl_RePost_ActiveAssets"
15840           If TableExists("zz_tbl_RePost_ActiveAssets") = True Then  ' ** Module Function: modFileUtilities.
                  ' ** Empty zz_tbl_RePost_ActiveAssets.
15850             Set qdf = .QueryDefs("zz_qry_Tmp_Table_Empty_96b_zz_tbl_RePost_ActiveAssets")
15860             qdf.Execute
15870           End If
15880         Case "zz_tbl_RePost_Debug"
15890           If TableExists("zz_tbl_RePost_Debug") = True Then  ' ** Module Function: modFileUtilities.
                  ' ** Empty zz_tbl_RePost_Debug.
15900             Set qdf = .QueryDefs("zz_qry_Tmp_Table_Empty_96c_zz_tbl_RePost_Debug")
15910             qdf.Execute
15920           End If
15930         Case "zz_tbl_RePost_Journal"
15940           If TableExists("zz_tbl_RePost_Journal") = True Then  ' ** Module Function: modFileUtilities.
                  ' ** Empty zz_tbl_RePost_Journal.
15950             Set qdf = .QueryDefs("zz_qry_Tmp_Table_Empty_96d_zz_tbl_RePost_Journal")
15960             qdf.Execute
15970           End If
15980         Case "zz_tbl_RePost_Journal_Error"
15990           If TableExists("zz_tbl_RePost_Journal_Error") = True Then  ' ** Module Function: modFileUtilities.
                  ' ** Empty zz_tbl_RePost_Journal_Error.
16000             Set qdf = .QueryDefs("zz_qry_Tmp_Table_Empty_96e_zz_tbl_RePost_Journal_Error")
16010             qdf.Execute
16020           End If
16030         Case "zz_tbl_RePost_Ledger"
16040           If TableExists("zz_tbl_RePost_Ledger") = True Then  ' ** Module Function: modFileUtilities.
                  ' ** Empty zz_tbl_RePost_Ledger.
16050             Set qdf = .QueryDefs("zz_qry_Tmp_Table_Empty_96f_zz_tbl_RePost_Ledger")
16060             qdf.Execute
16070           End If
16080         Case "zz_tbl_RePost_MasterAsset"
16090           If TableExists("zz_tbl_RePost_MasterAsset") = True Then  ' ** Module Function: modFileUtilities.
                  ' ** Empty zz_tbl_RePost_MasterAsset.
16100             Set qdf = .QueryDefs("zz_qry_Tmp_Table_Empty_96g_zz_tbl_RePost_MasterAsset")
16110             qdf.Execute
16120           End If
16130         Case "zz_tbl_RePost_Posting"
16140           If TableExists("zz_tbl_RePost_Posting") = True Then  ' ** Module Function: modFileUtilities.
                  ' ** Empty zz_tbl_RePost_Posting.
16150             Set qdf = .QueryDefs("zz_qry_Tmp_Table_Empty_96h_zz_tbl_RePost_Posting")
16160             qdf.Execute
16170           End If
16180         End Select
16190       Next
16200       .Close
16210     End With

16220     Beep

16230   End If  ' ** blnSkip.

EXITP:
16240   Set qdf = Nothing
16250   Set dbs = Nothing
16260   EmptyTable_RePost = blnRetVal
16270   Exit Function

ERRH:
16280   blnRetVal = False
16290   Select Case ERR.Number
        Case Else
16300     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
16310   End Select
16320   Resume EXITP

End Function

Public Function EmptyTable_TempZZ() As Boolean
' ** Empty zz_tbl's that hold temporary data.
' ** SEE DeleteTable_ZZs(), BELOW!

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTable_TempZZ"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, tdf As DAO.TableDef
        Dim lngZZTbls As Long, arr_varZZTbl() As Variant
        Dim strSQL As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

16410   blnRetVal = True

16420   lngZZTbls = 0&
16430   ReDim arr_varZZTbl(0)

16440   Set dbs = CurrentDb
16450   With dbs

          ' ** zz_.. Tables:
          ' **   tblDatabase_Table, just zz_tbl_RePost_.. tables; 13.
          ' **   zz_qry_Database_Table_16a
          ' **     Handled by EmptyTable_RePost(), above.
          ' **   tblDatabase_Table, just zz_tbl_tmp.. tables; 3.
          ' **   zz_qry_Database_Table_17a
          'CHECK THESE OUT! DELETED!
          ' **     zz_tbl_tmpAccount_01
          ' **     zz_tbl_tmpMasterAsset
          ' **     zz_tbl_tmpWindowsSecurityEvents
          ' **   tblDatabase_Table, just zz_tbl_.. tables; 63.
          ' **   zz_qry_Database_Table_18a
          ' **     Standard development tables.
          ' **     Exceptions: zz_tbl_Client_.., zz_tbl_Dev_..
          ' **   tblDatabase_Table, just zz_tmp.. tables; 20.
          ' **   zz_qry_Database_Table_19a
          ' **     All backups of permanent temporary tables
          ' **     All may be emptied below.
          ' **   tblDatabase_Table, just zz_.. tables; 6.
          ' **   zz_qry_Database_Table_20a
          ' **     All but zz_USysRibbons may be emptied below.

16460     For Each tdf In .TableDefs
16470       With tdf
16480         If Left(.Name, 3) = "zz_" Then
16490           If (Left(.Name, Len("zz_tbl_Client_")) = "zz_tbl_Client_") Or _
                    (Left(.Name, Len("zz_tbl_Dev_")) = "zz_tbl_Dev_") Or _
                    (Left(.Name, Len("zz_tbl_RePost_")) = "zz_tbl_RePost_") Or _
                    .Name = "zz_USysRibbons" Then
                  ' ** Don't empty these!
                Else
16500             lngZZTbls = lngZZTbls + 1&
16510             lngE = lngZZTbls - 1&
16520             ReDim Preserve arr_varZZTbl(lngE)
16530             arr_varZZTbl(lngE) = .Name
16540           End If
16550         End If
16560       End With
16570     Next

16580     For lngX = 0& To (lngZZTbls - 1&)
16590       If TableExists(CStr(arr_varZZTbl(lngX))) = True Then  ' ** Module Function: modFileUtilities.
16600         strSQL = "DELETE [" & arr_varZZTbl(lngX) & "].* FROM [" & arr_varZZTbl(lngX) & "];"
16610         Set qdf = .CreateQueryDef("", strSQL)
16620         qdf.Execute
16630       End If
16640     Next

16650     .Close
16660   End With

16670   Beep

EXITP:
16680   Set qdf = Nothing
16690   Set tdf = Nothing
16700   Set dbs = Nothing
16710   EmptyTable_TempZZ = blnRetVal
16720   Exit Function

ERRH:
16730   blnRetVal = False
16740   Select Case ERR.Number
        Case Else
16750     Beep
16760     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
16770   End Select
16780   Resume EXITP

End Function

Public Function EmptyTable_AllTmp_Renumber() As Boolean
' ** Renumber the Table_Empty queries, when a new one is inserted.

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTable_AllTmp_Renumber"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim lngLastNum As Long
        Dim blnRenumbered As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        Const TBL_EMPTY As String = "qryTmp_Table_Empty_"

        Const strHighNum As String = "_165_"  ' ** Don't hit this number; already occupied.
        Const lngHighNum As Long = 165&

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const Q_NAM_OLD As Integer = 0
        Const Q_NUM_OLD As Integer = 1
        Const Q_TNAM    As Integer = 2
        Const Q_NAM_NEW As Integer = 3
        Const Q_NUM_NEW As Integer = 4

16810   blnRetVal = True
16820   blnRenumbered = False

16830   lngQrys = 0&
16840   ReDim arr_varQry(Q_ELEMS, 0)

16850   Set dbs = CurrentDb
16860   With dbs

16870     For Each qdf In .QueryDefs
16880       With qdf
16890         If Left(.Name, Len(TBL_EMPTY)) = TBL_EMPTY Then
16900           If InStr(.Name, strHighNum) = 0 Then
16910             strTmp01 = Mid(.Name, Len(TBL_EMPTY))    ' ** Includes starting underscore.
16920             intPos01 = InStr(2, strTmp01, "_")         ' ** Find next underscore.
16930             strTmp02 = Left(strTmp01, intPos01)       ' ** Includes next underscore.
16940             If Left(strTmp02, 4) <> "_000" Then  ' ** 3-digit numbers now!
16950               lngTmp03 = Val(Mid(Left(strTmp02, (Len(strTmp02) - 1)), 3))  ' ** Query number.
16960               strTmp01 = Mid(strTmp01, (intPos01 + 1))                       ' ** Table name.
16970               If lngTmp03 < lngHighNum Then
16980                 lngQrys = lngQrys + 1&
16990                 lngE = lngQrys - 1&
17000                 ReDim Preserve arr_varQry(Q_ELEMS, lngE)
17010                 arr_varQry(Q_NAM_OLD, lngE) = .Name
17020                 arr_varQry(Q_NUM_OLD, lngE) = lngTmp03
17030                 arr_varQry(Q_TNAM, lngE) = strTmp01
17040                 arr_varQry(Q_NAM_NEW, lngE) = vbNullString
17050                 arr_varQry(Q_NUM_NEW, lngE) = CLng(0)
17060               Else
                      Stop
17070                 If (lngTmp03 >= 94& And lngTmp03 <= 96&) And (Asc(Mid(strTmp02, 4, 1)) >= 97 And Asc(Mid(strTmp02, 4, 1)) <= 122) Then
                        ' ** These should remain as-is.
17080                 Else
17090                   blnRetVal = False
17100                   Beep
17110                   Debug.Print "'HIGH NUM HIT! " & .Name
17120                   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
17130                   DoEvents
17140                   Exit For
17150                 End If
17160               End If  ' ** lngHighNum.
17170             Else
                    ' ** For now, I'll handle these manually.
                    'qryTmp_Table_Empty_000a_m_VP
                    'qryTmp_Table_Empty_000b_m_VD
                    'qryTmp_Table_Empty_000c_m_VA
                    'qryTmp_Table_Empty_000d__~rmcp
17180             End If  ' ** _000.
17190           Else
17200             Exit For
17210           End If  ' ** strHighNum.
17220         End If  ' ** TBL_EMPTY.
17230       End With  ' ** qdf.
17240     Next  ' ** qdf.

17250     If blnRetVal = True Then

            ' ** Find out where it's out of whack.
17260       lngLastNum = 0&
17270       For lngX = 0& To (lngQrys - 1&)
17280         If arr_varQry(Q_NUM_OLD, lngX) = lngLastNum + 1& Then
                ' ** Sequence OK so far.
17290           lngLastNum = lngLastNum + 1&
17300           arr_varQry(Q_NUM_NEW, lngX) = lngLastNum  ' ** And leave Q_NAM_NEW empty.
17310         Else
17320           lngLastNum = lngLastNum + 1&
17330           arr_varQry(Q_NUM_NEW, lngX) = lngLastNum
17340           arr_varQry(Q_NAM_NEW, lngX) = TBL_EMPTY & Right("000" & CStr(lngLastNum), 3) & "_" & arr_varQry(Q_TNAM, lngX)
17350         End If
17360       Next

            ' ** Check the sequence.
17370       lngLastNum = 0&
17380       For lngX = 0& To (lngQrys - 1&)
17390         If arr_varQry(Q_NUM_NEW, lngX) <> (lngLastNum + 1&) Then
17400           Stop
17410         Else
17420           lngLastNum = lngLastNum + 1&
17430         End If
17440       Next

            ' ** Now renumber them.
17450       For lngX = 0& To (lngQrys - 1&)
17460         If arr_varQry(Q_NAM_NEW, lngX) <> vbNullString Then
17470           blnRenumbered = True
17480           DoCmd.Rename arr_varQry(Q_NAM_NEW, lngX), acQuery, arr_varQry(Q_NAM_OLD, lngX)
17490         End If
17500       Next

17510     End If  ' ** blnRetVal.

17520     .Close
17530   End With  ' ** dbs.

17540   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

17550   Beep
17560   Select Case blnRenumbered
        Case True
17570     Debug.Print "'DONE! " & THIS_PROC & "()"
17580   Case False
17590     Debug.Print "'NONE FOUND!"
17600   End Select

EXITP:
17610   Set qdf = Nothing
17620   Set dbs = Nothing
17630   EmptyTable_AllTmp_Renumber = blnRetVal
17640   Exit Function

ERRH:
17650   blnRetVal = False
17660   Select Case ERR.Number
        Case Else
17670     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
17680   End Select
17690   Resume EXITP

End Function

Public Function CAM() As Boolean

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "CAM"

        Dim blnRetVal As Boolean

17710   blnRetVal = CloseAllModules  ' ** Function: Below.

EXITP:
17720   CAM = blnRetVal
17730   Exit Function

ERRH:
17740   blnRetVal = False
17750   Select Case ERR.Number
        Case Else
17760     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
17770   End Select
17780   Resume EXITP

End Function

Public Function CAMF() As Boolean

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "CAMF"

        Dim blnRetVal As Boolean

17810   blnRetVal = CloseAllModulesF  ' ** Function: Below.

EXITP:
17820   CAMF = blnRetVal
17830   Exit Function

ERRH:
17840   blnRetVal = False
17850   Select Case ERR.Number
        Case Else
17860     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
17870   End Select
17880   Resume EXITP

End Function

Public Function CAMR() As Boolean

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "CAMR"

        Dim blnRetVal As Boolean

17910   blnRetVal = CloseAllModulesR  ' ** Function: Below.

EXITP:
17920   CAMR = blnRetVal
17930   Exit Function

ERRH:
17940   blnRetVal = False
17950   Select Case ERR.Number
        Case Else
17960     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
17970   End Select
17980   Resume EXITP

End Function

Public Function CAF() As Boolean

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "CAF"

        Dim blnRetVal As Boolean

18010   blnRetVal = CloseAllForms  ' ** Function: Below.

EXITP:
18020   CAF = blnRetVal
18030   Exit Function

ERRH:
18040   blnRetVal = False
18050   Select Case ERR.Number
        Case Else
18060     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
18070   End Select
18080   Resume EXITP

End Function

Public Function CAR() As Boolean

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "CAR"

        Dim blnRetVal As Boolean

18110   blnRetVal = CloseAllReports  ' ** Function: Below.

EXITP:
18120   CAR = blnRetVal
18130   Exit Function

ERRH:
18140   blnRetVal = False
18150   Select Case ERR.Number
        Case Else
18160     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
18170   End Select
18180   Resume EXITP

End Function

Public Function CloseAllModules() As Boolean

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "CloseAllModules"

        Dim mdl As Access.Module
        Dim intMods As Integer
        Dim intX As Integer
        Dim blnRetVal As Boolean

18210   blnRetVal = True

18220   intMods = Application.Modules.Count
18230   For intX = (intMods - 1) To 0 Step -1
18240     Set mdl = Application.Modules(intX)
18250     With mdl
18260       If .Name <> "zz_mod_ModuleFormatFuncs" Then
18270         DoCmd.Close acModule, .Name
18280         If Left(.Name, 5) <> "Form_" Or Left(.Name, 7) <> "Report_" Then
                ' ** Only save form modules when the form is closed.
                'DoCmd.Save acModule, .Name
18290         End If
18300       End If
18310     End With
18320   Next

18330   Beep

EXITP:
18340   Set mdl = Nothing
18350   CloseAllModules = blnRetVal
18360   Exit Function

ERRH:
18370   blnRetVal = False
18380   Select Case ERR.Number
        Case Else
18390     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
18400   End Select
18410   Resume EXITP

End Function

Public Function CloseAllModulesF() As Boolean

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "CloseAllModulesF"

        Dim mdl As Access.Module
        Dim intMods As Integer
        Dim intX As Integer
        Dim blnRetVal As Boolean

18510   blnRetVal = True

18520   intMods = Application.Modules.Count
18530   For intX = (intMods - 1) To 0 Step -1
18540     Set mdl = Application.Modules(intX)
18550     With mdl
18560       If Left(.Name, 5) = "Form_" Then
18570         DoCmd.Close acModule, .Name
18580       End If
18590     End With
18600   Next

18610   Beep

EXITP:
18620   Set mdl = Nothing
18630   CloseAllModulesF = blnRetVal
18640   Exit Function

ERRH:
18650   blnRetVal = False
18660   Select Case ERR.Number
        Case Else
18670     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
18680   End Select
18690   Resume EXITP

End Function

Public Function CloseAllModulesR() As Boolean

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "CloseAllModulesR"

        Dim mdl As Access.Module
        Dim intMods As Integer
        Dim intX As Integer
        Dim blnRetVal As Boolean

18710   blnRetVal = True

18720   intMods = Application.Modules.Count
18730   For intX = (intMods - 1) To 0 Step -1
18740     Set mdl = Application.Modules(intX)
18750     With mdl
18760       If Left(.Name, 7) = "Report_" Then
18770         DoCmd.Close acModule, .Name
18780       End If
18790     End With
18800   Next

18810   Beep

EXITP:
18820   Set mdl = Nothing
18830   CloseAllModulesR = blnRetVal
18840   Exit Function

ERRH:
18850   blnRetVal = False
18860   Select Case ERR.Number
        Case Else
18870     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
18880   End Select
18890   Resume EXITP

End Function

Public Function CloseAllForms() As Boolean

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "CloseAllForms"

        Dim intFrms As Integer
        Dim intX As Integer
        Dim blnRetVal As Boolean

18910   blnRetVal = True

18920   intFrms = Application.Forms.Count
18930   For intX = (intFrms - 1) To 0 Step -1
18940     DoCmd.Close acForm, Application.Forms(intX).Name, acSaveYes
18950   Next

18960   Beep

EXITP:
18970   CloseAllForms = blnRetVal
18980   Exit Function

ERRH:
18990   blnRetVal = False
19000   Select Case ERR.Number
        Case Else
19010     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
19020   End Select
19030   Resume EXITP

End Function

Public Function CloseAllReports() As Boolean

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "CloseAllReports"

        Dim intRpts As Integer
        Dim intX As Integer
        Dim blnRetVal As Boolean

19110   blnRetVal = True

19120   intRpts = Application.Reports.Count
19130   For intX = (intRpts - 1) To 0 Step -1
19140     DoCmd.Close acReport, Application.Reports(intX).Name, acSaveNo
19150   Next

19160   Beep

EXITP:
19170   CloseAllReports = blnRetVal
19180   Exit Function

ERRH:
19190   blnRetVal = False
19200   Select Case ERR.Number
        Case Else
19210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19220   End Select
19230   Resume EXITP

End Function

Public Function RetSuper() As Boolean
' ** Puts 'superuser' into ActiveControl, for when you're
' ** entering data into raw tables and you need the Username.

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "RetSuper"

        Dim ctl As Access.Control
        Dim blnRetVal As Boolean

19310   blnRetVal = True

19320   Set ctl = Screen.ActiveControl
19330   ctl.text = "Superuser"

EXITP:
19340   Set ctl = Nothing
19350   RetSuper = blnRetVal
19360   Exit Function

ERRH:
19370   Select Case ERR.Number
        Case 2135  ' ** This property is read-only and can't be set.
          ' ** Ignore.
19380     Beep
19390   Case Else
19400     Beep
19410     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()", _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
19420   End Select
19430   Resume EXITP

End Function

Public Function Setup_AutoNumber() As Boolean
' ** Reset the AutoNumber field of all TrustDta.mdb tables that
' ** have an AutoNumber field, and that accept new records.

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "Setup_AutoNumber"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngAutos As Long, arr_varAuto As Variant
        Dim lngSuccess As Long, lngFail As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

        Const A_NAM As Integer = 1

19510   blnRetVal = True

19520   Set dbs = CurrentDb
19530   With dbs
19540     Set qdf = .QueryDefs("qrySystemUpdate_14_AutoNumber")
19550     Set rst = qdf.OpenRecordset
19560     With rst
19570       .MoveLast
19580       lngAutos = .RecordCount
19590       .MoveFirst
19600       arr_varAuto = .GetRows(lngAutos)
            ' ****************************************************
            ' ** Array: arr_varAuto()
            ' **
            ' **   Field  Element  Name               Constant
            ' **   =====  =======  =================  ==========
            ' **     1       0     mtbl_ID
            ' **     2       1     mtbl_NAME          A_NAM
            ' **     3       2     mtbl_AUTONUMBER
            ' **     4       3     mtbl_NEWRecs
            ' **     5       4     mtbl_DTA
            ' **
            ' ****************************************************
19610       .Close
19620     End With
19630     .Close
19640   End With

19650   lngSuccess = 0&: lngFail = 0&
19660   For lngX = 0& To (lngAutos - 1&)
19670     blnRetVal = ChangeSeed_Ext(arr_varAuto(A_NAM, lngX))  ' ** Module Function: modAutonumberFieldFuncs
19680     If blnRetVal = True Then
19690       lngSuccess = lngSuccess + 1&
19700     Else
19710       lngFail = lngFail + 1&
19720     End If
19730   Next

19740   Win_Mod_Restore  ' ** Procedure: Below.

19750   Debug.Print "'AUTONUM RESET - OK: " & CLng(lngSuccess) & "  ERR: " & CStr(lngFail)

19760   Beep

EXITP:
19770   Set rst = Nothing
19780   Set qdf = Nothing
19790   Set dbs = Nothing
19800   Setup_AutoNumber = blnRetVal
19810   Exit Function

ERRH:
19820   blnRetVal = False
19830   Select Case ERR.Number
        Case Else
19840     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
19850   End Select
19860   Resume EXITP

End Function

Public Function Setup_Demo() As Boolean

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "Setup_Demo"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnRetVal As Boolean

      #If IsDemo Then
        Const strZetaPathFile As String = "C:\VictorGCS_Clients\TrustAccountant\NewDemo\Trust.mdb"
      #Else
        Const strZetaPathFile As String = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Trust.mdb"
      #End If

19910   blnRetVal = True

19920   If Len(TA_SEC) > Len(TA_SEC2) Then

19930     If TableExists("tblTemplate_Zeta3") = False Then  ' ** Module Function: modFileUtilities.
19940       DoCmd.TransferDatabase acLink, "Microsoft Access", strZetaPathFile, acTable, "tblTemplate_Zeta3", "tblTemplate_Zeta3"
19950     End If
19960     If TableExists("tblTemplate_Zeta4") = False Then
19970       DoCmd.TransferDatabase acLink, "Microsoft Access", strZetaPathFile, acTable, "tblTemplate_Zeta4", "tblTemplate_Zeta4"
19980     End If

19990     DoEvents

          ' ** Reset the AutoNumber field in tblSecurity_User.
20000     ChangeSeed_Ext "tblSecurity_User"  ' ** Module Function: modAutonumberFieldFuncs.
          ' ** Reset the AutoNumber field in tblSecurity_GroupUser.
20010     ChangeSeed_Ext "tblSecurity_GroupUser"  ' ** Module Function: modAutonumberFieldFuncs.

20020     Set dbs = CurrentDb
20030     With dbs
            ' ** Check for 'Expr1:' nomenclature.
20040       Qry_RemExpr_rel False, True, "qrySecurity_User_11"  ' ** Module Function: modQueryFunctions1.
            ' ** Append qrySecurity_User_10 (tblTemplate_Zeta4, just 'TADemo', 'mary', 'bill'), not in Users, to Users.
20050       Set qdf = .QueryDefs("qrySecurity_User_11")
20060       qdf.Execute dbFailOnError
            ' ** Append qrySecurity_User_12 (Users, just 'TADemo', 'mary', 'bill'), not in tblSecurity_User, to tblSecurity_User.
20070       Set qdf = .QueryDefs("qrySecurity_User_13")
20080       qdf.Execute dbFailOnError
            ' ** Append qrySecurity_User_14 (tblSecurity_User, with tblSecurity_Group, just 'TADemo', 'mary', 'bill',
            ' ** for secgrp_name = 'Users'; Cartesian), not in tblSecurity_GroupUser, to tblSecurity_GroupUser.
20090       Set qdf = .QueryDefs("qrySecurity_User_16")
20100       qdf.Execute dbFailOnError
            ' ** Append qrySecurity_User_17 (tblSecurity_User, with tblSecurity_Group, just 'TADemo', 'mary', 'bill',
            ' ** for secgrp_name = 'Admins', 'DataEntry'; Cartesian), not in tblSecurity_GroupUser, to tblSecurity_GroupUser.
20110       Set qdf = .QueryDefs("qrySecurity_User_19")
20120       qdf.Execute dbFailOnError
20130       .Close
20140     End With

          'TableDelete "tblTemplate_Zeta3"  ' ** Module Function: modFileUtilities.
          'TableDelete "tblTemplate_Zeta4"  ' ** Module Function: modFileUtilities.

20150   End If

EXITP:
20160   Set qdf = Nothing
20170   Set dbs = Nothing
20180   Setup_Demo = blnRetVal
20190   Exit Function

ERRH:
20200   blnRetVal = False
20210   Select Case ERR.Number
        Case Else
20220     Beep
20230     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()", _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
20240   End Select
20250   Resume EXITP

End Function

Public Function Setup_Location() As Boolean
' ** Called by:
' **   zz_mcr_Clean_RecurItems.

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "Setup_Location"

        Dim blnRetVal As Boolean

20310   blnRetVal = True

        ' ** Empty tblTemplate_Location of all but the first record.
20320   TableEmpty "tblTemplate_Location"  ' ** Module Function: modFileUtilities.

        ' ** Reset its AutoNumber field to 2.
20330   ChangeSeed_Ext "tblTemplate_Location"  ' ** Module Function: modAutonumberFieldFuncs.

20340   Beep

EXITP:
20350   Setup_Location = blnRetVal
20360   Exit Function

ERRH:
20370   blnRetVal = False
20380   Select Case ERR.Number
        Case Else
20390     Beep
20400     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
20410   End Select
20420   Resume EXITP

End Function

Public Function Setup_m_REVCODE() As Boolean
' ** Called by:
' **   zz_mcr_Clean_RecurItems.

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "Setup_m_REVCODE"

        Dim blnRetVal As Boolean

20510   blnRetVal = True

        ' ** Empty tblTemplate_m_REVCODE of all but the first four records.
20520   TableEmpty "tblTemplate_m_REVCODE"  ' ** Module Function: modFileUtilities.

        ' ** Reset its AutoNumber field to 5.
20530   ChangeSeed_Ext "tblTemplate_m_REVCODE"  ' ** Module Function: modAutonumberFieldFuncs.

20540   Beep

EXITP:
20550   Setup_m_REVCODE = blnRetVal
20560   Exit Function

ERRH:
20570   blnRetVal = False
20580   Select Case ERR.Number
        Case Else
20590     Beep
20600     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
20610   End Select
20620   Resume EXITP

End Function

Public Function Setup_Recur() As Boolean
' ** Called by:
' **   zz_mcr_Clean_RecurItems.

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "Setup_Recur"

        Dim blnRetVal As Boolean

20710   blnRetVal = True

        ' ** Empty tblTemplate_RecurringItems of all but the first two records.
20720   TableEmpty "tblTemplate_RecurringItems"  ' ** Module Function: modFileUtilities.

        ' ** Reset its AutoNumber field to 3.
20730   ChangeSeed_Ext "tblTemplate_RecurringItems"  ' ** Module Function: modAutonumberFieldFuncs.

20740   Beep

EXITP:
20750   Setup_Recur = blnRetVal
20760   Exit Function

ERRH:
20770   blnRetVal = False
20780   Select Case ERR.Number
        Case Else
20790     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
20800   End Select
20810   Resume EXITP

End Function

Public Function Setup_SecUser() As Boolean
' ** Called by:
' **   zz_mcr_Clean_SecurityUser

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "Setup_SecUser"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnRetVal As Boolean

20910   blnRetVal = True

20920   Set dbs = CurrentDb
20930   With dbs
          ' ** Delete tblSecurity_User, where secusr_id > 5.
20940     Set qdf = .QueryDefs("qrySecurity_User_04")
20950     qdf.Execute
20960     .Close
20970   End With

        ' ** secusr_id  secusr_name secusr_default
        ' **     1      admin           -1
        ' **     2      Creator         -1
        ' **     3      Engine          -1
        ' **     4      superuser       -1
        ' **     5      TAAdmin         -1
        ' **     6      TADemo          -1
        ' **     7      Bill             0
        ' **     8      Mary             0

20980   ChangeSeed_Ext "tblSecurity_User"  ' ** Module Function: modAutonumberFieldFuncs.
20990   ChangeSeed_Ext "tblSecurity_GroupUser"  ' ** Module Function: modAutonumberFieldFuncs.

        ' ** Set Demo status.
21000   If Len(TA_SEC) > Len(TA_SEC2) Then
21010     Setup_Demo  ' ** Function: Below.
21020   End If

21030   Beep

EXITP:
21040   Set qdf = Nothing
21050   Set dbs = Nothing
21060   Setup_SecUser = blnRetVal
21070   Exit Function

ERRH:
21080   blnRetVal = False
21090   Select Case ERR.Number
        Case Else
21100     Beep
21110     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
21120   End Select
21130   Resume EXITP

End Function

Public Function Setup_VersionKey() As Boolean
' ** Called by:
' **   zz_mcr_Clean_RecurItems.

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "Setup_VersionKey"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnRetVal As Boolean

21210   blnRetVal = True

21220   Set dbs = CurrentDb
21230   With dbs
          ' ** Empty tblVersion_Conversion.
21240     TableEmpty "tblVersion_Conversion"  ' ** Module Function: modFileUtilities.
          ' ** Empty tblVersion_Key.
21250     TableEmpty "tblVersion_Key"  ' ** Module Function: modFileUtilities.
21260     .Close
21270   End With

        ' ** Reset their AutoNumber fields.
21280   ChangeSeed_Ext "tblVersion_Key"  ' ** Module Function: modAutonumberFieldFuncs.
21290   ChangeSeed_Ext "tblVersion_Conversion"  ' ** Module Function: modAutonumberFieldFuncs.

21300   Beep

EXITP:
21310   Set qdf = Nothing
21320   Set dbs = Nothing
21330   Setup_VersionKey = blnRetVal
21340   Exit Function

ERRH:
21350   blnRetVal = False
21360   Select Case ERR.Number
        Case Else
21370     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
21380   End Select
21390   Resume EXITP

End Function

Public Function VBA_PrefLoadSave() As Boolean
' ** Document all Pref_Load(), Pref_Save() calls to tblPreference_VBComponent.

21400 On Error GoTo ERRH

        Const THIS_PROC As String = "VBA_PrefLoadSave"

        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strModName As String, strFrmName As String, strProcName As String, strCodeLine As String
        Dim strLine As String
        Dim lngLines As Long, lngDecLines As Long
        Dim lngMods As Long, arr_varMod() As Variant
        Dim lngSubs As Long, arr_varSub() As Variant
        Dim lngThisDbsID As Long, lngVBComID As Long, lngModsAdded As Long
        Dim lngDels As Long, arr_varDel() As Variant
        Dim lngRecs As Long
        Dim blnFound As Boolean, blnAddAll As Boolean, blnAdd As Boolean, blnDelete As Boolean, blnExit As Boolean
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varMod().
        Const M_ELEMS As Integer = 15  ' ** Array's first-element UBound().
        Const M_DID1  As Integer = 0
        Const M_FID1  As Integer = 1
        Const M_FNAM  As Integer = 2
        Const M_VID1  As Integer = 3
        Const M_VNAM1 As Integer = 4
        Const M_PID1  As Integer = 5
        Const M_PNAM1 As Integer = 6
        Const M_LIN1  As Integer = 7
        Const M_COD1  As Integer = 8
        Const M_DID2  As Integer = 9
        Const M_VID2  As Integer = 10
        Const M_VNAM2 As Integer = 11
        Const M_PID2  As Integer = 12
        Const M_PNAM2 As Integer = 13
        Const M_LIN2  As Integer = 14
        Const M_COD2  As Integer = 15

        ' ** Array: arr_varDel().
        Const D_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const D_PVID As Integer = 0
        Const D_FID  As Integer = 1
        Const D_FNAM As Integer = 2
        Const D_PID  As Integer = 3
        Const D_PNAM As Integer = 4

        ' ** Array: arr_varSub().
        Const S_ELEMS As Integer = 9  ' ** Array's first-element UBound().
        Const S_PAR    As Integer = 0
        Const S_SUB1   As Integer = 1
        Const S_SUB2   As Integer = 2
        Const S_PELEM  As Integer = 3
        Const S_SELEM1 As Integer = 4
        Const S_SELEM2 As Integer = 5
        Const S_STR1   As Integer = 6
        Const S_STR2   As Integer = 7
        Const S_STR3   As Integer = 8
        Const S_LNG4   As Integer = 9

21410   blnRetVal = True

21420   lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

21430   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
21440   DoEvents

21450   lngMods = 0&
21460   ReDim arr_varMod(M_ELEMS, 0)

21470   lngSubs = 0&
21480   ReDim arr_varSub(S_ELEMS, 0)

        'TRANS AUDIT NOT WORKING RIGHT!

        ' ** Forms having Pref_Load() in the subform, and Pref_Save() in the parent.
21490   For lngZ = 1& To 6&
21500     lngSubs = lngSubs + 1&
21510     lngE = lngSubs - 1&
21520     ReDim Preserve arr_varSub(S_ELEMS, lngE)
21530     Select Case lngZ
          Case 1&
            ' ** Fee.
21540       arr_varSub(S_PAR, lngE) = "frmFeeSchedules"
21550       arr_varSub(S_SUB1, lngE) = "frmFeeSchedules_Sub"
21560       arr_varSub(S_SUB2, lngE) = "frmFeeSchedules_Detail_Sub"
21570     Case 2&
            ' ** Trans.
21580       arr_varSub(S_PAR, lngE) = "frmTransaction_Audit"
21590       arr_varSub(S_SUB1, lngE) = "frmTransaction_Audit_Sub"
21600       arr_varSub(S_SUB2, lngE) = "frmTransaction_Audit_Sub_ds"
21610     Case 3&
            ' ** Prof.
21620       arr_varSub(S_PAR, lngE) = "frmAccountProfile"
21630       arr_varSub(S_SUB1, lngE) = "frmAccountProfile_Sub"
21640       arr_varSub(S_SUB2, lngE) = vbNullString
21650     Case 4&
            ' ** ProfAdd.
21660       arr_varSub(S_PAR, lngE) = "frmAccountProfile_Add"
21670       arr_varSub(S_SUB1, lngE) = "frmAccountProfile_Add_Sub"
21680       arr_varSub(S_SUB2, lngE) = vbNullString
21690     Case 5&
            ' ** RevFreq.
21700       arr_varSub(S_PAR, lngE) = "frmAccountProfile_ReviewFreq"
21710       arr_varSub(S_SUB1, lngE) = "frmAccountProfile_ReviewFreq_Sub"
21720       arr_varSub(S_SUB2, lngE) = vbNullString
21730     Case 6&
            ' ** StmtFreq.
21740       arr_varSub(S_PAR, lngE) = "frmAccountProfile_StatementFreq"
21750       arr_varSub(S_SUB1, lngE) = "frmAccountProfile_StatementFreq_Sub"
21760       arr_varSub(S_SUB2, lngE) = vbNullString
21770     End Select
21780     arr_varSub(S_PELEM, lngE) = CLng(-1)
21790     arr_varSub(S_SELEM1, lngE) = CLng(-1)
21800     arr_varSub(S_SELEM2, lngE) = CLng(-1)
21810     arr_varSub(S_STR1, lngE) = vbNullString
21820     arr_varSub(S_STR2, lngE) = vbNullString
21830     arr_varSub(S_STR3, lngE) = vbNullString
21840     arr_varSub(S_LNG4, lngE) = vbNullString
21850   Next  ' ** lngZ.

21860   Debug.Print "'SUBS: " & CStr(lngSubs)
21870   DoEvents

21880   Set vbp = Application.VBE.ActiveVBProject
21890   With vbp
21900     strModName = vbNullString: strFrmName = vbNullString: strProcName = vbNullString: strCodeLine = vbNullString
21910     For Each vbc In .VBComponents
21920       With vbc

21930         strModName = .Name
21940         If Left(strModName, 6) <> "Report" And strModName <> "modPreferenceFuncs" And _
                  strModName <> "zz_mod_MDEPrepFuncs" Then
21950           If Left(strModName, 4) = "Form" Then  ' ** Might they be in a Standard Module?
21960             strFrmName = Mid(strModName, 6)
21970           End If

21980           Set cod = .CodeModule
21990           With cod
22000             lngLines = .CountOfLines
22010             lngDecLines = .CountOfDeclarationLines
22020             blnFound = False
22030             For lngX = lngDecLines To lngLines
22040               strLine = Trim(.Lines(lngX, 1))
22050               If strLine <> vbNullString Then
22060                 If Left(strLine, 1) <> "'" Then

22070                   intPos01 = InStr(strLine, "Pref_Load")
22080                   intPos02 = InStr(strLine, "Pref_Save")
22090                   intPos03 = InStr(strLine, "'")  ' ** The position of any remark also on the line.

                        ' ** frmFeeSchedules_Sub:                   Pref_Load()
                        ' **   frmFeeSchedules:                     Prev_Save()
                        ' ** frmFeeSchedules_Detail_Sub:            Pref_Load()
                        ' **   frmFeeSchedules:                     Pref_Save()

                        ' ** frmTransaction_Audit_Sub:              Pref_Load()
                        ' **   frmTransaction_Audit:                Pref_Save()
                        ' ** frmTransaction_Audit_Sub_ds:           Pref_Load()
                        ' **   frmTransaction_Audit:                Pref_Save()

                        ' ** frmAccountProfile_Sub:                 Pref_Load()
                        ' **   frmAccountProfile:                   Pref_Save()

                        ' ** frmAccountProfile_Add_Sub:             Pref_Load()
                        ' **   frmAccountProfile_Add:               Pref_Save()

                        ' ** frmAccountProfile_ReviewFreq           Pref_Load()
                        ' **   frmAccountProfile_ReviewFreq_Sub     Pref_Save()

                        ' ** frmAccountProfile_StatementFreq        Pref_Load()
                        ' **   frmAccountProfile_StatementFreq_Sub  Pref_Save()

                        'If strModName = "Form_frmTransaction_Audit_Sub" Or strModName = "Form_frmTransaction_Audit_Sub_ds" Then
                        'Stop
                        'End If

22100                   If intPos01 > 0 Then
22110                     If intPos03 = 0 Or intPos03 > intPos01 Then
22120                       If blnFound = False Then
22130                         blnFound = True
22140                         strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
22150                         intPos04 = InStr(strLine, " ")
22160                         If intPos04 > 0 Then  ' ** They should always be numbered.
22170                           strCodeLine = Trim(Left(strLine, intPos04))
22180                           If IsNumeric(strCodeLine) = False Then strCodeLine = vbNullString
22190                         End If
22200                         lngMods = lngMods + 1&
22210                         lngE = lngMods - 1&
22220                         ReDim Preserve arr_varMod(M_ELEMS, lngE)
                              ' ****************************************************
                              ' ** Array: arr_varMod()
                              ' **
                              ' **   Field  Element  Name               Constant
                              ' **   =====  =======  =================  ==========
                              ' **     1       0     dbs_id1            M_DID1
                              ' **     2       1     frm_id1            M_FID1
                              ' **     3       2     frm_name           M_FNAM
                              ' **     4       3     vbcom_id1          M_VID1
                              ' **     5       4     vbcom_name1        M_VNAM1
                              ' **     6       5     vbcomproc_id1      M_PID1
                              ' **     7       6     vbcomproc_name1    M_PNAM1
                              ' **     8       7     prefvbcom_line1    M_LIN1
                              ' **     9       8     prefvbcom_code1    M_COD1
                              ' **    10       9     dbs_id2            M_DID2
                              ' **    11      10     vbcom_id2          M_VID2
                              ' **    12      11     vbcom_name2        M_VNAM2
                              ' **    13      12     vbcomproc_id2      M_PID2
                              ' **    14      13     vbcomproc_name2    M_PNAM2
                              ' **    15      14     prefvbcom_line2    M_LIN2
                              ' **    16      15     prefvbcom_code2    M_COD2
                              ' **
                              ' ****************************************************
22230                         arr_varMod(M_DID1, lngE) = lngThisDbsID
22240                         arr_varMod(M_FID1, lngE) = CLng(0)
22250                         If strFrmName <> vbNullString Then
22260                           arr_varMod(M_FNAM, lngE) = strFrmName
22270                         Else
22280                           arr_varMod(M_FNAM, lngE) = Null
22290                         End If
22300                         arr_varMod(M_VID1, lngE) = CLng(0)
22310                         arr_varMod(M_VNAM1, lngE) = strModName
22320                         arr_varMod(M_PID1, lngE) = CLng(0)
22330                         arr_varMod(M_PNAM1, lngE) = strProcName
22340                         arr_varMod(M_LIN1, lngE) = lngX
22350                         If strCodeLine <> vbNullString Then
22360                           arr_varMod(M_COD1, lngE) = strCodeLine
22370                         Else
22380                           arr_varMod(M_COD1, lngE) = Null
22390                         End If
22400                         arr_varMod(M_DID2, lngE) = lngThisDbsID
22410                         arr_varMod(M_VID2, lngE) = CLng(0)
22420                         arr_varMod(M_VNAM2, lngE) = Null
22430                         arr_varMod(M_PID2, lngE) = CLng(0)
22440                         arr_varMod(M_PNAM2, lngE) = Null
22450                         arr_varMod(M_LIN2, lngE) = Null
22460                         arr_varMod(M_COD2, lngE) = Null
22470                         strProcName = vbNullString: strCodeLine = vbNullString

22480                         For lngZ = 0& To (lngSubs - 1&)
22490                           If strFrmName = arr_varSub(S_SUB1, lngZ) Or strFrmName = arr_varSub(S_SUB2, lngZ) Then
22500                             If arr_varSub(S_STR1, lngZ) <> vbNullString Then
22510                               arr_varMod(M_VNAM2, lngE) = arr_varSub(S_STR1, lngZ)
22520                               arr_varMod(M_PNAM2, lngE) = arr_varSub(S_STR2, lngZ)
22530                               arr_varMod(M_LIN2, lngE) = arr_varSub(S_LNG4, lngZ)
22540                               arr_varMod(M_COD2, lngE) = arr_varSub(S_STR3, lngZ)
22550                             Else
                                    ' ** Just wait till it's finished.
22560                             End If
22570                             Exit For
22580                           End If
22590                         Next  ' ** lngZ

22600                         If strModName = "Form_frmTransaction_Audit_Sub" Or strModName = "Form_frmTransaction_Audit_Sub_ds" Then
22610                           For lngZ = 0& To (lngSubs - 1&)
22620                             If strModName = "Form_" & arr_varSub(S_SUB1, lngZ) Or strModName = "Form_" & arr_varSub(S_SUB2, lngZ) Then
                                    ' ** These are the Pref_Save's.
22630                               arr_varSub(S_STR1, lngZ) = strModName
22640                               arr_varSub(S_STR2, lngZ) = .ProcOfLine(lngX, vbext_pk_Proc)
22650                               intPos04 = InStr(strLine, " ")
22660                               If intPos04 > 0 Then  ' ** They should always be numbered.
22670                                 arr_varSub(S_STR3, lngZ) = Trim(Left(strLine, intPos04))
22680                                 If IsNumeric(arr_varSub(S_STR2, lngZ)) = False Then strCodeLine = vbNullString
22690                               End If
22700                               arr_varSub(S_LNG4, lngZ) = lngX
22710                               blnExit = True
22720                               Exit For  ' ** This Exit's for lngZ.
22730                             End If
22740                           Next  ' lngZ.
22750                         End If

22760                       Else
22770                         If strFrmName <> "frmFeeSchedules" And strFrmName <> "frmTransaction_Audit" Then
                                ' ** 2 Pref_Load's?
22780                           Debug.Print "'2ND PREF_LOAD() FOUND!  " & strModName
22790                           DoEvents
22800                         End If
22810                       End If
22820                     End If  ' ** intPos03.
22830                   ElseIf intPos02 > 0 Then
22840                     If intPos03 = 0 Or intPos03 > intPos02 Then
22850                       If blnFound = True Then
22860                         strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
22870                         intPos04 = InStr(strLine, " ")
22880                         If intPos04 > 0 Then  ' ** They should always be numbered.
22890                           strCodeLine = Trim(Left(strLine, intPos04))
22900                           If IsNumeric(strCodeLine) = False Then strCodeLine = vbNullString
22910                         End If
22920                         arr_varMod(M_VID2, lngE) = CLng(0)
22930                         arr_varMod(M_VNAM2, lngE) = strModName
22940                         arr_varMod(M_PID2, lngE) = CLng(0)
22950                         arr_varMod(M_PNAM2, lngE) = strProcName
22960                         arr_varMod(M_LIN2, lngE) = lngX
22970                         If strCodeLine <> vbNullString Then
22980                           arr_varMod(M_COD2, lngE) = strCodeLine
22990                         Else
23000                           arr_varMod(M_COD2, lngE) = Null
23010                         End If
23020                         Exit For
23030                       Else
23040                         blnExit = False
23050                         For lngZ = 0& To (lngSubs - 1&)
23060                           If strModName = "Form_" & arr_varSub(S_PAR, lngZ) Then
23070                             For lngY = 0& To (lngMods - 1&)
23080                               If arr_varMod(M_FNAM, lngY) = arr_varSub(S_SUB1, lngZ) Then
23090                                 blnFound = True
23100                                 arr_varSub(S_SELEM1, lngZ) = lngY
23110                                 Exit For
23120                               End If
23130                             Next  ' ** lngY.
23140                             If blnFound = False Then
                                    ' ** We encountered the parent before its Sub's.
                                    ' ** Save data for the Sub's.
23150                             Else
23160                               If arr_varSub(S_SUB2, lngZ) <> vbNullString Then
23170                                 blnFound = False
23180                                 For lngY = 0& To (lngMods - 1&)
23190                                   If arr_varMod(M_FNAM, lngY) = arr_varSub(S_SUB2, lngZ) Then
23200                                     blnFound = True
23210                                     arr_varSub(S_SELEM2, lngZ) = lngY
23220                                     Exit For
23230                                   End If
23240                                 Next  ' ** lngY.
23250                                 If blnFound = False Then
                                        ' ** One before, one after?
                                        ' ** Save data for the Sub's.
23260                                 End If
23270                               End If
23280                             End If
                                  ' ** These are the Pref_Save's.
23290                             arr_varSub(S_STR1, lngZ) = strModName
23300                             arr_varSub(S_STR2, lngZ) = .ProcOfLine(lngX, vbext_pk_Proc)
23310                             intPos04 = InStr(strLine, " ")
23320                             If intPos04 > 0 Then  ' ** They should always be numbered.
23330                               arr_varSub(S_STR3, lngZ) = Trim(Left(strLine, intPos04))
23340                               If IsNumeric(arr_varSub(S_STR2, lngZ)) = False Then strCodeLine = vbNullString
23350                             End If
23360                             arr_varSub(S_LNG4, lngZ) = lngX
23370                             blnExit = True
23380                             Exit For  ' ** This Exit's for lngZ.
23390                           End If
23400                         Next  ' ** lngZ.
23410                         Select Case blnExit
                              Case True
23420                           Exit For
23430                         Case False
                                ' ** Either no Pref_Load(), or Pref_Save() found before Pref_Load()!
23440                           Debug.Print "'OUT OF ORDER!  " & strModName
23450                           DoEvents
23460                         End Select
23470                       End If  ' ** blnFound.
23480                     End If  ' ** intPos03.
23490                   End If  ' ** intPos01, intPos02.

23500                 End If  ' ** Remark.
23510               End If  ' ** vbNullString.
23520             Next  ' ** lngX.
23530           End With  ' ** cod.
23540           Set cod = Nothing

23550         End If  ' ** strModName.

23560       End With  ' ** vbc.
23570     Next  ' ** vbc.
23580   End With  ' ** vbp.
23590   Set cod = Nothing
23600   Set vbc = Nothing
23610   Set vbp = Nothing

23620   Debug.Print "'PREF_LOADS/SAVES: " & CStr(lngMods)
23630   DoEvents

        'For lngX = 0& To (lngSubs - 1&)
        '  Debug.Print "'PAR: " & arr_varSub(S_PAR, lngX) & "  SUB1: " & arr_varSub(S_SUB1, lngX) & "  " & CStr(arr_varSub(S_SELEM1, lngX)) & "  " & _
        '    "SUB2: " & arr_varSub(S_SUB2, lngX) & "  " & CStr(arr_varSub(S_SELEM2, lngX))
        '  Debug.Print "'  STR1: " & arr_varSub(S_STR1, lngX) & "  STR2: " & arr_varSub(S_STR2, lngX) & "  " & _
        '    "STR3: " & arr_varSub(S_STR3, lngX) & "  LNG4: " & CStr(arr_varSub(S_LNG4, lngX))
        'Next
        'SUBS: 6
        'PREF_LOADS/SAVES: 104
        'PAR: frmFeeSchedules  SUB1: frmFeeSchedules_Sub  -1  SUB2: frmFeeSchedules_Detail_Sub  -1
        '  STR1: Form_frmFeeSchedules  STR2: Form_Unload  STR3: 1510  LNG4: 239
        'PAR: frmTransaction_Audit  SUB1: frmTransaction_Audit_Sub  -1  SUB2: frmTransaction_Audit_Sub_ds  -1
        '  STR1: Form_frmTransaction_Audit_Sub_ds  STR2: Form_Load  STR3: 410  LNG4: 129
        'PAR: frmAccountProfile  SUB1: frmAccountProfile_Sub  -1  SUB2:   -1
        '  STR1: Form_frmAccountProfile  STR2: Form_Unload  STR3: 4420  LNG4: 622
        'PAR: frmAccountProfile_Add  SUB1: frmAccountProfile_Add_Sub  -1  SUB2:   -1
        '  STR1: Form_frmAccountProfile_Add  STR2: Form_Unload  STR3: 2920  LNG4: 428
        'PAR: frmAccountProfile_ReviewFreq  SUB1: frmAccountProfile_ReviewFreq_Sub  -1  SUB2:   -1
        '  STR1: Form_frmAccountProfile_ReviewFreq  STR2: Form_Unload  STR3: 1010  LNG4: 151
        'PAR: frmAccountProfile_StatementFreq  SUB1: frmAccountProfile_StatementFreq_Sub  -1  SUB2:   -1
        '  STR1: Form_frmAccountProfile_StatementFreq  STR2: Form_Unload  STR3: 1010  LNG4: 153

        ' ** Take a breather!
        'Stop

23640   If lngMods > 0& Then

23650     Set dbs = CurrentDb
23660     With dbs

            ' ** Check on frmFeeSchedules.
            ' ** frmFeeSchedules_Sub:                  Pref_Load()
            ' **   frmFeeSchedules:                    Prev_Save()
            ' ** frmFeeSchedules_Detail_Sub:           Pref_Load()
            ' **   frmFeeSchedules:                    Pref_Save()
            ' ** Check on frmTransaction_Audit.
            ' ** frmTransaction_Audit_Sub:             Pref_Load()
            ' **   frmTransaction_Audit:               Pref_Save()
            ' ** frmTransaction_Audit_Sub_ds:          Pref_Load()
            ' **   frmTransaction_Audit:               Pref_Save()
            ' ** Check on frmAccountProfile.
            ' ** frmAccountProfile_Sub:                Pref_Load()
            ' **   frmAccountProfile:                  Pref_Save()
            ' ** Check on frmAccountProfile_Add.
            ' ** frmAccountProfile_Add_Sub:            Pref_Load()
            ' **   frmAccountProfile_Add:              Pref_Save()
            ' ** Check on frmAccountProfile_ReviewFreq.
            ' ** frmAccountProfile_ReviewFreq_Sub:     Pref_Load()
            ' **   frmAccountProfile_ReviewFreq:       Pref_Save()
            ' ** Check on frmAccountProfile_StatementFreq.
            ' ** frmAccountProfile_StatementFreq_Sub:  Pref_Load()
            ' **   frmAccountProfile_StatementFreq:    Pref_Save()

23670       For lngZ = 0& To (lngSubs - 1&)

23680         If arr_varSub(S_PAR, lngZ) <> "frmTransaction_Audit" Then
23690           ' ** First, there should be no entry for the parent.
23700           For lngX = 0& To (lngMods - 1&)
23710             If arr_varMod(M_FNAM, lngX) = arr_varSub(S_PAR, lngZ) Then
23720               ' ** Don't save this!
23730               arr_varSub(S_PELEM, lngZ) = lngX
23740             End If
23750           Next  ' ** lngX.
23760         End If

              'If arr_varSub(S_PAR, lngZ) = "frmTransaction_Audit" Then
              'Stop
              'End If

23770         ' ** Next, find the 1st Sub.
23780         If arr_varSub(S_SELEM1, lngZ) = -1& Then
23790           For lngX = 0& To (lngMods - 1&)
23800             If arr_varMod(M_FNAM, lngX) = arr_varSub(S_SUB1, lngZ) Then
23810               arr_varSub(S_SELEM1, lngZ) = lngX
23820               Exit For
23830             End If
23840           Next  ' ** lngX.
23850         End If
23860         If arr_varSub(S_SELEM1, lngZ) >= 0& Then
23870           If IsNull(arr_varMod(M_VNAM2, arr_varSub(S_SELEM1, lngZ))) = True Then
23880             If arr_varSub(S_STR1, lngZ) <> vbNullString Then
23890               arr_varMod(M_VNAM2, arr_varSub(S_SELEM1, lngZ)) = arr_varSub(S_STR1, lngZ)
23900               arr_varMod(M_PNAM2, arr_varSub(S_SELEM1, lngZ)) = arr_varSub(S_STR2, lngZ)
23910               arr_varMod(M_LIN2, arr_varSub(S_SELEM1, lngZ)) = arr_varSub(S_LNG4, lngZ)
23920               arr_varMod(M_COD2, arr_varSub(S_SELEM1, lngZ)) = arr_varSub(S_STR3, lngZ)
23930             End If
23940           End If
23950         End If

23960         If arr_varSub(S_SUB2, lngZ) <> vbNullString Then
23970           ' ** And finally, find the 2nd Sub.
23980           If arr_varSub(S_SELEM2, lngZ) = -1& Then
23990             For lngX = 0& To (lngMods - 1&)
24000               If arr_varMod(M_FNAM, lngX) = arr_varSub(S_SUB2, lngZ) Then
24010                 arr_varSub(S_SELEM2, lngZ) = lngX
24020                 Exit For
24030               End If
24040             Next  ' ** lngX.
24050           End If
24060           If arr_varSub(S_SELEM2, lngZ) >= 0& Then
24070             If IsNull(arr_varMod(M_VNAM2, arr_varSub(S_SELEM2, lngZ))) = True Then
24080               If arr_varSub(S_STR1, lngZ) <> vbNullString Then
24090                 arr_varMod(M_VNAM2, arr_varSub(S_SELEM2, lngZ)) = arr_varSub(S_STR1, lngZ)
24100                 arr_varMod(M_PNAM2, arr_varSub(S_SELEM2, lngZ)) = arr_varSub(S_STR2, lngZ)
24110                 arr_varMod(M_LIN2, arr_varSub(S_SELEM2, lngZ)) = arr_varSub(S_LNG4, lngZ) + 1&
24120                 arr_varMod(M_COD2, arr_varSub(S_SELEM2, lngZ)) = CStr(Val(arr_varSub(S_STR3, lngZ)) + 10)
24130               End If
24140             End If
24150           End If
24160         End If

24170       Next  ' ** lngZ.

24180       For lngX = 0& To (lngMods - 1&)
24190         If IsNull(arr_varMod(M_VNAM1, lngX)) = True Then
24200           Stop
24210         ElseIf IsNull(arr_varMod(M_VNAM2, lngX)) = True Then
24220           Stop
24230         Else
24240           If arr_varMod(M_VNAM1, lngX) = vbNullString Then
24250             Stop
24260           ElseIf arr_varMod(M_VNAM2, lngX) = vbNullString Then
24270             Stop
24280           End If
24290         End If
24300       Next

            ' ** Take another breather!
            'Stop

24310       For lngX = 0& To (lngMods - 1&)
24320         lngVBComID = 0&
24330         If IsNull(arr_varMod(M_FNAM, lngX)) = False Then
24340           Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
24350           With rst
24360             .FindFirst "[dbs_id] = " & CStr(arr_varMod(M_DID1, lngX)) & " And [frm_name] = '" & arr_varMod(M_FNAM, lngX) & "'"
24370             If .NoMatch = False Then
24380               arr_varMod(M_FID1, lngX) = ![frm_id]
24390             Else
24400               blnRetVal = False
24410               Debug.Print "'FORM NOT FOUND!  " & arr_varMod(M_FNAM, lngX)
24420               DoEvents
24430             End If
24440             .Close
24450           End With  ' ** rst.
24460           Set rst = Nothing
24470         End If  ' ** strFrmName.
24480         Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
24490         With rst
24500           .FindFirst "[dbs_id] = " & CStr(arr_varMod(M_DID1, lngX)) & " And [vbcom_name] = '" & arr_varMod(M_VNAM1, lngX) & "'"
24510           If .NoMatch = False Then
24520             lngVBComID = ![vbcom_id]
24530             arr_varMod(M_VID1, lngX) = lngVBComID
24540           Else
24550             blnRetVal = False
24560             Debug.Print "'MOD NOT FOUND!  " & arr_varMod(M_VNAM1, lngX)
24570             DoEvents
                  Stop
24580           End If
24590           .Close
24600         End With  ' ** rst.
24610         Set rst = Nothing
24620         If lngVBComID > 0& Then
24630           Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
24640           With rst
24650             .FindFirst "[dbs_id] = " & CStr(arr_varMod(M_DID1, lngX)) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
                    "[vbcomproc_name] = '" & arr_varMod(M_PNAM1, lngX) & "'"
24660             If .NoMatch = False Then
24670               arr_varMod(M_PID1, lngX) = ![vbcomproc_id]
24680             Else
24690               blnRetVal = False
24700               Debug.Print "'PROC NOT FOUND!  " & arr_varMod(M_PNAM1, lngX) & "  IN  " & arr_varMod(M_VNAM1, lngX)
24710               DoEvents
24720             End If
24730           End With  ' ** rst.
24740           Set rst = Nothing
24750           lngVBComID = 0&
24760           Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
24770           With rst
24780             .FindFirst "[dbs_id] = " & CStr(arr_varMod(M_DID2, lngX)) & " And [vbcom_name] = '" & arr_varMod(M_VNAM2, lngX) & "'"
24790             If .NoMatch = False Then
24800               lngVBComID = ![vbcom_id]
24810               arr_varMod(M_VID2, lngX) = lngVBComID
24820             Else
24830               blnRetVal = False
24840               Debug.Print "'MOD NOT FOUND!  " & arr_varMod(M_VNAM2, lngX)
24850               DoEvents
                    Stop
24860             End If
24870             .Close
24880           End With  ' ** rst.
24890           Set rst = Nothing
24900           If lngVBComID > 0& Then
24910             Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
24920             With rst
24930               .FindFirst "[dbs_id] = " & CStr(arr_varMod(M_DID2, lngX)) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
                      "[vbcomproc_name] = '" & arr_varMod(M_PNAM2, lngX) & "'"
24940               If .NoMatch = False Then
24950                 arr_varMod(M_PID2, lngX) = ![vbcomproc_id]
24960               Else
24970                 blnRetVal = False
24980                 Debug.Print "'PROC NOT FOUND!  " & arr_varMod(M_PNAM2, lngX) & "  IN  " & arr_varMod(M_VNAM2, lngX)
24990                 DoEvents
25000               End If
25010             End With  ' ** rst.
25020             Set rst = Nothing
25030           End If  ' ** lngVBComID.
25040         End If  ' ** lngVBComID.
25050         If blnRetVal = False Then
25060           Exit For
25070         End If
25080       Next  ' ** lngX.

25090       If blnRetVal = True Then

25100         lngModsAdded = 0&
25110         Set rst = .OpenRecordset("tblPreference_VBComponent", dbOpenDynaset, dbConsistent)
25120         With rst
25130           blnAddAll = False
25140           If .BOF = True And .EOF = True Then
25150             blnAddAll = True
25160           Else
25170             .MoveFirst
25180           End If
25190           For lngX = 0& To (lngMods - 1&)
25200             blnAdd = False
25210             Select Case blnAddAll
                  Case True
25220               blnAdd = True
25230             Case False
25240               .MoveFirst
25250               .FindFirst "[dbs_id1] = " & CStr(arr_varMod(M_DID1, lngX)) & " And " & _
                      "[frm_id1] = " & CStr(arr_varMod(M_FID1, lngX)) & " And " & _
                      "[vbcom_id1] = " & CStr(arr_varMod(M_VID1, lngX))
25260               Select Case .NoMatch
                    Case True
25270                 blnAdd = True
25280               Case False
25290                 .Edit
25300               End Select
25310             End Select
25320             If blnAdd = True Then
25330               .AddNew
                    ' ** ![prefvbcom_id] : AutoNumber.
25340               ![dbs_id1] = arr_varMod(M_DID1, lngX)
25350               ![frm_id1] = arr_varMod(M_FID1, lngX)
25360               ![vbcom_id1] = arr_varMod(M_VID1, lngX)
25370               lngModsAdded = lngModsAdded + 1&
25380             End If
25390             ![vbcomproc_id1] = arr_varMod(M_PID1, lngX)
25400             ![prefvbcom_line1] = arr_varMod(M_LIN1, lngX)
25410             ![prefvbcom_code1] = arr_varMod(M_COD1, lngX)
25420             ![dbs_id2] = arr_varMod(M_DID2, lngX)
25430             ![vbcom_id2] = arr_varMod(M_VID2, lngX)
25440             ![vbcomproc_id2] = arr_varMod(M_PID2, lngX)
25450             ![prefvbcom_line2] = arr_varMod(M_LIN2, lngX)
25460             ![prefvbcom_code2] = arr_varMod(M_COD2, lngX)
25470             ![prefvbcom_datemodified] = Now()
25480             .Update
25490           Next  ' ** lngX.
25500           .Close
25510         End With  ' ** rst.
25520         Set rst = Nothing

25530         lngDels = 0&
25540         ReDim arr_varDel(D_ELEMS, 0)

              ' ** Check for obsolete table entries.
25550         Set rst = .OpenRecordset("tblPreference_VBComponent", dbOpenDynaset, dbReadOnly)
25560         With rst
25570           .MoveLast
25580           lngRecs = .RecordCount
25590           .MoveFirst
25600           For lngX = 1& To lngRecs
25610             If ![dbs_id1] = lngThisDbsID Then
25620               blnFound = False
25630               For lngY = 0& To (lngMods - 1&)
25640                 If arr_varMod(M_DID1, lngY) = ![dbs_id1] And arr_varMod(M_FID1, lngY) = ![frm_id1] And _
                          arr_varMod(M_VID1, lngY) = ![vbcom_id1] And arr_varMod(M_LIN1, lngY) = ![prefvbcom_line1] And _
                          arr_varMod(M_LIN2, lngY) = ![prefvbcom_line2] Then
25650                   blnFound = True
25660                   Exit For
25670                 End If
25680               Next  ' ** lngY.
25690               If blnFound = False Then
25700                 lngDels = lngDels + 1&
25710                 lngE = lngDels - 1&
25720                 ReDim Preserve arr_varDel(D_ELEMS, lngE)
25730                 arr_varDel(D_PVID, lngE) = ![prefvbcom_id]
25740                 arr_varDel(D_FID, lngE) = ![frm_id1]
25750                 arr_varDel(D_FNAM, lngE) = Null
25760                 arr_varDel(D_PID, lngE) = ![vbcomproc_id1]
25770                 arr_varDel(D_PNAM, lngE) = Null
25780               End If  ' ** blnFound.
25790             End If  ' ** lngThisDbsID.
25800             If lngX < lngRecs Then .MoveNext
25810           Next  ' ** lngX.
25820           .Close
25830         End With  ' ** rst.
25840         Set rst = Nothing

25850         Debug.Print "'DELS: " & CStr(lngDels)
25860         DoEvents

25870         If lngDels > 0& Then

25880           For lngX = 0& To (lngDels - 1&)
25890             varTmp00 = DLookup("[frm_name]", "tblForm", "[frm_id] = " & CStr(arr_varDel(D_FID, lngX)))
25900             If IsNull(varTmp00) = False Then
25910               arr_varDel(D_FNAM, lngX) = varTmp00
25920             Else
25930               Stop
25940             End If
25950             varTmp00 = DLookup("[vbcomproc_name]", "tblVBComponent_Procedure", "[vbcomproc_id] = " & CStr(arr_varDel(D_PID, lngX)))
25960             If IsNull(varTmp00) = False Then
25970               arr_varDel(D_PNAM, lngX) = varTmp00
25980             Else
25990               Stop
26000             End If
26010           Next  ' ** lngX.

26020           For lngX = 0& To (lngDels - 1&)
26030             blnDelete = True
26040             Debug.Print "'DEL?  " & arr_varDel(D_FNAM, lngX) & "  " & arr_varDel(D_PNAM, lngX) & "()?"
26050             Stop
26060             If blnDelete = True Then
                    ' ** Delete tblPreference_VBComponent, by specified [pvid].
26070               Set qdf = .QueryDefs("qryPreferences_12")
26080               With qdf.Parameters
26090                 ![pvid] = arr_varDel(D_PVID, lngX)
26100               End With
26110               qdf.Execute
26120               Set qdf = Nothing
26130             End If
26140           Next  ' ** lngX.

26150         End If  ' ** lngDels.

26160       End If  ' ** blnRetVal.

26170       .Close
26180     End With

26190   End If  ' ** lngMods.

26200   Debug.Print "'MODS ADDED: " & CStr(lngModsAdded)
26210   DoEvents

26220   Beep
26230   Debug.Print "'DONE!"
26240   DoEvents

EXITP:
26250   Set cod = Nothing
26260   Set vbc = Nothing
26270   Set vbp = Nothing
26280   Set rst = Nothing
26290   Set qdf = Nothing
26300   Set dbs = Nothing
26310   VBA_PrefLoadSave = blnRetVal
26320   Exit Function

ERRH:
26330   blnRetVal = False
26340   Select Case ERR.Number
        Case Else
26350     Beep
26360     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
26370   End Select
26380   Resume EXITP

End Function

Public Function EmptyTable_Check() As Boolean
' ** Check for new temp and staging tables.
' **
' ** Empty Tmp Queries:
' **   1. Tmps are 1st.
' **   2. Non-tbl-prefixed tables.
' **   2. Followed by Templates.
' **   3. Then Staging and regular named tables.
' **   4. Then some special Tmps.
' **   5. tmpXAdmins.
' **   6. CheckReconcile.

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTable_Check"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, qdf As DAO.QueryDef
        Dim lngTbls As Long, arr_varTbl() As Variant
        Dim strTblName As String, strQryName As String
        Dim blnFound As Boolean
        Dim intPos01 As Integer
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varTbl().
        Const T_ELEMS As Integer = 7  ' ** Array's first-element UBound().
        Const T_TNAM  As Integer = 0
        Const T_STG   As Integer = 1
        Const T_TMPLT As Integer = 2
        Const T_BAK   As Integer = 3
        Const T_TMP   As Integer = 4
        Const T_ZZ    As Integer = 5
        Const T_FND   As Integer = 6
        Const T_QRY   As Integer = 7

        Const TBL_EMPTY As String = "qryTmp_Table_Empty_"

26410   blnRetVal = True

'TBLS: 229
'STAGING QRY NOT FOUND: tblReport_List_Staging_bak
'STAGING QRY NOT FOUND: tblReport_List_Staging_tmp

'TMP QRY NOT FOUND: tmpXAdmin_m_REVCODE
'TMP QRY NOT FOUND: tmpXAdmin_MasterAsset_02

'TEMPLATES W/O QRYS: tblTemplate_AssetType
'TEMPLATES W/O QRYS: tblTemplate_Database
'TEMPLATES W/O QRYS: tblTemplate_Database_Table
'TEMPLATES W/O QRYS: tblTemplate_Database_Table_Link
'TEMPLATES W/O QRYS: tblTemplate_Form
'TEMPLATES W/O QRYS: tblTemplate_Form_Control
'TEMPLATES W/O QRYS: tblTemplate_Form_Graphics
'TEMPLATES W/O QRYS: tblTemplate_Form_Graphics_PictureData
'TEMPLATES W/O QRYS: tblTemplate_Form_Graphics_PictureData2
'TEMPLATES W/O QRYS: tblTemplate_HiddenType
'TEMPLATES W/O QRYS: tblTemplate_InvestmentObjective
'TEMPLATES W/O QRYS: tblTemplate_Journal_Field
'TEMPLATES W/O QRYS: tblTemplate_m_REVCODE_TYPE
'TEMPLATES W/O QRYS: tblTemplate_m_TBL
'TEMPLATES W/O QRYS: tblTemplate_Preference_Control
'TEMPLATES W/O QRYS: tblTemplate_RecurringType
'TEMPLATES W/O QRYS: tblTemplate_TaxCode
'TEMPLATES W/O QRYS: tblTemplate_TaxCode_Old
'TEMPLATES W/O QRYS: tblTemplate_USysRibbons
'TEMPLATES W/O QRYS: tblTemplate_Zeta1
'TEMPLATES W/O QRYS: tblTemplate_Zeta2
'TEMPLATES W/O QRYS: tblTemplate_Zeta3
'TEMPLATES W/O QRYS: tblTemplate_Zeta4
'TEMPLATES W/O QRYS: tblTemplate_Zeta5
'TEMPLATES W/O QRYS: tblTemplate_Zeta6
'TEMPLATES W/O QRYS: tblTemplate_Zeta7
'TEMPLATES W/O QRYS: tblTemplate_Zeta8
'TEMPLATES W/O QRYS: tblTemplate_Zeta9

'BAK TBLS:
'  tblImportExport_Columns_bak
'  tblImportExport_Specifications_bak
'  tblPricing_MasterAsset_History_bak
'  tblRelation_bak
'  tblRelation_Field_bak
'  tblReport_List_bak
'  tblReport_List_Control_bak
'  tblReport_List_Report_Alt_bak
'  tblReport_List_Report_Alt_Control_bak
'  tblReport_List_Report_bak
'  tblReport_List_Sort_bak
'  tblVersion_Conversion_bak

26420   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
26430   DoEvents

26440   lngTbls = 0&
26450   ReDim arr_varTbl(T_ELEMS, 0)

26460   Set dbs = CurrentDb
26470   With dbs

26480     For Each tdf In .TableDefs
26490       With tdf
26500         If InStr(.Name, "staging") > 0 Then
26510           lngTbls = lngTbls + 1&
26520           lngE = lngTbls - 1&
26530           ReDim Preserve arr_varTbl(T_ELEMS, lngE)
26540           arr_varTbl(T_TNAM, lngE) = .Name
26550           arr_varTbl(T_STG, lngE) = CBool(True)
26560           arr_varTbl(T_TMPLT, lngE) = CBool(False)
26570           arr_varTbl(T_BAK, lngE) = CBool(False)
26580           arr_varTbl(T_TMP, lngE) = CBool(False)
26590           arr_varTbl(T_ZZ, lngE) = CBool(False)
26600           arr_varTbl(T_FND, lngE) = CBool(False)
26610           arr_varTbl(T_QRY, lngE) = Null
26620         ElseIf Left(.Name, 11) = "tblTemplate" Then
26630           lngTbls = lngTbls + 1&
26640           lngE = lngTbls - 1&
26650           ReDim Preserve arr_varTbl(T_ELEMS, lngE)
26660           arr_varTbl(T_TNAM, lngE) = .Name
26670           arr_varTbl(T_STG, lngE) = CBool(False)
26680           arr_varTbl(T_TMPLT, lngE) = CBool(True)
26690           arr_varTbl(T_BAK, lngE) = CBool(False)
26700           arr_varTbl(T_TMP, lngE) = CBool(False)
26710           arr_varTbl(T_ZZ, lngE) = CBool(False)
26720           arr_varTbl(T_FND, lngE) = CBool(False)
26730           arr_varTbl(T_QRY, lngE) = Null
26740         ElseIf InStr(.Name, "_bak") > 0 Then
26750           lngTbls = lngTbls + 1&
26760           lngE = lngTbls - 1&
26770           ReDim Preserve arr_varTbl(T_ELEMS, lngE)
26780           arr_varTbl(T_TNAM, lngE) = .Name
26790           arr_varTbl(T_STG, lngE) = CBool(False)
26800           arr_varTbl(T_TMPLT, lngE) = CBool(False)
26810           arr_varTbl(T_BAK, lngE) = CBool(True)
26820           arr_varTbl(T_TMP, lngE) = CBool(False)
26830           arr_varTbl(T_ZZ, lngE) = CBool(False)
26840           arr_varTbl(T_FND, lngE) = CBool(False)
26850           arr_varTbl(T_QRY, lngE) = Null
26860         ElseIf Left(.Name, 3) = "tmp" Then
26870           lngTbls = lngTbls + 1&
26880           lngE = lngTbls - 1&
26890           ReDim Preserve arr_varTbl(T_ELEMS, lngE)
26900           arr_varTbl(T_TNAM, lngE) = .Name
26910           arr_varTbl(T_STG, lngE) = CBool(False)
26920           arr_varTbl(T_TMPLT, lngE) = CBool(False)
26930           arr_varTbl(T_BAK, lngE) = CBool(False)
26940           arr_varTbl(T_TMP, lngE) = CBool(True)
26950           arr_varTbl(T_ZZ, lngE) = CBool(False)
26960           arr_varTbl(T_FND, lngE) = CBool(False)
26970           arr_varTbl(T_QRY, lngE) = Null
26980         ElseIf Left(.Name, 2) = "zz" Then
26990           lngTbls = lngTbls + 1&
27000           lngE = lngTbls - 1&
27010           ReDim Preserve arr_varTbl(T_ELEMS, lngE)
27020           arr_varTbl(T_TNAM, lngE) = .Name
27030           arr_varTbl(T_STG, lngE) = CBool(False)
27040           arr_varTbl(T_TMPLT, lngE) = CBool(False)
27050           arr_varTbl(T_BAK, lngE) = CBool(False)
27060           arr_varTbl(T_TMP, lngE) = CBool(False)
27070           arr_varTbl(T_ZZ, lngE) = CBool(True)
27080           arr_varTbl(T_FND, lngE) = CBool(False)
27090           arr_varTbl(T_QRY, lngE) = Null
27100         ElseIf InStr(.Name, "tmp") > 0 Then
27110           lngTbls = lngTbls + 1&
27120           lngE = lngTbls - 1&
27130           ReDim Preserve arr_varTbl(T_ELEMS, lngE)
27140           arr_varTbl(T_TNAM, lngE) = .Name
27150           arr_varTbl(T_STG, lngE) = CBool(False)
27160           arr_varTbl(T_TMPLT, lngE) = CBool(False)
27170           arr_varTbl(T_BAK, lngE) = CBool(False)
27180           arr_varTbl(T_TMP, lngE) = CBool(True)
27190           arr_varTbl(T_ZZ, lngE) = CBool(False)
27200           arr_varTbl(T_FND, lngE) = CBool(False)
27210           arr_varTbl(T_QRY, lngE) = Null
27220         End If
27230       End With  ' ** tdf.
27240     Next  ' ** tdf.
27250     Set tdf = Nothing

27260     Debug.Print "'TBLS: " & CStr(lngTbls)
27270     DoEvents

27280     If lngTbls > 0& Then

27290       For Each qdf In .QueryDefs
27300         With qdf
27310           If Left(.Name, Len(TBL_EMPTY)) = TBL_EMPTY Then
27320             strQryName = .Name
27330             intPos01 = CharPos(strQryName, 4, "_")  ' ** Module Function: modStringFuncs.
27340             strTblName = Mid(strQryName, (intPos01 + 1))
27350             For lngX = 0& To (lngTbls - 1&)
27360               If arr_varTbl(T_TNAM, lngX) = strTblName Then
27370                 arr_varTbl(T_FND, lngX) = True
27380                 arr_varTbl(T_QRY, lngX) = strQryName
27390                 Exit For
27400               End If
27410             Next  ' ** lngX.
27420           End If
27430         End With  ' ** qdf.
27440       Next  ' ** qdf.
27450       DoEvents

            ' ** Staging tables.
27460       blnFound = False
27470       For lngX = 0& To (lngTbls - 1&)
27480         If arr_varTbl(T_STG, lngX) = True Then
27490           If arr_varTbl(T_FND, lngX) = False Then
27500             Debug.Print "'STAGING QRY NOT FOUND: " & arr_varTbl(T_TNAM, lngX)
27510             DoEvents
27520             blnFound = True
27530           End If
27540         End If
27550       Next  ' ** lngX.
27560       If blnFound = True Then
27570         Debug.Print
27580         DoEvents
27590       End If

            ' ** Temp tables.
27600       blnFound = False
27610       For lngX = 0& To (lngTbls - 1&)
27620         If arr_varTbl(T_TMP, lngX) = True Then
27630           If arr_varTbl(T_FND, lngX) = False Then
27640             Debug.Print "'TMP QRY NOT FOUND: " & arr_varTbl(T_TNAM, lngX)
27650             DoEvents
27660             blnFound = True
27670           End If
27680         End If
27690       Next  ' ** lngX.
27700       If blnFound = True Then
27710         Debug.Print
27720         DoEvents
27730       End If

            ' ** Template tables.
27740       blnFound = False
27750       For lngX = 0& To (lngTbls - 1&)
27760         If arr_varTbl(T_TMPLT, lngX) = True Then
27770           If arr_varTbl(T_FND, lngX) = False Then
27780             Debug.Print "'TEMPLATES W/O QRYS: " & arr_varTbl(T_TNAM, lngX)
27790             DoEvents
27800             blnFound = True
27810           End If
27820         End If
27830       Next  ' ** lngX.
27840       If blnFound = True Then
27850         Debug.Print
27860         DoEvents
27870       End If

            ' ** Bak tables.
27880       Debug.Print "'BAK TBLS:"
27890       DoEvents
27900       For lngX = 0& To (lngTbls - 1&)
27910         If arr_varTbl(T_BAK, lngX) = True Then
27920           Debug.Print "'  " & arr_varTbl(T_TNAM, lngX)
27930           DoEvents
27940         End If
27950       Next  ' ** lngX.

27960     End If  ' ** lngTbls.

27970     .Close
27980   End With  ' ** dbs.
27990   Set dbs = Nothing

        'TBLS: 231
        'STAGING QRY NOT FOUND: tblReport_List_Staging_bak
        'STAGING QRY NOT FOUND: tblReport_List_Staging_tmp

        'TMP QRY NOT FOUND: tmpXAdmin_m_REVCODE
        'TMP QRY NOT FOUND: tmpXAdmin_MasterAsset_02

        'TEMPLATES W/O QRYS: tblTemplate_AssetType
        'TEMPLATES W/O QRYS: tblTemplate_Database
        'TEMPLATES W/O QRYS: tblTemplate_Database_Table
        'TEMPLATES W/O QRYS: tblTemplate_Database_Table_Link
        'TEMPLATES W/O QRYS: tblTemplate_Form
        'TEMPLATES W/O QRYS: tblTemplate_Form_Control
        'TEMPLATES W/O QRYS: tblTemplate_Form_Graphics
        'TEMPLATES W/O QRYS: tblTemplate_Form_Graphics_PictureData
        'TEMPLATES W/O QRYS: tblTemplate_HiddenType
        'TEMPLATES W/O QRYS: tblTemplate_InvestmentObjective
        'TEMPLATES W/O QRYS: tblTemplate_Journal_Field
        'TEMPLATES W/O QRYS: tblTemplate_m_REVCODE_TYPE
        'TEMPLATES W/O QRYS: tblTemplate_m_TBL
        'TEMPLATES W/O QRYS: tblTemplate_Preference_Control
        'TEMPLATES W/O QRYS: tblTemplate_RecurringType
        'TEMPLATES W/O QRYS: tblTemplate_TaxCode
        'TEMPLATES W/O QRYS: tblTemplate_TaxCode_Old
        'TEMPLATES W/O QRYS: tblTemplate_USysRibbons
        'TEMPLATES W/O QRYS: tblTemplate_Zeta1
        'TEMPLATES W/O QRYS: tblTemplate_Zeta2
        'TEMPLATES W/O QRYS: tblTemplate_Zeta3
        'TEMPLATES W/O QRYS: tblTemplate_Zeta4
        'TEMPLATES W/O QRYS: tblTemplate_Zeta5
        'TEMPLATES W/O QRYS: tblTemplate_Zeta6
        'TEMPLATES W/O QRYS: tblTemplate_Zeta7
        'TEMPLATES W/O QRYS: tblTemplate_Zeta8
        'TEMPLATES W/O QRYS: tblTemplate_Zeta9

        'BAK TBLS:
        '  tblImportExport_Columns_bak
        '  tblImportExport_Specifications_bak
        '  tblPricing_MasterAsset_History_bak
        '  tblRelation_bak
        '  tblRelation_Field_bak
        '  tblReport_List_bak
        '  tblReport_List_Control_bak
        '  tblReport_List_Report_Alt_bak
        '  tblReport_List_Report_Alt_Control_bak
        '  tblReport_List_Report_bak
        '  tblReport_List_Sort_bak
        '  tblReport_VBComponent_bak
        '  tblSecurity_GroupUser_bak
        '  tblVersion_Conversion_bak
        'DONE!

28000   Beep

28010   Debug.Print "'DONE!"
28020   DoEvents

EXITP:
28030   Set qdf = Nothing
28040   Set tdf = Nothing
28050   Set dbs = Nothing
28060   EmptyTable_Check = blnRetVal
28070   Exit Function

ERRH:
28080   blnRetVal = False
28090   Select Case ERR.Number
        Case Else
28100     Beep
28110     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
28120   End Select
28130   Resume EXITP

End Function

Public Function Frm_PopUpChk() As Boolean
' ** Check report forms to make sure they're not PopUp.

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "Frm_PopUpChk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim lngFrms As Long, arr_varFrm As Variant
        Dim strMsg As String
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varFrm().
        'Const F_DID  As Integer = 0  'dbs_id
        'Const F_FID  As Integer = 1  'frm_id
        Const F_FNAM As Integer = 2  'frm_name
        Const F_POP  As Integer = 3  'IsPopUp

28210   blnRetVal = True

28220   If Parse_Ext(CurrentAppName) = gstrExt_AppDev Then  ' ** Module Functions: modFileUtilities.
          ' ** Errors with an MDE.
28230     Set dbs = CurrentDb
28240     With dbs
            ' ** tblForm_Non_PopUp.
28250       Set qdf = .QueryDefs("qryForm_Non_PopUp_05")
28260       Set rst = qdf.OpenRecordset
28270       With rst
28280         .MoveLast
28290         lngFrms = .RecordCount
28300         .MoveFirst
28310         arr_varFrm = .GetRows(lngFrms)
              ' **********************************************
              ' ** Array: arr_varFrm()
              ' **
              ' **   Field  Element  Name         Constant
              ' **   =====  =======  ===========  ==========
              ' **     1       0     dbs_id       F_DID
              ' **     2       1     frm_id       F_FID
              ' **     3       2     frm_name     F_FNAM
              ' **     4       3     IsPopUp      F_POP
              ' **
              ' **********************************************
28320         .Close
28330       End With
28340       Set rst = Nothing
28350       Set qdf = Nothing
28360       .Close
28370     End With
28380     Set dbs = Nothing

28390     If lngFrms > 0& Then

28400       For lngX = 0& To (lngFrms - 1&)
28410         DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
28420         Set frm = Forms(arr_varFrm(F_FNAM, lngX))
28430         With frm
28440           arr_varFrm(F_POP, lngX) = .PopUp
28450         End With
28460         Set frm = Nothing
28470         DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX), acSaveNo
28480       Next

28490       lngY = 0&
28500       For lngX = 0& To (lngFrms - 1&)
28510         If arr_varFrm(F_POP, lngX) = True Then
28520           lngY = lngY + 1&
28530         End If
28540       Next

28550       If lngY > 0& Then
28560         strMsg = CStr(lngY) & " form"
28570         If lngY > 1& Then
28580           strMsg = strMsg & "s are "
28590         Else
28600           strMsg = strMsg & " is "
28610         End If
28620         strMsg = strMsg & "PopUp that shouldn't be!" & vbCrLf
28630         For lngX = 0& To (lngFrms - 1&)
28640           If arr_varFrm(F_POP, lngX) = True Then
28650             strMsg = strMsg & arr_varFrm(F_FNAM, lngX) & vbCrLf
28660           End If
28670         Next
28680         Beep
28690         DoCmd.Hourglass False
28700         MsgBox strMsg, vbCritical + vbOKOnly, "PopUps Found"
28710       End If

28720     End If

28730   End If

EXITP:
28740   Set frm = Nothing
28750   Set rst = Nothing
28760   Set qdf = Nothing
28770   Set dbs = Nothing
28780   Frm_PopUpChk = blnRetVal
28790   Exit Function

ERRH:
28800   blnRetVal = False
28810   Select Case ERR.Number
        Case Else
28820     Beep
28830     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
28840   End Select
28850   Resume EXITP

End Function

Public Function SwapDemoEmpty() As Boolean
' ** Swap data between Demo and Empty data files.

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "SwapDemoEmpty"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strFileArch1 As String, strFileArch2 As String, strFileDta1 As String, strFileDta2 As String
        Dim strPath As String, strPathFile1 As String, strPathFile2 As String
        Dim msgResponse As VbMsgBoxResult
        Dim intMode As Integer
        Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long
        Dim blnRetVal As Boolean

28910   blnRetVal = True

28920   strPath = CurrentBackendPath  ' ** Module Function: modFileUtilities

28930   strFileArch1 = "TrstArch_bak.mdb"
28940   strFileArch2 = "TrstArch_bak_WmB_ti.mdb"
28950   strFileDta1 = "TrustDta_bak.mdb"
28960   strFileDta2 = "TrustDta_bak_WmB_ti.mdb"

28970   intMode = 0
28980   lngTmp02 = DCount("*", "account")
28990   If lngTmp02 = 2 Then
          ' ** Current data is empty, so swap in demo.
29000     intMode = 1  ' ** Empty-to-Demo.
29010   Else
          ' ** Current data may be demo.
29020     varTmp00 = DLookup("[shortname]", "account", "[accountno] = '11'")
29030     Select Case IsNull(varTmp00)
          Case True
            ' ** Some data other than our demo data.
29040       Beep
29050       msgResponse = MsgBox("Current data is neither Demo, nor Empty." & vbCrLf & vbCrLf & _
              "Swap in empty?", vbQuestion + vbOKCancel, "Unknown Data To Empty")
29060       If msgResponse = vbOK Then
29070         intMode = 2
29080       Else
29090         blnRetVal = False
29100       End If
29110     Case False
29120       If varTmp00 = "William B. Johnson Trust" Then
              ' ** Yes, our demo data.
29130         intMode = 2  ' ** Demo-to-Empty.
29140       Else
              ' ** Who knows!?
29150         Beep
29160         msgResponse = MsgBox("Current data is neither Demo, nor Empty." & vbCrLf & vbCrLf & _
                "Swap in empty?", vbQuestion + vbOKCancel, "Unknown Data To Empty")
29170         If msgResponse = vbOK Then
29180           intMode = 2
29190         Else
29200           blnRetVal = False
29210         End If
29220       End If
29230     End Select
29240   End If

29250   If blnRetVal = True Then
29260     Select Case intMode
          Case 1
            ' ** Empty-to-Demo.
29270       strPathFile1 = strPath & LNK_SEP & strFileArch2
29280       blnRetVal = FileExists(strPathFile1)  ' ** Module Function: modFileUtilities.
29290       Select Case blnRetVal
            Case True
29300         strPathFile2 = strPath & LNK_SEP & strFileDta2
29310         blnRetVal = FileExists(strPathFile2)  ' ** Module Function: modFileUtilities.
29320         Select Case blnRetVal
              Case True
                ' ** Proceed with swap.
29330         Case False
29340           Beep
29350           MsgBox "The file '" & strPathFile2 & "' could not be found.", vbCritical + vbOKOnly, "File Not Found"
29360         End Select
29370       Case False
29380         Beep
29390         MsgBox "The file '" & strPathFile1 & "' could not be found.", vbCritical + vbOKOnly, "File Not Found"
29400       End Select
29410     Case 2
            ' ** Demo-to-Empty.
29420       strPathFile1 = strPath & LNK_SEP & strFileArch1
29430       blnRetVal = FileExists(strPathFile1)  ' ** Module Function: modFileUtilities.
29440       Select Case blnRetVal
            Case True
29450         strPathFile2 = strPath & LNK_SEP & strFileDta1
29460         blnRetVal = FileExists(strPathFile2)  ' ** Module Function: modFileUtilities.
29470         Select Case blnRetVal
              Case True
                ' ** Proceed with swap.
29480         Case False
29490           Beep
29500           MsgBox "The file '" & strPathFile2 & "' could not be found.", vbCritical + vbOKOnly, "File Not Found"
29510         End Select
29520       Case False
29530         Beep
29540         MsgBox "The file '" & strPathFile1 & "' could not be found.", vbCritical + vbOKOnly, "File Not Found"
29550       End Select
29560     End Select
29570   End If  ' ** blnRetval.

29580   If blnRetVal = True Then

29590     strPathFile1 = strPath & LNK_SEP & gstrFile_ArchDataName
29600     strPathFile2 = strPath & LNK_SEP & gstrFile_DataName

          ' ** Delete the current data files.
29610     Kill strPathFile1
29620     DoEvents
29630     Kill strPathFile2
29640     DoEvents

29650     Select Case intMode
          Case 1
            ' ** Empty-to-Demo.
29660       strPathFile1 = strPath & LNK_SEP & gstrFile_ArchDataName
29670       strPathFile2 = strPath & LNK_SEP & strFileArch2
29680       FileCopy strPathFile2, strPathFile1
29690       DoEvents
29700       strPathFile1 = strPath & LNK_SEP & gstrFile_DataName
29710       strPathFile2 = strPath & LNK_SEP & strFileDta2
29720       FileCopy strPathFile2, strPathFile1
29730       DoEvents
29740     Case 2
            ' ** Demo-to-Empty.
29750       strPathFile1 = strPath & LNK_SEP & gstrFile_ArchDataName
29760       strPathFile2 = strPath & LNK_SEP & strFileArch1
29770       FileCopy strPathFile2, strPathFile1
29780       DoEvents
29790       strPathFile1 = strPath & LNK_SEP & gstrFile_DataName
29800       strPathFile2 = strPath & LNK_SEP & strFileDta1
29810       FileCopy strPathFile2, strPathFile1
29820       DoEvents
29830     End Select

29840   End If  ' ** blnRetval.

29850   Beep

EXITP:
29860   SwapDemoEmpty = blnRetVal
29870   Exit Function

ERRH:
29880   blnRetVal = False
29890   Select Case ERR.Number
        Case Else
29900     Beep
29910     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
29920   End Select
29930   Resume EXITP

End Function
