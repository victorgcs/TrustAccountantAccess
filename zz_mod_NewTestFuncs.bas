Attribute VB_Name = "zz_mod_NewTestFuncs"
Option Compare Database
Option Explicit

'VGC 05/09/2015: CHANGES!

' **************************************
' ** VGC 04/01/2010:
' ** NOTE REQUIRED ZZ_'S IN DOC FUNCS.
' **   tblQuery_Staging2
' **   zz_qry_System_20 - zz_qry_System_31
' **   zz_qry_System_84a
' **   zz_qry_System_84a_01
' **   zz_qry_System_84a_02
' **   zz_qry_System_84b
' **   zz_qry_System_84f
' **   zz_qry_System_84g
' **   zz_qry_System_84j
' **   zz_qry_System_84k
' **   zz_qry_System_85a
' **   zz_qry_System_85a_01
' **   zz_qry_System_85b_01
' **   zz_qry_System_85b_02
' **   zz_qry_System_85b_03
' **   zz_qry_System_85b_04
' **   zz_qry_System_85b_05
' **   zz_qry_System_85b_06
' **   zz_qry_System_85cx
' ** ELSEWHERE:
' **   zz_qry_System_19_01
' **   zz_qry_System_19_02
' **   zz_qry_System_19_03
' **   zz_qry_System_46_01
' **   zz_qry_System_46_02
' **   zz_qry_System_70a
' **   zz_qry_System_70b
' **   zz_qry_System_72
' **************************************

'CurrentDb.Containers![Databases].Documents("SummaryInfo").Properties("Title") = "Trust Accountant™"
'CurrentDb.Containers![Databases].Documents.Count = 4:
'0 AccessLayout
'1 MSysDb  (Startup Properties, etc.)
'    .Documents("MSysDb").Properties("AppTitle") = Trust Accountant™
'2 SummaryInfo
'    .Documents("SummaryInfo").Properties("Title") = Trust Accountant™
'    .Documents("SummaryInfo").Properties("Author") = gb, mike, vgc, et al
'    .Documents("SummaryInfo").Properties("Manager") = Rich McCabe
'    .Documents("SummaryInfo").Properties("Company") = Delta Data, Inc.
'3 UserDefined
'    .Documents("UserDefined").Properties("AppVersion") = 2.1.146
'    .Documents("UserDefined").Properties("AppDate") = 12/07/07

Private Const THIS_NAME As String = "zz_mod_NewTestFuncs"
' **

Public Function LicenseInfo_Get() As Boolean
' ** Display the current license information to the Immediate Window.
' ** See also zz_mod_MDEPrepFuncs.xIniFile_Set().

100   On Error GoTo ERRH

        Const THIS_PROC As String = "LicenseInfo_Get"

        Dim strSection As String
        Dim strLicTo As String, strExpires As String, strLimit As String, strPExpires As String
        Dim strWorkDir As String, strLicDir As String, strFile As String
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     strSection = "License"
130     strFile = "TA.lic"

        'strWorkDir = "C:\VictorGCS_Clients\TrustAccountant\NewDemo\"
        'strWorkDir = "C:\VictorGCS_Clients\TrustAccountant\NewUpgrade\"
140     strWorkDir = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\"
        'strLicDir = "Untouched\"
150     strLicDir = "DemoDatabase\"
        'strLicDir = "EmptyDatabase\"
        'strLicDir = "TestDatabase\"
160     strFile = strWorkDir & strLicDir & strFile

        'If gstrTrustDataLocation = vbNullString Then
        '  IniFile_GetDataLoc  ' ** Module Function: modUtilities.
        'End If
        'strFile = gstrTrustDataLocation & strFile

170     strLicTo = DecodeString(IniFile_Get(strSection, "Firm", "", strFile))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
180     strExpires = DecodeString(IniFile_Get(strSection, "Expires", "", strFile))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
190     strPExpires = DecodeString(IniFile_Get(strSection, "Pricing", "", strFile))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
200     strLimit = DecodeString(IniFile_Get(strSection, "Limit", "", strFile))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.

210     Debug.Print "'" & strLicTo
220     Debug.Print "'" & strExpires
230     Debug.Print "'" & strPExpires
240     Debug.Print "'" & strLimit

EXITP:
250     LicenseInfo_Get = blnRetVal
260     Exit Function

ERRH:
270     blnRetVal = False
280     Select Case ERR.Number
        Case Else
290       Beep
300       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
310     End Select
320     Resume EXITP

End Function

Public Function LicenseInfo_Set() As Boolean
' ** Example:
' **   strSection = "License"
' **   strSubSection = "Firm"
' **   strValue = EncodeString(Me.txtLicensedTo)
' **   strFile = gstrTrustDataLocation & "TA.lic"
' ** See also zz_mod_MDEPrepFuncs.xIniFile_Set().

400   On Error GoTo ERRH

        Const THIS_PROC As String = "LicenseInfo_Set"

        Dim strSection As String, strSubSection As String, strValue As String, strFile As String
        Dim strWorkDir As String, strLicDir As String
        Dim blnRetVal As Boolean

410     blnRetVal = True

420     strSection = "License"
430     strFile = "TA.lic"

        'strWorkDir = "C:\VictorGCS_Clients\TrustAccountant\NewDemo\"
        'strWorkDir = "C:\VictorGCS_Clients\TrustAccountant\NewUpgrade\"
440     strWorkDir = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\"
450     strLicDir = "Untouched\"
        'strLicDir = "DemoDatabase\"
        'strLicDir = "EmptyDatabase\"
        'strLicDir = "TestDatabase\"
460     strFile = strWorkDir & strLicDir & strFile

        'If gstrTrustDataLocation = vbNullString Then
        '  IniFile_GetDataLoc  ' ** Module Function: modUtilities.
        'End If
        'strFile = gstrTrustDataLocation & strFile

        'strSubSection = "Firm"  ' ** Licensed To.
        'strSubSection = "Limit"  ' ** Limit code.
        'strValue = EncodeString("400")
470     strSubSection = "Expires"  ' ** Expires.
        'strSubSection = "Pricing"  ' ** Pricing.
480     strValue = EncodeString("01/01/2011")

490     blnRetVal = IniFile_Set(strSection, strSubSection, strValue, strFile)  ' ** Module Procedure: modStartupFuncs.

500     Beep

EXITP:
510     LicenseInfo_Set = blnRetVal
520     Exit Function

ERRH:
530     blnRetVal = False
540     Select Case ERR.Number
        Case Else
550       Beep
560       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
570     End Select
580     Resume EXITP

End Function
