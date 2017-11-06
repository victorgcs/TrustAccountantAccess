Attribute VB_Name = "modPreferenceFuncs"
Option Compare Database
Option Explicit

'VGC 09/05/2017: CHANGES!

Private Const THIS_NAME As String = "modPreferenceFuncs"
' **

Public Sub Pref_Load(strFormName As String)
' ** Load user preferences for the specified form.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
        Dim strFormName_Parent As String
        Dim lngRecs As Long
        Dim blnRetVal As Boolean
        Dim intPos01 As Integer
        Dim lngX As Long

110     gblnBadLink = False
120     blnRetVal = TableExists("tblPreference_User", , , True) ' ** Module Function: modFileUtilities.

130     Select Case gblnBadLink
        Case True
          ' ** It's a bad link to the database!
140       gblnBadLink = False
150       Pref_LoadTemplate strFormName  ' ** Procedure: Below.
160     Case False
170       Select Case blnRetVal
          Case True

            ' ** Determine whether we're dealing with a subform.
180         intPos01 = InStr(strFormName, "_Sub")
190         If intPos01 > 0 Then
200           strFormName_Parent = Left(strFormName, (intPos01 - 1))
210         Else
220           strFormName_Parent = vbNullString
230         End If
240         If strFormName_Parent = "frmFeeSchedules_Detail" Then strFormName_Parent = "frmFeeSchedules"

250         Set dbs = CurrentDb
260         With dbs
              ' ** Retrieve user saved preference by [frmnam], [usr].
270           Set qdf = .QueryDefs("qryPreferences_02a")  '##dbs_id
280           With qdf.Parameters
290             ![frmnam] = strFormName  ' ** This remains listed as the actual form.
300             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
310           End With
320           Set rst = qdf.OpenRecordset
330           With rst
340             If .BOF = True And .EOF = True Then
                  ' ** No preferences, or just nothing saved for this form.
350             Else
360               .MoveLast
370               lngRecs = .RecordCount
380               .MoveFirst
390               Select Case strFormName_Parent
                  Case vbNullString
400                 Set frm = Forms(strFormName)
410               Case Else
420                 Set frm = Forms(strFormName_Parent).Controls(strFormName).Form  ' ** Set this to the subform.
430               End Select
440               For lngX = 1& To lngRecs
450                 Set ctl = frm.Controls(rst![ctl_name])
460                 With ctl
470                   If .Name = rst![ctl_name] Then
480                     Select Case rst![datatype_db_type]
                        Case dbBoolean
490                       .Value = rst![prefuser_boolean]
500                       If strFormName = "frmTransaction_Audit_Sub" Then
510                         If Right(rst![ctl_name], 4) <> "_chk" Then
520                           If rst![prefuser_boolean] = True Then
530                             frm.FilterRecs_Clr (rst![ctl_name] & "_AfterUpdate"), True  ' ** Form Procedure: frmTransaction_Audit_Sub.
540                           End If
550                         Else
560                           If rst![prefuser_boolean] = False Then
570                             frm.Print_Chk (rst![ctl_name] & "_AfterUpdate")  ' ** Form Procedure: frmTransaction_Audit_Sub.
580                           End If
590                         End If
600                       End If
610                     Case dbInteger
620   On Error Resume Next
630                       .Value = rst![prefuser_integer]
640   On Error GoTo ERRH
650                     Case dbLong
660   On Error Resume Next
670                       .Value = rst![prefuser_long]
680   On Error GoTo ERRH
690                       If strFormName = "frmTransaction_Audit_Sub" Then
700                         If rst![ctl_name] <> "assetno" And rst![ctl_name] <> "revcode_ID" And rst![ctl_name] <> "revcode_TYPE" Then
710                           If IsNull(rst![prefuser_long]) = False Then
720                             If rst![prefuser_long] > 0& Then
730                               frm.FilterRecs_Clr (rst![ctl_name] & "_AfterUpdate"), True  ' ** Form Procedure: frmTransaction_Audit_Sub.
740                             End If
750                           End If
760                         End If
770                       End If
780                     Case dbCurrency
790                       .Value = rst![prefuser_currency]
800                     Case dbSingle
810                       .Value = rst![prefuser_single]
820                     Case dbDouble
830                       .Value = rst![prefuser_double]
840                     Case dbDate
850                       .Value = rst![prefuser_date]
860                       If strFormName = "frmTransaction_Audit_Sub" Then
870                         If IsNull(rst![prefuser_date]) = False Then
880                           frm.FilterRecs_Clr (rst![ctl_name] & "_AfterUpdate"), True  ' ** Form Procedure: frmTransaction_Audit_Sub.
890                         End If
900                       End If
910                     Case dbText
920                       If strFormName = "frmLocations_Add_Purchase" And .Name = "Loc_State" Then
930                         frm.Controls("Loc_State_Pref").Value = rst![prefuser_text]
940                       Else
950                         .Value = rst![prefuser_text]
960                         If strFormName = "frmTransaction_Audit_Sub" Then
970                           If rst![ctl_name] <> "CurrentFilter" Then
980                             If IsNull(rst![prefuser_text]) = False Then
990                               If rst![prefuser_text] <> vbNullString Then
1000                                frm.FilterRecs_Clr (rst![ctl_name] & "_AfterUpdate"), True  ' ** Form Procedure: frmTransaction_Audit_Sub.
1010                              End If
1020                            End If
1030                          End If
1040                        End If
1050                      End If
1060                    End Select
1070                  End If
1080                End With
1090                If lngX < lngRecs Then .MoveNext
1100              Next
1110            End If
1120            .Close
1130          End With
1140          .Close
1150        End With

1160      Case False
1170        Set frm = Forms(strFormName)
1180        With frm
1190          For Each ctl In .FormHeader.Controls
1200            With ctl
1210              If .Name = "chkNoLink" Then
1220                .Value = True
1230                Exit For
1240              End If
1250            End With
1260          Next
1270        End With
1280      End Select  ' ** blnRetVal.
1290    End Select  ' ** gblnBadLink.

EXITP:
1300    Set ctl = Nothing
1310    Set frm = Nothing
1320    Set rst = Nothing
1330    Set qdf = Nothing
1340    Set dbs = Nothing
1350    Exit Sub

ERRH:
1360    Select Case ERR.Number
        Case Else
1370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1380    End Select
1390    Resume EXITP

End Sub

Public Sub Pref_Save(strFormName As String)
' ** Save user preferences for the specified form.

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_Save"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim frm As Access.Form, ctl As Access.Control
        Dim strFormName_Parent As String
        Dim blnIsState As Boolean, blnHasState As Boolean, strStateControl As String, varStateCode As Variant
        Dim lngRecs As Long
        Dim blnAdd As Boolean, blnRetVal As Boolean, blnUpdate As Boolean
        Dim intPos01 As Integer
        Dim varTmp00 As Variant
        Dim lngX As Long

1410    blnUpdate = True
1420    blnHasState = False: strStateControl = vbNullString  ' ** There can be only 1 .._State_Pref per form.

        ' ** Determine whether we're dealing with a subform.
1430    intPos01 = InStr(strFormName, "_Sub")
1440    If intPos01 > 0 Then
1450      strFormName_Parent = Left(strFormName, (intPos01 - 1))
1460    Else
1470      strFormName_Parent = vbNullString
1480    End If
        ' ** This should be renamed to frmFeeSchedules_Sub_Detail!
1490    If strFormName_Parent = "frmFeeSchedules_Detail" Then strFormName_Parent = "frmFeeSchedules"

        ' ** We need the parent in order to address the control reference,
        ' ** but the prefs are listed under the subform's name.

1500    gblnBadLink = False
1510    blnRetVal = TableExists("tblPreference_Control", , , True) ' ** Module Function: modFileUtilities.

1520    Select Case gblnBadLink
        Case True
          ' ** It's a bad link to the database!
1530      gblnBadLink = False
1540      Pref_SaveTemplate strFormName  ' ** Procedure: Below.
1550    Case False

1560      Set dbs = CurrentDb
1570      With dbs
            ' ** tblPreference_Control, just dbs_id = 1, specified by [frmnam].
1580        Set qdf1 = .QueryDefs("qryPreferences_01a")  '##dbs_id
1590        With qdf1.Parameters
1600          ![frmnam] = strFormName  ' ** This remains listed as the actual form.
1610        End With
1620        Select Case strFormName_Parent
            Case vbNullString
1630          blnRetVal = IsLoaded(strFormName, acForm)  ' ** Module Function: modFileUtilities.
1640        Case Else
1650          blnRetVal = IsLoaded(strFormName_Parent, acForm)  ' ** Module Function: modFileUtilities.
1660        End Select
1670        If blnRetVal = True Then  ' ** Module Function: modFileUtilities.
1680          Set rst1 = qdf1.OpenRecordset
1690          If rst1.BOF = True And rst1.EOF = True Then
                ' ** No preferences designated for this form.
1700            rst1.Close
1710          Else
1720            rst1.MoveLast
1730            lngRecs = rst1.RecordCount
1740            rst1.MoveFirst
1750            Select Case strFormName_Parent
                Case vbNullString
1760              Set frm = Forms(strFormName)
1770            Case Else
1780              Set frm = Forms(strFormName_Parent).Controls(strFormName).Form  ' ** Set this to the subform.
1790            End Select
1800            For lngX = 1& To lngRecs
1810              blnAdd = False: blnIsState = False
1820              If InStr(rst1![ctl_name], "_State_Pref") > 0 Then
1830                blnIsState = True
1840                strStateControl = rst1![ctl_name]
1850              End If
                  ' ** tblPreference_User, with frm_name, ctl_name.
1860              Set qdf2 = .QueryDefs("qryPreferences_15_01")  '##dbs_id
1870              Set rst2 = qdf2.OpenRecordset
1880              With rst2
1890                If .BOF = True And .EOF = True Then
                      ' ** None saved whatsoever!
1900                  blnAdd = True
1910                Else
1920                  .FindFirst "[frm_name] = '" & rst1![frm_name] & "' And [ctl_name] = '" & rst1![ctl_name] & "' And " & _
                        "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
1930                  If .NoMatch = True Then
1940                    blnAdd = True
1950                  Else
1960  On Error Resume Next
1970                    varTmp00 = frm.Controls(rst1![ctl_name])
1980                    If ERR.Number <> 0 Then
                          ' ** Error: 2427  You entered an expression that has no value.
1990  On Error GoTo ERRH
2000                      blnUpdate = False
2010                    Else
2020  On Error GoTo ERRH
2030                    End If
2040                    If blnUpdate = True Then
2050                      .Edit
2060                      Select Case rst1![datatype_db_type]
                          Case dbBoolean
2070                        If IsNull(frm.Controls(rst1![ctl_name])) = True Then
2080                          ![prefuser_boolean] = False
2090                        Else
2100                          ![prefuser_boolean] = frm.Controls(rst1![ctl_name])
2110                        End If
2120                      Case dbInteger
2130                        ![prefuser_integer] = frm.Controls(rst1![ctl_name])
2140                      Case dbLong
2150                        ![prefuser_long] = frm.Controls(rst1![ctl_name])
2160                      Case dbCurrency
2170                        ![prefuser_currency] = frm.Controls(rst1![ctl_name])
2180                      Case dbSingle
2190                        ![prefuser_single] = frm.Controls(rst1![ctl_name])
2200                      Case dbDouble
2210                        ![prefuser_double] = frm.Controls(rst1![ctl_name])
2220                      Case dbDate
2230                        ![prefuser_date] = frm.Controls(rst1![ctl_name])
2240                      Case dbText
2250                        If IsNull(frm.Controls(rst1![ctl_name])) = True Then
2260                          Select Case blnIsState
                              Case True
                                ' ** Don't Null-out any existing state code prefs.
2270                          Case False
2280                            ![prefuser_text] = Null
2290                          End Select
2300                        Else
2310                          If Trim(frm.Controls(rst1![ctl_name])) = vbNullString Then
2320                            Select Case blnIsState
                                Case True
                                  ' ** Don't Null-out any existing state code prefs.
2330                            Case False
2340                              ![prefuser_text] = Null
2350                            End Select
2360                          Else
2370                            ![prefuser_text] = frm.Controls(rst1![ctl_name])
2380                            If blnIsState = True Then
2390                              blnHasState = True
2400                              varStateCode = frm.Controls(rst1![ctl_name])
2410                            End If
2420                          End If
2430                        End If
2440                      End Select
2450                      ![DateModified] = Now()
2460                      .Update
2470                    End If  ' ** blnUpdate.
2480                  End If  ' ** NoMatch.
2490                End If  ' ** BOF, EOF.
2500                If blnAdd = True Then
2510  On Error Resume Next
2520                  varTmp00 = frm.Controls(rst1![ctl_name])
2530                  If ERR.Number <> 0 Then
2540  On Error GoTo ERRH
'2530  On Error GoTo 0
                        ' ** Error: 2427  You entered an expression that has no value.
2550                    blnUpdate = False
2560                  Else
2570  On Error GoTo ERRH
'2560  On Error GoTo 0
2580                  End If
2590                  If blnUpdate = True Then
2600                    .AddNew
2610                    ![dbs_id] = rst1![dbs_id]
2620                    ![frm_id] = rst1![frm_id]
2630                    ![ctl_id] = rst1![ctl_id]
2640                    ![prefctl_id] = rst1![prefctl_id]
2650                    Select Case rst1![datatype_db_type]
                        Case dbBoolean
2660                      If IsNull(frm.Controls(rst1![ctl_name])) = True Then
2670                        ![prefuser_boolean] = False
2680                      Else
2690                        ![prefuser_boolean] = frm.Controls(rst1![ctl_name])
2700                      End If
2710                    Case dbInteger
2720                      ![prefuser_integer] = frm.Controls(rst1![ctl_name])
2730                    Case dbLong
2740                      ![prefuser_long] = frm.Controls(rst1![ctl_name])
2750                    Case dbCurrency
2760                      ![prefuser_currency] = frm.Controls(rst1![ctl_name])
2770                    Case dbSingle
2780                      ![prefuser_single] = frm.Controls(rst1![ctl_name])
2790                    Case dbDouble
2800                      ![prefuser_double] = frm.Controls(rst1![ctl_name])
2810                    Case dbDate
2820                      ![prefuser_date] = frm.Controls(rst1![ctl_name])
2830                    Case dbText
2840                      If IsNull(frm.Controls(rst1![ctl_name])) = True Then
2850                        ![prefuser_text] = Null
2860                      Else
2870                        If Trim(frm.Controls(rst1![ctl_name])) = vbNullString Then
2880                          ![prefuser_text] = Null
2890                        Else
2900                          ![prefuser_text] = frm.Controls(rst1![ctl_name])
2910                          If blnIsState = True Then
2920                            blnHasState = True
2930                            varStateCode = frm.Controls(rst1![ctl_name])
2940                          End If
2950                        End If
2960                      End If
2970                    End Select
2980                    ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2990                    ![DateCreated] = Now()
3000                    ![DateModified] = Now()
3010                    .Update
3020                  End If  ' ** blnUpdate.
3030                End If  ' ** blnAdd.
3040                .Close
3050              End With  ' ** rst2.
3060              If lngX < lngRecs Then rst1.MoveNext
3070            Next  ' ** lngX.
3080            rst1.Close
3090            If blnHasState = True Then
3100              Pref_State varStateCode, strFormName, strStateControl  ' ** Procedure: Below.
3110            End If
3120          End If  ' ** BOF, EOF.
3130        End If  ' ** blnRetVal.
3140        .Close
3150      End With  ' ** dbs.
3160    End Select  ' ** gblnBadLink.

EXITP:
3170    Set ctl = Nothing
3180    Set frm = Nothing
3190    Set rst1 = Nothing
3200    Set rst2 = Nothing
3210    Set qdf1 = Nothing
3220    Set qdf2 = Nothing
3230    Set dbs = Nothing
3240    Exit Sub

ERRH:
3250    Select Case ERR.Number
        Case 2310  ' ** You entered an expression that has no value.
          ' ** It closed before the procedure could complete.
3260    Case 2450  ' ** Microsoft Access can't find the form '|' referred to in a macro expression or Visual Basic code.
          ' ** It closed before the procedure could complete.
3270    Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

End Sub

Public Function Pref_RemNot(strFormName As String, strType As String, Optional varCtlName As Variant) As Boolean
' ** Remember NOT!

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_RemNot"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strThisUser As String
        Dim blnRetVal As Boolean

3410    blnRetVal = True

3420    strThisUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

        ' ** Delete residual prefs when Remember check boxes are unchecked.
3430    Set dbs = CurrentDb
3440    With dbs
3450      Select Case strType
          Case "Dates"  ' ** Various controls.
3460        Select Case strFormName
            Case "frmStatementParameters"
3470          Select Case IsMissing(varCtlName)
              Case True
                ' ** Delete tblPreference_User, linked to qryPreferences_17_01 (tblPreference_Control,
                ' ** just frmStatementParameters, for All Dates), by specified [usr].
3480            Set qdf = .QueryDefs("qryPreferences_18_01")  '##dbs_id
3490          Case False
3500            Select Case varCtlName
                Case "Trans"
                  ' ** Delete tblPreference_User, linked to qryPreferences_17_03 (tblPreference_Control,
                  ' ** just frmStatementParameters, for TransDateStart, TransDateEnd), by specified [usr].
3510              Set qdf = .QueryDefs("qryPreferences_18_03")  '##dbs_id
3520            Case "Asset"
                  ' ** Delete tblPreference_User, linked to qryPreferences_17_04 (tblPreference_Control,
                  ' ** just frmStatementParameters, for AssetListDate), by specified [usr].
3530              Set qdf = .QueryDefs("qryPreferences_18_04")  '##dbs_id
3540            Case "Stmt"
                  ' ** Delete tblPreference_User, linked to qryPreferences_17_05 (tblPreference_Control,
                  ' ** just frmStatementParameters, for cmbMonth, StatementsYear), by specified [usr].
3550              Set qdf = .QueryDefs("qryPreferences_18_05")  '##dbs_id
3560            End Select
3570          End Select
3580        End Select
3590        With qdf.Parameters
3600          ![usr] = strThisUser
3610        End With
3620        qdf.Execute
3630      Case "Account"  ' ** Just cmbAccounts.
3640        Select Case IsMissing(varCtlName)
            Case True
              ' ** Delete tblPreference_User, linked to qryPreferences_17_02 (tblPreference_Control,
              ' ** for cmbAccounts, by specified [fnam]), by specified [usr].
3650          Set qdf = .QueryDefs("qryPreferences_18_02")  '##dbs_id
3660          With qdf.Parameters
3670            ![fnam] = strFormName
3680            ![usr] = strThisUser
3690          End With
3700          qdf.Execute
3710        Case False
3720          Select Case varCtlName
              Case "LastAcctNo"
                ' ** Delete tblPreference_User, linked to qryPreferences_17_06 (tblPreference_Control,
                ' ** just frmJournal, frmJournal_Columns, for LastAcctNo), by specified [usr].
3730            Set qdf = .QueryDefs("qryPreferences_18_06")  '##dbs_id
3740            With qdf.Parameters
3750              ![usr] = strThisUser
3760            End With
3770          End Select
3780          qdf.Execute
3790        End Select
3800      Case "Asset"
            ' **
3810      End Select
3820      .Close
3830    End With

EXITP:
3840    Set qdf = Nothing
3850    Set dbs = Nothing
3860    Pref_RemNot = blnRetVal
3870    Exit Function

ERRH:
3880    blnRetVal = False
3890    Select Case ERR.Number
        Case 2310  ' ** You entered an expression that has no value.
          ' ** It closed before the procedure could complete.
3900    Case 2450  ' ** Microsoft Access can't find the form '|' referred to in a macro expression or Visual Basic code.
          ' ** It closed before the procedure could complete.
3910    Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Function

Public Sub Pref_State(varStateCode As Variant, strForm As String, strControl As String)
' ** Manage the 2-letter state code preferences.
' ** Once a user enters a state code anywhere
' ** within Trust Accountant, make that the default
' ** code in all the state fields. If they've got
' ** one out-of-state address, it'll continue to
' ** offer that until they put in something else.
'StateCodeQrySet  ' ** Module Function: modPreferenceFuncs.
'frmAccountProfile
'frmAccountProfile_Add
'frmCompanyInfo
'frmLocations
'frmLocations_Add
'frmLocations_Add_Purchase
'frmRecurringItems
'frmRecurringItems_Add
'frmRecurringItems_Add_Misc
'frmVersion_Input

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_State"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngPrefs As Long, arr_varPref As Variant
        Dim strThisUser As String
        Dim lngRecs As Long
        Dim lngX As Long, lngY As Long

        ' ** Array: arr_varPref().
        Const P_PID  As Integer = 0
        Const P_DID  As Integer = 1
        'Const P_DNAM As Integer = 2
        Const P_FID  As Integer = 3
        Const P_FNAM As Integer = 4
        Const P_CID  As Integer = 5
        Const P_CNAM As Integer = 6
        'Const P_TYP  As Integer = 7
        Const P_HASP As Integer = 8

4010    If IsNull(varStateCode) = False And strForm <> vbNullString And strControl <> vbNullString Then
          ' ** Currently, these are all considered one preference.
          ' ** They could, however, be handled separately.
          ' **   Acct_State_Pref
          ' **   CoInfo_State_Pref
          ' **   Loc_State_Pref
          ' **   Recur_State_Pref

4020      strThisUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

4030      Set dbs = CurrentDb
4040      With dbs

            ' ** tblPreference_Control, just dbs_id = 1, all '.._State_Pref' controls.
4050        Set qdf = .QueryDefs("qryPreferences_19_01")  '##dbs_id
4060        Set rst = qdf.OpenRecordset
4070        With rst
4080          If .BOF = True And .EOF = True Then
                ' ** Shouldn't happen!
4090          Else
4100            .MoveLast
4110            lngPrefs = .RecordCount
4120            .MoveFirst
4130            arr_varPref = .GetRows(lngPrefs)
                ' *****************************************************
                ' ** Array: arr_varPref()
                ' **
                ' **   Field  Element  Name                Constant
                ' **   =====  =======  ==================  ==========
                ' **     1       0     prefctl_id          P_PID
                ' **     2       1     dbs_id              P_DID
                ' **     3       2     dbs_name            P_DNAM
                ' **     4       3     frm_id              P_FID
                ' **     5       4     frm_name            P_FNAM
                ' **     6       5     ctl_id              P_CID
                ' **     7       6     ctl_name            P_CNAM
                ' **     8       7     datatype_db_type    P_TYP
                ' **     9       8     HasPref             P_HASP
                ' **
                ' *****************************************************
4140          End If
4150          .Close
4160        End With
4170        Set rst = Nothing
4180        Set qdf = Nothing

            ' ** tblPreference_User, linked to tblPreference_Control, just
            ' ** dbs_id = 1, '.._State_Pref' controls, by specified [usr].
4190        Set qdf = .QueryDefs("qryPreferences_19_02")  '##dbs_id
4200        With qdf.Parameters
4210          ![usr] = strThisUser
4220        End With
4230        Set rst = qdf.OpenRecordset
4240        With rst
4250          If .BOF = True And .EOF = True Then
                ' ** This will be the first.
4260          Else
4270            .MoveLast
4280            lngRecs = .RecordCount
4290            .MoveFirst
                ' ** Put the current state code in all state prefs.
4300            For lngX = 1& To lngRecs
4310              For lngY = 0& To (lngPrefs - 1&)
4320                If arr_varPref(P_PID, lngY) = ![prefctl_id] Then
4330                  arr_varPref(P_HASP, lngY) = CBool(True)
4340                  If arr_varPref(P_FNAM, lngY) = strForm And arr_varPref(P_CNAM, lngY) = strControl Then
                        ' ** This current one was handled above.
4350                  Else
4360                    Select Case IsNull(![prefuser_text])
                        Case True
4370                      .Edit
4380                      ![prefuser_text] = varStateCode
4390                      ![DateModified] = Now()
4400                      .Update
4410                    Case False
4420                      If ![prefuser_text] <> varStateCode Then
4430                        .Edit
4440                        ![prefuser_text] = varStateCode
4450                        ![DateModified] = Now()
4460                        .Update
4470                      End If
4480                    End Select
4490                  End If
4500                  Exit For
4510                End If
4520              Next  ' ** lngPrefs: lngY.
4530              If lngX < lngRecs Then .MoveNext
4540            Next  ' ** lngRecs: lngX.
4550          End If
4560          .Close
4570        End With  ' ** rst.
4580        Set rst = Nothing
4590        Set qdf = Nothing

            ' ** Add any prefs they're currently missing.
4600        Set rst = .OpenRecordset("tblPreference_User", dbOpenDynaset, dbAppendOnly)
4610        With rst
4620          For lngX = 0& To (lngPrefs - 1&)
4630            If arr_varPref(P_HASP, lngX) = False Then
4640              .AddNew
4650              ![dbs_id] = arr_varPref(P_DID, lngX)  '##dbs_id
4660              ![frm_id] = arr_varPref(P_FID, lngX)
4670              ![ctl_id] = arr_varPref(P_CID, lngX)
4680              ![prefctl_id] = arr_varPref(P_PID, lngX)
4690              ![prefuser_text] = varStateCode
4700              ![Username] = strThisUser
4710              ![DateCreated] = Now()
4720              ![DateModified] = Now()
4730              .Update
4740            End If
4750          Next
4760          .Close
4770        End With

4780        .Close
4790      End With  ' ** dbs.

4800    End If

EXITP:
4810    Set rst = Nothing
4820    Set qdf = Nothing
4830    Set dbs = Nothing
4840    Exit Sub

ERRH:
4850    Select Case ERR.Number
        Case Else
4860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4870    End Select
4880    Resume EXITP

End Sub

Public Function Pref_Sync() As Boolean

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_Sync"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, fld As DAO.Field
        Dim lngNoLongers As Long, arr_varNoLonger As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

4910    blnRetVal = True

4920    If gstrTrustDataLocation = vbNullString Then
4930      IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
4940    End If

4950    If gstrTrustDataLocation <> vbNullString Then
4960      If TableExists("tblPreference_Control") = True And _
              TableExists("tblPreference_User") = True Then  ' ** Module Function: modFileUtilities.

4970        Set dbs = CurrentDb
4980        With dbs

              ' ** tblPreference_User, not in tblPreference_Control.
4990          Set qdf = .QueryDefs("qryPreferences_13")  '##dbs_id
5000          Set rst = qdf.OpenRecordset
5010          With rst
5020            If .BOF = True And .EOF = True Then
                  ' ** All's well.
5030              lngNoLongers = 0&
5040            Else
5050              .MoveLast
5060              lngNoLongers = .RecordCount
5070              .MoveFirst
5080              arr_varNoLonger = .GetRows(lngNoLongers)
5090            End If
5100            .Close
5110          End With

              ' ** Delete obsolete preferences.
5120          If lngNoLongers > 0& Then
5130            For lngX = 0& To (lngNoLongers - 1&)
                  ' ** Delete tblPreference_User, by specified [prfusrid].
5140              Set qdf = .QueryDefs("qryPreferences_14_01")  '##dbs_id
5150              With qdf.Parameters
5160                ![prfusrid] = arr_varNoLonger(0, lngX)
5170              End With
5180              qdf.Execute
5190            Next
5200          End If

              ' ** Delete tblPreference_User, for non-existant Username's.
              ' ** Because of the CascadeDelete from Users to tblPreference_User,
              ' ** this should always come up empty. Just to be sure!
5210          Set qdf = .QueryDefs("qryPreferences_14_02")  '##dbs_id
5220          qdf.Execute

5230          .Close
5240        End With

5250      Else
            ' ** Should their presence have been already assured?
5260      End If
5270    End If

EXITP:
5280    Set fld = Nothing
5290    Set rst = Nothing
5300    Set qdf = Nothing
5310    Set dbs = Nothing
5320    Pref_Sync = blnRetVal
5330    Exit Function

ERRH:
5340    blnRetVal = False
5350    Select Case ERR.Number
        Case Else
5360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5370    End Select
5380    Resume EXITP

End Function

Public Sub Pref_LoadTemplate(strFormName As String)
' ** Load user preferences for the specified form, using template tables.

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_LoadTemplate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
        Dim strFormName_Parent As String
        Dim lngRecs As Long
        Dim intPos01 As Integer
        Dim lngX As Long

5410    If TableExists("tblTemplate_Preference_User") = True Then  ' ** Module Function: modFileUtilities.

          ' ** Determine whether we're dealing with a subform.
5420      intPos01 = InStr(strFormName, "_Sub")
5430      If intPos01 > 0 Then
5440        strFormName_Parent = Left(strFormName, (intPos01 - 1))
5450      Else
5460        strFormName_Parent = vbNullString
5470      End If

5480      Set dbs = CurrentDb
5490      With dbs
            ' ** Retrieve user saved preference (template), just dbs_id = 1,  by specified [frmnam], [usr].
5500        Set qdf = .QueryDefs("qryPreferences_02b")  '##dbs_id
5510        With qdf.Parameters
5520          ![frmnam] = strFormName  ' ** This remains listed as the actual form.
5530          ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
5540        End With
5550        Set rst = qdf.OpenRecordset
5560        With rst
5570          If .BOF = True And .EOF = True Then
                ' ** No preferences, or just nothing saved for this form.
5580          Else
5590            .MoveLast
5600            lngRecs = .RecordCount
5610            .MoveFirst
5620            Select Case strFormName_Parent
                Case vbNullString
5630              Set frm = Forms(strFormName)
5640            Case Else
5650              Set frm = Forms(strFormName_Parent).Controls(strFormName).Form  ' ** Set this to the subform.
5660            End Select
5670            For lngX = 1& To lngRecs
5680              Set ctl = frm.Controls(rst![ctl_name])
5690              With ctl
5700                If .Name = rst![ctl_name] Then
5710                  Select Case rst![datatype_db_type]
                      Case dbBoolean
5720                    .Value = rst![prefuser_boolean]
5730                    If strFormName = "frmTransaction_Audit_Sub" Then
5740                      If Right(rst![ctl_name], 4) <> "_chk" Then
5750                        If rst![prefuser_boolean] = True Then
5760                          frm.FilterRecs_Clr (rst![ctl_name] & "_AfterUpdate"), True  ' ** Form Procedure: frmTransaction_Audit_Sub.
5770                        End If
5780                      Else
5790                        If rst![prefuser_boolean] = False Then
5800                          frm.Print_Chk (rst![ctl_name] & "_AfterUpdate")  ' ** Form Procedure: frmTransaction_Audit_Sub.
5810                        End If
5820                      End If
5830                    End If
5840                  Case dbInteger
5850                    .Value = rst![prefuser_integer]
5860                  Case dbLong
5870  On Error Resume Next
5880                    .Value = rst![prefuser_long]
5890  On Error GoTo ERRH
5900                    If strFormName = "frmTransaction_Audit_Sub" Then
5910                      If rst![ctl_name] <> "assetno" And rst![ctl_name] <> "revcode_ID" And rst![ctl_name] <> "revcode_TYPE" Then
5920                        If IsNull(rst![prefuser_long]) = False Then
5930                          If rst![prefuser_long] > 0& Then
5940                            frm.FilterRecs_Clr (rst![ctl_name] & "_AfterUpdate"), True  ' ** Form Procedure: frmTransaction_Audit_Sub.
5950                          End If
5960                        End If
5970                      End If
5980                    End If
5990                  Case dbCurrency
6000                    .Value = rst![prefuser_currency]
6010                  Case dbSingle
6020                    .Value = rst![prefuser_single]
6030                  Case dbDouble
6040                    .Value = rst![prefuser_double]
6050                  Case dbDate
6060                    .Value = rst![prefuser_date]
6070                    If strFormName = "frmTransaction_Audit_Sub" Then
6080                      If IsNull(rst![prefuser_date]) = False Then
6090                        frm.FilterRecs_Clr (rst![ctl_name] & "_AfterUpdate"), True  ' ** Form Procedure: frmTransaction_Audit_Sub.
6100                      End If
6110                    End If
6120                  Case dbText
6130                    If strFormName = "frmLocations_Add_Purchase" And .Name = "Loc_State" Then
6140                      frm.Controls("Loc_State_Pref").Value = rst![prefuser_text]
6150                    Else
6160                      .Value = rst![prefuser_text]
6170                      If strFormName = "frmTransaction_Audit_Sub" Then
6180                        If rst![ctl_name] <> "CurrentFilter" Then
6190                          If IsNull(rst![prefuser_text]) = False Then
6200                            If rst![prefuser_text] <> vbNullString Then
6210                              frm.FilterRecs_Clr (rst![ctl_name] & "_AfterUpdate"), True  ' ** Form Procedure: frmTransaction_Audit_Sub.
6220                            End If
6230                          End If
6240                        End If
6250                      End If
6260                    End If
6270                  End Select
6280                End If
6290              End With
6300              If lngX < lngRecs Then .MoveNext
6310            Next
6320          End If
6330          .Close
6340        End With
6350        .Close
6360      End With

6370    Else
6380      Set frm = Forms(strFormName)
6390      With frm
6400        For Each ctl In .FormHeader.Controls
6410          With ctl
6420            If .Name = "chkNoLink" Then
6430              .Value = True
6440              Exit For
6450            End If
6460          End With
6470        Next
6480      End With
6490    End If

EXITP:
6500    Set ctl = Nothing
6510    Set frm = Nothing
6520    Set rst = Nothing
6530    Set qdf = Nothing
6540    Set dbs = Nothing
6550    Exit Sub

ERRH:
6560    Select Case ERR.Number
        Case Else
6570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6580    End Select
6590    Resume EXITP

End Sub

Public Sub Pref_SaveTemplate(strFormName As String)
' ** Save user preferences for the specified form, using template tables.

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_SaveTemplate"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim frm As Access.Form, ctl As Access.Control
        Dim strFormName_Parent As String
        Dim blnIsState As Boolean, blnHasState As Boolean, strStateControl As String, varStateCode As Variant
        Dim lngRecs As Long
        Dim blnAdd As Boolean, blnRetVal As Boolean, blnUpdate As Boolean
        Dim intPos01 As Integer
        Dim varTmp00 As Variant
        Dim lngX As Long

6610    blnUpdate = True
6620    blnHasState = False: strStateControl = vbNullString  ' ** There can be only 1 .._State_Pref per form.

        ' ** Determine whether we're dealing with a subform.
6630    intPos01 = InStr(strFormName, "_Sub")
6640    If intPos01 > 0 Then
6650      strFormName_Parent = Left(strFormName, (intPos01 - 1))
6660    Else
6670      strFormName_Parent = vbNullString
6680    End If

6690    Set dbs = CurrentDb
6700    With dbs
          ' ** Preferences available, specified by [frmnam].
6710      Set qdf1 = .QueryDefs("qryPreferences_01b")  '##dbs_id
6720      With qdf1.Parameters
6730        ![frmnam] = strFormName  ' ** This remains listed as the actual form.
6740      End With
6750      Select Case strFormName_Parent
          Case vbNullString
6760        blnRetVal = IsLoaded(strFormName)  ' ** Module Function: modFileUtilities.
6770      Case Else
6780        blnRetVal = IsLoaded(strFormName_Parent)  ' ** Module Function: modFileUtilities.
6790      End Select
6800      If blnRetVal = True Then  ' ** Module Function: modFileUtilities.
6810        Set rst1 = qdf1.OpenRecordset
6820        If rst1.BOF = True And rst1.EOF = True Then
              ' ** No preferences designated for this form.
6830          rst1.Close
6840        Else
6850          rst1.MoveLast
6860          lngRecs = rst1.RecordCount
6870          rst1.MoveFirst
6880          Select Case strFormName_Parent
              Case vbNullString
6890            Set frm = Forms(strFormName)
6900          Case Else
6910            Set frm = Forms(strFormName_Parent).Controls(strFormName).Form  ' ** Set this to the subform.
6920          End Select
6930          For lngX = 1& To lngRecs
6940            blnAdd = False: blnIsState = False
6950            If InStr(rst1![ctl_name], "_State_Pref") > 0 Then
6960              blnIsState = True
6970              strStateControl = rst1![ctl_name]
6980            End If
                ' ** tblTemplate_Preference_User, with frm_name, ctl_name.
6990            Set qdf2 = .QueryDefs("qryPreferences_15_02")  '##dbs_id
7000            Set rst2 = qdf2.OpenRecordset
7010            With rst2
7020              If .BOF = True And .EOF = True Then
                    ' ** None saved whatsoever!
7030                blnAdd = True
7040              Else
7050                .FindFirst "[frm_name] = '" & rst1![frm_name] & "' And [ctl_name] = '" & rst1![ctl_name] & "' And " & _
                      "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
7060                If .NoMatch = True Then
7070                  blnAdd = True
7080                Else
7090  On Error Resume Next
7100                  varTmp00 = frm.Controls(rst1![ctl_name])
7110                  If ERR.Number <> 0 Then
7120  On Error GoTo ERRH
                        ' ** Error: 2427  You entered an expression that has no value.
7130                    blnUpdate = False
7140                  Else
7150  On Error GoTo ERRH
7160                  End If
7170                  If blnUpdate = True Then
7180                    .Edit
7190                    Select Case rst1![datatype_db_type]
                        Case dbBoolean
7200                      If IsNull(frm.Controls(rst1![ctl_name])) = True Then
7210                        ![prefuser_boolean] = False
7220                      Else
7230                        ![prefuser_boolean] = frm.Controls(rst1![ctl_name])
7240                      End If
7250                    Case dbInteger
7260                      ![prefuser_integer] = frm.Controls(rst1![ctl_name])
7270                    Case dbLong
7280                      ![prefuser_long] = frm.Controls(rst1![ctl_name])
7290                    Case dbCurrency
7300                      ![prefuser_currency] = frm.Controls(rst1![ctl_name])
7310                    Case dbSingle
7320                      ![prefuser_single] = frm.Controls(rst1![ctl_name])
7330                    Case dbDouble
7340                      ![prefuser_double] = frm.Controls(rst1![ctl_name])
7350                    Case dbDate
7360                      ![prefuser_date] = frm.Controls(rst1![ctl_name])
7370                    Case dbText
7380                      If IsNull(frm.Controls(rst1![ctl_name])) = True Then
7390                        Select Case blnIsState
                            Case True
                              ' ** Don't Null-out any existing state code prefs.
7400                        Case False
7410                          ![prefuser_text] = Null
7420                        End Select
7430                      Else
7440                        If Trim(frm.Controls(rst1![ctl_name])) = vbNullString Then
7450                          Select Case blnIsState
                              Case True
                                ' ** Don't Null-out any existing state code prefs.
7460                          Case False
7470                            ![prefuser_text] = Null
7480                          End Select
7490                        Else
7500                          ![prefuser_text] = frm.Controls(rst1![ctl_name])
7510                          If blnIsState = True Then
7520                            blnHasState = True
7530                            varStateCode = frm.Controls(rst1![ctl_name])
7540                          End If
7550                        End If
7560                      End If
7570                    End Select
7580                    ![DateModified] = Now()
7590                    .Update
7600                  End If  ' ** blnUpdate.
7610                End If  ' ** NoMatch.
7620              End If  ' ** BOF, EOF.
7630              If blnAdd = True Then
7640  On Error Resume Next
7650                varTmp00 = frm.Controls(rst1![ctl_name])
7660                If ERR.Number <> 0 Then
7670  On Error GoTo ERRH
                      ' ** Error: 2427  You entered an expression that has no value.
7680                  blnUpdate = False
7690                Else
7700  On Error GoTo ERRH
7710                End If
7720                If blnUpdate = True Then
7730                  .AddNew
7740                  ![dbs_id] = rst1![dbs_id]
7750                  ![frm_id] = rst1![frm_id]
7760                  ![ctl_id] = rst1![ctl_id]
7770                  ![prefctl_id] = rst1![prefctl_id]
7780                  Select Case rst1![datatype_db_type]
                      Case dbBoolean
7790                    If IsNull(frm.Controls(rst1![ctl_name])) = True Then
7800                      ![prefuser_boolean] = False
7810                    Else
7820                      ![prefuser_boolean] = frm.Controls(rst1![ctl_name])
7830                    End If
7840                  Case dbInteger
7850                    ![prefuser_integer] = frm.Controls(rst1![ctl_name])
7860                  Case dbLong
7870                    ![prefuser_long] = frm.Controls(rst1![ctl_name])
7880                  Case dbCurrency
7890                    ![prefuser_currency] = frm.Controls(rst1![ctl_name])
7900                  Case dbSingle
7910                    ![prefuser_single] = frm.Controls(rst1![ctl_name])
7920                  Case dbDouble
7930                    ![prefuser_double] = frm.Controls(rst1![ctl_name])
7940                  Case dbDate
7950                    ![prefuser_date] = frm.Controls(rst1![ctl_name])
7960                  Case dbText
7970                    If IsNull(frm.Controls(rst1![ctl_name])) = True Then
7980                      ![prefuser_text] = Null
7990                    Else
8000                      If Trim(frm.Controls(rst1![ctl_name])) = vbNullString Then
8010                        ![prefuser_text] = Null
8020                      Else
8030                        ![prefuser_text] = frm.Controls(rst1![ctl_name])
8040                        If blnIsState = True Then
8050                          blnHasState = True
8060                          varStateCode = frm.Controls(rst1![ctl_name])
8070                        End If
8080                      End If
8090                    End If
8100                  End Select
8110                  ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
8120                  ![DateCreated] = Now()
8130                  ![DateModified] = Now()
8140                  .Update
8150                End If  ' ** blnUpdate.
8160              End If  ' ** blnAdd.
8170              .Close
8180            End With  ' ** rst2.
8190            If lngX < lngRecs Then rst1.MoveNext
8200          Next  ' ** lngX.
8210          rst1.Close
8220          If blnHasState = True Then
8230            Pref_State varStateCode, strFormName, strStateControl  ' ** Procedure: Below.
8240          End If
8250        End If  ' ** BOF, EOF.
8260      End If  ' ** blnRetVal.
8270      .Close
8280    End With  ' ** dbs.

EXITP:
8290    Set ctl = Nothing
8300    Set frm = Nothing
8310    Set rst1 = Nothing
8320    Set rst2 = Nothing
8330    Set qdf1 = Nothing
8340    Set qdf2 = Nothing
8350    Set dbs = Nothing
8360    Exit Sub

ERRH:
8370    Select Case ERR.Number
        Case 2310  ' ** You entered an expression that has no value.
          ' ** It closed before the procedure could complete.
8380    Case 2450  ' ** Microsoft Access can't find the form '|' referred to in a macro expression or Visual Basic code.
          ' ** It closed before the procedure could complete.
8390    Case Else
8400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8410    End Select
8420    Resume EXITP

End Sub

Public Function Pref_ReportPath(varURP As Variant, strFormName As String) As Variant
' ** if the user has a UserReportPath preference from another screen, use that.

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_ReportPath"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRptPath As String
        Dim lngRecs As Long
        Dim blnContinue As Boolean
        Dim varRetVal As Variant

8510    varRetVal = Null

8520    blnContinue = True

8530    If IsNull(varURP) = False Then
8540      If Trim(varURP) <> vbNullString Then
8550        blnContinue = False
8560        varRetVal = varURP
8570      End If
8580    End If

8590    If blnContinue = True Then
8600      Set dbs = CurrentDb
8610      With dbs
            ' ** qryReport_Path_02 (qryReport_Path_01 (tblPreference_Control, just 'UserReportPath'),
            ' ** linked to tblPreference_User, by specified [usr]), grouped by prefuser_text.
8620        Set qdf = .QueryDefs("qryReport_Path_03")
8630        With qdf.Parameters
8640          ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
8650        End With
8660        Set rst = qdf.OpenRecordset
8670        With rst
8680          If .BOF = True And .EOF = True Then
                ' ** Nope.
8690          Else
8700            .MoveLast
8710            lngRecs = .RecordCount
8720            .MoveFirst
8730            If lngRecs = 1& Then
                  ' ** If they only have one, or all of them are the same, use it!
8740              strRptPath = ![prefuser_text]
8750            End If
8760          End If
8770          .Close
8780        End With
8790        .Close
8800      End With
8810      If strRptPath <> vbNullString Then
8820        varRetVal = strRptPath
8830      End If
8840    End If

EXITP:
8850    Set rst = Nothing
8860    Set qdf = Nothing
8870    Set dbs = Nothing
8880    Pref_ReportPath = varRetVal
8890    Exit Function

ERRH:
8900    varRetVal = RET_ERR
8910    Select Case ERR.Number
        Case Else
8920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8930    End Select
8940    Resume EXITP

End Function

Public Function Pref_HasPref(strFrmName As String, strCtlName As String) As Boolean
' ** Determine whether the user has ever had a preference for this control.

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_HasPref"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngThisDbsID As Long
        Dim blnRetVal As Boolean

9010    blnRetVal = False

9020    If strFrmName <> vbNullString And strCtlName <> vbNullString Then
9030      lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
9040      Set dbs = CurrentDb
9050      With dbs
            ' ** tblPreference_Control, linked to tblPreference_User, by specified [dbid], [fnam], [cnam], [usr].
9060        Set qdf = .QueryDefs("qryPreferences_11_01")
9070        With qdf.Parameters
9080          ![dbid] = lngThisDbsID
9090          ![fnam] = strFrmName
9100          ![cnam] = strCtlName
9110          ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
9120        End With
9130        Set rst = qdf.OpenRecordset
9140        With rst
9150          If .BOF = True And .EOF = True Then
                ' ** Nope.
9160          Else
9170            .MoveFirst
9180            If IsNull(![prefuser_id]) = False Then
9190              blnRetVal = True
9200            End If
9210          End If
9220          .Close
9230        End With
9240        .Close
9250      End With
9260    End If

EXITP:
9270    Set rst = Nothing
9280    Set qdf = Nothing
9290    Set dbs = Nothing
9300    Pref_HasPref = blnRetVal
9310    Exit Function

ERRH:
9320    blnRetVal = False
9330    Select Case ERR.Number
        Case Else
9340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9350    End Select
9360    Resume EXITP

End Function

Public Function Pref_CurrID() As Boolean
' ** Returns whether the user has currency turned on for transaction maintenance.

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_CurrID"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

9410    blnRetVal = False

9420    Set dbs = CurrentDb
9430    With dbs
          ' ** tblPreference_User, for 'chkIncludeCurrency' on 'frmPostingDate',
          ' ** just dbs_id = 1, by specified [usr].
9440      Set qdf = .QueryDefs("qryPreferences_07_01")
9450      With qdf.Parameters
9460        ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
9470      End With
9480      Set rst = qdf.OpenRecordset
9490      With rst
9500        If .BOF = True And .EOF = True Then
              ' ** No preference!
9510        Else
9520          .MoveFirst
9530          blnRetVal = ![prefuser_boolean]
9540        End If
9550        .Close
9560      End With
9570      .Close
9580    End With

EXITP:
9590    Set rst = Nothing
9600    Set qdf = Nothing
9610    Set dbs = Nothing
9620    Pref_CurrID = blnRetVal
9630    Exit Function

ERRH:
9640    blnRetVal = False
9650    Select Case ERR.Number
        Case Else
9660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9670    End Select
9680    Resume EXITP

End Function

Public Function Pref_Suppress() As Boolean
' ** Return Values:
' **   True: Preference chkDefaultSuppress is checked.
' **   False: Preference chkOptionSuppress is checked.

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_Suppress"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim blnAddAll As Boolean, blnHasDef As Boolean, blnHasOpt As Boolean
        Dim blnTmp01 As Boolean, blnTmp02 As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

9710    blnRetVal = True

9720    Set dbs = CurrentDb
9730    With dbs

          ' ** tblPreference_User, for 'chkDefaultSuppress', 'chkOptionSuppress',
          ' ** just dbs_id = 1, by specified [usr].
9740      Set qdf = .QueryDefs("qryPreferences_08_01")
9750      With qdf.Parameters
9760        ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
9770      End With
9780      Set rst = qdf.OpenRecordset
9790      blnAddAll = False: blnHasDef = False: blnHasOpt = False
9800      blnTmp01 = False: blnTmp02 = False
9810      With rst
9820        If .BOF = True And .EOF = True Then
9830          blnAddAll = True
9840        Else
9850          .MoveLast
9860          lngRecs = .RecordCount
9870          .MoveFirst
9880          For lngX = 1& To lngRecs
9890            Select Case ![ctl_name]
                Case "chkDefaultSuppress"
9900              blnHasDef = True
9910              blnTmp01 = ![prefuser_boolean]
9920            Case "chkOptionSuppress"
9930              blnHasOpt = True
9940              blnTmp02 = ![prefuser_boolean]
9950            End Select
9960            If lngX < lngRecs Then .MoveNext
9970          Next
9980        End If
9990        .Close
10000     End With
10010     Set rst = Nothing
10020     Set qdf = Nothing
10030     DoEvents

10040     If blnAddAll = True Then
10050       blnTmp01 = True
            ' ** Append 'chkDefaultSuppress' record to tblPreference_User,
            ' ** just dbs_id = 1, by specified [usr], [pbln].
10060       Set qdf = .QueryDefs("qryPreferences_08_02")
10070       With qdf.Parameters
10080         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10090         ![pbln] = blnTmp01  ' ** Default to suppress.
10100       End With
10110       qdf.Execute
10120       Set qdf = Nothing
10130       blnTmp02 = False
            ' ** Append 'chkOptionSuppress' record to tblPreference_User,
            ' ** just dbs_id = 1, by specified [usr], [pbln].
10140       Set qdf = .QueryDefs("qryPreferences_08_03")
10150       With qdf.Parameters
10160         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10170         ![pbln] = blnTmp02
10180       End With
10190       qdf.Execute
10200       Set qdf = Nothing
10210     ElseIf blnHasDef = False Then
10220       blnTmp01 = (Not blnTmp02)
            ' ** Append 'chkDefaultSuppress' record to tblPreference_User,
            ' ** just dbs_id = 1, by specified [usr], [pbln].
10230       Set qdf = .QueryDefs("qryPreferences_08_02")
10240       With qdf.Parameters
10250         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10260         ![pbln] = blnTmp01
10270       End With
10280       qdf.Execute
10290       Set qdf = Nothing
10300     ElseIf blnHasOpt = False Then
10310       blnTmp02 = (Not blnTmp01)
            ' ** Append 'chkOptionSuppress' record to tblPreference_User,
            ' ** just dbs_id = 1, by specified [usr], [pbln].
10320       Set qdf = .QueryDefs("qryPreferences_08_03")
10330       With qdf.Parameters
10340         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10350         ![pbln] = blnTmp02
10360       End With
10370       qdf.Execute
10380       Set qdf = Nothing
10390     End If
10400     DoEvents

10410     If blnTmp01 = True Then
10420       blnRetVal = True
10430     ElseIf blnTmp02 = True Then
10440       blnRetVal = False
10450     End If

10460     .Close
10470   End With

EXITP:
10480   Set rst = Nothing
10490   Set qdf = Nothing
10500   Set dbs = Nothing
10510   Pref_Suppress = blnRetVal
10520   Exit Function

ERRH:
10530   blnRetVal = True
10540   Select Case ERR.Number
        Case Else
10550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10560   End Select
10570   Resume EXITP

End Function

Public Function Pref_GetBln(strFrmName As String, strCtlName As String) As Boolean
' ** Return the value of one specified Boolean preference.

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "Pref_GetBln"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngThisDbsID As Long
        Dim blnRetVal As Boolean

10610   blnRetVal = False

10620   If strFrmName <> vbNullString And strCtlName <> vbNullString Then
10630     lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
10640     Set dbs = CurrentDb
10650     With dbs
            ' ** tblPreference_Control, linked to tblPreference_User, with
            ' ** prefuser_boolean, by specified [dbid], [fnam], [cnam], [usr].
10660       Set qdf = .QueryDefs("qryPreferences_11_02")
10670       With qdf.Parameters
10680         ![dbid] = lngThisDbsID
10690         ![fnam] = strFrmName
10700         ![cnam] = strCtlName
10710         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10720       End With
10730       Set rst = qdf.OpenRecordset
10740       With rst
10750         If .BOF = True And .EOF = True Then
                ' ** None found.
10760         Else
10770           .MoveFirst
10780           If IsNull(![prefuser_id]) = False Then
10790             blnRetVal = ![prefuser_boolean]
10800           End If
10810         End If
10820         .Close
10830       End With
10840       .Close
10850     End With
10860   End If

EXITP:
10870   Set rst = Nothing
10880   Set qdf = Nothing
10890   Set dbs = Nothing
10900   Pref_GetBln = blnRetVal
10910   Exit Function

ERRH:
10920   blnRetVal = False
10930   Select Case ERR.Number
        Case Else
10940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10950   End Select
10960   Resume EXITP

End Function

Public Function StateCodeQrySet(frm As Access.Form) As Boolean

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "StateCodeQrySet"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim grp As DAO.Group, usr As DAO.User
        Dim strFormName As String, strCurrentUser As String
        Dim strQryName_DefUser As String, strQryName_ThisUser As String
        Dim blnAdmin As Boolean
        Dim lngRecs As Long, lngCnt As Long
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

11010   blnRetVal = False
11020   strFormName = vbNullString

11030   strFormName = frm.Name
11040   If strFormName <> vbNullString Then
11050     Set dbs = CurrentDb
11060     With dbs

            ' ** qryState_09 (tblPreference_User, linked to tblPreference_Control,
            ' ** just 'state_query' control, by specified GetDefaultUser()), with
            ' ** qryState_06 (qryState_05 (qryState_04 (tblQuery, just state combo
            ' ** box queries), with Boolean code types), just default query of
            ' ** states only), default pref and current pref; Cartesian.
11070       Set qdf = .QueryDefs("qryState_10")  '##dbs_id
11080       Set rst = qdf.OpenRecordset
11090       If rst.BOF = True And rst.EOF = True Then
11100         rst.Close
              ' ** Append default state_query pref to tblPreference_User.
11110         Set qdf = .QueryDefs("qryState_11")  '##dbs_id
11120         qdf.Execute
11130         DoEvents
11140       Else
11150         With rst
11160           .MoveFirst
11170           Select Case IsNull(![rowsrc_rowsource])
                Case True
11180             strQryName_DefUser = ![rowsrc_rowsource_default]
11190           Case False
11200             strQryName_DefUser = ![rowsrc_rowsource]
11210           End Select
11220           .Close
11230         End With  ' ** rst.
11240       End If

            ' ** If this user is in admins, and her pref is different
            ' ** from the default user's, edit the default user to match.
11250       strCurrentUser = CurrentUser()  ' ** Internal Access Function: Trust Accountant login.
11260       strQryName_ThisUser = vbNullString
11270       If strCurrentUser <> GetDefaultUser Then  ' ** Module Function: modFileUtilities.

11280         blnAdmin = False
11290         For Each grp In DBEngine.Workspaces(0).Groups
11300           If grp.Name = "Admins" Then
11310             For Each usr In grp.Users
11320               If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
11330                 blnAdmin = True
11340                 Exit For
11350               End If
11360             Next
11370           End If
11380         Next

11390         If blnAdmin = True Then
                ' ** tblPreference_Control, linked to tblPreference_User, for 'state_query', by specified [usr].
11400           varTmp00 = DLookup("[prefuser_text]", "qryState_14", "[Username] = '" & strCurrentUser & "'")  '##dbs_id
11410           Select Case IsNull(varTmp00)
                Case True
                  ' ** No pref for this user, so just use default user pref.
11420           Case False
11430             strQryName_ThisUser = varTmp00
11440             If ((strQryName_ThisUser <> vbNullString) And (strQryName_ThisUser <> strQryName_DefUser)) Then
                    ' ** Update tblPreference_User, for 'state_query', by specified [usr], [ctlname, [usrtxt].
11450               Set qdf = .QueryDefs("qryState_15")  '##dbs_id
11460               With qdf.Parameters
11470                 ![usr] = GetDefaultUser()  ' ** Module Function: modFileUtilities
11480                 ![ctlnam] = "state_query"
11490                 ![usrtxt] = strQryName_ThisUser
11500               End With
11510               qdf.Execute
11520               DoEvents
11530             End If
11540           End Select
11550         End If  ' ** blnAdmin.

11560       End If  ' ** strCurrentUser.

            ' ** qryState_07 (tblForm_Control_RowSource, with qryState_06 (qryState_05 (qryState_04
            ' ** (tblQuery, just state combo box queries), with Boolean code types), just default
            ' ** query of states only), just state RowSource's, with rowsrc_rowsource_default; Cartesian),
            ' ** linked to qryState_12 (qryState_08 (qryState_07 (tblForm_Control_RowSource,
            ' ** with qryState_06 (qryState_05 (qryState_04 (tblQuery, just state combo box queries),
            ' ** with Boolean code types), just default query of states only), just state RowSource's,
            ' ** with rowsrc_rowsource_default; Cartesian), grouped by frm_name, frm_parent_sub, with cnt),
            ' ** with qryState_10 (qryState_09 (tblPreference_User, linked to tblPreference_Control,
            ' ** just 'state_query' control, by specified GetDefaultUser()), with
            ' ** qryState_06 (qryState_05 (qryState_04 (tblQuery, just state combo
            ' ** box queries), with Boolean code types), just default query of
            ' ** states only), default pref and current pref; Cartesian), with rowsrc_rowsource_new;
            ' ** Cartesian), all controls, with current pref.
11570       Set qdf = .QueryDefs("qryState_13")  '##dbs_id
11580       Set rst = qdf.OpenRecordset
11590       With rst
11600         If .BOF = True And .EOF = True Then
                ' ** After all this?
11610         Else
11620           .MoveLast
11630           lngRecs = .RecordCount
11640           .MoveFirst
11650           lngCnt = 0&
                'THIS MISSES THE SUBFORMS WHEN FED THE PARENT!
11660           For lngX = 1& To lngRecs
11670             If strFormName = ![frm_name] Then
11680               blnRetVal = True
11690               If ![cnt] = 1 Then
11700                 lngCnt = lngCnt + 1&
11710                 If frm.Controls(![ctl_name]).RowSource <> ![rowsrc_rowsource_new] Then
11720                   frm.Controls(![ctl_name]).RowSource = ![rowsrc_rowsource_new]
11730                 End If
11740                 Exit For
11750               Else
11760                 lngCnt = lngCnt + 1&
11770                 If frm.Controls(![ctl_name]).RowSource <> ![rowsrc_rowsource_new] Then
11780                   frm.Controls(![ctl_name]).RowSource = ![rowsrc_rowsource_new]
11790                 End If
11800                 If lngCnt = ![cnt] Then
11810                   Exit For
11820                 End If
11830               End If
11840             Else
11850               If IsNull(![frm_parent_sub]) = False Then
11860                 If strFormName = ![frm_parent_sub] Then
11870                   blnRetVal = True
11880                   If ![cnt] = 1 Then
11890                     lngCnt = lngCnt + 1&
11900                     If frm.Controls(![frm_name]).Form.Controls(![ctl_name]).RowSource <> ![rowsrc_rowsource_new] Then
11910                       frm.Controls(![frm_name]).Form.Controls(![ctl_name]).RowSource = ![rowsrc_rowsource_new]
11920                     End If
11930                     Exit For
11940                   Else
11950                     lngCnt = lngCnt + 1&
11960                     If frm.Controls(![frm_name]).Form.Controls(![ctl_name]).RowSource <> ![rowsrc_rowsource_new] Then
11970                       frm.Controls(![frm_name]).Form.Controls(![ctl_name]).RowSource = ![rowsrc_rowsource_new]
11980                     End If
11990                     If lngCnt = ![cnt] Then
12000                       Exit For
12010                     End If
12020                   End If
12030                 End If
12040               End If
12050             End If
12060             If lngX < lngRecs Then .MoveNext
12070           Next  ' ** lngX.
12080         End If
12090         .Close
12100       End With  ' ** rst.

12110       .Close
12120     End With  ' ** dbs.
12130   End If

EXITP:
12140   Set usr = Nothing
12150   Set grp = Nothing
12160   Set rst = Nothing
12170   Set qdf = Nothing
12180   Set dbs = Nothing
12190   StateCodeQrySet = blnRetVal
12200   Exit Function

ERRH:
12210   DoCmd.Hourglass False
12220   blnRetVal = False
12230   Select Case ERR.Number
        Case Else
12240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12250   End Select
12260   Resume EXITP

End Function
