Attribute VB_Name = "modStartupFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modStartupFuncs"

'VGC 09/24/2017: CHANGES!

' ** CHECK TrustAux.mdb LOCATION, BELOW!!
' **   1730        strAuxiliaryDatabase = gstrDir_Dev & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & gstrFile_AuxDataName

' ** THIS HAS THAT SINGLE BEEP THAT SOUNDS DURING INITIALIZATION! SEE 'BEEP HERE!

'Public gdatLoadStart As Date, gdatLoadEnd As Date

' ** Progress Bar Variables.
Private frmPBar As Access.Form
Private ctlPBox As Access.Rectangle
Private dblPBar_MaxWidth As Double
Private dblPBar_CurWidth As Double
Private dblPBar_SubSteps As Double
Private dblPBar_ThisStep As Double
Private dblPBar_ThisSubStep As Double
Private dblPBar_ThisIncrement As Double
Private dblPBar_ThisSubIncrement As Double
Private blnPBar As Boolean

' ** Array: arr_varPBar_Step().
Private dblPBar_Steps As Double, arr_varPBar_Step() As Variant
Private Const PB_ELEMS As Integer = 4  ' ** Array's first-element UBound().
Private Const PB_STEP As Integer = 0
Private Const PB_NAME As Integer = 1
Private Const PB_PCT  As Integer = 2
Private Const PB_INCR As Integer = 3
Private Const PB_SUBS As Integer = 4

' ** Array: arr_varLink().
'Private Const L_ID   As Integer = 0
Private Const L_NAM  As Integer = 1
'Private Const L_AUT  As Integer = 2
'Private Const L_ORD  As Integer = 3
'Private Const L_NEW  As Integer = 4
'Private Const L_ACT  As Integer = 5
Private Const L_DTA  As Integer = 6
Private Const L_ARC  As Integer = 7
Private Const L_AUX  As Integer = 8
Private Const L_SRC  As Integer = 9
'Private Const L_TYP  As Integer = 10
'Private Const L_LNKT As Integer = 11
Private Const L_LNKC As Integer = 12
Private Const L_FND  As Integer = 13
Private Const L_FIX  As Integer = 14

Private blnWindowVisible As Boolean
' **

Public Function InitializeTables() As Boolean
' ** Called by:
' **   frmLinkData.cmdRelink_Click()
' **   frmMenu_Title.Form_Open()
' **   frmMenu_Title.TA_Logo_img_MouseUp()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "InitializeTables"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, tdf As DAO.TableDef, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst As DAO.Recordset
        Dim rstProg As DAO.Recordset, rstData As DAO.Recordset, frm As Access.Form
        Dim strDatabase As String, strArchiveDatabase As String, strAuxiliaryDatabase As String
        Dim strPath As String, strSpecPathFile As String
        Dim strIs_License_String As String, strVerNum As String
        Dim dblFullVerD As Double, dblFullVerP As Double
        Dim strMsg As String, strTitle1 As String
        Dim lngLinks As Long, arr_varLink As Variant
        Dim lngDels As Long, arr_varDel() As Variant
        Dim lngLinkCnt As Long
        Dim blnLinkSuccess As Boolean, blnNoVersion As Boolean, blnFound As Boolean
        Dim blnDoFrm As Boolean, blnAuxLoc As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intPos01 As Integer, intPos02 As Integer
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, dblZ As Double, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varDel().
        Const D_ELEMS As Integer = 1
        Const D_NAM As Integer = 0
        Const D_FND As Integer = 1

110     Select Case IsLoaded("frmMenu_Title", acForm)  ' ** Module Function: modFileUtilities.
        Case True
120       Set frm = Forms("frmMenu_Title")
130       blnDoFrm = True
140     Case False
150       blnDoFrm = False
160     End Select

        ' ** Get the data location from DDTrust.ini.
170     gstrTrustDataLocation = "NONE"
180     strDatabase = "NONE"
190     strArchiveDatabase = "NONE"
200     strAuxiliaryDatabase = "NONE"
210     blnRetVal = IniFile_GetDataLoc  ' ** Function: Below.
220     blnAuxLoc = False

        ' ****************************************
        ' ** Progress Bar setup.
        ' ****************************************
230     blnPBar = IsLoaded("frmLinkData", acForm)  ' ** Module Function: modFileUtilities.
240     If blnPBar = True Then

250       Set frmPBar = Forms("frmLinkData")
260       With frmPBar
270         Set ctlPBox = .ProgBar_box
280         dblPBar_MaxWidth = ctlPBox.Width
290         If ctlPBox.Visible = False Then ctlPBox.Visible = True
300         .ProgBar_Width_Link True, 1  ' ** Form Procedure: frmLinkData.
310       End With
320       DoEvents

330       dblPBar_Steps = 23#
340       ReDim arr_varPBar_Step(PB_ELEMS, dblPBar_Steps)  ' ** We won't be using Zero.
350       For dblZ = 1# To dblPBar_Steps
360         arr_varPBar_Step(PB_STEP, dblZ) = dblZ
370         Select Case dblZ
            Case 1#
              ' ** 1. INI file 1.
380           arr_varPBar_Step(PB_NAME, dblZ) = "INI file 1."
390           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.01)
400           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
410         Case 2#
              ' ** 2. Data paths.
420           arr_varPBar_Step(PB_NAME, dblZ) = "Data paths."
430           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
440           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
450         Case 3#
              ' ** 3. Data file check.
460           arr_varPBar_Step(PB_NAME, dblZ) = "Data file check."
470           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
480           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
490         Case 4#
              ' ** 4. Icon check.
500           arr_varPBar_Step(PB_NAME, dblZ) = "Icon check."
510           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.01)
520           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
530         Case 5#
              ' ** 5. EULA check (Demo check 1).
540           arr_varPBar_Step(PB_NAME, dblZ) = "EULA check."
550           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
560           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
570         Case 6#
              ' ** 6. INI file 2.
580           arr_varPBar_Step(PB_NAME, dblZ) = "INI file 2."
590           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
600           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
610         Case 7#
              ' ** 7. File extention check.
620           arr_varPBar_Step(PB_NAME, dblZ) = "File extention check."
630           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.01)
640           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
650         Case 8#
              ' ** 8. Link list.
660           arr_varPBar_Step(PB_NAME, dblZ) = "Link list."
670           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
680           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
690         Case 9#
              ' ** 9. Version check 1.
700           arr_varPBar_Step(PB_NAME, dblZ) = "Version check 1."
710           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
720           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
730         Case 10#
              ' ** 10. Version check 2.
740           arr_varPBar_Step(PB_NAME, dblZ) = "Version check 2."
750           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
760           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
770         Case 11#
              ' ** 11. Current links.
780           arr_varPBar_Step(PB_NAME, dblZ) = "Current links."
790           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.05)
800           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(True)
810         Case 12#
              ' ** 12. Extra links.
820           arr_varPBar_Step(PB_NAME, dblZ) = "Extra links."
830           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.05)
840           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(True)
850         Case 13#
              ' ** 13. Link TrustDta.mdb.
860           arr_varPBar_Step(PB_NAME, dblZ) = "Link TrustDta.mdb."
870           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.26)
880           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(True)
890         Case 14#
              ' ** 14. Link TrstArch.mdb.
900           arr_varPBar_Step(PB_NAME, dblZ) = "Link TrstArch.mdb."
910           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.04)
920           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(True)
930         Case 15#
              ' ** 15. Link TrustAux.mdb.
940           arr_varPBar_Step(PB_NAME, dblZ) = "Link TrustAux.mdb."
950           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.26)
960           arr_varPBar_Step(PB_SUBS, dblZ) = CBool(True)
970         Case 16#
              ' ** 16. Hidden check.
980           arr_varPBar_Step(PB_NAME, dblZ) = "Hidden check."
990           arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
1000          arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
1010        Case 17#
              ' ** 17. License check 1.
1020          arr_varPBar_Step(PB_NAME, dblZ) = "License check."
1030          arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
1040          arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
1050        Case 18#
              ' ** 18. Demo check 2.
1060          arr_varPBar_Step(PB_NAME, dblZ) = "Demo check."
1070          arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
1080          arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
1090        Case 19#
              ' ** 19. Conversion check.
1100          arr_varPBar_Step(PB_NAME, dblZ) = "Conversion check."
1110          arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
1120          arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
1130        Case 20#
              ' ** 20. License form.
1140          arr_varPBar_Step(PB_NAME, dblZ) = "License form."
1150          arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
1160          arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
1170        Case 21#
              ' ** 21. Table update.
1180          arr_varPBar_Step(PB_NAME, dblZ) = "Table update."
1190          arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02)
1200          arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
1210        Case 22#
              ' ** 22. Foreign currency check.
1220          arr_varPBar_Step(PB_NAME, dblZ) = "Foreign currency check."
1230          arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.04)
1240          arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
1250        Case 23#
              ' ** 22. Finished.
1260          arr_varPBar_Step(PB_NAME, dblZ) = "Finished."
1270          arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.01)
1280          arr_varPBar_Step(PB_SUBS, dblZ) = CBool(False)
1290        End Select
1300        arr_varPBar_Step(PB_INCR, dblZ) = (arr_varPBar_Step(PB_PCT, dblZ) * dblPBar_MaxWidth)
1310      Next  ' ** dblPBar_Steps: dblZ.

          ' ** 1  arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.01) 1
          ' ** 2  arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 4
          ' ** 3  arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 5
          ' ** 4  arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.01) 6
          ' ** 5  arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 8
          ' ** 6  arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 10
          ' ** 7  arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.01) 11
          ' ** 8  arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 13
          ' ** 9  arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 15
          ' ** 10 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 17
          ' ** 11 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.05) 22
          ' ** 12 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.05) 27
          ' ** 13 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.26) 53  'TrustDta
          ' ** 14 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.04) 57
          ' ** 15 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.26) 83  'TrustAux
          ' ** 16 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 85
          ' ** 17 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 87
          ' ** 18 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 89
          ' ** 19 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 91
          ' ** 20 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 93
          ' ** 21 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.02) 95
          ' ** 22 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.04) 99  'Currency
          ' ** 23 arr_varPBar_Step(PB_PCT, dblZ) = CDbl(0.01) 100

1320      dblPBar_SubSteps = 0#
1330      dblPBar_CurWidth = 0#
1340      dblPBar_ThisSubStep = 0#
1350      dblPBar_ThisIncrement = 0#
1360      dblPBar_ThisSubIncrement = 0#

1370    End If  ' ** blnPBar.
        ' ****************************************

        ' ******************************
1380    If blnPBar = True Then
          ' ** 1. INI file 1.
1390      dblPBar_ThisStep = 1#
1400      dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
1410      If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
            ' ** None.
1420      Else
1430        dblPBar_SubSteps = 0#
1440        dblPBar_ThisSubStep = 0#
1450        dblPBar_ThisSubIncrement = 0#
1460        dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
1470        frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
1480      End If

1490      DoEvents
1500    End If  ' ** blnPBar.

        ' ******************************

        ' ** STATUS SUB 1: DDTrust.ini
1510    If blnDoFrm = True Then
1520      frm.InitMsg_lbl3.Caption = "DDTrust.ini Check"
1530      frm.InitMsg_lbl3.Visible = True
1540      DoEvents
1550    End If

        ' ** If the initial IniFile_GetDataLoc() failed,
        ' ** make sure gstrTrustDataLocation doesn't have a message instead of a path.
        ' ** If it has a message (and I haven't traced where or how it might get that message),
        ' ** write the message to DDTrust.ini (for what purpose, I have no idea!).
1560    If blnRetVal = False Then
1570      If gstrTrustDataLocation = "INSTALL" Or gstrTrustDataLocation = "NONE" Or gstrTrustDataLocation = RET_ERR Then
1580        If IniFile_Set("Files", "Location", gstrTrustDataLocation, CurrentAppPath & LNK_SEP & gstrFile_INI) = False Then  ' ** Module Function: modStartupFuncs, modFileUtilities.
1590          blnRetVal = False
1600          msgResponse = MsgBox("Unable to write INI file.", vbCritical + vbOKOnly, "Error")
1610        Else
1620          blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
1630        End If
1640      End If
1650    End If

1660    If blnRetVal = True Then

          ' ******************************
1670      If blnPBar = True Then
            ' ** 2. Data paths.
1680        dblPBar_ThisStep = 2#
1690        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
1700        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
1710        Else
1720          dblPBar_SubSteps = 0#
1730          dblPBar_ThisSubStep = 0#
1740          dblPBar_ThisSubIncrement = 0#
1750          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
1760          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
1770        End If
1780        DoEvents
1790      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 2: Path Check
1800      If blnDoFrm = True Then
1810        frm.InitMsg_lbl3.Caption = "Path Check"
1820        DoEvents
1830      End If

          ' ** If gstrTrustDataLocation has a message, get a path a different way.
1840      If gstrTrustDataLocation = "INSTALL" Or gstrTrustDataLocation = "NONE" Or gstrTrustDataLocation = RET_ERR Then
1850        If gblnDev_Debug Then
1860          gstrTrustDataLocation = gstrDir_Dev
1870        Else
1880          gstrTrustDataLocation = CurrentAppPath  ' ** Module Function: modFileUtilities.
1890        End If
1900        gstrTrustDataLocation = gstrTrustDataLocation & LNK_SEP
1910      End If
1920      strDatabase = gstrTrustDataLocation & gstrFile_DataName
1930      strArchiveDatabase = gstrTrustDataLocation & gstrFile_ArchDataName

          ' ** New TrustAux.mdb location!
1940      varTmp00 = DLookup("[seclic_auxloc]", "tblSecurity_License")
1950      Select Case IsNull(varTmp00)
          Case True
1960        gstrTrustAuxLocation = gstrTrustDataLocation
1970        strAuxiliaryDatabase = gstrTrustAuxLocation & gstrFile_AuxDataName
1980      Case False
1990        Select Case varTmp00
            Case True
2000          blnAuxLoc = True
2010          gstrTrustAuxLocation = CurrentAppPath & LNK_SEP  ' ** Module Function: modFileUtilities.
2020          strAuxiliaryDatabase = gstrTrustAuxLocation & gstrFile_AuxDataName
2030        Case False
2040          gstrTrustAuxLocation = gstrTrustDataLocation
2050          strAuxiliaryDatabase = gstrTrustAuxLocation & gstrFile_AuxDataName
2060        End Select
2070      End Select

2080      If blnAuxLoc = False Then
2090        If (InStr(gstrTrustDataLocation, gstrDir_Dev) > 0) Or (InStr(gstrTrustDataLocation, gstrDir_DevClient) > 0) Then
              ' ** gstrTrustDataLocation = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\EmptyDatabase\"  '## OK
              ' ** gstrDir_Dev = "C:\VictorGCS_Clients\TrustAccountant\NewWorking"  '## OK
              ' ** gstrDir_DevClient = "C:\VictorGCS_Clients\TrustAccountant\Clients"
2100          If Len(TA_SEC) > Len(TA_SEC2) Then
2110            strAuxiliaryDatabase = gstrDir_Dev & LNK_SEP & gstrDir_DevDemo & LNK_SEP & gstrFile_AuxDataName
2120          Else
2130            strAuxiliaryDatabase = gstrDir_Dev & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & gstrFile_AuxDataName
2140          End If
2150        Else
2160          strAuxiliaryDatabase = gstrTrustDataLocation & gstrFile_AuxDataName
2170        End If
2180      End If  ' ** blnAuxLoc.

2190  On Error Resume Next
2200      blnLinkSuccess = False
2210      Set wrk = DBEngine.Workspaces(0)

          ' ******************************
2220      If blnPBar = True Then
            ' ** 3. Data file check.
2230        dblPBar_ThisStep = 3#
2240        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
2250        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
2260        Else
2270          dblPBar_SubSteps = 0#
2280          dblPBar_ThisSubStep = 0#
2290          dblPBar_ThisSubIncrement = 0#
2300          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
2310          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
2320        End If
2330        DoEvents
2340      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 3: Database Check
2350      If blnDoFrm = True Then
2360        frm.InitMsg_lbl3.Caption = "Database Check"
2370        DoEvents
2380      End If

          ' ** THIS IS THE FIRST CHECK FOR THE PRESENSE OF THE DATABASE, WITH SEARCH DIALOG IF NOT FOUND.
2390      Do While blnLinkSuccess = False
2400        Set gdbsDBLock = wrk.OpenDatabase(strDatabase, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
2410        If ERR <> 0 Or gblnChangeDB = True Then
2420          If gblnChangeDB = False Then
2430            strMsg = "Your data file was not found." & vbCrLf & vbCrLf & "Database = " & strDatabase & vbCrLf & vbCrLf & _
                  "Message: " & ERR.description & vbCrLf & vbCrLf & "Would you like to search for the data file?"
2440          Else
2450            strMsg = "Do you wish to search for a different location of your data file?"
2460          End If
2470          If MsgBox(strMsg, vbCritical + vbYesNo, "File Not Found") = vbYes Then
2480            strTitle1 = "Find Data File Location"
2490            strPath = GetFolderPathSIS(strTitle1)  ' ** Module Function: modBrowseFilesAndFolders.
2500            If strPath <> vbNullString Then
2510              gstrTrustDataLocation = strPath & LNK_SEP
2520              strDatabase = gstrTrustDataLocation & gstrFile_DataName
2530              strArchiveDatabase = gstrTrustDataLocation & gstrFile_ArchDataName
2540              If blnAuxLoc = False Then
2550                strAuxiliaryDatabase = gstrTrustDataLocation & gstrFile_AuxDataName
2560              End If
2570            End If
2580            gblnChangeDB = False
2590          Else
2600            If gblnChangeDB = False Then
2610              gdbsDBLock.Close
2620              blnRetVal = False: Exit Do
2630            Else
2640              blnRetVal = False
2650              Exit Do
2660            End If
2670          End If
2680        Else
2690          blnLinkSuccess = True
2700        End If
2710      Loop

          ' ** It's within this If/EndIf that wrk is set (the If/EndIf loop way above, not the one on the next line).
2720      If blnRetVal = False Then
2730        gdbsDBLock.Close
2740        wrk.Close
2750      End If

2760    End If  ' ** blnRetVal.

        ' ******************************
2770    If blnPBar = True Then
          ' ** 4. Icon check.
2780      dblPBar_ThisStep = 4#
2790      dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
2800      If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
            ' ** None.
2810      Else
2820        dblPBar_SubSteps = 0#
2830        dblPBar_ThisSubStep = 0#
2840        dblPBar_ThisSubIncrement = 0#
2850        dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
2860        frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
2870      End If
2880      DoEvents
2890    End If  ' ** blnPBar.
        ' ******************************

        ' ** STATUS SUB 4: Icon Check
2900    If blnDoFrm = True Then
2910      frm.InitMsg_lbl3.Caption = "Icon Check"
2920      DoEvents
2930    End If

        ' ** Check the icon path to make sure it's set to (x86) on a 64-bit machine.
2940    If blnRetVal = True Then
2950      strTmp01 = CurrentDb.Properties("AppIcon")
2960      If Parse_Path(strTmp01) <> CurrentAppPath Then  ' ** Module Functions: modFileUtilities.
2970        AppIcon_Let  ' ** Module Function: modAppVersionFuncs.
2980      End If
2990    End If

        ' ******************************
3000    If blnPBar = True Then
          ' ** 5. EULA check.
3010      dblPBar_ThisStep = 5#
3020      dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
3030      If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
            ' ** None.
3040      Else
3050        dblPBar_SubSteps = 0#
3060        dblPBar_ThisSubStep = 0#
3070        dblPBar_ThisSubIncrement = 0#
3080        dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
3090        frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
3100      End If
3110      DoEvents
3120    End If  ' ** blnPBar.
        ' ******************************

        ' ** STATUS SUB 5: EULA Check
3130    If blnDoFrm = True Then
3140      frm.InitMsg_lbl3.Caption = "Demo Check 1"
3150      DoEvents
3160    End If

        ' ** If it's a demo copy, check the EULA.
3170    If blnRetVal = True Then
3180      If Len(TA_SEC) > Len(TA_SEC2) Then
3190        strMsg = DemoLicense_Get  ' ** Module Functions: modSecurityFunctions.
3200        If strMsg <> RET_ERR And strMsg <> "#MISSING" Then
3210          If Left(strMsg, 5) <> "#EULA" Then
3220            gdatAccept = CDate(Left(strMsg, (InStr(strMsg, "~") - 1)))
3230            gstrAccept = Mid(strMsg, (InStr(strMsg, "~") + 1))
3240          Else
3250            gblnSetFocus = True
3260            DoCmd.OpenForm "frmEula", acNormal, , , , acDialog, frm.Name
3270            DoCmd.Hourglass True
3280            DoEvents
3290            If InStr(gstrAccept, "Accept") = 0 Then
3300              blnRetVal = False
3310            End If
3320          End If
3330        Else
              ' ** There was an error getting the demo license stuff!
3340          blnRetVal = False
3350          DoCmd.Hourglass False
3360          Beep
3370          MsgBox "Your End User License Agreement is invalid." & vbCrLf & vbCrLf & _
                "Please contact Delta Data, Inc.", vbCritical + vbOKOnly, "License Invalid"
3380        End If
3390      End If
3400      If blnRetVal = False Then
3410        gdbsDBLock.Close
3420        wrk.Close
3430      End If
3440    End If  ' ** blnRetVal.

3450    If blnRetVal = True Then

          ' ******************************
3460      If blnPBar = True Then
            ' ** 6. INI file 2.
3470        dblPBar_ThisStep = 6#
3480        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
3490        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
3500        Else
3510          dblPBar_SubSteps = 0#
3520          dblPBar_ThisSubStep = 0#
3530          dblPBar_ThisSubIncrement = 0#
3540          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
3550          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
3560        End If
3570        DoEvents
3580      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 6: INI File Check 2
3590      If blnDoFrm = True Then
3600        frm.InitMsg_lbl3.Caption = "INI Check 2"
3610        DoEvents
3620      End If

3630      gdbsDBLock.Close
3640      wrk.Close

          ' ** When it's satisfied a proper gstrTrustDataLocation is present, write to DDTrust.ini.
3650      IniFile_Set "Files", "Location", gstrTrustDataLocation, CurrentAppPath & LNK_SEP & gstrFile_INI  ' ** Module Function: Below, modFileUtilities.

3660  On Error GoTo ERRH

          ' ******************************
3670      If blnPBar = True Then
            ' ** 7. File extention check.
3680        dblPBar_ThisStep = 7#
3690        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
3700        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
3710        Else
3720          dblPBar_SubSteps = 0#
3730          dblPBar_ThisSubStep = 0#
3740          dblPBar_ThisSubIncrement = 0#
3750          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
3760          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
3770        End If
3780        DoEvents
3790      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 7: File Extention Check
3800      If blnDoFrm = True Then
3810        frm.InitMsg_lbl3.Caption = "File Extention Check"
3820        DoEvents
3830      End If

          ' ** This now ignores Trust Import and Trust Administration.
3840      Setup_FileExt blnAuxLoc, gstrTrustDataLocation  ' ** Function: Below.

          ' ******************************
3850      If blnPBar = True Then
            ' ** 8. Link list.
3860        dblPBar_ThisStep = 8#
3870        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
3880        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
3890        Else
3900          dblPBar_SubSteps = 0#
3910          dblPBar_ThisSubStep = 0#
3920          dblPBar_ThisSubIncrement = 0#
3930          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
3940          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
3950        End If
3960        DoEvents
3970      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 8: Link List Check
3980      If blnDoFrm = True Then
3990        frm.InitMsg_lbl3.Caption = "Link List Check"
4000        DoEvents
4010      End If

          ' ** Get the current list of tables to be linked from a local table.
4020      arr_varLink = GetLinkList  ' ** Function: Below.
4030      lngLinks = UBound(arr_varLink, 2) + 1&
4040      If arr_varLink(L_NAM, 0) = RET_ERR Then
4050        blnRetVal = False
4060      End If
          ' ********************************************************
          ' ** Array: arr_varLink()
          ' **
          ' **   Field  Element  Name                   Constant
          ' **   =====  =======  =====================  ==========
          ' **     1       0     mtbl_ID                L_ID
          ' **     2       1     mtbl_NAME              L_NAM
          ' **     3       2     mtbl_AUTONUMBER        L_AUT
          ' **     4       3     mtbl_ORDER             L_ORD
          ' **     5       4     mtbl_NEWRecs           L_NEW
          ' **     6       5     mtbl_ACTIVE            L_ACT
          ' **     7       6     mtbl_DTA               L_DTA
          ' **     8       7     mtbl_ARCH              L_ARC
          ' **     9       8     mtbl_AUX               L_AUX
          ' **    10       9     mtbl_SOURCE            L_SRC
          ' **    11      10     contype_type           L_TYP
          ' **    12      11     tbllnk_connect         L_LNKT
          ' **    13      12     tbllnk_connect_CURR    L_LNKC
          ' **    14      13     tbllnk_fnd             L_FND
          ' **    15      14     tbllnk_fix             L_FIX
          ' **
          ' ********************************************************

4070    End If  ' ** blnRetVal.

        ' ** Tables we no longer use:
        'TableDelete ("assetsub")  ' ** Delete it if it's there, but don't put it back.
        'TableDelete ("feefreq")  ' ** Delete it if it's there, but don't put it back.
        'TableDelete ("jurisdiction")  ' ** Delete it if it's there, but don't put it back.
        'TableDelete ("Lock")  ' ** Delete it if it's there, but don't put it back.
        'TableDelete ("MasterAsset Temp")  ' ** Delete it if it's there, but don't put it back.
        'TableDelete ("reviewfreq")  ' ** Delete it if it's there, but don't put it back.
        'TableDelete ("statementfreq")  ' ** Delete it if it's there, but don't put it back.
        'TableDelete ("tblaveragePrice")  ' ** Delete it if it's there, but don't put it back.

4080    If blnRetVal = True Then

          ' ******************************
4090      If blnPBar = True Then
            ' ** 9. Version check 1.
4100        dblPBar_ThisStep = 9#
4110        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
4120        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
4130        Else
4140          dblPBar_SubSteps = 0#
4150          dblPBar_ThisSubStep = 0#
4160          dblPBar_ThisSubIncrement = 0#
4170          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
4180          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
4190        End If
4200        DoEvents
4210      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 9: Version Check 1
4220      If blnDoFrm = True Then
4230        frm.InitMsg_lbl3.Caption = "Version Check 1"
4240        DoEvents
4250      End If

          ' ** Check Trust Accountant version information.
4260      strVerNum = vbNullString
4270      Set wrk = DBEngine.Workspaces(0)
4280      Set gdbsDBLock = wrk.OpenDatabase(strDatabase, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
4290      With gdbsDBLock

            ' ** Look first for m_VD, in order to check backend version.
4300        blnFound = False
4310        For Each tdf In .TableDefs
4320          With tdf
4330            If .Name = "m_VD" Then
4340              blnFound = True
4350              Exit For
4360            End If
4370          End With
4380        Next

4390        If blnFound = True Then
              ' ** Get its version number.
4400  On Error Resume Next
4410          Set rst = .OpenRecordset("m_VD", dbOpenDynaset, dbReadOnly)
4420          blnNoVersion = rst.BOF
4430          If ERR.Number <> 0 Then
4440  On Error GoTo ERRH
4450            blnNoVersion = True
4460            blnRetVal = False
4470            MsgBox "An important table was not found in the linked data file." & vbCrLf & vbCrLf & _
                  "Either you need to run a Conversion to bring it up to date," & vbCrLf & _
                  "or the file, TrustDta.mdb, is corrupted." & vbCrLf & vbCrLf & _
                  "Contact Delta Data, Inc., for assistance.", _
                  vbCritical + vbOKOnly, ("Trust Accountant Version Mismatch" & Space(20) & "1")
4480          Else
4490  On Error GoTo ERRH
4500            If rst.BOF = True And rst.EOF = True Then
4510              blnNoVersion = True
4520              blnRetVal = False
4530              MsgBox "An important table was not found in the linked data file." & vbCrLf & vbCrLf & _
                    "Either you need to run a Conversion to bring it up to date," & vbCrLf & _
                    "or the file, TrustDta.mdb, is corrupted." & vbCrLf & vbCrLf & _
                    "Contact Delta Data, Inc., for assistance.", _
                    vbCritical + vbOKOnly, ("Trust Accountant Version Mismatch" & Space(20) & "2")
4540            Else
4550              rst.MoveFirst
4560              strVerNum = GetVersionInfo(rst![vd_MAIN], rst![vd_MINOR], rst![vd_REVISION])  ' ** Function: Below.
4570            End If
4580            rst.Close
4590          End If
4600          Set rst = Nothing
4610        Else
              ' ** Real early version, or bad MDB.
4620          blnRetVal = False
4630          MsgBox "An important table was not found in the linked data file." & vbCrLf & vbCrLf & _
                "Either you need to run a Conversion to bring it up to date," & vbCrLf & _
                "or the file, TrustDta.mdb, is corrupted." & vbCrLf & vbCrLf & _
                "Contact Delta Data, Inc., for assistance.", _
                vbCritical + vbOKOnly, ("Trust Accountant Version Mismatch" & Space(20) & "3")
4640        End If

4650        .Close
4660      End With
4670      wrk.Close

4680    End If  ' ** blnRetVal.

4690    If blnRetVal = True Then

          ' ******************************
4700      If blnPBar = True Then
            ' ** 10. Version check 2.
4710        dblPBar_ThisStep = 10#
4720        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
4730        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
4740        Else
4750          dblPBar_SubSteps = 0#
4760          dblPBar_ThisSubStep = 0#
4770          dblPBar_ThisSubIncrement = 0#
4780          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
4790          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
4800        End If
4810        DoEvents
4820      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 10: Version Check 2
4830      If blnDoFrm = True Then
4840        frm.InitMsg_lbl3.Caption = "Version Check 2"
4850        DoEvents
4860      End If

          ' ** Continue processing the Trust Accountant version information.
4870      blnRetVal = False  ' ** It'll be easier going the other way here.
4880      If strVerNum <> vbNullString Then
            ' ** This new version, from v2.1.64 forward, should always have, and only have, a corresponding TrustDta.mdb and Archive.
            ' ** Since it's now so simple to convert previous versions, let them know that's what they must do.
            ' ** Remember, the conversion always starts with an empty v2.1.64, .65, .66, and then transfers everything.
            ' ** NOTE: strVerNum comes from m_VP, which has numeric fields, so a Zero in vp_REVISION is read a ".0", not ".00"
4890        intPos01 = InStr(strVerNum, ".")
4900        If intPos01 > 0 Then
4910          If Val(Left(strVerNum, (intPos01 - 1))) >= 2 Then
4920            intPos02 = InStr((intPos01 + 1), strVerNum, ".")
4930            If intPos02 > 0 Then
4940              If Val(Mid(strVerNum, (intPos01 + 1), 1)) = 1 Then
4950                If Val(Mid(strVerNum, (intPos02 + 1))) >= 64 Then
4960                  blnRetVal = True
4970                End If
4980              ElseIf Val(Mid(strVerNum, (intPos01 + 1), 1)) >= 2 Then
4990                blnRetVal = True
5000              End If
5010            End If
5020          End If
5030        End If
5040      End If
5050      If blnRetVal = False Then
5060        MsgBox "Your data is from a previous version." & vbCrLf & vbCrLf & _
              "The version expected is v" & AppVersion_Get2 & ", and your data file appears to be v" & strVerNum & vbCrLf & _
              "Contact Delta Data, Inc., for assistance.", _
              vbCritical + vbOKOnly, "Conversion Required"  ' ** Module Function: AppVersionFuncs.
5070      End If

5080    End If  ' ** blnRetVal.

        ' ********************************************************
        ' ** Array: arr_varLink()
        ' **
        ' **   Field  Element  Name                   Constant
        ' **   =====  =======  =====================  ==========
        ' **     1       0     mtbl_ID                L_ID
        ' **     2       1     mtbl_NAME              L_NAM
        ' **     3       2     mtbl_AUTONUMBER        L_AUT
        ' **     4       3     mtbl_ORDER             L_ORD
        ' **     5       4     mtbl_NEWRecs           L_NEW
        ' **     6       5     mtbl_ACTIVE            L_ACT
        ' **     7       6     mtbl_DTA               L_DTA
        ' **     8       7     mtbl_ARCH              L_ARC
        ' **     9       8     mtbl_AUX               L_AUX
        ' **    10       9     mtbl_SOURCE            L_SRC
        ' **    11      10     contype_type           L_TYP
        ' **    12      11     tbllnk_connect         L_LNKT
        ' **    13      12     tbllnk_connect_CURR    L_LNKC
        ' **    14      13     tbllnk_fnd             L_FND
        ' **    15      14     tbllnk_fix             L_FIX
        ' **
        ' ********************************************************

        ' ** Check all current table links.
5090    If blnRetVal = True Then

5100      Set dbs = CurrentDb
5110      With dbs

            ' ******************************
5120        If blnPBar = True Then
              ' ** 11. Current links.
5130          dblPBar_ThisStep = 11#
5140          dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
5150          If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
                ' ** YES!
5160            dblPBar_SubSteps = .TableDefs.Count
5170            dblPBar_ThisSubStep = 0#
5180            dblPBar_ThisSubIncrement = (dblPBar_ThisIncrement / dblPBar_SubSteps)
5190          Else
5200            dblPBar_SubSteps = 0#
5210            dblPBar_ThisSubStep = 0#
5220            dblPBar_ThisSubIncrement = 0#
5230            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
5240            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
5250          End If
5260          DoEvents
5270        End If  ' ** blnPBar.
            ' ******************************

            ' ** STATUS SUB 11: Link Connection Check
5280        If blnDoFrm = True Then
5290          frm.InitMsg_lbl3.Caption = "Connection Check"
5300          DoEvents
5310        End If

            ' ** Check current Connect string against what it's supposed to be.
5320        For Each tdf In .TableDefs
5330          With tdf
5340            If .Connect <> vbNullString Then
5350              For lngX = 0& To (lngLinks - 1&)
5360                If arr_varLink(L_NAM, lngX) = .Name Then
5370                  arr_varLink(L_FND, lngX) = CBool(True)
5380                  arr_varLink(L_LNKC, lngX) = .Connect
5390                  intPos01 = InStr(.Connect, LNK_IDENT)
5400                  If intPos01 > 0 Then
5410                    strTmp01 = Mid(arr_varLink(L_LNKC, lngX), (intPos01 + Len(LNK_IDENT)))
5420                    If arr_varLink(L_DTA, lngX) = True Then
5430                      If strTmp01 <> strDatabase Then
                            ' ** Current link not right.
5440                        arr_varLink(L_FIX, lngX) = CBool(True)
5450                      End If
5460                    ElseIf arr_varLink(L_ARC, lngX) = True Then
5470                      If strTmp01 <> strArchiveDatabase Then
                            ' ** Current link not right.
5480                        arr_varLink(L_FIX, lngX) = CBool(True)
5490                      End If
5500                    ElseIf arr_varLink(L_AUX, lngX) = True Then
5510                      If strTmp01 <> strAuxiliaryDatabase Then
                            ' ** Current link not right.
5520                        arr_varLink(L_FIX, lngX) = CBool(True)
5530                      End If
5540                    End If
5550                  Else
                        ' ** Current link not right.
5560                    arr_varLink(L_FIX, lngX) = CBool(True)
5570                  End If
5580                End If
5590              Next  ' ** lngLinks: lngX.
5600            End If
5610          End With
              ' ********************
5620          If blnPBar = True Then
5630            dblPBar_ThisSubStep = dblPBar_ThisSubStep + 1&
5640            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisSubIncrement)  'COULD DO THISWIDTH = X * INCRE, THEN CURWIDTH + THISWIDTH
5650            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
5660            DoEvents
5670          End If
              ' ********************
5680        Next  ' ** TableDefs: tdf.

            ' ******************************
5690        If blnPBar = True Then
              ' ** 12. Extra links.
5700          dblPBar_ThisStep = 12#
5710          dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
5720          If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
                ' ** YES!
5730            dblPBar_SubSteps = .TableDefs.Count
5740            dblPBar_ThisSubStep = 0#
5750            dblPBar_ThisSubIncrement = (dblPBar_ThisIncrement / dblPBar_SubSteps)
5760          Else
5770            dblPBar_SubSteps = 0#
5780            dblPBar_ThisSubStep = 0#
5790            dblPBar_ThisSubIncrement = 0#
5800            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
5810            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
5820          End If
5830          DoEvents
5840        End If  ' ** blnPBar.
            ' ******************************

            ' ** STATUS SUB 12: Extra Link Check
5850        If blnDoFrm = True Then
5860          frm.InitMsg_lbl3.Caption = "Extra Link Check"
5870          DoEvents
5880        End If

            ' ** Check for extraneous links.
5890        lngDels = 0&
5900        ReDim arr_varDel(D_ELEMS, 0)
5910        For Each tdf In .TableDefs
5920          With tdf
5930            If .Connect <> vbNullString Then
5940              blnFound = False
5950              For lngX = 0& To (lngLinks - 1&)
5960                If arr_varLink(L_NAM, lngX) = .Name Then
5970                  blnFound = True
5980                  Exit For
5990                End If
6000              Next
6010              If blnFound = False Then
6020                lngDels = lngDels + 1&
6030                lngE = lngDels - 1&
6040                ReDim Preserve arr_varDel(D_ELEMS, lngE)
6050                arr_varDel(D_NAM, lngE) = .Name
6060                arr_varDel(D_FND, lngE) = CBool(False)
6070              End If
6080            End If
6090          End With
              ' ********************
6100          If blnPBar = True Then
6110            dblPBar_ThisSubStep = dblPBar_ThisSubStep + 1&
6120            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisSubIncrement)
6130            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
6140            DoEvents
6150          End If
              ' ********************
6160        Next

6170        .Close
6180      End With

          ' ** Delete extraneous links.
6190      If lngDels > 0& Then
6200        For lngX = 0& To (lngDels - 1&)
6210          TableDelete CStr(arr_varDel(D_NAM, lngX))  ' ** Module Function: modFileUtilities.
6220        Next
6230      End If

6240    End If  ' ** blnRetVal.

        ' ** Add or change any missing or broken table links.
6250    If blnRetVal = True Then

6260      Set dbs = CurrentDb
6270      With dbs

            ' ******************************
6280        If blnPBar = True Then
              ' ** 13. Link TrustDta.mdb.
6290          dblPBar_ThisStep = 13#
6300          dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
6310          If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
                ' ** YES!
6320            dblPBar_SubSteps = lngLinks
6330            dblPBar_ThisSubStep = 0#
6340            dblPBar_ThisSubIncrement = (dblPBar_ThisIncrement / dblPBar_SubSteps)
6350          Else
6360            dblPBar_SubSteps = 0#
6370            dblPBar_ThisSubStep = 0#
6380            dblPBar_ThisSubIncrement = 0#
6390            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
6400            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
6410          End If
6420          DoEvents
6430        End If  ' ** blnPBar.
            ' ******************************

            ' ** STATUS SUB 13: TrustDta Link
6440        If blnDoFrm = True Then
6450          frm.InitMsg_lbl3.Caption = "Link TrustDta "
6460          lngLinkCnt = 0&
6470          DoEvents
6480        End If

            ' ** First, TrustDta.
6490        For lngX = 0& To (lngLinks - 1&)
6500          If arr_varLink(L_DTA, lngX) = True Then
6510            lngLinkCnt = lngLinkCnt + 1&
6520            If blnDoFrm = True Then
6530              frm.InitMsg_lbl3.Caption = "Link TrustDta: " & CStr(lngLinkCnt) & " of " & CStr(lngLinks)
6540              DoEvents
6550            End If
6560            If arr_varLink(L_FND, lngX) = True Then
6570              If arr_varLink(L_FIX, lngX) = True Then
6580                If arr_varLink(L_NAM, lngX) = "tblDataTypeDb1" Or arr_varLink(L_NAM, lngX) = "tblDataTypeDb" Then
6590                  TableDelete "tblDataTypeDb1"  ' ** Module Function: modFileUtilities.
6600                  DoCmd.TransferDatabase acLink, "Microsoft Access", strDatabase, acTable, "tblDataTypeDb", "tblDataTypeDb1"
6610                Else
6620                  TableDelete CStr(arr_varLink(L_NAM, lngX))  ' ** Module Function: modFileUtilities.
6630                  DoCmd.TransferDatabase acLink, "Microsoft Access", strDatabase, acTable, arr_varLink(L_NAM, lngX), arr_varLink(L_NAM, lngX)
6640                End If
6650                DoEvents
6660                .TableDefs.Refresh
6670                .TableDefs.Refresh
6680                strTmp01 = .TableDefs(arr_varLink(L_NAM, lngX)).Connect
6690                arr_varLink(L_LNKC, lngX) = strTmp01
6700              End If
6710            Else
6720              If arr_varLink(L_NAM, lngX) = "tblDataTypeDb1" Or arr_varLink(L_NAM, lngX) = "tblDataTypeDb" Then
6730                TableDelete "tblDataTypeDb1"  ' ** Module Function: modFileUtilities.
6740                DoCmd.TransferDatabase acLink, "Microsoft Access", strDatabase, acTable, "tblDataTypeDb", "tblDataTypeDb1"
6750              Else
6760                TableDelete CStr(arr_varLink(L_NAM, lngX))  ' ** Module Function: modFileUtilities.
6770                If arr_varLink(L_NAM, lngX) = "_~xusr" Then
6780  On Error Resume Next
6790                  DoCmd.TransferDatabase acLink, "Microsoft Access", strDatabase, acTable, arr_varLink(L_NAM, lngX), arr_varLink(L_NAM, lngX)
6800                  If ERR.Number <> 0 Then
6810                    blnRetVal = False
6820                    DoCmd.Hourglass False
6830                    Select Case ERR.Number
                        Case 3011  ' ** The Microsoft Jet database engine could not find the object '|'.
6840                      Beep
6850                      MsgBox "If this is a new Upgrade, files necessary for proper conversion were not copied correctly." & vbCrLf & vbCrLf & _
                            "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "Installation Error"
6860                    Case Else
6870                      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6880                    End Select
6890  On Error GoTo ERRH
6900                    Exit For
6910                  Else
6920  On Error GoTo ERRH
6930                  End If
6940                Else
6950                  DoCmd.TransferDatabase acLink, "Microsoft Access", strDatabase, acTable, arr_varLink(L_NAM, lngX), arr_varLink(L_NAM, lngX)
6960                End If
6970              End If
6980              DoEvents
6990              .TableDefs.Refresh
7000              .TableDefs.Refresh
7010              strTmp01 = .TableDefs(arr_varLink(L_NAM, lngX)).Connect
7020              arr_varLink(L_LNKC, lngX) = strTmp01
7030            End If
7040          End If
              ' ********************
7050          If blnPBar = True Then
7060            dblPBar_ThisSubStep = dblPBar_ThisSubStep + 1&
7070            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisSubIncrement)
7080            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
7090            DoEvents
7100          End If
              ' ********************
7110        Next  ' ** lngLinks: lngX.

            ' ******************************
7120        If blnPBar = True Then
              ' ** 14. Link TrstArch.mdb.
7130          dblPBar_ThisStep = 14#
7140          dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
7150          If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
                ' ** YES!
7160            dblPBar_SubSteps = lngLinks
7170            dblPBar_ThisSubStep = 0#
7180            dblPBar_ThisSubIncrement = (dblPBar_ThisIncrement / dblPBar_SubSteps)
7190          Else
7200            dblPBar_SubSteps = 0#
7210            dblPBar_ThisSubStep = 0#
7220            dblPBar_ThisSubIncrement = 0#
7230            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
7240            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
7250          End If
7260          DoEvents
7270        End If  ' ** blnPBar.
            ' ******************************

            ' ** STATUS SUB 14: TrstArch Link
7280        If blnDoFrm = True Then
7290          frm.InitMsg_lbl3.Caption = "Link TrstArch "
7300          DoEvents
7310        End If

            ' ** Then, TrstArch.
7320        For lngX = 0& To (lngLinks - 1&)
7330          If arr_varLink(L_ARC, lngX) = True Then
7340            lngLinkCnt = lngLinkCnt + 1&
7350            If blnDoFrm = True Then
7360              frm.InitMsg_lbl3.Caption = "Link TrstArch: " & CStr(lngLinkCnt) & " of " & CStr(lngLinks)
7370              DoEvents
7380            End If
7390            If IsNull(arr_varLink(L_SRC, lngX)) = False Then
7400              If arr_varLink(L_FND, lngX) = True Then
7410                If arr_varLink(L_FIX, lngX) = True Then
7420                  TableDelete CStr(arr_varLink(L_NAM, lngX))  ' ** Module Function: modFileUtilities.
7430                  DoCmd.TransferDatabase acLink, "Microsoft Access", strArchiveDatabase, acTable, arr_varLink(L_SRC, lngX), arr_varLink(L_NAM, lngX)
7440                  DoEvents
7450                  .TableDefs.Refresh
7460                  .TableDefs.Refresh
7470                  strTmp01 = .TableDefs(arr_varLink(L_NAM, lngX)).Connect
7480                  arr_varLink(L_LNKC, lngX) = strTmp01
7490                End If
7500              Else
7510                TableDelete CStr(arr_varLink(L_NAM, lngX))  ' ** Module Function: modFileUtilities.
7520                DoCmd.TransferDatabase acLink, "Microsoft Access", strArchiveDatabase, acTable, arr_varLink(L_SRC, lngX), arr_varLink(L_NAM, lngX)
7530                DoEvents
7540                .TableDefs.Refresh
7550                .TableDefs.Refresh
7560                strTmp01 = .TableDefs(arr_varLink(L_NAM, lngX)).Connect
7570                arr_varLink(L_LNKC, lngX) = strTmp01
7580              End If
7590            Else
7600              If arr_varLink(L_FND, lngX) = True Then
7610                If arr_varLink(L_FIX, lngX) = True Then
7620                  TableDelete CStr(arr_varLink(L_NAM, lngX))  ' ** Module Function: modFileUtilities.
7630                  DoCmd.TransferDatabase acLink, "Microsoft Access", strArchiveDatabase, acTable, arr_varLink(L_NAM, lngX), arr_varLink(L_NAM, lngX)
7640                  DoEvents
7650                  .TableDefs.Refresh
7660                  .TableDefs.Refresh
7670                  strTmp01 = .TableDefs(arr_varLink(L_NAM, lngX)).Connect
7680                  arr_varLink(L_LNKC, lngX) = strTmp01
7690                End If
7700              Else
7710                TableDelete CStr(arr_varLink(L_NAM, lngX))  ' ** Module Function: modFileUtilities.
7720                DoCmd.TransferDatabase acLink, "Microsoft Access", strArchiveDatabase, acTable, arr_varLink(L_NAM, lngX), arr_varLink(L_NAM, lngX)
7730                DoEvents
7740                .TableDefs.Refresh
7750                .TableDefs.Refresh
7760                strTmp01 = .TableDefs(arr_varLink(L_NAM, lngX)).Connect
7770                arr_varLink(L_LNKC, lngX) = strTmp01
7780              End If
7790            End If
7800          End If
              ' ********************
7810          If blnPBar = True Then
7820            dblPBar_ThisSubStep = dblPBar_ThisSubStep + 1&
7830            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisSubIncrement)
7840            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
7850            DoEvents
7860          End If

              ' ********************
7870        Next  ' ** lngLinks: lngX.

            ' ******************************
7880        If blnPBar = True Then
              ' ** 15. Link TrustAux.mdb.
7890          dblPBar_ThisStep = 15#
7900          dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
7910          If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
                ' ** YES!
7920            dblPBar_SubSteps = lngLinks
7930            dblPBar_ThisSubStep = 0#
7940            dblPBar_ThisSubIncrement = (dblPBar_ThisIncrement / dblPBar_SubSteps)
7950          Else
7960            dblPBar_SubSteps = 0#
7970            dblPBar_ThisSubStep = 0#
7980            dblPBar_ThisSubIncrement = 0#
7990            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
8000            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
8010          End If
8020          DoEvents
8030        End If  ' ** blnPBar.
            ' ******************************

            ' ** STATUS SUB 15: TrustAux Link
8040        If blnDoFrm = True Then
8050          frm.InitMsg_lbl3.Caption = "Link TrustAux "
8060          DoEvents
8070        End If

            ' ** And finally, TrustAux.
8080        For lngX = 0& To (lngLinks - 1&)
8090          If arr_varLink(L_AUX, lngX) = True Then
8100            lngLinkCnt = lngLinkCnt + 1&
8110            If blnDoFrm = True Then
8120              frm.InitMsg_lbl3.Caption = "Link TrustAux: " & CStr(lngLinkCnt) & " of " & CStr(lngLinks)
8130              DoEvents
8140            End If
8150            If arr_varLink(L_FND, lngX) = True Then
8160              If arr_varLink(L_FIX, lngX) = True Then
                    ' ** One last check to see if relinking is necessary.
8170                If Mid(.TableDefs(arr_varLink(L_NAM, lngX)).Connect, _
                        (InStr(.TableDefs(arr_varLink(L_NAM, lngX)).Connect, LNK_IDENT) + Len(LNK_IDENT))) <> strAuxiliaryDatabase Then
8180                  TableDelete CStr(arr_varLink(L_NAM, lngX))  ' ** Module Function: modFileUtilities.
8190                  DoCmd.TransferDatabase acLink, "Microsoft Access", strAuxiliaryDatabase, acTable, arr_varLink(L_NAM, lngX), arr_varLink(L_NAM, lngX)
8200                  DoEvents
8210                  .TableDefs.Refresh
8220                  .TableDefs.Refresh
8230                  strTmp01 = .TableDefs(arr_varLink(L_NAM, lngX)).Connect
8240                  arr_varLink(L_LNKC, lngX) = strTmp01
8250                End If
8260              End If
8270            Else
8280              TableDelete CStr(arr_varLink(L_NAM, lngX))  ' ** Module Function: modFileUtilities.
8290              DoCmd.TransferDatabase acLink, "Microsoft Access", strAuxiliaryDatabase, acTable, arr_varLink(L_NAM, lngX), arr_varLink(L_NAM, lngX)
8300              DoEvents
8310              .TableDefs.Refresh
8320              .TableDefs.Refresh
8330              strTmp01 = .TableDefs(arr_varLink(L_NAM, lngX)).Connect
8340              arr_varLink(L_LNKC, lngX) = strTmp01
8350            End If
8360          End If
              ' ********************
8370          If blnPBar = True Then
8380            dblPBar_ThisSubStep = dblPBar_ThisSubStep + 1&
8390            dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisSubIncrement)
8400            frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
8410            DoEvents
8420          End If
              ' ********************
8430        Next  ' ** lngLinks: lngX.

8440        .Close
8450      End With

          ' ******************************
8460      If blnPBar = True Then
            ' ** 16. Hidden check.
8470        dblPBar_ThisStep = 16#
8480        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
8490        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
8500        Else
8510          dblPBar_SubSteps = 0#
8520          dblPBar_ThisSubStep = 0#
8530          dblPBar_ThisSubIncrement = 0#
8540          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
8550          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
8560        End If
8570        DoEvents
8580      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 16: Hidden Object Check
8590      If blnDoFrm = True Then
8600        frm.InitMsg_lbl3.Caption = "Hidden Object Check"
8610        DoEvents
8620      End If

          ' ** Make sure certain things are hidden (this is an additional step in the new startup routine).
8630      If Security_Hidden_Load = True Then  ' ** Module Function: modSecurityFunctions.
8640        For lngX = 0& To (glngHids - 1&)
8650  On Error Resume Next
              ' ** If some of my zz_qry_System_nn queries aren't there, don't blow the whole thing!
8660          Application.SetHiddenAttribute garr_varHid(H_TYP, lngX), garr_varHid(H_NAM, lngX), True
8670  On Error GoTo ERRH
8680        Next
8690      End If

8700    End If  ' ** blnRetVal.

8710    If blnRetVal = True Then

          ' ******************************
8720      If blnPBar = True Then
            ' ** 17. License check 1.
8730        dblPBar_ThisStep = 17#
8740        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
8750        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
8760        Else
8770          dblPBar_SubSteps = 0#
8780          dblPBar_ThisSubStep = 0#
8790          dblPBar_ThisSubIncrement = 0#
8800          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
8810          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
8820        End If
8830        DoEvents
8840      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 17: License Check 1
8850      If blnDoFrm = True Then
8860        frm.InitMsg_lbl3.Caption = "License Check 1"
8870        DoEvents
8880      End If

          ' ** Do a preliminary check of the  current License status in TA.lic.
8890      strIs_License_String = DecodeString(IniFile_Get("License", "Firm", _
            EncodeString("Call Delta Data, Inc., for Licensing info."), gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Procedure: modStartupFuncs, modCodeUtilities.

8900      Set dbs = CurrentDb
8910      Set rst = dbs.OpenRecordset("License Name")
8920      rst.MoveFirst
8930      rst.Edit
8940      rst![License name] = strIs_License_String
8950      rst.Update
8960      rst.Close
8970      Set rst = Nothing
          ' ** Now, we must have version information, so see if it matches.
8980      Set qdf1 = dbs.QueryDefs("qrySystemStartup_06_m_VP")
8990      Set rstProg = qdf1.OpenRecordset  '("SELECT m_VP.* FROM m_VP", dbOpenSnapshot)
9000      Set qdf2 = dbs.QueryDefs("qrySystemStartup_07_m_VD")
9010      Set rstData = qdf2.OpenRecordset  '("SELECT m_VD.* FROM m_VD", dbOpenSnapshot)
9020      If rstProg.RecordCount < 1 Then
9030        MsgBox "Your program file is missing version information. Unable to continue.", _
              vbCritical + vbOKOnly, "Version Info Missing"
9040        blnRetVal = False
9050        rstProg.Close
9060        rstData.Close
9070        dbs.Close
9080      Else
9090        If rstData.RecordCount < 1 Then  ' ** No records.
9100          MsgBox "Your data is from a previous version." & vbCrLf & vbCrLf & _
                "You must first run the conversion utility to continue.", vbCritical + vbOKOnly, "Conversion Required"
9110          blnRetVal = False
9120          rstProg.Close
9130          rstData.Close
9140          dbs.Close
9150        Else
              ' ** If we got here, we must have both version tables, and at least one record in each.
              ' ** The current version must be the first record in the table.
9160          rstProg.MoveFirst
9170          rstData.MoveFirst
9180          dblFullVerP = Val(Trim(str(rstProg("vp_MAIN"))) & "." & Trim(str(rstProg("vp_MINOR"))))
9190          dblFullVerD = Val(Trim(str(rstData("vd_MAIN"))) & "." & Trim(str(rstData("vd_MINOR"))))
              ' ** The MAIN and MINOR version numbers must match, though the REVISION numbers can be all over the map.
9200          If rstProg("vp_MAIN") <> rstData("vd_MAIN") Or rstProg("vp_MINOR") <> rstData("vd_MINOR") Then
9210            strMsg = "Your data is a different version than this program expects." & vbCrLf & vbCrLf
9220            strMsg = strMsg & "Data: " & str(dblFullVerD) & IIf(dblFullVerD = Round(dblFullVerD, 0), ".0", "") & vbCrLf & "Program Expected: " & str(dblFullVerP) & IIf(dblFullVerP = Round(dblFullVerP, 0), ".0", "") & vbCrLf & vbCrLf
9230            If dblFullVerD < dblFullVerP Then
9240              strMsg = strMsg & "You must first run the conversion utility to continue."
9250            Else
9260              strMsg = strMsg & "You need an updated program to work with this data."
9270            End If
9280            MsgBox strMsg, vbCritical + vbOKOnly, "Version Mismatch"
9290            blnRetVal = False
9300            rstProg.Close
9310            rstData.Close
9320            dbs.Close
9330          Else
9340            rstProg.Close
9350            rstData.Close
9360            dbs.Close
9370          End If
9380        End If
9390      End If
9400      Set rstProg = Nothing
9410      Set rstData = Nothing
9420      Set qdf1 = Nothing
9430      Set qdf2 = Nothing
9440      Set dbs = Nothing
9450    End If  ' ** blnRetVal.

        ' ** Debug and Demo checks.
9460    If blnRetVal = True Then

          ' ******************************
9470      If blnPBar = True Then
            ' ** 18. Demo check 2.
9480        dblPBar_ThisStep = 18#
9490        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
9500        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
9510        Else
9520          dblPBar_SubSteps = 0#
9530          dblPBar_ThisSubStep = 0#
9540          dblPBar_ThisSubIncrement = 0#
9550          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
9560          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
9570        End If
9580        DoEvents
9590      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 18: Demo Check 2
9600      If blnDoFrm = True Then
9610        frm.InitMsg_lbl3.Caption = "Demo Check 2"
9620        DoEvents
9630      End If

9640      If gblnDev_Debug Then
9650        MsgBox "Global Debug Variable ACTIVE!", vbExclamation + vbOKOnly, "Warning: Debug Active"
9660      End If
9670      If gblnDemo Then
            ' ** Running in DEMO mode.
9680      Else
            ' ** Running in NORMAL mode.
            ' ** Currently, the Demo version uses the Username of 'TADemo', so it won't get caught here.
9690        If Trim(UCase(CurrentUser)) = "DEMO" Then  ' ** Internal Access Function: Trust Accountant login.
9700          MsgBox "The 'Demo' user is only for demonstration purposes." & vbCrLf & vbCrLf & _
                "Please run the program as a valid user.", vbCritical + vbOKOnly, "Demo Mode"
9710          blnRetVal = False
9720        End If
9730      End If

9740    End If  ' ** blnRetVal.

9750    If blnRetVal = True Then

          ' ******************************
9760      If blnPBar = True Then
            ' ** 19. Conversion check.
9770        dblPBar_ThisStep = 19#
9780        dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
9790        If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
9800        Else
9810          dblPBar_SubSteps = 0#
9820          dblPBar_ThisSubStep = 0#
9830          dblPBar_ThisSubIncrement = 0#
9840          dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
9850          frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
9860        End If
9870        DoEvents
9880      End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 19: Conversion Preference Check
9890      If blnDoFrm = True Then
9900        frm.InitMsg_lbl3.Caption = "Conversion Preference Check"
9910        DoEvents
9920      End If

          ' ** Check for ConversionCheck() preference record.
          ' ** Once this preference has been set True, the startup
          ' ** routine will no longer check the \Convert_New directory.
9930      blnFound = False
9940      Set dbs = CurrentDb
9950      With dbs

            ' ** tblPreference_User, for 'chkConversionCheck', by specified [usr].
            'This is the first spot it detects the wrong data files, i.e., it put
            'the old data in \Database instead of the empties ready for conversion.
            'WAIT A MINUTE! IT CHECKS FOR VERSION ABOVE, SO HOW CAN IT GET PAST THAT?
9960        Set qdf1 = .QueryDefs("qryPreferences_06_01")  '##dbs_id
9970        With qdf1.Parameters
9980          ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
9990        End With
10000       Set rst = qdf1.OpenRecordset
10010       With rst
10020         If .BOF = True And .EOF = True Then
                ' ** No preference for ConversionCheck().
10030         Else
10040           blnFound = True
10050           .MoveFirst
10060           If IsLoaded("frmMenu_Title", acForm) = True Then  ' ** Module Functions: modFileUtilities.
10070             Forms("frmMenu_Title").chkConversionCheck = ![prefuser_boolean]
10080           Else
                  ' ** May not be open yet, or may be from frmLinkData.
10090           End If
10100         End If
10110         .Close
10120       End With
10130       Set rst = Nothing
10140       Set qdf1 = Nothing

10150       If blnFound = False Then
              ' ** Append default 'chkConversionCheck' record (False) to tblPreference_User, by specified [usr].
10160         Set qdf1 = .QueryDefs("qryPreferences_06_02")  '##dbs_id
10170         With qdf1.Parameters
10180           ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10190         End With
10200         qdf1.Execute
10210         Set qdf1 = Nothing
10220       End If

10230       blnFound = False
            ' ** tblPreference_User, for 'chkIncludeCurrency' on 'frmPostingDate', just dbs_id = 1, by specified [usr].
10240       Set qdf1 = .QueryDefs("qryPreferences_07_01")
10250       With qdf1.Parameters
10260         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10270       End With
10280       Set rst = qdf1.OpenRecordset
10290       With rst
10300         If .BOF = True And .EOF = True Then
                ' ** No preference for chkIncludeCurrency.
10310         Else
10320           .MoveFirst
10330           blnFound = True
10340         End If
10350         .Close
10360       End With
10370       Set qdf1 = Nothing
10380       Set rst = Nothing

10390       If blnFound = False Then
              ' ** Append 'chkIncludeCurrency' record to tblPreference_User, just dbs_id = 1, by specified [usr], [pbln].
10400         Set qdf1 = .QueryDefs("qryPreferences_07_02")
10410         With qdf1.Parameters
10420           ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10430           ![pbln] = False
10440         End With
10450         qdf1.Execute
10460         Set qdf1 = Nothing
10470       End If

10480       .Close
10490     End With  ' ** dbs.
10500     Set dbs = Nothing

10510   End If  ' ** blnRetVal.

10520   If blnRetVal = True Then

10530     If blnPBar = False Then
            ' ** Reset the frmMenu_Title caption.
10540       With Forms("frmMenu_Title")
10550         .cmdMenu.Visible = True
10560         .cmdMenu_box1.Visible = True
10570         .cmdMenu_box2.Visible = True
10580         .cmdQuit.Visible = True
10590         .cmdQuit_box1.Visible = True
10600         .cmdQuit_box2.Visible = True
              ' ** Initializing message now has its own label!
10610         .cmdMenu.SetFocus
10620       End With
10630     Else
10640       If IsLoaded("frmLinkData", acForm) = False Then  ' ** Module Function: modFileUtilities.
10650         DoCmd.OpenForm "frmLinkData"
10660         Set frmPBar = Forms("frmLinkData")
10670         With frmPBar
10680           Set ctlPBox = .ProgBar_box
                'Set ctlPBar = .ProgBar_bar
10690         End With
10700       End If
10710     End If  ' ** blnPBar.

          ' ******************************
10720     If blnPBar = True Then
            ' ** 20. License form.
10730       dblPBar_ThisStep = 20#
10740       dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
10750       If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
10760       Else
10770         dblPBar_SubSteps = 0#
10780         dblPBar_ThisSubStep = 0#
10790         dblPBar_ThisSubIncrement = 0#
10800         dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
10810         frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
10820       End If
10830       DoEvents
10840     End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 20: License Check 2
10850     If blnDoFrm = True Then
10860       frm.InitMsg_lbl3.Caption = "License Check 2"
10870       DoEvents
10880     End If

          ' ** MAKE SURE A FAILED LICENSE CHECK EXITS!
10890     gblnMessage = False
          ' ** Check and confirm an up-to-date license.
10900     If Security_LicenseChk = False Then  ' ** Module Function: modSecurityFunctions
            ' ** When this isn't designated acDialog, InitializeTables() will continue. It should wait!
10910       gblnClosing = True  ' ** Borrowing this variable.
10920       DoCmd.OpenForm "frmLicense", , , , , acDialog, "frmMenu_Title"
10930       Select Case gblnClosing
            Case True
10940         DoCmd.Hourglass True
10950         DoEvents
10960       Case False
              ' ** Comes back False if canceled.
10970         blnRetVal = False
              ' ** gblnMessage = True means Quit has already been invoked.
10980       End Select
10990     End If

11000   End If  ' ** blnRetVal.

11010   If blnRetVal = True Then

          ' ** STATUS SUB
          ' ** Update table description.
          'TableDescriptionUpdate (gstrTrustDataLocation & gstrFile_DataName)  ' ** Module Function: modFileUtilities.
          'TableDescriptionUpdate (gstrTrustDataLocation & gstrFile_ArchDataName)  ' ** Module Function: modFileUtilities.
          'TableDescriptionUpdate (gstrTrustDataLocation & gstrFile_AuxDataName)  ' ** Module Function: modFileUtilities.

          ' ******************************
11020     If blnPBar = True Then
            ' ** 21. Table update.
11030       dblPBar_ThisStep = 21#
11040       dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
11050       If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
11060       Else
11070         dblPBar_SubSteps = 0#
11080         dblPBar_ThisSubStep = 0#
11090         dblPBar_ThisSubIncrement = 0#
11100         dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
11110         frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
11120       End If
11130       DoEvents
11140     End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 21: Database Path Records
11150     If blnDoFrm = True Then
11160       frm.InitMsg_lbl3.Caption = "Database Path Records"
11170       DoEvents
11180     End If

          ' ** Use this path whether or not TrustImport is present.
11190     strSpecPathFile = "FileSpec.mdb"
11200     strTmp01 = CurrentAppPath  ' ** Module Function: modFileUtilities.
11210     strTmp01 = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
11220     strTmp01 = strTmp01 & LNK_SEP & "Import"
11230     strSpecPathFile = strTmp01 & LNK_SEP & strSpecPathFile

          ' ** Update paths in tblDatabase and tblDatabase_Table_Link.
11240     Set dbs = CurrentDb
11250     With dbs
            ' ** Update qrySecurity_Stat_02 (qrySecurity_Stat_01 (tblDatabase_Table_Link,
            ' ** by specified [apath], [dpath], [spath], [prfx]), with tbllnk_connect_new).
11260       Set qdf1 = .QueryDefs("qrySecurity_Stat_03")
11270       With qdf1.Parameters
11280         ![apath] = gstrTrustAuxLocation
11290         ![dpath] = gstrTrustDataLocation  ' ** Includes final backslash.
11300         ![spath] = strSpecPathFile
11310         ![prfx] = ";" & LNK_IDENT
11320       End With
11330       qdf1.Execute
11340       Set qdf1 = Nothing
            ' ** Update qrySecurity_Stat_05 (qrySecurity_Stat_04 (tblDatabase,
            ' ** by specified [ppath], [apath], [dpath]), with dbs_path_new).
11350       Set qdf1 = .QueryDefs("qrySecurity_Stat_06")
11360       With qdf1.Parameters
11370         ![ppath] = CurrentAppPath  ' ** Module Function: modFileUtilities.                        ' ** Without final backslash.
11380         ![apath] = IIf(Right(gstrTrustAuxLocation, 1) = LNK_SEP, _
                Left(gstrTrustAuxLocation, (Len(gstrTrustAuxLocation) - 1)), gstrTrustAuxLocation)     ' ** Without final backslash.
11390         ![dpath] = IIf(Right(gstrTrustDataLocation, 1) = LNK_SEP, _
                Left(gstrTrustDataLocation, (Len(gstrTrustDataLocation) - 1)), gstrTrustDataLocation)  ' ** Without final backslash.
11400       End With
11410       qdf1.Execute
11420       Set qdf1 = Nothing
11430       .Close
11440     End With  ' ** dbs.
11450     Set dbs = Nothing

          ' ******************************
11460     If blnPBar = True Then
            ' ** 22. Foreign currency check.
11470       dblPBar_ThisStep = 22#
11480       dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
11490       If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
11500       Else
11510         dblPBar_SubSteps = 0#
11520         dblPBar_ThisSubStep = 0#
11530         dblPBar_ThisSubIncrement = 0#
11540         dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
11550         frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
11560       End If
11570       DoEvents
11580     End If  ' ** blnPBar.
          ' ******************************

          ' ** STATUS SUB 22: Foreign Currency Check
11590     If blnDoFrm = True Then
11600       frm.InitMsg_lbl3.Caption = "Foreign Currency Check"
11610       DoEvents
11620     End If

11630     gblnHasForEx = HasForEx_All  ' ** Module Function: modCurrencyFuncs.
11640     HasForEx_Load  ' ** Module Procedure: modCurrencyFuncs.

          ' ******************************
11650     If blnPBar = True Then
            ' ** 23. Finished.
11660       dblPBar_ThisStep = 23#
11670       dblPBar_ThisIncrement = arr_varPBar_Step(PB_INCR, dblPBar_ThisStep)
11680       If arr_varPBar_Step(PB_SUBS, dblPBar_ThisStep) = True Then
              ' ** None.
11690       Else
11700         dblPBar_SubSteps = 0#
11710         dblPBar_ThisSubStep = 0#
11720         dblPBar_ThisSubIncrement = 0#
11730         dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
11740         frmPBar.ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Form Procedure: frmLinkData.
11750       End If
11760       DoEvents
11770     End If  ' ** blnPBar.
          ' ******************************

11780   End If  ' ** blnRetVal.

11790   If blnRetVal = False Then
          ' ** gblnMessage = True, Quit has already been invoked.
          ' ** Will it matter if we do it again?
11800     QuitNow  ' ** Procedure: Below.
11810   End If

EXITP:
11820   Set ctlPBox = Nothing
11830   Set frmPBar = Nothing
11840   Set rstProg = Nothing
11850   Set rstData = Nothing
11860   Set tdf = Nothing
11870   Set rst = Nothing
11880   Set qdf1 = Nothing
11890   Set qdf2 = Nothing
11900   Set dbs = Nothing
11910   Set wrk = Nothing
11920   InitializeTables = blnRetVal
11930   Exit Function

ERRH:
11940   blnRetVal = False
11950   DoCmd.Hourglass False
11960   Select Case ERR.Number
        Case 3045  ' ** Couldn't use '|'; file already in use.
11970     MsgBox "This Program is already in use. Aborting...", vbCritical + vbOKOnly, "Trust Accountant Already Open"
11980     blnRetVal = False
11990     QuitNow  ' ** Procedure: Below.
12000   Case Else
12010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12020   End Select
12030   Resume EXITP

End Function

Public Sub QuitNow()
' ** Set options back to users original settings.

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "QuitNow"

12110   OpenAllDatabases False  ' ** Procedure: Below.

12120   SetOption_Run False  ' ** Function: Below.

12130   DoCmd.Quit

EXITP:
12140   Exit Sub

ERRH:
12150   Select Case ERR.Number
        Case Else
12160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12170   End Select
12180   Resume EXITP

End Sub

Public Function SizeChecks() As Integer
' ** Called by:
' **   frmMenu_Title
' **     Form_Open()
' ** Returns:
' **   SZ_OK   : All's well.
' **   SZ_COMP : TrustDta.mdb too big.
' **   SZ_RECS : A table has too many records.
' **   SZ_ERR  : Unexpected error.

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "SizeChecks"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strDatabase As String
        Dim blnWarnSeen As Boolean, blnWarnRecCnt As Boolean, blnWarnFileSize As Boolean, blnHasWarnPref As Boolean
        Dim lngSize As Long
        Dim strMessage As String
        Dim strSQL As String
        Dim lngRecs As Long, lngX As Long
        Dim intRetVal As Integer

12210   intRetVal = SZ_OK

12220   lngSize = 0
12230   strMessage = ""

        ' ** Find database location.
12240   strDatabase = gstrTrustDataLocation & gstrFile_DataName

12250   If FileExists(strDatabase) = False Then
12260     Beep
12270     MsgBox "Problem with data file:" & vbCrLf & vbCrLf & "    " & strDatabase, vbCritical + vbOKOnly, ("File Not Found" & Space(40))
12280     intRetVal = SZ_OK
12290   Else

12300     lngSize = FileLen(strDatabase)

12310     blnWarnRecCnt = True: blnWarnFileSize = True  ' ** Default to warnings ON
12320     blnHasWarnPref = False: blnWarnSeen = False   ' ** Default to no preference record, and warnings not seen.

          ' ** tblPreference_User, for 'chkWarningSeen', 'chkFileSizeWarning', 'chkRecordCountWarning', by specified [usr].
          ' ** I'd like to keep the warning options on frmLinkData hidden until they've seen the a message at least once.
12330     Set dbs = CurrentDb
12340     With dbs
12350       Set qdf = .QueryDefs("qryPreferences_03")  '##dbs_id
12360       With qdf.Parameters
12370         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
12380       End With
12390       Set rst1 = qdf.OpenRecordset
12400       With rst1
12410         If .BOF = True And .EOF = True Then
                ' ** No preferences for warning messages.
12420         Else
12430           .MoveLast
12440           lngRecs = .RecordCount
12450           .MoveFirst
12460           For lngX = 1& To lngRecs  ' ** Should be 3 max.
12470             Select Case ![ctl_name]
                  Case "chkWarningSeen"
12480               blnWarnSeen = ![prefuser_boolean]
12490               blnHasWarnPref = True
12500             Case "chkRecordCountWarning"
12510               blnWarnRecCnt = ![prefuser_boolean]
12520             Case "chkFileSizeWarning"
12530               blnWarnFileSize = ![prefuser_boolean]
12540             End Select
12550             If lngX < lngRecs Then .MoveNext
12560           Next
12570         End If
12580         .Close
12590       End With
12600       If blnHasWarnPref = False Then
              ' ** Append default 'chkWarningSeen' record (False) to tblPreference_User, by specified [usr].
12610         Set qdf = .QueryDefs("qryPreferences_04")  '##dbs_id
12620         With qdf.Parameters
12630           ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
12640         End With
12650         qdf.Execute
12660       End If
12670     End With

          ' ** Check if TrustDta.mdb is getting particularly large.
12680     If lngSize > glngWarnSize And (blnWarnSeen = False Or (blnWarnSeen = True And blnWarnFileSize = True)) Then
12690       intRetVal = SZ_COMP
12700       strMessage = "The size of your database has grown to the point where it is recommended that you compact it." & vbCrLf & _
              "  Main Menu --> Utility Menu --> Compact Data File" & vbCrLf & vbCrLf & _
              "If, after compacting, you continue to get this warning message, please contact Delta Data, Inc."
            ' ** Update tblPreference_User, for 'chkWarningSeen' = True, by specified [usr].
12710       Set qdf = dbs.QueryDefs("qryPreferences_05")  '##dbs_id
12720       With qdf.Parameters
12730         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
12740       End With
12750       qdf.Execute
12760     End If

12770     If intRetVal = SZ_OK And (blnWarnSeen = False Or (blnWarnSeen = True And blnWarnRecCnt = True)) Then
            ' ** Check if any tables have too many records.
12780       strSQL = vbNullString
12790       Set rst1 = CurrentDb.OpenRecordset("tblTemplate_m_TBL", dbOpenDynaset, dbReadOnly)
12800       rst1.MoveLast
12810       lngRecs = rst1.RecordCount
12820       rst1.MoveFirst
12830       For lngX = 1& To lngRecs
12840         If rst1![mtbl_ACTIVE] = True Then
12850           strSQL = "SELECT COUNT(*) AS NumRecs FROM [" & Trim(rst1![mtbl_NAME]) & "]"
12860           Set rst2 = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)
12870           If rst2![NumRecs] > glngWarnRecs Then
12880             intRetVal = SZ_RECS
12890             strMessage = "One or more of the tables in your database has a very large number of records in it." & vbCrLf & vbCrLf & _
                    "Please contact Delta Data, Inc., regarding this."
                  ' ** Update tblPreference_User, for 'chkWarningSeen' = True, by specified [usr].
12900             Set qdf = dbs.QueryDefs("qryPreferences_05")  '##dbs_id
12910             With qdf.Parameters
12920               ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
12930             End With
12940             qdf.Execute
12950             Exit For
12960           End If
12970         End If
12980         If lngX < lngRecs Then rst1.MoveNext
12990       Next
13000       rst1.Close
13010       rst2.Close
13020     End If

13030     dbs.Close

13040     If intRetVal > 0 Then
13050       If MsgBox(strMessage & vbCrLf & vbCrLf & "Would you like to continue into the system without compacting now?", _
                vbQuestion + vbYesNo + vbDefaultButton2, "Warning: Possible Degraded Response") = vbYes Then
13060         intRetVal = 0
13070       End If
13080     End If

13090   End If

EXITP:
13100   Set rst1 = Nothing
13110   Set rst2 = Nothing
13120   Set qdf = Nothing
13130   Set dbs = Nothing
13140   SizeChecks = intRetVal
13150   Exit Function

ERRH:
13160   intRetVal = SZ_ERR
13170   Select Case ERR.Number
        Case Else
13180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13190   End Select
13200   Resume EXITP

End Function

Public Function IniFile_Set(strSection As String, strSubSection As String, strValue As String, strFile As String) As Boolean
' ** Example:
' **   strSection = "License"
' **   strSubSection = "Firm"
' **   strValue = EncodeString(Me.txtLicensedTo)
' **   strFile = gstrTrustDataLocation & gstrFile_LIC
' ** See also: zz_mod_MDEPrepFuncs.xIniFile_Set()

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "IniFile_Set"

        Dim lngRetVal As Long
        Dim blnRetVal As Boolean

13310   lngRetVal = WritePrivateProfileStringA(strSection, strSubSection, strValue, strFile)
13320   If lngRetVal = 0 Then
13330     blnRetVal = False
13340   Else
13350     blnRetVal = True
13360   End If

EXITP:
13370   IniFile_Set = blnRetVal
13380   Exit Function

ERRH:
13390   Select Case ERR.Number
        Case Else
13400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13410   End Select
13420   Resume EXITP

End Function

Public Function IniFile_Del(strSection As String, strSubSection As String, strFile As String) As Boolean
' ** Example:
' **   strSection = "License"
' **   strSubSection = "Firm"
' **   strValue = EncodeString(Me.txtLicensedTo)
' **   strFile = gstrTrustDataLocation & gstrFile_LIC

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "IniFile_Del"

        Dim intLen As Long
        Dim blnRetVal As Boolean

13510   intLen = WritePrivateProfileStringA(strSection, strSubSection, "", strFile)
13520   blnRetVal = True

EXITP:
13530   IniFile_Del = blnRetVal
13540   Exit Function

ERRH:
13550   Select Case ERR.Number
        Case Else
13560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13570   End Select
13580   Resume EXITP

End Function

Public Function IniFile_GetDataLoc() As Boolean
' ** gstrTrustDataLocation  INCLUDES FINAL BACKSLASH!

13600 On Error GoTo ERRH

        Dim blnRetVal As Boolean

        Const THIS_PROC As String = "IniFile_GetDataLoc"

13610   blnRetVal = False  ' ** Unless proven otherwise.

13620   gstrTrustDataLocation = IniFile_Get("Files", "Location", RET_ERR, CurrentAppPath & LNK_SEP & gstrFile_INI)  ' ** Module Function: modFileUtilities.
        'gstrTrustDataLocation = IniFile_Get("Files", "Location", RET_ERR, "C:\Program Files\Delta Data\Trust Accountant" & LNK_SEP & gstrFile_INI)
13630   If gstrTrustDataLocation <> RET_ERR Then
13640     blnRetVal = True
13650   End If

EXITP:
13660   IniFile_GetDataLoc = blnRetVal
13670   Exit Function

ERRH:
13680   DoCmd.Hourglass False
13690   blnRetVal = False
13700   Select Case ERR.Number
        Case Else
13710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13720   End Select
13730   Resume EXITP

End Function

Public Function IniFile_Get(strSection As String, strSubSection As String, strDefault As String, strFile As String) As String
' ** Example:
' **   strSection - the brackected section name.
' **   lstrSubString - the variable in the section.
' **   strDefault - the value to return if the subsection is not found.
' **   strFile - full path name location of the INI file (extension is needed).
' **   RETURNS - the value in the INI file or the default value.

13800 On Error GoTo ERRH

        Dim intLen As Long
        Dim strReturn As String
        Dim strRetVal As String

        Const THIS_PROC As String = "IniFile_Get"

13810   strRetVal = vbNullString
13820   strReturn = String(256, 0)

13830   intLen = GetPrivateProfileStringA(strSection, strSubSection, strDefault, strReturn, 255, strFile)
13840   If intLen = 0 Then
13850     strReturn = Trim(strDefault)
13860   Else
13870     strRetVal = Left(strReturn, intLen)
13880   End If

EXITP:
13890   IniFile_Get = strRetVal
13900   Exit Function

ERRH:
13910   DoCmd.Hourglass False
13920   Select Case ERR.Number
        Case Else
13930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13940   End Select
13950   Resume EXITP

End Function

Public Function CoOptions_Read() As Boolean
' ** Called by:
' **   modStatementParamFuncs1:
' **     Month_AfterUpdate_SP()
' **   modVersionConvertFuncs1:
' **     Version_Upgrade_02()
' **   frmJournal_Columns_Sub:
' **     Form_Open()
' **   frmJournal_Sub3_Purchase:
' **     Form_Load()
' **   frmJournal_Sub4_Sold:
' **     Form_Load()
' **   frmJournal_Sub5_Misc:
' **     Form_Load()
' **   frmMenu_Main:
' **     Form_Open()
' **   frmMenu_Maintenance:
' **     cmdMaint05_Click()
' **   frmOptions:
' **     Form_Open()
' **   frmRpt_CourtReports_CA:
' **     Form_Open()
' **     PreviewOrPrint()
' **     SendToFile_CA()
' **   frmRpt_CourtReports_FL:
' **     Form_Open()
' **     PreviewOrPrint()
' **     SendToFile_FL()
' **     SummaryNew_FL()
' **   frmRpt_CourtReports_NS:
' **     PreviewOrPrint()
' **   frmRpt_CourtReports_NY:
' **     Form_Open()
' **   rptCheckList_All:
' **     Report_Open()
' **   rptCheckList_Users:
' **     Report_Open()
' **   rptChecks_Blank:
' **     Report_Open()
' **   rptPostingJournal_Classic:
' **     Report_Open()
' **   rptPostingJournal_Classic_ForEx:
' **     Report_Open()
' **   rptPostingJournal_Column:
' **     Report_Open()
' **   modJrnlCol_Misc:
' **     JC_Msc_Pub_Reset()
' **   modQueryFunctions1:
' **     CoInfo()
' **     CoInfoGet()
' **     CoInfoGet_Block()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "CoOptions_Read"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngThisDbsID As Long
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

14010   blnRetVal = True   ' ** Unless proven otherwise.

        ' ** Initialize.
14020   gblnIncomeTaxCoding = False
14030   gblnRevenueExpenseTracking = False
14040   gblnAccountNoWithType = False
14050   gblnSeparateCheckingAccounts = False
14060   gblnTabCopyAccount = False
14070   gblnLinkRevTaxCodes = False
14080   gblnSpecialCapGainLoss = False
14090   gintSpecialCapGainLossOpt = 0

14100   Set dbs = CurrentDb
14110   With dbs

14120     Set qdf = .QueryDefs("qryCompanyInformation_01b")
14130     Set rst = qdf.OpenRecordset
14140     With rst

14150       If .BOF = True And .EOF = True Then  ' ** Need a blank record at least.
14160         .AddNew
14170         ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
14180         ![CoInfo_DateModified] = Now()
14190         .Update
14200       End If
14210       .MoveFirst

14220       gblnIncomeTaxCoding = ![IncomeTaxCoding]
14230       gblnRevenueExpenseTracking = ![RevenueExpenseTracking]
14240       gblnAccountNoWithType = ![AccountNoWithType]
14250       gblnSeparateCheckingAccounts = ![SeparateCheckingAccounts]
14260       gblnTabCopyAccount = ![TabCopyAccount]
14270       gblnLinkRevTaxCodes = ![LinkRevTaxCodes]
14280       gblnSpecialCapGainLoss = ![SpecialCapGainLoss]
14290       If ![SpecialCapGainLoss] = True Then
14300         Select Case IsNull(![SpecialCapGainLossOpt])
              Case True
14310           .Edit
14320           ![SpecialCapGainLossOpt] = 1
14330           .Update
14340         Case False
14350           If ![SpecialCapGainLossOpt] = 0 Then
14360             .Edit
14370             ![SpecialCapGainLossOpt] = 1
14380             .Update
14390           End If
14400         End Select
14410       End If
14420       gintSpecialCapGainLossOpt = ![SpecialCapGainLossOpt]
14430       gstrCo_Name = Nz(![CoInfo_Name], vbNullString)
14440       gstrCo_Address1 = Nz(![CoInfo_Address1], vbNullString)
14450       gstrCo_Address2 = Nz(![CoInfo_Address2], vbNullString)
14460       gstrCo_City = Nz(![CoInfo_City], vbNullString)
14470       gstrCo_State = Nz(![CoInfo_State], vbNullString)
14480       gstrCo_Zip = Nz(![CoInfo_Zip], vbNullString)
14490       gstrCo_Country = Nz(![CoInfo_Country], vbNullString)
14500       gstrCo_PostalCode = Nz(![CoInfo_PostalCode], vbNullString)
14510       gstrCo_Phone = Nz(![CoInfo_Phone], vbNullString)

14520       .Close
14530     End With
14540     Set rst = Nothing
14550     Set qdf = Nothing

14560     If gblnSpecialCapGainLoss = True Then
14570       varTmp00 = DLookup("[Dividend]", "AssetType", "[assettype] = '80'")
14580       If varTmp00 = False Then
              ' ** Update qryOptions_SpecialCap_01 (AssetType, just '80', with Dividend_new = True).
14590         Set qdf = .QueryDefs("qryOptions_SpecialCap_02")
14600         qdf.Execute
14610         Set qdf = Nothing
14620       End If
14630     End If

14640     lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

14650     .Close
14660   End With

EXITP:
14670   Set rst = Nothing
14680   Set qdf = Nothing
14690   Set dbs = Nothing
14700   CoOptions_Read = blnRetVal
14710   Exit Function

ERRH:
14720   blnRetVal = False
14730   Select Case ERR.Number
        Case Else
14740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14750   End Select
14760   Resume EXITP

End Function

Public Function CoOptions_Write(strField As String, blnValue As Boolean) As Boolean

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "CoOptions_Write"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnRetVal As Boolean

14810   blnRetVal = True

14820   Set dbs = CurrentDb
14830   With dbs
14840     Set rst = .OpenRecordset("CompanyInformation", dbOpenDynaset, dbConsistent)
14850     With rst
14860       If .BOF = True And .EOF = True Then
14870         .AddNew
14880       Else
14890         .Edit
14900       End If
14910       .Fields(strField) = blnValue
14920       ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
14930       ![CoInfo_DateModified] = Now()
14940       .Update
14950       .Close
14960     End With
14970     .Close
14980   End With

EXITP:
14990   Set rst = Nothing
15000   Set dbs = Nothing
15010   CoOptions_Write = blnRetVal
15020   Exit Function

ERRH:
15030   blnRetVal = False
15040   Select Case ERR.Number
        Case Else
15050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15060   End Select
15070   Resume EXITP

End Function

Public Function GetVersionInfo(varV_Main As Variant, varV_Minor As Variant, varV_Revision As Variant) As String
' ** Return a formatted version string

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "GetVersionInfo"

        Dim strTmp01 As String
        Dim strRetVal As String

15110   strRetVal = vbNullString

15120   If IsNull(varV_Main) = False Then
15130     strRetVal = CStr(varV_Main) & "."
15140     If IsNull(varV_Minor) = False Then
15150       strTmp01 = CStr(varV_Minor)
15160       If Len(strTmp01) = 2 Then
15170         If Right(strTmp01, 1) = "0" Then  ' ** The center number must just be 1 - 9.
15180           strTmp01 = Left(strTmp01, 1) & "."
15190           If IsNull(varV_Revision) = False Then
15200             strTmp01 = strTmp01 & Left(CStr(varV_Revision) & "00", 2)
15210           Else
15220             strTmp01 = strTmp01 & "00"
15230           End If
15240         Else
15250           strTmp01 = Left(strTmp01, 1) & "." & Mid(strTmp01, 2)
15260           If IsNull(varV_Revision) = False Then
15270             strTmp01 = strTmp01 & Left(CStr(varV_Revision) & "00", 2)
15280           End If
15290         End If
15300         strRetVal = strRetVal & strTmp01
15310       Else
15320         strRetVal = strRetVal & strTmp01 & "."
15330         If IsNull(varV_Revision) = False Then
15340           strRetVal = strRetVal & Left(CStr(varV_Revision) & "00", 2)
15350         Else
15360           strRetVal = strRetVal & "00"
15370         End If
15380       End If
15390     Else
15400       strRetVal = strRetVal & "0."
15410       If IsNull(varV_Revision) = False Then
15420         strRetVal = strRetVal & Left(CStr(varV_Revision) & "00", 2)
15430       Else
15440         strRetVal = strRetVal & "00"
15450       End If
15460     End If
15470   Else
15480     strRetVal = "2.x.x"
15490   End If

EXITP:
15500   GetVersionInfo = strRetVal
15510   Exit Function

ERRH:
15520   strRetVal = RET_ERR
15530   Select Case ERR.Number
        Case Else
15540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15550   End Select
15560   Resume EXITP

End Function

Public Function GetLinkList() As Variant

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "GetLinkList"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngLinks As Long, arr_varLink As Variant
        Dim arr_varTmp00() As Variant
        Dim varRetVal As Variant

15610   varRetVal = Empty

        ' ** Create a default array.
15620   ReDim arr_varTmp00(L_NAM, 0)
15630   arr_varTmp00(L_NAM, 0) = RET_ERR

15640   Set dbs = CurrentDb
15650   With dbs

          ' ** tblTemplate_m_TBL, linked to qrySystemStartup_11 (tblTemplate_Database_Table_Link,
          ' ** by specified CurrentAppName()), currently active tables.
15660     Set qdf = .QueryDefs("qrySystemStartup_04_m_TBL")
15670     Set rst = qdf.OpenRecordset
15680     With rst
15690       If .BOF = True And .EOF = True Then
15700         varRetVal = arr_varTmp00
15710       Else
15720         .MoveLast
15730         lngLinks = .RecordCount
15740         .MoveFirst
15750         arr_varLink = .GetRows(lngLinks)
              ' ********************************************************
              ' ** Array: arr_varLink()
              ' **
              ' **   Field  Element  Name                   Constant
              ' **   =====  =======  =====================  ==========
              ' **     1       0     mtbl_ID                L_ID
              ' **     2       1     mtbl_NAME              L_NAM
              ' **     3       2     mtbl_AUTONUMBER        L_AUT
              ' **     4       3     mtbl_ORDER             L_ORD
              ' **     5       4     mtbl_NEWRecs           L_NEW
              ' **     6       5     mtbl_ACTIVE            L_ACT
              ' **     7       6     mtbl_DTA               L_DTA
              ' **     8       7     mtbl_ARCH              L_ARC
              ' **     9       8     mtbl_AUX               L_AUX
              ' **    10       9     mtbl_SOURCE            L_SRC
              ' **    11      10     contype_type           L_TYP
              ' **    12      11     tbllnk_connect         L_LNKT
              ' **    13      12     tbllnk_connect_CURR    L_LNKC
              ' **    14      13     tbllnk_fnd             L_FND
              ' **    15      14     tbllnk_fix             L_FIX
              ' **
              ' ********************************************************
15760         varRetVal = arr_varLink
15770       End If
15780       .Close
15790     End With
15800     .Close
15810   End With

EXITP:
15820   GetLinkList = varRetVal
15830   Exit Function

ERRH:
15840   varRetVal = arr_varTmp00
15850   Select Case ERR.Number
        Case Else
15860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15870   End Select
15880   Resume EXITP

End Function

Public Function IsTAOpen(Optional varList As Variant, Optional varApp As Variant) As Boolean
' ** Check for another instance of Trust Accountant.
' ** No longer mistakes Windows Explorer window showing
' ** Trust Accountant directory for another instance.

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "IsTAOpen"

        Dim lngRet As Long, lngParam As Long
        Dim blnList As Boolean, blnAdmin As Boolean, blnImport As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

15910   blnRetVal = False  ' ** Default not open.

15920   Select Case IsMissing(varList)
        Case True
15930     blnList = False
15940     blnAdmin = False
15950     blnImport = False
15960   Case False
15970     blnList = CBool(varList)
15980     Select Case IsMissing(varApp)
          Case True
15990       blnAdmin = False
16000       blnImport = False
16010     Case False
16020       Select Case varApp
            Case "Admin"
16030         blnAdmin = True
16040         blnImport = False
16050       Case "Import"
16060         blnAdmin = False
16070         blnImport = True
16080       End Select
16090     End Select
16100   End Select

16110   glngClasses = 0&
16120   ReDim garr_varClass(CLS_ELEMS, 0)

16130   blnWindowVisible = True  ' ** Only check visible windows.

        ' ** List all open windows into garr_varClass() array.
16140   lngRet = EnumWindows(AddressOf Win_Class_Load, lngParam)  ' ** API Function: modWindowFunctions, Below.

16150   For lngX = 0& To (glngClasses - 1&)
          ' ** Positive hit must be Access (OMain),
          ' ** titled "Trust Accountant" (gstrRegKeyName),
          ' ** and not this instance (hWndAccessApp).
16160     If garr_varClass(CLS_CLASS, lngX) = "OMain" Then
16170       Select Case blnList
            Case True
16180         Debug.Print "'" & garr_varClass(CLS_TITLE, lngX)
              ' ** Trust Accountant™
              ' ** Trust Accountant™ Administration
              ' ** Trust Import™
16190       Case False
16200         If blnAdmin = True Then
16210           If InStr(garr_varClass(CLS_TITLE, lngX), gstrRegKeyName) > 0 And _
                    InStr(garr_varClass(CLS_TITLE, lngX), "Administration") > 0 Then
16220             blnRetVal = True
16230             Exit For
16240           End If
16250         ElseIf blnImport = True Then
16260           If InStr(garr_varClass(CLS_TITLE, lngX), "Trust Import") > 0 Then
16270             blnRetVal = True
16280             Exit For
16290           End If
16300         Else
16310           If InStr(garr_varClass(CLS_TITLE, lngX), gstrRegKeyName) > 0 And _
                    InStr(garr_varClass(CLS_TITLE, lngX), "Administration") = 0 And _
                    garr_varClass(CLS_HWND, lngX) <> hWndAccessApp Then
16320             blnRetVal = True
16330             Exit For
16340           End If
16350         End If
16360       End Select
16370     End If
16380   Next

        ' ** Some application class names:
        ' **   Application          Class Name
        ' **   ===================  ===========================
        ' **   Access               OMain
        ' **   Excel                XLMAIN
        ' **   FrontPage            FrontPageExplorerWindow40
        ' **   Outlook              rctrl_renwnd32
        ' **   PowerPoint 95        PP7FrameClass
        ' **   PowerPoint 97        PP97FrameClass
        ' **   PowerPoint 2000      PP9FrameClass
        ' **   PowerPoint XP        PP10FrameClass
        ' **   Project              JWinproj-WhimperMainClass
        ' **   Visual Basic Editor  wndclass_desked_gsk
        ' **   Word                 OpusApp
        ' **   Calculator           SciCalc
        ' **   Windows Explorer     ExploreWClass
        ' **   Windows Explorer     CabinetWClass
        ' **   Internet Explorer    IEFrame
        ' **   Windows Media Player WMPlayerApp
        ' **   Solitaire            Solitaire
        ' **   System Tray          Shell_TrayWnd
        ' **   Program Manager      Progman

EXITP:
16390   IsTAOpen = blnRetVal
16400   Exit Function

ERRH:
16410   Select Case ERR.Number
        Case Else
16420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16430   End Select
16440   Resume EXITP

End Function

Public Function IsHelpOpen() As Boolean
' ** My Acrobat ClassName: AcrobatSDIWindow

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "IsHelpOpen"

        Dim lngRet As Long
        Dim lngParam As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

16510   blnRetVal = False  ' ** Default not open.

16520   glngClasses = 0&
16530   ReDim garr_varClass(CLS_ELEMS, 0)

16540   blnWindowVisible = True  ' ** Only check visible windows.

        ' ** List all open windows into garr_varClass() array.
16550   lngRet = EnumWindows(AddressOf Win_Class_Load, lngParam)  ' ** API Function: modWindowFunctions, Below.

16560   For lngX = 0& To (glngClasses - 1&)
16570     If garr_varClass(CLS_CLASS, lngX) = "AcrobatSDIWindow" And _
              garr_varClass(CLS_HWND, lngX) <> hWndAccessApp Then
16580       blnRetVal = True
16590       Exit For
16600     End If
16610   Next

EXITP:
16620   IsHelpOpen = blnRetVal
16630   Exit Function

ERRH:
16640   blnRetVal = False
16650   Select Case ERR.Number
        Case Else
16660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16670   End Select
16680   Resume EXITP

End Function

Public Function IsLocalData() As Boolean

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "IsLocalData"

        Dim strCurAppPath As String, strCurDataPath As String
        Dim blnRetVal As Boolean

16710   blnRetVal = False

16720   If gstrTrustDataLocation = vbNullString Then
16730     IniFile_GetDataLoc  ' ** Procedure: Above.
16740   End If

16750   strCurAppPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
16760   strCurDataPath = gstrTrustDataLocation
16770   If Right(strCurDataPath, 1) = LNK_SEP Then strCurDataPath = Left(strCurDataPath, (Len(strCurDataPath) - 1))
16780   If strCurDataPath = strCurAppPath & LNK_SEP & "Database" Then
16790     blnRetVal = True
16800   End If

EXITP:
16810   IsLocalData = blnRetVal
16820   Exit Function

ERRH:
16830   blnRetVal = False
16840   Select Case ERR.Number
        Case Else
16850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16860   End Select
16870   Resume EXITP

End Function

Public Function SaveLoadTime() As Boolean

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "SaveLoadTime"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean, blnSkip As Boolean
        Dim blnRetVal As Boolean

16910   blnRetVal = True
16920   blnContinue = False

16930   blnSkip = True
16940   If blnSkip = False Then

16950     If blnContinue = True Then
16960       Set dbs = CurrentDb
16970       With dbs
              ' ** tblXAdmin_Load.
16980         Set qdf = .QueryDefs("qrySecurity_Stat_07")
16990         Set rst = qdf.OpenRecordset
17000         With rst
17010           If .BOF = True And .EOF = True Then
                  ' ** First recorded entry.
17020           Else
                  '.FindFirst "[xadload_start] = #" & Format(gdatLoadStart, "mm/dd/yyyy hh:nn:ss AM/PM") & "#"
17030             If .NoMatch = False Then
                    ' ** Already saved.
17040               blnContinue = False
17050             End If
17060           End If
17070           If blnContinue = True Then
                  '.AddNew
                  '![xadload_start] = gdatLoadStart
                  '![xadload_end] = gdatLoadEnd
                  '![xadload_time] = CDate(CDbl(gdatLoadEnd) - CDbl(gdatLoadStart))
                  '.Update
17080           End If
17090           .Close
17100         End With
17110         .Close
17120       End With
17130     End If  ' ** blnContinue.

17140   End If  ' ** blnSkip.

EXITP:
17150   Set rst = Nothing
17160   Set qdf = Nothing
17170   Set dbs = Nothing
17180   SaveLoadTime = blnRetVal
17190   Exit Function

ERRH:
17200   blnRetVal = False
17210   Select Case ERR.Number
        Case Else
17220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17230   End Select
17240   Resume EXITP

End Function

Private Function Win_Class_Load(ByVal lngHWnd As Long, ByVal lngParam As Long) As Boolean
' ** Called by:
' **   IsTAOpen(), Above.
' **   IsHelpOpen(), Above.

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "Win_Class_Load"

        Dim strClass As String, strTitle1 As String
        Dim strClassBuf As String * 255, strTitle1Buf As String * 255
        Dim blnFound As Boolean
        Dim intVis As Integer
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

17310   blnRetVal = True

17320   strClass = GetClassName(lngHWnd, strClassBuf, 255)  ' ** API Function: Above.
17330   strClass = StripNulls(strClassBuf)  ' ** Function: Below.
17340   strTitle1 = GetWindowText(lngHWnd, strTitle1Buf, 255)  ' ** API Function: Above.
17350   strTitle1 = StripNulls(strTitle1Buf)  ' ** Function: Below.

17360   If blnWindowVisible = False Then
          ' ** Check both visible and hidden windows.
17370     intVis = 1
17380   Else
          ' ** Check only visible windows.
17390     intVis = IsWindowVisible(lngHWnd)  ' ** API Function: Above.
17400   End If

        ' ** Check if Window is a parent and visible.
17410   If GetParent(lngHWnd) = 0 And intVis = 1 Then  ' ** API Function: Above.
17420     blnFound = False
17430     For lngX = 0& To (glngClasses - 1)
17440       If garr_varClass(CLS_CLASS, lngX) = strClass And _
                garr_varClass(CLS_TITLE, lngX) = strTitle1 And _
                garr_varClass(CLS_HWND, lngX) = lngHWnd Then
17450         blnFound = True
17460         Exit For
17470       End If
17480     Next
17490     If blnFound = False Then
17500       glngClasses = glngClasses + 1&
17510       lngE = glngClasses - 1&
17520       ReDim Preserve garr_varClass(CLS_ELEMS, lngE)
17530       garr_varClass(CLS_CLASS, lngE) = strClass
17540       garr_varClass(CLS_TITLE, lngE) = strTitle1
17550       garr_varClass(CLS_HWND, lngE) = lngHWnd
17560       garr_varClass(CLS_PARENT, lngE) = GetParent(lngHWnd)  ' ** API Function: modWindowFunctions.
17570       garr_varClass(CLS_VISIBLE, lngE) = IsWindowVisible(lngHWnd)  ' ** API Function: modWindowFunctions.
17580     End If
17590   End If

EXITP:
17600   Win_Class_Load = blnRetVal
17610   Exit Function

ERRH:
17620   blnRetVal = False
17630   Select Case ERR.Number
        Case Else
17640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17650   End Select
17660   Resume EXITP

End Function

Public Function Setup_FileExt(blnAuxLoc As Boolean, Optional varDataPath As Variant) As Boolean
' ** Called by:
' **   InitializeTables(), Above.

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "Setup_FileExt"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim strAppName As String, strAppPath As String, strDataPath As String
        Dim blnChanged As Boolean, blnLocal As Boolean
        Dim lngLoopCnt As Long
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

17710   blnRetVal = True
17720   blnChanged = False: blnLocal = False

17730   strAppName = CurrentAppName  ' ** Module Function: modFileUtilities.
17740   strAppPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
17750   Select Case IsMissing(varDataPath)
        Case True
17760     strDataPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.
17770   Case False
17780     strDataPath = Trim(varDataPath)
17790     If strDataPath = vbNullString Then
17800       strDataPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.
17810     Else
17820       If Right(strDataPath, 1) = LNK_SEP Then strDataPath = Left(strDataPath, (Len(strDataPath) - 1))
17830     End If
17840   End Select

17850   Set dbs = CurrentDb
17860   With dbs

          ' ** tblDatabase, just needed fields.
17870     Set qdf = .QueryDefs("qrySystemStartup_08a")
17880 On Error Resume Next
17890     Set rst = qdf.OpenRecordset
17900     If ERR.Number <> 0 Then
17910 On Error GoTo ERRH
            ' ** tblTemplate_Database, just needed fields.
17920       Set qdf = .QueryDefs("qrySystemStartup_08b")
17930       Set rst = qdf.OpenRecordset
17940       blnLocal = True
17950     Else
17960 On Error GoTo ERRH
17970     End If

17980     Select Case blnLocal
          Case True
17990       lngLoopCnt = 1&
18000     Case False
18010       lngLoopCnt = 2&
18020     End Select

18030     For lngX = 1& To lngLoopCnt

            ' ** See to it that both tables remain synchronized.
18040       If lngX = 2& Then
              ' ** tblTemplate_Database, just needed fields.
18050         Set qdf = .QueryDefs("qrySystemStartup_08b")
18060         Set rst = qdf.OpenRecordset
18070       End If

18080       With rst
18090         .MoveFirst

              ' ** Check for MDE vs. MDB.
18100         If ![dbs_name] <> strAppName Then  ' ** Only checked on first record, dbs_id = 1, this database.
18110           .Edit
18120           ![dbs_name] = strAppName
18130           ![dbs_datemodified] = Now()
18140           .Update
18150           blnChanged = True
18160         End If

18170         .MoveLast
18180         lngRecs = .RecordCount
18190         .MoveFirst

              ' ** Check for the current paths.
18200         For lngY = 1& To lngRecs
18210           If lngY = 1& Then  ' ** Only 1st record, Trust.mdb/Trust.mde.
18220             If ![dbs_path] <> strAppPath Then
18230               .Edit
18240               ![dbs_path] = strAppPath
18250               ![dbs_datemodified] = Now()
18260               .Update
18270               blnChanged = True
18280             End If
18290           ElseIf Left(![dbs_name], 8) = "TrstXAdm" Or Left(![dbs_name], 11) = "TrustImport" Then
                  ' ** Leave these alone!
18300           Else
18310             Select Case blnAuxLoc
                  Case True
18320               If Left(![dbs_name], 8) = "TrustAux" Then
18330                 If ![dbs_path] <> strAppPath Then
18340                   .Edit
18350                   ![dbs_path] = strAppPath
18360                   ![dbs_datemodified] = Now()
18370                   .Update
18380                   blnChanged = True
18390                 End If
18400               Else
18410                 If ![dbs_path] <> strDataPath Then
18420                   .Edit
18430                   ![dbs_path] = strDataPath
18440                   ![dbs_datemodified] = Now()
18450                   .Update
18460                   blnChanged = True
18470                 End If
18480               End If
18490             Case False
18500               If ![dbs_path] <> strDataPath Then
18510                 .Edit
18520                 ![dbs_path] = strDataPath
18530                 ![dbs_datemodified] = Now()
18540                 .Update
18550                 blnChanged = True
18560               End If
18570             End Select  ' ** blnAuxLoc.
18580           End If
18590           If lngY < lngRecs Then .MoveNext
18600         Next  ' ** lngY.

18610         .Close
18620       End With  ' ** rst.
18630       Set rst = Nothing

18640     Next  ' ** lngX.

          ' ** tblSecurity_License, just needed fields.
18650     Set qdf = .QueryDefs("qrySystemStartup_09")
18660     Set rst = qdf.OpenRecordset
18670     With rst
18680       .MoveFirst
18690       If ![seclic_clientpath_ta] <> strAppPath Then
18700         .Edit
18710         ![seclic_clientpath_ta] = strAppPath
18720         ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
18730         ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
18740         ![seclic_datemodified] = Now()
18750         .Update
18760         blnChanged = True
18770       End If
18780       If ![seclic_datapath_ta] <> strDataPath Then
18790         .Edit
18800         ![seclic_datapath_ta] = strDataPath
18810         ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
18820         ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
18830         ![seclic_datemodified] = Now()
18840         .Update
18850         blnChanged = True
18860       End If
18870       Select Case blnAuxLoc
            Case True
18880         If ![seclic_auxiliarypath] <> strAppPath Then
18890           .Edit
18900           ![seclic_auxiliarypath] = strAppPath
18910           ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
18920           ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
18930           ![seclic_datemodified] = Now()
18940           .Update
18950           blnChanged = True
18960         End If
18970       Case False
18980         If ![seclic_auxiliarypath] <> strDataPath Then
18990           .Edit
19000           ![seclic_auxiliarypath] = strDataPath
19010           ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
19020           ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
19030           ![seclic_datemodified] = Now()
19040           .Update
19050           blnChanged = True
19060         End If
19070       End Select  ' ** blnAuxLoc.
19080       .Close
19090     End With

19100     .Close
19110   End With

19120   If blnChanged = True Then
          'BEEP HERE!
          'Beep
19130   End If

EXITP:
19140   Set rst = Nothing
19150   Set qdf = Nothing
19160   Set dbs = Nothing
19170   Setup_FileExt = blnRetVal
19180   Exit Function

ERRH:
19190   blnRetVal = False
19200   Select Case ERR.Number
        Case Else
19210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19220   End Select
19230   Resume EXITP

End Function

Public Sub OpenAllDatabases(blnInit As Boolean)
' ** Open a handle to all databases and keep it open during the entire time the application runs.
' ** Params  : blnInit   TRUE to initialize (call when application starts).
' **                     FALSE to close (call when application ends).
' ** From    : Total Visual SourceBook

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "OpenAllDatabases"

        Dim strDbsName As String
        Dim strMsg As String
        Dim intX As Integer

        ' ** Maximum number of back end databases to link.
        Const MAX_DBS As Integer = 3

19310   Select Case blnInit
        Case True

19320     ReDim dbsOpen(1 To MAX_DBS)

19330     For intX = 1 To MAX_DBS

            ' ** Specify your back end databases.
19340       Select Case intX
            Case 1
19350         strDbsName = gstrTrustDataLocation & gstrFile_DataName
19360       Case 2
19370         strDbsName = gstrTrustDataLocation & gstrFile_ArchDataName
19380       Case 3
19390         strDbsName = gstrTrustAuxLocation & gstrFile_AuxDataName
19400       End Select
19410       strMsg = vbNullString

19420 On Error Resume Next
19430       Set dbsOpen(intX) = OpenDatabase(strDbsName)
19440       If ERR.Number > 0 Then
19450         strMsg = "Trouble opening database: " & strDbsName & vbCrLf & _
                "Make sure the drive is available." & vbCrLf & vbCrLf & _
                "Error:" & vbTab & vbTab & CStr(ERR.Number) & vbCrLf & _
                "Description:" & vbTab & ERR.description & vbCrLf & vbCrLf & _
                "Module:" & vbTab & vbTab & THIS_NAME & vbCrLf & _
                "Sub/Function:" & vbTab & THIS_PROC & "()" & vbCrLf & _
                "Line:" & vbTab & vbTab & CStr(Erl)
19460       End If

19470 On Error GoTo ERRH
19480       If strMsg <> vbNullString Then
19490         MsgBox strMsg, vbCritical + vbOKOnly, "Open Database Failed"
19500         Exit For
19510       End If

19520     Next  ' ** IntX.

19530   Case False

19540 On Error Resume Next
19550     For intX = 1 To MAX_DBS
19560       dbsOpen(intX).Close
19570     Next  ' ** IntX
19580 On Error GoTo ERRH

19590   End Select

EXITP:
19600   Exit Sub

ERRH:
19610   Select Case ERR.Number
        Case Else
19620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19630   End Select
19640   Resume EXITP

End Sub
