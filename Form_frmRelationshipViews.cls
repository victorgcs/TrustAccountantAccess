VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRelationshipViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRelationshipViews"

'VGC 11/23/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   About Tab:        A {tabRelation_pagAbout}
' **   Browse:           B {cmdSelectMDB}
' **   Close:            C {cmdClose}
' **   Delete Layout:    D {cmdDelete}
' **   External Tab:     E {tabRelation_pagExternalRel}
' **   Get BLOB:         G {cmdGetBlob}
' **   BLOB Tab:         I {tabRelation_pagBlobs}
' **   Local Tab:        L {tabRelation_pagLocalRel}
' **   Restore Local:    R {cmdRestore}
' **   Restore External: R {cmdRestoreExternal}
' **   Save Layout:      S {cmdSaveLayout}
' **   Lebans Web:       W {cmdWeb}
' **   Exit:             X {cmdClose}

' ** Minimum Left control-to-border on tab control: 135 Twips
' **   I'll use: 180 Twips
' ** Minimum Top control-to-border on tab control: 450 Twips (10 pt. font on 1-line tabs)
' **   I'll use: 480 Twips
'Top   : 90
'Left  : 105
'Width : 7665
'Height: 3645

' ** DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 2000 VBA
' **
' ** Copyright: Stephen Lebans - Lebans Holdings 1999 Inc.
' **            Please feel free to use this code within your own projects,
' **            both private and commercial, with no obligation.
' **            You may not resell this code by itself or as part of a collection.
' **
' ** Name:      Save/Restore Relationship Window view
' **
' ** Version:   1.3
' **
' ** Purpose:   1) To allow the saving of the layout of the Relationship window to a table.
' **            2) To allow the restoration of the layout of the Relationship Window from a table.
' **            3) To allow the saving/restoration of multiple Relationship Window views.
' **            4) To allow the importing of the layout of the Relationship Window from an external MDB.
' **  
' ** Requires: This Form must be bound to the table named tblRelationshipViews
' **
' ** Author:    Stephen Lebans
' **
' ** Email:     Stephen@lebans.com
' **
' ** Web Site:  www.lebans.com
' **
' ** Date:      February 04, 2004, 11:11:11 PM
' **
' ** Credits:   Steve Jorgensen for the idea!
' **
' ** BUGS:      Please report any bugs to:
' **            Stephen@lebans.com
' **
' ** What's Missing:
' **            DAO error handling
' **            All other Error handling
' **            Add it yourself!
' **
' ** How it Works:
' **            Walk through the source code!<grin>
' **
' ** Enjoy
' ** Stephen Lebans

'Private Type POINTAPI  ' ** Public Type: modWindowFunctions.
'  X As Long
'  Y As Long
'End Type

'Private Type RECT  ' ** Public Type: modWindowFunctions.
'  Left As Long
'  Top As Long
'  Right As Long
'  Bottom As Long
'End Type

Private Const AAAlength     As Integer = 12
Private Const FFFlength     As Integer = 8
Private Const Padding       As Integer = 12
Private Const NameLengthMax As Integer = 128
' ** 64 Char MAX for a DAO Table Name * 2 = Unicode

Private Type RELBLOB
  Sig As Long
  AAAs(1 To AAAlength) As Byte
  RelWinX1  As Long
  RelWinY1 As Long
  RelWinX2  As Long
  RelWinY2 As Long
  Blank As Long
  FFFs(1 To FFFlength) As Byte
  ClientRectX As Long
  ClientRectY As Long
  Pad(1 To Padding) As Byte
  NumWindows As Long
End Type

Private Type RELWINDOW
  RelWinX1  As Long
  RelWinY1 As Long
  RelWinX2  As Long
  RelWinY2 As Long
  Junk As Long
  WinName As String * NameLengthMax
  Junk1 As Long
  WinNameMaster As String * NameLengthMax
  'Pad(1 To Padding) As Byte
  Junk2 As Long
End Type

'Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long  ' ** Public Function: modWindowFunctions.

'Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" _
'  (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long  ' ** Public Function: modWindowFunctions.

'Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long  ' ** Public Function: modWindowFunctions.

'Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long  ' ** Public Function: modWindowFunctions.

'Private Declare Function SetWindowPos Lib "user32.dll" _
'  (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
'  ByVal Y As Long, ByVal Cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long  ' ** Public Function: modWindowFunctions.

'Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" _
'  (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  ' ** Public Function: modWindowFunctions.

'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)  ' ** Public Procedure: modWindowFunctions.

' ** SetWindowPos() Constants.
'Private Const SWP_SHOWWINDOW = &H40  ' ** Public Constant: modWindowFunctions.

' ** GetWindow() Constants.
'Private Const GW_HWNDNEXT = 2  ' ** Public Constant: modWindowFunctions.
'Private Const GW_CHILD = 5  ' ** Public Constant: modWindowFunctions.

Private blnDatabaseLinked As Boolean
Private strNewName As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim varTmp00 As Variant, strTmp01 As String

110     With Me

120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strCallingForm = Trim(.OpenArgs)  ' ** Unlikely to ever be called.
150         End If
160       End If

170       DoCmd.Restore

180       .PictureData = .Form_RelationshipViews_img

190       blnDatabaseLinked = False

200       varTmp00 = DLookup("[relview_name]", "qryRelationView_02")  ' ** This will return the first one in the query.
210       If IsNull(varTmp00) = False Then
220         If Trim(varTmp00) <> vbNullString Then
230           .relview_stored_lbx = 0  'varTmp00
240         End If
250       End If

260       strNewName = GetNewName  ' ** Function: Below.
270       .relview_name_new.DefaultValue = Chr(34) & strNewName & Chr(34)

280       strTmp01 = Parse_Path(CurrentDb.Name)  ' ** Module Function: modFileUtilities.

290       If InStr(strTmp01, gstrDir_Dev) > 0 Then
            ' ** All's well.
300       ElseIf InStr(strTmp01, Parse_Path(gstrDir_Dev)) > 0 Then  ' ** Module Function: modFileUtilities.
            ' ** It's in my TrustAccountant directory, but where?
310         If Parse_File(strTmp01) <> Parse_File(gstrDir_Dev) Then  ' ** Module Function: modFileUtilities.
320           Beep
330           MsgBox "gstrDir_Dev is not correctly set!", vbCritical + vbOKOnly, "Errors Coming"
340         Else
              ' ** Not sure what else.
350         End If
360       Else
            ' ** We'll see if it complains.
370       End If

380       .cmdRestore.SetFocus

390     End With

400     If Cancel = -1 Then
          ' ** Just for the record.
410       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
420     End If

EXITP:
430     Exit Sub

ERRH:
440     Select Case ERR.Number
        Case Else
450       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
460     End Select
470     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Alt keys.
550     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyB
570         With Me
580           intRetVal = 0
590           Select Case .tabRelation
              Case .tabRelation_pagExternalRel.PageIndex
600             .cmdSelectMDB.SetFocus
610             cmdSelectMDB_Click  ' ** Procedure: Below.
620           Case Else
630             Beep
640           End Select
650         End With
660       Case vbKeyD
670         With Me
680           intRetVal = 0
690           Select Case .tabRelation
              Case .tabRelation_pagLocalRel.PageIndex
700             .cmdDelete.SetFocus
710             cmdDelete_Click  ' ** Procedure: Below.
720           Case Else
730             Beep
740           End Select
750         End With
760       Case vbKeyG
770         With Me
780           intRetVal = 0
790           Select Case .tabRelation
              Case .tabRelation_pagBlobs.PageIndex
800             .cmdGetBlob.SetFocus
810             cmdGetBlob_Click  ' ** Procedure: Below.
820           Case Else
830             Beep
840           End Select
850         End With
860       Case vbKeyR
870         With Me
880           intRetVal = 0
890           Select Case .tabRelation
              Case .tabRelation_pagLocalRel.PageIndex
900             .cmdRestore.SetFocus
910             cmdRestore_Click  ' ** Procedure: Below.
920           Case .tabRelation_pagExternalRel.PageIndex
930             .cmdRestoreExternal.SetFocus
940             cmdRestoreExternal_Click  ' ** Procedure: Below.
950           Case Else
960             Beep
970           End Select
980         End With
990       Case vbKeyS
1000        With Me
1010          intRetVal = 0
1020          Select Case .tabRelation
              Case .tabRelation_pagLocalRel.PageIndex
1030            .cmdSaveLayout.SetFocus
1040            cmdSaveLayout_Click  ' ** Procedure: Below.
1050          Case Else
1060            Beep
1070          End Select
1080        End With
1090      Case vbKeyW
1100        With Me
1110          intRetVal = 0
1120          Select Case .tabRelation
              Case .tabRelation_pagAbout.PageIndex
1130            .cmdWeb.SetFocus
1140            cmdWeb_Click  ' ** Procedure: Below.
1150          Case Else
1160            Beep
1170          End Select
1180        End With
1190      Case vbKeyX
1200        With Me
1210          intRetVal = 0
1220          .cmdClose.SetFocus
1230          cmdClose_Click  ' ** Procedure: Below.
1240        End With
1250      End Select
1260    End If

EXITP:
1270    KeyCode = intRetVal
1280    Exit Sub

ERRH:
1290    intRetVal = 0
1300    Select Case ERR.Number
        Case Else
1310      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1320    End Select
1330    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1410    If blnDatabaseLinked = True Then
1420      DoCmd.DeleteObject acTable, "tblDatabase"
1430    End If

EXITP:
1440    Exit Sub

ERRH:
1450    Select Case ERR.Number
        Case Else
1460      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1470    End Select
1480    Resume EXITP

End Sub

Private Sub cmdClose_Click()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1510    DoCmd.Close acForm, THIS_NAME

EXITP:
1520    Exit Sub

ERRH:
1530    Select Case ERR.Number
        Case Else
1540      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1550    End Select
1560    Resume EXITP

End Sub

Private Sub tabRelation_Change()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "tabRelation_Change"

1610    With Me
1620      Select Case .tabRelation
          Case .tabRelation_pagLocalRel.PageIndex
1630        .relview_stored_lbx.Requery
1640      Case .tabRelation_pagExternalRel.PageIndex
            ' ** Nothing yet.
1650      Case .tabRelation_pagBlobs.PageIndex
            ' ** Nothing yet.
1660      Case .tabRelation_pagAbout.PageIndex
            ' ** Nothing yet.
1670      End Select
1680    End With

EXITP:
1690    Exit Sub

ERRH:
1700    Select Case ERR.Number
        Case Else
1710      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1720    End Select
1730    Resume EXITP

End Sub

Private Sub cmdRestore_Click()
' ** Restore the selected saved Relationship Window View.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRestore_Click"

        Dim lngX As Long
        Dim blnRetVal As Boolean

1810    blnRetVal = False

1820    With Me
1830      If .relview_stored_lbx.ListCount > 0 Then
1840        For lngX = 0 To (.relview_stored_lbx.ListCount - 1)
1850          If .relview_stored_lbx.Selected(lngX) = True Then
1860            blnRetVal = True
1870            Exit For
1880          End If
1890        Next
1900        If blnRetVal = False Then
1910          .relview_stored_lbx = .relview_stored_lbx.Column(0, 0)
1920          blnRetVal = True
1930          DoEvents
1940        End If
1950      Else
1960        blnRetVal = False
1970        Beep
1980      End If
1990      If blnRetVal = True Then
2000        blnRetVal = RestoreLayout(.relview_stored_lbx)  ' ** Function: Below.
2010        Beep
2020      End If
2030    End With

EXITP:
2040    Exit Sub

ERRH:
2050    Select Case ERR.Number
        Case Else
2060      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2070    End Select
2080    Resume EXITP

End Sub

Private Sub cmdDelete_Click()
' ** Delete the selected saved Relationship Window View.

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim msgResponse As VbMsgBoxResult

2110    With Me
2120      msgResponse = MsgBox("Are you sure you want to delete this View?" & vbCrLf & _
            "    " & .relview_stored_lbx.Value, vbQuestion + vbOKCancel, ("Delete Stored Relationship View" & Space(20)))
2130      If msgResponse = vbOK Then
            ' ** Delete tblRelation_View_Window, by specified [vewnam].
2140        Set dbs = CurrentDb
2150        With dbs
2160          Set qdf = .QueryDefs("qryRelationView_04")
2170          With qdf.Parameters
2180            ![vewnam] = Me.relview_stored_lbx
2190          End With
2200          qdf.Execute
2210          .Close
2220        End With
            ' ** Refresh the data on the Form.
2230        .Refresh
2240        strNewName = GetNewName  ' ** Function: Below.
2250        .relview_name_new.DefaultValue = Chr(34) & strNewName & Chr(34)
2260      End If
2270    End With

EXITP:
2280    Set qdf = Nothing
2290    Set dbs = Nothing
2300    Exit Sub

ERRH:
2310    Select Case ERR.Number
        Case Else
2320      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2330    End Select
2340    Resume EXITP

End Sub

Private Sub cmdSaveLayout_Click()
' ** Save the current coordinates of every Window in the Relationships Window.

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaveLayout_Click"

2410    SaveLayout     ' ** Procedure: Below.
2420    DoEvents

        ' ** Update our ListBox control.
2430    With Me
2440      .relview_stored_lbx.Requery
2450      DoEvents
2460      .relview_name_new = GetNewName  ' ** Function: Below.
2470    End With

EXITP:
2480    Exit Sub

ERRH:
2490    Select Case ERR.Number
        Case Else
2500      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2510    End Select
2520    Resume EXITP

End Sub

Private Sub relview_stored_lbx_DblClick(Cancel As Integer)
' ** Restore the selected saved Relationship Window View.

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "relview_stored_lbx_DblClick"

        Dim blnRetVal As Boolean

2610    With Me
2620      blnRetVal = RestoreLayout(.relview_stored_lbx)  ' ** Function: Below.
2630      Beep
2640    End With

EXITP:
2650    Exit Sub

ERRH:
2660    Select Case ERR.Number
        Case Else
2670      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2680    End Select
2690    Resume EXITP

End Sub

Private Sub cmdSelectMDB_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectMDB_Click"

        Dim strTmp01 As String

2710    With Me
          ' ** Grab all of the entries in the MsysObjects table with a TYPE  = -32758.
2720      strTmp01 = GetMDBRels(.ExternalMDB_lbx)  ' ** Function: Below.
          ' ** Set focus back to the first row (if any) of the "ExternalMDB_lbx" control.
2730      .ExternalMDB_lbx.Value = .ExternalMDB_lbx.ItemData(0)
2740    End With

EXITP:
2750    Exit Sub

ERRH:
2760    Select Case ERR.Number
        Case Else
2770      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2780    End Select
2790    Resume EXITP

End Sub

Private Sub cmdRestoreExternal_Click()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRestoreExternal_Click"

        Dim strTheUser As String, strTheMDB As String
        Dim intListIdx As Long
        Dim blnRetVal As Boolean

2810    With Me
          ' ** Verify the user has selected an entry from the ListBox.
2820      If .ExternalMDB_lbx.ListCount > 0 Then
2830        Select Case IsNull(.ExternalMDB_lbx.Column(0))
            Case True
2840          Beep
2850          MsgBox "You must select an entry from the ListBox.", vbInformation + vbOKOnly, "Nothing To Do"
2860        Case False
              ' ** Call Our GetBlob function.
              ' ** It will Decode the BLOB and create the necessary entries in our
              ' ** local Relationship View Layout table.
2870          strTheUser = .ExternalMDB_lbx.Column(0)
2880          strTheMDB = .ExternalMDB_lbx.Column(1)
2890          intListIdx = .ExternalMDB_lbx.ListIndex
2900          If Left(strTheUser, 1) = Chr(160) Or Left(strTheMDB, 1) = Chr(160) Then  ' ** Hard-Space, so this is line 2 of 2-line path.
2910            intListIdx = intListIdx - 1
2920            strTheUser = .ExternalMDB_lbx.Column(0, intListIdx)
2930            strTheMDB = .ExternalMDB_lbx.Column(1, intListIdx)
2940            .ExternalMDB_lbx = strTheUser  ' ** Move it back up 1 line.
2950            DoEvents
2960          End If
2970          GetBlob strTheUser, strTheMDB  ' ** Procedure: Below.
2980          blnRetVal = RestoreLayout(.ExternalMDB_lbx)  ' ** Function: Below.
              ' ** Update our ListBox control.
2990          .relview_stored_lbx.Requery
3000        End Select
3010      Else
3020        Beep
3030        MsgBox "There are no entries in the ListBox.", vbInformation + vbOKOnly, "Nothing To Do"
3040      End If
3050    End With

EXITP:
3060    Exit Sub

ERRH:
3070    Select Case ERR.Number
        Case Else
3080      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3090    End Select
3100    Resume EXITP

End Sub

Private Sub cmdGetBlob_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdGetBlob_Click"

        ' ** Call our GetBlob function with no Params so
        ' ** it will look to the local MSysObjects table
3210    GetBlob  ' ** Procedure: Below.

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3250    End Select
3260    Resume EXITP

End Sub

Private Sub cmdWeb_Click()
' ** Works fine!

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWeb_Click"

        ' ** Setup as a Hyperlinked control.
        'MsgBox "Not connected.", vbInformation + vbOKOnly, "Hyperlink Not Set Up"

EXITP:
3310    Exit Sub

ERRH:
3320    Select Case ERR.Number
        Case Else
3330      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3340    End Select
3350    Resume EXITP

End Sub

Private Sub SaveLayout()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "SaveLayout"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngWndMDI As Long, lngWndRel As Long, lngWndODsk As Long
        Dim typRECT As RECT
        Dim typPT As POINTAPI
        Dim lngWndTemp As Long
        Dim strPath As String, strPathFile As String
        Dim lngDbsID As Long, lngRelViewID As Long
        Dim blnContinue As Boolean
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngRetVal As Long

3410    blnContinue = True

3420    With Me

          ' ** Make sure they entered a unique name to store the Relationship View.
3430      If IsNull(.relview_name_new) = True Then
3440        blnContinue = False
3450        MsgBox "You must enter a name to identify this relationship view", vbInformation + vbOKOnly, "Entry Required"
3460        .relview_name_new.SetFocus
3470      Else
3480        If Trim(.relview_name_new) = vbNullString Then
3490          blnContinue = False
3500          MsgBox "You must enter a name to identify this relationship view", vbInformation + vbOKOnly, "Entry Required"
3510          .relview_name_new.SetFocus
3520        Else
3530          varTmp00 = DLookup("[relview_id]", "tblRelation_View", "[relview_name] = '" & .relview_name_new & "'")
3540          If IsNull(varTmp00) = False Then
3550            blnContinue = False
3560            MsgBox "The name entered is already present in the table." & vbCrLf & vbCrLf & _
                  "Enter a different name, or delete the previous saved view.", vbInformation + vbOKOnly, "Duplicate View Name"
3570          End If
3580        End If
3590      End If

3600      If blnContinue = True Then

3610        varTmp00 = Null
3620        Set dbs = CurrentDb
3630        With dbs

3640          strTmp01 = Parse_File(.Name)  ' ** Module Function: modFileUtilities.
3650          If Right(strTmp01, 1) = "e" Then strTmp01 = (Left(strTmp01, (Len(strTmp01) - 1)) & "b")
3660          If TableExistsX("tblDatabase") = False Then  ' ** Function: Below.
3670            strPath = Parse_Path(.Name)  ' ** Module Function: modFileUtilities.
3680            If InStr(strPath, gstrDir_Dev) > 0 Then
3690              If InStr(strPath, gstrDir_DevEmpty) > 0 Then
3700                If InStr(strPath, "_bak") > 0 Then
                      ' ** TrustAux.mdb should be in one directory up.
3710                  strPath = gstrDir_Dev & LNK_SEP & gstrDir_DevEmpty
3720                Else
                      ' ** TrustAux.mdb should be in this same directory.
3730                End If
3740              ElseIf InStr(strPath, gstrDir_DevDemo) > 0 Then
3750                If InStr(strPath, "NewDemo") > 0 Then
3760                  If InStr(strPath, "_bak") > 0 Then
                        ' ** TrustAux.mdb should be in one directory up.
3770                    strPath = gstrDir_Dev & LNK_SEP & gstrDir_DevDemo
3780                  Else
                        ' ** TrustAux.mdb should be in this same directory.
3790                  End If
3800                Else
                      ' ** Use TrustAux.mdb from the EmptyDatabase directory.
3810                  strPath = gstrDir_Dev & LNK_SEP & gstrDir_DevEmpty
3820                End If
3830              Else
3840                blnContinue = False
3850                Beep
3860                MsgBox "WHERE AM I?", vbExclamation + vbOKOnly, "File Not Found"
3870              End If
3880            Else
3890              If InStr(strPath, "Database") > 0 Then
                    ' ** TrustAux.mdb should be in this same directory.
3900              Else
                    ' ** TrustAux.mdb should be in the Database directory.
3910                strPath = strPath & LNK_SEP & "Database"
3920              End If
3930            End If
3940            If blnContinue = True Then
3950              strPathFile = strPath & LNK_SEP & "TrustAux.mdb"
3960              DoCmd.TransferDatabase acLink, "Microsoft Access", strPathFile, acTable, "tblDatabase", "tblDatabase"
3970              blnDatabaseLinked = True
3980              varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strTmp01 & "'")
3990            End If  ' ** blnContinue.
4000          Else
4010            varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strTmp01 & "'")
4020          End If

4030          .Close
4040        End With  ' ** dbs.
4050        Set dbs = Nothing

4060        Select Case IsNull(varTmp00)
            Case True
4070          blnContinue = False
4080          Beep
4090          MsgBox "[dbs_id] not found!", vbCritical + vbOKOnly, "Record Not Found"
4100        Case False
4110          lngDbsID = varTmp00
4120        End Select

4130      End If  ' ** blnContinue.

4140      If blnContinue = True Then
4150  On Error Resume Next
            ' ** Force the Relationships window to open.
4160        DoCmd.RunCommand acCmdRelationships
4170  On Error GoTo ERRH

            ' ** The Relationships Window is a child of the MDI Client window.
            ' ** Find MDIClient first.
4180        lngWndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)  ' ** API Function: modWindowFunctions.

            ' ** Find the Relationships Window.
4190        lngWndRel = FindWindowEx(lngWndMDI, 0&, "OSysRel", "Relationships")  ' ** API Function: modWindowFunctions.

4200        If lngWndRel = 0 Then
4210          blnContinue = False
4220          MsgBox "The Relationships Window is not open.", vbCritical + vbOKOnly, "Relationships Window Didn't Open"
4230        End If
4240      End If  ' ** blnContinue.

4250      If blnContinue = True Then
            ' ** The first child window is of class ODsk.
4260        lngWndODsk = FindWindowEx(lngWndRel, 0&, "ODsk", vbNullString)  ' ** API Function: modWindowFunctions.

            ' ** Let's get first Child Window of the ODsk window.
4270        lngWndTemp = GetWindow(lngWndODsk, GW_CHILD)  ' ** API Function: modWindowFunctions.
4280        If lngWndTemp = 0 Then
4290          blnContinue = False
4300          MsgBox "There are no Relationships!", vbCritical + vbOKOnly, "Nothing To Do"
4310        End If
4320      End If  ' ** blnContinue.

4330      If blnContinue = True Then

4340        Set dbs = CurrentDb
4350        With dbs
4360          Set rst = .OpenRecordset("tblRelation_View", dbOpenDynaset, dbConsistent)
4370          With rst
4380            .AddNew
4390            ![dbs_id] = lngDbsID
4400            ![relview_name] = Me.relview_name_new
4410            ![relview_user] = GetUserName  ' ** Module Function: modFileUtilities.
4420            ![relview_datemodified] = Now()
4430            .Update
4440            .Bookmark = .LastModified
4450            lngRelViewID = ![relview_id]
4460            .Close
4470          End With  ' ** rst.
4480          Set rst = Nothing
4490        End With  ' ** dbs.

4500        lngRetVal = GetWindowRect(lngWndTemp, typRECT)  ' ** API Function: modWindowFunctions.

            ' ** MUST store coordinates as Client Not Screen coords.
            ' ** Upper Left corner of window first.
4510        typPT.X = typRECT.Left
4520        typPT.Y = typRECT.Top
4530        lngRetVal = ScreenToClient(lngWndRel, typPT)  ' ** API Function: modWindowFunctions.
4540        typRECT.Left = typPT.X
4550        typRECT.Top = typPT.Y

            ' ** Lower right corner of window second.
4560        typPT.X = typRECT.Right
4570        typPT.Y = typRECT.Bottom
4580        lngRetVal = ScreenToClient(lngWndRel, typPT)  ' ** API Function: modWindowFunctions.
4590        typRECT.Right = typPT.X
4600        typRECT.Bottom = typPT.Y

            ' ** Get the Text for this Window.
4610        strTmp01 = Space(256)
4620        lngRetVal = GetWindowText(lngWndTemp, strTmp01, 256)  ' ** API Function: modWindowFunctions.
4630        strTmp01 = Left(strTmp01, lngRetVal)

            ' ** Update the fields.
4640        Set rst = dbs.OpenRecordset("tblRelation_View_Window", dbOpenDynaset, dbConsistent)
4650        With rst
4660          .AddNew
4670          ![dbs_id] = lngDbsID
4680          ![relview_id] = lngRelViewID
4690          ![relwin_window] = strTmp01
4700          ![relwin_X] = typRECT.Left
4710          ![relwin_X1] = typRECT.Right
4720          ![relwin_Y] = typRECT.Top
4730          ![relwin_Y1] = typRECT.Bottom
4740          ![relwin_user] = GetUserName  ' ** Module Function: modFileUtilities.
4750          ![relwin_datemodified] = Now()
4760          .Update
4770        End With  ' ** Leave rst open.

            ' ** Walk through every sibling window.
4780        Do

              ' ** Get the NEXT SIBLING Window.
4790          lngWndTemp = GetWindow(lngWndTemp, GW_HWNDNEXT)  ' ** API Function: modWindowFunctions.

              ' ** Check for a valid Window handle.
4800          If lngWndTemp <> 0 Then

4810            lngRetVal = GetWindowRect(lngWndTemp, typRECT)  ' ** API Function: modWindowFunctions.
                ' ** MUST store coords as Client Not Screen coords.
                ' ** Upper Left corner of window first.
4820            typPT.X = typRECT.Left
4830            typPT.Y = typRECT.Top
4840            lngRetVal = ScreenToClient(lngWndRel, typPT)  ' ** API Function: modWindowFunctions.
4850            typRECT.Left = typPT.X
4860            typRECT.Top = typPT.Y

                ' ** Lower right corner of window second.
4870            typPT.X = typRECT.Right
4880            typPT.Y = typRECT.Bottom
4890            lngRetVal = ScreenToClient(lngWndRel, typPT)  ' ** API Function: modWindowFunctions.
4900            typRECT.Right = typPT.X
4910            typRECT.Bottom = typPT.Y

                ' ** Get the Text for this Window.
4920            strTmp01 = Space(256)
4930            lngRetVal = GetWindowText(lngWndTemp, strTmp01, 256)  ' ** API Function: modWindowFunctions.
4940            strTmp01 = Left(strTmp01, lngRetVal)

                ' ** Update the fields.
4950            With rst
4960              .AddNew
4970              ![dbs_id] = lngDbsID
4980              ![relview_id] = lngRelViewID
4990              ![relwin_window] = strTmp01
5000              ![relwin_X] = typRECT.Left
5010              ![relwin_X1] = typRECT.Right
5020              ![relwin_Y] = typRECT.Top
5030              ![relwin_Y1] = typRECT.Bottom
5040              ![relwin_user] = GetUserName  ' ** Module Function: modFileUtilities.
5050              ![relwin_datemodified] = Now()
5060              .Update
5070            End With  ' ** rst.

5080          End If

              ' ** Start the process from the Top again.
              ' ** End this loop if no more Windows.
5090        Loop While lngWndTemp <> 0

            ' ** All done!
5100        rst.Close
5110        dbs.Close

5120      End If  ' ** blnContinue.

5130    End With

EXITP:
5140    Set rst = Nothing
5150    Set qdf = Nothing
5160    Set dbs = Nothing
5170    Exit Sub

ERRH:
5180    Select Case ERR.Number
        Case Else
5190      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
5200    End Select
5210    Resume EXITP

End Sub

Public Function RestoreLayout(ctl As Access.ListBox) As Boolean

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "RestoreLayout"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim colWindows As New VBA.Collection  ' ** Collection for Window Text and hWnd.
        Dim strSel As String
        Dim hWndMDI As Long, hWndRel As Long, hWndODsk As Long
        Dim hWndTemp As Long
        Dim lngRecs As Long
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long, lngX1 As Long, lngY1 As Long, lngZ As Long
        Dim lngRetVal As Long
        Dim blnRetVal As Boolean

5310    blnRetVal = True

        ' ** Ensure we are zero.
5320    hWndTemp = 0

5330  On Error Resume Next
        ' ** Force the Relationships window to open.
5340    DoCmd.RunCommand acCmdRelationships
5350  On Error GoTo ERRH

        ' ** Have they selected a previously saved Relationship View from our ListBox.
5360    If IsNull(ctl) = True Then
5370      blnRetVal = False
5380      Beep
5390      MsgBox "You must select an existing Relationship View.", vbInformation + vbOKOnly, "Entry Required"
5400    ElseIf ctl.ListCount = 0 Then
5410      Beep
5420      MsgBox "You must select an existing Relationship View.", vbInformation + vbOKOnly, "Entry Required"
5430    Else
5440      If Trim(ctl) = vbNullString Then
5450        blnRetVal = False
5460        MsgBox "You Must Select an Existing Relationship View.", vbInformation + vbOKOnly, "Entry Required"
5470      Else

            ' ** The Relationships Window is a child of the MDI Client window.
            ' ** Find MDIClient first.
5480        hWndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)  ' ** API Function: modWindowFunctions.

            ' ** Find the Relationships Window.
5490        hWndRel = FindWindowEx(hWndMDI, 0&, "OSysRel", "Relationships")  ' ** API Function: modWindowFunctions.

            ' ** Check for a valid Window handle.
5500        If hWndRel = 0 Then
5510          blnRetVal = False
5520          MsgBox "The Relationships Window is not open.", vbCritical + vbOKOnly, "Relationships Window Didn't Open"
5530        Else

              ' ** The first child window is of class ODsk.
5540          hWndODsk = FindWindowEx(hWndRel, 0&, "ODsk", vbNullString)  ' ** API Function: modWindowFunctions.

              ' ** Let's get first Child Window of the ODsk window.
5550          hWndTemp = GetWindow(hWndODsk, GW_CHILD)  ' ** API Function: modWindowFunctions.

5560          If hWndTemp = 0 Then
5570            blnRetVal = False
5580            MsgBox "Their are no Relationships!", vbCritical + vbOKOnly, "Nothing To Do"
5590          Else

5600            Set dbs = CurrentDb

                ' ** Create a recordset of the selected records.
                ' ** tblRelation_View_Window, by specified [vewnam].
5610            strSel = ctl
5620            Set qdf = dbs.QueryDefs("qryRelationView_03")
5630            With qdf.Parameters
5640              ![vewnam] = strSel
5650            End With
5660            Set rst = qdf.OpenRecordset

                ' ** Add this window to our collection.

                ' ** Grab the Windows Text.
5670            strTmp01 = Space(256)
5680            lngRetVal = GetWindowText(hWndTemp, strTmp01, 256)  ' ** API Function: modWindowFunctions.
5690            strTmp01 = Left(strTmp01, lngRetVal)
5700            colWindows.Add hWndTemp, strTmp01

                ' ** Loop through the rest of the sibling windows adding them to our collection.
5710            Do
5720              hWndTemp = GetWindow(hWndTemp, GW_HWNDNEXT)  ' ** API Function: modWindowFunctions.
5730              If hWndTemp = 0& Then
5740                Exit Do
5750              Else
                    ' ** Grab the Windows Text.
5760                strTmp01 = Space(256)
5770                lngRetVal = GetWindowText(hWndTemp, strTmp01, 256)  ' ** API Function: modWindowFunctions.
5780                strTmp01 = Left(strTmp01, lngRetVal)
5790                colWindows.Add hWndTemp, strTmp01
5800              End If
5810            Loop

                ' ** Walk through our windows collection and match Saved hWnd with current record.
5820            rst.MoveLast
5830            lngRecs = rst.RecordCount

                ' ** Move our recordset to the first row.
5840            rst.MoveFirst
                ' ** Loop through all of the rows.
5850            For lngZ = 1& To lngRecs
5860  On Error Resume Next
                  ' ** Get the Window handle by using the Window's Text as the Key.
5870              hWndTemp = colWindows.item(rst![relwin_window])
5880  On Error GoTo ERRH
5890              If hWndTemp <> 0 Then
                    ' ** Copy the window coords into our temp vars.
5900                With rst
5910                  lngX = ![relwin_X]
5920                  lngX1 = ![relwin_X1]
5930                  lngY = ![relwin_Y]
5940                  lngY1 = ![relwin_Y1]
5950                End With
                    ' ** Move and size the window to its original values.
5960                lngRetVal = SetWindowPos(hWndTemp, 0&, lngX, lngY, lngX1 - lngX, lngY1 - lngY, SWP_SHOWWINDOW)  ' ** API Function: modWindowFunctions.
5970              End If
                  ' ** Move to the next row.
5980              If lngZ < lngRecs Then rst.MoveNext
                  ' ** Reset.
5990              hWndTemp = 0
                  ' ** Keep looping!
6000            Next

                ' ** All done!
6010            rst.Close
6020            dbs.Close

                ' ** Return Success.

6030          End If
6040        End If
6050      End If
6060    End If

EXITP:
6070    Set rst = Nothing
6080    Set qdf = Nothing
6090    Set dbs = Nothing
        ' ** Delete our Collection object.
        ' ** Remove the first object each time through the loop
        ' ** until there are no objects left in the collection.
6100    For lngX = 1 To colWindows.Count
6110      colWindows.Remove 1
6120    Next
6130    Set colWindows = Nothing
6140    RestoreLayout = blnRetVal
6150    Exit Function

ERRH:
6160    blnRetVal = False
6170    Select Case ERR.Number
        Case Else
6180      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
6190    End Select
6200    Resume EXITP

End Function

Private Function GetMDBRels(ctl As Access.ListBox) As String
' ** Pops File Dialog window.
' ** Returns selected MDB.
' ** Fills control with any Relationship Window Layout BLOBS User NAME.

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "GetMDBRels"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strSQL As String, strMDBname As String
        Dim intType As Integer
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim intX As Integer
        Dim strRetVal As String

6310    strRetVal = vbNullString

        ' ** Pop the File Dialog Window.
6320    strMDBname = fFileDialog  ' ** Function: Below
6330    If Len(strMDBname & vbNullString) <> 0 Then

          ' ** Relationship Window Layout BLOB TYPE = -32758.
6340      intType = -32758
6350      strSQL = "SELECT MSysObjects.* FROM MSysObjects WHERE Type = " & CStr(intType)
6360      Set dbs = OpenDatabase(strMDBname, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
6370      Set rst = dbs.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

          ' ** Fill the Rowsource of the ListBox control.
          ' ** Build our string.
          ' ** Move to first row.
6380      rst.MoveFirst
6390      strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
6400      Do While Not rst.EOF
6410        strTmp01 = strTmp01 & rst![Name] & ";"
6420        strTmp02 = strMDBname
6430        strTmp03 = vbNullString
6440        intLen = Len(strTmp02)
6450        If intLen > 55 Then
              ' ** Divide it into 2 lines.
6460          For intX = 55 To 1 Step -1
6470            If Mid(strTmp02, intX, 1) = LNK_SEP Then
6480              strTmp03 = Mid(strTmp02, intX)
6490              strTmp02 = Left(strTmp02, (intX - 1))
6500              Exit For
6510            End If
6520          Next
6530        End If
6540        strTmp01 = strTmp01 & strTmp02 & ";"
6550        If strTmp03 <> vbNullString Then
6560          strTmp01 = strTmp01 & String(3, Chr(160)) & rst![Name] & ";" & String(3, Chr(160)) & strTmp03 & ";"
6570        End If
6580        rst.MoveNext
6590      Loop

6600      ctl.RowSource = strTmp01

          ' ** Return Success
6610      strRetVal = strMDBname

          ' ** Close our Recordset.
6620      rst.Close
          ' ** Release dbs.
6630      dbs.Close

6640    End If

EXITP:
6650    Set rst = Nothing
6660    Set dbs = Nothing
6670    GetMDBRels = strRetVal
6680    Exit Function

ERRH:
6690    strRetVal = RET_ERR
6700    Select Case ERR.Number
        Case Else
6710      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
6720    End Select
6730    Resume EXITP

End Function

Private Sub GetBlob(Optional strTheUser As String = vbNullString, Optional strTheMDB As String = vbNullString)
' ** Supply params if using External MDB.
' ** The MDB must include full path info.

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "GetBlob"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim typRB As RELBLOB
        Dim arr_typRW() As RELWINDOW
        Dim arr_bytA() As Byte
        Dim strSQL As String, strSel As String
        Dim lngDbsID As Long, lngRelViewID As Long
        Dim strTmp01 As String, lngTmp02 As Long
        Dim lngX As Long

        ' ** Read the Relationship window BLOB into our array.
        ' ** Assumes CurrentUser is the same user who setup and saved the current Relationship window.
        ' ** Layout for the internal tables. For an External MDB we supply the User!
6810    If Len(strTheUser & vbNullString) > 0 Then
6820      strSel = strTheUser
6830    Else
6840      strSel = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
6850    End If

6860    If Len(strTheMDB & vbNullString) > 0 Then
6870      Set dbs = OpenDatabase(strTheMDB, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
6880    Else
6890      Set dbs = CurrentDb()
6900    End If

6910    strSQL = "SELECT MSysObjects.* FROM MSysObjects WHERE [Name] = '" & strSel & "'"
6920    Set rst = dbs.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

        ' ** Get length of BLOB.
6930    lngTmp02 = rst![LVExtra].FieldSize()

6940    ReDim arr_bytA(0 To lngTmp02)
        ' ** Copy Blob to our array.
6950    arr_bytA = rst![LVExtra].GetChunk(0, lngTmp02)

        ' ** Close the Recordset.
6960    rst.Close
        ' ** Release the database.
6970    dbs.Close

        ' ** Fill in our RELBLOB header.
6980    CopyMemory typRB, arr_bytA(0), Len(typRB)  ' ** API Function: modWindowFunctions.

        ' ** Fill in our TextBox controls.
6990    With Me
7000      .txtAAAs = typRB.AAAs
7010      .txtBlank = typRB.Blank
7020      .txtFFFs = typRB.FFFs
7030      .txtNumWindows = typRB.NumWindows
7040      .txtPadding = typRB.Pad
7050      .txtSig = typRB.Sig
7060      .txtRelWinX1 = typRB.RelWinX1
7070      .txtRelWinX2 = typRB.RelWinX2
7080      .txtRelWinY1 = typRB.RelWinY1
7090      .txtRelWinY2 = typRB.RelWinY2
7100      .txtClientRectY = typRB.ClientRectY
7110      .txtClientRectX = typRB.ClientRectX
7120    End With

        ' ** First 68 Bytes are the Header.
        ' ** This is followed by (NumWindows + 1) * 284 bytes per record'
        ' ** Last record seems to be padding.
        ' ** Let's create an array of our RelWin structures.
7130    ReDim arr_typRW(0 To typRB.NumWindows)
        ' ** Fill in our array of structures.
7140    For lngX = 0 To typRB.NumWindows
7150      CopyMemory arr_typRW(lngX), arr_bytA((lngX * 284) + 68), 284  ' ** API Function: modWindowFunctions.
7160    Next lngX

7170    Set dbs = CurrentDb
7180    With dbs
7190      strTmp01 = Parse_File(.Name)  ' ** Module Function: modFileUtilities.
7200      If Right(strTmp01, 1) = "e" Then strTmp01 = (Left(strTmp01, (Len(strTmp01) - 1)) & "b")
7210      lngDbsID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strTmp01 & "'")
          ' ** Delete tblRelation_View_Window, by specified [dbid], [vewnam].
7220      Set qdf = .QueryDefs("qryRelationView_07")
7230      With qdf.Parameters
7240        ![dbid] = lngDbsID
7250        ![vewnam] = strSel
7260      End With
7270      qdf.Execute
7280      Set rst = .OpenRecordset("tblRelation_View", dbOpenDynaset, dbConsistent)
7290      With rst
7300        .AddNew
7310        ![dbs_id] = lngDbsID
7320        ![relview_name] = strSel
7330        ![relview_user] = GetUserName  ' ** Module Function: modFileUtilities.
7340        ![relview_datemodified] = Now()
7350        .Update
7360        .Bookmark = .LastModified
7370        lngRelViewID = ![relview_id]
7380        .Close
7390      End With
7400    End With  ' ** Leave dbs open.

        ' ** Now Let's display our Windows and their coords
        ' ** by adding them to our table.
7410    With dbs
7420      Set rst = .OpenRecordset("tblRelation_View_Window", dbOpenDynaset, dbConsistent)
7430      With rst
7440        For lngX = 0 To typRB.NumWindows - 1
              ' ** Add a new row to our table.
7450          .AddNew
7460          ![dbs_id] = lngDbsID
7470          ![relview_id] = lngRelViewID
7480          ![relwin_window] = StrConv(arr_typRW(lngX).WinName, vbFromUnicode)
7490          ![relwin_X] = arr_typRW(lngX).RelWinX1
7500          ![relwin_X1] = arr_typRW(lngX).RelWinX2
7510          ![relwin_Y] = arr_typRW(lngX).RelWinY1
7520          ![relwin_Y1] = arr_typRW(lngX).RelWinY2
7530          ![relwin_user] = GetUserName  ' ** Module Function: modFileUtilities.
7540          ![relwin_datemodified] = Now()
7550          .Update
7560        Next
7570        .Close
7580      End With
7590      .Close
7600    End With  ' ** dbs.

7610    With Me
7620      .relview_stored_lbx.Requery
7630    End With

EXITP:
7640    Set rst = Nothing
7650    Set qdf = Nothing
7660    Set dbs = Nothing
7670    Exit Sub

ERRH:
7680    Select Case ERR.Number
        Case Else
7690      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
7700    End Select
7710    Resume EXITP

End Sub

Private Function fFileDialog() As String
' ** Calls the API File Dialog Window.
' ** Returns full path to the existing File.

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "fFileDialog"

        Dim clsDialog As Object
        Dim strRetVal As String

7810    strRetVal = vbNullString

        ' ** Call the File Common Dialog Window.
7820    Set clsDialog = New clsCommonDialog

        ' ** Fill in our properties.
        'clsDialog.Filter = "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)
7830    clsDialog.Filter = "MDB (*.MDB)" & Chr$(0) & "*.MDB" & Chr$(0)
7840    clsDialog.Filter = clsDialog.Filter & "MDE (*.MDE)" & Chr$(0) & "*.MDE" & Chr$(0)

7850    clsDialog.hDC = 0
7860    clsDialog.MaxFileSize = 256
7870    clsDialog.Max = 256
7880    clsDialog.FileTitle = vbNullString
7890    clsDialog.DialogTitle = "Please Select an Existing Access Database File"
7900    clsDialog.InitDir = vbNullString
7910    clsDialog.DefaultExt = vbNullString

        ' ** Display the File Dialog.
7920    clsDialog.ShowOpen

        ' ** See if user clicked Cancel or even selected
        ' ** the very same file already selected.
7930    strRetVal = clsDialog.Filename
7940    If Len(strRetVal & vbNullString) = 0 Then
          ' ** Raise the exception.
7950      ERR.Raise vbObjectError + 513, "Form_frmRelationshipViews.fFileDialog", _
            "Please Select an Existing Access Database"
7960    End If

EXITP:
7970    Set clsDialog = Nothing
7980    fFileDialog = strRetVal
7990    Exit Function

ERRH:
8000    strRetVal = vbNullString
8010    Select Case ERR.Number
        Case Else
8020      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
8030    End Select
8040    Resume EXITP

End Function

Private Function GetNewName() As String

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "GetNewName"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRetVal As String

8110    strRetVal = vbNullString

8120    Set dbs = CurrentDb
8130    With dbs
          ' ** qryRelationView_05 (tblRelation_View, with HasSfx, relview_nam, relview_num),
          ' ** grouped, with cnt, HasSfx, relview_num.
8140      Set qdf = .QueryDefs("qryRelationView_06")
8150      Set rst = qdf.OpenRecordset
8160      With rst
8170        If .BOF = True And .EOF = True Then
              ' ** No saved layouts.
8180          strRetVal = "Trust_Accountant_01"
8190        Else
8200          .MoveLast
8210          If .RecordCount = 1 Then
8220            strRetVal = ![relview_nam] & "_" & Right("00" & CStr(![relview_num] + 1), 2)
8230          Else
                ' ** Inconsistent naming.
8240          End If
8250        End If
8260        .Close
8270      End With
8280      .Close
8290    End With

8300    If strRetVal = vbNullString Then
8310      strRetVal = "RelationshipView_01"
8320    End If

EXITP:
8330    Set rst = Nothing
8340    Set qdf = Nothing
8350    Set dbs = Nothing
8360    GetNewName = strRetVal
8370    Exit Function

ERRH:
8380    strRetVal = vbNullString
8390    Select Case ERR.Number
        Case Else
8400      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
8410    End Select
8420    Resume EXITP

End Function

Private Function TableExistsX(strTableName As String) As Boolean

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "TableExistsX"

        Dim dbs As DAO.Database, tdf As DAO.TableDef
        Dim blnRetVal As Boolean

8510    blnRetVal = False

8520    If Trim(strTableName) <> vbNullString Then
8530      Set dbs = CurrentDb
8540      With dbs
8550        For Each tdf In .TableDefs
8560          With tdf
8570            If .Name = strTableName Then
8580              blnRetVal = True
8590              Exit For
8600            End If
8610          End With
8620        Next
8630        .Close
8640      End With
8650    End If

EXITP:
8660    Set tdf = Nothing
8670    Set dbs = Nothing
8680    TableExistsX = blnRetVal
8690    Exit Function

ERRH:
8700    blnRetVal = False
8710    Select Case ERR.Number
        Case Else
8720      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
8730    End Select
8740    Resume EXITP

End Function
