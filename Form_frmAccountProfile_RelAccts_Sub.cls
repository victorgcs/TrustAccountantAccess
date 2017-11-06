VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountProfile_RelAccts_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountProfile_RelAccts_Sub"

'VGC 08/29/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Accounts:         A {frmAccountProfile_RelAccts_Sub on frmAccountProfile_RelAccts}
' **   Sort By Number:   B {opgOrderBy_optNumber on frmAccountProfile_RelAccts}
' **   Cancel:           C {cmdCancel on frmAccountProfile_RelAccts}
' **   Clear Closed:     E {cmdClearClosed on frmAccountProfile_RelAccts}
' **   Sort By Name:     M {opgOrderBy_optName on frmAccountProfile_RelAccts}
' **   Select None:      N {cmdSelectNone on frmAccountProfile_RelAccts}
' **   OK:               O {cmdOK on frmAccountProfile_RelAccts}
' **   Print Report:     P {cmdPrintReport on frmAccountProfile_RelAccts}
' **   Reset:            R {cmdSelectReset on frmAccountProfile_RelAccts}
' **   Sort By Selected: S {opgOrderBy_optSelected on frmAccountProfile_RelAccts}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

'IsHid.ForeColor = 6974207
'Asc("Û") = 219

' ** Array: arr_varRA().
Private lngRAs As Long, arr_varRA() As Variant    ' ** Complete, full list.
Private lngCRAs As Long, arr_varCRA As Variant  ' ** Closed accounts still present.
Private lngORAs As Long, arr_varORA As Variant  ' ** Original list on entry.
Private Const RA_ELEMS As Integer = 3  ' ** Array's first-element UBound()
Private Const RA_ACCT     As Integer = 0
Private Const RA_FOUND    As Integer = 1
Private Const RA_CLOSED   As Integer = 2
Private Const RA_LST_ELEM As Integer = 3

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountProfile_RelAccts" Then

160         lngRAs = 0&
170         ReDim arr_varRA(RA_ELEMS, 0)
180         lngCRAs = 0&
190         ReDim arr_varCRA(RA_ELEMS, 0)
200         lngORAs = 0&
210         ReDim arr_varORA(RA_ELEMS, 0)

220       Else
230         Cancel = -1
240       End If
250     End With

260     If Cancel = -1 Then
270       Beep
280       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
290     End If

EXITP:
300     Exit Sub

ERRH:
310     Select Case ERR.Number
        Case Else
320       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
330     End Select
340     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

410     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
420     intShiftDown = (Shift And acShiftMask) > 0
430     intAltDown = (Shift And acAltMask) > 0
440     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
450     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyUp
470         intRetVal = 0
480         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
490       Case vbKeyDown
500         intRetVal = 0
510         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
520       Case vbKeyEscape
530         With Me
540           intRetVal = 0
550           .Parent.cmdCancel_Click  ' ** Form Procedure: frmAccountAssets.
560         End With
570       End Select
580     End If

        ' ** Ctrl keys.
590     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
600       Select Case intRetVal
          Case vbKeyS
610         intRetVal = 0
620         cmdSave_Click  ' ** Procedure: Below.
630       Case vbKeyTab, vbKeyReturn
640         With Me
650           intRetVal = 0
660           DoCmd.SelectObject acForm, .Parent.Name, False
670           If .Parent.cmdPrintReport.Enabled = True Then
680             .Parent.cmdPrintReport.SetFocus
690           Else
700             .Parent.cmdCancel.SetFocus
710           End If
720         End With
730       End Select
740     End If

        ' ** Ctrl-Shift keys.
750     If intCtrlDown And (Not intAltDown) And intShiftDown Then
760       Select Case intRetVal
          Case vbKeyF
770         With Me
780           intRetVal = 0
790           DoCmd.SelectObject acForm, .Parent.Name, False
800           .Parent.FocusHolder.SetFocus
810         End With
820       Case vbKeyTab, vbKeyReturn
830         With Me
840           intRetVal = 0
850           DoCmd.SelectObject acForm, .Parent.Name, False
860           .Parent.cmdCancel.SetFocus
870         End With
880       End Select
890     End If

EXITP:
900     KeyCode = intRetVal
910     Exit Sub

ERRH:
920     intRetVal = 0
930     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
940     Case Else
950       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
960     End Select
970     Resume EXITP

End Sub

Public Sub cmdSave_Click()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1010    DoCmd.RunCommand acCmdSaveRecord
1020    DoEvents

EXITP:
1030    Exit Sub

ERRH:
1040    Select Case ERR.Number
        Case Else
1050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1060    End Select
1070    Resume EXITP

End Sub

Private Sub IsRelated_AfterUpdate()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "IsRelated_AfterUpdate"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim strRelAccts As String
        Dim lngX As Long, lngE As Long

1110    With Me

1120      cmdSave_Click  ' ** Procedure: Above.

1130      strRelAccts = vbNullString
1140      lngRAs = 0&
1150      ReDim arr_varRA(RA_ELEMS, 0)

1160      Set rst = .RecordsetClone
1170      With rst
1180        .MoveLast
1190        lngRecs = .RecordCount
1200        .MoveFirst
1210        For lngX = 1& To lngRecs
1220          If ![IsRelated] = True Then
1230            lngRAs = lngRAs + 1&
1240            lngE = lngRAs - 1&
1250            ReDim Preserve arr_varRA(RA_ELEMS, lngE)
1260            arr_varRA(RA_ACCT, lngE) = ![accountno]
1270            arr_varRA(RA_FOUND, lngE) = CBool(True)
1280            arr_varRA(RA_CLOSED, lngE) = CBool(False)
1290            arr_varRA(RA_LST_ELEM, lngE) = ![relacct_id]
1300            strRelAccts = strRelAccts & ![accountno] & ","
1310          End If
1320          If lngX < lngRecs Then .MoveNext
1330        Next  ' ** lngX.
1340        If Trim(strRelAccts) <> vbNullString And Trim(strRelAccts) <> "," Then
1350          strRelAccts = Left(strRelAccts, (Len(strRelAccts) - 1))
1360        End If
1370        .Close
1380      End With  ' ** rst.
1390      Set rst = Nothing

1400      If lngCRAs > 0& Then
            ' ** Add the closed accounts.
1410        For lngX = 0& To (lngCRAs - 1&)
1420          lngRAs = lngRAs + 1&
1430          lngE = lngRAs - 1&
1440          ReDim Preserve arr_varRA(RA_ELEMS, lngE)
1450          arr_varRA(RA_ACCT, lngE) = arr_varCRA(RA_ACCT, lngX)
1460          arr_varRA(RA_FOUND, lngE) = arr_varCRA(RA_FOUND, lngX)
1470          arr_varRA(RA_CLOSED, lngE) = arr_varCRA(RA_CLOSED, lngX)
1480          arr_varRA(RA_LST_ELEM, lngE) = arr_varCRA(RA_LST_ELEM, lngX)
1490          strRelAccts = strRelAccts & "," & arr_varCRA(RA_ACCT, lngX)
1500        Next
1510      End If
1520      .Parent.related_accountno = strRelAccts

1530    End With

EXITP:
1540    Set rst = Nothing
1550    Exit Sub

ERRH:
1560    Select Case ERR.Number
        Case Else
1570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1580    End Select
1590    Resume EXITP

End Sub

Public Sub HighlightRelated(arr_varRel As Variant, lngRels As Long, Optional varReset As Variant, Optional varIsOpen As Variant)
' ** Highlight the related accounts.

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "HighlightRelated"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strQry As String
        Dim lngRecs As Long
        Dim blnReset As Boolean, blnIsOpen As Boolean
        Dim lngX As Long, lngE As Long

1610    With Me

1620      .FocusHolder.SetFocus

1630      Select Case IsMissing(varReset)
          Case True
1640        blnReset = False
1650      Case False
1660        blnReset = varReset
1670      End Select

1680      Select Case IsMissing(varIsOpen)
          Case True
1690        blnIsOpen = False
1700      Case False
1710        blnIsOpen = varIsOpen
1720      End Select

1730      Select Case blnReset
          Case True

1740        lngORAs = 0&
1750        ReDim arr_varORA(RA_ELEMS, 0)

1760        For lngX = 0& To (lngRels - 1&)
1770          lngORAs = lngORAs + 1&
1780          lngE = lngORAs - 1&
1790          ReDim Preserve arr_varORA(RA_ELEMS, lngE)
              ' **************************************************
              ' ** Array: arr_varORA()
              ' **
              ' **   Field  Element  Name          Constant
              ' **   =====  =======  ============  =============
              ' **     1       0     accountno     RA_ACCT
              ' **     2       1     Found         RA_FOUND
              ' **     3       2     Closed        RA_CLOSED
              ' **     4       3     relacct_id    RA_LST_ELEM
              ' **
              ' **************************************************
1800          arr_varORA(RA_ACCT, lngE) = arr_varRel(RA_ACCT, lngX)
1810          arr_varORA(RA_FOUND, lngE) = arr_varRel(RA_FOUND, lngX)
1820          arr_varORA(RA_CLOSED, lngE) = arr_varRel(RA_CLOSED, lngX)
1830          arr_varORA(RA_LST_ELEM, lngE) = arr_varRel(RA_LST_ELEM, lngX)
1840        Next

1850        strQry = .RecordSource
1860        .IsRelated.Visible = False
1870        .IsRelated_box1.ForeColor = MY_CLR_BGE
1880        .IsRelated_box2.ForeColor = MY_CLR_BGE
1890        .Desc.ForeColor = CLR_WHT
1900        Set dbs = CurrentDb
1910        With dbs
1920          Set qdf = .QueryDefs(strQry)

1930          Set rst = qdf.OpenRecordset
1940          With rst
1950            .MoveLast
1960            lngRecs = .RecordCount
1970            .MoveFirst
1980            For lngX = 1& To lngRecs
1990              .Edit
2000              ![IsRelated] = False
2010              ![HasRelated] = False
2020              .Update
2030              If lngX < lngRecs Then .MoveNext
2040            Next
2050            .MoveFirst
2060            For lngX = 0& To (lngORAs - 1)
2070              .FindFirst "[accountno] = '" & arr_varORA(RA_ACCT, lngX) & "'"
2080              If .NoMatch = False Then
2090                arr_varORA(RA_LST_ELEM, lngX) = ![relacct_id]
2100                .Edit
2110                ![IsRelated] = True
2120                ![HasRelated] = True
2130                .Update
2140                arr_varORA(RA_FOUND, lngX) = CBool(True)
2150              End If
2160            Next
2170            .Close
2180          End With
2190          Set rst = Nothing
2200          .Close
2210        End With
2220        .IsRelated.Visible = True
2230        .IsRelated_box1.ForeColor = CLR_IEE_RED
2240        .IsRelated_box2.ForeColor = CLR_IEE_RED
2250        .Desc.ForeColor = CLR_BLK

2260      Case False

2270        lngRAs = 0&
2280        ReDim arr_varRA(RA_ELEMS, 0)

2290        For lngX = 0& To (lngRels - 1&)
2300          lngRAs = lngRAs + 1&
2310          lngE = lngRAs - 1&
2320          ReDim Preserve arr_varRA(RA_ELEMS, lngE)
              ' **************************************************
              ' ** Array: arr_varRA()
              ' **
              ' **   Field  Element  Name          Constant
              ' **   =====  =======  ============  =============
              ' **     1       0     accountno     RA_ACCT
              ' **     2       1     Found         RA_FOUND
              ' **     3       2     Closed        RA_CLOSED
              ' **     4       3     relacct_id    RA_LST_ELEM
              ' **
              ' **************************************************
2330          arr_varRA(RA_ACCT, lngE) = arr_varRel(RA_ACCT, lngX)
2340          arr_varRA(RA_FOUND, lngE) = arr_varRel(RA_FOUND, lngX)
2350          arr_varRA(RA_CLOSED, lngE) = arr_varRel(RA_CLOSED, lngX)
2360          arr_varRA(RA_LST_ELEM, lngE) = arr_varRel(RA_LST_ELEM, lngX)
2370        Next

2380        strQry = .RecordSource
2390        .IsRelated.Visible = False
2400        .IsRelated_box1.ForeColor = MY_CLR_BGE
2410        .IsRelated_box2.ForeColor = MY_CLR_BGE
2420        If blnIsOpen = False Then
2430          .Desc.ForeColor = CLR_WHT
2440        End If
2450        Set dbs = CurrentDb
2460        With dbs
2470          Set qdf = .QueryDefs(strQry)

2480          Set rst = qdf.OpenRecordset
2490          With rst
2500            .MoveLast
2510            lngRecs = .RecordCount
2520            .MoveFirst
2530            For lngX = 1& To lngRecs
2540              .Edit
2550              ![IsRelated] = False
2560              ![HasRelated] = False
2570              .Update
2580              If lngX < lngRecs Then .MoveNext
2590            Next
2600            .MoveFirst
2610            For lngX = 0& To (lngRAs - 1)
2620              .FindFirst "[accountno] = '" & arr_varRA(RA_ACCT, lngX) & "'"
2630              If .NoMatch = False Then
2640                arr_varRA(RA_LST_ELEM, lngX) = ![relacct_id]
2650                .Edit
2660                ![IsRelated] = True
2670                ![HasRelated] = True
2680                .Update
2690                arr_varRA(RA_FOUND, lngX) = CBool(True)
2700              End If
2710            Next
2720            .Close
2730          End With
2740          Set rst = Nothing
2750          .Close
2760        End With
2770        .IsRelated.Visible = True
2780        If blnIsOpen = False Then
2790          .IsRelated_box1.ForeColor = CLR_IEE_RED
2800          .IsRelated_box2.ForeColor = CLR_IEE_RED
2810          .Desc.ForeColor = CLR_BLK
2820        End If

2830      End Select

2840      .Requery

2850    End With

EXITP:
2860    Set rst = Nothing
2870    Set qdf = Nothing
2880    Set dbs = Nothing
2890    Exit Sub

ERRH:
2900    Select Case ERR.Number
        Case Else
2910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2920    End Select
2930    Resume EXITP

End Sub

Public Function RelArray_Get() As Variant

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "RelArray_Get"

        Dim arr_varRetVal As Variant

3010    arr_varRetVal = arr_varRA

EXITP:
3020    RelArray_Get = arr_varRetVal
3030    Exit Function

ERRH:
3040    arr_varRetVal = RET_ERR
3050    Select Case ERR.Number
        Case Else
3060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3070    End Select
3080    Resume EXITP

End Function

Public Sub RelArray_Set(arr_varRel As Variant, lngRels As Long)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "RelArray_Set"

        Dim lngX As Long, lngE As Long

3110    lngRAs = 0&
3120    ReDim arr_varRA(RA_ELEMS, 0)

3130    For lngX = 0& To (lngRels - 1&)
3140      lngRAs = lngRAs + 1&
3150      lngE = lngRAs - 1&
3160      ReDim Preserve arr_varRA(RA_ELEMS, lngE)
          ' **************************************************
          ' ** Array: arr_varRA()
          ' **
          ' **   Field  Element  Name          Constant
          ' **   =====  =======  ============  =============
          ' **     1       0     accountno     RA_ACCT
          ' **     2       1     Found         RA_FOUND
          ' **     3       2     Closed        RA_CLOSED
          ' **     4       3     relacct_id    RA_LST_ELEM
          ' **
          ' **************************************************
3170      arr_varRA(RA_ACCT, lngE) = arr_varRel(RA_ACCT, lngX)
3180      arr_varRA(RA_FOUND, lngE) = arr_varRel(RA_FOUND, lngX)
3190      arr_varRA(RA_CLOSED, lngE) = arr_varRel(RA_CLOSED, lngX)
3200      arr_varRA(RA_LST_ELEM, lngE) = arr_varRel(RA_LST_ELEM, lngX)
3210    Next

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Public Sub CRelArray_Set(arr_varCRel As Variant, lngClosed As Long)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "CRelArray_Set"

3310    arr_varCRA = arr_varCRel
3320    lngCRAs = lngClosed

EXITP:
3330    Exit Sub

ERRH:
3340    Select Case ERR.Number
        Case Else
3350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3360    End Select
3370    Resume EXITP

End Sub

Public Sub ORelArray_Set(arr_varORel As Variant, lngOrigs As Long)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "ORelArray_Set"

3410    arr_varORA = arr_varORel
3420    lngORAs = lngOrigs

EXITP:
3430    Exit Sub

ERRH:
3440    Select Case ERR.Number
        Case Else
3450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3460    End Select
3470    Resume EXITP

End Sub

Public Function RecCnt() As Long

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

3510    lngRetVal = 0&

3520    With Me
3530      Set rst = .RecordsetClone
3540      With rst
3550        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
3560        Else
3570          .MoveLast
3580          lngRetVal = .RecordCount
3590        End If
3600        .Close
3610      End With
3620    End With

EXITP:
3630    Set rst = Nothing
3640    RecCnt = lngRetVal
3650    Exit Function

ERRH:
3660    lngRetVal = 0&
3670    Select Case ERR.Number
        Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

3810    Select Case intWhere
        Case acCmdRecordsGoToFirst
3820      DoCmd.RunCommand acCmdRecordsGoToFirst
3830    Case acCmdRecordsGoToLast
3840      DoCmd.RunCommand acCmdRecordsGoToLast
3850    Case acCmdRecordsGoToPrevious
3860      DoCmd.RunCommand acCmdRecordsGoToPrevious
3870    Case acCmdRecordsGoToNext
3880      DoCmd.RunCommand acCmdRecordsGoToNext
3890    Case Else
3900      If IsMissing(varID) = False Then
3910        With Me
3920          Set rst = .RecordsetClone
3930          With rst
3940            .FindFirst "[accountno] = '" & varID & "'"
3950            If .NoMatch = False Then
3960              Me.Bookmark = .Bookmark
3970            End If
3980            .Close
3990          End With
4000        End With
4010      End If
4020    End Select

EXITP:
4030    Set rst = Nothing
4040    Exit Sub

ERRH:
4050    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4060    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4070    Case Else
4080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4090    End Select
4100    Resume EXITP

End Sub
