Attribute VB_Name = "modJrnlCol_Sort"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modJrnlCol_Sort"

'VGC 03/23/2017: CHANGES!

'tblJOURNAL_FIELD DOESN'T GET AUTOMATICALLY UPDATED!!!

Private Const strSortOrig As String = "[SpecSort2], [JrnlCol_ID]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngTpp As Long
Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
' **

Public Sub JC_Sort_Now(strProc As String, frmSub As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Sort_Now"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [SpecSort2], [JrnlCol_ID]"
        'Const strStdDesc As String = ", [SpecSort2] DESC, [JrnlCol_ID] DESC"

110     With frmSub

120       If lngSortLbl_Width = 0& Then
130         lngSortLbl_Width = .Sort_lbl.Width
140       End If
150       If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
170       End If

180       .Controls(strSortLbl).Visible = False
190       .Controls(strSortLine).Visible = False
200       .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.

210       If strProc = "Form_Load" Then
220         strCalled = "JrnlCol_ID"
230         strSortNow = strSortOrig
240         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
250         lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (2& * lngTpp))
260         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
270         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
280         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
290         .Controls(strSortLbl).Top = lngSortLbl_Top
300         .Controls(strSortLbl).Left = lngSortLbl_Left
310         .Controls(strSortLine).Top = lngSortLine_Top
320         .Controls(strSortLine).Left = lngSortLine_Left
330         .Controls(strSortLine).Width = lngSortLine_Width
340         .Controls(strSortLbl).Caption = strArwUp
350         .Controls(strSortLbl).ForeColor = CLR_DKBLU
360       Else
370         strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
380         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
390         lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
400         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
410         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
420         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
430         .Controls(strSortLbl).Top = lngSortLbl_Top
440         .Controls(strSortLbl).Left = lngSortLbl_Left
450         .Controls(strSortLine).Top = lngSortLine_Top
460         .Controls(strSortLine).Left = lngSortLine_Left
470         .Controls(strSortLine).Width = lngSortLine_Width
480         If strCalled = "JrnlCol_ID" Then
490           strSortAsc = strSortOrig
500           If strSortNow = strSortAsc Then
510             strSortNow = "[SpecSort2] DESC, [JrnlCol_ID] DESC"
520             lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
530             lngSortLine_Top = (lngSortLine_Top - lngTpp)
540             lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
550             .Controls(strSortLbl).Top = lngSortLbl_Top
560             .Controls(strSortLbl).Left = lngSortLbl_Left
570             .Controls(strSortLine).Top = lngSortLine_Top
580             .Controls(strSortLbl).Caption = strArwDn
590             .Controls(strSortLbl).ForeColor = CLR_DKRED
600           Else
610             strSortNow = strSortAsc
620             lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
630             lngSortLine_Top = (lngSortLine_Top - lngTpp)
640             lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
650             .Controls(strSortLbl).Top = lngSortLbl_Top
660             .Controls(strSortLbl).Left = lngSortLbl_Left
670             .Controls(strSortLine).Top = lngSortLine_Top
680             .Controls(strSortLbl).Caption = strArwUp
690             .Controls(strSortLbl).ForeColor = CLR_DKBLU
700           End If
710         Else
720           .Controls(strSortLbl).Caption = strArwUp
730           .Controls(strSortLbl).ForeColor = CLR_DKBLU
740           Select Case strCalled
              Case "posted"
750             strSortAsc = "[posted]" & strStdAsc
760             lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
770             lngSortLine_Top = (lngSortLine_Top - lngTpp)
780             lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
790             .Controls(strSortLbl).Top = lngSortLbl_Top
800             .Controls(strSortLbl).Left = lngSortLbl_Left
810             .Controls(strSortLine).Top = lngSortLine_Top
820           Case "transdate"
830             strSortAsc = "[transdate]" & strStdAsc
840           Case "accountno"
850             strSortAsc = "[accountno]" & strStdAsc
860           Case "accountno2"
                ' ** accountno2.Column(1) {shortname}
870             strSortAsc = "[shortname]" & strStdAsc
880           Case "journaltype"
890             strSortAsc = "[journaltype]" & strStdAsc
900           Case "assetno", "assetno_description"
                ' ** assetno.Column(1) {assetno_description}
910             strSortAsc = "[assetno_description], [journaltype_sortorder]" & strStdAsc
920           Case "Recur_Name"
                ' ** Recur_Name.Column(1) {Recur_Name}
930             strSortAsc = "[Recur_Name]" & strStdAsc
940           Case "assetdate", "assetdate_display"
950             strSortAsc = "[assetdate]" & strStdAsc
960           Case "shareface"
970             strSortAsc = "[shareface], [assetno]" & strStdAsc
980           Case "icash"
990             strSortAsc = "[icash]" & strStdAsc
1000          Case "pcash"
1010            strSortAsc = "[pcash]" & strStdAsc
1020          Case "cost"
1030            strSortAsc = "[cost]" & strStdAsc
1040          Case "Location_ID", "Loc_Name_display"
1050            strSortAsc = "[Loc_Name]" & strStdAsc
1060          Case "PrintCheck"
1070            strSortAsc = "[PrintCheck]" & strStdAsc  ' ** -1 comes before 0.
1080            lngSortLbl_Left = (lngSortLbl_Left + (0& * lngTpp))
1090            lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
1100            lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
1110            .Controls(strSortLbl).Left = lngSortLbl_Left
1120            .Controls(strSortLine).Left = lngSortLine_Left
1130            .Controls(strSortLine).Width = lngSortLine_Width
1140          Case "description"
1150            strSortAsc = "[description]" & strStdAsc
1160          Case "revcode_ID", "revcode_DESC_display"
                ' ** revcode_ID.Column(1) {revcode_DESC}
1170            strSortAsc = "[revcode_DESC]" & strStdAsc
1180          Case "taxcode", "taxcode_description_display"
                ' ** taxcode.Column(1) {taxcode_description}
1190            strSortAsc = "[taxcode_description]" & strStdAsc
1200          Case "Reinvested"
1210            strSortAsc = "[reinvested] & strstdasc"
1220            lngSortLbl_Top = (lngSortLbl_Top + lngTpp)
1230            lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
1240            lngSortLine_Top = (lngSortLine_Top + lngTpp)
1250            lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
1260            lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
1270            .Controls(strSortLbl).Top = lngSortLbl_Top
1280            .Controls(strSortLbl).Left = lngSortLbl_Left
1290            .Controls(strSortLine).Top = lngSortLine_Top
1300            .Controls(strSortLine).Left = lngSortLine_Left
1310            .Controls(strSortLine).Width = lngSortLine_Width
1320          Case "journal_USER"
1330            strSortAsc = "[journal_USER]" & strStdAsc
1340          End Select
1350          If strSortNow = strSortAsc Then
1360            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
1370            Select Case intCnt
                Case 1
1380              strTmp01 = strSortAsc & " DESC"
1390            Case 2
1400              intPos01 = InStr(strSortAsc, ",")
1410              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
1420              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
1430            Case 3
1440              intPos01 = InStr(strSortAsc, ",")
1450              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
1460              strTmp02 = Mid(strSortAsc, intPos01)
1470              intPos01 = InStr(2, strTmp02, ",")
1480              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
1490              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
1500            Case 4
1510              intPos01 = InStr(strSortAsc, ",")
1520              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
1530              strTmp02 = Mid(strSortAsc, intPos01)
1540              intPos01 = InStr(2, strTmp02, ",")
1550              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
1560              strTmp02 = Mid(strTmp02, intPos01)
1570              intPos01 = InStr(2, strTmp02, ",")
1580              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
1590              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
1600            Case 5
1610              intPos01 = InStr(strSortAsc, ",")
1620              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
1630              strTmp02 = Mid(strSortAsc, intPos01)
1640              intPos01 = InStr(2, strTmp02, ",")
1650              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
1660              strTmp02 = Mid(strTmp02, intPos01)
1670              intPos01 = InStr(2, strTmp02, ",")
1680              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
1690              strTmp02 = Mid(strTmp02, intPos01)
1700              intPos01 = InStr(2, strTmp02, ",")
1710              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
1720              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
1730            End Select
1740            strSortNow = strTmp01
1750            .Controls(strSortLbl).Caption = strArwDn
1760            .Controls(strSortLbl).ForeColor = CLR_DKRED
1770          Else
1780            strSortNow = strSortAsc
1790          End If
1800        End If
1810      End If
1820      Select Case strCalled
          Case "posted", "JrnlCol_ID", "PrintCheck"
1830        .Controls(strSortLbl).Left = (.Controls(strSortLbl).Left + (7& * lngTpp))
1840      Case Else
            ' ** Leave it as-is.
1850      End Select
1860      .Controls(strSortLbl).Visible = True
1870      .Controls(strSortLine).Visible = True
1880      .OrderBy = strSortNow
1890      .OrderByOn = True

1900    End With

EXITP:
1910    Exit Sub

ERRH:
1920    Select Case ERR.Number
        Case Else
1930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1940    End Select
1950    Resume EXITP

End Sub

Public Function JC_Sort_Recur_Chk(frmSub As Access.Form) As Boolean

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Sort_Recur_Chk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRecurName As String
        Dim lngRecurID As Long, strRecurType As String
        Dim blnRetVal As Boolean

2010    blnRetVal = True

2020    With frmSub
2030      If IsNull(.Recur_Name) = False Then
2040        strRecurName = Trim(.Recur_Name)
2050        If strRecurName <> vbNullString Then
2060          Set dbs = CurrentDb
2070          With dbs
2080            lngRecurID = 0&
2090            strRecurType = vbNullString
                ' ** RecurringItems, by specified [rnam].
2100            Set qdf = .QueryDefs("qryJournal_Columns_13")
2110            With qdf.Parameters
2120              ![rnam] = strRecurName
2130            End With
2140            Set rst = qdf.OpenRecordset
2150            With rst
2160              If .BOF = True And .EOF = True Then
                    ' ** Manually entered Description.
2170              Else
2180                .MoveFirst
2190                lngRecurID = ![RecurringItem_ID]
2200                strRecurType = ![Recur_Type]
2210              End If
2220              .Close
2230            End With  ' ** rst.
2240            .Close
2250          End With  ' ** dbs.
2260          If lngRecurID > 0& Then
2270            .RecurringItem_ID = lngRecurID
2280            Select Case .journaltype
                Case "Misc."
2290              If strRecurType <> "Misc" Then
2300                blnRetVal = False
2310                MsgBox "The Recurring Item chosen is inappropriate for this transaction." & vbCrLf & vbCrLf & _
                      "Please choose a Misc Recurring Item, or enter a unique Description.", vbInformation + vbOKOnly, "Invalid Entry"
2320              End If
2330              If blnRetVal = True Then
2340                If IsNull(.ICash) = False And IsNull(.PCash) = False Then
2350                  If .ICash <> 0@ And .PCash <> 0@ Then
2360                    Select Case strRecurName
                        Case RECUR_I_TO_P
2370                      If .ICash >= 0@ Or .PCash <= 0 Then
2380                        blnRetVal = False
2390                        MsgBox "The Recurring Item chosen is inappropriate for this transaction." & vbCrLf & vbCrLf & _
                              "The Transfer Recurring Item does not match the current" & vbCrLf & _
                              "Income Cash and Principal Cash", vbInformation + vbOKOnly, "Invalid Entry"
2400                      End If
2410                    Case RECUR_P_TO_I
2420                      If .ICash <= 0@ Or .PCash >= 0 Then
2430                        blnRetVal = False
2440                        MsgBox "The Recurring Item chosen is inappropriate for this transaction." & vbCrLf & vbCrLf & _
                              "The Transfer Recurring Item does not match the current" & vbCrLf & _
                              "Income Cash and Principal Cash", vbInformation + vbOKOnly, "Invalid Entry"
2450                      End If
2460                    End Select
2470                  Else
                        ' ** Cash checking done elsewhere.
2480                  End If
2490                Else
                      ' ** Cash checking done elsewhere.
2500                End If
2510              End If  ' ** blnRetVal.
2520            Case "Paid"
2530              If strRecurType <> "Payee" Then
2540                blnRetVal = False
2550                MsgBox "The Recurring Item chosen is inappropriate for this transaction." & vbCrLf & vbCrLf & _
                      "Please choose a Payee Recurring Item, or enter a unique Description.", vbInformation + vbOKOnly, "Invalid Entry"
2560              End If
2570            Case "Received"
2580              If strRecurType <> "Payor" Then
2590                blnRetVal = False
2600                MsgBox "The Recurring Item chosen is inappropriate for this transaction." & vbCrLf & vbCrLf & _
                      "Please choose a Payor Recurring Item, or enter a unique Description.", vbInformation + vbOKOnly, "Invalid Entry"
2610              End If
2620            End Select
2630            If IsNull(.Recur_Type) = True Then
2640              .Recur_Type = strRecurType
2650            End If
2660          Else
2670            If IsNull(.RecurringItem_ID) = False Then
2680              .RecurringItem_ID = Null
2690            End If
2700          End If
2710        Else
2720          blnRetVal = JC_Sort_Recur_Null(frmSub)  ' ** Function: Below.
2730        End If
2740      Else
            ' ** Make sure they're all Null.
2750        blnRetVal = JC_Sort_Recur_Null(frmSub)  ' ** Function: Below.
2760      End If
2770    End With

EXITP:
2780    Set rst = Nothing
2790    Set qdf = Nothing
2800    Set dbs = Nothing
2810    JC_Sort_Recur_Chk = blnRetVal
2820    Exit Function

ERRH:
2830    blnRetVal = False
2840    Select Case ERR.Number
        Case Else
2850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2860    End Select
2870    Resume EXITP

End Function

Public Function JC_Sort_Recur_Null(frmSub As Access.Form) As Boolean
' ** Called by:
' **   JC_Sort_Recur_Chk(), Above
' **   frmJournal_Columns_Sub:
' **     posted_AfterUpdate()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Sort_Recur_Null"

        Dim blnRetVal As Boolean

2910    blnRetVal = True

2920    With frmSub
2930      If IsNull(.Recur_Name) = False Or IsNull(.Recur_Type) = False Or IsNull(.RecurringItem_ID) = False Then
2940        .Recur_Name = Null
2950        .Recur_Type = Null
2960        .RecurringItem_ID = Null
2970      End If
2980    End With

EXITP:
2990    JC_Sort_Recur_Null = blnRetVal
3000    Exit Function

ERRH:
3010    blnRetVal = False
3020    Select Case ERR.Number
        Case Else
3030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3040    End Select
3050    Resume EXITP

End Function
