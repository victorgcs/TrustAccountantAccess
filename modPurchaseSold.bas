Attribute VB_Name = "modPurchaseSold"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modPurchaseSold"

'VGC 07/19/2017: CHANGES!

'##########################################
'CURRENCY NOT YET FINISHED!!
'##########################################

' ** Combo box column constants: saleAssetNo.
'Private Const CBX_A_DESC   As Integer = 0  'totdesc
'Private Const CBX_A_CUSIP  As Integer = 1  'cusip
Private Const CBX_A_ASTNO  As Integer = 2  'assetno
'Private Const CBX_A_RATE   As Integer = 3  'rate
'Private Const CBX_A_DUE    As Integer = 4  'due
'Private Const CBX_A_TYPE   As Integer = 5  'assettype
'Private Const CBX_A_TDESC  As Integer = 6  'assettype_description
Private Const CBX_A_TAX    As Integer = 7  'taxcode
Private Const CBX_A_CURRID As Integer = 8  'curr_id

Private lngAssetNo As Long, strAccountNo As String
' **

Public Function OpenLotInfoForm(Optional varFromSaleBtn As Variant, Optional varCallingForm As Variant) As Integer
' ** Return Values:
' **    0 OK.
' **   -1 Input missing.
' **   -2 No holdings.
' **   -3 Insufficient holdings.
' **   -4 Zero shares.
' **   -9 Data problem.
' ** Called by:
' **   frmJournal_Sub4_Sold:
' **     Form_KeyDown()
' **     cmdSaleLotInfo_Click()
' **   frmJournal_Columns:
' **     pcash_KeyDown()
' **     pcash_Exit()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "OpenLotInfoForm"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstAssets As DAO.Recordset, rstTransactions As DAO.Recordset
        Dim frm As Access.Form
        Dim strCallingForm As String, strDocName As String
        Dim strAssetType As String
        Dim lngJournalID As Long, dblShareface As Double, dblShareFaceAA As Double, varAssetDate As Variant
        Dim blnD4D As Boolean, blnZeroShares As Boolean
        Dim blnFromSaleBtn As Boolean, blnDoMsg1 As Boolean, blnDoMsg2 As Boolean, blnFound As Boolean
        Dim lngRecs As Long
        Dim strTmp01 As String, dblTmp02 As Double, dblTmp03 As Double
        Dim lngX As Long
        Dim intRetVal As Integer

110     DoCmd.Hourglass True  ' ** Make sure it's still running.
120     DoEvents

130     intRetVal = 0
140     blnD4D = False  ' ** Unless proven otherwise.
150     blnZeroShares = False

160     Select Case IsMissing(varFromSaleBtn)
        Case True
170       blnFromSaleBtn = False
180     Case False
190       blnFromSaleBtn = CBool(varFromSaleBtn)
200     End Select

210     Select Case IsMissing(varCallingForm)
        Case True
220       strCallingForm = "frmJournal_Sub4_Sold"
230     Case False
240       strCallingForm = varCallingForm
250       If strCallingForm = "modJrnlCol_Forms" Then strCallingForm = "frmJournal_Columns_Sub"
260     End Select

270     Set dbs = CurrentDb

280     Select Case strCallingForm
        Case "frmJournal_Sub4_Sold", "frmJournal"
290       If strCallingForm = "frmJournal_Sub4_Sold" Then strCallingForm = "frmJournal"  ' ** Just to simplify things.
300       Set frm = Forms(strCallingForm).frmJournal_Sub4_Sold.Form
          ' ** ActiveAssets, for Cost = Null.
310       Set qdf = dbs.QueryDefs("qryLotInformation_01a")
320     Case "frmJournal_Columns", "frmJournal_Columns_Sub"
330       If strCallingForm = "frmJournal_Columns_Sub" Then strCallingForm = "frmJournal_Columns"  ' ** Just to simplify things.
340       Set frm = Forms(strCallingForm).frmJournal_Columns_Sub.Form
          ' ** ActiveAssets, for Cost = Null.
350       Set qdf = dbs.QueryDefs("qryLotInformation_01b")
360     End Select
370     DoEvents

380     Set rstAssets = qdf.OpenRecordset
390     With rstAssets
400       If .BOF = True And .EOF = True Then
            ' ** All's well.
410       Else
420         intRetVal = -9
430         DoCmd.Hourglass False
440         MsgBox "An Active Asset was found with a Null cost!" & vbCrLf & vbCrLf & _
              "Please contact Delta Data, Inc., for help with resolving this.", vbCritical + vbOKOnly, "Zero Cost Asset Found"
450       End If
460       .Close
470     End With
480     Set rstAssets = Nothing
490     Set qdf = Nothing
500     DoEvents

510     If intRetVal = 0 Then
520       blnDoMsg1 = False
530       Select Case strCallingForm
          Case "frmJournal"
540         If IsNull(frm.saleAssetno.Column(2)) = True Then  ' ** assetno.
550           blnDoMsg1 = True
560         Else
570           If IsNumeric(frm.saleAssetno.Column(2)) = False Then  ' ** assetno.
580             blnDoMsg1 = True
590           Else
600             If CLng(frm.saleAssetno.Column(2)) = 0& Then  ' ** assetno.
610               blnDoMsg1 = True
620             End If
630           End If
640         End If
650       Case "frmJournal_Columns"
660         If IsNull(frm.assetno) = True Then
670           blnDoMsg1 = True
680         Else
690           If frm.assetno = 0& Then
700             blnDoMsg1 = True
710           End If
720         End If
730       End Select
740       If blnDoMsg1 = True Then
750         intRetVal = -1
760         DoCmd.Hourglass False
770         MsgBox "You must select an asset to continue.", vbInformation + vbOKOnly, "Entry Required"
780       End If
790     End If  ' ** intRetVal.
800     DoEvents

810     If intRetVal = 0 Then

820       Select Case strCallingForm
          Case "frmJournal"
830         lngAssetNo = CLng(frm.saleAssetno.Column(2)) ' ** assetno.
840         strAccountNo = frm.saleAccountNo
850         lngJournalID = frm.saleID
860         dblShareface = frm.saleShareFace
870         dblShareFaceAA = 0#
880         varAssetDate = frm.saleAssetDate
            ' ** MasterAsset, by specified [astno].
890         Set qdf = dbs.QueryDefs("qryLotInformation_20a")
900         With qdf.Parameters
910           ![astno] = lngAssetNo
920         End With
930       Case "frmJournal_Columns"
940         lngAssetNo = frm.assetno
950         strAccountNo = frm.accountno
960         lngJournalID = frm.JrnlCol_ID
970         dblShareface = frm.shareface
980         dblShareFaceAA = 0#
990         varAssetDate = frm.assetdate
            ' ** MasterAsset, by specified [astno].
1000        Set qdf = dbs.QueryDefs("qryLotInformation_20b")
1010        With qdf.Parameters
1020          ![astno] = lngAssetNo
1030        End With
1040      End Select

1050      Set rstAssets = qdf.OpenRecordset
1060      With rstAssets
1070        If .BOF = True And .EOF = True Then
1080          intRetVal = -9
1090          DoCmd.Hourglass False
1100          MsgBox "A problem has occurred retrieving information about the chosen asset." & vbCrLf & _
                "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "MasterAsset Table Error"
1110        Else
1120          .MoveFirst
1130          strAssetType = Trim(rstAssets![assettype])
1140          Select Case strAssetType
              Case "60", "80", "81"  ' ** Money Market Funds, Interest Bearing - Own, Interest Bearing - Other.
1150            blnD4D = True  ' ** Dollar-For-Dollar.
1160          End Select
1170        End If
1180        .Close
1190      End With
1200      Set rstAssets = Nothing
1210      Set qdf = Nothing
1220      DoEvents

1230    End If  ' ** intRetVal.

1240    If intRetVal = 0 Then

          'START WITH TAKING EVERYTHING FROM ACTIVE ASSETS FOR THIS ACCOUNTNO, ASSETNO!
1250      Select Case blnD4D  ' ** Dollar-For-Dollar.
          Case True  ' ** Just 60, 80, 81.

1260        Select Case strCallingForm
            Case "frmJournal"
              ' ** Empty tmpEdit04.
1270          Set qdf = dbs.QueryDefs("qryLotInformation_21a")
1280          qdf.Execute
1290          Set qdf = Nothing
1300          DoEvents
              ' ** Append qryLotInformation_22a (ActiveAssets, with add'l fields, by specified
              ' ** [actno], [astno], for frmJournal), linked to tblCurrency, to tmpEdit04.  #curr_id
1310          Set qdf = dbs.QueryDefs("qryLotInformation_22c")
1320          With qdf.Parameters
1330            ![actno] = strAccountNo
1340            ![astno] = lngAssetNo
1350          End With
1360        Case "frmJournal_Columns"
              ' ** Empty tmpEdit04.
1370          Set qdf = dbs.QueryDefs("qryLotInformation_21b")
1380          qdf.Execute
1390          Set qdf = Nothing
1400          DoEvents
              ' ** Append qryLotInformation_22b (ActiveAssets, with add'l fields, by specified
              ' ** [actno], [astno], for frmJournal_Columns), linked to tblCurrency, to tmpEdit04.  #curr_id
1410          Set qdf = dbs.QueryDefs("qryLotInformation_22d")
1420          With qdf.Parameters
1430            ![actno] = strAccountNo
1440            ![astno] = lngAssetNo
1450          End With
1460        End Select
1470        qdf.Execute
1480        Set qdf = Nothing
1490        DoEvents

1500        If intRetVal = 0 Then

1510          Set rstAssets = Nothing

1520          Select Case strCallingForm
              Case "frmJournal"
                ' ** Update qryLotInformation_11a (tmpEdit04, with qryLotInformation_10a (tblForm_Graphics, just
                ' ** frmTaxLot, for cmdPrintReport), with frm_id_new, frmgfx_id_new, frmgfx_alt_new; Cartesian).
1530            Set qdf = dbs.QueryDefs("qryLotInformation_12a")
1540          Case "frmJournal_Columns"
                ' ** Update qryLotInformation_11b (tmpEdit04, with qryLotInformation_10b (tblForm_Graphics, just
                ' ** frmJournal_Columns_TaxLot, for cmdPrintReport), with frm_id_new, frmgfx_id_new, frmgfx_alt_new; Cartesian).
1550            Set qdf = dbs.QueryDefs("qryLotInformation_12b")
1560          End Select
1570          qdf.Execute
1580          Set qdf = Nothing
1590          DoEvents

1600          Set rstAssets = dbs.OpenRecordset("tmpEdit04", dbOpenDynaset)

              'THIS GETS ALL SALES CURRENTLY IN JOURNAL!
1610          Select Case strCallingForm
              Case "frmJournal"
                ' ** Journal, just 'Sold'/'Withdrawn', with add'l fields, by specified [actno], [astno].  #D4D  #curr_id
1620            Set qdf = dbs.QueryDefs("qryLotInformation_23a")
1630            With qdf.Parameters
1640              ![actno] = strAccountNo
                  'NO! WE NEED TO SEE ITS CURRENT STATE!
1650              ![astno] = lngAssetNo  'CLng(IIf(blnFromSaleBtn = False, lngAssetNo, 0&)) ' ** assetno.
1660            End With
                ' ** Includes: Journal_ID.
1670          Case "frmJournal_Columns"
                ' ** tblJournal_Column, just 'Sold'/'Withdrawn', with add'l fields, by specified [actno], [astno].  #D4D  #curr_id
1680            Set qdf = dbs.QueryDefs("qryLotInformation_23b")
1690            With qdf.Parameters
1700              ![actno] = strAccountNo
1710              ![astno] = lngAssetNo
1720            End With
                ' ** Includes: Journal_ID, JrnlCol_ID.
1730          End Select
1740          Set rstTransactions = qdf.OpenRecordset
1750          DoEvents

              ' ** It's looking for any 'Sold' or 'Withdrawn' transactions
              ' ** in the Journal, for this AccountNo and this AssetNo.
1760          If rstTransactions.BOF = True And rstTransactions.EOF = True Then
                ' ** Nothing to do.
                ' ** I presume this means it didn't expect to find the one we're on! (It may be blnFromSaleBtn = True.)
1770          Else
1780            rstTransactions.MoveLast
1790            lngRecs = rstTransactions.RecordCount
1800            rstTransactions.MoveFirst
1810            For lngX = 1& To lngRecs
                  ' ** Since both input forms save the Sold before coming here, it's expecting to find the one we're on.
1820              If IsNull(rstTransactions![PurchaseDate]) = False Then
                    ' ** Another Sold/Withdrawn for the same accountno, assetno.
1830                rstAssets.FindFirst "[assetdate2] = #" & Format(rstTransactions![PurchaseDate], "mm/dd/yyyy hh:nn:ss") & "#"
                    ' ** I'm going to try this assetdate2 for a while; see if it makes any difference.
1840                Select Case rstAssets.NoMatch
                    Case True
1850                  intRetVal = -9
1860                  Select Case strCallingForm
                      Case "frmJournal"
1870                    DoCmd.Hourglass False
1880                    MsgBox "A Sold asset was not found in ActiveAssets." & vbCrLf & vbCrLf & _
                          "AccountNo: " & frm.saleAccountNo & vbCrLf & _
                          "AssetNo: " & CStr(IIf(blnFromSaleBtn = False, CLng(frm.saleAssetno.Column(2)), 0&)) & vbCrLf & _
                          "Purchase Date: " & Format(rstTransactions![PurchaseDate], "mm/dd/yyyy hh:nn:ss") & vbCrLf & vbCrLf & _
                          "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & vbCrLf & "Line: 1440", _
                          vbCritical + vbOKOnly, "Asset Not Found"
1890                  Case "frmJournal_Columns"
1900                    DoCmd.Hourglass False
1910                    MsgBox "A Sold asset was not found in ActiveAssets." & vbCrLf & vbCrLf & _
                          "AccountNo: " & frm.accountno & vbCrLf & _
                          "AssetNo: " & CStr(IIf(blnFromSaleBtn = False, CLng(frm.assetno), 0&)) & vbCrLf & _
                          "Purchase Date: " & Format(rstTransactions![PurchaseDate], "mm/dd/yyyy hh:nn:ss") & vbCrLf & vbCrLf & _
                          "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & vbCrLf & "Line: 1440", _
                          vbCritical + vbOKOnly, "Asset Not Found"
1920                  End Select
1930                  Exit For
1940                Case False
                      ' ** NoMatch is False.
1950                  rstAssets.Edit  ' ** This is tmpEdit04.
1960                  rstAssets![shareface] = rstAssets![shareface] - rstTransactions![shareface]
                      '##########################################################################################
                      'WAIT A MINUTE! IT'S JUST LOOKING FOR EXACTLY ZERO! WHAT IF IT'S LESS THAN ZERO?
                      'THIS SEEMS TO BE LOOKING AT ANY OTHER COMPLETED SOLD/WITHDRAWN, AND
                      'APPLYING THEM TO tmpEDIT04, BUT IT'S DOING A PISS-POOR JOB OF IT!
                      'IF THE TRANSACTION WOULD DRAW THE TAX LOT DOWN TO LESS THAN ZERO,
                      'THIS CODE NEEDS TO SAVE THAT DIFFERENCE AND APPLY IT TO ANOTHER TAX LOT!
                      'OOPS! WAIT!...
                      'IF IT'S COMPLETED, AND IN THE JOURNAL, AND HAS A PURCHASE DATE,
                      'IT SHOULDN'T HAVE BEEN ABLE TO DRAW IT DOWN TO LESS THAN ZERO,
                      'BECAUSE THAT WOULD HAVE GENERATED A MULTI-LOT SALE, AND OVERAGE
                      'WOULD GO TO A DIFFERENT TAX LOT WITH A DIFFERENT PURCHASE DATE!
                      'I GUESS IT'S OK!
                      '##########################################################################################
1970                  If ((rstAssets![shareface] < 0.0001) And (rstAssets![shareface] > -0.0001)) Then  ' ** Make sure it catches mushy Zeroes!
1980                    rstAssets![shareface] = 0#  ' ** shareface2 has the original shareface.
1990                    rstAssets![Cost] = 0@
2000                    rstAssets.Update
2010                  ElseIf rstAssets![shareface] < 0 Then  ' ** Not supposed to happen!
2020                    rstAssets![shareface] = 0#
2030                    rstAssets![Cost] = 0@
2040                    rstAssets.Update
2050                  Else
2060                    rstAssets![Cost] = rstAssets![Cost] + rstTransactions![Cost]
2070                    rstAssets.Update
2080                  End If
2090                  If lngX < lngRecs Then rstTransactions.MoveNext
2100                End Select
2110              Else
                    ' ** PurchaseDate is Null.
                    ' ** An incomplete Sold/Withdrawn was found, presumably not this one!
                    'NO, NOT PRESUMABLY! OF COURSE IT FOUND THIS ONE!!!
2120                blnFound = False
2130                Select Case strCallingForm
                    Case "frmJournal"
2140                  If rstTransactions![Journal_ID] = lngJournalID Then
2150                    blnFound = True
2160                  End If
2170                Case "frmJournal_Columns"
2180                  If rstTransactions![JrnlCol_ID] = lngJournalID Then
2190                    blnFound = True
2200                  End If
2210                End Select
2220                If blnFound = True Then
                      ' ** Found this one, AS EXPECTED!, so keep looking.
                      ' ** And if that's all there is, we're fine! Proceed!
2230                  If lngX < lngRecs Then rstTransactions.MoveNext
2240                Else
2250                  intRetVal = -9
2260                  Select Case strCallingForm
                      Case "frmJournal"
2270                    DoCmd.Hourglass False
2280                    MsgBox "An incomplete Sold/Withdrawn transaction was found." & vbCrLf & _
                          "Cancel this transaction and start over.", vbCritical + vbOKOnly, "Cancel Transaction"
2290                  Case "frmJournal_Columns"
                        ' ** I'm going to see that this doesn't happen!
2300                    If IsNull(rstTransactions![Journal_ID]) = True Then
2310                      If rstTransactions![JrnlCol_ID] <> frm.JrnlCol_ID Then
2320                        DoCmd.Hourglass False
2330                        MsgBox "Another uncommitted transaction was found for this same Account and Asset." & vbCrLf & _
                              "Cancel this one and finish the other.", vbCritical + vbOKOnly, "Cancel Transaction"
2340                      Else
2350                        DoCmd.Hourglass False
2360                        MsgBox "The code has found this transaction, when it shouldn't have!", _
                              vbCritical + vbOKOnly, "Cancel Transaction"
2370                      End If
2380                    Else
                          ' ** It came from Outer Space! (I mean, the Journal.)
2390                      DoCmd.Hourglass False
2400                      MsgBox "An incomplete Journal transaction other than this one was found.", _
                            vbCritical + vbOKOnly, "Cancel Transaction"
2410                    End If
2420                  End Select
2430                  Exit For
2440                End If
2450              End If  ' ** purchaseDate.
2460            Next  ' ** lngX.
2470          End If  ' ** BOF, EOF.
2480          DoEvents

              'THIS SHOULD HAVE DRAWN DOWN THE TAX LOT TABLE FROM EXISTING JOURNAL ENTRIES!
              ' ** The above section seems to be adjusting the Tax Lots in
              ' ** tmpEdit04 for any previous Sold/Withdrawn in the Journal.
              ' ** On the face of it, it appears to be doing what it's supposed to!

2490          Select Case intRetVal
              Case 0
2500            rstAssets.Close
2510            rstTransactions.Close
2520            Set rstAssets = Nothing
2530            Set rstTransactions = Nothing
2540            Set qdf = Nothing
2550          Case Else
2560            rstAssets.Close
2570            rstTransactions.Close
2580            Set rstAssets = Nothing
2590            Set rstTransactions = Nothing
2600            Set qdf = Nothing
2610            dbs.Close
2620            Set dbs = Nothing
2630          End Select  ' ** intRetVal.
2640          DoEvents

2650        Else
2660          DoCmd.Hourglass False
2670          MsgBox "Temp table error in procedure '" & THIS_PROC & "'.", vbCritical + vbOKOnly, "Error"
2680          dbs.Close
2690          Set dbs = Nothing
2700          DoEvents
2710        End If  ' ** intRetVal.

2720      Case False
            ' ** Not Dollar-For-Dollar.

2730        Select Case strCallingForm
            Case "frmJournal"
              ' ** Empty tmpEdit04.
2740          Set qdf = dbs.QueryDefs("qryLotInformation_21a")
2750          qdf.Execute
2760          Set qdf = Nothing
2770          DoEvents
              ' ** Append qryLotInformation_24a (ActiveAssets, with add'l fields, by specified
              ' ** [actno], [astno], for frmJournal), linked to tblCurrency, to tmpEdit04.  #curr_id
2780          Set qdf = dbs.QueryDefs("qryLotInformation_24c")
2790          With qdf.Parameters
2800            ![actno] = strAccountNo
2810            ![astno] = lngAssetNo
2820          End With
2830        Case "frmJournal_Columns"
              ' ** Empty tmpEdit04.
2840          Set qdf = dbs.QueryDefs("qryLotInformation_21b")
2850          qdf.Execute
2860          Set qdf = Nothing
2870          DoEvents
              ' ** Append qryLotInformation_24b (ActiveAssets, with add'l fields, by specified
              ' ** [actno], [astno], for frmJournal_Columns), linked to tblCurrency, to tmpEdit04.  #curr_id
2880          Set qdf = dbs.QueryDefs("qryLotInformation_24d")
2890          With qdf.Parameters
2900            ![actno] = strAccountNo
2910            ![astno] = lngAssetNo
2920          End With
2930        End Select
2940        qdf.Execute
2950        Set qdf = Nothing
2960        DoEvents

2970        Select Case intRetVal
            Case 0
              ' ** That's it, proceed to form!
2980          Set rstAssets = Nothing

2990          Select Case strCallingForm
              Case "frmJournal"
                ' ** Update qryLotInformation_11a (tmpEdit04, with qryLotInformation_10a (tblForm_Graphics, just
                ' ** frmTaxLot, for cmdPrintReport), with frm_id_new, frmgfx_id_new, frmgfx_alt_new; Cartesian).
3000            Set qdf = dbs.QueryDefs("qryLotInformation_12a")
3010          Case "frmJournal_Columns"
                ' ** Update qryLotInformation_11b (tmpEdit04, with qryLotInformation_10b (tblForm_Graphics, just
                ' ** frmJournal_Columns_TaxLot, for cmdPrintReport), with frm_id_new, frmgfx_id_new, frmgfx_alt_new; Cartesian).
3020            Set qdf = dbs.QueryDefs("qryLotInformation_12b")
3030          End Select
3040          qdf.Execute
3050          Set qdf = Nothing
3060          DoEvents

              'SHOULDN'T THIS BE WHERE WE GO THROUGH AND UPDATE TMPEDIT04 FOR EXISTING JOURNAL ENTRIES?
              'WHERE IS IT?!!!!!!!!!!!!!
              'IS IT BECAUSE THEY SHOULDN'T HAVE BEEN ALLOWED TO GET THIS FAR?
              'IT'S NEEDED FOR THE LOT INFO BUTTON!

3070          Set rstAssets = dbs.OpenRecordset("tmpEdit04", dbOpenDynaset)

              'THIS GETS ALL SALES CURRENTLY IN JOURNAL!
3080          Select Case strCallingForm
              Case "frmJournal"
                ' ** Journal, just 'Sold'/'Withdrawn', with add'l fields,
                ' ** by specified [actno], [astno], for frmJournal.  #NOT D4D  #curr_id
3090            Set qdf = dbs.QueryDefs("qryLotInformation_23c")
3100            With qdf.Parameters
3110              ![actno] = strAccountNo
3120              ![astno] = lngAssetNo
3130            End With
                ' ** Includes: Journal_ID.
3140          Case "frmJournal_Columns"
                ' ** tblJournal_Column, just 'Sold','Withdrawn', with add'l fields,
                ' ** by specified [actno], [astno], for frmJournal_Columns.  #NOT D4D  #curr_id
3150            Set qdf = dbs.QueryDefs("qryLotInformation_23d")
3160            With qdf.Parameters
3170              ![actno] = strAccountNo
3180              ![astno] = lngAssetNo
3190            End With
                ' ** Includes: Journal_ID, JrnlCol_ID.
3200          End Select
3210          Set rstTransactions = qdf.OpenRecordset
3220          DoEvents

              ' ** It's looking for any 'Sold' or 'Withdrawn' transactions
              ' ** in the Journal, for this AccountNo and this AssetNo.
3230          If rstTransactions.BOF = True And rstTransactions.EOF = True Then
                ' ** Nothing to do.
                ' ** I presume this means it didn't expect to find the one we're on! (It may be blnFromSaleBtn = True.)
3240          Else
3250            rstTransactions.MoveLast
3260            lngRecs = rstTransactions.RecordCount
3270            rstTransactions.MoveFirst
3280            For lngX = 1& To lngRecs
                  ' ** Since both input forms save the Sold before coming here, it's expecting to find the one we're on.
3290              If IsNull(rstTransactions![PurchaseDate]) = False Then
                    ' ** Another Sold/Withdrawn for the same accountno, assetno.
3300                rstAssets.FindFirst "[assetdate2] = #" & Format(rstTransactions![PurchaseDate], "mm/dd/yyyy hh:nn:ss") & "#"
                    ' ** I'm going to try this assetdate2 for a while; see if it makes any difference.
3310                Select Case rstAssets.NoMatch
                    Case True
3320                  intRetVal = -9
3330                  Select Case strCallingForm
                      Case "frmJournal"
3340                    DoCmd.Hourglass False
3350                    MsgBox "A Sold asset was not found in ActiveAssets." & vbCrLf & vbCrLf & _
                          "AccountNo: " & frm.saleAccountNo & vbCrLf & _
                          "AssetNo: " & CStr(IIf(blnFromSaleBtn = False, CLng(frm.saleAssetno.Column(2)), 0&)) & vbCrLf & _
                          "Purchase Date: " & Format(rstTransactions![PurchaseDate], "mm/dd/yyyy hh:nn:ss") & vbCrLf & vbCrLf & _
                          "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & vbCrLf & "Line: 1440", _
                          vbCritical + vbOKOnly, "Asset Not Found"
3360                  Case "frmJournal_Columns"
3370                    DoCmd.Hourglass False
3380                    MsgBox "A Sold asset was not found in ActiveAssets." & vbCrLf & vbCrLf & _
                          "AccountNo: " & frm.accountno & vbCrLf & _
                          "AssetNo: " & CStr(IIf(blnFromSaleBtn = False, CLng(frm.assetno), 0&)) & vbCrLf & _
                          "Purchase Date: " & Format(rstTransactions![PurchaseDate], "mm/dd/yyyy hh:nn:ss") & vbCrLf & vbCrLf & _
                          "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & vbCrLf & "Line: 1440", _
                          vbCritical + vbOKOnly, "Asset Not Found"
3390                  End Select
3400                  Exit For
3410                Case False
                      ' ** NoMatch is False.
3420                  rstAssets.Edit  ' ** This is tmpEdit04.
3430                  rstAssets![shareface] = rstAssets![shareface] - rstTransactions![shareface]
3440                  If ((rstAssets![shareface] < 0.0001) And (rstAssets![shareface] > -0.0001)) Then  ' ** Make sure it catches mushy Zeroes!
3450                    rstAssets![shareface] = 0#  ' ** shareface2 has the original shareface.
3460                    rstAssets![Cost] = 0@
3470                    rstAssets.Update
3480                  ElseIf rstAssets![shareface] < 0 Then  ' ** Not supposed to happen!
3490                    rstAssets![shareface] = 0#
3500                    rstAssets![Cost] = 0@
3510                    rstAssets.Update
3520                  Else
                        'WHAT IS THIS DOING?
                        'I BELIEVE THIS IS ADDING OTHER ENTRIES IN THE JOURNAL TO THE
                        'TAX LOTS TO BE DISPLAYED SO THEY REFLECT WHAT'S AVAILABLE AT THIS MOMENT.
3530                    rstAssets![Cost] = rstAssets![Cost] + rstTransactions![Cost]
                        '################################################
                        '## NOT DONE YET!
                        'rstAssets![cost_usd] = rstAssets![cost_usd] + (rstTransactions![Cost] * THE APPROPRIATE RATE!
                        '################################################
3540                    rstAssets.Update
3550                  End If
3560                  If lngX < lngRecs Then rstTransactions.MoveNext
3570                End Select
3580              Else
                    ' ** PurchaseDate is Null.
                    ' ** An incomplete Sold/Withdrawn was found, presumably not this one!
                    'NO, NOT PRESUMABLY! OF COURSE IT FOUND THIS ONE!!!
3590                blnFound = False
3600                Select Case strCallingForm
                    Case "frmJournal"
3610                  If rstTransactions![Journal_ID] = lngJournalID Then
3620                    blnFound = True
3630                  End If
3640                Case "frmJournal_Columns"
3650                  If rstTransactions![JrnlCol_ID] = lngJournalID Then
3660                    blnFound = True
3670                  End If
3680                End Select
3690                If blnFound = True Then
                      ' ** Found this one, AS EXPECTED!, so keep looking.
                      ' ** And if that's all there is, we're fine! Proceed!
3700                  If lngX < lngRecs Then rstTransactions.MoveNext
3710                Else
3720                  intRetVal = -9
3730                  Select Case strCallingForm
                      Case "frmJournal"
3740                    DoCmd.Hourglass False
3750                    MsgBox "An incomplete Sold/Withdrawn transaction was found." & vbCrLf & _
                          "Cancel this transaction and start over.", vbCritical + vbOKOnly, "Cancel Transaction"
3760                  Case "frmJournal_Columns"
                        ' ** I'm going to see that this doesn't happen!
3770                    If IsNull(rstTransactions![Journal_ID]) = True Then
3780                      If rstTransactions![JrnlCol_ID] <> frm.JrnlCol_ID Then
3790                        DoCmd.Hourglass False
3800                        MsgBox "Another uncommitted transaction was found for this same Account and Asset." & vbCrLf & _
                              "Cancel this one and finish the other.", vbCritical + vbOKOnly, "Cancel Transaction"
3810                      Else
3820                        DoCmd.Hourglass False
3830                        MsgBox "The code has found this transaction, when it shouldn't have!", _
                              vbCritical + vbOKOnly, "Cancel Transaction"
3840                      End If
3850                    Else
                          ' ** It came from Outer Space! (I mean, the Journal.)
3860                      DoCmd.Hourglass False
3870                      MsgBox "An incomplete Journal transaction other than this one was found.", _
                            vbCritical + vbOKOnly, "Cancel Transaction"
3880                    End If
3890                  End Select
3900                  Exit For
3910                End If
3920              End If  ' ** purchaseDate.
3930            Next  ' ** lngX.
3940          End If  ' ** BOF, EOF.

              'THIS SHOULD HAVE DRAWN DOWN THE TAX LOT TABLE FROM EXISTING JOURNAL ENTRIES!
              ' ** The above section seems to be adjusting the Tax Lots in
              ' ** tmpEdit04 for any previous Sold/Withdrawn in the Journal.
              ' ** On the face of it, it appears to be doing what it's supposed to!

3950          Select Case intRetVal
              Case 0
3960            rstAssets.Close
3970            rstTransactions.Close
3980            Set rstAssets = Nothing
3990            Set rstTransactions = Nothing
4000            Set qdf = Nothing
4010          Case Else
4020            rstAssets.Close
4030            rstTransactions.Close
4040            Set rstAssets = Nothing
4050            Set rstTransactions = Nothing
4060            Set qdf = Nothing
4070            dbs.Close
4080            Set dbs = Nothing
4090          End Select  ' ** intRetVal.
4100          DoEvents

4110        Case Else
4120          DoCmd.Hourglass False
4130          MsgBox "Temp table error in procedure '" & THIS_PROC & "'.", vbCritical + vbOKOnly, "Error"
4140          dbs.Close
4150          Set dbs = Nothing
4160          DoEvents
4170        End Select  ' ** intRetVal.

4180      End Select  ' ** blnD4D.

4190    End If  ' ** intRetVal.

4200    If intRetVal = 0 Then
4210      Select Case strCallingForm
          Case "frmJournal"
            ' ** Table tmpEdit04, all fields.  #curr_id
4220        Set qdf = dbs.QueryDefs("qryLotInformation_27a")
4230      Case "frmJournal_Columns"
            ' ** Table tmpEdit04, all fields.  #curr_id
4240        Set qdf = dbs.QueryDefs("qryLotInformation_27b")
4250      End Select
4260      Set rstAssets = qdf.OpenRecordset
4270      With rstAssets
4280        If .BOF = True And .EOF = True Then
4290          intRetVal = -2
4300          DoCmd.Hourglass False
4310          MsgBox "This account has not purchased any shares/face of this asset," & vbCrLf & _
                "or other transactions not yet posted have zeroed it out.", vbInformation + vbOKOnly, "Invalid Assignment"
4320        Else
4330          .MoveLast
4340          If .RecordCount = 1 And IsNull(![cusip]) Then
4350            intRetVal = -2
4360            DoCmd.Hourglass False
4370            MsgBox "This account has not purchased any shares/face of this asset," & vbCrLf & _
                  "or other transactions not yet posted have zeroed it out.", vbInformation + vbOKOnly, "Invalid Assignment"
4380          Else
4390            lngRecs = .RecordCount
4400            .MoveFirst
4410            dblShareFaceAA = 0#
4420            For lngX = 1& To lngRecs
4430              dblShareFaceAA = dblShareFaceAA + ![shareface]
4440              If lngX < lngRecs Then .MoveNext
4450            Next
4460            If dblShareFaceAA < 0.0001 Then
                  ' ** This just catches previously edited entries from above, and does not compare it to the amount being sold.
4470              intRetVal = -2
4480              DoCmd.Hourglass False
4490              MsgBox "This account has sold/withdrawn all shares/face for this asset." & vbCrLf & _
                    "There is nothing left to sell.", vbInformation + vbOKOnly, "Invalid Assignment"
4500            End If
4510          End If
4520        End If
4530        .Close
4540      End With
4550      Set rstAssets = Nothing
4560      Set qdf = Nothing
4570      If intRetVal <> 0 Then
4580        dbs.Close
4590        Set dbs = Nothing
4600      End If
4610    End If  ' ** intRetVal.
4620    DoEvents

4630    If intRetVal = 0 Then

4640      Select Case strCallingForm
          Case "frmJournal"
            ' ** Table tmpEdit04, all fields.  #curr_id
4650        Set qdf = dbs.QueryDefs("qryLotInformation_27a")
4660      Case "frmJournal_Columns"
            ' ** Table tmpEdit04, all fields.  #curr_id
4670        Set qdf = dbs.QueryDefs("qryLotInformation_27b")
4680      End Select
4690      Set rstAssets = qdf.OpenRecordset
4700      With rstAssets
4710        .MoveLast
4720        lngRecs = .RecordCount
4730        .MoveFirst
4740        dblTmp02 = 0#: dblTmp03 = 0#
4750        For lngX = 1& To lngRecs
4760          dblTmp02 = dblTmp02 + ![shareface]
4770          If ![IsAverage] = False Then
4780            .Edit
4790            Select Case IsNull(![shareface])
                Case True
4800              ![shareface] = 0#
4810              dblTmp03 = 0#
4820              blnZeroShares = True
4830            Case False
4840              If ![shareface] = 0# Then
4850                dblTmp03 = 0#
4860                blnZeroShares = True
4870              Else
                    '################################################
                    '## NOT DONE YET!
4880                dblTmp03 = Abs(![Cost] / ![shareface])  ' ** Each Tax Lot has its own priceperunit.
                    '################################################
4890              End If
4900            End Select
4910            If (dblTmp03 <> 1) And (dblTmp03 >= 0.9999 And dblTmp03 <= 1.0001) Then dblTmp03 = 1#
4920            ![priceperunit] = CDbl(Format(dblTmp03, "$#,##0.00000"))
4930            ![priceperunit_raw] = dblTmp03  ' ** Without rounding.
4940            .Update
4950          End If
4960          If lngX < lngRecs Then .MoveNext
4970        Next
4980        .Close
4990      End With
5000      Set rstAssets = Nothing
5010      Set qdf = Nothing
5020      DoEvents

5030      Select Case strCallingForm
          Case "frmJournal"
            ' ** Update qryLotInformation_25a (Table tmpEdit04, just negative averagepriceperunit, priceperunit).
5040        Set qdf = dbs.QueryDefs("qryLotInformation_26a")
5050      Case "frmJournal_Columns"
            ' ** Update qryLotInformation_25b (Table tmpEdit04, just negative averagepriceperunit, priceperunit).
5060        Set qdf = dbs.QueryDefs("qryLotInformation_26b")
5070      End Select
5080      qdf.Execute dbFailOnError
5090      Set qdf = Nothing
5100      dbs.Close
5110      Set dbs = Nothing
5120      DoEvents

5130      blnDoMsg1 = False: blnDoMsg2 = False
5140      If Round(dblShareface, 4) > Round(dblTmp02, 4) Then
5150        If ((dblShareface - dblTmp02) <= 0.0001) Then
5160          blnDoMsg2 = True
5170        Else
5180          blnDoMsg1 = True
5190        End If
5200      End If
5210      If blnDoMsg1 = True Then
5220        intRetVal = -3
5230        DoCmd.Hourglass False
5240        Beep
5250        MsgBox "You can only sell " & CStr(dblTmp02) & " of this asset because this" & vbCrLf & _
              "account has purchased only " & CStr(dblTmp02) & " shares/face of this asset.", _
              vbInformation + vbOKOnly, "Invalid Assignment"
5260      ElseIf blnDoMsg2 = True Then
5270        intRetVal = -3
5280        Beep
5290        DoCmd.Hourglass False
5300        MsgBox "Due to share/face rounding, the volume being" & vbCrLf & _
              "sold is technically more than what is available," & vbCrLf & _
              "though by less than 0.0001 units." & vbCrLf & vbCrLf & _
              "Try splitting the sale into two transactions," & vbCrLf & _
              "the first short by one or two shares, then" & vbCrLf & _
              "the remainder in a second very small sale.", _
              vbInformation + vbOKOnly, "Rounding Difficulty"
5310      End If  ' ** blnDoMsg1.

5320      If blnZeroShares = True Then
            ' ** If a particular tax lot was Zeroed by a previous transaction
            ' ** in the Journal, than this is to be expected. Just mark that
            ' ** tax lot, and don't let them choose it.
5330      End If

5340    End If  ' ** intRetVal.

        'THIS SHOULD HAVE DRAWN DOWN THE TAX LOT TABLE FROM EXISTING JOURNAL ENTRIES!
5350    If intRetVal = 0 Then
          ' ** I've now appended the non-D4D to tmpEdit04 as well, so all
          ' ** versions can use that table alone for the form's RecordSource.

5360      Select Case strCallingForm
          Case "frmJournal"

5370        dblTmp02 = lngJournalID
5380        dblTmp03 = dblShareface
5390        If IsNull(varAssetDate) = True Then
5400          strTmp01 = "Null"
5410        Else
5420          strTmp01 = Format(varAssetDate, "mm/dd/yyyy hh:nn:ss")
5430        End If
5440        DoEvents

            ' ** acDialog removed because it can cause problems!
5450        gblnSetFocus = True
5460        strDocName = "frmTaxLot"
5470        DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & CStr(dblTmp02) & "~" & _
              CStr(dblTmp03) & "~" & strTmp01 & "~" & IIf(blnD4D = True, "True", "False")
            ' ** strCallingForm ~ ID ~ ShareFace ~ AssetDate ~ Dollar-for-Dollar

5480      Case "frmJournal_Columns"

5490        dblTmp02 = lngJournalID
5500        dblTmp03 = dblShareface
5510        If IsNull(varAssetDate) = True Then
5520          strTmp01 = "Null"
5530        Else
5540          strTmp01 = Format(varAssetDate, "mm/dd/yyyy hh:nn:ss")
5550        End If

5560        If strCallingForm = "frmJournal_Columns" Or strCallingForm = "frmJournal_Columns_Sub" Then
5570          If Val(gstrSaleICash) = 0 And Val(gstrSalePCash) = 0 Then
5580            gstrSaleICash = CStr(CDbl(frm.ICash))
5590            gstrSaleICash = Rem_Dollar(gstrSaleICash)  ' ** Module Function: modStringFuncs.
5600            gstrSalePCash = CStr(CDbl(frm.PCash))
5610            gstrSalePCash = Rem_Dollar(gstrSalePCash)  ' ** Module Function: modStringFuncs.
5620            If Val(gstrSalePCash) = 0 Then
5630  On Error Resume Next
5640              gstrSalePCash = CStr(CDbl(Val(frm.PCash.text)))
5650  On Error GoTo ERRH
5660              gstrSalePCash = Rem_Dollar(gstrSalePCash)  ' ** Module Function: modStringFuncs.
5670            End If
5680          End If
5690        End If
5700        DoEvents

            ' ** acDialog removed because it can cause problems!
5710        gblnSetFocus = True
5720        strDocName = "frmJournal_Columns_TaxLot"
5730        DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & CStr(dblTmp02) & "~" & _
              CStr(dblTmp03) & "~" & strTmp01 & "~" & IIf(blnD4D = True, "True", "False")

5740      End Select  ' ** strCallingForm.

5750    End If  ' ** intRetVal.

EXITP:
5760    Set frm = Nothing
5770    Set rstTransactions = Nothing
5780    Set rstAssets = Nothing
5790    Set dbs = Nothing
5800    OpenLotInfoForm = intRetVal
5810    Exit Function

ERRH:
5820    DoCmd.Hourglass False
5830    Select Case ERR.Number
        Case 3075  ' ** '|' in query expression '|'.
5840      If gstrSaleAccountNumber = vbNullString Or gstrSaleAccountNumber = "0" Then
5850        intRetVal = -1
5860        MsgBox "There must be an account number to continue.", vbInformation + vbOKOnly, "Entry Required"
5870      Else
5880        If gstrSaleType = vbNullString Then
5890          intRetVal = -1
5900          MsgBox "You must choose one sale type to continue.", vbInformation + vbOKOnly, "Entry Required"
5910        Else
5920          If gstrSaleAsset = vbNullString Or gstrSaleAsset = "0" Then
5930            intRetVal = -1
5940            MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
5950          Else
5960            If gstrSaleShareFace = vbNullString Or Val(gstrSaleShareFace) = 0 Then
5970              intRetVal = -1
5980              MsgBox "The Share/Face must be greater than zero.", vbInformation + vbOKOnly, "Entry Required"
5990            End If
6000          End If
6010        End If
6020      End If
6030    Case Else
6040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6050    End Select
6060    Resume EXITP

End Function

Public Sub DistributeCost(frmSub As Access.Form, Optional varJTypeOrd As Variant, Optional varSName As Variant, Optional varAssetDesc As Variant, Optional varLocID As Variant, Optional varLocName As Variant, Optional varRevID As Variant, Optional varRevDesc As Variant, Optional varRevType As Variant, Optional varTax As Variant, Optional varTaxDesc As Variant, Optional varTaxType As Variant)
' ** Called by:
' **   frmJournal_Sub4_Sold:
' **     SaleDistribute()
' **       Called by: cmdSaleOK_Click()
' **   frmJournal_Columns:
' **     cmdSpecPurp_CostAdj_Distribute_Click()
' **   frmJournal_Columns_Sub:
' **     description_KeyDown()
' **     revcode_DESC_display_KeyDown()
' **     revcode_ID_AfterUpdate()
' **     revcode_ID_KeyDown()
' **     taxcode_description_display_KeyDown()
' **     taxcode_AfterUpdate()
' **     taxcode_KeyDown()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "DistributeCost"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstAssets As DAO.Recordset, rstJournal As DAO.Recordset
        Dim ctlAccountNo As Access.Control, tbxAssetDate As Access.TextBox, tbxPurchaseDate As Access.TextBox
        Dim tbxShareFace As Access.TextBox, tbxICash As Access.TextBox, tbxPCash As Access.TextBox, tbxCost As Access.TextBox
        Dim varAccountNo As Variant, varAssetNo As Variant, varJType As Variant
        Dim varTranDate As Variant, varAssetDateAdj As Variant, varThisAssetDate As Variant
        Dim varCostAdj As Variant, varThisCost As Variant
        Dim varDescription As Variant, varSum As Variant, varIsAverage As Variant
        Dim strCallingForm As String
        Dim lngJrnlColID As Long
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnContinue As Boolean

6110    blnContinue = True

6120    With frmSub

6130      strCallingForm = .Parent.Name

6140      Select Case strCallingForm
          Case "frmJournal"
6150        Set ctlAccountNo = .Controls("saleAccountNo")
6160        Set tbxAssetDate = .Controls("saleAssetDate")
6170        Set tbxPurchaseDate = .Controls("UnitPurchaseDate")
6180        Set tbxShareFace = .Controls("saleShareFace")
6190        Set tbxICash = .Controls("saleICash")
6200        Set tbxPCash = .Controls("salePCash")
6210        Set tbxCost = .Controls("saleCost")
6220        varAccountNo = ctlAccountNo
6230        varAssetNo = .saleAssetno.Column(2)
6240        varJType = .saleType
6250        varTranDate = .saleTransDate
6260        varAssetDateAdj = tbxAssetDate
6270        varCostAdj = tbxCost
6280        varDescription = .saleDescription
6290        varIsAverage = .saleIsAverage
            'varAccountno    = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleAccountNo
            'varAssetNo      = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleAssetno.Column(2)
            'varJType        = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleType
            'varShareFace    = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleShareFace
            'varICash        = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleICash
            'varPCash        = Forms("frmJournal").frmJournal_Sub4_Sold.Form.salePCash
            'varCostAdj         = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleCost
            'varTranDate     = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleTransDate
            'varPurchaseDate = Forms("frmJournal").frmJournal_Sub4_Sold.Form.UnitPurchaseDate
            'varAssetDateAdj    = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleAssetDate
            'varDescription  = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleDescription
6300      Case "frmJournal_Columns"
6310        Set ctlAccountNo = .Controls("accountno")
6320        Set tbxAssetDate = .Controls("assetdate")
6330        Set tbxPurchaseDate = .Controls("purchaseDate")
6340        Set tbxShareFace = .Controls("shareface")
6350        Set tbxICash = .Controls("ICash")
6360        Set tbxPCash = .Controls("PCash")
6370        Set tbxCost = .Controls("Cost")
6380        varAccountNo = ctlAccountNo
6390        varAssetNo = .assetno
6400        varJType = .journaltype
6410        varTranDate = .transdate
6420        varAssetDateAdj = tbxAssetDate
6430        varCostAdj = tbxCost
6440        varDescription = .description
6450        varIsAverage = .IsAverage
6460      End Select  ' ** strCallingForm.

6470      Set dbs = CurrentDb

          'strSQL = "SELECT account.shortname, masterasset.cusip, CStr([masterasset].[Description]) & " & _
          '  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
          '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy')) AS totdesc, " & _
          '  "ActiveAssets.assetno, ActiveAssets.transdate, ActiveAssets.postdate, ActiveAssets.accountno, ActiveAssets.shareface, " & _
          '  "ActiveAssets.due, ActiveAssets.rate, ActiveAssets.averagepriceperunit, ActiveAssets.priceperunit, ActiveAssets.icash, " & _
          '  "ActiveAssets.pcash, ActiveAssets.cost, ActiveAssets.assetdate, ActiveAssets.description, ActiveAssets.posted, " & _
          '  "assettype_description, assettype.assettype, ActiveAssets.IsAverage, ActiveAssets.[Location_ID] "
          'strSQL = strSQL & "FROM account INNER JOIN (((ActiveAssets INNER JOIN masterasset ON ActiveAssets.assetno = masterasset.assetno) " & _
          '  "INNER JOIN assettype ON masterasset.assettype = assettype.assettype) " & _
          '  "LEFT JOIN Location ON ActiveAssets.[Location_ID] = Location.[Location_ID]) ON account.accountno = ActiveAssets.accountno " & _
          '  "WHERE (((ActiveAssets.assetno)=" & CStr(varAssetNo) & ") " & _
          '  "AND ((ActiveAssets.accountno) = '" & varAccountno & "')) " & _
          '  "ORDER BY ActiveAssets.assetdate ASC;"

6480      Select Case strCallingForm
          Case "frmJournal"
            ' ** ActiveAssets, with add'l fields, by specified [actno], [astno].
6490        Set qdf = dbs.QueryDefs("qryJournal_Sale_07")
6500      Case "frmJournal_Columns"
            ' ** ActiveAssets, with add'l fields, by specified [actno], [astno].
6510        Set qdf = dbs.QueryDefs("qryJournal_Columns_33_01")
6520      End Select
6530      With qdf.Parameters
6540        ![actno] = CStr(varAccountNo)
6550        ![astno] = CLng(varAssetNo)
6560      End With

          ' ** Takes all current TaxLots for 1 account and 1 asset,
          ' ** sums all their costs,
          ' **   (rstAssets![Cost] / varSum) * varCostAdj
          ' **   (FIRST COST / ALL COSTS) * THIS COST
          ' **   ($10,000 / $50,000) * $1,000
          ' **   (10000/50000) * 1000 = $200
          ' ** Then adds new Journal entries for each Tax Lot.

6570      Set rstAssets = qdf.OpenRecordset
6580      With rstAssets
6590        If .BOF = True And .EOF = True Then
6600          lngRecs = 0&
6610          blnContinue = False
6620          MsgBox "This account has not purchased any shares/face of this asset.", vbInformation + vbOKOnly, "Invalid Distribution"
6630        Else
6640          .MoveLast
6650          lngRecs = .RecordCount
6660          .MoveFirst
6670          If lngRecs = 1& And IsNull(![cusip]) = True Then
6680            lngRecs = 0&
6690            blnContinue = False
6700            MsgBox "This account has not purchased any shares/face of this asset.", vbInformation + vbOKOnly, "Invalid Distribution"
6710          End If
6720        End If
6730      End With  ' ** rstAssets.

6740      If blnContinue = True Then

6750        varSum = 0
6760        With rstAssets
6770          For lngX = 1& To lngRecs
6780            varSum = varSum + ![Cost]
6790            If lngX < lngRecs Then .MoveNext
6800          Next
6810        End With  ' ** rstAssets.

6820        If (varSum = 0) And (lngRecs > 1) Then
6830          blnContinue = False
6840          MsgBox "This asset has multiple lots, all without cost information.", vbInformation + vbOKOnly, "Invalid Entry"
6850        End If

6860      Else
6870        rstAssets.Close
6880        dbs.Close
6890      End If  ' ** blnContinue.

6900      If blnContinue = True Then

            'varCostAdj = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleCost

6910        With rstAssets

6920          .MoveFirst
6930          varThisCost = ![Cost]
6940          If varSum <> 0 Then
6950            tbxCost = Format(((varThisCost / varSum) * varCostAdj), "Currency")
6960            gstrSaleCost = Format(((varThisCost / varSum) * varCostAdj), "Currency")
6970            gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
6980          Else
                ' ** Just use COST as total because only one record without any cost information.
6990            gstrSaleCost = Format(varCostAdj, "Currency")
7000            gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
7010          End If
7020          tbxPurchaseDate = ![assetdate]

7030          tbxShareFace = 0
7040          tbxPCash = 0
7050          tbxICash = 0

              'varAccountno = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleAccountNo
              'varJType = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleType
              'varTranDate = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleTransDate
              'varDescription = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleDescription

7060          If Not IsDate(varAssetDateAdj) Then
7070            tbxAssetDate = Now()
7080          Else
7090            If CDbl(varAssetDateAdj) = CLng(varAssetDateAdj) Then
7100              tbxAssetDate = tbxAssetDate + time()
7110            End If
7120          End If

7130          varAssetDateAdj = tbxAssetDate
              'varAssetNo = Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleAssetno

7140          Select Case strCallingForm
              Case "frmJournal"
7150            Set rstJournal = dbs.OpenRecordset("journal", dbOpenDynaset, dbConsistent)
7160          Case "frmJournal_Columns"
7170            Set rstJournal = dbs.OpenRecordset("tblJournal_Column", dbOpenDynaset, dbConsistent)
7180          End Select

7190          .MoveNext  ' ** To record 2.
7200          For lngX = 2& To lngRecs
7210            varThisCost = ![Cost]
7220            varThisAssetDate = ![assetdate]
7230            With rstJournal
7240              Select Case strCallingForm
                  Case "frmJournal"
7250                .AddNew
7260                ![journaltype] = varJType
7270                ![transdate] = varTranDate
7280                ![accountno] = varAccountNo
7290                ![assetdate] = varAssetDateAdj
7300                ![PurchaseDate] = varThisAssetDate
7310                ![assetno] = varAssetNo
7320                ![shareface] = 0#
7330                ![PCash] = 0@
7340                ![ICash] = 0@
7350                ![Cost] = Format(((varThisCost / varSum) * varCostAdj), "Currency")
7360                gstrSaleCost = Format(((varThisCost / varSum) * varCostAdj), "Currency")
7370                gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
7380                ![description] = varDescription
7390                ![IsAverage] = varIsAverage
7400                ![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
7410                .Update
7420              Case "frmJournal_Columns"
7430                .AddNew
7440                ![Journal_ID] = CLng(0)
7450                ![FocusHolder] = Null
7460                ![posted] = CBool(False)
7470                ![journaltype] = varJType
7480                ![journalSubtype] = Null
7490                ![journaltype_sortorder] = varJTypeOrd
7500                ![transdate] = varTranDate
7510                ![Calendar1] = Null
7520                ![accountno] = varAccountNo
7530                ![shortname] = varSName
7540                ![assetdate] = varAssetDateAdj
7550                ![assetdate_display] = CDate(Format(varAssetDateAdj, "mm/dd/yyyy"))
7560                ![Calendar2] = Null
7570                ![PurchaseDate] = varThisAssetDate
7580                ![assetno] = varAssetNo
7590                ![assetno_description] = varAssetDesc
7600                ![Recur_Name] = Null
7610                ![Recur_Type] = Null
7620                ![RecurringItem_ID] = Null
7630                ![shareface] = 0#
7640                ![pershare] = 0#
7650                ![PCash] = 0@
7660                ![ICash] = 0@
7670                ![Cost] = Format(((varThisCost / varSum) * varCostAdj), "Currency")
7680                ![Reinvested] = CBool(False)
7690                ![PrintCheck] = CBool(False)
7700                ![CheckNum] = Null
7710                ![JrnlMemo_Memo] = Null
7720                ![JrnlMemo_HasMemo] = CBool(False)
7730                ![description] = varDescription
7740                ![Location_ID] = varLocID
7750                ![Loc_Name] = varLocName
7760                If varLocID = 1& Then
7770                  ![Loc_Name_display] = Null
7780                Else
7790                  ![Loc_Name_display] = varLocName
7800                End If
7810                ![revcode_ID] = varRevID
7820                ![revcode_DESC] = IIf(varRevDesc = "Null", Null, varRevDesc)
7830                ![revcode_DESC_display] = IIf(varRevDesc = "Null", Null, IIf(InStr(varRevDesc, "Unspecified") > 0, Null, varRevDesc))
7840                ![revcode_TYPE] = varRevType
7850                ![taxcode] = varTax
7860                ![taxcode_description] = IIf(varTaxDesc = "Null", Null, varTaxDesc)
7870                ![taxcode_description_display] = IIf(varTaxDesc = "Null", Null, IIf(InStr(varTaxDesc, "Unspecified") > 0, Null, varTaxDesc))
7880                ![taxcode_type] = varTaxType
7890                ![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
7900                ![rate] = 0#
7910                ![due] = Null
7920                ![assettype] = Null
7930                ![IsAverage] = varIsAverage
7940                ![JrnlCol_DateModified] = Now()
7950                .Update
7960                .Bookmark = .LastModified
7970                lngJrnlColID = ![JrnlCol_ID]
7980                frmSub.NewRecAdd lngJrnlColID  ' ** Form Procedure: frmJournal_Columns_Sub.
7990              End Select  ' ** strCallingForm.
8000            End With  ' ** rstJournal.
8010            If lngX < lngRecs Then .MoveNext
8020          Next

8030          rstJournal.Close

8040          .Close
8050        End With  ' ** rstAssets.

8060        dbs.Close

8070        .Requery

8080      End If  ' ** blnContinue.

8090    End With  ' ** frmSub.

EXITP:
8100    Set ctlAccountNo = Nothing
8110    Set tbxAssetDate = Nothing
8120    Set tbxPurchaseDate = Nothing
8130    Set tbxShareFace = Nothing
8140    Set tbxICash = Nothing
8150    Set tbxPCash = Nothing
8160    Set tbxCost = Nothing
8170    Set rstJournal = Nothing
8180    Set rstAssets = Nothing
8190    Set qdf = Nothing
8200    Set dbs = Nothing
8210    Exit Sub

ERRH:
8220    If ERR.Number <> 0 Then
8230      Select Case ERR.Number
          Case 3075  ' ** '|' in query expression '|'.
8240        If gstrSaleAccountNumber = vbNullString Or gstrSaleAccountNumber = "0" Then
8250          MsgBox "There must be an account number to continue.", vbInformation + vbOKOnly, "Entry Required"
8260          ctlAccountNo.SetFocus
8270        End If
8280        If gstrSaleType = vbNullString Then
8290          MsgBox "You must choose one sale type to continue.", vbInformation + vbOKOnly, "Entry Required"
8300        End If
8310        If gstrSaleAsset = vbNullString Or gstrSaleAsset = "0" Then
8320          MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
8330        End If
8340        If gstrSaleShareFace = vbNullString Or Val(gstrSaleShareFace) = 0 Then
8350          MsgBox "The Share/Face must be greater than zero.", vbInformation + vbOKOnly, "Entry Required"
8360        End If
8370      Case Else
8380        zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8390      End Select
8400    End If
8410    Resume EXITP

End Sub

Public Function StockSplits() As Boolean

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "StockSplits"

        Dim blnRetVal As Boolean

8510    blnRetVal = True

        ' ** 15 Recent Stock Splits:
        ' ** =======================
        ' ** 1.  Amphenol Corp. - Class A (APH)
        ' **       Stock Split: 2-for-1 on October 10, 2014
        ' ** 2.  Apple Inc. (AAPL)
        ' **       assetno:     48
        ' **       CUSIP:       037833100
        ' **       Stock Split: 7-for-1 on June 9, 2014
        ' **       transdate:   12/31/2013
        ' ** 3.  CF Industries Holdings Inc
        ' **       Stock Split: 5-for-1 on June 18, 2015
        ' ** 4.  EOG Resources Inc. (EOG)
        ' **       Stock Split: 2-for-1 on April 1, 2014
        ' ** 5.  Hanesbrands Inc. (HBI)
        ' **       Stock Split: 4-for-1 on March 4, 2015
        ' ** 6.  Marathon Petroleum Corp. (MPC)
        ' **       assetno:     247
        ' **       CUSIP:       56585A102
        ' **       Stock Split: 2-for-1 on June 11, 2015
        ' **       transdate:   12/31/2013
        ' ** 7.  Netflix (NFLX)
        ' **       Stock Split: 7-for-1 on July 2, 2015
        ' ** 8.  PPG Industries Inc. (PPG)
        ' **       Stock Split: 2-for-1 on June 15, 2015
        ' ** 9.  Qorvo Inc. (QRVO)
        ' **       Stock Split: 1-for-4 on January 2, 2015
        ' ** 10. Ross Stores Inc. (ROST)
        ' **       Stock Split: 2-for-1 on June 12, 2015
        ' ** 11. Starbucks Corp. (SBUX)
        ' **       Stock Split: 2-for-1 on April 9, 2015
        ' ** 12. Torchmark Corp. (TMK)
        ' **       Stock Split: 3-for-2 on July 2, 2014
        ' ** 13. Under Armour Inc. (UA)
        ' **       assetno:     308
        ' **       CUSIP:       904311107
        ' **       Stock Split: 2-for-1 on April 15, 2014
        ' ** 14. Union Pacific Corp. (UNP)
        ' **       Stock Split: 2-for-1 on June 9, 2014
        ' ** 15. Visa Inc. - Class A (V)
        ' **       assetno:     140
        ' **       CUSIP:       92826C839
        ' **       Stock Split: 4-for-1 on March 19, 2015
        ' **       transdate:   12/31/2012
        ' **       transdate:   12/31/2013
        ' **       transdate:   12/30/2014
        ' **       transdate:   01/31/2015
        ' **       transdate:   02/28/2015

        '  IIf(InStr([description],'Amphenol')>0 Or InStr([description],'Apple')>0 Or InStr([description],'CF Industries')>0 Or InStr([description],'EOG Resources')>0 Or InStr([description],'Hanesbrands')>0 Or InStr([description],'Marathon')>0 Or InStr([description],'Netflix')>0 Or InStr([description],'PPG Industries')>0 Or InStr([description],'Qorvo')>0 Or InStr([description],'Ross Stores')>0 Or InStr([description],'Starbucks')>0 Or InStr([description],'Torchmark')>0 Or InStr([description],'Under Armour')>0 Or InStr([description],'Union Pacific')>0 Or InStr([description],'Visa')>0,-1,0)

EXITP:
8520    StockSplits = blnRetVal
8530    Exit Function

ERRH:
8540    blnRetVal = False
8550    Select Case ERR.Number
        Case Else
8560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8570    End Select
8580    Resume EXITP

End Function

Public Sub subRevCode4(blnSpecialCap As Boolean, intSpecialCapOpt As Integer, frm As Access.Form)
' ** cmbRevenueCodes:
' **   RowSource is 0-Based:
' **     Col 0: revcode_ID
' **     Col 1: revcode_DESC
' **     Col 2: revcode_TYPE
' **     Col 3: revcode_TYPE_Code (I/E)
' **     Col 4: taxcode_type
' **     Col 5: taxcode_type_Code (I/D)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "subRevCode4"

        Dim curICash As Currency, curPCash As Currency, curCost As Currency
        Dim strRevCode As String, lngTaxcode As Long

8610    With frm
8620  On Error Resume Next
8630      strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
8640  On Error GoTo ERRH
8650  On Error Resume Next
8660      lngTaxcode = Nz(.cmbTaxCodes, 0&)
8670  On Error GoTo ERRH
8680      curICash = Nz(.saleICash)
8690      curPCash = Nz(.salePCash)
8700      curCost = Nz(.saleCost, 0)
8710      If IsNull(.saleType) = False Then
8720        Select Case .saleType
            Case "Sold"
              ' ** INCOME.
8730          If .cmbRevenueCodes.RowSource <> "qryRevCodeComboI" Then
8740            .cmbRevenueCodes.RowSource = "qryRevCodeComboI"
8750            .cmbRevenueCodes.Requery
8760          End If
8770          If strRevCode <> "I" Then
8780            .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
8790          End If
8800          If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
8810            .cmbRevenueCodes.Enabled = True
8820            .cmbRevenueCodes.BorderColor = CLR_LTBLU2
8830            .cmbRevenueCodes.BackStyle = acBackStyleNormal
8840            .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
8850            .cmbRevenueCodes_lbl_box.Visible = False
8860          ElseIf curICash <> 0 Then
8870            If gblnRevenueExpenseTracking = True Then
8880              .cmbRevenueCodes.Enabled = True
8890              .cmbRevenueCodes.BorderColor = CLR_LTBLU2
8900              .cmbRevenueCodes.BackStyle = acBackStyleNormal
8910              .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
8920              .cmbRevenueCodes_lbl_box.Visible = False
8930            End If
8940          Else
8950            .cmbRevenueCodes.Enabled = False
8960            .cmbRevenueCodes.BorderColor = WIN_CLR_DISR
8970            .cmbRevenueCodes.BackStyle = acBackStyleTransparent
8980            .cmbRevenueCodes_lbl.BackStyle = acBackStyleTransparent
8990            .cmbRevenueCodes_lbl_box.Visible = True
9000          End If
9010        Case "Withdrawn", "Cost Adj."
              ' ** ALL.
9020          If .cmbRevenueCodes.RowSource <> "qryRevCodeComboIE" Then
9030            .cmbRevenueCodes.RowSource = "qryRevCodeComboIE"
9040            .cmbRevenueCodes.Requery
9050          End If
9060          If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
9070            Select Case IsNull(.cmbRevenueCodes)
                Case True
9080              .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
9090            Case False
9100              If .cmbRevenueCodes > 0& Then
9110                If gblnLinkRevTaxCodes = True Then
9120                  If lngTaxcode > 0& Then  ' ** Tax Code takes precedence over Revenue Code!
9130                    If .cmbTaxCodes.Column(2) = 1 And strRevCode <> "I" Then  ' ** taxcode_type, Income.
9140                      .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
9150                    ElseIf .cmbTaxCodes.Column(2) = 2 And strRevCode <> "E" Then  ' ** taxcode_type, Deduction.
9160                      .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
9170                    Else
                          ' ** Let it stand.
9180                    End If
9190                  Else
                        ' ** Let it stand.
9200                  End If
9210                Else
                      ' ** Let it stand.
9220                End If
9230              Else
9240                .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
9250              End If
9260            End Select
9270          Else
9280            If IsNull(.cmbRevenueCodes) = False Then
9290              If .cmbRevenueCodes > 0& Then
9300                If gblnLinkRevTaxCodes = True Then
9310                  If lngTaxcode > 0& Then  ' ** Tax Code takes precedence over Revenue Code!
9320                    If .cmbTaxCodes.Column(2) = 1 And strRevCode <> "I" Then  ' ** taxcode_type, Income.
9330                      .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
9340                    ElseIf .cmbTaxCodes.Column(2) = 2 And strRevCode <> "E" Then  ' ** taxcode_type, Deduction.
9350                      .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
9360                    End If
9370                  End If
9380                End If
9390              Else
9400                .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
9410              End If
9420            Else
9430              .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
9440            End If
9450          End If
9460        Case "Liability"
              ' ** EXPENSE.
9470          If .cmbRevenueCodes.RowSource <> "qryRevCodeComboE" Then
9480            .cmbRevenueCodes.RowSource = "qryRevCodeComboE"
9490            .cmbRevenueCodes.Requery
9500          End If
9510          If strRevCode <> "E" Then
9520            .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
9530          End If
9540          If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
9550            .cmbRevenueCodes.Enabled = True
9560            .cmbRevenueCodes.BorderColor = CLR_LTBLU2
9570            .cmbRevenueCodes.BackStyle = acBackStyleNormal
9580            .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
9590            .cmbRevenueCodes_lbl_box.Visible = False
9600          ElseIf curICash <> 0 Then
9610            If gblnRevenueExpenseTracking = True Then
9620              .cmbRevenueCodes.Enabled = True
9630              .cmbRevenueCodes.BorderColor = CLR_LTBLU2
9640              .cmbRevenueCodes.BackStyle = acBackStyleNormal
9650              .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
9660              .cmbRevenueCodes_lbl_box.Visible = False
9670            End If
9680          ElseIf curPCash <> 0 Then
9690            If gblnRevenueExpenseTracking = True Then
9700              .cmbRevenueCodes.Enabled = True
9710              .cmbRevenueCodes.BorderColor = CLR_LTBLU2
9720              .cmbRevenueCodes.BackStyle = acBackStyleNormal
9730              .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
9740              .cmbRevenueCodes_lbl_box.Visible = False
9750            End If
9760          Else
9770            .cmbRevenueCodes.Enabled = False
9780            .cmbRevenueCodes.BorderColor = WIN_CLR_DISR
9790            .cmbRevenueCodes.BackStyle = acBackStyleTransparent
9800            .cmbRevenueCodes_lbl.BackStyle = acBackStyleTransparent
9810            .cmbRevenueCodes_lbl_box.Visible = True
9820          End If
9830        End Select
9840      End If
9850    End With

EXITP:
9860    Exit Sub

ERRH:
9870    Select Case ERR.Number
        Case Else
9880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9890    End Select
9900    Resume EXITP

End Sub

Public Sub subTaxCode4(blnSpecialCap As Boolean, intSpecialCapOpt As Integer, frm As Access.Form)
' ** cmbTaxCodes:
' **   RowSource is 0-Based:
' **     Col 0: taxcode
' **     Col 1: taxcode_description
' **     Col 2: taxcode_type
' **     Col 3: taxcode_type_Code (I/D)
' **     Col 4: revcode_TYPE
' **     Col 5: revcode_TYPE_Code (I/E)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "subTaxCode4"

        Dim curICash As Currency, curCost As Currency
        Dim strRevCode As String, lngTaxcode As Long

10010   With frm
10020 On Error Resume Next
10030     strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
10040 On Error GoTo ERRH
10050 On Error Resume Next
10060     lngTaxcode = Nz(.cmbTaxCodes, 0&)
10070 On Error GoTo ERRH
10080     curICash = Nz(.saleICash)
10090     curCost = Nz(.saleCost, 0)
10100     If IsNull(.saleType) = False Then
10110       Select Case .saleType
            Case "Sold"
              ' ** INCOME.
10120         If .cmbTaxCodes.RowSource <> "qryTaxCode_02" Then
10130           .cmbTaxCodes.RowSource = "qryTaxCode_02"  ' ** INCOME.
10140           .cmbTaxCodes.Requery
10150           If lngTaxcode > 0& Then
10160             If .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
10170               .cmbTaxCodes = .saleAssetno.Column(CBX_A_TAX)  ' ** All AssetType-based Tax Codes are INCOME.
10180             End If
10190           End If
10200         End If
10210         If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
10220           .cmbTaxCodes.Enabled = True
10230           .cmbTaxCodes.BorderColor = CLR_LTBLU2
10240           .cmbTaxCodes.BackStyle = acBackStyleNormal
10250           .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
10260           .cmbTaxCodes_lbl_box.Visible = False
10270         ElseIf curICash <> 0 Then
10280           If gblnIncomeTaxCoding = True Then
10290             .cmbTaxCodes.Enabled = True
10300             .cmbTaxCodes.BorderColor = CLR_LTBLU2
10310             .cmbTaxCodes.BackStyle = acBackStyleNormal
10320             .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
10330             .cmbTaxCodes_lbl_box.Visible = False
10340           End If
10350         Else
10360           .cmbTaxCodes.Enabled = False
10370           .cmbTaxCodes.BorderColor = WIN_CLR_DISR
10380           .cmbTaxCodes.BackStyle = acBackStyleTransparent
10390           .cmbTaxCodes_lbl.BackStyle = acBackStyleTransparent
10400           .cmbTaxCodes_lbl_box.Visible = True
10410         End If
10420       Case "Withdrawn", "Cost Adj."
10430         If .cmbTaxCodes.RowSource <> "qryTaxCode_05" Then
10440           .cmbTaxCodes.RowSource = "qryTaxCode_05"  ' ** ALL.
10450           .cmbTaxCodes.Requery
10460         End If
10470         If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
10480           Select Case IsNull(.cmbTaxCodes)
                Case True
10490             .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
10500           Case False
10510             If .cmbTaxCodes > 0& Then
10520               If gblnLinkRevTaxCodes = True Then
10530                 If .cmbTaxCodes.Column(2) = 1 And strRevCode = "E" Then  ' ** taxcode_type, Income.
10540                   .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
10550                 ElseIf .cmbTaxCodes.Column(2) = 2 And strRevCode = "I" Then  ' ** taxcode_type, Deduction.
10560                   .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
10570                 Else
                        ' ** Let it stand.
10580                 End If
10590               Else
                      ' ** Let it stand.
10600               End If
10610             Else
10620               .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
10630             End If
10640           End Select
10650         Else
10660           If IsNull(.cmbTaxCodes) = True Then
10670             .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
10680           Else
10690             If gblnLinkRevTaxCodes = True Then
10700               If .cmbTaxCodes > 0 Then  ' ** Tax Code takes precedence over Revenue Code!
10710                 If .cmbTaxCodes.Column(2) = 1 And strRevCode = "E" Then  ' ** taxcode_type, Income.
10720                   .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
10730                 ElseIf .cmbTaxCodes.Column(2) = 2 And strRevCode = "I" Then  ' ** taxcode_type, Deduction.
10740                   .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
10750                 End If
10760               End If
10770             End If
10780           End If
10790         End If
10800       Case "Liability"
              ' ** EXPENSE.
10810         If .cmbTaxCodes.RowSource <> "qryTaxCode_04" Then
10820           .cmbTaxCodes.RowSource = "qryTaxCode_04"  ' ** EXPENSE LTD.
10830           .cmbTaxCodes.Requery
10840         End If
10850         If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
10860           Select Case IsNull(.cmbTaxCodes)
                Case True
10870             .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
10880           Case False
10890             If lngTaxcode > 0& Then
10900               If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
10910                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
10920               Else
                      ' ** Let it stand.
10930               End If
10940             Else
10950               .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
10960             End If
10970           End Select
10980         Else
10990           If IsNull(.cmbTaxCodes) = True Then
11000             .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
11010           Else
11020             If lngTaxcode > 0& Then
11030               If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
11040                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
11050               End If
11060             End If
11070           End If
11080         End If
11090       End Select
11100     End If
11110   End With

EXITP:
11120   Exit Sub

ERRH:
11130   Select Case ERR.Number
        Case Else
11140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11150   End Select
11160   Resume EXITP

End Sub

Public Sub SaleLotInfo_PS(blnClickedLotInfo As Boolean, blnBeenToLotInfo As Boolean, blnGoingToReport As Boolean, blnGoneToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_ID As Long, frm As Access.Form)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "SaleLotInfo_PS"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strMsg As String
        Dim intStyle As Integer
        Dim strTitle1 As String
        Dim blnContinue As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim lngRetVal As Long

11210   blnContinue = True

11220   With frm

11230     blnClickedLotInfo = True

11240     .salePCash.SetFocus

11250     If IsNull(.salePCash) = False Then
11260       If .salePCash <> 0 Then
11270         blnContinue = False
11280       End If
11290     End If
11300     If IsNull(.saleCost) = False Then
11310       If .saleCost <> 0 Then
11320         blnContinue = False
11330       End If
11340     End If

11350     Select Case blnContinue
          Case True

11360       .cmdSaleLotInfo.SetFocus

11370       If IsNull(.saleAccountNo) = True Then
11380         blnContinue = False
11390         MsgBox "An Account Number must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
11400         .saleAccountNo.SetFocus
11410       Else
11420         If Trim(.saleAccountNo) = vbNullString Then
11430           blnContinue = False
11440           MsgBox "An Account Number must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
11450           .saleAccountNo.SetFocus
11460         Else
11470           If IsNull(.saleAssetno) = True Then
11480             blnContinue = False
11490             MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
11500             .saleAssetno.SetFocus
11510           Else
11520             If .saleAssetno = 0& Then
11530               blnContinue = False
11540               MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
11550               .saleAssetno.SetFocus
11560             End If
11570           End If
11580         End If
11590       End If

11600       If blnContinue = True Then
11610         gstrFormQuerySpec = .Parent.Name  ' ** In case it's been Nulled out by one of the Map forms.
11620         .FromSaleBtn = True
11630         lngRetVal = OpenLotInfoForm(True, THIS_NAME)  ' ** Function: Above.  'OK!
              ' ** Return Values:
              ' **    0 OK.
              ' **   -1 Input missing.
              ' **   -2 No holdings.
              ' **   -3 Insufficient holdings.
              ' **   -4 Zero shares.
              ' **   -9 Data problem.
11640         DoEvents
11650         blnBeenToLotInfo = True
11660         If gblnGoToReport = True Then
11670           blnGoneToReport = True
11680           lngGTR_ID = .saleID
11690           Forms("frmTaxLot").TimerInterval = 50&
11700           blnGoingToReport = False
11710           blnGTR_Emblem = False
11720           .GoToReport_arw_tl_taxlot_img.Visible = False
11730           .cmdSaleCancel.SetFocus
11740         End If
11750       End If

11760     Case False
11770       .cmdSaleLotInfo.SetFocus
11780       strMsg = "The Lot Info screen is unavailable by button at this time. It will" & vbCrLf & _
              "automatically pop up when you press Enter or Tab from the Cost field." & vbCrLf & vbCrLf & _
              "You may only browse Tax Lots by clicking" & vbCrLf & _
              "here before cash or cost has been specified."
11790       intStyle = vbInformation + vbOKOnly + vbDefaultButton1
11800       strTitle1 = "Lot Info Unavailable By Button" & Space(40)
11810       msgResponse = MsgBox(strMsg, intStyle, strTitle1)
11820     End Select

11830     blnClickedLotInfo = False

11840   End With

EXITP:
11850   Set qdf = Nothing
11860   Set dbs = Nothing
11870   Exit Sub

ERRH:
11880   Select Case ERR.Number
        Case Else
11890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11900   End Select
11910   Resume EXITP

End Sub

Public Sub ChkAverage_PS_Sold(frm As Access.Form)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkAverage_PS_Sold"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String, lngAssetNo As Long

12010   With frm
12020     .saleIsAverage = False
12030     If IsNull(.saleAssetno) = False Then
12040       strAccountNo = .saleAccountNo
12050       lngAssetNo = .saleAssetno
12060       Set dbs = CurrentDb
            ' ** ActiveAssets, just IsAverage = True, grouped, by accountno, assetno, IsAverage, with cnt.
12070       Set qdf = dbs.QueryDefs("qryJournal_Sale_04d")
12080       Set rst = qdf.OpenRecordset
12090       If rst.BOF = True And rst.EOF = True Then
              ' ** No averaging anywhere.
12100         rst.Close
12110       Else
12120         rst.Close
              ' ** ActiveAssets, grouped, by accountno, assetno, IsAverage,
              ' ** with cnt, by specified [actno], [astno].
12130         Set qdf = dbs.QueryDefs("qryJournal_Sale_04e")
12140         With qdf.Parameters
12150           ![actno] = strAccountNo
12160           ![astno] = lngAssetNo
12170         End With
12180         Set rst = qdf.OpenRecordset
12190         If rst.BOF = True And rst.EOF = True Then
                ' ** This asset not averaged for this account.
12200           rst.Close
12210         Else
12220           rst.MoveFirst
12230           .saleIsAverage = True
12240           rst.Close
12250         End If
12260       End If
12270       dbs.Close
12280     End If
12290   End With

EXITP:
12300   Set rst = Nothing
12310   Set qdf = Nothing
12320   Set dbs = Nothing
12330   Exit Sub

ERRH:
12340   Select Case ERR.Number
        Case Else
12350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12360   End Select
12370   Resume EXITP

End Sub

Public Sub SoldCostedSet_PS(blnOKCancel As Boolean, blnSoldCosted As Boolean, frm As Access.Form)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "SoldCostedSet_PS"

        Dim varTmp00 As Variant

12410   With frm
12420     blnSoldCosted = blnOKCancel
12430     Select Case blnSoldCosted
          Case True
12440       .saleType.Locked = True
12450       .saleShareFace.Locked = True
12460       Select Case .saleType
            Case "Sold", "Liability"
12470         .saleICash.Locked = True
12480         .salePCash.Locked = True
12490       Case "Withdrawn", "Cost Adj."
              ' ** Leave them unlocked, so they look disabled.
12500       End Select
12510       .saleCost.Locked = True
12520       .cmdPaidTotal.Enabled = False
12530     Case False
12540       .saleType.Locked = False
12550       .saleShareFace.Locked = False
12560       Select Case .saleType
            Case "Sold", "Liability"
12570         .saleICash.Locked = False
12580         .salePCash.Locked = False
12590       Case "Withdrawn", "Cost Adj."
              ' ** They're disabled.
12600       End Select
12610       .saleCost.Locked = False
12620       .cmdPaidTotal.Enabled = False
12630       varTmp00 = DCount("*", "journal", "[journaltype] = 'Paid'")
12640       If IsNull(varTmp00) = False Then
12650         If varTmp00 > 0 Then
12660           .cmdPaidTotal.Enabled = True
12670         End If
12680       End If
12690     End Select
12700   End With

EXITP:
12710   Exit Sub

ERRH:
12720   Select Case ERR.Number
        Case Else
12730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12740   End Select
12750   Resume EXITP

End Sub

Public Sub Sub3Purch_RevCode(frm As Access.Form)
' ** cmbRevenueCodes:
' **   RowSource is 0-Based:
' **     Col 0: revcode_ID
' **     Col 1: revcode_DESC
' **     Col 2: revcode_TYPE
' **     Col 3: revcode_TYPE_Code (I/E)
' **     Col 4: taxcode_type
' **     Col 5: taxcode_type_Code (I/D)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "Sub3Purch_RevCode"

        Dim strRevCode As String

12810   With frm
12820 On Error Resume Next
12830     strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
12840 On Error GoTo ERRH
12850     If IsNull(.purchaseType) = False Then
12860       Select Case .purchaseType
            Case "Purchase", "Deposit"
              ' ** INCOME.
12870         If .cmbRevenueCodes.RowSource <> "qryRevCodeComboI" Then
12880           .cmbRevenueCodes.RowSource = "qryRevCodeComboI"
12890           .cmbRevenueCodes.Requery
12900         End If
12910         If strRevCode <> "I" Then
12920           .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
12930         End If
12940       Case "Liability"
              ' ** EXPENSE.
12950         If .cmbRevenueCodes.RowSource <> "qryRevCodeComboE" Then
12960           .cmbRevenueCodes.RowSource = "qryRevCodeComboE"
12970           .cmbRevenueCodes.Requery
12980         End If
12990         If strRevCode <> "E" Then
13000           .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
13010         End If
13020       End Select
13030     End If
13040   End With

EXITP:
13050   Exit Sub

ERRH:
13060   Select Case ERR.Number
        Case Else
13070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13080   End Select
13090   Resume EXITP

End Sub

Public Sub Sub3Purch_TaxCode(frm As Access.Form)
' ** cmbTaxCodes:
' **   RowSource is 0-Based:
' **     Col 0: taxcode
' **     Col 1: taxcode_description
' **     Col 2: taxcode_type
' **     Col 3: taxcode_type_Code (I/D)
' **     Col 4: revcode_TYPE
' **     Col 5: revcode_TYPE_Code (I/E)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "Sub3Purch_TaxCode"

        Dim lngTaxcode As Long

13110   With frm
13120 On Error Resume Next
13130     lngTaxcode = Nz(.cmbTaxCodes, 0&)
13140 On Error GoTo ERRH
13150     If IsNull(.purchaseType) = False Then
13160       Select Case .purchaseType
            Case "Purchase", "Deposit"
              ' ** INCOME.
13170         If .cmbTaxCodes.RowSource <> "qryTaxCode_02" Then
13180           .cmbTaxCodes.RowSource = "qryTaxCode_02"  ' ** INCOME.
13190           .cmbTaxCodes.Requery
13200           If lngTaxcode > 0& Then
13210             If .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
13220               .cmbTaxCodes = .purchaseAssetNo.Column(CBX_A_TAX)  ' ** All AssetType-based Tax Codes are INCOME.
13230             End If
13240           End If
13250         End If
13260       Case "Liability"
              ' ** EXPENSE.
13270         If .cmbTaxCodes.RowSource <> "qryTaxCode_04" Then
13280           .cmbTaxCodes.RowSource = "qryTaxCode_04"  ' ** EXPENSE LTD.
13290           .cmbTaxCodes.Requery
13300           If lngTaxcode > 0& Then
13310             If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
13320               .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
13330             End If
13340           End If
13350         End If
13360       End Select
13370     End If
13380   End With

EXITP:
13390   Exit Sub

ERRH:
13400   Select Case ERR.Number
        Case Else
13410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13420   End Select
13430   Resume EXITP

End Sub

Public Sub Sub3Purch_Changed(blnChanged As Boolean, frm As Access.Form)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "Sub3Purch_Changed"

13510   With frm
13520     Select Case blnChanged
          Case True
13530       gblnPurchaseChanged = True
13540       .NavigationButtons = False
13550       DoCmd.SelectObject acForm, .Parent.Name, False
13560       With .Parent
13570         .FocusHolder.SetFocus
13580         .opgJournal.Enabled = False
13590         .opgJournal_optPurchase_lbl_box.Visible = True
13600         .cmdSwitch.Enabled = False
13610         .cmdSwitch_raised_img_dis.Visible = True
13620         .cmdSwitch_raised_img.Visible = False
13630         .cmdSwitch_raised_semifocus_dots_img.Visible = False
13640         .cmdSwitch_raised_focus_img.Visible = False
13650         .cmdSwitch_raised_focus_dots_img.Visible = False
13660         .cmdSwitch_sunken_focus_dots_img.Visible = False
13670         .frmJournal_Sub3_Purchase.SetFocus
13680       End With
13690       .cmdPurchaseClose.Enabled = False
13700       .cmdPurchaseOK.Enabled = True
13710       .cmdPurchaseCancel.Enabled = True
13720       .cmdPurchaseAssetNew.Enabled = True
13730       .cmdPurchaseLocNew.Enabled = True
13740       .Parent.NavVis False  ' ** Form Procedure: frmJournal.
13750     Case False
13760       gblnPurchaseChanged = False
13770       .NavigationButtons = True
13780       DoCmd.SelectObject acForm, .Parent.Name, False
13790       With .Parent
13800         .opgJournal.Enabled = True
13810         .opgJournal_optPurchase_lbl_box.Visible = False
13820         .cmdSwitch.Enabled = True
13830         .cmdSwitch_raised_img.Visible = True
13840         .cmdSwitch_raised_img_dis.Visible = False
13850         .cmdSwitch_raised_semifocus_dots_img.Visible = False
13860         .cmdSwitch_raised_focus_img.Visible = False
13870         .cmdSwitch_raised_focus_dots_img.Visible = False
13880         .cmdSwitch_sunken_focus_dots_img.Visible = False
13890       End With
13900       .cmdPurchaseClose.Enabled = True
13910       .cmdPurchaseOK.Enabled = False
13920       .cmdPurchaseCancel.Enabled = False
13930       .cmdPurchaseAssetNew.Enabled = False
13940       .cmdPurchaseLocNew.Enabled = False
13950       .Parent.NavVis True  ' ** Form Procedure: frmJournal.
13960     End Select
13970   End With

EXITP:
13980   Exit Sub

ERRH:
13990   Select Case ERR.Number
        Case Else
14000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14010   End Select
14020   Resume EXITP

End Sub
