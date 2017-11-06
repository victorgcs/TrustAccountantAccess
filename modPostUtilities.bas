Attribute VB_Name = "modPostUtilities"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modPostUtilities"

'VGC 10/29/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const HasRepost = 0  ' ** 0 = Repost module not included; -1 = Repost module included.
' ** Also in:
' **   zz_mod_MDEPrepFuncs
' **   zz_mod_ModuleMiscFuncs

'###########################################################################
'SHOULD ANY ATTEMPT BE MADE TO CROSS-CHECK TRANSDATE WITH PRICING HISTORY?
'###########################################################################

' ** ARE THERE ANY SYSTEM-GENERATED ENTRIES ANYMORE?
' **   "System Generated Item", "System"
' ** qryJournal_User_01a assigns this to Journal entries
' ** that have no journal_USER. This should no longer happen.

' ** Liability Criteria, Liability Rules:
' **   ONLY PCash and Cost define it!
' **   Liability with '+' (positive) OR ZERO pcash : Purchase  with  (negative) cost
' **   Liability with '-' (negative) pcash         : Sold      with  (positive) cost
' ** A Sale may have '-' (negative) icash, which
' ** is for interest tacked on top of the principal.
' ** A Purcahse should never have an icash amount.

' ** In Journal reports, all these fields are present and used in varying ways:
' **   Jcomment
' **   totdesc
' **   totdesc1
' **   totdesc2
' ** Some of the information is showing up twice!
' **   Jcomment:
' **     IIf(IsNull([journal].[description])=True,Null,IIf(InStr([journal].[description],'Stock Split')>0,IIf(InStr([journal].[description],',')>0,Left([journal].[description],(InStr([journal].[description],',')-1)),IIf(InStr([journal].[description],'Payable')>0,Left([journal].[description],(InStr([journal].[description],'Payable')-1)),[journal].[description])),[journal].[description]))
' **     Just [description], and no other stuff added.  .._02a
' **   totdescx:
' **     IIf(IsNull([RecurringItem]),'',IIf([journaltype]='Received',[RecurringItem],IIf([journaltype]='Paid',[RecurringItem],[RecurringItem]))) & IIf(IsNull([assetno])=False,IIf(IsNull([assetdate])=True Or ([journaltype] In ('Dividend','Interest')),'',Format([assetdate],'mm/dd/yyyy') & ' ') & '^SHAREFACE^' & CStr([Description]) & ' ' & IIf([rate]>0,' ' & Format([rate],'#,##0.000%'),'') & IIf(IsNull([due])=False,'  Due ' & Format([due],'mm/dd/yyyy')),' ' & [jcomment] & IIf([SingleUser]=True,'',' posted by ' & [journal_user]))
' **     Includes [RecurringItem], and both [description] and [Jcomment]!  .._03a
' **   totdesc1:
' **     IIf(InStr(Nz([totdescx],''),'^SHAREFACE^')=0,Nz([totdescx],''),Left([totdescx],(InStr(Nz([totdescx],''),'^SHAREFACE^')-1)) & IIf([shareface]-CLng([shareface])=0,Format([shareface],'#,##0'),Format([shareface],'#,##0.0000')) & ' ' & Mid([totdescx],(InStr(Nz([totdescx],''),'^SHAREFACE^')+Len('^SHAREFACE^'))))
' **     Replaces totdescx, with ^SHAREFACE^ replaced by actual value.  .._03c
' **   totdesc2:
' **     IIf(IsNull([assetno]),'',[Jcomment] & IIf([SingleUser]=True,'',' posted by ' & [journal_user]))
' **     Includes [Jcomment]!  .._03a
' **   totdesc:
' **      Trim(IIf(InStr(Nz([totdescx],''),'^SHAREFACE^')=0,Nz([totdescx],''),Left([totdescx],(InStr(Nz([totdescx],''),'^SHAREFACE^')-1)) & IIf([shareface]-CLng([shareface])=0,Format([shareface],'#,##0'),Format([shareface],'#,##0.0000')) & ' ' & Mid([totdescx],(InStr(Nz([totdescx],''),'^SHAREFACE^')+Len('^SHAREFACE^')))) & ' ' & [totdesc2])
' **     Replaces totdescx, with ^SHAREFACE^ replaced by actual value, and adds [totdesc2]!  .._03c
' **
' ** So, totdesc has RecurringItem, description, and Jcomment from totdescx, plus Jcomment again from totdesc2!
' ** SOME THINGS COULD BE THERE 3 TIMES!
' **

Public Const POST_INIT    As Integer = 0 ' ** Initialized.
Public Const POST_NOTRANS As Integer = 1 ' ** No transactions.
Public Const POST_CANCEL  As Integer = 2 ' ** User Canceled.
Public Const POST_NULL    As Integer = 3 ' ** Null asset pricing.
Public Const POST_ROLLBK  As Integer = 4 ' ** Problem, caused rollback.
Public Const POST_ERROR   As Integer = 5 ' ** Other Error, from error handler.
Public Const POST_DONE    As Integer = 6 ' ** Post successful.

Private lngAssetDateMin As Long
' **

Public Function PostTransactions(Optional varAdminResponse As Variant) As Integer
' ** Posting procedure.
' ** Return codes:
' **   0  POST_INIT     Initialized.
' **   1  POST_NOTRANS  No transactions.
' **   2  POST_CANCEL   User Canceled.
' **   3  POST_NULL     Null asset pricing.
' **   4  POST_ROLLBK   Problem, caused rollback.
' **   5  POST_ERROR    Other Error, from error handler.
' **   6  POST_DONE     Post successful.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "PostTransactions"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, qdf3 As DAO.QueryDef
        Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim rstJournal As DAO.Recordset, rstAccount As DAO.Recordset, rstLedger As DAO.Recordset
        Dim rstMasterAsset As DAO.Recordset, rstActiveAssets As DAO.Recordset
        Dim rstErrLog As DAO.Recordset
        Dim grp As DAO.Group, usr As DAO.User
        Dim ctlPostMsg As Access.Label
        Dim strAccountNo As String, strAssetNo As String
        Dim datAssetDate As Date, datTransDate As Date, datPosted As Date  'Date type is 8 bytes; same as Double.
        Dim datPurchDate As Date
        Dim lngMax As Long, strErr As String
        Dim msgResponse As VbMsgBoxResult
        Dim intStyle As Integer, strTitle1 As String, strMsg As String
        Dim strUser As String, strPostType As String, strAdminResponse As String
        Dim strFindFirst As String
        Dim blnSkip As Boolean
        Dim blnRollback As Boolean, intRollbackPoint As Integer
        Dim blnInTrans As Boolean
        Dim blnHasForExJ As Boolean, lngCurrID As Long
        Dim lngImports As Long, arr_varImport() As Variant
        Dim lngPosPays As Long, arr_varPosPay() As Variant
        Dim lngRecs1 As Long, lngRecs2 As Long
        Dim lngRecsToPost As Long, lngRecsToPost_Selected As Long
        Dim lngDate_Year As Long, lngDate_Month As Long, lngDate_Day As Long
        Dim lngDate_Hour As Long, lngDate_Minute As Long, lngDate_Second As Long
        Dim strTmp01 As String, lngTmp02 As Long, dblTmp03 As Double, dblTmp04 As Double, dblTmp05 As Double
        Dim curTmp06 As Currency, curTmp07 As Currency, curTmp08 As Currency, dblTmp09 As Double
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim intRetVal As Integer, blnRetVal As Boolean

        Const POST_NON As String = "0"
        Const POST_ALL As String = "1"
        Const POST_SYS As String = "2"
        Const POST_USR As String = "3"

        ' ** Array: arr_varImport().
        Const I_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const I_JID  As Integer = 0
        Const I_ATYP As Integer = 1
        Const I_CNUM As Integer = 2
        Const I_PDAT As Integer = 3

        ' ** Array: arr_varPosPay().
        Const P_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const P_ID    As Integer = 0
        Const P_JNO   As Integer = 1
        Const P_ATYP  As Integer = 2
        Const P_PPDID As Integer = 3

      #If HasRepost Then
110     If gblnNoErrHandle_Repost = True Then
120   On Error GoTo 0
130     End If
      #End If

140     DoCmd.Hourglass True
150     DoEvents

        ' ***************************************
        ' ***************************************
        ' ** Step 1. Initializing
        ' ***************************************
        ' ***************************************
160     Set ctlPostMsg = Forms("frmMenu_Post").cmdPost_msg_lbl
170     ctlPostMsg.Caption = "1 of 15 - Initializing"
        'ctlPostMsg.Visible = True
180     DoEvents

190     intRetVal = POST_INIT  ' ** 0. Initialized.
200     gblnAdmin = False
210     blnSkip = False: blnHasForExJ = False
220     blnRollback = False
230     intRollbackPoint = 0
240     strPostType = POST_NON
250     blnInTrans = False
260     lngRecsToPost_Selected = 0&: lngCurrID = 0&

270     Select Case IsMissing(varAdminResponse)
        Case True
280       strAdminResponse = "All"
290     Case False
300       strAdminResponse = varAdminResponse
310     End Select

        ' ** Check Journal for presence of records. Save a step if no transactions...
      #If HasRepost Then
320     Select Case gblnRePost
        Case True
330       lngRecsToPost = DCount("ID", "zz_tbl_RePost_Journal")
340     Case False
350       lngRecsToPost = DCount("ID", "journal")
360     End Select
      #Else
370     lngRecsToPost = DCount("ID", "journal")
      #End If
380     If lngRecsToPost = 0 Then
390       intRetVal = POST_NOTRANS  ' ** 1. No transactions.
400       DoCmd.Hourglass False
410       MsgBox "No Transactions to Post!  ¹", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
420     Else

          ' ** Security Check.
430       For Each grp In DBEngine.Workspaces(0).Groups
440         If grp.Name = SGRP_ADMINS Then
              ' ** Any User found in Admins (¹) Group permitted access.
450           For Each usr In grp.Users
460             If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
470               gblnAdmin = True
480               Exit For
490             End If
500           Next
510         End If
520       Next
530       Set usr = Nothing
540       Set grp = Nothing
550       DoEvents

          ' ** But don't make us crazy with choices if there really aren't any....
          ' ** So if all the entries belong to the current user, just skip the choice...

          ' ** Check for Null assets.
560       Set dbs = CurrentDb
570       With dbs
            ' ** qryPost_Journal_13_01 (MasterAsset, just
            ' ** marketvaluecurrent = Null), grouped, with cnt_astno.
580         Set qdf1 = .QueryDefs("qryPost_Journal_13_02")
590         Set rst1 = qdf1.OpenRecordset
600         If rst1.BOF = True And rst1.EOF = True Then
              ' ** No Nulls; all's well.
610         Else
620           rst1.MoveFirst
630           If IsNull(rst1![cnt_astno]) = False Then
640             lngTmp02 = rst1![cnt_astno]
650             If lngTmp02 = 0& Then
                  ' ** No Nulls; all's well.
660               rst1.Close
670             Else
680               intRetVal = POST_NULL
690               strMsg = vbNullString
700               If lngTmp02 <= 2& Then
710                 rst1.Close
720                 Set rst1 = Nothing
730                 Set qdf1 = Nothing
                    ' ** MasterAsset, just marketvaluecurrent = Null
740                 Set qdf1 = .QueryDefs("qryPost_Journal_13_01")
750                 Set rst1 = qdf1.OpenRecordset
760                 With rst1
770                   .MoveFirst
780                   If lngTmp02 = 1& Then
790                     strMsg = "There is an unpriced asset present, and posting cannot continue." & vbCrLf
800                     strMsg = strMsg & "CUSIP: " & ![cusip] & vbCrLf
810                     strMsg = strMsg & "Desc: " & ![description]
820                   Else
830                     strMsg = "There are two unpriced assets present, and posting cannot continue." & vbCrLf
840                     strMsg = strMsg & "CUSIP: " & ![cusip] & vbCrLf
850                     strMsg = strMsg & "Desc: " & ![description]
860                     .MoveNext
870                     strMsg = strMsg & "CUSIP: " & ![cusip] & vbCrLf
880                     strMsg = strMsg & "Desc: " & ![description]
890                   End If
900                   strMsg = strMsg & vbCrLf & vbCrLf
910                   .Close
920                 End With  ' ** rst1.
930               Else
940                 strMsg = "There are several unpriced assets present, and posting cannot continue." & vbCrLf & vbCrLf
950               End If
960               strMsg = strMsg & "Please go to Asset Pricing on the Asset Menu and assure all" & vbCrLf & _
                    "assets have a Current Market Unit Value (which can be Zero)."
970               Beep
980               DoCmd.Hourglass False
990               MsgBox strMsg, vbInformation + vbOKOnly, "Null Asset Pricing"
1000            End If
1010          Else
                ' ** No Nulls; all's well.
1020            rst1.Close
1030          End If
1040        End If
1050        Set rst1 = Nothing
1060        Set qdf1 = Nothing
1070        .Close
1080      End With  ' ** dbs.
1090      Set dbs = Nothing

1100    End If  ' ** lngRecsToPost.

1110    If lngRecsToPost > 0 And intRetVal = POST_INIT Then

          ' ***************************************
          ' ***************************************
          ' ** Step 2. User counts.
          ' ***************************************
          ' ***************************************
1120      ctlPostMsg.Caption = "2 of 15 - User Counts"
1130      DoEvents

      #If HasRepost Then
1140      Select Case gblnRePost
          Case True
1150        lngTmp02 = DCount("journal_USER", "zz_tbl_RePost_Journal", "journal_USER = CurrentUser")  ' ** Internal Access Function: Trust Accountant login.
1160      Case False
1170        lngTmp02 = DCount("journal_USER", "journal", "journal_USER = CurrentUser")  ' ** Internal Access Function: Trust Accountant login.
1180      End Select
      #Else
1190      lngTmp02 = DCount("journal_USER", "journal", "journal_USER = CurrentUser")  ' ** Internal Access Function: Trust Accountant login.
      #End If
1200      If lngTmp02 = lngRecsToPost Then
            ' ** If the only entries in the Journal are the current user's,
            ' ** don't bother checking for Admin status.
1210        strUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1220        strPostType = POST_USR
1230      Else
            ' ** If multiple users are in the Journal, check for Admin status.
1240        Select Case gblnAdmin
            Case True
      #If HasRepost Then
1250          Select Case gblnRePost
              Case True
1260            strUser = "All"
1270            lngRecsToPost = DCount("ID", "zz_tbl_RePost_Journal")
1280          Case False
                ' ** Check if Admin already responded to this.
1290            If strAdminResponse = vbNullString Then
                  ' ** If Admin, either this one, or its duplicate in cmdPost_Click(), should be the 1st message they receive.
1300              strUser = SelectUser  ' ** Function: Below.
1310            Else
1320              strUser = strAdminResponse
1330            End If
1340            Select Case strUser
                Case "All"
1350              lngRecsToPost = DCount("ID", "journal")
1360            Case Else
1370              lngRecsToPost = DCount("journal_USER", "journal", "journal_USER = '" & strUser & "'")
1380              lngRecsToPost_Selected = lngRecsToPost
1390            End Select
1400          End Select
      #Else
              ' ** Check if Admin already responded to this.
1410          If strAdminResponse = vbNullString Then
                ' ** If Admin, either this one, or its duplicate in cmdPost_Click(), should be the 1st message they receive.
1420            strUser = SelectUser  ' ** Function: Below.
1430          Else
1440            strUser = strAdminResponse
1450          End If
1460          Select Case strUser
              Case "All"
1470            lngRecsToPost = DCount("ID", "journal")
1480          Case Else
1490            lngRecsToPost = DCount("journal_USER", "journal", "journal_USER = '" & strUser & "'")
1500            lngRecsToPost_Selected = lngRecsToPost
1510          End Select
      #End If
              ' ** See if we really want to post entries. Do it now while we still have the exact selection text.
1520          If strUser = vbNullString Then
1530            blnSkip = True
1540          Else
1550            DoCmd.Hourglass False
                ' ** If Admin, this should only be the 2nd message they receive.
1560            strMsg = "Are you sure you want to post journal entries for " & strUser & " at this time?"
1570            intStyle = vbQuestion + vbYesNo + vbDefaultButton1
1580            strTitle1 = "Posting Journal Entries For " & strUser
      #If HasRepost Then
1590            Select Case gblnRePost
                Case True
1600              msgResponse = vbYes
1610            Case False
1620              msgResponse = MsgBox(strMsg, intStyle, strTitle1)
1630            End Select
      #Else
1640            msgResponse = MsgBox(strMsg, intStyle, strTitle1)
      #End If
1650            Select Case msgResponse
                Case vbNo
1660              intRetVal = POST_CANCEL  ' ** 2. User Canceled.
1670              blnSkip = True  ' ** Nope, get out.
1680            Case Else
1690              DoCmd.Hourglass True
1700              DoEvents
1710              Select Case strUser
                  Case "All"
1720                strUser = vbNullString
1730                strPostType = POST_ALL
      #If HasRepost Then
1740                Select Case gblnRePost
                    Case True
1750                  lngRecsToPost = DCount("ID", "zz_tbl_RePost_Journal")
1760                Case False
1770                  lngRecsToPost = DCount("ID", "journal")
1780                End Select
      #Else
1790                lngRecsToPost = DCount("ID", "journal")
      #End If
1800              Case "System"
1810                strUser = vbNullString
1820                strPostType = POST_SYS
                    ' *************************
                    ' ** Open Database.
                    ' *************************
1830                Set dbs = CurrentDb
      #If HasRepost Then
1840                Select Case gblnRePost
                    Case True
                      ' ** zz_tbl_RePost_Journal, with cnt of System entries.
1850                  Set qdf1 = dbs.QueryDefs("zz_qry_RePost_Journal_04")
1860                Case False
                      ' ** Journal, with cnt of System entries.
1870                  Set qdf1 = dbs.QueryDefs("qryPost_Journal_10")
1880                End Select
      #Else
                    ' ** Journal, with cnt of System entries.
1890                Set qdf1 = dbs.QueryDefs("qryPost_Journal_10")
      #End If
1900                Set rst1 = qdf1.OpenRecordset()
1910                With rst1
1920                  If .BOF = True And .EOF = True Then
1930                    lngRecsToPost = 0&
1940                  Else
1950                    .MoveFirst
1960                    lngRecsToPost = ![cnt]
1970                  End If
1980                  .Close
1990                End With
2000                Set rst1 = Nothing
2010                Set qdf1 = Nothing
2020                dbs.Close
2030                Set dbs = Nothing
2040                DoEvents
                    ' *************************
                    ' ** Close Database.
                    ' *************************
2050              Case Else
                    ' ** As chosen.
2060                strPostType = POST_USR
      #If HasRepost Then
2070                Select Case gblnRePost
                    Case True
2080                  lngRecsToPost = DCount("journal_USER", "zz_tbl_RePost_Journal", "journal_USER = '" & strUser & "'")
2090                Case False
2100                  lngRecsToPost = DCount("journal_USER", "journal", "journal_USER = '" & strUser & "'")
2110                End Select
      #Else
2120                lngRecsToPost = DCount("journal_USER", "journal", "journal_USER = '" & strUser & "'")
      #End If
2130              End Select
2140            End Select
2150          End If
2160        Case False
      #If HasRepost Then
2170          Select Case gblnRePost
              Case True
2180            lngRecsToPost = DCount("journal_USER", "zz_tbl_RePost_Journal", "journal_USER = CurrentUser")  ' ** Internal Access Function: Trust Accountant login.
2190          Case False
2200            lngRecsToPost = DCount("journal_USER", "journal", "journal_USER = CurrentUser")  ' ** Internal Access Function: Trust Accountant login.
2210          End Select
      #Else
2220          lngRecsToPost = DCount("journal_USER", "journal", "journal_USER = CurrentUser")  ' ** Internal Access Function: Trust Accountant login.
      #End If
2230          strUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2240          strPostType = POST_USR
2250        End Select
2260      End If

2270      If blnSkip = False Then
2280        gstrJournalUser = vbNullString  ' ** Just used to pass from frmMenu_Post_Clear_Multi.
2290        If gblnDemo Then
2300          blnSkip = True
2310          strPostType = POST_NON
2320        End If
2330      End If
          ' ** Nothing remains open at this point.

2340      If blnSkip = False Then

            ' ***************************************
            ' ***************************************
            ' ** Step 3. Journal counts.
            ' ***************************************
            ' ***************************************
2350        ctlPostMsg.Caption = "3 of 15 - Journal Counts"
2360        DoEvents

            ' ** To get an accurate count of transactions being posted,
            ' ** use qryPost_Journal_02, and qryPost_Journal_03 queries.
            ' ** Though lngRecsToPost has been set above, just replace it.
2370        lngRecsToPost = TransCount  ' ** Function: Below.

2380        lngRecs1 = DCount("*", "journal", "[curr_id]<>150")
2390        If lngRecs1 > 0& Then
2400          blnHasForExJ = True
2410        End If

            ' *************************
            ' ** Open Workspace.
            ' *************************
2420        Set wrk = DBEngine.Workspaces(0)
            ' *************************
            ' ** Open Database.
            ' *************************
2430        Set dbs = wrk.Databases(0)

2440        DoCmd.Hourglass True
2450        DoEvents

            ' *************************
            ' ** BeginTrans.
            ' *************************
2460        wrk.BeginTrans
2470        blnInTrans = True

            ' ** Set a common [posted] and use throughout this procedure.
2480        datPosted = Now()

            ' ***************************************
            ' ***************************************
            ' ** Step 4. AssetDate check.
            ' ***************************************
            ' ***************************************
2490        ctlPostMsg.Caption = "4 of 15 - AssetDate Check"
2500        DoEvents

            ' ** 0. Check for presense of assetdate timestamp.
2510        With dbs
              ' ** Journal, just assetdate's without a timestamp.  '#jnox
2520          Set qdf1 = .QueryDefs("qryPost_Journal_12_" & strPostType)
2530          Select Case strPostType
              Case "3"
2540            With qdf1.Parameters
2550              ![usr] = strUser
2560            End With
2570          Case Else
                ' ** Nothing else.
2580          End Select
2590          Set rstJournal = qdf1.OpenRecordset
2600          With rstJournal
2610            If .BOF = True And .EOF = True Then
                  ' ** Good!
2620            Else
2630              .MoveLast
2640              lngTmp02 = .RecordCount
2650              .MoveFirst
2660              For lngX = 1& To lngTmp02
2670                .Edit
2680                ![assetdate] = ![assetdate_new]
2690                .Update
2700                If lngX < lngTmp02 Then .MoveNext
2710              Next
2720            End If
2730            .Close
2740          End With
2750          Set rstJournal = Nothing
2760          Set qdf1 = Nothing
2770          DoEvents
2780        End With  ' ** dbs.

            ' ***************************************
            ' ***************************************
            ' ** Step 5. Update Account Table.
            ' ***************************************
            ' ***************************************
2790        ctlPostMsg.Caption = "5 of 15 - Update Account Table"
2800        DoEvents

            ' ** 1. Update Account table for fields [icash], [pcash], [cost].
      #If HasRepost Then
2810        Select Case gblnRePost
            Case True
2820          Set rstAccount = dbs.OpenRecordset("zz_tbl_RePost_Account", dbOpenDynaset)  ' ** All account table records.
2830          Set qdf1 = dbs.QueryDefs("zz_qry_RePost_Journal_05_" & strPostType)  ' ** Journal is being read only at this point.
2840        Case False
2850          Set rstAccount = dbs.OpenRecordset("account", dbOpenDynaset)  ' ** All account table records.
2860          Select Case blnHasForExJ
              Case True
2870            Set qdf1 = dbs.QueryDefs("qryPost_Journal_01_4_" & strPostType)
2880          Case False
2890            Set qdf1 = dbs.QueryDefs("qryPost_Journal_01_" & strPostType)  ' ** Journal is being read only at this point.  '#jnox
2900          End Select
2910        End Select
      #Else
2920        Set rstAccount = dbs.OpenRecordset("account", dbOpenDynaset)  ' ** All account table records.
2930        Select Case blnHasForExJ
            Case True
2940          Set qdf1 = dbs.QueryDefs("qryPost_Journal_01_4_" & strPostType)
2950        Case False
2960          Set qdf1 = dbs.QueryDefs("qryPost_Journal_01_" & strPostType)  ' ** Journal is being read only at this point.  '#jnox
2970        End Select
      #End If
2980        If strPostType = POST_USR Then
2990          With qdf1.Parameters
3000            ![usr] = strUser
3010          End With
3020        End If
3030        Set rstJournal = qdf1.OpenRecordset()  ' ** Journal table records matching strPostType.
3040        With rstJournal
3050          .MoveLast
3060          lngRecs1 = .RecordCount
3070          .MoveFirst
3080          For lngX = 1& To lngRecs1
3090            strAccountNo = rstJournal![accountno]
3100            With rstAccount
3110              Select Case rstJournal![journaltype]  ' ** Both journaltype and journaltypex are in query.
                  Case "Purchase", "Deposit", "Sold", "Withdrawn", "Liability", "Misc.", "Paid", "Received", "Cost Adj."
3120                .FindFirst "accountno = '" & strAccountNo & "'"
3130                Select Case .NoMatch
                    Case True
3140                  strErr = "Unable to process record. Account " & strAccountNo & _
                        " was not found in active assets table. Rolling back changes!  Ref#1"
3150                  blnRollback = True
3160                  intRollbackPoint = 1
3170                Case False
3180                  .Edit
3190                  Select Case blnHasForExJ
                      Case True
                        ' ** If curr_id = 150, it just has the same values.
3200                    ![ICash] = ![ICash] + rstJournal![icash_usd]  ' ** dbCurrency
3210                    ![PCash] = ![PCash] + rstJournal![pcash_usd]  ' ** dbCurrency
3220                    ![Cost] = ![Cost] + rstJournal![cost_usd]     ' ** dbCurrency
3230                  Case False
3240                    ![ICash] = ![ICash] + rstJournal![ICash]  ' ** dbCurrency/dbDouble
3250                    ![PCash] = ![PCash] + rstJournal![PCash]  ' ** dbCurrency/dbDouble
3260                    ![Cost] = ![Cost] + rstJournal![Cost]     ' ** dbCurrency/dbDouble
3270                  End Select
                      ' ** Liability Note: These should be fine with no-cash transactions; Cost will go down, then up.
3280                  .Update  ' ** Should also have [modified_date] and [modified_user] fields.
3290                End Select
3300              Case "Interest", "Dividend"
3310                .FindFirst "accountno = '" & strAccountNo & "'"
3320                Select Case .NoMatch
                    Case True
3330                  strErr = "Unable to process record. Account " & strAccountNo & _
                        " was not found in active assets table. Rolling back changes!  Ref#2"
3340                  blnRollback = True
3350                  intRollbackPoint = 2
3360                Case False
3370                  .Edit
3380                  Select Case blnHasForExJ
                      Case True
3390                    ![ICash] = ![ICash] + rstJournal![icash_usd]
3400                  Case False
3410                    ![ICash] = ![ICash] + rstJournal![ICash]
3420                  End Select
3430                  .Update
3440                End Select
3450              End Select
3460              If blnRollback = True Then Exit For
                  ' #### rstAccount changed ####
3470            End With  ' ** rstAccount.
3480            If lngX < lngRecs1 Then .MoveNext
3490          Next  ' ** lngX.
3500        End With  ' ** rstJournal.
            ' ** rstAccount now updated with new amounts, rstJournal only read.
            ' ** Both rstAccount and rstJournal remain open.

            ' ** We're still within BeginTrans.
3510        If blnRollback = False Then

3520          rstJournal.Close
3530          rstAccount.Close
3540          Set rstJournal = Nothing
3550          Set rstAccount = Nothing
3560          Set qdf1 = Nothing
3570          DoEvents

              ' ***************************************
              ' ***************************************
              ' ** Step 6. New ActiveAssets Entries.
              ' ***************************************
              ' ***************************************
3580          ctlPostMsg.Caption = "6 of 15 - New ActiveAssets Entries"
3590          DoEvents

              ' ** 2. Append new ActiveAssets records.
              ' **    Deposits and Purchases that don't already have an entry with the same assetdate.
              ' **    Also positive Liabilities: pcash >= 0.
              ' **    datAssetDate is [assetdate] from the Journal entry.
              ' **    NOTE:
              ' **    ONCE THE 1ST ONE IS APPENDED, DOES THIS MEAN ADDITIONAL
              ' **    ENTRIES WITH THE SAME assetdate WILL ERROR OUT AS DUPES?
              ' **    YES! 02/17/09: Ignore for now; check and handle in next major upgrade, per Rich.
              ' **    Process would be to skip dupe accountno/assetno/assetdate combos here,
              ' **    then update that single new entry for subsequent items.
      #If HasRepost Then
3600          Select Case gblnRePost
              Case True
3610            Set rstActiveAssets = dbs.OpenRecordset("zz_tbl_RePost_ActiveAssets", dbOpenDynaset)  ' ** ActiveAssets table accepting new records.
3620            Set qdf1 = dbs.QueryDefs("zz_qry_RePost_ActiveAssets_01_" & strPostType)
3630          Case False
3640            Set rstActiveAssets = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)  ' ** ActiveAssets table accepting new records.
3650            Select Case blnHasForExJ
                Case True
3660              Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_01_5_" & strPostType)
3670            Case False
3680              Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_01_" & strPostType)
3690            End Select
3700          End Select
      #Else
3710          Set rstActiveAssets = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)  ' ** ActiveAssets table accepting new records.
3720          Select Case blnHasForExJ
              Case True
3730            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_01_5_" & strPostType)
3740          Case False
3750            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_01_" & strPostType)
3760          End Select
      #End If
              ' ** Liability Note: The query says '[pcash] >= 0', which includes,
              ' ** in error I think, transactions without ANY cash involved.
              ' ** I would say that no-cash Liabilities only get added if [cost] < 0.
              ' ** If [cost] > 0, then those are Sales.
              ' ** Added qualification:
              ' ** IIf([journaltype]<>'Liability',0,IIf([journal].[icash]=0 And [journal].[pcash]=0,IIf([journal].[cost]<0,-1,0),-1)) = -1
              ' ** This should leave alone Liabilities that DO include cash, to be processed as they've always been handled.
3770          If strPostType = POST_USR Then
3780            With qdf1.Parameters
3790              ![usr] = strUser
3800            End With
3810          End If
3820          Set rstJournal = qdf1.OpenRecordset()
3830          With rstJournal
3840            If .BOF = True And .EOF = True Then
                  ' ** No Deposits, Purchases, etc.; only Withdrawn, Sold, etc.
3850            Else
3860              .MoveLast
3870              lngRecs1 = .RecordCount
3880              .MoveFirst
3890              For lngX = 1& To lngRecs1
3900                strAccountNo = ![accountno]
3910                strAssetNo = CStr(![assetno])
3920                Select Case IsNull(![assetdate])
                    Case True
3930                  strErr = "Unable to process record. Asset = " & strAssetNo & _
                        ", Account = " & strAccountNo & ": Asset Date required. Rolling back changes!  Ref#4"
3940                  blnRollback = True
3950                  intRollbackPoint = 4
                      'DOES THIS ALSO ROLL BACK CHANGES MADE TO ACCOUNT, ABOVE?
3960                Case False
3970                  datAssetDate = ![assetdate]
3980                  If CLng(datAssetDate) < lngAssetDateMin Then
3990                    strErr = "Unable to process record. Asset = " & strAssetNo & _
                          ", Account = " & strAccountNo & ": Asset Date " & Format(![assetdate], "mm/dd/yyyy") & _
                          " appears inappropriate. Rolling back changes!  Ref#3"
4000                    blnRollback = True
4010                    intRollbackPoint = 3
      #If HasRepost Then
4020                    If gblnRePost = True Then
4030                      glngRePostErrNum = 2113&  ' ** The value you entered isn't valid for this field.
4040                      gintRePostErrLine = 2170
4050                      gstrRePostErrMsg = strErr
4060  On Error Resume Next
4070                      garr_varRePost(RP_JORIG, 0) = DLookup("[journalno_orig]", "zz_tbl_RePost_Journal", "[ID] = " & CStr(rstJournal![ID]))
4080                      garr_varRePost(RP_ID, 0) = rstJournal![ID]
4090                      garr_varRePost(RP_ACTNO, 0) = rstJournal![accountno]
4100                      garr_varRePost(RP_ASTNO, 0) = rstJournal![assetno]
4110                      garr_varRePost(RP_ASTDAT, 0) = datAssetDate
4120                      garr_varRePost(RP_TRNDAT, 0) = datTransDate
4130                      garr_varRePost(RP_RSTNAM, 0) = rstJournal.Name
4140                      garr_varRePost(RP_LNGX, 0) = lngX
4150                      garr_varRePost(RP_RECS1, 0) = lngRecs1
4160                      If gblnNoErrHandle_Repost = True Then
4170  On Error GoTo 0
4180                      Else
4190  On Error GoTo ERRH
4200                      End If
4210                    End If
      #End If
4220                  End If
4230                End Select
4240                Select Case IsNull(![transdate])
                    Case True
4250                  strErr = "Unable to process record. Asset = " & strAssetNo & _
                        ", Account = " & strAccountNo & ": Transaction Date required. Rolling back changes!  Ref#6"
4260                  blnRollback = True
4270                  intRollbackPoint = 6
                      'DOES THIS ALSO ROLL BACK CHANGES MADE TO ACCOUNT, ABOVE?
4280                Case False
4290                  datTransDate = ![transdate]
4300                  If CLng(datTransDate) < lngAssetDateMin Then
4310                    strErr = "Unable to process record. Asset = " & strAssetNo & _
                          ", Account = " & strAccountNo & ": Transaction Date " & Format(![transdate], "mm/dd/yyyy") & _
                          " appears inappropriate. Rolling back changes!  Ref#5"
4320                    blnRollback = True
4330                    intRollbackPoint = 5
4340                  End If
4350                End Select
4360                If blnRollback = False Then
4370                  With rstActiveAssets
4380                    .AddNew
4390                    ![assetno] = rstJournal![assetno]                 ' ** dbLong (Required)
4400                    ![accountno] = rstJournal![accountno]             ' ** dbText (Required)
4410                    ![assetdate] = datAssetDate                       ' ** dbDate
4420                    ![transdate] = datTransDate                       ' ** dbDate (Required)
4430                    ![postdate] = Null                                ' ** dbDate
4440                    ![shareface] = Nz(rstJournal![shareface], 0#)     ' ** dbDouble
4450                    ![due] = rstJournal![due]                         ' ** dbDate
4460                    ![rate] = Nz(rstJournal![rate], 0#)               ' ** dbDouble
                        ' ** VGC 12/28/2009: What if I put in the calc for just this one when it's first entered?
                        ' ** It could still get updated later.
4470                    If Nz(rstJournal![shareface], 0#) > 0# And Nz(rstJournal![Cost], 0@) > 0@ Then
                          ' ** VGC 10/24/2014: This will be updated below.
4480                      ![averagepriceperunit] = Abs(rstJournal![Cost] / rstJournal![shareface])  ' ** dbDouble
4490                      If Nz(rstJournal![pershare], 0#) = 0# Then
4500                        ![priceperunit] = Abs(rstJournal![Cost] / rstJournal![shareface])       ' ** dbDouble
4510                      Else
4520                        ![priceperunit] = Abs(Nz(rstJournal![pershare], 0#))  ' ** dbDouble
4530                      End If
4540                    Else
4550                      ![averagepriceperunit] = 0#                      ' ** dbDouble
4560                      ![priceperunit] = Abs(Nz(rstJournal![pershare], 0#))  ' ** dbDouble
4570                    End If
4580                    ![ICash] = Nz(rstJournal![ICash], 0@)             ' ** dbCurrency/dbDouble
4590                    ![PCash] = Nz(rstJournal![PCash], 0@)             ' ** dbCurrency/dbDouble
4600                    ![Cost] = Nz(rstJournal![Cost], 0@)               ' ** dbCurrency/dbDouble
4610                    If IsNull(rstJournal![description]) = False Then  ' ** journal:      dbText(200)
4620                      If Len(Trim(rstJournal![description])) > 0 Then '** ActiveAssets: dbText(150)
4630                        ![description] = Left(Trim(rstJournal![description]), 150)
4640                      End If
4650                    End If
4660                    ![posted] = datPosted                             ' ** dbDate
                        ' ** VGC 10/24/2014: IsAverage now set when asset chosen.
4670                    ![IsAverage] = rstJournal![IsAverage]             ' ** dbBoolean
4680                    ![Location_ID] = rstJournal![Location_ID]         ' ** dbLong
4690                    ![curr_id] = rstJournal![curr_id]                 ' ** dbLong
                        'SHOULD ANY ATTEMPT BE MADE TO CROSS-CHECK TRANSDATE WITH PRICING HISTORY?
4700                    Select Case blnHasForExJ
                        Case True
4710                      ![cost_usd] = rstJournal![cost_usd]
4720                      If rstJournal![curr_id] = 150& Then
4730                        ![market_usd] = CCur(Round((rstJournal![shareface] * rstJournal![marketvaluecurrent]), 2))
4740                      Else
4750                        ![market_usd] = CCur(Round(((rstJournal![shareface] * rstJournal![marketvaluecurrent]) * rstJournal![curr_rate2]), 2))
4760                      End If
4770                    Case False
4780                      ![cost_usd] = rstJournal![Cost]
4790                      ![market_usd] = CCur(Round((rstJournal![shareface] * rstJournal![marketvaluecurrent]), 2))
4800                    End Select
      #If HasRepost Then
4810                    Select Case gblnRePost
                        Case True
4820  On Error Resume Next
4830                      .Update
4840                      If ERR.Number <> 0 Then
4850                        If ERR.Number = 3022 Then  ' ** Duplicate.
                              ' ** If this is a stock split, move it to the next posting.
4860                          If gblnNoErrHandle_Repost = True Then
4870  On Error GoTo 0
4880                          Else
4890  On Error GoTo ERRH
4900                          End If
4910                          glngRePost_MoveJIDs = glngRePost_MoveJIDs + 1&
4920                          ReDim Preserve garr_varRePost_MoveJID(MJ_ELEMS, (glngRePost_MoveJIDs - 1&))
4930                          garr_varRePost_MoveJID(MJ_JID, (glngRePost_MoveJIDs - 1&)) = rstJournal![ID]
4940                        Else
4950                          blnRollback = True
4960                          Set rstErrLog = dbs.OpenRecordset("tblErrorLog", dbOpenDynaset, dbConsistent)
4970                          MsgBox "An error has occurred while adding a new Tax Lot." & vbCrLf & vbCrLf & _
                                "Error: " & CStr(ERR.Number) & vbCrLf & "Description: " & ERR.description & vbCrLf & _
                                "Module: " & THIS_NAME & vbCrLf & "Procedure: " & THIS_PROC & vbCrLf & vbCrLf & _
                                "Contact Delta Data, Inc., with this information.", vbCritical + vbOKOnly, "Error: " & CStr(ERR.Number)
4980                          blnRetVal = zErrorWriteRecord(ERR.Number, ERR.description, THIS_NAME, THIS_PROC, Erl, rstErrLog)  ' ** Module Function: modErrorHandler.
4990                          rstErrLog.Close
5000                          If gblnNoErrHandle_Repost = True Then
5010  On Error GoTo 0
5020                          Else
5030  On Error GoTo ERRH
5040                          End If
5050                        End If
5060                      Else
5070                        If gblnNoErrHandle_Repost = True Then
5080  On Error GoTo 0
5090                        Else
5100  On Error GoTo ERRH
5110                        End If
5120                      End If
5130                    Case False
5140                      .Update
5150                    End Select
      #Else
5160                    .Update
      #End If
5170                  End With
5180                End If  ' ** blnRollback.
5190                If blnRollback = True Then Exit For
5200                If lngX < lngRecs1 Then .MoveNext
5210              Next  ' ** lngX.
5220            End If
5230          End With  ' ** rstJournal.
              ' ** rstActiveAssets now has new records.
              ' ** Fields mis-assigned:
              ' **   [journal].[posted] dbBoolean    ' ** Treated as -1, so date conversion gives 12/30/1899,
              ' **   [ActiveAssets].[posted] dbDate  ' ** because zero in Windows/Access is 12/31/1899.
              ' **   Should be:
              ' **   Now() Or Date() to [ActiveAssets].[posted]
              ' ** Fields not used; no value appended:
              ' **   [ActiveAssets].[postdate] dbDate  Is it used anywhere?
              ' **   [ActiveAssets].[averagepriceperunit] dbDouble
              ' ** Fields with differing attributes:
              ' **   [journal].[description] dbText (200)
              ' **   [ActiveAssets].[description] dbText (150)

5240        End If  ' ** blnRollback.
            ' ** We're still within BeginTrans.
            ' ** Both rstJournal and rstActiveAssets remain open.

5250        If blnRollback = False Then

5260          rstJournal.Close
5270          rstActiveAssets.Close
5280          Set rstJournal = Nothing
5290          Set rstActiveAssets = Nothing
5300          Set qdf1 = Nothing
5310          DoEvents

              ' ***************************************
              ' ***************************************
              ' ** Step 7. Update ActiveAssets.
              ' ***************************************
              ' ***************************************
5320          ctlPostMsg.Caption = "7 of 15 - Update ActiveAssets"
5330          DoEvents

              ' ** 3. Update ActiveAssets amounts.
              ' **    Withdrawals, Sales, and Cost Adjustments, as well as negative Liabilities: pcash < 0.
      #If HasRepost Then
5340          Select Case gblnRePost
              Case True
5350            Set rstActiveAssets = dbs.OpenRecordset("zz_tbl_RePost_ActiveAssets", dbOpenDynaset)
5360            Set qdf1 = dbs.QueryDefs("zz_qry_RePost_ActiveAssets_02_" & strPostType)  ' ** journal table entries to use for updating ActiveAssets table.
5370          Case False
5380            Set rstActiveAssets = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)
5390            Select Case blnHasForExJ
                Case True
5400              Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_02_5_" & strPostType)
5410            Case False
5420              Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_02_" & strPostType)  ' ** journal table entries to use for updating ActiveAssets table.
5430            End Select
5440          End Select
      #Else
5450          Set rstActiveAssets = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)
5460          Select Case blnHasForExJ
              Case True
5470            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_02_5_" & strPostType)
5480          Case False
5490            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_02_" & strPostType)  ' ** journal table entries to use for updating ActiveAssets table.
5500          End Select
      #End If
              ' ** Liability Note: As above, added this qualification:
              ' ** IIf([journaltype]<>'Liability',0,IIf([icash]=0 And [pcash]=0,IIf([cost]>0,-1,0),-1)) = -1
5510          If strPostType = POST_USR Then
5520            With qdf1.Parameters
5530              ![usr] = strUser
5540            End With
5550          End If
5560          Set rstJournal = qdf1.OpenRecordset()
5570          With rstJournal
5580            If .BOF = True And .EOF = True Then
                  ' ** Wouldn't have gotten this far if there were none.
5590            Else
5600              .MoveLast
5610              lngRecs1 = .RecordCount
5620              .MoveFirst
5630              For lngX = 1& To lngRecs1

5640                If IsNull(![PurchaseDate]) = False Then

5650                  datPurchDate = ![PurchaseDate]
5660                  lngDate_Year = DatePart("yyyy", ![PurchaseDate])
5670                  lngDate_Month = DatePart("m", ![PurchaseDate])
5680                  lngDate_Day = DatePart("d", ![PurchaseDate])
5690                  lngDate_Hour = DatePart("h", ![PurchaseDate])
5700                  lngDate_Minute = DatePart("n", ![PurchaseDate])
5710                  lngDate_Second = DatePart("s", ![PurchaseDate])
                      ' ** Make sure the PurchaseDate matches exactly.
5720                  strFindFirst = "[assetno] = " & CStr(![assetno]) & " AND [accountno] = '" & ![accountno] & "'" & _
                        " AND DatePart('yyyy', [assetdate]) = " & CStr(lngDate_Year) & _
                        " AND DatePart('m', [assetdate]) = " & CStr(lngDate_Month) & _
                        " AND DatePart('d', [assetdate]) = " & CStr(lngDate_Day) & _
                        " AND DatePart('h', [assetdate]) = " & CStr(lngDate_Hour) & _
                        " AND DatePart('n', [assetdate]) = " & CStr(lngDate_Minute) & _
                        " AND DatePart('s', [assetdate]) = " & CStr(lngDate_Second)
                      '"' AND [assetdate] = #" & datPurchDate & "#"
                      'rstJournal![assetno]    ' ** dbDouble (Required)
                      'rstJournal![accountno]  ' ** dbText (Required)
                      'rstJournal![assetdate]  ' ** dbDate
                      'rstJournal![shareface]  ' ** dbDouble
                      'rstJournal![cost]       ' ** dbCurrency
                      ' ************************************************************************************************
                      ' ** DatePart Function:
                      ' **
                      ' **   Syntax:
                      ' **     DatePart(interval, date[,firstdayofweek[, firstweekofyear]])
                      ' **
                      ' **   The DatePart function syntax has these named arguments:
                      ' **     Part             Description
                      ' **     ===============  ======================================================================
                      ' **     interval         Required. String expression that is the interval of time you want to
                      ' **                      return.
                      ' **     date             Required. Variant (Date) value that you want to evaluate.
                      ' **     firstdayofweek   Optional. A constant that specifies the first day of the week.
                      ' **                      If not specified, Sunday is assumed.
                      ' **     firstweekofyear  Optional. A constant that specifies the first week of the year.
                      ' **                      If not specified, the first week is assumed to be the week in which
                      ' **                      January 1 occurs.
                      ' **
                      ' **     Return           A Variant (Integer) containing the specified part of a given date.
                      ' **
                      ' ************************************************************************************************
5730                  With rstActiveAssets
5740                    .FindFirst strFindFirst
5750                    Select Case .NoMatch
                        Case True
                          ' ** Shouldn't happen!
5760                    Case False
5770                      lngCurrID = ![curr_id]
5780                      dblTmp03 = Nz(![shareface], 0#)
5790                      dblTmp04 = Nz(![priceperunit], 0#)
5800                      If lngCurrID = 150& Then
5810                        curTmp06 = Nz(![Cost], 0@)
5820                      Else
5830                        dblTmp05 = Nz(![Cost], 0#)
5840                      End If
5850                      curTmp07 = Nz(![cost_usd], 0@)
5860                      curTmp08 = Nz(![market_usd], 0@)
5870                      .Edit  ' ** This one is Plus/Minus because it's Sold/Withdrawn/Cost Adj./-Liability.
                          'VGC 10/23/2014: WHAT ABOUT A COST ADJ. ON AN AVERAGED ASSET?
                          'I THINK IT MIGHT BE HANDLED IN 6. BELOW!
                          'NO, IT ISN'T!
5880                      If rstJournal![journaltype] = "Cost Adj." Then
5890                        If (dblTmp03 - rstJournal![shareface]) = 0# Then
5900                          ![priceperunit] = 0#
5910                        Else
                              ' ** VGC 10/24/2014: Average should change averagepriceperunit.
                              ' ** rstJournal![shareface] should be Zero.
5920                          Select Case rstJournal![IsAverage]
                              Case True
                                'VALUES ARE NOT READJUSTED BELOW!
5930                            If lngCurrID = 150& Then
5940                              ![averagepriceperunit] = Abs((curTmp06 + rstJournal![Cost]) / (dblTmp03 - rstJournal![shareface]))
                                  'DO I WANT TO CHANGE PRICE PER UNIT BASED ON ORIGINAL priceperunit?
5950                              ![priceperunit] = Abs(((![shareface] * ![priceperunit]) + rstJournal![Cost]) / (dblTmp03 - rstJournal![shareface]))
5960                            Else
5970                              ![averagepriceperunit] = Abs((dblTmp05 + rstJournal![Cost]) / (dblTmp03 - rstJournal![shareface]))
5980                              ![priceperunit] = Abs(((![shareface] * ![priceperunit]) + rstJournal![Cost]) / (dblTmp03 - rstJournal![shareface]))
5990                            End If
6000                          Case False
6010                            If lngCurrID = 150& Then
6020                              ![priceperunit] = Abs((curTmp06 + rstJournal![Cost]) / (dblTmp03 - rstJournal![shareface]))
6030                            Else
6040                              ![priceperunit] = Abs((dblTmp05 + rstJournal![Cost]) / (dblTmp03 - rstJournal![shareface]))
6050                            End If
6060                          End Select
6070                        End If
6080                      End If
6090                      ![shareface] = dblTmp03 - rstJournal![shareface]
                          'SHOULD ANY ATTEMPT BE MADE TO CROSS-CHECK TRANSDATE WITH PRICING HISTORY?
6100                      If lngCurrID = 150& Then
6110                        ![Cost] = curTmp06 + rstJournal![Cost]
6120                        ![cost_usd] = curTmp07 + rstJournal![Cost]
6130                        ![market_usd] = curTmp08 + CCur(Round((rstJournal![shareface] * rstJournal![marketvaluecurrent]), 2))
6140                      Else
6150                        ![Cost] = dblTmp05 + rstJournal![Cost]
6160                        ![cost_usd] = curTmp07 + rstJournal![cost_usd]
6170                        dblTmp05 = Round(((rstJournal![shareface] * rstJournal![marketvaluecurrent]) * rstJournal![curr_rate2]), 2)
6180                        ![market_usd] = curTmp08 + dblTmp05
6190                      End If
6200                      .Update
6210                    End Select
6220                  End With  ' ** rstActiveAssets.

6230                End If  ' ** [PurchaseDate].

6240                If blnRollback = True Then Exit For
                    ' #### rstActiveAssets changed, rstAccount changed, rstJournal not changed ####
6250                If lngX < lngRecs1 Then .MoveNext
6260              Next  ' ** lngX.
6270            End If  ' ** BOF, EOF.
6280          End With  ' ** rstJournal.
              ' ** [ActiveAssets] table now has new and/or updated records.

6290        End If  ' ** blnRollback.
            ' ** Both rstJournal and rstActiveAssets remain open.
            ' ** We're still within BeginTrans.

            ' ** If there are averaged Cost Adj's, those values must now be re-adjusted.
6300        If blnRollback = False Then

6310          Select Case blnHasForExJ
              Case True
                ' ** qryPost_ActiveAssets_07_5_1 (qryPost_ActiveAssets_07_4_1 (Journal, just 'Cost Adj.',
                ' ** with asset_trans_date; for IsAverage = True), linked to qryPost_ActiveAssets_00_07
                ' ** (qryPost_ActiveAssets_00_06 (qryPost_ActiveAssets_00_05 (Journal, linked to
                ' ** qryPost_ActiveAssets_00_04 (qryPost_ActiveAssets_00_03 (Union of qryPost_ActiveAssets_00_01
                ' ** (tblCurrency_History, just needed fields), qryPost_ActiveAssets_00_02 (tblCurrency,
                ' ** just needed fields)), grouped by curr_id, curr_date, with Max(currhist_id)), grouped by curr_id,
                ' ** with Max(curr_date)), linked back to qryPost_ActiveAssets_00_04 (qryPost_ActiveAssets_00_03
                ' ** (Union of qryPost_ActiveAssets_00_01 (tblCurrency_History, just needed fields),
                ' ** qryPost_ActiveAssets_00_02 (tblCurrency, just needed fields)), grouped by curr_id, curr_date,
                ' ** with Max(currhist_id)), with currhist_id), linked to tblCurrency_History, tblCurrency),
                ' ** with .._usd fields), grouped and summed, by accountno, assetno.
6320            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_08_4_" & strPostType)
6330          Case False
                ' ** qryPost_ActiveAssets_07_1 (Journal, just 'Cost Adj.', for
                ' ** IsAverage = True), grouped and summed, by accountno, assetno.
6340            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_08_" & strPostType)
6350          End Select
6360          If strPostType = POST_USR Then
6370            With qdf1.Parameters
6380              ![usr] = strUser
6390            End With
6400          End If
6410          Set rst1 = qdf1.OpenRecordset
6420          With rst1
6430            If .BOF = True And .EOF = True Then
                  ' ** No averaged Cost Adj's.
6440            Else
6450              .MoveLast
6460              lngRecs1 = .RecordCount
6470              .MoveFirst
6480              For lngX = 1& To lngRecs1
6490                Select Case blnHasForExJ
                    Case True
                      ' ** qryPost_ActiveAssets_09_2 (qryPost_ActiveAssets_09_1 (ActiveAssets, linked to MasterAsset,
                      ' ** with asset_trans_date, by  specified [actno], [astno]), linked to qryPost_ActiveAssets_00_07
                      ' ** (qryPost_ActiveAssets_00_06 (qryPost_ActiveAssets_00_05 (Journal, linked to
                      ' ** qryPost_ActiveAssets_00_04 (qryPost_ActiveAssets_00_03 (Union of qryPost_ActiveAssets_00_01
                      ' ** (tblCurrency_History, just needed fields), qryPost_ActiveAssets_00_02 (tblCurrency, just
                      ' ** needed fields)), grouped by curr_id, curr_date, with Max(currhist_id)), grouped by curr_id,
                      ' ** with Max(curr_date)), linked back to qryPost_ActiveAssets_00_04 (qryPost_ActiveAssets_00_03
                      ' ** (Union of qryPost_ActiveAssets_00_01 (tblCurrency_History, just needed fields),
                      ' **  qryPost_ActiveAssets_00_02 (tblCurrency, just needed fields)), grouped by curr_id,
                      ' ** curr_date, with Max(currhist_id)), with currhist_id), linked to tblCurrency_History,
                      ' ** tblCurrency), grouped and summed, by accountno, assetno.  #curr_id
6500                  Set qdf2 = dbs.QueryDefs("qryPost_ActiveAssets_10_1")
6510                Case False
                      ' ** For each accountno/assetno pair in Journal that have an averaged Cost Adj.
                      ' ** qryPost_ActiveAssets_09 (ActiveAssets, by  specified [actno], [astno]),
                      ' ** grouped and summed, by accountno, assetno.  #curr_id
6520                  Set qdf2 = dbs.QueryDefs("qryPost_ActiveAssets_10")
6530                End Select
6540                With qdf2.Parameters
6550                  ![actno] = rst1![accountno]
6560                  ![astno] = rst1![assetno]
6570                End With
6580                Set rst2 = qdf2.OpenRecordset
6590                With rst2
6600                  .MoveFirst               ' ** Just 1 totals record.
6610                  dblTmp03 = ![shareface]  ' ** Total shareface.
6620                  Select Case blnHasForExJ
                      Case True
6630                    If lngCurrID = 150& Then
6640                      curTmp06 = ![Cost]       ' ** Total cost, already adjusted above.
6650                      curTmp07 = ![cost_usd]
6660                      curTmp08 = ![market_usd]
6670                      dblTmp09 = ![marketvaluecurrent]
6680                    Else
6690                      dblTmp05 = ![Cost]
6700                      curTmp07 = ![cost_usd]
6710                      curTmp08 = ![market_usd]
6720                      dblTmp09 = ![marketvaluecurrent]
6730                    End If
6740                  Case False
6750                    curTmp06 = ![Cost]
6760                  End Select
6770                  .Close
6780                End With  ' ** rst2.
6790                Set rst2 = Nothing
6800                Set qdf2 = Nothing
6810                DoEvents
6820                Select Case blnHasForExJ
                    Case True
6830                  If lngCurrID = 150& Then
6840                    dblTmp04 = Abs(curTmp06 / dblTmp03)  ' ** Overall averagepriceperunit.
6850                  Else
6860                    dblTmp04 = Abs(dblTmp05 / dblTmp03)
6870                  End If
6880                Case False
6890                  dblTmp04 = Abs(curTmp06 / dblTmp03)
6900                End Select
6910                Select Case blnHasForExJ
                    Case True
                      ' ** qryPost_ActiveAssets_09_1 (ActiveAssets, linked to MasterAsset, with asset_trans_date,
                      ' ** by  specified [actno], [astno]), linked to qryPost_ActiveAssets_00_07 (qryPost_ActiveAssets_00_06
                      ' ** (qryPost_ActiveAssets_00_05 (Journal, linked to qryPost_ActiveAssets_00_04 (qryPost_ActiveAssets_00_03
                      ' ** (Union of qryPost_ActiveAssets_00_01 (tblCurrency_History, just needed fields),
                      ' ** qryPost_ActiveAssets_00_02 (tblCurrency, just needed fields)), grouped by curr_id, curr_date, with
                      ' ** Max(currhist_id)), grouped by curr_id, with Max(curr_date)), linked back to qryPost_ActiveAssets_00_04
                      ' ** (qryPost_ActiveAssets_00_03 (Union of qryPost_ActiveAssets_00_01 (tblCurrency_History, just needed fields),
                      ' ** qryPost_ActiveAssets_00_02 (tblCurrency, just needed fields)), grouped by curr_id, curr_date,
                      ' ** with Max(currhist_id)), with currhist_id), linked to tblCurrency_History, tblCurrency.  #curr_id
6920                  Set qdf2 = dbs.QueryDefs("qryPost_ActiveAssets_09_2")
6930                Case False
                      ' ** ActiveAssets, by  specified [actno], [astno].  #curr_id
6940                  Set qdf2 = dbs.QueryDefs("qryPost_ActiveAssets_09")
6950                End Select
6960                With qdf2.Parameters
6970                  ![actno] = rst1![accountno]
6980                  ![astno] = rst1![assetno]
6990                End With
7000                Set rst2 = qdf2.OpenRecordset
7010                With rst2
7020                  .MoveLast
7030                  lngRecs2 = .RecordCount
7040                  .MoveFirst
                      'IF IT'S A DEPOSIT OR PURCHASE, THE ASSETDATE WILL GIVE THAT DATE'S CURR_RATE,
                      'IF IT'S A WITHDRAWAL OR SOLD, PURCHASE DATE IS THE ORIGINAL, BUT ASSETDATE
                      'WILL BE THE DATE OF SALE, AND SO CURR_RATE WILL BE FOR THAT DAY!
7050                  For lngY = 1& To lngRecs2
                        ' ** For each Tax Lot held by this account.
7060                    .Edit
7070                    ![averagepriceperunit] = dblTmp04
                        ' ** We can leave priceperunit alone.
7080                    Select Case blnHasForExJ
                        Case True
7090                      If lngCurrID = 150& Then
7100                        ![Cost] = CCur(Round((![shareface] * dblTmp04), 2))
7110                        ![cost_usd] = CCur(Round((![shareface] * dblTmp04), 2))
7120                        ![market_usd] = CCur(Round((![shareface] * dblTmp09), 2))
7130                      Else
7140                        ![Cost] = CDbl(Round((![shareface] * dblTmp04), 2))
7150                        ![cost_usd] = CCur(Round(((![shareface] * dblTmp04) * ![curr_rate2]), 2))
7160                        ![market_usd] = CCur(Round(((![shareface] * dblTmp09) * ![curr_rate2]), 2))
7170                      End If
7180                    Case False
7190                      ![Cost] = CCur(Round((![shareface] * dblTmp04), 2))
7200                      ![cost_usd] = CCur(Round((![shareface] * dblTmp04), 2))
7210                      ![market_usd] = CCur(Round((![shareface] * dblTmp09), 2))
7220                    End Select
7230                    .Update
7240                    If lngY < lngRecs2 Then .MoveNext
7250                  Next  ' ** lngY.
7260                  .Close
7270                End With  ' ** rst2.
7280                Set rst2 = Nothing
7290                Set qdf2 = Nothing
7300                DoEvents
7310                If lngX < lngRecs1 Then .MoveNext  ' ** Next accountno/assetno pair.
7320              Next  ' ** lngX.
7330            End If  ' ** BOF, EOF.
7340            .Close
7350          End With  ' ** rst1.
7360          Set rst1 = Nothing
7370          Set qdf1 = Nothing
7380          DoEvents

7390        End If  ' ** blnRollback.
            ' ** Both rstJournal and rstActiveAssets remain open.
            ' ** We're still within BeginTrans.

7400        If blnRollback = False Then

7410          rstJournal.Close
7420          rstActiveAssets.Close
7430          Set rstJournal = Nothing
7440          Set rstActiveAssets = Nothing
7450          DoEvents

              ' ***************************************
              ' ***************************************
              ' ** Step 8. Adjust ActiveAssets.
              ' ***************************************
              ' ***************************************
7460          ctlPostMsg.Caption = "8 of 15 - Adjust ActiveAssets"
7470          DoEvents

              ' ** 4. Adjust ActiveAssets amounts.
              ' **    Only updates Journal entries that have a PurchaseDate!
              ' **    PurchaseDate is added only under these 3 circumstances, otherwise it remains blank:
              ' **    1. PurchaseDate field populated when frmJournal_Sub4_Sold Distribute button is pressed.
              ' **    2. PurchaseDate field populated when frmLotInformation OK button is pressed.
              ' **    3. PurchaseDate field populated when frmSweeper OK button is pressed.
      #If HasRepost Then
7480          Select Case gblnRePost
              Case True
7490            Set rstActiveAssets = dbs.OpenRecordset("zz_tbl_RePost_ActiveAssets", dbOpenDynaset)
7500            Set qdf1 = dbs.QueryDefs("zz_qry_RePost_ActiveAssets_03_" & strPostType)
7510          Case False
7520            Set rstActiveAssets = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)
7530            Select Case blnHasForExJ
                Case True
7540              Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_03_5_" & strPostType)
7550            Case False
7560              Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_03_" & strPostType)
7570            End Select
7580          End Select
      #Else
7590          Set rstActiveAssets = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)
              ' ** This does not include Cost Adj.
7600          Select Case blnHasForExJ
              Case True
7610            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_03_5_" & strPostType)
7620          Case False
7630            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_03_" & strPostType)
7640          End Select
      #End If
              ' ** Liability Note: As above, added this qualification:
              ' ** [journaltype] = 'Liability', [PurchaseDate] Is Not Null, [cost] > 0, [pcash] = 0, [icash] = 0
7650          If strPostType = POST_USR Then
7660            With qdf1.Parameters
7670              ![usr] = strUser
7680            End With
7690          End If
7700          Set rstJournal = qdf1.OpenRecordset()  ' ** Assets to append from journal table to ActiveAssets table.
7710          With rstJournal
7720            If .BOF = True And .EOF = True Then
                  ' ** Wouldn't have gotten this far if there were none.
7730            Else
7740              .MoveLast
7750              lngRecs1 = .RecordCount
7760              .MoveFirst
7770              For lngX = 1& To lngRecs1
7780                datPurchDate = ![PurchaseDate]  ' ** VGC 05/06/08: Switch from assetdate to purchaseDate!
7790                strFindFirst = "[assetno] = " & CStr(![assetno]) & " AND [accountno] = '" & _
                      ![accountno] & "' AND Format([assetdate],'mm/dd/yyyy hh:nn:ss') = '" & Format(datPurchDate, "mm/dd/yyyy hh:nn:ss") & "'"
7800                With rstActiveAssets
7810                  .FindFirst strFindFirst
7820                  Select Case .NoMatch
                      Case True
7830                    strErr = "Unable to process record. Asset = " & CStr(rstJournal![assetno]) & _
                          ", Account = " & rstJournal![accountno] & ", Assetdate = " & rstJournal![PurchaseDate] & _
                          " was not found in active assets table. Rolling back changes!  Ref#8"
7840                    blnRollback = True
7850                    intRollbackPoint = 8
7860                  Case False
                        ' ** There should only be one record.
7870                    If rstJournal![journaltype] = "Liability" And _
                            rstJournal![ICash] = 0 And rstJournal![PCash] = 0 And rstJournal![Cost] > 0 Then
                          ' ** Liability Note: As above, added this special case.
7880                      dblTmp03 = -Nz(![shareface], 0#)
7890                    Else
7900                      dblTmp03 = Nz(![shareface], 0#)
7910                    End If
7920                    dblTmp04 = Nz(![priceperunit], 0#)
7930                    Select Case blnHasForExJ
                        Case True
7940                      If lngCurrID = 150& Then
7950                        curTmp06 = Nz(![Cost], 0@)
                            ' ** These aren't used below!
                            'curTmp07 = Nz(![cost_usd], 0@)
                            'curTmp08 = Nz(![market_usd], 0@)
7960                      Else
7970                        dblTmp05 = Nz(![Cost], 0@)
                            ' ** These aren't used below!
                            'curTmp07 = Nz(![cost_usd], 0@)
                            'curTmp08 = Nz(![market_usd], 0@)
7980                      End If
7990                    Case False
8000                      curTmp06 = Nz(![Cost], 0@)
                          ' ** These aren't used below!
                          'curTmp07 = Nz(![cost_usd], 0@)
                          'curTmp08 = Nz(![market_usd], 0@)
8010                    End Select
8020                    .Edit  ' ** This one is Plus/Plus because it's Purchase/Deposit/+Liability.
8030                    ![shareface] = dblTmp03 + rstJournal![shareface]
8040                    If (dblTmp03 + rstJournal![shareface]) = 0# Then
8050                      ![priceperunit] = 0#
8060                    Else
8070                      Select Case blnHasForExJ
                          Case True
8080                        If lngCurrID = 150& Then
8090                          ![priceperunit] = Abs(curTmp06 + rstJournal![Cost]) / (dblTmp03 + rstJournal![shareface])
8100                        Else
8110                          ![priceperunit] = Abs(dblTmp05 + rstJournal![Cost]) / (dblTmp03 + rstJournal![shareface])
8120                        End If
8130                      Case False
8140                        ![priceperunit] = Abs(curTmp06 + rstJournal![Cost]) / (dblTmp03 + rstJournal![shareface])
8150                      End Select
8160                    End If
8170                    Select Case blnHasForExJ
                        Case True
8180                      If lngCurrID = 150& Then
8190                        ![Cost] = curTmp06 + rstJournal![Cost]
8200                        ![cost_usd] = curTmp06 + rstJournal![Cost]
8210                        ![market_usd] = CCur(Round((![shareface] * rstJournal![marketvaluecurrent]), 2))
8220                      Else
8230                        ![Cost] = dblTmp05 + rstJournal![Cost]
8240                        ![cost_usd] = CCur(Round(((dblTmp05 + rstJournal![Cost]) * rstJournal![curr_rate2]), 2))
8250                        ![market_usd] = CCur(Round(((![shareface] * rstJournal![marketvaluecurrent]) * rstJournal![curr_rate2]), 2))
8260                      End If
8270                    Case False
8280                      ![Cost] = curTmp06 + rstJournal![Cost]
8290                      ![cost_usd] = curTmp06 + rstJournal![Cost]
8300                      ![market_usd] = CCur(Round((![shareface] * rstJournal![marketvaluecurrent]), 2))
8310                    End Select
8320                    .Update
8330                  End Select
8340                End With  ' ** rstActiveAssets.
8350                If blnRollback = True Then Exit For
                    ' #### rstActiveAssets changed, rstAccount changed, rstJournal not changed ####
8360                If lngX < lngRecs1 Then .MoveNext
8370              Next
8380            End If
8390          End With  ' ** rstJournal.
              ' ** rstActiveAssets now has updated records.
              ' ** rstJournal and rstActiveAssets remain open.

      #If HasRepost Then
8400          If gblnRePost = True Then
8410            If glngRePost_MoveJIDs > 0& Then
8420              blnRollback = True
8430              intRollbackPoint = 7
8440              strErr = "Error: 3022" & vbCrLf & "The changes you requested to the table were not successful because " & _
                    "they would create duplicate values in the index, primary key, or relationship."
8450              gstrRePostErrMsg = strErr
8460            End If
8470          End If
      #End If

8480        End If  ' ** blnRollback.
            ' ** We're still within BeginTrans.

8490        If blnRollback = False Then

8500          rstJournal.Close
8510          rstActiveAssets.Close
8520          Set rstJournal = Nothing
8530          Set rstActiveAssets = Nothing
8540          Set qdf1 = Nothing
8550          DoEvents

              ' ***************************************
              ' ***************************************
              ' ** Step 9. Delete ActiveAssets.
              ' ***************************************
              ' ***************************************
8560          ctlPostMsg.Caption = "9 of 15 - Delete ActiveAssets"
8570          DoEvents

              ' ** 5. Delete ActiveAssets records.
              ' **    ActiveAsset records deleted that have a 0 shareface.
      #If HasRepost Then
8580          Select Case gblnRePost
              Case True
8590            Set rstActiveAssets = dbs.OpenRecordset("zz_tbl_RePost_ActiveAssets", dbOpenDynaset)
8600          Case False
8610            Set rstActiveAssets = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)
8620          End Select
      #Else
8630          Set rstActiveAssets = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)
      #End If
8640          With rstActiveAssets
8650            If .BOF = True And .EOF = True Then
                  ' ** New system; nothing there yet.
8660            Else
8670              .MoveLast
8680              lngRecs1 = .RecordCount
8690              For lngX = lngRecs1 To 1& Step -1&
8700                If ![shareface] < 0.0001 And ![shareface] > -0.0001 Then
8710                  .Delete
8720                End If
8730                If lngX > 1& Then .MovePrevious
8740              Next
8750            End If
8760            .Close
8770          End With
8780          Set rstActiveAssets = Nothing
8790          DoEvents

              ' ***************************************
              ' ***************************************
              ' ** Step 10. Tax Lot Averaging.
              ' ***************************************
              ' ***************************************
8800          ctlPostMsg.Caption = "10 of 15 - Tax Lot Averaging"
8810          DoEvents

              ' ** 6. Update ActiveAssets average price per unit.
              ' ** WHY DOES IT APPEAR TO UPDATING EVERYBODY'S AVERAGEPRICEPERUNIT, NOT JUST THE NEW POSTING?
              ' ** How 'bout this:
              ' ** If this asset has ANY IsAverage = True, then they all should.
              ' ** Find only ActiveAssets for this accountno and assetno, and update just them.
              ' ** VGC 10/24/2014: A new purchase goes in above, with its original cost.
              ' ** Then, this section will update all based on total cost, including the original
              ' ** added above, divided by the new total shares. It will only do this for a
              ' ** Purchase of an already averaged asset. Totals updated in the Account table aren't affected.
      #If HasRepost Then
8820          Select Case gblnRePost
              Case True
                ' #### TO COME ####
                'Set qdf1 = dbs.QueryDefs("zz_qry_RePost_ActiveAssets_06_" & strPostType)
8830          Case False
                ' ** Journal, grouped by accountno, assetno, for 'Deposit', 'Purchase', 'Sold', 'Withdrawn', 'Liability'.
8840            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_06_" & strPostType)
8850          End Select
      #Else
              ' ** Journal, grouped by accountno, assetno, for 'Deposit', 'Purchase', 'Sold', 'Withdrawn', 'Liability'.
8860          Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_06_" & strPostType)
      #End If
8870          If strPostType = POST_USR Then
8880            With qdf1.Parameters
8890              ![usr] = strUser
8900            End With
8910          End If
8920          Set rstJournal = qdf1.OpenRecordset()
8930          With rstJournal
8940            If .BOF = True And .EOF = True Then
                  ' ** No asset pruchases or sales.
8950            Else
8960              .MoveLast
8970              lngRecs1 = .RecordCount
8980              .MoveFirst
8990              For lngX = 1& To lngRecs1
      #If HasRepost Then
9000                Select Case gblnRePost
                    Case True
                      ' ** zz_tbl_RePost_MasterAsset, linked to zz_tbl_RePost_ActiveAssets,
                      ' ** grouped and summed by assetno, with shareface_sum.
9010                  Set qdf2 = dbs.QueryDefs("zz_qry_RePost_ActiveAssets_05")
9020                Case False
9030                  Select Case blnHasForExJ
                      Case True
                        ' ** qryPost_ActiveAssets_04_01 (xx), linked to qryPost_ActiveAssets_04_10 (xx),
                        ' ** grouped and summed, with Min(IsAverage).  #curr_id
9040                    Set qdf2 = dbs.QueryDefs("qryPost_ActiveAssets_04_11")
9050                  Case False
                        ' ** ActiveAssets, grouped and summed, with Min(IsAverage).  #curr_id
9060                    Set qdf2 = dbs.QueryDefs("qryPost_ActiveAssets_04")
9070                  End Select
9080                End Select
      #Else
9090                Select Case blnHasForExJ
                    Case True
                      ' ** qryPost_ActiveAssets_04_01 (xx), linked to qryPost_ActiveAssets_04_10 (xx),
                      ' ** grouped and summed, with Min(IsAverage).  #curr_id
9100                  Set qdf2 = dbs.QueryDefs("qryPost_ActiveAssets_04_11")
9110                Case False
                      ' ** ActiveAssets, grouped and summed, with Min(IsAverage).  #curr_id
9120                  Set qdf2 = dbs.QueryDefs("qryPost_ActiveAssets_04")
9130                End Select
      #End If
9140                Set rstActiveAssets = qdf2.OpenRecordset()
9150                With rstActiveAssets
9160                  .MoveFirst
9170                  strAccountNo = rstJournal![accountno]
9180                  strAssetNo = CStr(rstJournal![assetno])
9190                  .FindFirst "[accountno] = '" & strAccountNo & "' And [assetno] = " & strAssetNo
9200                  If .NoMatch = False Then
                        ' ** VGC 10/24/2014: At least 1 Tax Lot, for this accountno and assetno, must be averaged to hit this.
9210                    If ![IsAverage] = True Then
9220                      If Nz(![shareface], 0#) = 0# Then
9230                        dblTmp03 = 0#
9240                      Else
9250                        dblTmp03 = (Nz(![Cost], 0#) / ![shareface])
9260                      End If
      #If HasRepost Then
9270                      Select Case gblnRePost
                          Case True
9280                        Set rst1 = dbs.OpenRecordset("zz_tbl_RePost_ActiveAssets", dbOpenDynaset)
9290                      Case False
9300                        Set rst1 = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)
9310                      End Select
      #Else
9320                      Set rst1 = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset)
      #End If
9330                      strFindFirst = "[accountno] = '" & strAccountNo & "' AND [assetno] = " & strAssetNo
9340                      rst1.Filter = strFindFirst
9350                      Set rst2 = rst1.OpenRecordset
9360                      With rst2
9370                        .MoveLast
9380                        lngRecs2 = .RecordCount
9390                        .MoveFirst
9400                        For lngZ = 1& To lngRecs2
9410                          .Edit
                              ' ** VGC 10/24/2014: This wasn't changing cost, only averageprice per unit!
9420                          ![averagepriceperunit] = Abs(dblTmp03)
9430                          Select Case blnHasForExJ
                              Case True
9440                            If lngCurrID = 150& Then
9450                              ![Cost] = CCur(Round((![shareface] * dblTmp03), 2))
9460                              ![cost_usd] = CCur(Round((![shareface] * dblTmp03), 2))
9470                              ![market_usd] = CCur(Round((![shareface] * rstActiveAssets![marketvaluecurrent]), 2))
9480                            Else
9490                              ![Cost] = CDbl(Round((![shareface] * dblTmp03), 2))
9500                              ![cost_usd] = CCur(Round(((![shareface] * dblTmp03) * rstActiveAssets![curr_rate2]), 2))
9510                              ![market_usd] = CCur(Round(((![shareface] * _
                                    rstActiveAssets![marketvaluecurrent]) * rstActiveAssets![curr_rate2]), 2))
9520                            End If
9530                          Case False
9540                            ![Cost] = CCur(Round((![shareface] * dblTmp03), 2))
9550                            ![cost_usd] = CCur(Round((![shareface] * dblTmp03), 2))
9560                            ![market_usd] = CCur(Round((![shareface] * rstActiveAssets![marketvaluecurrent]), 2))
9570                          End Select
9580                          If ![IsAverage] = False Then
9590                            ![IsAverage] = True
9600                          End If
9610                          .Update
                              ' ** VGC 10/24/2014: This wasn't advancing through the records!
9620                          If lngZ < lngRecs2 Then .MoveNext
9630                        Next  ' ** lngZ.
9640                        .Close
9650                      End With  ' ** rst2.
9660                      Set rst2 = Nothing
9670                      rst1.Close
9680                      Set rst1 = Nothing
9690                      DoEvents
9700                    End If  ' ** IsAverage.
9710                  End If  ' ** .NoMatch.
9720                  .Close
9730                End With  ' ** rstActiveAssets.
9740                If lngX < lngRecs1 Then .MoveNext
9750              Next  ' ** lngX, each accountno/assetno pair.
9760            End If
9770            .Close
9780          End With  ' ** rstJournal.

9790          Set rstJournal = Nothing
9800          Set rstActiveAssets = Nothing
9810          Set qdf1 = Nothing
9820          Set qdf2 = Nothing
9830          DoEvents

              ' ***************************************
              ' ***************************************
              ' ** Step 11. Update MasterAsset.
              ' ***************************************
              ' ***************************************
9840          ctlPostMsg.Caption = "11 of 15 - Update MasterAsset"
9850          DoEvents

              ' ** 7. Update masterasset records.
              ' **    MasterAsset table updated with ActiveAssets' shareface total.
      #If HasRepost Then
9860          Select Case gblnRePost
              Case True
9870            Set rstMasterAsset = dbs.OpenRecordset("zz_tbl_RePost_MasterAsset", dbOpenDynaset)
                ' ** Append zz_tbl_RePost_ActiveAssets to ActiveAssets.
9880            Set qdf1 = dbs.QueryDefs("zz_qry_RePost_ActiveAssets_06")
9890          Case False
9900            Set rstMasterAsset = dbs.OpenRecordset("masterasset", dbOpenDynaset)
9910            Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_05")
9920          End Select
      #Else
9930          Set rstMasterAsset = dbs.OpenRecordset("masterasset", dbOpenDynaset)
9940          Set qdf1 = dbs.QueryDefs("qryPost_ActiveAssets_05")
      #End If
9950          Set rstActiveAssets = qdf1.OpenRecordset()
9960          With rstActiveAssets
9970            If .BOF = True And .EOF = True Then
                  ' ** Wouldn't have gotten this far if there were none.
9980            Else
9990              .MoveLast
10000             lngRecs1 = .RecordCount
10010             .MoveFirst
10020             For lngX = 1& To lngRecs1
10030               strAssetNo = CStr(![assetno])
10040               dblTmp03 = Nz(![shareface_sum], 0)
10050               strFindFirst = "[assetno] = " & strAssetNo
10060               With rstMasterAsset
10070                 .FindFirst strFindFirst
10080                 If .NoMatch = False Then
10090                   .Edit
10100                   ![shareface] = dblTmp03
10110                   .Update
10120                 End If
10130               End With  ' ** rstMasterAsset.
10140               If lngX < lngRecs1 Then .MoveNext
10150             Next
10160           End If
10170           .Close
10180         End With  ' ** rstActiveAssets.
10190         rstMasterAsset.Close
10200         Set rstActiveAssets = Nothing
10210         Set rstMasterAsset = Nothing
10220         Set qdf1 = Nothing
10230         DoEvents

              ' ***************************************
              ' ***************************************
              ' ** Step 12. Zero-Out MasterAsset.
              ' ***************************************
              ' ***************************************
10240         ctlPostMsg.Caption = "12 of 15 - Zero-Out MasterAsset"
10250         DoEvents

              ' ** 8. Update masterasset records.
              ' **    When no assetno, zero-out shareface field.
      #If HasRepost Then
10260         Select Case gblnRePost
              Case True
10270           Set rst1 = dbs.OpenRecordset("zz_tbl_RePost_MasterAsset", dbOpenDynaset)
10280         Case False
10290           Set rst1 = dbs.OpenRecordset("masterasset", dbOpenDynaset)
10300         End Select
      #Else
10310         Set rst1 = dbs.OpenRecordset("masterasset", dbOpenDynaset)
      #End If
10320         strFindFirst = "[assetno] Is Null"
10330         rst1.Filter = strFindFirst
10340         Set rstMasterAsset = rst1.OpenRecordset()
10350         With rstMasterAsset
10360           If .BOF = True And .EOF = True Then
                  ' ** No Null assetno.
10370           Else
10380             .MoveLast
10390             lngRecs1 = .RecordCount
10400             .MoveFirst
10410             For lngX = 1& To lngRecs1
10420               ![shareface] = 0#
10430               If lngX < lngRecs1 Then .MoveNext
10440             Next
10450           End If
10460           .Close
10470         End With  ' ** rstMasterAsset.
10480         rst1.Close
10490         Set rstMasterAsset = Nothing
10500         Set rst1 = Nothing
10510         DoEvents

10520         lngPosPays = 0&
10530         ReDim arr_varPosPay(P_ELEMS, 0)

              ' ** Get last journalno.
      #If HasRepost Then
10540         Select Case gblnRePost
              Case True
10550           Set qdf1 = dbs.QueryDefs("zz_qry_RePost_Ledger_02")
10560         Case False
10570           Set qdf1 = dbs.QueryDefs("qryPost_Ledger_01")
10580         End Select
      #Else
10590         Set qdf1 = dbs.QueryDefs("qryPost_Ledger_01")
      #End If
10600         Set rstLedger = qdf1.OpenRecordset()
10610         With rstLedger
10620           If .BOF = True And .EOF = True Then
10630             lngMax = 0&
10640           Else
10650             .MoveFirst
10660             Select Case IsNull(![journalno])
                  Case True
10670               lngMax = 0&
10680             Case False
10690               lngMax = ![journalno]
10700             End Select
10710           End If
10720           .Close
10730         End With
10740         Set rstLedger = Nothing
10750         Set qdf1 = Nothing
10760         DoEvents

              ' ***************************************
              ' ***************************************
              ' ** Step 13. New Ledger Entries.
              ' ***************************************
              ' ***************************************
10770         ctlPostMsg.Caption = "13 of 15 - New Ledger Entries"
10780         DoEvents

              ' ** 9. Add new ledger records.
              ' **    Every Journal record is appended.
              ' **    This says assetdate is copied intact from Journal!!!!
              ' **    Assetdate fields in both tables have a short-date input mask! Isn't that wrong?
              ' **    Removed: 12/21/07
      #If HasRepost Then
10790         Select Case gblnRePost
              Case True
10800           Set rstLedger = dbs.OpenRecordset("zz_tbl_RePost_Ledger", dbOpenDynaset, dbAppendOnly)
10810           Set qdf1 = dbs.QueryDefs("zz_qry_RePost_Ledger_03_" & strPostType)
10820         Case False

10830           Set rstLedger = dbs.OpenRecordset("ledger", dbOpenDynaset, dbAppendOnly)
10840           Set qdf1 = dbs.QueryDefs("qryPost_Ledger_02_" & strPostType)

10850           lngImports = 0&
10860           ReDim arr_varImport(I_ELEMS, 0)

10870           Set rst1 = dbs.OpenRecordset("tblJournal_Import", dbOpenDynaset, dbReadOnly)
10880           With rst1
10890             If .BOF = True And .EOF = True Then
                    ' ** None of these are imported.
10900             Else
10910               .MoveLast
10920               lngRecs2 = .RecordCount
10930               .MoveFirst
10940               For lngX = 1& To lngRecs2
10950                 lngImports = lngImports + 1&
10960                 lngE = lngImports - 1&
10970                 ReDim Preserve arr_varImport(I_ELEMS, lngE)
10980                 arr_varImport(I_JID, lngE) = ![Journal_ID]
10990                 arr_varImport(I_ATYP, lngE) = ![assettype]
11000                 arr_varImport(I_CNUM, lngE) = ![CheckNum]
                      ' ** The AssetType field contains the date it was exported from Trust Import: impxl1_dateexport.
                      ' ** The CheckNum field contains the tblImport_Journal_Archive ID: impjarch_id.
                      ' ** The Ledger's PostDate field (currently not used) will contain the export day as the integer
                      ' ** portion of the field, and the archive ID incorporated within the decimal portion.
                      ' ** So, impxl1_dateexport =
                      ' **   Format([postdate], "mm/dd/yyyy")
                      ' ** And, impjarch_id =
                      ' **   IIf(IsNull([postdate])=True,Null,IIf(InStr(CStr(CDbl([postdate])),'.')=0,Null,
                      ' **     CLng(Left(Mid(Mid(CStr(CDbl([postdate])),(InStr(CStr(CDbl([postdate])),'.')+1)),4) &
                      ' **     String(6,'0'),Val(Left(Mid(CStr(CDbl([postdate])),(InStr(CStr(CDbl([postdate])),'.')+1)),2))))))
                      ' ** The 1st character of AssetType gets us within 255 days of the export date,
                      ' **   and the 2nd character gets us to the exact day.
                      ' ** For the archive ID, start with just the decimal portion of the PostDate field.
                      ' **   The 1st position (10th) is always 0.
                      ' **   The 2nd position (100th) tells us the length of the archive id (not likely to get into the millions).
                      ' **   The 3rd position (1000th) is always 0.
                      ' **   The 4th position and beyond (10000th) is the archive ID without any right-most 0's, if present.
                      ' **   So, .0103 = 3, .0207 = 70, .050123 = 12300, .0409876 = 9876.
11010                 dblTmp04 = CDbl(CStr(CLng(Asc(Left(![assettype], 1)) * 255#) + CLng(Asc(Right(![assettype], 1)))) & _
                        ".0" & CStr(Len(CStr(![CheckNum]))) & "0" & CStr(![CheckNum]))
11020                 arr_varImport(I_PDAT, lngE) = CDate(dblTmp04)
11030                 If lngX < lngRecs2 Then .MoveNext
11040               Next
11050             End If
11060             .Close
11070           End With  ' ** rst1.
11080           Set rst1 = Nothing
11090           DoEvents

11100         End Select
              ' ** rstLedger remains open.

      #Else

11110         Set rstLedger = dbs.OpenRecordset("ledger", dbOpenDynaset, dbAppendOnly)
11120         Set qdf1 = dbs.QueryDefs("qryPost_Ledger_02_" & strPostType)

11130         lngImports = 0&
11140         ReDim arr_varImport(I_ELEMS, 0)

11150         Set rst1 = dbs.OpenRecordset("tblJournal_Import", dbOpenDynaset, dbReadOnly)
11160         With rst1
11170           If .BOF = True And .EOF = True Then
                  ' ** None of these are imported.
11180           Else
11190             .MoveLast
11200             lngRecs2 = .RecordCount
11210             .MoveFirst
11220             For lngX = 1& To lngRecs2
11230               lngImports = lngImports + 1&
11240               lngE = lngImports - 1&
11250               ReDim Preserve arr_varImport(I_ELEMS, lngE)
11260               arr_varImport(I_JID, lngE) = ![Journal_ID]
11270               arr_varImport(I_ATYP, lngE) = ![assettype]
11280               arr_varImport(I_CNUM, lngE) = ![CheckNum]
                    ' ** The AssetType field contains the date it was exported from Trust Import: impxl1_dateexport.
                    ' ** The CheckNum field contains the tblImport_Journal_Archive ID: impjarch_id.
                    ' ** The Ledger's PostDate field (currently not used) will contain the export day as the integer
                    ' ** portion of the field, and the archive ID incorporated within the decimal portion.
                    ' ** So, impxl1_dateexport =
                    ' **   Format([postdate], "mm/dd/yyyy")
                    ' ** And, impjarch_id =
                    ' **   IIf(IsNull([postdate])=True,Null,IIf(InStr(CStr(CDbl([postdate])),'.')=0,Null,
                    ' **     CLng(Left(Mid(Mid(CStr(CDbl([postdate])),(InStr(CStr(CDbl([postdate])),'.')+1)),4) &
                    ' **     String(6,'0'),Val(Left(Mid(CStr(CDbl([postdate])),(InStr(CStr(CDbl([postdate])),'.')+1)),2))))))
                    ' ** The 1st character of AssetType gets us within 255 days of the export date,
                    ' **   and the 2nd character gets us to the exact day.
                    ' ** For the archive ID, start with just the decimal portion of the PostDate field.
                    ' **   The 1st position (10th) is always 0.
                    ' **   The 2nd position (100th) tells us the length of the archive id (not likely to get into the millions).
                    ' **   The 3rd position (1000th) is always 0.
                    ' **   The 4th position and beyond (10000th) is the archive ID without any right-most 0's, if present.
                    ' **   So, .0103 = 3, .0207 = 70, .050123 = 12300, .0409876 = 9876.
11290               dblTmp04 = CDbl(CStr(CLng(Asc(Left(![assettype], 1)) * 255#) + CLng(Asc(Right(![assettype], 1)))) & _
                      ".0" & CStr(Len(CStr(![CheckNum]))) & "0" & CStr(![CheckNum]))
11300               arr_varImport(I_PDAT, lngE) = CDate(dblTmp04)
11310               If lngX < lngRecs2 Then .MoveNext
11320             Next
11330           End If
11340           .Close
11350         End With  ' ** rst1
11360         Set rst1 = Nothing
11370         DoEvents
              ' ** rstLedger remains open.

      #End If

11380         If strPostType = POST_USR Then
11390           With qdf1.Parameters
11400             ![usr] = strUser
11410           End With
11420         End If

11430         Set rstJournal = qdf1.OpenRecordset()
11440         With rstJournal
                ' ** It shouldn't get here with no Journal records!
11450           .MoveLast
11460           lngRecs1 = .RecordCount
11470           .MoveFirst
11480           For lngX = 1& To lngRecs1
11490             If Nz(rstJournal![shareface], 0#) = 0# And Nz(rstJournal![ICash], 0@) = 0@ And _
                      Nz(rstJournal![PCash], 0@) = 0@ And Nz(rstJournal![Cost], 0@) = 0@ Then
                    ' ** Don't post zero-value entries!
11500             Else
11510               lngMax = lngMax + 1&
11520               With rstLedger
11530                 .AddNew
11540                 ![journalno] = lngMax                              ' ** dbLong
11550                 ![Location_ID] = Nz(rstJournal![Location_ID], 0&)  ' ** dbLong
11560                 ![assetno] = Nz(rstJournal![assetno], 0#)          ' ** dbLong
11570                 ![accountno] = Trim(rstJournal![accountno])        ' ** dbText (15)
11580                 ![shareface] = Nz(rstJournal![shareface], 0#)      ' ** dbDouble
11590                 ![rate] = Nz(rstJournal![rate], 0#)                ' ** dbDouble
11600                 ![pershare] = Nz(rstJournal![pershare], 0#)        ' ** dbDouble
11610                 ![assetdate] = rstJournal![assetdate]              ' ** dbDate
11620                 ![journaltype] = rstJournal![journaltype]          ' ** dbText (13)
11630                 If rstJournal![journaltype] = "Paid" And IsNull(rstJournal![assettype]) = False And lngImports = 0& Then
                        ' ** This should be an item included in a POSPay file.
11640                   lngPosPays = lngPosPays + 1&
11650                   lngE = lngPosPays - 1&
11660                   ReDim Preserve arr_varPosPay(P_ELEMS, lngE)
11670                   arr_varPosPay(P_ID, lngE) = rstJournal![ID]
11680                   arr_varPosPay(P_JNO, lngE) = lngMax
11690                   arr_varPosPay(P_ATYP, lngE) = rstJournal![assettype]
11700                   arr_varPosPay(P_PPDID, lngE) = Null
11710                 End If
11720                 ![transdate] = rstJournal![transdate]              ' ** dbDate
11730                 ![ICash] = Nz(rstJournal![ICash], 0@)              ' ** dbCurrency
11740                 ![PCash] = Nz(rstJournal![PCash], 0@)              ' ** dbCurrency
11750                 ![Cost] = Nz(rstJournal![Cost], 0@)                ' ** dbCurrency
11760                 ![description] = rstJournal![description]          ' ** dbText (150)  dbText (200)
                      ' ** The query source, above, sets [posted] = Now().
                      ' ** Let's use a common [posted] value throughout this procedure.
11770                 ![posted] = datPosted                              ' ** dbDate
11780                 ![due] = rstJournal![due]                          ' ** dbDate
11790                 ![taxcode] = Nz(rstJournal![taxcode], 1&)          ' ** dbLong  (VGC 08/25/11: Default to Unspecified.)
11800                 ![RecurringItem] = rstJournal![RecurringItem]      ' ** dbText (50)
11810                 ![PurchaseDate] = rstJournal![PurchaseDate]        ' ** dbDate
11820                 ![revcode_ID] = Nz(rstJournal![revcode_ID], 1&)    ' ** dbLong  (VGC 04/12/08: Default to Unspecified.)
11830                 ![journal_USER] = rstJournal![journal_USER]        ' ** dbText (20)
11840                 If rstJournal![journaltype] = "Paid" Then
11850                   ![CheckNum] = rstJournal![CheckNum]              ' ** dbLong
11860                 Else
                        ' ** Used by 'Dividend', 'Interest', 'Purchase' for referencing reinvestments.
11870                   ![CheckNum] = Null
11880                 End If
11890                 ![CheckPaid] = False                               ' ** dbBoolean  (VGC 08/25/11: Already being set to False)
11900                 ![curr_id] = rstJournal![curr_id]
11910                 If lngImports = 0& Then
11920                   ![postdate] = Null                               ' ** dbDate
11930                 Else
11940                   For lngY = 0& To (lngImports - 1&)
11950                     If arr_varImport(I_JID, lngY) = rstJournal![ID] Then
11960                       ![postdate] = arr_varImport(I_PDAT, lngY)
11970                       Exit For
11980                     End If
11990                   Next
12000                 End If
12010                 .Update
12020               End With  ' ** rstLedger.
12030             End If
12040             If lngX < lngRecs1 Then .MoveNext
12050           Next
12060           .Close
12070         End With  ' ** rstJournal.
12080         rstLedger.Close
12090         Set rstJournal = Nothing
12100         Set rstLedger = Nothing
12110         Set qdf1 = Nothing
12120         DoEvents

              ' ***************************************
              ' ***************************************
              ' ** Step 14. Delete Journal Entries.
              ' ***************************************
              ' ***************************************
12130         ctlPostMsg.Caption = "14 of 15 - Delete Journal Entries"
12140         DoEvents

              ' ** Empty tblJournal_MiscSold.
12150         Select Case strPostType
              Case POST_ALL
                ' ** Empty tblJournal_MiscSold, all entries.
12160           Set qdf1 = dbs.QueryDefs("qryJournal_MiscSold_04_01")
12170           qdf1.Execute
12180         Case POST_SYS
                ' ** Delete tblJournal_MiscSold, via subquery to
                ' ** qryJournal_MiscSold_04_02_01 (Journal, just journal_USER = Null, 'System').
12190           Set qdf1 = dbs.QueryDefs("qryJournal_MiscSold_04_02")
12200           qdf1.Execute
12210         Case POST_USR
                ' ** Delete tblJournal_MiscSold, via subquery to
                ' ** qryJournal_MiscSold_04_03_01 (Journal, by specified [usr]).
12220           Set qdf1 = dbs.QueryDefs("qryJournal_MiscSold_04_03")
12230           With qdf1.Parameters
12240             ![usr] = strUser
12250           End With
12260           qdf1.Execute
12270         End Select
12280         Set qdf1 = Nothing
12290         DoEvents

              ' ** 10. Delete posted journal records.
              ' **     Empty the Journal table of posted entries.
      #If HasRepost Then
12300         Select Case gblnRePost
              Case True
12310           Set rstJournal = dbs.OpenRecordset("zz_tbl_RePost_Journal", dbOpenDynaset)
12320         Case False
12330           Set rstJournal = dbs.OpenRecordset("journal", dbOpenDynaset)
12340         End Select
      #Else
12350         Set rstJournal = dbs.OpenRecordset("journal", dbOpenDynaset)
      #End If
12360         With rstJournal
12370           .MoveLast
12380           lngRecs1 = .RecordCount
12390           For lngX = lngRecs1 To 1& Step -1&
12400             Select Case strPostType
                  Case POST_ALL
12410               .Delete
12420             Case POST_SYS
12430               If IsNull(![journal_USER]) = True Then
12440                 .Delete
12450               End If
12460             Case POST_USR
12470               If ![journal_USER] = strUser Then
12480                 .Delete
12490               End If
12500             End Select
12510             If lngX > 1& Then .MovePrevious
12520           Next
12530           .Close
12540         End With  ' ** rstJournal.
12550         Set rstJournal = Nothing
12560         DoEvents

12570         wrk.CommitTrans dbForceOSFlush
              ' ** The flush may release the databases sooner, and might prevent the lockup-on-exit problem.
12580         blnInTrans = False
              ' *************************
              ' ** CommitTrans.
              ' *************************

              ' ** Update POSPay table for posted checks.
12590         If lngPosPays > 0& Then
12600           lngTmp02 = 0&
12610           For lngZ = 0& To (lngPosPays - 1&)
12620             If Len(arr_varPosPay(P_ATYP, lngZ)) = 1 Then
12630               lngTmp02 = Asc(arr_varPosPay(P_ATYP, lngZ))
12640             Else
12650               lngTmp02 = ((Asc(Left(arr_varPosPay(P_ATYP, lngZ), 1)) * 255) + Asc(Right(arr_varPosPay(P_ATYP, lngZ), 1)))
12660             End If
12670             arr_varPosPay(P_PPDID, lngZ) = lngTmp02
12680           Next  ' ** lngZ.
12690           For lngZ = 0& To (lngPosPays - 1&)
                  ' ** Update qryCheckPOSPay_16_09 (tblCheckPOSPay_Detail,
                  ' ** with journalno_new, by specified [ppdid], [jno]).
12700             Set qdf1 = dbs.QueryDefs("qryCheckPOSPay_16_10")
12710             With qdf1.Parameters
12720               ![ppdid] = arr_varPosPay(P_PPDID, lngZ)
12730               ![jno] = arr_varPosPay(P_JNO, lngZ)
12740             End With
12750             qdf1.Execute
12760             Set qdf1 = Nothing
12770           Next  ' ** lngZ.
12780         End If  ' ** lngPosPays.

              ' ** BeginTrans still in effect for blnRollback = True.

12790         DoCmd.Hourglass False
      #If HasRepost Then
12800         If gblnRePost = False Then
12810           If lngRecsToPost_Selected > 0& Then lngRecsToPost = lngRecsToPost_Selected
12820           Select Case lngRecsToPost
                Case 1&
12830             strTmp01 = "1 Transaction Posted."
12840           Case Else
12850             strTmp01 = CStr(lngRecsToPost) & " Transactions Posted."
12860           End Select
12870           intRetVal = POST_DONE  ' ** 5. Post successful.
12880           MsgBox strTmp01, vbInformation + vbOKOnly, "Post Successful"
12890         End If
      #Else
12900         If lngRecsToPost_Selected > 0& Then lngRecsToPost = lngRecsToPost_Selected
12910         Select Case lngRecsToPost
              Case 1&
12920           strTmp01 = "1 Transaction Posted."
12930         Case Else
12940           strTmp01 = CStr(lngRecsToPost) & " Transactions Posted."
12950         End Select
12960         intRetVal = POST_DONE  ' ** 5. Post successful.
12970         MsgBox strTmp01, vbInformation + vbOKOnly, "Post Successful"
      #End If

12980       End If  ' ** blnRollback.

12990       If blnRollback = False Then
13000         dbs.Close
13010         Set dbs = Nothing
13020         wrk.Close
13030         Set wrk = Nothing
13040         DoEvents
13050       End If

13060     End If  ' ** blnSkip.

13070   End If  ' ** lngRecsToPost.
        ' ** BeginTrans still in effect for blnRollback = True.

        ' ***************************************
        ' ***************************************
        ' ** Step 15. Finish.
        ' ***************************************
        ' ***************************************
13080   ctlPostMsg.Caption = "15 of 15 - Finished"
13090   DoEvents

13100   If blnRollback = True Then
          ' ***************************************
          ' ***************************************
          ' ** Step 16. Rollback.
          ' ***************************************
          ' ***************************************
13110     ctlPostMsg.Caption = "16 - Rollback"
13120     DoEvents
13130     intRetVal = POST_ROLLBK  ' ** 3. Problem, causing rollback.
      #If HasRepost Then
13140     If gblnRePost = False Then
13150       Beep
13160       MsgBox strErr, vbCritical + vbOKOnly, "Invalid Entry"
            'Debug.Print "'ROLLBACK: " & CStr(intRollbackPoint) & " " & strErr
13170     End If
      #Else
13180     Beep
13190     MsgBox strErr, vbCritical + vbOKOnly, "Invalid Entry"
      #End If
13200     wrk.Rollback
      #If HasRepost Then
13210     gintRePostBreak = intRollbackPoint + 10
      #End If
13220     Select Case intRollbackPoint
          Case 1
13230       rstJournal.Close
13240       rstAccount.Close
'1 rstJournal
'1 rstAccount
13250     Case 2
13260       rstJournal.Close
13270       rstAccount.Close
'2 rstJournal
'2 rstAccount
13280     Case 3
13290       rstJournal.Close
13300       rstActiveAssets.Close
'3 rstJournal
'3 rstActiveAssets
13310     Case 4
13320       rstJournal.Close
13330       rstActiveAssets.Close
'4 rstJournal
'4 rstActiveAssets
13340     Case 5
13350       rstJournal.Close
13360       rstActiveAssets.Close
'5 rstJournal
'5 rstActiveAssets
13370     Case 6
13380       rstJournal.Close
13390       rstActiveAssets.Close
'6 rstJournal
'6 rstActiveAssets
13400     Case 7
13410       rstJournal.Close
13420       rstActiveAssets.Close
'7 rstJournal
'7 rstActiveAssets
13430     Case 8
13440       rstJournal.Close
13450       rstActiveAssets.Close
'8 rstJournal
'8 rstActiveAssets
13460     End Select
13470     Set rstJournal = Nothing
13480     Set rstAccount = Nothing
13490     Set rstActiveAssets = Nothing
13500     dbs.Close
13510     Set dbs = Nothing
13520     wrk.Close
13530     Set wrk = Nothing
13540     DoEvents
13550   End If  ' ** blnRollback.

13560   DoCmd.Hourglass False

EXITP:
13570   Set ctlPostMsg = Nothing
13580   Set usr = Nothing
13590   Set grp = Nothing
13600   Set qdf1 = Nothing
13610   Set qdf2 = Nothing
13620   Set qdf3 = Nothing
13630   Set rst1 = Nothing
13640   Set rst2 = Nothing
13650   Set rstJournal = Nothing
13660   Set rstAccount = Nothing
13670   Set rstActiveAssets = Nothing
13680   Set rstMasterAsset = Nothing
13690   Set rstLedger = Nothing
13700   Set rstErrLog = Nothing
13710   Set dbs = Nothing
13720   Set wrk = Nothing
13730   PostTransactions = intRetVal
13740   Exit Function

ERRH:
13750   DoCmd.Hourglass False
13760   Select Case ERR.Number
        Case 3021  ' ** No current record.
13770     intRetVal = POST_NOTRANS  ' ** 1. No transactions.
13780     MsgBox "No Transactions to Post!  ²", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
13790     If blnInTrans = True Then
13800       wrk.Rollback
13810       wrk.Close
13820     End If
13830     Resume EXITP
13840   Case Else
      #If HasRepost Then
13850     If ERR.Number = 3022 Then  ' ** The changes you requested to the table were not successful because they
13860       gintRePostBreak = 2      ' ** would create duplicate values in the index, primary key, or relationship.
13870     End If                     ' ** See RePost_All_Load() in modRePostFuncs.
13880     glngRePostErrNum = ERR.Number
13890     gintRePostErrLine = Erl
13900 On Error Resume Next
13910     garr_varRePost(RP_ID, 0) = rstJournal![ID]
13920     garr_varRePost(RP_ACTNO, 0) = rstJournal![accountno]
13930     garr_varRePost(RP_ASTNO, 0) = rstJournal![assetno]
13940     garr_varRePost(RP_ASTDAT, 0) = datAssetDate
13950     garr_varRePost(RP_TRNDAT, 0) = datTransDate
13960     garr_varRePost(RP_RSTNAM, 0) = rstJournal.Name
13970     garr_varRePost(RP_LNGX, 0) = lngX
13980     garr_varRePost(RP_RECS1, 0) = lngRecs1
      #End If
13990     Beep
14000     If blnInTrans = True Then
14010       wrk.Rollback
14020       wrk.Close
14030     End If
14040     intRetVal = POST_ERROR  ' ** 4. Other Error, from error handler.
14050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14060     Resume EXITP
14070   End Select

End Function

Public Function SelectUser() As String

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "SelectUser"

        Dim strRetVal As String

14110   strRetVal = vbNullString

14120   DoCmd.Hourglass False
14130   gstrJournalUser = vbNullString  ' ** Don't clear this on the form!
14140   DoCmd.OpenForm "frmMenu_Post_Clear_Multi", , , , , acDialog, "frmMenu_Post" & "~" & "PostJournal"
14150   DoCmd.Hourglass True
14160   DoEvents

14170   strRetVal = gstrJournalUser

EXITP:
14180   SelectUser = strRetVal
14190   Exit Function

ERRH:
14200   strRetVal = RET_ERR
14210   Select Case ERR.Number
        Case Else
14220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14230   End Select
14240   Resume EXITP

End Function

Private Function TransCount() As Long
' ** This is the same code as cmdPostingJournal_Click() in the frmMenu_Post form.

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "TransCount"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRetVal As Long

14310   lngRetVal = 0&

14320   JrnlChk_EmptyRecs  ' ** Procedure: Below.

14330   Set dbs = CurrentDb
14340   If gblnAdmin = True Then
          ' ** Journal, with add'l fields, all journal_USER's.  #curr_id
14350     Set qdf = dbs.QueryDefs("qryPost_Journal_02a")  ' ** All journal_USER's.  '#jnox
14360   Else
          ' ** qryPost_Journal_02b (Journal, with add'l fields, by specified journal_USER), by specified [usr].
14370     Set qdf = dbs.QueryDefs("qryPost_Journal_02c")  ' ** Specified journal_USER.  '#jnox
14380     With qdf.Parameters
14390       ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
14400     End With
14410   End If
14420   Set rst = qdf.OpenRecordset
14430   If rst.BOF = True And rst.EOF = True Then
14440     rst.Close
14450     dbs.Close
14460   Else
14470     rst.Close
14480     If gblnAdmin = True Then
            ' ** qryPost_Journal_04a (qryPost_Journal_02a (Journal, with add'l fields,
            ' ** all journal_USER's), grouped and summed; all journal_USER's), grouped,
            ' ** with Count of Journal transactions to be posted; all journal_USER's.
14490       Set qdf = dbs.QueryDefs("qryPost_Journal_11a")  '#jnox
14500     Else
            ' ** qryPost_Journal_04b (qryPost_Journal_02b (Journal, with add'l fields,
            ' ** by specified journal_USER), grouped and summed; specified journal_USER,
            ' ** by specified [usr]), grouped, with Count of Journal transactions to be
            ' ** posted; specified journal_USER.
14510       Set qdf = dbs.QueryDefs("qryPost_Journal_11b")  '#jnox
14520       With qdf.Parameters
14530         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
14540       End With
14550     End If
14560     Set rst = qdf.OpenRecordset
14570     With rst
14580       .MoveFirst
14590       lngRetVal = ![cnt]
14600       .Close
14610     End With
14620     dbs.Close
14630   End If

EXITP:
14640   Set rst = Nothing
14650   Set qdf = Nothing
14660   Set dbs = Nothing
14670   TransCount = lngRetVal
14680   Exit Function

ERRH:
14690   lngRetVal = 0&
14700   Select Case ERR.Number
        Case Else
14710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14720   End Select
14730   Resume EXITP

End Function

Public Sub JrnlChk_EmptyRecs()
' ** Collect all these Journal/Ledger checks...

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlChk_EmptyRecs"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long

14810   If gblnAdmin = True And gblnGoToReport = False Then
14820     lngRecs = 0&
14830     Set dbs = CurrentDb
14840     With dbs
            ' ** Journal, just completely empty records.
14850       Set qdf = .QueryDefs("qryJournal_04")
14860       Set rst = qdf.OpenRecordset
14870       With rst
14880         If .BOF = True And .EOF = True Then
                ' ** Everything's OK.
14890         Else
14900           .MoveLast
14910           lngRecs = .RecordCount
14920         End If
14930         .Close
14940       End With
14950       Set rst = Nothing
14960       Set qdf = Nothing
14970       If lngRecs > 0& Then
              ' ** Delete qryJournal_04 (Journal, just completely empty records).
14980         Set qdf = .QueryDefs("qryJournal_05")
14990         qdf.Execute
              'Beep
              'MsgBox "There appears to be a completely empty record in the Journal," & vbCrLf & _
              '  "which can cause problems when creating new entries." & vbCrLf & vbCrLf & _
              '  "To assure the Journal is clean, post all pending transactions, then clear all transactions." & vbCrLf & vbCrLf & _
              '  "If the problem persists, contact Delta Data, Inc.", vbCritical + vbOKOnly, "Empty Journal Record"
15000       End If
15010       .Close
15020     End With
15030   End If

EXITP:
15040   Set rst = Nothing
15050   Set qdf = Nothing
15060   Set dbs = Nothing
15070   Exit Sub

ERRH:
15080   Select Case ERR.Number
        Case Else
15090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15100   End Select
15110   Resume EXITP

End Sub

Public Function TotDescChk(varInput As Variant, varJComment As Variant, varJrnlUser As Variant) As Variant
' ** varInput is totdesc.
' ** Called by queries.

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "TotDescChk"

        Dim intCnt As Integer
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim varRetVal As Variant

15210   varRetVal = Null

        ' ** Also lookd for dupes of 'posted by'.
15220   If IsNull(varInput) = False Then
15230     strTmp01 = Trim(varInput)
15240     strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString
15250     If strTmp01 <> vbNullString Then
15260       Select Case IsNull(varJComment)
            Case True
15270         strTmp02 = vbNullString
15280       Case False
15290         strTmp02 = Trim(varJComment)
15300       End Select
15310       If strTmp02 <> vbNullString Then
              ' ** Check for presense of JComment 2 or more times in totdesc.
15320         intCnt = CharCnt(strTmp01, strTmp02, True)  ' ** Module Function: modStringFuncs.
15330         If intCnt > 1 Then
15340           intPos01 = InStr(strTmp01, strTmp02)
15350           intPos02 = CharPos(strTmp01, 2, strTmp02)  ' ** Module Function: modStringFuncs.
15360           intPos03 = CharPos(strTmp01, 3, strTmp02)  ' ** Module Function: modStringFuncs.
15370           If intPos03 > 0 Then
                  ' ** Take out the 3rd one first.
15380             strTmp03 = Left(strTmp01, (intPos03 - 1))
15390             strTmp04 = Mid(strTmp01, intPos03)
15400             strTmp04 = Trim(StringReplace(strTmp04, strTmp02, vbNullString))  ' ** Module Function: modStringFuncs.
15410             strTmp01 = Trim((strTmp03 & strTmp04))
15420             strTmp01 = Rem_Spaces(strTmp01)  ' ** Module Function: modStringFuncs.
15430           End If
                ' ** Now take out the 2nd one.
15440           strTmp03 = Left(strTmp01, (intPos02 - 1))
15450           strTmp04 = Mid(strTmp01, intPos02)
15460           strTmp04 = Trim(StringReplace(strTmp04, strTmp02, vbNullString))  ' ** Module Function: modStringFuncs.
15470           strTmp01 = Trim((strTmp03 & strTmp04))
15480           strTmp01 = Rem_Spaces(strTmp01)  ' ** Module Function: modStringFuncs.
15490         End If
15500       End If  ' ** strTmp02.
15510       strTmp03 = vbNullString: strTmp04 = vbNullString
15520       Select Case IsNull(varJrnlUser)
            Case True
15530         strTmp02 = vbNullString
15540       Case False
15550         strTmp02 = Trim(varJrnlUser)
15560       End Select
15570       If strTmp02 <> vbNullString Then
15580         strTmp02 = "posted by " & strTmp02
              ' ** Check for presense of 'posted by' 2 or more times in totdesc.
15590         intCnt = CharCnt(strTmp01, strTmp02, True)  ' ** Module Function: modStringFuncs.
15600         If intCnt > 1 Then
15610           intPos01 = InStr(strTmp01, strTmp02)
15620           intPos02 = CharPos(strTmp01, 2, strTmp02)  ' ** Module Function: modStringFuncs.
15630           intPos03 = CharPos(strTmp01, 3, strTmp02)  ' ** Module Function: modStringFuncs.
15640           If intPos03 > 0 Then
                  ' ** Take out the 3rd one first.
15650             strTmp03 = Left(strTmp01, (intPos03 - 1))
15660             strTmp04 = Mid(strTmp01, intPos03)
15670             strTmp04 = Trim(StringReplace(strTmp04, strTmp02, vbNullString))  ' ** Module Function: modStringFuncs.
15680             strTmp01 = Trim((strTmp03 & strTmp04))
15690             strTmp01 = Rem_Spaces(strTmp01)  ' ** Module Function: modStringFuncs.
15700           End If
                ' ** Now take out the 2nd one.
15710           strTmp03 = Left(strTmp01, (intPos02 - 1))
15720           strTmp04 = Mid(strTmp01, intPos02)
15730           strTmp04 = Trim(StringReplace(strTmp04, strTmp02, vbNullString))  ' ** Module Function: modStringFuncs.
15740           strTmp01 = Trim((strTmp03 & strTmp04))
15750           strTmp01 = Rem_Spaces(strTmp01)  ' ** Module Function: modStringFuncs.
15760         End If
15770       End If  ' ** strTmp02.
15780       If strTmp01 <> vbNullString Then
15790         varRetVal = strTmp01
15800       End If
15810     End If  ' ** vbNullString.
15820   End If  ' ** Null.

EXITP:
15830   TotDescChk = varRetVal
15840   Exit Function

ERRH:
15850   varRetVal = RET_ERR
15860   Select Case ERR.Number
        Case Else
15870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15880   End Select
15890   Resume EXITP

End Function
