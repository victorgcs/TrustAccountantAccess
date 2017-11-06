Attribute VB_Name = "zz_mod_SullivanFuncs"
Option Compare Database
Option Explicit

'VGC 03/22/2015: CHANGES!

' ** ½ ¼
Private Const THIS_NAME As String = "zz_mod_SullivanFuncs"
' **

Public Function Sul_Tbl_Chk() As Boolean

  Const THIS_PROC As String = "Sul_Tbl_Chk"

  Dim dbs As DAO.Database, rst As DAO.Recordset, fld As DAO.Field
  Dim strTblName As String
  Dim lngRecs As Long, lngFlds As Long
  Dim lngEmptyFlds As Long
  Dim blnEmpty As Boolean
  Dim lngX As Long, lngY As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  strTblName = "tblSullivan_AccountSynopsis_excel_do"
  'strTblName = "tblSullivan_TransactionType_excel"
  'strTblName = "tblSullivan_AssetCodes_excel"

  Debug.Print "'TBL: " & strTblName
  DoEvents

  Set dbs = CurrentDb
  With dbs
    lngEmptyFlds = 0&
    Set rst = .OpenRecordset(strTblName, dbOpenDynaset, dbReadOnly)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      lngFlds = .Fields.Count
      For lngX = 0& To (lngFlds - 1&)
        blnEmpty = True
        .MoveFirst
        Set fld = .Fields(lngX)
        With fld
          If Left(.Name, 5) = "Field" Then
            For lngY = 1& To lngRecs
              If IsNull(.Value) = False Then
                blnEmpty = False
                Exit For
              End If
              If lngY < lngRecs Then rst.MoveNext
            Next  ' ** lngY.
            If blnEmpty = True Then
              lngEmptyFlds = lngEmptyFlds + 1&
              Debug.Print "'EMPTY!  " & .Name
              DoEvents
            End If
          End If
        End With  ' ** fld.

      Next  ' ** lngX.

    End With  ' ** rst.
    Set rst = Nothing


    .Close
  End With  ' dbs.
  Set dbs = Nothing

  Debug.Print "'EMPTY FLDS: " & CStr(lngEmptyFlds)
  DoEvents

'TBL: tblSullivan_AccountSynopsis_excel_do
'EMPTY!  Field24
'EMPTY!  Field30
'EMPTY!  Field55
'EMPTY!  Field56
'EMPTY!  Field58
'EMPTY!  Field60
'EMPTY!  Field63
'EMPTY!  Field66
'EMPTY!  Field70
'EMPTY!  Field76
'EMPTY!  Field78
'EMPTY!  Field80
'EMPTY!  Field81
'EMPTY!  Field83
'EMPTY!  Field85
'EMPTY!  Field87
'EMPTY!  Field89
'EMPTY!  Field91
'EMPTY!  Field99
'EMPTY FLDS: 19
'DONE!

'EMPTY!  Field4
'EMPTY!  Field5
'EMPTY!  Field8
'EMPTY!  Field13
'EMPTY!  Field15
'EMPTY!  Field16
'EMPTY!  Field19
'EMPTY!  Field21
'EMPTY!  Field23
'EMPTY!  Field24
'EMPTY!  Field26
'EMPTY!  Field28
'EMPTY!  Field30
'EMPTY!  Field31
'EMPTY!  Field35
'EMPTY!  Field37
'EMPTY!  Field38
'EMPTY!  Field40
'EMPTY!  Field41
'EMPTY!  Field44
'EMPTY!  Field46
'EMPTY!  Field47
'EMPTY!  Field50
'EMPTY!  Field53
'EMPTY!  Field55
'EMPTY!  Field56
'EMPTY!  Field57
'EMPTY!  Field59
'EMPTY FLDS: 28
'DONE!
  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set fld = Nothing
  Set rst = Nothing
  Set dbs = Nothing

  Sul_Tbl_Chk = blnRetVal

End Function

Public Function Sul_Fld_Chk() As Boolean

  Const THIS_PROC As String = "Sul_Fld_Chk"

  Dim dbs As DAO.Database, rst As DAO.Recordset, fld As DAO.Field
  Dim strTblName As String, strFind1 As String
  Dim lngRecs As Long, lngFlds As Long, lngFldNum As Long
  Dim blnFound As Boolean, lngRecsFound As Long
  Dim lngX As Long, lngY As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  strFind1 = "Start"
  strTblName = "tblSullivan_AccountSynopsis_excel"

  Set dbs = CurrentDb
  With dbs
    lngRecsFound = 0&
    Set rst = .OpenRecordset(strTblName, dbOpenDynaset, dbReadOnly)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      lngFlds = .Fields.Count
      .MoveFirst
      For lngX = 1& To lngRecs
        blnFound = False: lngFldNum = 0&
        For lngY = 0& To (lngFlds - 1&)
          Set fld = .Fields(lngY)
          With fld
            If IsNull(.Value) = False Then
              If .Type = dbText Then
                If InStr(.Value, strFind1) > 0 Then
                  blnFound = True
                  lngFldNum = lngY
                  Exit For
                End If
              End If
            End If
          End With  ' ** fld.
        Next  ' ** lngY.
        If blnFound = True Then
          Debug.Print "'REC: " & CStr(lngX) & "  FLD: " & CStr(lngFldNum)
          DoEvents
          lngRecsFound = lngRecsFound + 1&
        End If
        If lngX < lngRecs Then .MoveNext
      Next  ' ** lngX.

    End With  ' ** rst.
    Set rst = Nothing


    .Close
  End With  ' dbs.
  Set dbs = Nothing

  Debug.Print "'RECS FOUND: " & CStr(lngRecsFound)
  DoEvents

  Debug.Print "'DONE!"
  DoEvents

'Account Type:
'Account:
' Account_Code  Account_Name
'Admin. Last Reviewed:
'Admin. Officer:
'Admin. Review Cycle:
'Birth Date:
'Branch Code:
'Cap Gain Rate:
'Capacity:
'Checking Account:
'Corpus Balance:
'Cover Sheet:
'Custom Fields:
'Date Closed:
'Date Opened:
'Death Date Spouse:
'Death Date:
'Discretionary:
'Dist. of Principal:
'Fee Cycle:
'Fee Table:
'Fee Year End Date:
'Income Cash Balance:
'Income Reserve:
'Invest Income Balance:
'Invest. Last Reviewed:
'Invest. Officer:
'Invest. Review Cycle:
'Investment Balance:
'Investment Model:
'Investment Objective:
'Investment Powers:
'LT Loss Carry Fwd:
'Links:
'Multiple Copies:
'Net Cash Overdraft:
'Overdraft Allowed:
'Principal Cash Balance:
'Principal Only:
'Principal Reserve:
'Proxy Powers:
'Purge:
'Real Estate Officer:
'Report Notes:
'ST Loss Carry Fwd:
'Specific Powers:
'Statement Cycle:
'Status Code:
'Sweep Asset:
'Tax ID #:
'Tax Lot:
'Tax Officer:
'Tax Type:


'LawFirm
'GovernDesc
'GovernDate

'Type
'Beneficial Owner
'Category
'Area
'Exclusion Category
'Exclusion Value
'Name
'Action
'Accu
'Net
'Trade
'Stmt
'Bene
'Dist
'Dist%
'Stop
'Trust%


  Beep

  Set fld = Nothing
  Set rst = Nothing
  Set dbs = Nothing

  Sul_Fld_Chk = blnRetVal

End Function

Public Function Sul_TblDate() As Boolean

  Const THIS_PROC As String = "Sul_TblDate"

  Dim dbs As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field
  Dim strFormat As String
  Dim lngTypos As Long
  Dim intPos1 As Integer
  Dim strTmp01 As String
  Dim blnRetVal As Boolean

  blnRetVal = True

  lngTypos = 0&
  Set dbs = CurrentDb
  With dbs
    For Each tdf In .TableDefs
      With tdf
        If Left(.Name, 11) = "tblSullivan" Then
          For Each fld In .Fields
            With fld
              If .Type = dbDate Then
On Error Resume Next
                strFormat = .Properties("Format")
                If ERR.Number = 0 Then
On Error GoTo 0
                  intPos1 = InStr(strFormat, "/yy")
                  If intPos1 > 0 Then
                    strTmp01 = Mid(strFormat, intPos1)
                    If InStr(strTmp01, "yyyy") = 0 Then
                      ' ** Oops!
                      Debug.Print "'TBL: " & tdf.Name
                      DoEvents
                      lngTypos = lngTypos + 1&
                    End If
                  End If
                Else
On Error GoTo 0
                End If
              End If
            End With
          Next
        End If
      End With
    Next
    Set tdf = Nothing
    .Close
  End With
  Set dbs = Nothing

  Debug.Print "'TYPOS: " & CStr(lngTypos)
  DoEvents

'TBL: tblSullivan_AccountCode
'TBL: tblSullivan_AccountCode_excel_do
'TBL: tblSullivan_AccountCode_InUse
'TBL: tblSullivan_AccountCode_Primary
'TBL: tblSullivan_AccountList
'TBL: tblSullivan_AccountList_excel
'TBL: tblSullivan_AccountSynopsis_excel
'TBL: tblSullivan_AccountSynopsis_excel_do
'TBL: tblSullivan_AssetHolder_excel
'TBL: tblSullivan_AssetHolder_excel_do
'TBL: tblSullivan_AssetLocation_excel_do
'TBL: tblSullivan_AssetMaster_excel_do
'TBL: tblSullivan_Tmp01
'TBL: tblSullivan_Tmp02
'TBL: tblSullivan_Tmp03
'TBL: tblSullivan_Tmp04
'TBL: tblSullivan_Transaction_excel
'TBL: tblSullivan_Transaction_excel_do
'TBL: tblSullivan_TransType
'TBL: tblSullivan_TransType_excel
'TBL: tblSullivan_TransType_excel_do
'TBL: tblSullivan_TransType_Primary
'TBL: tblSullivan_TransType_Report
'TBL: tblSullivan_WhereHeld
'TBL: tblSullivan_WhereHeld_excel
'TBL: tblSullivan_WhereHeld_excel_do
'TYPOS: 26
'DONE!

  Debug.Print "'DONE!"
  DoEvents

  Beep

'sulam_issue_new_new: IIf(IsNull([sulam_issue])=True,Null,StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(Rem_Spaces(Rem_CrLf([sulam_issue]))
','QUARTERLY','QTR')
','CERTIFICATE OF DEPOSIT','CD')
','CERT OF DEPOSIT','CD')
','ONE YEAR','1 YEAR')
','WEEK','WK')
','MONTH','MO')
','YEAR','YR')
','RENEWABLE','RENEW')
','MATURITY','MAT')
','MATURING','MAT')
','VARIABLE','VAR')
','AUTOMATICALLY','AUTO')
','INTEREST','INT')
','CHECK','CHK')
','SINGLE','SNGL')
','SPECIAL','SPEC')
','RE-INVESTED','RE-INV')
','REINVESTED','RE-INV')
','PAYMENT','PMT')
','CTF DEP','CD')
','DATED','DTD'))

'sulam_issue_new_new: IIf(IsNull([sulam_issue])=True,0,
'IIf(InStr([sulam_issue],'CERTIFICATE OF DEPOSIT')>0 Or
'InStr([sulam_issue],'CERT OF DEPOSIT')>0 Or
'InStr([sulam_issue],'ONE YEAR')>0 Or
'InStr([sulam_issue],'WEEK')>0 Or
'InStr([sulam_issue],'MONTH')>0 Or
'InStr([sulam_issue],'YEAR')>0 Or
'InStr([sulam_issue],'RENEWABLE')>0 Or
'InStr([sulam_issue],'MATURITY')>0 Or
'InStr([sulam_issue],'MATURING')>0 Or
'InStr([sulam_issue],'VARIABLE')>0 Or
'InStr([sulam_issue],'AUTOMATICALLY')>0 Or
'InStr([sulam_issue],'INTEREST')>0 Or
'InStr([sulam_issue],'CHECK')>0 Or
'InStr([sulam_issue],'SINGLE')>0 Or
'InStr([sulam_issue],'SPECIAL')>0 Or
'InStr([sulam_issue],'RE-INVESTED')>0 Or
'InStr([sulam_issue],'REINVESTED')>0,-1,0))
' ,'PAYMENT'
' ,'CTF DEP'
' ,'DATED'
' ,'ANNUAL'
' ,'YIELD'
' ,'RENEWAL'

'IIf(IsNull([sulam_issue])=True,0,IIf(InStr([sulam_issue],'CERTIFICATE OF DEPOSIT')>0 Or InStr([sulam_issue],'CERT OF DEPOSIT')>0 Or InStr([sulam_issue],'ONE YEAR')>0 Or InStr([sulam_issue],'WEEK')>0 Or InStr([sulam_issue],'MONTH')>0 Or InStr([sulam_issue],'YEAR')>0 Or InStr([sulam_issue],'RENEWABLE')>0 Or InStr([sulam_issue],'MATURITY')>0 Or InStr([sulam_issue],'MATURING')>0 Or InStr([sulam_issue],'VARIABLE')>0 Or InStr([sulam_issue],'AUTOMATICALLY')>0 Or InStr([sulam_issue],'INTEREST')>0 Or InStr([sulam_issue],'CHECK')>0 Or InStr([sulam_issue],'SINGLE')>0 Or InStr([sulam_issue],'SPECIAL')>0 Or InStr([sulam_issue],'RE-INVESTED')>0 Or InStr([sulam_issue],'REINVESTED')>0,-1,0))

' ,'OPENED','OPND'
' ,'MUNICIPAL','MUNI'
' ,'BONDS','BND'
' ,'BOND','BND'
' ,'COUPON','CPN'

'sulam_issue_new_new: IIf(IsNull([sulam_issue_new])=True,Null,StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(Rem_CrLf([sulam_issue_new]),'OPENED','OPND'),'MUNICIPAL','MUNI'),'BONDS','BND'),'BOND','BND'),'COUPON','CPN'))


'Or InStr([sulam_issue],'OPENED')>0 Or InStr([sulam_issue],'MUNICIPAL')>0 Or InStr([sulam_issue],'BOND')>0 Or InStr([sulam_issue],'COUPON')>0



  Set fld = Nothing
  Set tdf = Nothing
  Set dbs = Nothing

  Sul_TblDate = blnRetVal

End Function

Public Function Sul_Dots(varInput As Variant) As Integer

  Const THIS_PROC As String = "Sul_Dots"

  Dim strTmp01 As String
  Dim intRetVal As Integer

  intRetVal = 0

  If IsNull(varInput) = False Then
    strTmp01 = Trim(varInput)
    If InStr(strTmp01, "..........") > 0 Then
      intRetVal = 10
    ElseIf InStr(strTmp01, ".........") > 0 Then
      intRetVal = 9
    ElseIf InStr(strTmp01, "........") > 0 Then
      intRetVal = 8
    ElseIf InStr(strTmp01, ".......") > 0 Then
      intRetVal = 7
    ElseIf InStr(strTmp01, "......") > 0 Then
      intRetVal = 6
    ElseIf InStr(strTmp01, ".....") > 0 Then
      intRetVal = 5
    ElseIf InStr(strTmp01, "....") > 0 Then
      intRetVal = 4
    ElseIf InStr(strTmp01, "...") > 0 Then
      intRetVal = 3
    ElseIf InStr(strTmp01, "..") > 0 Then
      intRetVal = 2
    End If
  End If

  Sul_Dots = intRetVal

End Function

Public Function Sul_TransAcct_Order() As Boolean

  Const THIS_PROC As String = "Sul_TransAcct_Order"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngAccts As Long, arr_varAcct As Variant
  Dim lngRecs As Long, lngOrder As Long, lngLastAcct As Long
  Dim lngX As Long, lngY As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varAcct().
  Const A_SULTID As Integer = 0
  Const A_ACTNO  As Integer = 1
  Const A_SHORT  As Integer = 2
  Const A_ORD    As Integer = 3
  Const A_ISACCT As Integer = 4
  Const A_FIRST  As Integer = 5
  Const A_LAST   As Integer = 6
  Const A_DAT    As Integer = 7

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    ' ** zzz_qry_Sullivan_Transactionx_09_07 (zzz_qry_Sullivan_Transactionx_09_04
    ' ** (zzz_qry_Sullivan_Transactionx_09_01 (tblSullivan_Transaction_excel_do,
    ' ** just IsAcct = True, with Field1_account_num, Field2_account_description),
    ' ** linked to zzz_qry_Sullivan_Transactionx_09_03 (zzz_qry_Sullivan_Transactionx_09_02
    ' ** (zzz_qry_Sullivan_Transactionx_09_01 (tblSullivan_Transaction_excel_do,
    ' ** just IsAcct = True, with Field1_account_num, Field2_account_description),
    ' ** grouped by Field1_account_num, with cnt), linked to tblSullivan_AccountSynopsis),
    ' ** with sulas_account_short_name, sult_id_last), linked to tblMark_AutoNum,
    ' ** with sult_order, sult_order_prev), linked to itself, with sult_id_first, sult_id_last.
    Set qdf = .QueryDefs("zzz_qry_Sullivan_Transactionx_09_08")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngAccts = .RecordCount
      .MoveFirst
      arr_varAcct = .GetRows(lngAccts)
      ' *************************************************************
      ' ** Array: arr_varAcct()
      ' **
      ' **   Field  Element  Name                        Constant
      ' **   =====  =======  ==========================  ==========
      ' **     1       0     sult_id                     A_SULTID
      ' **     2       1     Field1_account_num          A_ACTNO
      ' **     3       2     sulas_account_short_name    A_SHORT
      ' **     4       3     sult_order                  A_ORD
      ' **     5       4     IsAcct                      A_ISACCT
      ' **     6       5     sult_id_first               A_FIRST
      ' **     7       6     sult_id_last                A_LAST
      ' **     8       7     sult_datemodified           A_DAT
      ' **
      ' *************************************************************
      .Close
    End With  ' ** rst.
    Set rst = Nothing
    Set qdf = Nothing

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Debug.Print "'ACCT SEPS: " & CStr(lngAccts)
    DoEvents

    Set rst = .OpenRecordset("tblSullivan_Transaction_excel_do", dbOpenDynaset, dbConsistent)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      Debug.Print "'RECS: " & CStr(lngRecs)
      DoEvents
      Debug.Print "'|";
      lngLastAcct = 0&
      For lngX = 1& To lngRecs
        lngOrder = 0&
        For lngY = lngLastAcct To (lngAccts - 1&)
          If ![sult_id] >= arr_varAcct(A_FIRST, lngY) And ![sult_id] <= arr_varAcct(A_LAST, lngY) Then
            lngOrder = arr_varAcct(A_ORD, lngY)
            lngLastAcct = lngY
            Exit For
          End If
        Next  ' ** lngY.
        .Edit
        ![sult_order] = lngOrder
        ![sult_datemodified] = Now()
        .Update
        If (lngX Mod 10000) = 0 Then
          Debug.Print "|  " & CStr(lngX)
          Debug.Print "'|";
        ElseIf (lngX Mod 1000) = 0 Then
          Debug.Print "|";
        ElseIf (lngX Mod 100) = 0 Then
          Debug.Print ".";
        End If
        DoEvents
        If lngX < lngRecs Then .MoveNext
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.
    Set rst = Nothing
    Debug.Print

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Debug.Print "'DONE!"
  DoEvents

'ACCT SEPS: 8487
'RECS: 59642
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|.........|  10000
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|.........|  20000
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|.........|  30000
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|.........|  40000
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|.........|  50000
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|......
'DONE!
  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Sul_TransAcct_Order = blnRetVal

End Function

Public Function Sul_QryMove() As Boolean

  Const THIS_PROC As String = "Sul_QryMove"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strQryName As String, strQryStart As String, strQryEnd As String, strQryNewSeq As String
  Dim strSQL As String, strDesc As String
  Dim strQryNum As String, lngQryNum As Long
  Dim lngQrysCreated As Long, lngQrysDeleted As Long
  Dim blnStart As Boolean, blnSkip As Boolean
  Dim intPos1 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngW As Long, lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 8  ' ** Array's first-element UBound().
  Const Q_QNAM1 As Integer = 0
  Const Q_SQL1  As Integer = 1
  Const Q_DSC1  As Integer = 2
  Const Q_SFX1  As Integer = 3
  Const Q_TYP   As Integer = 4
  Const Q_QNAM2 As Integer = 5
  Const Q_SQL2  As Integer = 6
  Const Q_DSC2  As Integer = 7
  Const Q_SFX2  As Integer = 8

  Const QRY_BASE As String = "zzz_qry_Sullivan_Transactionx_27_"

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    For lngW = 1& To 4&

      Select Case lngW
      Case 1&
        strQryStart = "zzz_qry_Sullivan_Transactionx_27_06_05"
        strQryEnd = "zzz_qry_Sullivan_Transactionx_27_06_16"
        strQryNewSeq = "06_05_01"
      Case 2&
        strQryStart = "zzz_qry_Sullivan_Transactionx_27_06_17"
        strQryEnd = "zzz_qry_Sullivan_Transactionx_27_06_24"
        strQryNewSeq = "06_06_01"
      Case 3&
        strQryStart = "zzz_qry_Sullivan_Transactionx_27_06_25"
        strQryEnd = "zzz_qry_Sullivan_Transactionx_27_06_44"
        strQryNewSeq = "06_07_01"
      Case 4&
        strQryStart = "zzz_qry_Sullivan_Transactionx_27_06_45"
        strQryEnd = "zzz_qry_Sullivan_Transactionx_27_06_58"
        strQryNewSeq = "06_08_01"
      End Select

      lngQrys = 0&
      ReDim arr_varQry(Q_ELEMS, 0)

      blnStart = False
      For Each qdf In .QueryDefs
        With qdf
          If Left(.Name, Len(QRY_BASE)) = QRY_BASE Then
            If blnStart = False Then
              If .Name = strQryStart Then
                blnStart = True
              End If
            End If
            If blnStart = True Then
              lngQrys = lngQrys + 1&
              lngE = lngQrys - 1&
              ReDim Preserve arr_varQry(Q_ELEMS, lngE)
              arr_varQry(Q_QNAM1, lngE) = .Name
              arr_varQry(Q_SQL1, lngE) = .SQL
              arr_varQry(Q_DSC1, lngE) = .Properties("Description")
              strTmp01 = Mid(.Name, (Len(QRY_BASE) + 1))
              arr_varQry(Q_SFX1, lngE) = strTmp01
              arr_varQry(Q_TYP, lngE) = .Type
              arr_varQry(Q_QNAM2, lngE) = Null
              arr_varQry(Q_SQL2, lngE) = Null
              arr_varQry(Q_DSC2, lngE) = Null
              arr_varQry(Q_SFX2, lngE) = Null
            End If
          End If
          If .Name = strQryEnd Then
            Exit For
          End If
        End With
      Next  ' ** qdf
      Set qdf = Nothing

      strQryNum = vbNullString: lngQryNum = 0&
      strQryName = vbNullString: strSQL = vbNullString: strDesc = vbNullString

      If lngQrys > 0& Then

        ' ** Create the new names.
        For lngX = 0& To (lngQrys - 1&)
          strQryName = arr_varQry(Q_QNAM1, lngX)
          strTmp01 = QRY_BASE
          lngQryNum = lngQryNum + 1&
          strQryNum = Right("00" & CStr(lngQryNum), 2)
          strTmp01 = strTmp01 & Left(strQryNewSeq, 6) & strQryNum
          arr_varQry(Q_QNAM2, lngX) = strTmp01
          strTmp02 = Mid(strTmp01, (Len(QRY_BASE) + 1))
          arr_varQry(Q_SFX2, lngX) = strTmp02
        Next  ' ** lngX.

        ' ** Swap SQL references for others within the group.
        For lngX = 0& To (lngQrys - 1&)
          strSQL = arr_varQry(Q_SQL1, lngX)
          For lngY = 0& To (lngQrys - 1&)
            If lngY <> lngX Then  ' ** Never self-referencial.
              ' ** zzz_qry_Sullivan_Transactionx_27_06_05.
              strTmp01 = arr_varQry(Q_QNAM2, lngY)
              strTmp01 = "xxx" & Mid(strTmp01, 4)  ' ** To avoid infinite loop.
              strSQL = StringReplace(strSQL, CStr(arr_varQry(Q_QNAM1, lngY)), strTmp01)  ' ** Module Function: modStringFuncs.
              strTmp02 = "zzz" & Mid(strTmp01, 4)  ' ** Now set it back.
              strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            End If
          Next  ' ** lngY.
          arr_varQry(Q_SQL2, lngX) = strSQL
        Next  ' ** lngX.

        ' ** Swap Description references for others within the group.
        For lngX = 0& To (lngQrys - 1&)
          ' **                 .._27_06_05, linked back to .._26_01; 3.
          ' **                 .._27_06_07, linked back to .._26_01; 27.
          ' **                     Append .._27_06_10 to tblSullivan_Tmp14.
          ' **                     Update .._27_06_09.
          ' **                 tblSullivan_AssetMaster, via subquery to .._27_06_26; 6.
          strDesc = arr_varQry(Q_DSC1, lngX)
          For lngY = 0& To (lngQrys - 1&)
            strTmp01 = arr_varQry(Q_SFX2, lngY)
            strTmp01 = Left(strTmp01, 1) & "x" & Mid(strTmp01, 2)
            strDesc = StringReplace(strDesc, CStr(arr_varQry(Q_SFX1, lngY)), strTmp01)  ' ** Module Function: modStringFuncs.
            strTmp02 = Left(strTmp01, 1) & Mid(strTmp01, 3)
            strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
          Next  ' ** lngY.
          arr_varQry(Q_DSC2, lngX) = strDesc
        Next  ' ** lngX.

        Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
        DoEvents

        lngQrysCreated = 0&: lngQrysDeleted = 0&

        blnSkip = False
        If blnSkip = False Then

          For lngX = 0& To (lngQrys - 1&)
            Set qdf = .CreateQueryDef(arr_varQry(Q_QNAM2, lngX), arr_varQry(Q_SQL2, lngX))
            With qdf
              Set prp = .CreateProperty("Description", dbText, arr_varQry(Q_DSC2, lngX))
On Error Resume Next
              .Properties.Append prp
              If ERR.Number <> 0 Then
On Error GoTo 0
                .Properties("Description") = arr_varQry(Q_DSC2, lngX)
              Else
On Error GoTo 0
              End If
            End With
            Set prp = Nothing
            Set qdf = Nothing
            lngQrysCreated = lngQrysCreated + 1&
          Next  ' ** lngX.
          Set prp = Nothing
          Set qdf = Nothing
          .QueryDefs.Refresh

          Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
          DoEvents

        End If  ' ** blnSkip.

        blnSkip = False
        If blnSkip = False Then

          For lngX = (lngQrys - 1&) To 0& Step -1&
            DoCmd.DeleteObject acQuery, arr_varQry(Q_QNAM1, lngX)
            lngQrysDeleted = lngQrysDeleted + 1&
          Next
          .QueryDefs.Refresh

          Debug.Print "'QRYS DELETED: " & CStr(lngQrysDeleted)
          DoEvents

        End If  ' ** blnSkip.

      Else
        Beep!
      End If  ' ** lngQrys.

'Exit For
    Next  ' ** lngW.

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set prp = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Sul_QryMove = blnRetVal

End Function

Public Function Sul_GetPct(varInput As Variant) As Variant

  Const THIS_PROC As String = "Sul_GetPct"

  Dim intPos1 As Integer, intLen As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim intX As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  If IsNull(varInput) = False Then
    strTmp01 = Trim(varInput)
    intPos1 = InStr(strTmp01, "%")
    If intPos1 > 0 Then
      strTmp02 = Left(strTmp01, intPos1)
      intLen = Len(strTmp02)
      For intX = intLen To 1 Step -1
        strTmp03 = Mid(strTmp02, intX, 1)
        Select Case strTmp03
        Case " ", ",", ")", "("
          strTmp02 = Mid(strTmp02, (intX + 1))
          Exit For
        End Select
      Next
      varRetVal = strTmp02
    End If
  End If

  Sul_GetPct = varRetVal

End Function

Public Function Sul_QryDel() As Boolean

  Const THIS_PROC As String = "Sul_QryDel"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim lngQrysDeleted As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 0
  Const Q_QNAM As Integer = 0

  blnRetVal = True

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)  ' ** Array's first-element UBound().

  Set dbs = CurrentDb
  With dbs
    For Each qdf In .QueryDefs
      With qdf
        If InStr(.Name, "Sullivan") > 0 Then
          lngQrys = lngQrys + 1&
          lngE = lngQrys - 1&
          ReDim Preserve arr_varQry(Q_ELEMS, lngE)
          arr_varQry(Q_QNAM, lngE) = .Name
        End If
      End With
    Next
    .Close
  End With

  Debug.Print "'QRYS: " & CStr(lngQrys)
  DoEvents

  If lngQrys > 0& Then

    lngQrysDeleted = 0&
    For lngX = 0& To (lngQrys - 1&)
      DoCmd.DeleteObject acQuery, arr_varQry(Q_QNAM, lngX)
      DoEvents
      lngQrysDeleted = lngQrysDeleted + 1&
    Next
    CurrentDb.QueryDefs.Refresh

    Debug.Print "'QRYS DELETED: " & CStr(lngQrysDeleted)
    DoEvents

  Else
    Debug.Print "'NONE FOUND!"
  End If

  Beep
  Debug.Print "'DONE!"

  Set qdf = Nothing
  Set dbs = Nothing

  Sul_QryDel = blnRetVal

End Function

Public Function Sul_TblDel() As Boolean

  Const THIS_PROC As String = "Sul_TblDel"

  Dim dbs As DAO.Database, tdf As DAO.TableDef
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngTblsDeleted As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 0
  Const T_TNAM As Integer = 0

  blnRetVal = True

  lngTbls = 0&
  ReDim arr_varTbl(T_ELEMS, 0)  ' ** Array's first-element UBound().

  Set dbs = CurrentDb
  With dbs
    For Each tdf In .TableDefs
      With tdf
        If InStr(.Name, "Sullivan") > 0 Then
          lngTbls = lngTbls + 1&
          lngE = lngTbls - 1&
          ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          arr_varTbl(T_TNAM, lngE) = .Name
        End If
      End With
    Next
    .Close
  End With

  Debug.Print "'TBLS: " & CStr(lngTbls)
  DoEvents

  If lngTbls > 0& Then

    lngTblsDeleted = 0&
    For lngX = 0& To (lngTbls - 1&)
      DoCmd.DeleteObject acTable, arr_varTbl(T_TNAM, lngX)
      DoEvents
      lngTblsDeleted = lngTblsDeleted + 1&
    Next
    CurrentDb.TableDefs.Refresh

    Debug.Print "'TBLS DELETED: " & CStr(lngTblsDeleted)
    DoEvents

  Else
    Debug.Print "'NONE FOUND!"
  End If

  Beep
  Debug.Print "'DONE!"

  Set tdf = Nothing
  Set dbs = Nothing

  Sul_TblDel = blnRetVal

End Function
