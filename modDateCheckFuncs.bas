Attribute VB_Name = "modDateCheckFuncs"
Option Compare Database
Option Explicit

'VGC 06/07/2012: CHANGES!

Private Const THIS_NAME As String = "modDateCheckFuncs"
' **

Public Function CheckDates() As Boolean
' ** Check predate in the Account table, the earliest and latest Balance Date in the Balance table,
' ** and Statement_Date in the Statement Date table to insure all dates are sensible.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "CheckDates"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngRecs As Long
        Dim varTmp00 As Variant, datTmp01 As Date
        Dim lngX As Long
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     Set dbs = CurrentDb
130     With dbs

          ' ** If a new account, with a predate after Statement_Date, now has older transactions posted (filling in a history
          ' ** of transactions), the predate should always stay the same; only the Balance Date should get moved back.

          ' ** Update [predate] in the Account table; this usually does nothing.
140       Set qdf1 = .QueryDefs("qryStatementParameters_05")
150       qdf1.Execute
160       Set qdf1 = Nothing

          ' ** Append Balance table records if there are none; this usually does nothing.
170       Set qdf1 = .QueryDefs("qryStatementParameters_06b")
180       qdf1.Execute
190       Set qdf1 = Nothing

          ' ** If that initial Balance record is all zeroes, and now there are earlier posted transactions,
          ' ** just change the date on that initial Balance record, don't add a new zero entry.

          ' ** qryStatementParameters_03e, just Bal_Date_Min discrepancies.
200       Set qdf1 = .QueryDefs("qryStatementParameters_06c")
210       Set rst1 = qdf1.OpenRecordset
220       If rst1.BOF = True And rst1.EOF = True Then
            ' ** No discrepancies.
230         rst1.Close
240         Set rst1 = Nothing
250         Set qdf1 = Nothing
260       Else
270         rst1.MoveLast
280         lngRecs = rst1.RecordCount
290         rst1.MoveFirst
300         For lngX = 1& To lngRecs
              ' ** Balance table, just zero value entries, by specified [actno].
310           Set qdf2 = .QueryDefs("qryStatementParameters_06d")
320           With qdf2.Parameters
330             ![actno] = rst1![accountno]
340           End With
350           Set rst2 = qdf2.OpenRecordset
360           If rst2.BOF = True And rst2.EOF = True Then
                ' ** If it doesn't have an initial zero entry, add it.
370             rst2.Close
380             Set rst2 = Nothing
390             Set qdf2 = Nothing
                ' ** Append initial zero entry to Balance table, by specified [actno], [baldat].
400             Set qdf2 = .QueryDefs("qryStatementParameters_06e")
410             With qdf2.Parameters
420               ![actno] = rst1![accountno]
430               ![baldat] = rst1![Bal_Date_Minx]
440             End With
450             qdf2.Execute
460             Set qdf2 = Nothing
470           Else
                ' ** If it has an initial zero entry, and it's later than Bal_Date_Minx, edit it.
480             rst2.MoveFirst
490             rst2.Edit
500             rst2![balance date] = rst1![Bal_Date_Minx]
510             rst2.Update
520             rst2.Close
530             Set rst2 = Nothing
540             Set qdf2 = Nothing
550           End If
560           If lngX < lngRecs Then rst1.MoveNext
570         Next
580         rst1.Close
590         Set rst1 = Nothing
600         Set qdf1 = Nothing
610       End If  ' ** BOF/EOF.

          ' ** Balance table, not in qryStatementParameters_08a (zero value entries); these are the only ones with non-zero values.
620       Set qdf1 = .QueryDefs("qryStatementParameters_08b")
630       Set rst1 = qdf1.OpenRecordset
640       If rst1.BOF = True And rst1.EOF = True Then
            ' ** No real data in Balance table, so Statement_Date must be earlier than first transaction.
650         rst1.Close
660         Set rst1 = Nothing
670         Set qdf1 = Nothing
            ' ** Ledger, with Min(transdate).
680         Set qdf1 = .QueryDefs("qryStatementParameters_08g")
690         Set rst1 = qdf1.OpenRecordset
700         If rst1.BOF = True And rst1.EOF = True Then
              ' ** Wow! Must be a brand new installation. Make sure Statement_Date is 01/01/1990.
710           rst1.Close
720           Set rst1 = Nothing
730           Set qdf1 = Nothing
740           datTmp01 = #1/1/1990#
750           Set rst1 = .OpenRecordset("Statement Date", dbOpenDynaset)
760           If rst1.BOF = True And rst1.EOF = True Then
                ' ** This should never happen!
770             rst1.AddNew
780             rst1![Statement_Date] = datTmp01
790             rst1.Update
800           Else
810             rst1.MoveFirst
820             If rst1![Statement_Date] <> datTmp01 Then
830               rst1.Edit
840               rst1![Statement_Date] = datTmp01
850               rst1.Update
860             End If
870           End If
880           rst1.Close
890           Set rst1 = Nothing
900         Else
              ' ** qryStatementParameters_08g will return an EMPTY record if there's nothing yet in the Ledger!
910           rst1.MoveFirst
920           If IsNull(rst1![transdate]) = True Then
930             datTmp01 = #1/1/1990#
940           Else
950             datTmp01 = rst1![transdate]
960           End If
970           rst1.Close
980           Set rst1 = Nothing
990           Set qdf1 = Nothing
1000          Set rst1 = .OpenRecordset("Statement Date", dbOpenDynaset)
1010          If rst1.BOF = True And rst1.EOF = True Then
                ' ** This should never happen!
1020            rst1.AddNew
1030            rst1![Statement_Date] = datTmp01 - 1
1040            rst1.Update
1050          Else
1060            rst1.MoveFirst
1070            If rst1![Statement_Date] >= datTmp01 And datTmp01 <> #1/1/1990# Then
1080              rst1.Edit
1090              rst1![Statement_Date] = datTmp01 - 1
1100              rst1.Update
1110            End If
1120          End If
1130          rst1.Close
1140          Set rst1 = Nothing
1150        End If
1160      Else
1170        rst1.Close
1180        Set rst1 = Nothing
1190        Set qdf1 = Nothing
            ' ** qryStatementParameters_08e (max Balance Date and Statement_Date), if dates are not the same.
1200        Set qdf1 = .QueryDefs("qryStatementParameters_08f")
1210        Set rst1 = qdf1.OpenRecordset
1220        If rst1.BOF = True And rst1.EOF = True Then
              ' ** Everything's fine.
1230        Else
1240          rst1.MoveFirst
1250          datTmp01 = rst1![balance_date]
1260          rst1.Close
1270          Set rst1 = Nothing
1280          Set qdf1 = Nothing
1290          Set rst1 = .OpenRecordset("Statement Date", dbOpenDynaset)
1300          If rst1.BOF = True And rst1.EOF = True Then
                ' ** This should never happen!
1310            rst1.AddNew
1320            rst1![Statement_Date] = datTmp01
1330            rst1.Update
1340          Else
1350            rst1.MoveFirst
1360            rst1.Edit
1370            rst1![Statement_Date] = datTmp01
1380            rst1.Update
1390          End If
1400        End If
1410        rst1.Close
1420        Set rst1 = Nothing
1430        Set qdf1 = Nothing
1440      End If  ' ** BOF/EOF.

          ' ** All of this assumes a Null date is OK to proceed!

          ' ** PostingDate table, Username = Null.
1450      Set qdf1 = dbs.QueryDefs("qryPostingDate_03")
1460      Set rst1 = qdf1.OpenRecordset
1470      If rst1.BOF = True And rst1.EOF = True Then
            ' ** All's well.
1480        rst1.Close
1490        Set rst1 = Nothing
1500        Set qdf1 = Nothing
1510      Else
1520        rst1.Close
1530        Set rst1 = Nothing
1540        Set qdf1 = Nothing
            ' ** Delete PostingDate for Username = Null.
1550        Set qdf1 = .QueryDefs("qryPostingDate_08")
1560        qdf1.Execute
1570        Set qdf1 = Nothing
1580      End If  ' ** BOF/EOF.
1590      varTmp00 = DCount("*", "PostingDate")
1600      If varTmp00 = 0 Then
            ' ** Append new record to PostingDate table, with Posting_Date = Date(), by specified [usr].
1610        Set qdf1 = .QueryDefs("qryPostingDate_09")
1620        With qdf1.Parameters
1630          ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1640        End With
1650        qdf1.Execute
1660        Set qdf1 = Nothing
1670      Else
            ' ** PostingDate, by specified [usr].
1680        Set qdf1 = .QueryDefs("qryPostingDate_07")
1690        With qdf1.Parameters
1700          ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1710        End With
1720        Set rst1 = qdf1.OpenRecordset
1730        If rst1.BOF = True And rst1.EOF = True Then
              ' ** This user not there.
1740          rst1.Close
1750          Set rst1 = Nothing
1760          Set qdf1 = Nothing
              ' ** Append new record to PostingDate table, with Posting_Date = Date(), by specified [usr].
1770          Set qdf1 = .QueryDefs("qryPostingDate_09")
1780          With qdf1.Parameters
1790            ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1800          End With
1810          qdf1.Execute
1820          Set qdf1 = Nothing
1830        Else
              ' ** All's well
1840          rst1.Close
1850          Set rst1 = Nothing
1860          Set qdf1 = Nothing
1870        End If
1880      End If

1890      .Close
1900    End With  ' ** dbs.

EXITP:
1910    Set rst1 = Nothing
1920    Set rst2 = Nothing
1930    Set qdf1 = Nothing
1940    Set qdf2 = Nothing
1950    Set dbs = Nothing
1960    CheckDates = blnRetVal
1970    Exit Function

ERRH:
1980    blnRetVal = False
1990    Select Case ERR.Number
        Case Else
2000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2010    End Select
2020    Resume EXITP

End Function
