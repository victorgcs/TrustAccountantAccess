Attribute VB_Name = "zz_mod_RevzonFuncs"
Option Compare Database
Option Explicit

'VGC 06/08/2014: CHANGES!

'AccuTech
Private Const THIS_NAME As String = "zz_mod_RevzonFuncs"
' **

Public Function RZ_Avg1() As Boolean

  Const THIS_PROC As String = "RZ_Avg1"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim blnMultiWS As Boolean, blnMultiWSCovered As Boolean, blnWSCovered As Boolean, blnMultiLot As Boolean
  Dim blnMatchPair1 As Boolean, blnMatchPair2 As Boolean, blnMatchPair3 As Boolean
  Dim blnFromTheTop As Boolean
  Dim datThisDate As Date, datTestDate As Date
  Dim lngWSs As Long, lngDPs As Long
  Dim lngRecs As Long
  Dim blnSkip As Boolean, blnDate1 As Boolean, blnDate2 As Boolean, blnDate3 As Boolean
  Dim varTmp00 As Variant, lngTmp01 As Long, dblTmp02 As Double, dblTmp03 As Double
  Dim lngW As Long, lngX As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    ' ** From the very beginning.
    blnFromTheTop = False

    blnDate1 = False: blnDate2 = False: blnDate3 = False

    For lngW = 211& To 300&

      lngWSs = 0&: lngDPs = 0&
      gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

      If blnFromTheTop = True And lngW = 1& Then
        blnSkip = False
      Else
        blnSkip = True
      End If
      If blnSkip = False Then
        ' ** First time only.

        ' ** Empty zz_tbl_Revzon_01.
        Set qdf = .QueryDefs("zzz_qry_Revzon_29_01")
        qdf.Execute
        Set qdf = Nothing
        ' ** Empty zz_tbl_Revzon_02.
        Set qdf = .QueryDefs("zzz_qry_Revzon_29_02")
        qdf.Execute
        Set qdf = Nothing
        ' ** Empty zz_tbl_Revzon_03.
        Set qdf = .QueryDefs("zzz_qry_Revzon_29_04")
        qdf.Execute
        Set qdf = Nothing
        ' ** Empty zz_tbl_Revzon_06.
        Set qdf = .QueryDefs("zzz_qry_Revzon_29_08")
        qdf.Execute
        Set qdf = Nothing

        ' **********************************************
        ' ** Append Code: A.
        ' ** zz_tbl_Revzon_02
        ' ** zz_tbl_Revzon_03
        ' **********************************************
        gsCoCity = Right("000" & CStr(lngW), 3) & ".A"
        ' **********************************************

        ' ** Append zzz_qry_Revzon_20_04 (xx) to zz_tbl_Revzon_01.
        Set qdf = .QueryDefs("zzz_qry_Revzon_20_05")
        qdf.Execute
        Set qdf = Nothing
        ' ** Append zzz_qry_Revzon_20_07 (xx) to zz_tbl_Revzon_02.
        Set qdf = .QueryDefs("zzz_qry_Revzon_20_08")
        qdf.Execute
        Set qdf = Nothing

      End If  ' ** blnSkip.

      gsCoCity = vbNullString

      ' ** Empty zz_tbl_Revzon_02_tmp.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_03")
      qdf.Execute
      Set qdf = Nothing

      ' ** Empty zz_tbl_Revzon_03_tmp.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_05")
      qdf.Execute
      Set qdf = Nothing

      ' ** Empty zz_tbl_Revzon_05.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_07")
      qdf.Execute
      Set qdf = Nothing
      DoEvents

      ' ** Append zzz_qry_Revzon_22_02 (xx) to zz_tbl_Revzon_05
      Set qdf = .QueryDefs("zzz_qry_Revzon_22_02a")
      qdf.Execute
      Set qdf = Nothing
      DoEvents

      varTmp00 = DLookup("[transdate]", "zz_tbl_Revzon_05")
      If IsNull(varTmp00) = False Then
        datThisDate = CDate(varTmp00)
      End If
      Debug.Print "'" & Left((CStr(lngW) & ".") & Space(4), 4) & " THIS DATE: " & datThisDate
      DoEvents

      ' ** zzz_qry_Revzon_22_03 (xx), just this day's dp's.
      Set qdf = .QueryDefs("zzz_qry_Revzon_23_01")
      Set rst = qdf.OpenRecordset
      If rst.BOF = True And rst.EOF = True Then
        ' ** No dp's today.
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
        Debug.Print "'DP'S: " & CStr(lngDPs)
        DoEvents
      Else
        ' ** We have dp's today.
        rst.MoveLast
        lngDPs = rst.RecordCount
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
        Debug.Print "'DP'S: " & CStr(lngDPs)
        DoEvents

        ' **********************************************
        ' ** Append Code: B.
        ' ** zz_tbl_Revzon_02
        ' ** zz_tbl_Revzon_03
        ' **********************************************
        gsCoCity = Right("000" & CStr(lngW), 3) & ".B"
        ' **********************************************

        ' **********************************************
        ' ** Append Code: B.
        ' ** zz_tbl_Revzon_02_tmp
        ' ** zz_tbl_Revzon_03_tmp
        ' **********************************************
        gsCoZip = Right("000" & CStr(lngW), 3) & ".B"
        ' **********************************************

        ' ** Append zzz_qry_Revzon_23_02 (xx) to zz_tbl_Revzon_01.
        Set qdf = .QueryDefs("zzz_qry_Revzon_23_03")
        qdf.Execute dbFailOnError
        Set qdf = Nothing
        ' ** Append zzz_qry_Revzon_23_05 (xx) to zz_tbl_Revzon_02.
        Set qdf = .QueryDefs("zzz_qry_Revzon_23_06")
        qdf.Execute dbFailOnError
        Set qdf = Nothing
        ' ** Append zzz_qry_Revzon_23_09 (xx) to zz_tbl_Revzon_02_tmp.
        Set qdf = .QueryDefs("zzz_qry_Revzon_23_10")
        qdf.Execute dbFailOnError
        Set qdf = Nothing
      End If  ' ** dp's.

      blnMultiWS = False
      gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

      ' ** zzz_qry_Revzon_03 (xx), linked to zzz_qry_Revzon_23_12 (xx), just this_day's ws's.
      Set qdf = .QueryDefs("zzz_qry_Revzon_24_01")
      Set rst = qdf.OpenRecordset
      If rst.BOF = True And rst.EOF = True Then
        ' ** No ws's today.
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
        Debug.Print "'WS'S: " & CStr(lngWSs)
        DoEvents
      Else
        ' ** We've got ws's.
        rst.MoveLast
        lngWSs = rst.RecordCount
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
        Debug.Print "'WS'S: " & CStr(lngWSs)
        DoEvents

'If lngW = 12& Then
'Stop
'End If
        ' ** zzz_qry_Revzon_24_01 (xx), grouped, by accountno, assetno, with cnt > 1, multiple ws's on the same day.
        Set qdf = .QueryDefs("zzz_qry_Revzon_24_02_01")
        Set rst = qdf.OpenRecordset
        If rst.BOF = True And rst.EOF = True Then
          ' ** Keep it simple.
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
        Else
          ' ** Yup!
          blnMultiWS = True
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
        End If

        blnMatchPair1 = False

        ' ** Empty zz_tbl_Revzon_16.
        Set qdf = .QueryDefs("zzz_qry_Revzon_29_18")
        qdf.Execute
        Set qdf = Nothing
        ' ** Append zzz_qry_Revzon_24_06_00_04 (xx) to zz_tbl_Revzon_16.
        Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_00_04a")
        qdf.Execute
        Set qdf = Nothing

        ' ** Empty zz_tbl_Revzon_10.
        Set qdf = .QueryDefs("zzz_qry_Revzon_29_12")
        qdf.Execute
        Set qdf = Nothing
        ' ** Append zzz_qry_Revzon_24_06_00_06 (xx) to zz_tbl_Revzon_10.
        Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_00a")
        qdf.Execute
        Set qdf = Nothing

        ' ** zzz_qry_Revzon_24_06 (xx), linked to zz_tbl_Revzon_10, just Sx1, matching pairs.
        Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01")
        Set rst = qdf.OpenRecordset
        If rst.BOF = True And rst.EOF = True Then
          ' ** OK.
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
        Else
          ' ** Oooh! Goody!
          blnMatchPair1 = True
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
          ' ** zzz_qry_Revzon_24_06_01_01 (xx), linked to zzz_qry_Revzon_23_07_12 (xx), any with 2 dp's.
          Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_02")
          Set rst = qdf.OpenRecordset
          If rst.BOF = True And rst.EOF = True Then
            ' ** Good!
            rst.Close
            Set rst = Nothing
            Set qdf = Nothing

            ' **********************************************
            ' ** Append Code: C.
            ' ** zz_tbl_Revzon_02_tmp
            ' ** zz_tbl_Revzon_03_tmp
            ' **********************************************
            gsCoZip = Right("000" & CStr(lngW), 3) & ".C"
            ' **********************************************

            ' ** Append zzz_qry_Revzon_24_06_01_03 (xx) to zz_tbl_Revzon_03_tmp.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_05")
            qdf.Execute dbFailOnError
            Set qdf = Nothing

            ' **********************
            ' ** Check for closed!
            ' **********************
            ' ** zzz_qry_Revzon_24_06_01_04 (xx), just dp_closed = True.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_07")
            Set rst = qdf.OpenRecordset
            If rst.BOF = True And rst.EOF = True Then
              ' ** Fine.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
            Else
              ' ** Make sure everybody knows about it!
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
              ' ** Append zzz_qry_Revzon_24_06_01_07 (xx) to zz_tbl_Revzon_06.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_08")
              qdf.Execute dbFailOnError
              Set qdf = Nothing
            End If

            ' ** Append zzz_qry_Revzon_24_06_01_04 (xx) to zz_tbl_Revzon_02_tmp.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_06")
            qdf.Execute dbFailOnError
            Set qdf = Nothing

          Else
            ' ** Not dealt with.  'HIT: 195. 06/18/2013
            rst.Close
            Set rst = Nothing
            Set qdf = Nothing

            ' ** zzz_qry_Revzon_24_06_01_12 (zzz_qry_Revzon_24_06_01_02 (xx), just matching
            ' ** pairs, grouped by journalno_ws, with Min(PurchaseDate), cnt), just cnt > 1.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_13")
            Set rst = qdf.OpenRecordset
            If rst.BOF = True And rst.EOF = True Then
              ' ** Good, I don't feel like more complications right now!
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing

              ' **********************************************
              ' ** Append Code: D.
              ' ** zz_tbl_Revzon_02_tmp
              ' ** zz_tbl_Revzon_03_tmp
              ' **********************************************
              gsCoZip = Right("000" & CStr(lngW), 3) & ".D"
              ' **********************************************

              ' ** Append zzz_qry_Revzon_24_06_01_14 (xx) to zz_tbl_Revzon_03_tmp.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_16")
              qdf.Execute dbFailOnError
              Set qdf = Nothing

              ' **********************
              ' ** Check for closed!
              ' **********************
              ' ** zzz_qry_Revzon_24_06_01_15 (xx), just dp_closed = True.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_18")
              Set rst = qdf.OpenRecordset
              If rst.BOF = True And rst.EOF = True Then
                ' ** Fine.
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
              Else
                ' ** Make sure everybody knows about it!
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_01_18 (xx) to zz_tbl_Revzon_06.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_19")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
              End If

              ' ** Append zzz_qry_Revzon_24_06_01_15 (xx) to zz_tbl_Revzon_02_tmp.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_17")
              qdf.Execute dbFailOnError
              Set qdf = Nothing

              ' ** Append zzz_qry_Revzon_24_06_01_20 (zzz_qry_Revzon_24_06_01_11 (xx),
              ' ** non-matching avg_upd's, as zz_tbl_Revzon_02 records) to zz_tbl_Revzon_02_tmp.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_01_21")
              qdf.Execute dbFailOnError
              Set qdf = Nothing

            Else
              ' ** Oy! It just goes on and on...  Not dealt with yet.
              ' ** These would be more than 1 ws -- of the same amount -- matching at least 1 dp!
              ' ** (Are there 2 completely independent sets? 1 dp matching 2 ws's?)
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
              Stop
            End If
          End If
        End If

        blnWSCovered = False
        gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

        ' ** zzz_qry_Revzon_24_06 (xx), linked to zz_tbl_Revzon_10, just Sx2, dp covers ws.
        Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_02")
        Set rst = qdf.OpenRecordset
        If rst.BOF = True And rst.EOF = True Then
          ' ** OK.
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
        Else
          ' ** Simple coverage.
          blnWSCovered = True
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
        End If

        blnMultiLot = False

        ' ** zzz_qry_Revzon_24_06 (xx), linked to zz_tbl_Revzon_10, just Sx3, ws not covered.
        Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03")
        Set rst = qdf.OpenRecordset
        If rst.BOF = True And rst.EOF = True Then
          ' ** OK.
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
        Else
          ' ** Multi-Lot sale.
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
          ' ** zzz_qry_Revzon_24_06_03_01 (xx), linked to zz_tbl_Revzon_02_tmp.
          Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_02")
          Set rst = qdf.OpenRecordset
          If rst.BOF = True And rst.EOF = True Then
            ' ** None in temp table.
            blnMultiLot = True
            rst.Close
            Set rst = Nothing
            Set qdf = Nothing
            ' ** zzz_qry_Revzon_24_06_03 (xx), linked to zzz_qry_Revzon_24_06_03_03 (xx), with .._new, .._gtot fields.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_05")
            Set rst = qdf.OpenRecordset
            rst.MoveLast
            lngRecs = rst.RecordCount
            rst.Close
            ' ** zzz_qry_Revzon_24_06 (xx), linked to zz_tbl_Revzon_10, just Sx3, ws not covered.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03")
            Set rst = qdf.OpenRecordset
            rst.MoveLast
            If lngRecs > (2 * rst.RecordCount) Then
              ' ** There are more than 2 dp's per ws.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing

              gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

              Select Case lngRecs
              Case 3&  '72. 12/10/2012    MULTI-LOT WITH 3 DP'S!
                ' ** May have only 2 that are needed.

                ' ** Empty zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_19")
                qdf.Execute
                Set qdf = Nothing
                ' ** Empty zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_20")
                qdf.Execute
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_05 (zzz_qry_Revzon_24_06_03 (xx), linked to
                ' ** zzz_qry_Revzon_24_06_03_03 (xx), with .._new, .._gtot fields) to zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_05_00a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_06_03_04 (zzz_qry_Revzon_24_06_03_06_03_03 (xx),
                ' ** linked to zz_tbl_Revzon_17, for assetdate1) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_03_04a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_03_05 (zzz_qry_Revzon_24_06_03_06_03_03 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate2) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_03_05a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_03_06 (zzz_qry_Revzon_24_06_03_06_03_03 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate3) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_03_06a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' **********************************************
                ' ** Append Code: E.
                ' ** zz_tbl_Revzon_02_tmp
                ' ** zz_tbl_Revzon_03_tmp
                ' **********************************************
                gsCoZip = Right("000" & CStr(lngW), 3) & ".E"
                ' **********************************************

                ' ** Append zzz_qry_Revzon_24_06_03_06_03_08 (zz_tbl_Revzon_18, as zz_tbl_Revzon_03_tmp
                ' ** records; 2, excludes Zeroes) to zz_tbl_Revzon_03_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_03_10")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' **********************
                ' ** Check for closed!
                ' **********************
                ' ** zzz_qry_Revzon_24_06_03_06_03_09 (xx), just dp_closed = True.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_03_12")
                Set rst = qdf.OpenRecordset
                If rst.BOF = True And rst.EOF = True Then
                  ' ** Fine.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                Else
                  ' ** Make sure everybody knows about it!
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_06_03_06_03_12 (xx) to zz_tbl_Revzon_06.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_03_13")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing
                End If

                ' ** Append zzz_qry_Revzon_24_06_03_06_03_09 (zz_tbl_Revzon_18, as zz_tbl_Revzon_02_tmp
                ' ** records; 3, includes Zeroes) to zz_tbl_Revzon_02_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_03_11")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

              Case 4&  '190. 6/07/2013    MULTI-LOT WITH 4 DP'S!

                ' ** Empty zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_19")
                qdf.Execute
                Set qdf = Nothing
                ' ** Empty zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_20")
                qdf.Execute
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_05 (zzz_qry_Revzon_24_06_03 (xx), linked to
                ' ** zzz_qry_Revzon_24_06_03_03 (xx), with .._new, .._gtot fields) to zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_05_00a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_06_04_04 (zzz_qry_Revzon_24_06_03_06_04_03,
                ' ** linked to zz_tbl_Revzon_17, for assetdate1) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_04_04a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_04_05 (zzz_qry_Revzon_24_06_03_06_04_03,
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate2) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_04_05a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_04_06 (zzz_qry_Revzon_24_06_03_06_04_03,
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate3) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_04_06a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_04_07 (zzz_qry_Revzon_24_06_03_06_04_03,
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate4) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_04_07a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' **********************************************
                ' ** Append Code: F.
                ' ** zz_tbl_Revzon_02_tmp
                ' ** zz_tbl_Revzon_03_tmp
                ' **********************************************
                gsCoZip = Right("000" & CStr(lngW), 3) & ".F"
                ' **********************************************

                ' ** Append zzz_qry_Revzon_24_06_03_06_04_09 (zz_tbl_Revzon_18, as zz_tbl_Revzon_03_tmp
                ' ** records; 2, excludes Zeroes) to zz_tbl_Revzon_03_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_04_11")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' **********************
                ' ** Check for closed!
                ' **********************
                ' ** zzz_qry_Revzon_24_06_03_06_04_10 (zz_tbl_Revzon_18, as zz_tbl_Revzon_02_tmp
                ' ** records; 4, includes Zeroes), just dp_closed = True.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_04_13")
                Set rst = qdf.OpenRecordset
                If rst.BOF = True And rst.EOF = True Then
                  ' ** Fine.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                Else
                  ' ** Make sure everybody knows about it!
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_06_03_06_04_13 (zzz_qry_Revzon_24_06_03_06_04_10,
                  ' ** just dp_closed = True) to zz_tbl_Revzon_06.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_04_14")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing
                End If

                ' ** Append zzz_qry_Revzon_24_06_03_06_04_10 (zz_tbl_Revzon_18, as zz_tbl_Revzon_02_tmp
                ' ** records; 4, includes Zeroes) to zz_tbl_Revzon_02_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_04_12")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

              Case 5&

                ' ** Empty zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_19")
                qdf.Execute
                Set qdf = Nothing
                ' ** Empty zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_20")
                qdf.Execute
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_05 (zzz_qry_Revzon_24_06_03 (xx), linked to
                ' ** zzz_qry_Revzon_24_06_03_03 (xx), with .._new, .._gtot fields) to zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_05_00a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_06_05_05 (zzz_qry_Revzon_24_06_03_06_05_04 (xx),
                ' ** linked to zz_tbl_Revzon_17, for assetdate1) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_05_05a") '68. 12/04/2012  INVALID USE OF NULL!
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_05_06 (zzz_qry_Revzon_24_06_03_06_05_04 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate2) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_05_06a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_05_07 (zzz_qry_Revzon_24_06_03_06_05_04 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate3) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_05_07a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_05_08 (zzz_qry_Revzon_24_06_03_06_05_04 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate4) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_05_08a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_05_09 (zzz_qry_Revzon_24_06_03_06_05_04 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate5) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_05_09a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' **********************************************
                ' ** Append Code: G.
                ' ** zz_tbl_Revzon_02_tmp
                ' ** zz_tbl_Revzon_03_tmp
                ' **********************************************
                gsCoZip = Right("000" & CStr(lngW), 3) & ".G"
                ' **********************************************

                ' ** Append zzz_qry_Revzon_24_06_03_06_05_11 (zz_tbl_Revzon_18, as zz_tbl_Revzon_03_tmp records) to zz_tbl_Revzon_03_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_05_13")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' **********************
                ' ** Check for closed!
                ' **********************
                ' ** zzz_qry_Revzon_24_06_03_06_05_12 (xx), just dp_closed = True.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_05_15")
                Set rst = qdf.OpenRecordset
                If rst.BOF = True And rst.EOF = True Then
                  ' ** Fine.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                Else
                  ' ** Make sure everybody knows about it!
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_06_03_06_05_15 (xx) to zz_tbl_Revzon_06.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_05_16")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing
                End If

                ' ** Append zzz_qry_Revzon_24_06_03_06_05_12 (zz_tbl_Revzon_18,
                ' ** as zz_tbl_Revzon_02_tmp records) to zz_tbl_Revzon_02_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_05_14")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

              Case 8&  '142. 03/27/2013    MULTI-LOT WITH 8 DP'S!
                ' ** May have only 6 that are needed.

                ' ** Empty zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_19")
                qdf.Execute
                Set qdf = Nothing
                ' ** Empty zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_20")
                qdf.Execute
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_05 (zzz_qry_Revzon_24_06_03 (xx), linked to
                ' ** zzz_qry_Revzon_24_06_03_03 (xx), with .._new, .._gtot fields) to zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_05_00a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_06_08_06 (zzz_qry_Revzon_24_06_03_06_08_05 (xx),
                ' ** linked to zz_tbl_Revzon_17, for assetdate1) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_06a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_08_07 (zzz_qry_Revzon_24_06_03_06_08_05 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate2) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_07a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_08_08 (zzz_qry_Revzon_24_06_03_06_08_05 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate3) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_08a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_08_09 (zzz_qry_Revzon_24_06_03_06_08_05 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate4) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_09a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_08_10 (zzz_qry_Revzon_24_06_03_06_08_05 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate5) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_10a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_08_11 (zzz_qry_Revzon_24_06_03_06_08_05 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate6) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_11a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_08_12 (zzz_qry_Revzon_24_06_03_06_08_05 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate7) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_12a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_08_13 (zzz_qry_Revzon_24_06_03_06_08_05 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate8) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_13a")

                ' **********************************************
                ' ** Append Code: H.
                ' ** zz_tbl_Revzon_02_tmp
                ' ** zz_tbl_Revzon_03_tmp
                ' **********************************************
                gsCoZip = Right("000" & CStr(lngW), 3) & ".H"
                ' **********************************************

                ' ** Append zzz_qry_Revzon_24_06_03_06_08_15 (zz_tbl_Revzon_18, as zz_tbl_Revzon_03_tmp
                ' ** records; 6, excludes Zeroes) to zz_tbl_Revzon_03_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_17")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' **********************
                ' ** Check for closed!
                ' **********************
                ' ** zzz_qry_Revzon_24_06_03_06_08_16 (zzz_qry_Revzon_24zz_tbl_Revzon_18, as
                ' ** zz_tbl_Revzon_02_tmp records; 8, includes Zeroes), just dp_closed = True.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_19")
                Set rst = qdf.OpenRecordset
                If rst.BOF = True And rst.EOF = True Then
                  ' ** Fine.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                Else
                  ' ** Make sure everybody knows about it!
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_06_03_06_08_19 (zzz_qry_Revzon_24_06_03_06_08_16 (xx),
                  ' ** just dp_closed = True) to zz_tbl_Revzon_06.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_20")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing
                End If

                ' ** Append zzz_qry_Revzon_24_06_03_06_08_16 (zz_tbl_Revzon_18, as zz_tbl_Revzon_02_tmp
                ' ** records; 8, includes Zeroes) to zz_tbl_Revzon_02_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_08_18")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

              Case 13&  '211. 07/17/2013    MULTI-LOT WITH 13 DP'S!

                ' ** Empty zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_19")
                qdf.Execute
                Set qdf = Nothing
                ' ** Empty zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_20")
                qdf.Execute
                Set qdf = Nothing
                ' ** Empty zz_tbl_Revzon_19.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_21")
                qdf.Execute
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_05 (zzz_qry_Revzon_24_06_03 (xx), linked to
                ' ** zzz_qry_Revzon_24_06_03_03 (xx), with .._new, .._gtot fields) to zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_05_00a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_06_13_09 (zzz_qry_Revzon_24_06_03_06_13_08,
                ' ** linked to zz_tbl_Revzon_17, for assetdate01) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_09a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate02) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_10a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate03) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_11a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate04) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_12a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate05) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_13a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate06) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_14a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate07) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_15a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate08) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_16a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate09) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_17a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate10) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_18a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate11) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_19a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate12) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_20a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_06_13_10 (zzz_qry_Revzon_24_06_03_06_13_08 (xx),
                ' ** linked to zz_tbl_Revzon_17, zz_tbl_Revzon_18, for assetdate13) to zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_21a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' **********************************************
                ' ** Append Code: I.
                ' ** zz_tbl_Revzon_02_tmp
                ' ** zz_tbl_Revzon_03_tmp
                ' **********************************************
                gsCoZip = Right("000" & CStr(lngW), 3) & ".I"
                ' **********************************************

                ' ** Append zzz_qry_Revzon_24_06_03_06_13_23 (zz_tbl_Revzon_18, as
                ' ** zz_tbl_Revzon_03_tmp records; 5, excludes Zeroes) to zz_tbl_Revzon_03_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_25")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' **********************
                ' ** Check for closed!
                ' **********************
                ' ** zzz_qry_Revzon_24_06_03_06_13_24 (zz_tbl_Revzon_18, as
                ' ** zz_tbl_Revzon_02_tmp records; 13, includes Zeroes), just dp_closed = True; 0!
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_27")
                Set rst = qdf.OpenRecordset
                If rst.BOF = True And rst.EOF = True Then
                  ' ** Fine.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                Else
                  ' ** Make sure everybody knows about it!
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon._24_06_03_06_13_27 (zzz_qry_Revzon_24_06_03_06_13_24
                  ' ** (xx), just dp_closed = True) to zz_tbl_Revzon_06.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_28")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing
                End If

                ' ** Append zzz_qry_Revzon_24_06_03_06_13_24 (zz_tbl_Revzon_18, as
                ' ** zz_tbl_Revzon_02_tmp records; 13, includes Zeroes) to zz_tbl_Revzon_02_tmp.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_13_26")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

              Case 38&  '279. 10/24/2013    MULTI-LOT WITH 38 DP'S!

                ' ** Empty zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_19")
                qdf.Execute
                Set qdf = Nothing
                ' ** Empty zz_tbl_Revzon_18.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_20")
                qdf.Execute
                Set qdf = Nothing
                ' ** Empty zz_tbl_Revzon_19.
                Set qdf = .QueryDefs("zzz_qry_Revzon_29_21")
                qdf.Execute
                Set qdf = Nothing

                ' ** Append zzz_qry_Revzon_24_06_03_05 (zzz_qry_Revzon_24_06_03 (xx), linked to
                ' ** zzz_qry_Revzon_24_06_03_03 (xx), with .._new, .._gtot fields) to zz_tbl_Revzon_17.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_05_00a")
                qdf.Execute dbFailOnError
                Set qdf = Nothing

                ' ** zz_tbl_Revzon_17, all records, all fields.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_06_38_01")
                Set rst = qdf.OpenRecordset
                rst.MoveLast
                lngRecs = rst.RecordCount
                rst.MoveFirst
                dblTmp02 = rst![shareface_ws]
                lngTmp01 = 0&: dblTmp03 = 0#
                For lngX = 1& To lngRecs
                  dblTmp03 = dblTmp03 + rst![shareface_dp]
                  If dblTmp03 >= dblTmp02 Then
                    lngTmp01 = lngX
                    Exit For
                  End If
                  If lngX < lngRecs Then rst.MoveNext
                Next  ' ** lngX.
                If lngTmp01 > 0& Then
                  ' ** This gives us the minimum number of dp's that cover the ws.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                  ' ** I'd like to use pre-existing groups to handle it.
                  Select Case lngTmp01
                  Case 1&

                  Case 2&

                  Case 3&

                  Case 4&

                  Case 5&

                  Case 6&

                  End Select

                Else
                  ' ** Somethin's crazy!
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                  Stop
                End If


              Case Else
                Debug.Print "'MULTI-LOT WITH " & CStr(lngRecs) & " DP'S!"
                Stop
              End Select

            Else
              ' ** Looks OK.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing

              ' **********************************************
              ' ** Append Code: J.
              ' ** zz_tbl_Revzon_02_tmp
              ' ** zz_tbl_Revzon_03_tmp
              ' **********************************************
              gsCoZip = Right("000" & CStr(lngW), 3) & ".J"
              ' **********************************************

              ' ** Append zzz_qry_Revzon_24_06_03_08 (xx) to zz_tbl_Revzon_03_tmp.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_10")
              qdf.Execute dbFailOnError
              Set qdf = Nothing

              ' **********************
              ' ** Check for closed!
              ' **********************
              ' ** zzz_qry_Revzon_24_06_03_09 (xx), just dp_closed = True.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_12")
              Set rst = qdf.OpenRecordset
              If rst.BOF = True And rst.EOF = True Then
                ' ** Fine.
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
              Else
                ' ** Make sure everybody knows about it!
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_06_03_12 (xx) to zz_tbl_Revzon_06.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_13")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
              End If

              ' ** Append zzz_qry_Revzon_24_06_03_09 (xx) to zz_tbl_Revzon_02_tmp.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_06_03_11")
              qdf.Execute dbFailOnError
              Set qdf = Nothing

            End If
          Else
            ' ** Not dealt with.
            rst.Close
            Set rst = Nothing
            Set qdf = Nothing
            Stop
          End If
        End If

        gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

        If blnMatchPair1 = False And blnMultiLot = False Then

          ' ** zzz_qry_Revzon_24_06 (xx), grouped, by accountno, assetno, with cnt > 1, multiple dp's for the same ws.
          Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_01")
          Set rst = qdf.OpenRecordset
          If rst.BOF = True And rst.EOF = True Then
            ' ** 1-to-1, good!
            rst.Close
            Set rst = Nothing
            Set qdf = Nothing
          Else
            ' ** Multiple dp's available.
            rst.Close
            Set rst = Nothing
            Set qdf = Nothing

            ' ** Empty zz_tbl_Revzon_14.
            Set qdf = .QueryDefs("zzz_qry_Revzon_29_16")
            qdf.Execute
            Set qdf = Nothing
            ' ** Append zzz_qry_Revzon_24_07_02 (xx) to zz_tbl_Revzon_14.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_02_0")
            qdf.Execute
            Set qdf = Nothing

            ' ** Empty zz_tbl_Revzon_10.
            Set qdf = .QueryDefs("zzz_qry_Revzon_29_12")
            qdf.Execute
            Set qdf = Nothing
            ' ** Append zzz_qry_Revzon_24_07_02_00_04 (xx) to zz_tbl_Revzon_10.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_02_00a")    'SYSTEM RESOURCES EXCEEDED!  FIXED!
            qdf.Execute
            Set qdf = Nothing

            ' ** to zz_tbl_Revzon_14 (xx), linked to zz_tbl_Revzon_10, just Sx1, matching pairs.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_03")
            Set rst = qdf.OpenRecordset
            If rst.BOF = True And rst.EOF = True Then
              ' ** Fine.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
            Else
              ' ** Not dealt with.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
              Stop
            End If

            ' ** to zz_tbl_Revzon_14 (xx), linked to zz_tbl_Revzon_10, just Sx2, dp covers ws.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_04")
            Set rst = qdf.OpenRecordset
            If rst.BOF = True And rst.EOF = True Then
              ' ** OK.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
            Else
              ' ** It's covered, and extras (if any) excluded in zzz_qry_Revzon_24_08
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
              'Stop
            End If

            ' ** to zz_tbl_Revzon_14 (xx), linked to zz_tbl_Revzon_10, just Sx3, ws not covered.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_05")
            Set rst = qdf.OpenRecordset
            If rst.BOF = True And rst.EOF = True Then
              ' ** OK.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
            Else
              ' ** Not dealt with.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing

              ' ** Empty zz_tbl_Revzon_08.
              Set qdf = .QueryDefs("zzz_qry_Revzon_29_10")
              qdf.Execute
              Set qdf = Nothing
              ' ** Append zzz_qry_Revzon_24_07_05 (xx) to zz_tbl_Revzon_08.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_05_00a")
              qdf.Execute dbFailOnError
              Set qdf = Nothing

              ' **********************************************
              ' ** Append Code: K.
              ' ** zz_tbl_Revzon_02_tmp
              ' ** zz_tbl_Revzon_03_tmp
              ' **********************************************
              gsCoZip = Right("000" & CStr(lngW), 3) & ".K"
              ' **********************************************

              ' ** Append zzz_qry_Revzon_24_07_06_09 (xx) to zz_tbl_Revzon_03_tmp.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_06_11")
              qdf.Execute dbFailOnError
              Set qdf = Nothing

              ' **********************
              ' ** Check for closed!
              ' **********************
              ' ** zzz_qry_Revzon_24_07_06_10 (xx), just dp_closed = True.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_06_13")
              Set rst = qdf.OpenRecordset
              If rst.BOF = True And rst.EOF = True Then
                ' ** Fine.
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
              Else
                ' ** Make sure everybody knows about it!
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
                ' ** Append zzz_qry_Revzon_24_07_06_13 (xx) to zz_tbl_Revzon_06.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_06_14")
                qdf.Execute dbFailOnError
                Set qdf = Nothing
              End If

              ' ** Append zzz_qry_Revzon_24_07_06_10 (xx) to zz_tbl_Revzon_02_tmp.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_07_06_12")
              qdf.Execute dbFailOnError
              Set qdf = Nothing

            End If
          End If

        End If  ' ** blnMatchPair1.

        blnMatchPair2 = False: blnMatchPair3 = False: blnMultiWSCovered = False
        gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

        If blnMultiWS = True Then

          ' ** Empty zz_tbl_Revzon_10.
          Set qdf = .QueryDefs("zzz_qry_Revzon_29_12")
          qdf.Execute
          Set qdf = Nothing
          ' ** Append zzz_qry_Revzon_24_07_02_00 (zzz_qry_Revzon_24_09_01 (xx), grouped, with Max(Sx1,Sx2,Sx3)) to zz_tbl_Revzon_10.
          Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_01_00a")
          qdf.Execute
          Set qdf = Nothing

          ' ** zzz_qry_Revzon_24_09_01 (xx), linked to zz_tbl_Revzon_10, just Sx1, matching pairs.
          Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_02")
          Set rst = qdf.OpenRecordset
          If rst.BOF = True And rst.EOF = True Then
            ' ** OK.
            rst.Close
            Set rst = Nothing
            Set qdf = Nothing
          Else
            ' ** Hot diggity!
            blnMatchPair3 = True
            rst.Close
            Set rst = Nothing
            Set qdf = Nothing

            ' ** Empty zz_tbl_Revzon_04.
            Set qdf = .QueryDefs("zzz_qry_Revzon_29_06")
            qdf.Execute
            Set qdf = Nothing
            ' ** Empty zz_tbl_Revzon_13.
            Set qdf = .QueryDefs("zzz_qry_Revzon_29_15")
            qdf.Execute
            Set qdf = Nothing
            ' ** Append zzz_qry_Revzon_24_09_02_05 (zzz_qry_Revzon_24_09_02_04 (xx),
            ' ** linked back to zz_tbl_Revzon_02) to zz_tbl_Revzon_13.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_02_05a")
            qdf.Execute
            Set qdf = Nothing
            ' ** Append zzz_qry_Revzon_24_09_02_06 (zzz_qry_Revzon_24_09_02_04 (xx),
            ' ** linked back to zz_tbl_Revzon_02_tmp) to zz_tbl_Revzon_13.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_02_06a")
            qdf.Execute
            Set qdf = Nothing
            ' ** Append zzz_qry_Revzon_24_09_02_11 (xx) to zz_tbl_Revzon_04.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_02_11a")
            qdf.Execute dbFailOnError
            Set qdf = Nothing
            ' ** Append zzz_qry_Revzon_24_09_02_13 (xx) to zz_tbl_Revzon_04.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_02_13a")    'SYSTEM RESOURCES EXCEEDED!  FIXED!
            qdf.Execute dbFailOnError
            Set qdf = Nothing

            ' **********************************************
            ' ** Append Code: L.
            ' ** zz_tbl_Revzon_02_tmp
            ' ** zz_tbl_Revzon_03_tmp
            ' **********************************************
            gsCoZip = Right("000" & CStr(lngW), 3) & ".L"
            ' **********************************************

            ' ** Append zzz_qry_Revzon_24_09_02_10 (xx) to zz_tbl_Revzon_03_tmp.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_02_14")    'THE CHANGES YOU REQUESTED...  FIXED!
            qdf.Execute dbFailOnError
            Set qdf = Nothing

            ' **********************
            ' ** Check for closed!
            ' **********************
            ' ** zzz_qry_Revzon_24_09_02_11 (xx), just dp_closed = True.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_02_16")
            Set rst = qdf.OpenRecordset
            If rst.BOF = True And rst.EOF = True Then
              ' ** Fine.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
            Else
              ' ** Make sure everybody knows about it!
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
              ' ** Append zzz_qry_Revzon_24_09_02_16 (xx) to zz_tbl_Revzon_06.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_02_17")
              qdf.Execute dbFailOnError
              Set qdf = Nothing
            End If

            ' ** Append zz_tbl_Revzon_04 to zz_tbl_Revzon_02_tmp.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_02_15")
            qdf.Execute dbFailOnError
            Set qdf = Nothing

          End If

          gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

          If blnMatchPair3 = False Then

            ' ** zzz_qry_Revzon_24_09_01 (xx), linked to zz_tbl_Revzon_10, just Sx2, dp covers ws.
            Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03")
            Set rst = qdf.OpenRecordset
            If rst.BOF = True And rst.EOF = True Then
              ' ** OK.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing
            Else
              ' ** ws is covered.
              rst.Close
              Set rst = Nothing
              Set qdf = Nothing

              ' ** Empty zz_tbl_Revzon_11.
              Set qdf = .QueryDefs("zzz_qry_Revzon_29_13")
              qdf.Execute
              Set qdf = Nothing
              ' ** Append zzz_qry_Revzon_24_09_03_04_00 (zzz_qry_Revzon_24_09_01 (xx), grouped, with Max(Sx1,Sx2,Sx3)) to zz_tbl_Revzon_11.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_04_00a")
              qdf.Execute
              Set qdf = Nothing

              ' ** zzz_qry_Revzon_24_09_03_04 (xx), linked to zz_tbl_Revzon_11, just Sx1, matching pairs.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_05")
              Set rst = qdf.OpenRecordset
              If rst.BOF = True And rst.EOF = True Then
                ' ** OK.
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
              Else
                ' ** Wipe-Out!
                blnMatchPair2 = True
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
              End If

              If blnMatchPair2 = False Then

                ' ** zzz_qry_Revzon_24_09_03_04 (xx), linked to zz_tbl_Revzon_11, just Sx2, dp covers ws.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06")
                Set rst = qdf.OpenRecordset
                If rst.BOF = True And rst.EOF = True Then
                  ' ** OK.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                Else
                  ' ** Not dealt with.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing

                  ' ** Empty zz_tbl_Revzon_04.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_29_06")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Empty zz_tbl_Revzon_07.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_29_09")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Empty zz_tbl_Revzon_08.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_29_10")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Empty zz_tbl_Revzon_13.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_29_15")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_03(xx) to zz_tbl_Revzon_08.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_03a")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_14 (xx) to zz_tbl_Revzon_04.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_14a")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_08 (zzz_qry_Revzon_24_09_03_06_07 (xx),
                  ' ** linked back to zz_tbl_Revzon_02) to zz_tbl_Revzon_13.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_08a")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_09 (zzz_qry_Revzon_24_09_03_06_07 (xx),
                  ' ** linked back to zz_tbl_Revzon_02_tmp) to zz_tbl_Revzon_13.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_09a")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_10 (xx), not in zzz_qry_Revzon_24_09_03_06 (xx), to zz_tbl_Revzon_07.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_15")    'SYSTEM RESOURCES EXCEEDED!  FIXED!
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_17 (xx) to zz_tbl_Revzon_04.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_17a")
                  qdf.Execute
                  Set qdf = Nothing

                  ' ** Update zzz_qry_Revzon_24_09_03_06_17_01 (zz_tbl_Revzon_04, with dp_comp_new).
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_17_02")
                  qdf.Execute
                  Set qdf = Nothing

                  ' **********************************************
                  ' ** Append Code: M.
                  ' ** zz_tbl_Revzon_02_tmp
                  ' ** zz_tbl_Revzon_03_tmp
                  ' **********************************************
                  gsCoZip = Right("000" & CStr(lngW), 3) & ".M"
                  ' **********************************************

                  ' ** Append zzz_qry_Revzon_24_09_03_06_13 (xx) to zz_tbl_Revzon_03_tmp.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_18")
                  qdf.Execute
                  Set qdf = Nothing

                  ' **********************
                  ' ** Check for closed!
                  ' **********************
                  ' ** zzz_qry_Revzon_24_09_03_06_14 (xx), just dp_closed = True.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_29")
                  Set rst = qdf.OpenRecordset
                  If rst.BOF = True And rst.EOF = True Then
                    ' ** Fine.
                    rst.Close
                    Set rst = Nothing
                    Set qdf = Nothing
                  Else
                    ' ** Make sure everybody knows about it!
                    rst.Close
                    Set rst = Nothing
                    Set qdf = Nothing
                    ' ** Append zzz_qry_Revzon_24_09_03_06_29 (xx) to zz_tbl_Revzon_06.
                    Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_30")
                    qdf.Execute dbFailOnError
                    Set qdf = Nothing
                  End If

                  ' ** Append zz_tbl_Revzon_04 to zz_tbl_Revzon_02_tmp.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_19")
                  qdf.Execute
                  Set qdf = Nothing

                  ' ** Empty zz_tbl_Revzon_04.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_29_06")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Empty zz_tbl_Revzon_07.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_29_09")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Empty zz_tbl_Revzon_08.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_29_10")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_20 (xx) to zz_tbl_Revzon_08.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_20a")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_23 (xx) to zz_tbl_Revzon_04.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_23a")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_10 (xx), not in zzz_qry_Revzon_24_09_03_06 (xx), to zz_tbl_Revzon_07.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_24")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_06_26 (xx) to zz_tbl_Revzon_04.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_26a")
                  qdf.Execute
                  Set qdf = Nothing

                  gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

                  ' **********************************************
                  ' ** Append Code: N.
                  ' ** zz_tbl_Revzon_02_tmp
                  ' ** zz_tbl_Revzon_03_tmp
                  ' **********************************************
                  gsCoZip = Right("000" & CStr(lngW), 3) & ".N"
                  ' **********************************************

                  ' ** Append zzz_qry_Revzon_24_09_03_06_22 (xx) to zz_tbl_Revzon_03_tmp.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_27")
                  qdf.Execute
                  Set qdf = Nothing

                  ' **********************
                  ' ** Check for closed!
                  ' **********************
                  ' ** zz_tbl_Revzon_04, just dp_closed = True.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_31")
                  Set rst = qdf.OpenRecordset
                  If rst.BOF = True And rst.EOF = True Then
                    ' ** Fine.
                    rst.Close
                    Set rst = Nothing
                    Set qdf = Nothing
                  Else
                    ' ** Make sure everybody knows about it!
                    rst.Close
                    Set rst = Nothing
                    Set qdf = Nothing
                    ' ** Append zzz_qry_Revzon_24_09_03_06_31 (xx) to zz_tbl_Revzon_06.
                    Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_32")
                    qdf.Execute dbFailOnError
                    Set qdf = Nothing
                  End If

                  ' ** Append zz_tbl_Revzon_04 to zz_tbl_Revzon_02_tmp.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_06_28")
                  qdf.Execute
                  Set qdf = Nothing
                  blnMultiWSCovered = True

                End If

              End If  ' ** blnMatchPair2.

              gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

              If blnMatchPair2 = False And blnMultiWSCovered = False Then

                ' ** zzz_qry_Revzon_24_09_03_04 (xx), linked to zz_tbl_Revzon_11, just Sx3, ws not covered.
                Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07")
                Set rst = qdf.OpenRecordset
                If rst.BOF = True And rst.EOF = True Then
                  ' ** OK.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing
                Else
                  ' ** Not dealt with.
                  rst.Close
                  Set rst = Nothing
                  Set qdf = Nothing

                  ' ** Empty zz_tbl_Revzon_09.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_29_11")
                  qdf.Execute
                  Set qdf = Nothing
                  ' ** Append zzz_qry_Revzon_24_09_03_07 (xx) to zz_tbl_Revzon_09.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07_00a")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing

                  ' **********************************************
                  ' ** Append Code: O.
                  ' ** zz_tbl_Revzon_02_tmp
                  ' ** zz_tbl_Revzon_03_tmp
                  ' **********************************************
                  gsCoZip = Right("000" & CStr(lngW), 3) & ".O"
                  ' **********************************************

                  ' ** Append zzz_qry_Revzon_24_09_03_07_12 (xx) to zz_tbl_Revzon_03_tmp.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07_14")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing

                  ' **********************
                  ' ** Check for closed!
                  ' **********************
                  ' ** zzz_qry_Revzon_24_09_03_07_13, just dp_closed = True.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07_23")
                  Set rst = qdf.OpenRecordset
                  If rst.BOF = True And rst.EOF = True Then
                    ' ** Fine.
                    rst.Close
                    Set rst = Nothing
                    Set qdf = Nothing
                  Else
                    ' ** Make sure everybody knows about it!
                    rst.Close
                    Set rst = Nothing
                    Set qdf = Nothing
                    ' ** Append zzz_qry_Revzon_24_09_03_07_23 (xx) to zz_tbl_Revzon_06.
                    Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07_24")
                    qdf.Execute dbFailOnError
                    Set qdf = Nothing
                  End If

                  ' ** Append zzz_qry_Revzon_24_09_03_07_13 (xx) to zz_tbl_Revzon_02_tmp.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07_15")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing

                  gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

                  ' **********************************************
                  ' ** Append Code: P.
                  ' ** zz_tbl_Revzon_02_tmp
                  ' ** zz_tbl_Revzon_03_tmp
                  ' **********************************************
                  gsCoZip = Right("000" & CStr(lngW), 3) & ".P"
                  ' **********************************************

                  ' ** Append zzz_qry_Revzon_24_09_03_07_19 (xx) to zz_tbl_Revzon_03_tmp.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07_21")    'THE CHANGES YOU REQUESTED...  FIXED!
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing

                  ' **********************
                  ' ** Check for closed!
                  ' **********************
                  ' ** zzz_qry_Revzon_24_09_03_07_20 (xx), just dp_closed = True.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07_25")
                  Set rst = qdf.OpenRecordset
                  If rst.BOF = True And rst.EOF = True Then
                    ' ** Fine.
                    rst.Close
                    Set rst = Nothing
                    Set qdf = Nothing
                  Else
                    ' ** Make sure everybody knows about it!
                    rst.Close
                    Set rst = Nothing
                    Set qdf = Nothing
                    ' ** Append zzz_qry_Revzon_24_09_03_07_25 (xx) to zz_tbl_Revzon_06.
                    Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07_26")
                    qdf.Execute dbFailOnError
                    Set qdf = Nothing
                  End If

                  ' ** Append zzz_qry_Revzon_24_09_03_07_20 (xx) to zz_tbl_Revzon_02_tmp.
                  Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_03_07_22")
                  qdf.Execute dbFailOnError
                  Set qdf = Nothing

                End If
              End If  ' ** blnMatchPair2, blnMultiWSCovered.
            End If  ' ** dp covers ws.

            gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

            If blnMultiWSCovered = False Then
              ' ** zzz_qry_Revzon_24_09_01 (xx), linked to zz_tbl_Revzon_10, just Sx3, ws not covered.
              Set qdf = .QueryDefs("zzz_qry_Revzon_24_09_04")
              Set rst = qdf.OpenRecordset
              If rst.BOF = True And rst.EOF = True Then
                ' ** OK.
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
              Else
                ' ** Not dealt with.
                rst.Close
                Set rst = Nothing
                Set qdf = Nothing
                Stop
              End If
            End If  ' ** blnMultiWSCovered.

          End If  ' ** blnMatchPair3.

        End If  ' ** blnMultiWS.

      End If  ' ** ws's.

      gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

      ' ** Empty zz_tbl_Revzon_04.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_06")
      qdf.Execute
      Set qdf = Nothing

      ' ** Matching Pairs
      gsCoPhone = Right("000" & CStr(lngW), 3) & ".K1"

      ' ** Append zzz_qry_Revzon_24_14 (xx) to zz_tbl_Revzon_04.
      Set qdf = .QueryDefs("zzz_qry_Revzon_24_14a")
      qdf.Execute
      Set qdf = Nothing

      ' ** Empty zz_tbl_Revzon_15.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_17")
      qdf.Execute
      Set qdf = Nothing
      ' ** Append zzz_qry_Revzon_24_15_01 (xx) to zz_tbl_Revzon_15.
      Set qdf = .QueryDefs("zzz_qry_Revzon_24_15_01a")
      qdf.Execute dbFailOnError
      Set qdf = Nothing

      ' ** Matching Pair avg_upd.
      gsCoPhone = Right("000" & CStr(lngW), 3) & ".K2"

      ' ** Append zzz_qry_Revzon_24_16 (xx) to zz_tbl_Revzon_04.
      Set qdf = .QueryDefs("zzz_qry_Revzon_24_16a")  'SYSTEM RESOURCES EXCEEDED!
      qdf.Execute
      Set qdf = Nothing

      If blnWSCovered = True Then

        ' ** DP covers WS.
        gsCoPhone = Right("000" & CStr(lngW), 3) & ".K3"

        ' ** Append zzz_qry_Revzon_24_18 (xx) to zz_tbl_Revzon_04.
        Set qdf = .QueryDefs("zzz_qry_Revzon_24_18a")
        qdf.Execute
        Set qdf = Nothing

      End If

      gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

      ' ** Empty zz_tbl_Revzon_12.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_14")
      qdf.Execute
      Set qdf = Nothing

      ' ** Matching Pairs.
      gsCoState = Right("000" & CStr(lngW), 3) & ".K1"

      ' ** Append zzz_qry_Revzon_24_13 (zzz_qry_Revzon_24_12 (xx), as zz_tbl_Revzon_03 record) to zz_tbl_Revzon_12.
      Set qdf = .QueryDefs("zzz_qry_Revzon_24_13a")
      qdf.Execute
      Set qdf = Nothing

      ' ** DP covers WS.
      gsCoState = Right("000" & CStr(lngW), 3) & ".K2"

      ' ** Append zzz_qry_Revzon_24_17 (zzz_qry_Revzon_24_08 (xx), as zz_tbl_Revzon_03 records) to zz_tbl_Revzon_12.
      Set qdf = .QueryDefs("zzz_qry_Revzon_24_17a")
      qdf.Execute
      Set qdf = Nothing

      gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

      ' ** zz_tbl_Revzon_12, just Len(dp_comp) <> 15, with dp_comp_new.
      Set qdf = .QueryDefs("zzz_qry_Revzon_25_01_01")
      Set rst = qdf.OpenRecordset
      If rst.BOF = True And rst.EOF = True Then
        ' ** No discrepancies.
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
      Else
        ' ** Discrepancies found in dp_comp.
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
        ' ** Update zzz_qry_Revzon_25_01_01 (zz_tbl_Revzon_12, just Len(dp_comp) <> 15, with dp_comp_new).
        Set qdf = .QueryDefs("zzz_qry_Revzon_25_01_02")
        qdf.Execute
        Set qdf = Nothing
      End If

      ' ** zz_tbl_Revzon_04, just Len(dp_comp) <> 15, with dp_comp_new.
      Set qdf = .QueryDefs("zzz_qry_Revzon_25_02_01")
      Set rst = qdf.OpenRecordset
      If rst.BOF = True And rst.EOF = True Then
        ' ** No discrepancies.
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
      Else
        ' ** Discrepancies found in dp_comp.
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
        ' ** Update zzz_qry_Revzon_25_02_01 (zz_tbl_Revzon_04, just Len(dp_comp) <> 15, with dp_comp_new).
        Set qdf = .QueryDefs("zzz_qry_Revzon_25_02_02")
        qdf.Execute
        Set qdf = Nothing
      End If

      ' **********************************************
      ' ** Append Code: Q.
      ' ** zz_tbl_Revzon_02_tmp
      ' ** zz_tbl_Revzon_03_tmp
      ' **********************************************
      gsCoZip = Right("000" & CStr(lngW), 3) & ".Q"
      ' **********************************************

      ' ** Append zzz_qry_Revzon_25_01 (xx) to zz_tbl_Revzon_03_tmp.
      Set qdf = .QueryDefs("zzz_qry_Revzon_25_03")    'SYSTEM RESOURCES EXCEEDED!  FIXED!
      qdf.Execute
      Set qdf = Nothing

      ' **********************
      ' ** Check for closed!
      ' **********************
      ' ** zzz_qry_Revzon_25_02 (xx), just dp_closed = True.
      Set qdf = .QueryDefs("zzz_qry_Revzon_25_05")
      Set rst = qdf.OpenRecordset
      If rst.BOF = True And rst.EOF = True Then
        ' ** Fine.
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
      Else
        ' ** Make sure everybody knows about it!
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
        ' ** Append zzz_qry_Revzon_25_05 (xx) to zz_tbl_Revzon_06.
        Set qdf = .QueryDefs("zzz_qry_Revzon_25_06")
        qdf.Execute dbFailOnError
        Set qdf = Nothing
      End If

      ' ** Append zzz_qry_Revzon_25_02 (xx) to zz_tbl_Revzon_02_tmp.
      Set qdf = .QueryDefs("zzz_qry_Revzon_25_04")
      qdf.Execute
      Set qdf = Nothing

      gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

      ' ** zzz_qry_Revzon_24_02_02 (xx), not in zz_tbl_Revzon_03_tmp, remaining ws's for date.
      Set qdf = .QueryDefs("zzz_qry_Revzon_26_01")
      Set rst = qdf.OpenRecordset
      If rst.BOF = True And rst.EOF = True Then
        ' ** OK.
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
      Else
        ' ** But wait, there's more!
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing

        ' **********************************************
        ' ** Append Code: R.
        ' ** zz_tbl_Revzon_02_tmp
        ' ** zz_tbl_Revzon_03_tmp
        ' **********************************************
        gsCoZip = Right("000" & CStr(lngW), 3) & ".R"
        ' **********************************************

        ' **********************
        ' ** Check for closed!
        ' **********************
        ' ** zzz_qry_Revzon_27_02 (xx), just dp_closed = True.
        Set qdf = .QueryDefs("zzz_qry_Revzon_27_05")
        Set rst = qdf.OpenRecordset
        If rst.BOF = True And rst.EOF = True Then
          ' ** Fine.
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
        Else
          ' ** Make sure everybody knows about it!
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
          ' ** Append zzz_qry_Revzon_27_05 (xx) to zz_tbl_Revzon_06.
          Set qdf = .QueryDefs("zzz_qry_Revzon_27_06")
          qdf.Execute dbFailOnError
          Set qdf = Nothing
        End If

        ' ** Append zzz_qry_Revzon_27_02 (xx) to zz_tbl_Revzon_02_tmp.
        Set qdf = .QueryDefs("zzz_qry_Revzon_27_03")
        qdf.Execute dbFailOnError
        Set qdf = Nothing
        ' ** Append zzz_qry_Revzon_27_01 (xx) to zz_tbl_Revzon_03_tmp.
        Set qdf = .QueryDefs("zzz_qry_Revzon_27_04")
        qdf.Execute dbFailOnError
        Set qdf = Nothing

      End If  ' ** Remaining ws's.

      gsCoCity = vbNullString: gsCoZip = vbNullString: gsCoPhone = vbNullString: gsCoState = vbNullString

      ' ** Update zzz_qry_Revzon_28_00_01 (zz_tbl_Revzon_03_tmp, with dp_comp_new).
      Set qdf = .QueryDefs("zzz_qry_Revzon_28_08_02")
      qdf.Execute dbFailOnError
      Set qdf = Nothing

      ' ** Update zzz_qry_Revzon_28_01_01 (zz_tbl_Revzon_02_tmp, with dp_comp_new).
      Set qdf = .QueryDefs("zzz_qry_Revzon_28_09_02")    '17.  09/19/2012  THE CHANGES YOU REQUESTED...
      qdf.Execute dbFailOnError
      Set qdf = Nothing

      ' ** zz_tbl_Revzon_03_tmp, grouped, by journalno_ws, dp_cnt, with cnt > 1.
      Set qdf = .QueryDefs("zzz_qry_Revzon_28_08_05")
      Set rst = qdf.OpenRecordset
      If rst.BOF = True And rst.EOF = True Then
        ' ** No bad dp_cnt's.
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
      Else
        ' ** Some numbers are off, and it'll cause an error!
        rst.Close
        Set rst = Nothing
        Set qdf = Nothing
        ' ** Update zzz_qry_Revzon_28_00_07 (zz_tbl_Revzon_03_tmp, with DLookups() to
        ' ** zzz_qry_Revzon_28_00_06 (zzz_qry_Revzon_28_00_05 (zz_tbl_Revzon_03_tmp,
        ' ** grouped, by journalno_ws, dp_cnt, with cnt > 1), linked back to
        ' ** zz_tbl_Revzon_03_tmp, for rz3_id2, with dp_cnt_new, dp_comp_new)).
        Set qdf = .QueryDefs("zzz_qry_Revzon_28_08_08")
        qdf.Execute
        Set qdf = Nothing
      End If

      ' **********************************************
      ' ** Append Code: S.
      ' ** zz_tbl_Revzon_02
      ' ** zz_tbl_Revzon_03
      ' **********************************************
      gsCoCity = Right("000" & CStr(lngW), 3) & ".S"

      ' ** Append zz_tbl_Revzon_03_tmp to zz_tbl_Revzon_03.
      Set qdf = .QueryDefs("zzz_qry_Revzon_28_01")    '31. 10/09/2012  THE CHANGES YOU REQUESTED...
      qdf.Execute dbFailOnError
      Set qdf = Nothing

      ' ** Append zz_tbl_Revzon_02_tmp to zz_tbl_Revzon_02.
      Set qdf = .QueryDefs("zzz_qry_Revzon_28_02")    '52.  11/09/2012  THE CHANGES YOU REQUESTED...
      qdf.Execute dbFailOnError                       '267. 10/07/2013  THE CHANGES YOU REQUESTED...
      Set qdf = Nothing                               '270. 10/10/2013  THE CHANGES YOU REQUESTED...
                                                      'THE LAST 2 HAD 2ND COPIES WITH AN EARLY DP_COMP!
      ' ** Empty zz_tbl_Revzon_02_tmp.                'WHERE AND WHY IS IT GETTING THESE?
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_03")
      qdf.Execute dbFailOnError
      Set qdf = Nothing
      ' ** Empty zz_tbl_Revzon_03_tmp.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_05")
      qdf.Execute dbFailOnError
      Set qdf = Nothing
      ' ** Empty zz_tbl_Revzon_04.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_06")
      qdf.Execute
      Set qdf = Nothing
      ' ** Empty zz_tbl_Revzon_07.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_09")
      qdf.Execute
      Set qdf = Nothing
      ' ** Empty zz_tbl_Revzon_08.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_10")
      qdf.Execute
      Set qdf = Nothing
      ' ** Empty zz_tbl_Revzon_09.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_11")
      qdf.Execute
      Set qdf = Nothing
      ' ** Empty zz_tbl_Revzon_14.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_16")
      qdf.Execute
      Set qdf = Nothing
      ' ** Empty zz_tbl_Revzon_17.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_19")
      qdf.Execute
      Set qdf = Nothing
      ' ** Empty zz_tbl_Revzon_18.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_20")
      qdf.Execute
      Set qdf = Nothing
      ' ** Empty zz_tbl_Revzon_19.
      Set qdf = .QueryDefs("zzz_qry_Revzon_29_21")
      qdf.Execute
      Set qdf = Nothing
      DoEvents

If lngW = 150& Then
Stop
End If

    Next  ' ** lngW.

    .Close
  End With
  Set dbs = Nothing

  Beep
  Debug.Print "'DONE!"

'1.   THIS DATE: 08/16/2012
'DP'S: 386
'WS'S: 15
'2.   THIS DATE: 08/17/2012
'DP'S: 0
'WS'S: 12
'3.   THIS DATE: 08/20/2012
'DP'S: 0
'WS'S: 1
'4.   THIS DATE: 08/21/2012
'DP'S: 1
'WS'S: 13
'5.   THIS DATE: 08/23/2012
'DP'S: 12
'WS'S: 12
'6.   THIS DATE: 08/24/2012
'DP'S: 0
'WS'S: 22
'7.   THIS DATE: 08/27/2012
'DP'S: 0
'WS'S: 12
'8.   THIS DATE: 08/29/2012
'DP'S: 0
'WS'S: 9
'9.   THIS DATE: 08/31/2012
'DP'S: 14
'WS'S: 7
'10.  THIS DATE: 09/05/2012
'DP'S: 0
'WS'S: 3
'11.  THIS DATE: 09/06/2012
'DP'S: 0
'WS'S: 5
'12.  THIS DATE: 09/07/2012
'DP'S: 12
'WS'S: 17
'13.  THIS DATE: 09/12/2012
'DP'S: 9
'WS'S: 33
'14.  THIS DATE: 09/14/2012
'DP'S: 0
'WS'S: 14
'15.  THIS DATE: 09/17/2012
'DP'S: 0
'WS'S: 9
'16.  THIS DATE: 09/18/2012
'DP'S: 1
'WS'S: 8
'17.  THIS DATE: 09/19/2012
'DP'S: 14
'WS'S: 4
'18.  THIS DATE: 09/20/2012
'DP'S: 0
'WS'S: 16
'19.  THIS DATE: 09/21/2012
'DP'S: 0
'WS'S: 6
'20.  THIS DATE: 09/24/2012
'DP'S: 0
'WS'S: 10
'21.  THIS DATE: 09/25/2012
'DP'S: 0
'WS'S: 26
'22.  THIS DATE: 09/26/2012
'DP'S: 15
'WS'S: 9
'23.  THIS DATE: 09/27/2012
'DP'S: 5
'WS'S: 12
'24.  THIS DATE: 09/28/2012
'DP'S: 0
'WS'S: 1
'25.  THIS DATE: 10/01/2012
'DP'S: 0
'WS'S: 14
'26.  THIS DATE: 10/03/2012
'DP'S: 15
'WS'S: 0
'27.  THIS DATE: 10/04/2012
'DP'S: 0
'WS'S: 3
'28.  THIS DATE: 10/05/2012
'DP'S: 0
'WS'S: 6
'29.  THIS DATE: 10/08/2012
'DP'S: 2
'WS'S: 0
'30.  THIS DATE: 10/09/2012
'DP'S: 0
'WS'S: 9
'31.  THIS DATE: 10/10/2012
'DP'S: 0
'WS'S: 4
'32.  THIS DATE: 10/11/2012
'DP'S: 1
'WS'S: 24
'33.  THIS DATE: 10/12/2012
'DP'S: 0
'WS'S: 5
'34.  THIS DATE: 10/15/2012
'DP'S: 1
'WS'S: 12
'35.  THIS DATE: 10/16/2012
'DP'S: 0
'WS'S: 17
'36.  THIS DATE: 10/17/2012
'DP'S: 13
'WS'S: 13
'37.  THIS DATE: 10/18/2012
'DP'S: 0
'WS'S: 7
'38.  THIS DATE: 10/19/2012
'DP'S: 0
'WS'S: 3
'39.  THIS DATE: 10/22/2012
'DP'S: 3
'WS'S: 4
'40.  THIS DATE: 10/23/2012
'DP'S: 0
'WS'S: 11
'41.  THIS DATE: 10/24/2012
'DP'S: 5
'WS'S: 42
'42.  THIS DATE: 10/25/2012
'DP'S: 19
'WS'S: 18
'43.  THIS DATE: 10/26/2012
'DP'S: 0
'WS'S: 9
'44.  THIS DATE: 10/29/2012
'DP'S: 0
'WS'S: 6
'45.  THIS DATE: 10/31/2012
'DP'S: 29
'WS'S: 28
'46.  THIS DATE: 11/01/2012
'DP'S: 0
'WS'S: 2
'47.  THIS DATE: 11/02/2012
'DP'S: 0
'WS'S: 1
'48.  THIS DATE: 11/05/2012
'DP'S: 0
'WS'S: 28
'49.  THIS DATE: 11/06/2012
'DP'S: 1
'WS'S: 12
'50.  THIS DATE: 11/07/2012
'DP'S: 9
'WS'S: 3
'51.  THIS DATE: 11/08/2012
'DP'S: 0
'WS'S: 12
'52.  THIS DATE: 11/09/2012
'DP'S: 0
'WS'S: 6
'53.  THIS DATE: 11/12/2012
'DP'S: 0
'WS'S: 5
'54.  THIS DATE: 11/13/2012
'DP'S: 0
'WS'S: 1
'55.  THIS DATE: 11/14/2012
'DP'S: 0
'WS'S: 161
'56.  THIS DATE: 11/15/2012
'DP'S: 13
'WS'S: 11
'57.  THIS DATE: 11/16/2012
'DP'S: 1
'WS'S: 10
'58.  THIS DATE: 11/19/2012
'DP'S: 0
'WS'S: 10
'59.  THIS DATE: 11/20/2012
'DP'S: 0
'WS'S: 18
'60.  THIS DATE: 11/21/2012
'DP'S: 6
'WS'S: 6
'61.  THIS DATE: 11/22/2012
'DP'S: 0
'WS'S: 1
'62.  THIS DATE: 11/26/2012
'DP'S: 2
'WS'S: 23
'63.  THIS DATE: 11/27/2012
'DP'S: 0
'WS'S: 6
'64.  THIS DATE: 11/28/2012
'DP'S: 9
'WS'S: 8
'65.  THIS DATE: 11/29/2012
'DP'S: 0
'WS'S: 11
'66.  THIS DATE: 11/30/2012
'DP'S: 2
'WS'S: 7
'67.  THIS DATE: 12/03/2012
'DP'S: 0
'WS'S: 10
'68.  THIS DATE: 12/04/2012
'DP'S: 0
'WS'S: 10
'69.  THIS DATE: 12/05/2012
'DP'S: 13
'WS'S: 15
'70.  THIS DATE: 12/06/2012
'DP'S: 0
'WS'S: 7
'71.  THIS DATE: 12/07/2012
'DP'S: 0
'WS'S: 3
'72.  THIS DATE: 12/10/2012
'DP'S: 3
'WS'S: 11
'73.  THIS DATE: 12/12/2012
'DP'S: 12
'WS'S: 12
'74.  THIS DATE: 12/13/2012
'DP'S: 3
'WS'S: 0
'75.  THIS DATE: 12/14/2012
'DP'S: 0
'WS'S: 25
'76.  THIS DATE: 12/16/2012
'DP'S: 0
'WS'S: 1
'77.  THIS DATE: 12/17/2012
'DP'S: 3
'WS'S: 5
'78.  THIS DATE: 12/18/2012
'DP'S: 1
'WS'S: 33
'79.  THIS DATE: 12/19/2012
'DP'S: 6
'WS'S: 5
'80.  THIS DATE: 12/20/2012
'DP'S: 0
'WS'S: 13
'81.  THIS DATE: 12/21/2012
'DP'S: 0
'WS'S: 7
'82.  THIS DATE: 12/24/2012
'DP'S: 0
'WS'S: 7
'83.  THIS DATE: 12/26/2012
'DP'S: 0
'WS'S: 11
'84.  THIS DATE: 12/27/2012
'DP'S: 0
'WS'S: 1
'85.  THIS DATE: 12/28/2012
'DP'S: 15
'WS'S: 14
'86.  THIS DATE: 12/31/2012
'DP'S: 1
'WS'S: 11
'87.  THIS DATE: 01/03/2013
'DP'S: 1
'WS'S: 12
'88.  THIS DATE: 01/04/2013
'DP'S: 7
'WS'S: 6
'89.  THIS DATE: 01/07/2013
'DP'S: 1
'WS'S: 1
'90.  THIS DATE: 01/08/2013
'DP'S: 0
'WS'S: 8
'91.  THIS DATE: 01/09/2013
'DP'S: 2
'WS'S: 17
'92.  THIS DATE: 01/10/2013
'DP'S: 6
'WS'S: 12
'93.  THIS DATE: 01/11/2013
'DP'S: 0
'WS'S: 8
'94.  THIS DATE: 01/14/2013
'DP'S: 0
'WS'S: 14
'95.  THIS DATE: 01/15/2013
'DP'S: 0
'WS'S: 12
'96.  THIS DATE: 01/16/2013
'DP'S: 18
'WS'S: 18
'97.  THIS DATE: 01/17/2013
'DP'S: 0
'WS'S: 2
'98.  THIS DATE: 01/18/2013
'DP'S: 1
'WS'S: 29
'99.  THIS DATE: 01/22/2013
'DP'S: 0
'WS'S: 9
'100. THIS DATE: 01/23/2013
'DP'S: 1
'WS'S: 3
'101. THIS DATE: 01/24/2013
'DP'S: 0
'WS'S: 17
'102. THIS DATE: 01/25/2013
'DP'S: 16
'WS'S: 13
'103. THIS DATE: 01/29/2013
'DP'S: 0
'WS'S: 12
'104. THIS DATE: 01/30/2013
'DP'S: 5
'WS'S: 6
'105. THIS DATE: 01/31/2013
'DP'S: 8
'WS'S: 164
'106. THIS DATE: 02/01/2013
'DP'S: 1
'WS'S: 3
'107. THIS DATE: 02/04/2013
'DP'S: 0
'WS'S: 4
'108. THIS DATE: 02/05/2013
'DP'S: 0
'WS'S: 3
'109. THIS DATE: 02/06/2013
'DP'S: 0
'WS'S: 26
'110. THIS DATE: 02/07/2013
'DP'S: 18
'WS'S: 7
'111. THIS DATE: 02/08/2013
'DP'S: 0
'WS'S: 11
'112. THIS DATE: 02/11/2013
'DP'S: 0
'WS'S: 3
'113. THIS DATE: 02/12/2013
'DP'S: 1
'WS'S: 7
'114. THIS DATE: 02/13/2013
'DP'S: 0
'WS'S: 4
'115. THIS DATE: 02/14/2013
'DP'S: 0
'WS'S: 10
'116. THIS DATE: 02/15/2013
'DP'S: 15
'WS'S: 21
'117. THIS DATE: 02/19/2013
'DP'S: 0
'WS'S: 8
'118. THIS DATE: 02/20/2013
'DP'S: 0
'WS'S: 14
'119. THIS DATE: 02/21/2013
'DP'S: 11
'WS'S: 24
'120. THIS DATE: 02/22/2013
'DP'S: 0
'WS'S: 10
'121. THIS DATE: 02/25/2013
'DP'S: 5
'WS'S: 23
'122. THIS DATE: 02/26/2013
'DP'S: 2
'WS'S: 0
'123. THIS DATE: 02/27/2013
'DP'S: 1
'WS'S: 15
'124. THIS DATE: 02/28/2013
'DP'S: 24
'WS'S: 8
'125. THIS DATE: 03/01/2013
'DP'S: 5
'WS'S: 3
'126. THIS DATE: 03/04/2013
'DP'S: 0
'WS'S: 5
'127. THIS DATE: 03/05/2013
'DP'S: 1
'WS'S: 7
'128. THIS DATE: 03/06/2013
'DP'S: 0
'WS'S: 34
'129. THIS DATE: 03/07/2013
'DP'S: 10
'WS'S: 2
'130. THIS DATE: 03/11/2013
'DP'S: 5
'WS'S: 12
'131. THIS DATE: 03/12/2013
'DP'S: 1
'WS'S: 6
'132. THIS DATE: 03/13/2013
'DP'S: 0
'WS'S: 17
'133. THIS DATE: 03/14/2013
'DP'S: 0
'WS'S: 19
'134. THIS DATE: 03/15/2013
'DP'S: 17
'WS'S: 14
'135. THIS DATE: 03/18/2013
'DP'S: 0
'WS'S: 4
'136. THIS DATE: 03/19/2013
'DP'S: 2
'WS'S: 1
'137. THIS DATE: 03/20/2013
'DP'S: 0
'WS'S: 20
'138. THIS DATE: 03/21/2013
'DP'S: 2
'WS'S: 14
'139. THIS DATE: 03/22/2013
'DP'S: 10
'WS'S: 15
'140. THIS DATE: 03/25/2013
'DP'S: 0
'WS'S: 14
'141. THIS DATE: 03/26/2013
'DP'S: 0
'WS'S: 3
'142. THIS DATE: 03/27/2013
'DP'S: 0
'WS'S: 13
'143. THIS DATE: 03/28/2013
'DP'S: 25
'WS'S: 2
'144. THIS DATE: 03/29/2013
'DP'S: 0
'WS'S: 4
'145. THIS DATE: 03/30/2013
'DP'S: 0
'WS'S: 1
'146. THIS DATE: 04/01/2013
'DP'S: 0
'WS'S: 10
'147. THIS DATE: 04/02/2013
'DP'S: 1
'WS'S: 35
'148. THIS DATE: 04/03/2013
'DP'S: 0
'WS'S: 7
'149. THIS DATE: 04/04/2013
'DP'S: 16
'WS'S: 7
'150. THIS DATE: 04/05/2013
'DP'S: 1
'WS'S: 3
'151. THIS DATE: 04/08/2013
'DP'S: 2
'WS'S: 10
'152. THIS DATE: 04/09/2013
'DP'S: 10
'WS'S: 7
'153. THIS DATE: 04/10/2013
'DP'S: 0
'WS'S: 4
'154. THIS DATE: 04/11/2013
'DP'S: 0
'WS'S: 10
'155. THIS DATE: 04/12/2013
'DP'S: 0
'WS'S: 13
'156. THIS DATE: 04/15/2013
'DP'S: 0
'WS'S: 6
'157. THIS DATE: 04/16/2013
'DP'S: 0
'WS'S: 12
'158. THIS DATE: 04/17/2013
'DP'S: 15
'WS'S: 0
'159. THIS DATE: 04/18/2013
'DP'S: 0
'WS'S: 17
'160. THIS DATE: 04/19/2013
'DP'S: 0
'WS'S: 2
'161. THIS DATE: 04/22/2013
'DP'S: 2
'WS'S: 12
'162. THIS DATE: 04/23/2013
'DP'S: 0
'WS'S: 7
'163. THIS DATE: 04/24/2013
'DP'S: 2
'WS'S: 6
'164. THIS DATE: 04/25/2013
'DP'S: 25
'WS'S: 11
'165. THIS DATE: 04/26/2013
'DP'S: 2
'WS'S: 6
'166. THIS DATE: 04/29/2013
'DP'S: 24
'WS'S: 26
'167. THIS DATE: 04/30/2013
'DP'S: 4
'WS'S: 14
'168. THIS DATE: 05/02/2013
'DP'S: 0
'WS'S: 33
'169. THIS DATE: 05/03/2013
'DP'S: 15
'WS'S: 7
'170. THIS DATE: 05/06/2013
'DP'S: 0
'WS'S: 12
'171. THIS DATE: 05/07/2013
'DP'S: 0
'WS'S: 4
'172. THIS DATE: 05/08/2013
'DP'S: 0
'WS'S: 3
'173. THIS DATE: 05/09/2013
'DP'S: 10
'WS'S: 6
'174. THIS DATE: 05/10/2013
'DP'S: 1
'WS'S: 8
'175. THIS DATE: 05/13/2013
'DP'S: 1
'WS'S: 5
'176. THIS DATE: 05/14/2013
'DP'S: 0
'WS'S: 8
'177. THIS DATE: 05/15/2013
'DP'S: 0
'WS'S: 17
'178. THIS DATE: 05/17/2013
'DP'S: 1
'WS'S: 33
'179. THIS DATE: 05/20/2013
'DP'S: 7
'WS'S: 25
'180. THIS DATE: 05/21/2013
'DP'S: 0
'WS'S: 1
'181. THIS DATE: 05/22/2013
'DP'S: 0
'WS'S: 13
'182. THIS DATE: 05/23/2013
'DP'S: 17
'WS'S: 45
'183. THIS DATE: 05/24/2013
'DP'S: 0
'WS'S: 1
'184. THIS DATE: 05/28/2013
'DP'S: 2
'WS'S: 33
'185. THIS DATE: 05/30/2013
'DP'S: 2
'WS'S: 15
'186. THIS DATE: 05/31/2013
'DP'S: 17
'WS'S: 4
'187. THIS DATE: 06/03/2013
'DP'S: 1
'WS'S: 7
'188. THIS DATE: 06/04/2013
'DP'S: 0
'WS'S: 9
'189. THIS DATE: 06/06/2013
'DP'S: 13
'WS'S: 3
'190. THIS DATE: 06/07/2013
'DP'S: 0
'WS'S: 9
'191. THIS DATE: 06/12/2013
'DP'S: 1
'WS'S: 7
'192. THIS DATE: 06/13/2013
'DP'S: 5
'WS'S: 11
'193. THIS DATE: 06/14/2013
'DP'S: 9
'WS'S: 14
'194. THIS DATE: 06/17/2013
'DP'S: 3
'WS'S: 0
'195. THIS DATE: 06/18/2013
'DP'S: 5
'WS'S: 24
'196. THIS DATE: 06/19/2013
'DP'S: 1
'WS'S: 4
'197. THIS DATE: 06/20/2013
'DP'S: 8
'WS'S: 13
'198. THIS DATE: 06/21/2013
'DP'S: 0
'WS'S: 14
'199. THIS DATE: 06/26/2013
'DP'S: 2
'WS'S: 22
'200. THIS DATE: 06/27/2013
'DP'S: 9
'WS'S: 11
'201. THIS DATE: 06/28/2013
'DP'S: 0
'WS'S: 11
'202. THIS DATE: 07/01/2013
'DP'S: 0
'WS'S: 2
'203. THIS DATE: 07/02/2013
'DP'S: 13
'WS'S: 11
'204. THIS DATE: 07/03/2013
'DP'S: 2
'WS'S: 6
'205. THIS DATE: 07/05/2013
'DP'S: 21
'WS'S: 3
'206. THIS DATE: 07/09/2013
'DP'S: 1
'WS'S: 8
'207. THIS DATE: 07/11/2013
'DP'S: 0
'WS'S: 20
'208. THIS DATE: 07/12/2013
'DP'S: 4
'WS'S: 6
'209. THIS DATE: 07/15/2013
'DP'S: 0
'WS'S: 413
'210. THIS DATE: 07/16/2013
'DP'S: 1
'WS'S: 4
'211. THIS DATE: 07/17/2013
'DP'S: 0
'WS'S: 23
'212. THIS DATE: 07/18/2013
'DP'S: 0
'WS'S: 11
'213. THIS DATE: 07/19/2013
'DP'S: 0
'WS'S: 12
'214. THIS DATE: 07/22/2013
'DP'S: 26
'WS'S: 20
'215. THIS DATE: 07/23/2013
'DP'S: 3
'WS'S: 2
'216. THIS DATE: 07/24/2013
'DP'S: 0
'WS'S: 18
'217. THIS DATE: 07/25/2013
'DP'S: 1
'WS'S: 41
'218. THIS DATE: 07/26/2013
'DP'S: 3
'WS'S: 25
'219. THIS DATE: 07/29/2013
'DP'S: 0
'WS'S: 1
'220. THIS DATE: 07/30/2013
'DP'S: 3
'WS'S: 24
'221. THIS DATE: 07/31/2013
'DP'S: 18
'WS'S: 4
'222. THIS DATE: 08/01/2013
'DP'S: 0
'WS'S: 2
'223. THIS DATE: 08/02/2013
'DP'S: 4
'WS'S: 12
'224. THIS DATE: 08/05/2013
'DP'S: 2
'WS'S: 3
'225. THIS DATE: 08/06/2013
'DP'S: 0
'WS'S: 2
'226. THIS DATE: 08/07/2013
'DP'S: 0
'WS'S: 9
'227. THIS DATE: 08/08/2013
'DP'S: 16
'WS'S: 8
'228. THIS DATE: 08/12/2013
'DP'S: 0
'WS'S: 26
'229. THIS DATE: 08/13/2013
'DP'S: 0
'WS'S: 9
'230. THIS DATE: 08/14/2013
'DP'S: 27
'WS'S: 14
'231. THIS DATE: 08/15/2013
'DP'S: 0
'WS'S: 12
'232. THIS DATE: 08/16/2013
'DP'S: 5
'WS'S: 8
'233. THIS DATE: 08/19/2013
'DP'S: 0
'WS'S: 10
'234. THIS DATE: 08/20/2013
'DP'S: 2
'WS'S: 4
'235. THIS DATE: 08/21/2013
'DP'S: 1
'WS'S: 9
'236. THIS DATE: 08/22/2013
'DP'S: 11
'WS'S: 9
'237. THIS DATE: 08/23/2013
'DP'S: 0
'WS'S: 6
'238. THIS DATE: 08/26/2013
'DP'S: 0
'WS'S: 9
'239. THIS DATE: 08/27/2013
'DP'S: 1
'WS'S: 8
'240. THIS DATE: 08/28/2013
'DP'S: 0
'WS'S: 4
'241. THIS DATE: 08/29/2013
'DP'S: 41
'WS'S: 37
'242. THIS DATE: 08/30/2013
'DP'S: 1
'WS'S: 1
'243. THIS DATE: 09/03/2013
'DP'S: 3
'WS'S: 9
'244. THIS DATE: 09/04/2013
'DP'S: 2
'WS'S: 44
'245. THIS DATE: 09/05/2013
'DP'S: 24
'WS'S: 9
'246. THIS DATE: 09/06/2013
'DP'S: 1
'WS'S: 11
'247. THIS DATE: 09/09/2013
'DP'S: 0
'WS'S: 8
'248. THIS DATE: 09/10/2013
'DP'S: 0
'WS'S: 6
'249. THIS DATE: 09/11/2013
'DP'S: 0
'WS'S: 14
'250. THIS DATE: 09/12/2013
'DP'S: 0
'WS'S: 7
'251. THIS DATE: 09/13/2013
'DP'S: 14
'WS'S: 9
'252. THIS DATE: 09/16/2013
'DP'S: 0
'WS'S: 3
'253. THIS DATE: 09/17/2013
'DP'S: 0
'WS'S: 4
'254. THIS DATE: 09/19/2013
'DP'S: 7
'WS'S: 7
'255. THIS DATE: 09/20/2013
'DP'S: 0
'WS'S: 2
'256. THIS DATE: 09/22/2013
'DP'S: 0
'WS'S: 1
'257. THIS DATE: 09/23/2013
'DP'S: 0
'WS'S: 33
'258. THIS DATE: 09/24/2013
'DP'S: 0
'WS'S: 33
'259. THIS DATE: 09/25/2013
'DP'S: 4
'WS'S: 5
'260. THIS DATE: 09/26/2013
'DP'S: 3
'WS'S: 21
'261. THIS DATE: 09/27/2013
'DP'S: 2
'WS'S: 15
'262. THIS DATE: 09/30/2013
'DP'S: 28
'WS'S: 7
'263. THIS DATE: 10/01/2013
'DP'S: 2
'WS'S: 7

'264. THIS DATE: 10/02/2013
'DP'S: 0
'WS'S: 31
'265. THIS DATE: 10/03/2013
'DP'S: 0
'WS'S: 30
'266. THIS DATE: 10/04/2013
'DP'S: 21
'WS'S: 6
'267. THIS DATE: 10/07/2013
'DP'S: 297
'WS'S: 13
'268. THIS DATE: 10/08/2013
'DP'S: 1
'WS'S: 10
'269. THIS DATE: 10/09/2013
'DP'S: 0
'WS'S: 18
'270. THIS DATE: 10/10/2013
'DP'S: 13
'WS'S: 18
'271. THIS DATE: 10/11/2013
'DP'S: 0
'WS'S: 6
'272. THIS DATE: 10/15/2013
'DP'S: 0
'WS'S: 14
'273. THIS DATE: 10/16/2013
'DP'S: 0
'WS'S: 24
'274. THIS DATE: 10/17/2013
'DP'S: 0
'WS'S: 7
'275. THIS DATE: 10/18/2013
'DP'S: 10
'WS'S: 11
'276. THIS DATE: 10/21/2013
'DP'S: 0
'WS'S: 13
'277. THIS DATE: 10/22/2013
'DP'S: 0
'WS'S: 9
'278. THIS DATE: 10/23/2013
'DP'S: 0
'WS'S: 14
'279. THIS DATE: 10/24/2013
'DP'S: 16
'WS'S: 5

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  RZ_Avg1 = blnRetVal

End Function

Public Function RZ_EmptyTmps() As Boolean

  Const THIS_PROC As String = "RZ_EmptyTmps"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef
  Dim blnRetVal As Boolean

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs
    ' ** Empty zz_tbl_Revzon_01.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_01")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_02.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_02")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_02_tmp.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_03")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_03.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_04")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_03_tmp.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_05")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_04.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_06")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_05.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_07")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_06.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_08")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_07.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_09")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_08.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_10")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_09.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_11")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_10.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_12")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_11.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_13")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_12.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_14")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_13.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_15")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_14.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_16")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_15.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_17")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_16.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_18")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_17.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_19")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_18.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_20")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Revzon_19.
    Set qdf = .QueryDefs("zzz_qry_Revzon_29_21")
    qdf.Execute
    Set qdf = Nothing
    .Close
  End With

  Beep

  Set qdf = Nothing
  Set dbs = Nothing

  RZ_EmptyTmps = blnRetVal

End Function

Public Function RZ_DelObjs() As Boolean

  Const THIS_PROC As String = "RZ_DelObjs"

  Dim dbs As DAO.Database, tdf As DAO.TableDef, qdf As DAO.QueryDef
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim blnDelete As Boolean
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const T_TNAM As Integer = 0

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const Q_QNAM As Integer = 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngTbls = 0&
  ReDim arr_varTbl(T_ELEMS, 0)

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    For Each tdf In .TableDefs
      With tdf
        If Left(.Name, 12) = "tblAccuTech_" Then
          lngTbls = lngTbls + 1&
          lngE = lngTbls - 1&
          ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          arr_varTbl(T_TNAM, lngE) = .Name
        ElseIf Left(.Name, 10) = "tblRevzon_" Then
          lngTbls = lngTbls + 1&
          lngE = lngTbls - 1&
          ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          arr_varTbl(T_TNAM, lngE) = .Name
        ElseIf Left(.Name, 18) = "zz_tbl_Germantown_" Then
          lngTbls = lngTbls + 1&
          lngE = lngTbls - 1&
          ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          arr_varTbl(T_TNAM, lngE) = .Name
        ElseIf Left(.Name, 14) = "zz_tbl_Revzon_" Then
          lngTbls = lngTbls + 1&
          lngE = lngTbls - 1&
          ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          arr_varTbl(T_TNAM, lngE) = .Name
        End If
      End With  ' ** tdf.
    Next  ' ** tdf.

    For Each qdf In .QueryDefs
      With qdf
        If Left(.Name, 16) = "zz_qry_AccuTech_" Then
          lngQrys = lngQrys + 1&
          lngE = lngQrys - 1&
          ReDim Preserve arr_varQry(Q_ELEMS, lngE)
          arr_varQry(Q_QNAM, lngE) = .Name
        ElseIf Left(.Name, 15) = "zzz_qry_Revzon_" Then
          lngQrys = lngQrys + 1&
          lngE = lngQrys - 1&
          ReDim Preserve arr_varQry(Q_ELEMS, lngE)
          arr_varQry(Q_QNAM, lngE) = .Name
        End If
      End With  ' ** qdf.
    Next  ' ** qdf.

    .Close
  End With

  If lngTbls > 0& Or lngQrys > 0& Then

    Debug.Print "'DEL TBLS: " & CStr(lngTbls)
    DoEvents

    Debug.Print "'DEL QRYS: " & CStr(lngQrys)
    DoEvents

    If lngTbls > 0& Then

      blnDelete = True
      Debug.Print "'DELETE THESE TABLES?"
      Stop

      If blnDelete = True Then
        For lngX = 0& To (lngTbls - 1&)
          DoCmd.DeleteObject acTable, arr_varTbl(T_TNAM, lngX)
          DoEvents
        Next  ' ** lngX.
      End If  ' ** blnDelete.

    End If  ' ** lngTbls.

    If lngQrys > 0& Then

      blnDelete = True
      Debug.Print "'DELETE THESE QUERIES?"
      Stop

      If blnDelete = True Then
        For lngX = 0& To (lngQrys - 1&)
          DoCmd.DeleteObject acQuery, arr_varQry(Q_QNAM, lngX)
          DoEvents
        Next  ' ** lngX.
      End If  ' ** blnDelete.

    End If  ' ** lngQrys.

  Else
    Debug.Print "'NONE FOUND!"
  End If  ' ** lngTbls, lngQrys.

  Debug.Print "'DONE!"

'DEL TBLS: 56
'DEL QRYS: 882
'DONE!
  Beep

  Set qdf = Nothing
  Set tdf = Nothing
  Set dbs = Nothing

  RZ_DelObjs = blnRetVal

End Function
