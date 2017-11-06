VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssetPricing_Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssetPricing_Import"

'VGC 09/05/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Path:             A {priceimport_path}
' **   Browse:           B {cmdBrowse}
' **   Close/Cancel:     C {cmdCancel}
' **   File:             F {priceimport_file}
' **   Import:           N {cmdImport}
' **   Print:            P {cmdPrint}
' **   Preview:          W {cmdPreview}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}
' **   Save:             S {cmdSave}
' **

' ***********************************************
' ** Easter Egg (Secret Key):
' ** priceimport_id_lbx_MouseUp
' ***********************************************

'BASIC QUESTIONS:
'1. As it stands now, they will have already chosen a directory to save the exports in, which could be anywhere.
'2. Should we copy/move the files to a standard directory?
'3. Can they have multiple files to choose from?
'4. If their exports are always of the same name, should we rename them with a date once imported?
'5.

'WE NEED TO KNOW UNDER WHAT CONDITIONS THE Yield SHOULDN'T BE UPDATED!

' *********************************************************************************
' ** Note: NewDate on form should be the same as Appraisal Date within EstateVal.
' *********************************************************************************

' ** Rich's examples:
' **   Set 1:
' **     Stocks:
' **       14  GENERAL ELECTRIC CO (369604103)
' **       47  LUCENT TECHNOLOGIES INC (549463107)
' **       1   AT&T CORP (001957109)
' **       53  IOWA FIN AUTH HOSP FAC REV     REVS (46246NHH2)
' **     Files:
' **       tapricing.rtf
' **   Set 2:
' **     Stocks:
' **       14  GENERAL ELECTRIC CO (369604103)
' **       57  SHORT-TERM INVTS TR (825252687)
' **             AIM Liquidated Assets Reserve Class (AIM Liquid Assets Res)
' **             57  825252687 Short-Term Invts Tr (AIM Liquid Assets Res) 70  Miscellaneous
' **     Files:
' **       northbranch (Appraisal Date Evaluation).rtf
' **       northbranch (Appraisal Date Evaluation).txt
' **       tapricing (Appraisal Date Evaluation).rtf
' **       tapricing (Appraisal Date Evaluation).txt
' **   Set 3:
' **     Stocks:
' **       14  GENERAL ELECTRIC CO (369604103)
' **     Files:
' **       rich (Alternate Date Evaluation).rtf
' **       rich.txt
' **       tapricing.txt

' ** List box column constants: priceimport_id_lbx.
'Private Const LBX_PI_ID   As Integer = 0  ' ** priceimport_id.
Private Const LBX_PI_NAME As Integer = 1  ' ** priceimport_name.
'Private Const LBX_PI_PUB  As Integer = 2  ' ** priceimport_publisher.
'Private Const LBX_PI_DESC As Integer = 3  ' ** priceimport_description.
Private Const LBX_PI_EXT  As Integer = 4  ' ** priceimport_extension.

' ** Array: arr_varFilt() - Windows Filter String.
Private lngFilts As Long, arr_varFilt As Variant
'Private Const FLT_ID  As Integer = 0
Private Const FLT_NAM As Integer = 1
Private Const FLT_EXT As Integer = 2
'Private Const FLT_ORD As Integer = 3
Private Const FLT_PID As Integer = 4

' ** Array: arr_varFileType() - all recognized file extensions.
Private lngFileTypes As Long, arr_varFileType() As Variant
Private Const FT_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const FT_ID  As Integer = 0  ' ** priceimport_id.
Private Const FT_EXT As Integer = 1  ' ** Extension.
Private Const FT_PF  As Integer = 2  ' ** Path\File.

Private strIMPORT_PricingPath As String
Private strIMPORT_FileType As String  ' ** 4 characters: includes period.
Private blnIMPORT_EmptyPricingPath As Boolean

Private blnGoingToReport As Boolean
Private Const dblProgBar_Max As Double = 100#
Private dblProgBar_Incr As Double
Private dblProgBox_Width As Double
Private dblProgBar_Len As Double
Private blnProgBarOff As Boolean
Private lngTpp As Long, lngPriceImportPath_Height As Long, lngBrowse_Top As Long
Private intPct As Integer
Private blnUpdateSuccess As Boolean, blnAfterUpdate As Boolean
Private blnClearPathFile As Boolean, lngClearPathFileID As Long
Private blnBrowse_Focus As Boolean, blnBrowse_MouseDown As Boolean
Private strCallingForm As String, strCallersCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String, blnTmp02 As Boolean
        Dim lngX As Long, lngE As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strCallingForm = Trim(.OpenArgs)
150           intPos01 = InStr(strCallingForm, "~")
160           If intPos01 > 0 Then

170             strCallersCallingForm = Mid(strCallingForm, (intPos01 + 1))
180             strCallingForm = Left(strCallingForm, (intPos01 - 1))

190             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

210             gstrReportQuerySpec = vbNullString
220             blnUpdateSuccess = False: blnProgBarOff = False
230             strIMPORT_PricingPath = vbNullString: strIMPORT_FileType = vbNullString: blnIMPORT_EmptyPricingPath = False
240             blnAfterUpdate = False: blnClearPathFile = False: lngClearPathFileID = 0&
250             blnBrowse_Focus = False: blnBrowse_MouseDown = False
260             lngPriceImportPath_Height = .priceimport_path.Height
270             lngBrowse_Top = .cmdBrowse.Top

                ' ** Just to be sure.
280             If gstrTrustDataLocation = vbNullString Then
290               blnTmp02 = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
300             End If

310             .priceimport_lbl3.Caption = gstrTrustDataLocation & gstrDir_Pricing

320             Set dbs = CurrentDb
330             With dbs

                  ' ** Empty tblPricing_Cusip.
340               Set qdf = .QueryDefs("qryPricing_06")
350               qdf.Execute

                  ' ** Get the list of Windows filter strings from tblPricing_FileType.
360               Set qdf = .QueryDefs("qryPricing_FileType_01")
370               Set rst = qdf.OpenRecordset
380               With rst
390                 .MoveLast
400                 lngFilts = .RecordCount
410                 .MoveFirst
420                 arr_varFilt = .GetRows(lngFilts)
                    ' *******************************************************
                    ' ** Array: arr_varFilt()
                    ' **
                    ' **   Field  Element  Name                  Constant
                    ' **   =====  =======  ====================  ==========
                    ' **     1       0     filetype_id           FLT_ID
                    ' **     2       1     filetype_name         FLT_NAM
                    ' **     3       2     filetype_extension    FLT_EXT
                    ' **     4       3     filetype_order        FLT_ORD
                    ' **     5       4     priceimport_id        FLT_PID
                    ' **
                    ' *******************************************************
430                 .Close
440               End With

450               lngFileTypes = 0&
460               ReDim arr_varFileType(FT_ELEMS, 0)
                  ' ***************************************************
                  ' ** Array: arr_varFileType()
                  ' **
                  ' **   Field  Element  Name              Constant
                  ' **   =====  =======  ================  ==========
                  ' **     1       0     priceimport_id    FT_ID
                  ' **     2       1     Extension w/.     FT_EXT
                  ' **     3       2     Path\File         FT_PF
                  ' **
                  ' ***************************************************

                  ' ** Get a list of all the recognized file extensions from tblPricing_Import.
470               Set qdf = .QueryDefs("qryPricing_03")
480               Set rst = qdf.OpenRecordset
490               With rst
500                 .MoveLast
510                 lngRecs = .RecordCount
520                 .MoveFirst
530                 For lngX = 1& To lngRecs
540                   strTmp01 = ![priceimport_extension]
550                   intPos01 = InStr(strTmp01, ";")
560                   If intPos01 = 0 Then
570                     lngFileTypes = lngFileTypes + 1&
580                     lngE = lngFileTypes - 1&
590                     ReDim Preserve arr_varFileType(FT_ELEMS, lngE)
600                     arr_varFileType(FT_ID, lngE) = ![priceimport_id]
610                     arr_varFileType(FT_EXT, lngE) = "." & strTmp01  ' ** Include the period to be consistent with strIMPORT_FileType.
620                     arr_varFileType(FT_PF, lngE) = vbNullString
630                   Else
640                     Do While intPos01 > 0
650                       lngFileTypes = lngFileTypes + 1&
660                       lngE = lngFileTypes - 1&
670                       ReDim Preserve arr_varFileType(FT_ELEMS, lngE)
680                       arr_varFileType(FT_ID, lngE) = ![priceimport_id]
690                       arr_varFileType(FT_EXT, lngE) = "." & Left(strTmp01, (intPos01 - 1))
700                       arr_varFileType(FT_PF, lngE) = vbNullString
710                       strTmp01 = Mid(strTmp01, (intPos01 + 1))
720                       intPos01 = InStr(strTmp01, ";")
730                       If intPos01 = 0 Then
                            ' ** Last one.
740                         lngFileTypes = lngFileTypes + 1&
750                         lngE = lngFileTypes - 1&
760                         ReDim Preserve arr_varFileType(FT_ELEMS, lngE)
770                         arr_varFileType(FT_ID, lngE) = ![priceimport_id]
780                         arr_varFileType(FT_EXT, lngE) = "." & strTmp01
790                         arr_varFileType(FT_PF, lngE) = vbNullString
800                         Exit Do
810                       End If
820                     Loop
830                   End If
840                   If lngX < lngRecs Then .MoveNext
850                 Next
860                 .Close
870               End With

880               .Close
890             End With

                ' ** Create the Pricing folder if it doesn't already exist.
900             If DirExists(gstrTrustDataLocation & gstrDir_Pricing) = False Then  ' ** gstrTrustDataLocation includes trailing backslash.
910               MkDir (gstrTrustDataLocation & gstrDir_Pricing)
920               strIMPORT_PricingPath = (gstrTrustDataLocation & gstrDir_Pricing)
930               blnIMPORT_EmptyPricingPath = True
940             Else
950               strIMPORT_PricingPath = (gstrTrustDataLocation & gstrDir_Pricing)
960               If Dir(strIMPORT_PricingPath & LNK_SEP & "*.*") = vbNullString Then
970                 blnIMPORT_EmptyPricingPath = True
980               End If
990             End If

                ' ** Disable till an import's been chosen.
1000            .priceimport_path.Enabled = False
1010            .priceimport_path.BorderColor = WIN_CLR_DISR
1020            .priceimport_path.BackStyle = acBackStyleTransparent
1030            .priceimport_path.Visible = False
1040            .priceimport_file.Enabled = False
1050            .priceimport_file.BorderColor = WIN_CLR_DISR
1060            .priceimport_file.BackStyle = acBackStyleTransparent
1070            .priceimport_file.Visible = False
                '.priceimport_lbl.BackStyle = acBackStyleTransparent
1080            .priceimport_lbl.ForeColor = WIN_CLR_DISF
1090            .priceimport_lbl_dim_hi.Visible = True
1100            .cmdBrowse.Enabled = False
1110            .cmdImport.Enabled = False
1120            .priceimport_file_lbl2.Visible = False
1130            .priceimport_file_lbl3.Visible = False

                ' ** Disable till the import's given us something to view.
1140            .cmdPreview.Enabled = False
1150            .cmdPrint.Enabled = False
1160            .cmdPreviewPrint_lbl.ForeColor = WIN_CLR_DISF
1170            .cmdPreviewPrint_lbl_dim_hi.Visible = True

                ' ** Width of full progress bar.
1180            dblProgBox_Width = .ProgBar_box.Width
1190            dblProgBar_Incr = (dblProgBox_Width / dblProgBar_Max)  ' ** 1/100th of the width.
1200            dblProgBar_Len = 0#: intPct = 0
1210            ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Procedure: Below.
                '.ProgBar_bar.Width = dblProgBar_Len
1220            .ProgBar_lbl.Caption = CStr(intPct) & "%"
1230            .ProgBar_lbl.ForeColor = WIN_CLR_DISF
1240            .ProgBar_lbl_dim_hi.Caption = CStr(intPct) & "%"
1250            ProgBar_Width_Pric False, 1  ' ** Procedure: Below.
                '.ProgBar_bar.Visible = False
1260            .ProgBar_box.Visible = False
1270            .ProgBar_box2.Visible = False
1280            .ProgBar_lbl.Visible = False
1290            .ProgBar_lbl_dim_hi.Visible = False
1300            DoEvents

1310            If IsNull(.priceimport_file) = False Then
1320              .priceimport_file_lbl3.Visible = True
1330            End If
1340            priceimport_id_lbx_AfterUpdate  ' ** Procedure: Below.
1350            priceimport_file_AfterUpdate  ' ** Procedure: Below.
1360            chkPageOf_AfterUpdate  ' ** Procedure: Below.

1370            If gblnGoToReport = False Then
1380              DoCmd.Hourglass False
1390            End If

1400            .priceimport_id_lbx.SetFocus

1410          Else
1420            Cancel = -1
1430          End If
1440        Else
1450          Cancel = -1
1460        End If
1470      Else
1480        Cancel = -1
1490      End If
1500    End With

1510    If Cancel = -1 Then
1520      Beep
1530      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1540    End If

EXITP:
1550    Set rst = Nothing
1560    Set qdf = Nothing
1570    Set dbs = Nothing
1580    Exit Sub

ERRH:
1590    DoCmd.Hourglass False
1600    Select Case ERR.Number
        Case Else
1610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1620    End Select
1630    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1720    intShiftDown = (Shift And acShiftMask) > 0
1730    intAltDown = (Shift And acAltMask) > 0
1740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1760      Select Case intRetVal
          Case vbKeyEscape
1770        intRetVal = 0
1780        cmdCancel_Click  ' ** Procedure: Below.
1790      End Select
1800    End If

        ' ** Ctrl keys.
1810    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1820      Select Case intRetVal
          Case vbKeyG
1830        With Me
1840          intRetVal = 0
1850          If .chkPageOf.Enabled = True Then
1860            .chkPageOf.SetFocus
1870            .chkPageOf = (Not .chkPageOf)
1880            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1890          End If
1900        End With
1910      Case vbKeyS
1920        intRetVal = 0
1930        cmdSave_Click  ' ** Procedure: Below.
1940      End Select
1950    End If

        ' ** Ctrl-Shift keys.
1960    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1970      Select Case intRetVal
          Case vbKeyF
1980        With Me
1990          intRetVal = 0
2000          .FocusHolder.SetFocus
2010        End With
2020      End Select
2030    End If

EXITP:
2040    KeyCode = intRetVal
2050    Exit Sub

ERRH:
2060    intRetVal = 0
2070    Select Case ERR.Number
        Case Else
2080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2090    End Select
2100    Resume EXITP

End Sub

Private Sub Form_Timer()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strFormName As String

2210    With Me
2220      .TimerInterval = 0
2230      If gblnSetFocus = True Then
2240        gblnSetFocus = False
2250        strFormName = vbNullString
2260  On Error Resume Next
2270        Do Until strFormName = THIS_NAME
2280          DoCmd.SelectObject acForm, THIS_NAME, False
2290          strFormName = Screen.ActiveForm.Name
2300        Loop
2310  On Error GoTo ERRH
2320        .priceimport_id_lbx.SetFocus
2330      ElseIf blnProgBarOff = True Then
2340        blnProgBarOff = False
2350        ProgBar_Width_Pric False, 1  ' ** Procedure: Below.
            '.ProgBar_bar.Visible = False
2360        .ProgBar_box.Visible = False
2370        .ProgBar_box2.Visible = False
2380        .ProgBar_lbl.Visible = False
2390        .ProgBar_lbl_dim_hi.Visible = False
2400      ElseIf blnClearPathFile = True Then
2410        blnClearPathFile = False
2420        cmdSave_Click  ' ** Procedure: Below.
2430        Set dbs = CurrentDb
2440        With dbs
              ' ** Null-out priceimport_path, priceimport_file, by specified [pricimpid].
2450          Set qdf = .QueryDefs("qryPricing_08")
2460          With qdf.Parameters
2470            ![pricimpid] = lngClearPathFileID
2480          End With
2490          qdf.Execute dbFailOnError
2500          lngClearPathFileID = 0&
2510          .Close
2520        End With
2530        DoEvents
2540        .Requery
2550        .priceimport_id_lbx.Requery
2560        .priceimport_file_lbl2.Visible = False
2570        .priceimport_file_lbl2.Caption = vbNullString
2580        .priceimport_file_lbl3.Visible = False
2590        .cmdImport.Enabled = False
2600      End If
2610      If gblnGoToReport = True Then
2620        DoCmd.Hourglass True  ' ** Make sure it's still running.
2630        DoEvents
2640        Select Case blnGoingToReport
            Case True
2650          blnGoingToReport = False
2660          gblnGoToReport = False
2670          .cmdCancel.SetFocus
2680          .cmdPreviewPrint_lbl.ForeColor = WIN_CLR_DISF
2690          .cmdPreviewPrint_lbl_dim_hi.Visible = True
2700          .cmdPreview.Enabled = False
2710          .cmdPrint.Enabled = False
2720          DoCmd.Hourglass False
2730        Case False
2740          blnGoingToReport = True
2750          .cmdPreviewPrint_lbl.ForeColor = CLR_DKGRY2
2760          .cmdPreviewPrint_lbl_dim_hi.Visible = False
2770          .cmdPreview.Enabled = True
2780          .cmdPrint.Enabled = True
2790          .cmdPrint.SetFocus
2800          Beep
2810          DoEvents
2820          Forms(strCallingForm).TimerInterval = 0&
2830          Forms(strCallingForm).GTREmblem_Off  ' ** Form Procedure: frmAssetPricing.
2840          Forms(strCallingForm).GoToReport_arw_price_img.Visible = False
2850          Forms(strCallingForm).GoToReport_arw_date_img.Visible = False
2860          DoEvents
2870          .TimerInterval = (GTR_WAIT * 2&)
2880        End Select
2890      End If
2900    End With

EXITP:
2910    Set qdf = Nothing
2920    Set dbs = Nothing
2930    Exit Sub

ERRH:
2940    DoCmd.Hourglass False
2950    Select Case ERR.Number
        Case Else
          'Case 2115  ' ** The macro or function set to the BeforeUpdate or ValidationRule property for this field is preventing Microsoft Access from saving the data in the field.
          'Case 2118  ' ** You must save the current field before you run the Requery action.
2960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970    End Select
2980    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, frm As Access.Form

3010    DoCmd.Hourglass True
3020    DoEvents

3030    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3040    gstrReportQuerySpec = vbNullString

        ' ** Don't leave any reports open.
3050    Do While Reports.Count > 0&
3060      DoCmd.Close acReport, Reports(0).Name
3070    Loop
3080    DoEvents

3090    Set dbs = CurrentDb
3100    With dbs
          ' ** Empty tblPricing_Staging.
3110      Set qdf = .QueryDefs("qryPricing_04e")
3120      qdf.Execute
3130      Set qdf = Nothing
3140      .Close
3150    End With
3160    Set dbs = Nothing
3170    DoEvents

3180    If strCallingForm <> vbNullString Then
3190      DoCmd.OpenForm strCallingForm, , , , , , strCallersCallingForm
3200      Set frm = Forms(strCallingForm)
3210  On Error Resume Next
3220      frm.cmdPrice.SetFocus
3230      gblnSetFocus = True
3240      frm.TimerInterval = 100&
3250  On Error GoTo ERRH
3260      If blnUpdateSuccess = True Then
3270        frm.frmAssetPricing_Sub.Form.yield_upd.Visible = True
3280        frm.frmAssetPricing_Sub.Form.marketvalue_upd.Visible = True
3290        If frm.chkIncludeCurrency = True Then
3300          frm.frmAssetPricing_Sub.Form.marketvalue_usd_upd.Visible = True
3310        End If
3320      End If
3330    End If

EXITP:
3340    Set frm = Nothing
3350    Set qdf = Nothing
3360    Set dbs = Nothing
3370    Exit Sub

ERRH:
3380    DoCmd.Hourglass False
3390    Select Case ERR.Number
        Case Else
3400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3410    End Select
3420    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3510    With Me
3520      If .cmdBrowse_raised_focus_dots_img.Visible = True Or .cmdBrowse_raised_focus_img.Visible = True Then
3530        Select Case blnBrowse_Focus
            Case True
3540          .cmdBrowse_raised_semifocus_dots_img.Visible = True
3550          .cmdBrowse_raised_img.Visible = False
3560        Case False
3570          .cmdBrowse_raised_img.Visible = True
3580          .cmdBrowse_raised_semifocus_dots_img.Visible = False
3590        End Select
3600        .cmdBrowse_raised_focus_img.Visible = False
3610        .cmdBrowse_raised_focus_dots_img.Visible = False
3620        .cmdBrowse_sunken_focus_dots_img.Visible = False
3630        .cmdBrowse_raised_img_dis.Visible = False
3640      End If
3650    End With

EXITP:
3660    Exit Sub

ERRH:
3670    Select Case ERR.Number
        Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3810    With Me

3820      DoCmd.Hourglass True
3830      DoEvents

3840      Set dbs = CurrentDb
3850      With dbs
            ' ** Update qryPricing_20 (tblPricing_Import, linked to tblPricing_Staging, with .._new fields).
3860        Set qdf = .QueryDefs("qryPricing_21")
3870        qdf.Execute
3880        Set qdf = Nothing
3890        .Close
3900      End With
3910      Set dbs = Nothing
3920      DoEvents

3930      DoCmd.Close acForm, THIS_NAME

3940    End With

EXITP:
3950    Set qdf = Nothing
3960    Set dbs = Nothing
3970    Exit Sub

ERRH:
3980    DoCmd.Hourglass False
3990    Select Case ERR.Number
        Case Else
4000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4010    End Select
4020    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4120    intShiftDown = (Shift And acShiftMask) > 0
4130    intAltDown = (Shift And acAltMask) > 0
4140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyTab
4170        With Me
4180          intRetVal = 0
4190          .priceimport_id_lbx.SetFocus
4200        End With
4210      End Select
4220    End If

        ' ** Shift keys.
4230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4240      Select Case intRetVal
          Case vbKeyTab
4250        With Me
4260          intRetVal = 0
4270          If .cmdPrint.Enabled = True Then
4280            .cmdPrint.SetFocus
4290          ElseIf .cmdImport.Enabled = True Then
4300            .cmdImport.SetFocus
4310          ElseIf .priceimport_file.Enabled = True Then
4320            .priceimport_file.SetFocus
4330          Else
4340            .priceimport_id_lbx.SetFocus
4350          End If
4360        End With
4370      End Select
4380    End If

EXITP:
4390    KeyCode = intRetVal
4400    Exit Sub

ERRH:
4410    intRetVal = 0
4420    Select Case ERR.Number
        Case Else
4430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4440    End Select
4450    Resume EXITP

End Sub

Public Sub cmdSave_Click()
' ** This is a fake OnClick event.

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

4510    DoCmd.RunCommand acCmdSaveRecord

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub priceimport_id_lbx_AfterUpdate()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_id_lbx_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strPath As String, strFile As String
        Dim lngDiff As Long

4610    With Me

4620      .cmdImport.Enabled = False
4630      .cmdPreview.Enabled = False
4640      .cmdPrint.Enabled = False
4650      .cmdPreviewPrint_lbl.ForeColor = WIN_CLR_DISF
4660      .cmdPreviewPrint_lbl_dim_hi.Visible = True
4670      dblProgBar_Len = 0#: intPct = 0
4680      ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Procedure: Below.
          '.ProgBar_bar.Width = dblProgBar_Len
4690      .ProgBar_lbl.Caption = CStr(intPct) & "%"
4700      .ProgBar_lbl_dim_hi.Caption = CStr(intPct) & "%"
4710      If .cmdImport.Enabled = False Then
4720        .ProgBar_lbl.ForeColor = WIN_CLR_DISF
4730        If .ProgBar_lbl.Visible = True Then
4740          .ProgBar_lbl_dim_hi.Visible = True
4750        End If
4760      Else
4770        .ProgBar_lbl.ForeColor = CLR_VDKGRY
4780        .ProgBar_lbl_dim_hi.Visible = False
4790      End If
4800      DoEvents

4810      If IsNull(.priceimport_id_lbx) = False Then
4820        If .priceimport_id_lbx > 0& Then

4830          Set rst = .RecordsetClone
4840          With rst
4850            .FindFirst "[priceimport_id] = " & CStr(Me.priceimport_id_lbx)
4860            If .NoMatch = False Then
4870              Me.Bookmark = .Bookmark
4880            End If
4890            .Close
4900          End With
4910          .priceimport_path.Enabled = True
4920          .priceimport_path.Visible = True
4930          .priceimport_path.BorderColor = CLR_LTBLU2
4940          .priceimport_path.BackStyle = acBackStyleNormal
4950          .priceimport_path_empty.Visible = False
4960          .priceimport_file.Enabled = True
4970          .priceimport_file.BorderColor = CLR_LTBLU2
4980          .priceimport_file.BackStyle = acBackStyleNormal
4990          .priceimport_file.Visible = True
5000          .priceimport_file_empty.Visible = False
5010          .priceimport_lbl.ForeColor = CLR_VDKGRY
5020          .priceimport_lbl_dim_hi.Visible = False
5030          .cmdBrowse.Enabled = True

              ' ** Empty the temporary tables used for the chosen import.
              'CAN THEY IMPORT ONE, AND THEN IMPORT ANOTHER WITHOUT LEAVING THE SCREEN?
5040          Set dbs = CurrentDb
5050          Select Case .priceimport_id_lbx.Column(LBX_PI_NAME)
              Case "EstateVal"
5060            DoCmd.Hourglass True
                ' ** Empty tblPricing_EstateVal1.
5070            Set qdf = dbs.QueryDefs("qryPricing_EstateVal_01")
5080            qdf.Execute
                ' ** Empty tblPricing_EstateVal2.
5090            Set qdf = dbs.QueryDefs("qryPricing_EstateVal_02")
5100            qdf.Execute
5110          Case "Appraise"
5120            DoCmd.Hourglass True
                ' ** Empty tblPricing_Appraise_Raw.
5130            Set qdf = dbs.QueryDefs("qryPricing_Appraise_01")
5140            qdf.Execute
                ' ** Empty tblPricing_Appraise_Header.
5150            Set qdf = dbs.QueryDefs("qryPricing_Appraise_02")
5160            qdf.Execute
                ' ** Empty tblPricing_Appraise_Data.
5170            Set qdf = dbs.QueryDefs("qryPricing_Appraise_03")
5180            qdf.Execute
                ' ** Empty tblPricing_Appraise1.
5190            Set qdf = dbs.QueryDefs("qryPricing_Appraise_30")
5200            qdf.Execute
                ' ** Empty tblPricing_Appraise2.
5210            Set qdf = dbs.QueryDefs("qryPricing_Appraise_31")
5220            qdf.Execute
5230          End Select
5240          dbs.Close
5250          DoEvents

              ' ** If there's already a file listed, and it's there, enable cmdImport.
5260          If IsNull(.priceimport_path) = False And IsNull(.priceimport_file) = False Then
5270            strPath = .priceimport_path
5280            strFile = .priceimport_file
5290            If FileExists(strPath & LNK_SEP & strFile) = True Then
5300              .cmdImport.Enabled = True
5310              .ProgBar_lbl.ForeColor = CLR_VDKGRY
                  '.ProgBar_lbl_dim_hi.Visible = False
5320            End If
5330            lngDiff = (15& * lngTpp)
5340            If (Len(strPath) > 50) And (.priceimport_path.Height = lngPriceImportPath_Height) Then
5350              .cmdBrowse.Top = (lngBrowse_Top + lngDiff)
5360              .cmdBrowse_raised_focus_dots_img.Top = .cmdBrowse.Top
5370              .cmdBrowse_raised_focus_img.Top = .cmdBrowse.Top
5380              .cmdBrowse_raised_img.Top = .cmdBrowse.Top
5390              .cmdBrowse_raised_img_dis.Top = .cmdBrowse.Top
5400              .cmdBrowse_raised_semifocus_dots_img.Top = .cmdBrowse.Top
5410              .cmdBrowse_sunken_focus_dots_img.Top = .cmdBrowse.Top
5420              .priceimport_path.Height = (lngPriceImportPath_Height + lngDiff)
5430            ElseIf (Len(strPath) <= 50) And (.priceimport_path.Height > lngPriceImportPath_Height) Then
5440              .priceimport_path.Height = lngPriceImportPath_Height
5450              .cmdBrowse.Top = lngBrowse_Top
5460              .cmdBrowse_raised_focus_dots_img.Top = .cmdBrowse.Top
5470              .cmdBrowse_raised_focus_img.Top = .cmdBrowse.Top
5480              .cmdBrowse_raised_img.Top = .cmdBrowse.Top
5490              .cmdBrowse_raised_img_dis.Top = .cmdBrowse.Top
5500              .cmdBrowse_raised_semifocus_dots_img.Top = .cmdBrowse.Top
5510              .cmdBrowse_sunken_focus_dots_img.Top = .cmdBrowse.Top
5520            End If
5530          End If

5540  On Error Resume Next
5550          .priceimport_path.SetFocus
5560  On Error GoTo ERRH

5570          DoCmd.Hourglass False

5580        Else
5590          .priceimport_path_empty.Visible = True
5600          .priceimport_path.Enabled = False
5610          .priceimport_path.Visible = False
5620          .priceimport_file_empty.Visible = True
5630          .priceimport_file.Enabled = False
5640          .priceimport_file.Visible = False
              '.priceimport_lbl.BackStyle = acBackStyleTransparent
5650          .priceimport_lbl.ForeColor = WIN_CLR_DISF
5660          .priceimport_lbl_dim_hi.Visible = True
5670          .cmdBrowse.Enabled = False
5680          If .cmdBrowse.Top > lngBrowse_Top Then
5690            .cmdBrowse.Top = lngBrowse_Top
5700            .cmdBrowse_raised_focus_dots_img.Top = lngBrowse_Top
5710            .cmdBrowse_raised_focus_img.Top = lngBrowse_Top
5720            .cmdBrowse_raised_img.Top = lngBrowse_Top
5730            .cmdBrowse_raised_img_dis.Top = lngBrowse_Top
5740            .cmdBrowse_raised_semifocus_dots_img.Top = lngBrowse_Top
5750            .cmdBrowse_sunken_focus_dots_img.Top = lngBrowse_Top
5760          End If
5770        End If
5780      Else
5790        .priceimport_path_empty.Visible = True
5800        .priceimport_path.Enabled = False
5810        .priceimport_path.Visible = False
5820        .priceimport_file_empty.Visible = True
5830        .priceimport_file.Enabled = False
5840        .priceimport_file.Visible = False
            '.priceimport_lbl.BackStyle = acBackStyleTransparent
5850        .priceimport_lbl.ForeColor = WIN_CLR_DISF
5860        .priceimport_lbl_dim_hi.Visible = True
5870        .cmdBrowse.Enabled = False
5880        If .cmdBrowse.Top > lngBrowse_Top Then
5890          .cmdBrowse.Top = lngBrowse_Top
5900          .cmdBrowse_raised_focus_dots_img.Top = lngBrowse_Top
5910          .cmdBrowse_raised_focus_img.Top = lngBrowse_Top
5920          .cmdBrowse_raised_img.Top = lngBrowse_Top
5930          .cmdBrowse_raised_img_dis.Top = lngBrowse_Top
5940          .cmdBrowse_raised_semifocus_dots_img.Top = lngBrowse_Top
5950          .cmdBrowse_sunken_focus_dots_img.Top = lngBrowse_Top
5960        End If
5970      End If

5980      Select Case .cmdBrowse.Enabled
          Case True
5990        .cmdBrowse_raised_img.Visible = True
6000        .cmdBrowse_raised_semifocus_dots_img.Visible = False
6010        .cmdBrowse_raised_focus_img.Visible = False
6020        .cmdBrowse_raised_focus_dots_img.Visible = False
6030        .cmdBrowse_sunken_focus_dots_img.Visible = False
6040        .cmdBrowse_raised_img_dis.Visible = False
6050      Case False
6060        .cmdBrowse_raised_img_dis.Visible = True
6070        .cmdBrowse_raised_img.Visible = False
6080        .cmdBrowse_raised_semifocus_dots_img.Visible = False
6090        .cmdBrowse_raised_focus_img.Visible = False
6100        .cmdBrowse_raised_focus_dots_img.Visible = False
6110        .cmdBrowse_sunken_focus_dots_img.Visible = False
6120      End Select

6130    End With

EXITP:
6140    Set rst = Nothing
6150    Set qdf = Nothing
6160    Set dbs = Nothing
6170    Exit Sub

ERRH:
6180    Select Case ERR.Number
          'Case 2110  ' ** Access can't move the focus to the control '|'.
        Case Else
6190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6200    End Select
6210    Resume EXITP

End Sub

Private Sub priceimport_id_lbx_KeyDown(KeyCode As Integer, Shift As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_id_lbx_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6320    intShiftDown = (Shift And acShiftMask) > 0
6330    intAltDown = (Shift And acAltMask) > 0
6340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6360      Select Case intRetVal
          Case vbKeyTab
6370        With Me
6380          intRetVal = 0
6390          If .priceimport_path.Enabled = True Then
6400            .priceimport_path.SetFocus
6410          Else
6420            .cmdCancel.SetFocus
6430          End If
6440        End With
6450      Case vbKeyReturn
6460        With Me
6470          If .priceimport_path.Enabled = False Then
                ' ** Let the AfterUpdate() handle it.
6480          Else
6490            intRetVal = 0
6500            .priceimport_path.SetFocus
6510          End If
6520        End With
6530      End Select
6540    End If

        ' ** Shift keys.
6550    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6570        With Me
6580          intRetVal = 0
6590          .cmdCancel.SetFocus
6600        End With
6610      End Select
6620    End If

EXITP:
6630    KeyCode = intRetVal
6640    Exit Sub

ERRH:
6650    intRetVal = 0
6660    Select Case ERR.Number
        Case Else
6670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6680    End Select
6690    Resume EXITP

End Sub

Private Sub priceimport_id_lbx_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ** Easter Egg (Secret Key) to Null-out Last File Name (priceimport_file).

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_id_lbx_MouseUp"

6710    With Me
6720      If Button = acRightButton And Shift = acCtrlMask + acAltMask Then
6730        If IsNull(.priceimport_id_lbx) = False Then
6740          If .priceimport_id_lbx > 0& Then
6750            If IsNull(.priceimport_id) = False Then
6760              lngClearPathFileID = .priceimport_id
6770              blnClearPathFile = True
6780              .TimerInterval = 500
6790            End If
6800          End If
6810        End If
6820      End If
6830    End With

EXITP:
6840    Exit Sub

ERRH:
6850    Select Case ERR.Number
        Case Else
6860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6870    End Select
6880    Resume EXITP

End Sub

Private Sub priceimport_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_box_MouseMove"

6910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6920    Exit Sub

ERRH:
6930    Select Case ERR.Number
        Case Else
6940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6950    End Select
6960    Resume EXITP

End Sub

Private Sub priceimport_path_BeforeUpdate(Cancel As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_path_BeforeUpdate"

7010    With Me
7020      If IsNull(.priceimport_path) = False Then
7030        If DirExists(.priceimport_path) = False Then
7040          MsgBox "Folder not found.", vbInformation + vbOKOnly, "Folder Not Found"
7050          Cancel = -1
7060        End If
7070      End If
7080    End With

EXITP:
7090    Exit Sub

ERRH:
7100    Select Case ERR.Number
        Case Else
7110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7120    End Select
7130    Resume EXITP

End Sub

Private Sub priceimport_path_AfterUpdate()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_path_AfterUpdate"

        Dim strPath As String, strFile As String

7210    With Me
7220      If IsNull(.priceimport_path) = False Then
7230        strPath = .priceimport_path
7240        strPath = Trim(strPath)
7250        If strPath <> vbNullString Then
              ' ** Strip the trailing backslash if they put one in.
7260          If Right(.priceimport_path, 1) = LNK_SEP Then
7270            strPath = Left(.priceimport_path, (Len(.priceimport_path) - 1))
7280            .priceimport_path = strPath
7290          End If
7300          If IsNull(.priceimport_file) = False And DirExists(strPath) = True Then  ' ** Module Function: modFileUtilities.
7310            strFile = .priceimport_file
7320            strFile = Trim(strFile)
7330            If strFile <> vbNullString Then
7340              If FileExists(strPath & LNK_SEP & strFile) = True Then
7350                .cmdImport.Enabled = True
7360                .ProgBar_lbl.ForeColor = CLR_VDKGRY
7370                .ProgBar_lbl_dim_hi.Visible = False
7380              End If
7390            End If
7400          End If
7410        End If
7420      End If
7430    End With

EXITP:
7440    Exit Sub

ERRH:
7450    Select Case ERR.Number
        Case Else
7460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7470    End Select
7480    Resume EXITP

End Sub

Private Sub priceimport_path_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_path_MouseMove"

7510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7520    Exit Sub

ERRH:
7530    Select Case ERR.Number
        Case Else
7540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7550    End Select
7560    Resume EXITP

End Sub

Private Sub priceimport_file_BeforeUpdate(Cancel As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_file_BeforeUpdate"

7610    With Me
7620      If IsNull(.priceimport_file) = False Then
7630        If FileExists(.priceimport_path & LNK_SEP & .priceimport_file) = False Then
7640          MsgBox "File not found.", vbInformation + vbOKOnly, "File Not Found"
7650          Cancel = -1
7660        End If
7670      End If
7680    End With

EXITP:
7690    Exit Sub

ERRH:
7700    Select Case ERR.Number
        Case Else
7710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7720    End Select
7730    Resume EXITP

End Sub

Private Sub priceimport_file_AfterUpdate()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_file_AfterUpdate"

        Dim strPath As String, strFile As String

7810    With Me
7820      blnAfterUpdate = True
7830      If IsNull(.priceimport_file) = False Then
7840        strFile = .priceimport_file
7850        strFile = Trim(strFile)
7860        If strFile <> vbNullString Then
7870          If IsNull(.priceimport_path) = False Then
7880            strPath = .priceimport_path
7890            strPath = Trim(strPath)
7900            If strPath <> vbNullString Then
7910              If FileExists(strPath & LNK_SEP & strFile) = True Then
7920                .cmdImport.Enabled = True
7930                .ProgBar_lbl.ForeColor = CLR_VDKGRY
7940                .ProgBar_lbl_dim_hi.Visible = False
                    ' ** Get the date and timestamp of the chosen file.
7950                .priceimport_file_lbl2.Caption = Format(FileDateTime(strPath & LNK_SEP & strFile), "mm/dd/yyyy hh:nn:ss AM/PM")
7960                .priceimport_file_lbl2.Visible = True
7970                .priceimport_file_lbl3.Visible = False
7980              End If
7990            End If
8000          End If
8010        End If
8020      End If
8030    End With

EXITP:
8040    Exit Sub

ERRH:
8050    Select Case ERR.Number
        Case Else
8060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8070    End Select
8080    Resume EXITP

End Sub

Private Sub priceimport_file_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_file_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8120    intShiftDown = (Shift And acShiftMask) > 0
8130    intAltDown = (Shift And acAltMask) > 0
8140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8170        With Me
8180          intRetVal = 0
8190          If .cmdBrowse.Enabled = True Then
8200            .cmdBrowse.SetFocus
8210          ElseIf .cmdImport.Enabled = True Then
8220            .cmdImport.SetFocus
8230          ElseIf .cmdPreview.Enabled = True Then
8240            .cmdPreview.SetFocus
8250          Else
8260            .cmdCancel.SetFocus
8270          End If
8280        End With
8290      End Select
8300    End If

        ' ** Shift keys.
8310    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8330        With Me
8340          intRetVal = 0
8350          .priceimport_path.SetFocus
8360        End With
8370      End Select
8380    End If

EXITP:
8390    KeyCode = intRetVal
8400    Exit Sub

ERRH:
8410    intRetVal = 0
8420    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
8430    Case Else
8440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8450    End Select
8460    Resume EXITP

End Sub

Private Sub priceimport_file_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_file_MouseMove"

8510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8520    Exit Sub

ERRH:
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Sub

Private Sub priceimport_file_Exit(Cancel As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "priceimport_file_Exit"

8610    With Me
8620      If IsNull(.priceimport_file) = False And blnAfterUpdate = False Then
            ' ** A file may be there, but the AfterUpdate event didn't fire.
8630        priceimport_file_AfterUpdate  ' ** Procedure: Above.
8640      End If
8650      blnAfterUpdate = False
8660    End With

EXITP:
8670    Exit Sub

ERRH:
8680    Select Case ERR.Number
        Case Else
8690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8700    End Select
8710    Resume EXITP

End Sub

Private Sub cmdBrowse_GotFocus()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_GotFocus"

8810    With Me
8820      blnBrowse_Focus = True
8830      .cmdBrowse_raised_semifocus_dots_img.Visible = True
8840      .cmdBrowse_raised_img.Visible = False
8850      .cmdBrowse_raised_focus_img.Visible = False
8860      .cmdBrowse_raised_focus_dots_img.Visible = False
8870      .cmdBrowse_sunken_focus_dots_img.Visible = False
8880      .cmdBrowse_raised_img_dis.Visible = False
8890    End With

EXITP:
8900    Exit Sub

ERRH:
8910    Select Case ERR.Number
        Case Else
8920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8930    End Select
8940    Resume EXITP

End Sub

Private Sub cmdBrowse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_MouseDown"

9010    With Me
9020      blnBrowse_MouseDown = True
9030      .cmdBrowse_sunken_focus_dots_img.Visible = True
9040      .cmdBrowse_raised_img.Visible = False
9050      .cmdBrowse_raised_semifocus_dots_img.Visible = False
9060      .cmdBrowse_raised_focus_img.Visible = False
9070      .cmdBrowse_raised_focus_dots_img.Visible = False
9080      .cmdBrowse_raised_img_dis.Visible = False
9090    End With

EXITP:
9100    Exit Sub

ERRH:
9110    Select Case ERR.Number
        Case Else
9120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9130    End Select
9140    Resume EXITP

End Sub

Private Sub cmdBrowse_Click()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strFilter As String
        Dim strSearchPathFile As String, strThisPricePath As String, strPath As String
        Dim lngDiff As Long
        Dim blnContinue As Boolean
        Dim strTmp01 As String
        Dim lngX As Long

9210    blnContinue = True

9220    strFilter = vbNullString

9230    With Me

          ' ** Create the Windows Filter String.
9240      For lngX = 0& To (lngFilts - 1&)
9250        If arr_varFilt(FLT_PID, lngX) = .priceimport_id Then
9260          strFilter = strFilter & arr_varFilt(FLT_NAM, lngX) & vbNullChar & arr_varFilt(FLT_EXT, lngX) & vbNullChar
9270        End If
9280      Next

9290      If blnIMPORT_EmptyPricingPath = True Then
            ' ** Check for both old and new directories:
            ' **   Old: gstrDir_PricingOld ("C:\My Documents\EVP Systems\EstateVal\")
            ' **   New: gstrTrustDataLocation & "Pricing\" (strIMPORT_PricingPath)
9300        If DirExists(gstrDir_PricingOld) = True Then
              ' ** Use gstrDir_PricingOld in the popup.
9310          strThisPricePath = Left(gstrDir_PricingOld, (Len(gstrDir_PricingOld) - 1))  ' ** Strip the trailing backslash.
9320          If Dir(strThisPricePath & LNK_SEP & "*.*") = vbNullString Then
                ' ** It's empty, so disregard it.
9330            strThisPricePath = strIMPORT_PricingPath  ' ** The Pricing folder checked for during Form_Open().
9340          End If
9350        Else
9360          strThisPricePath = strIMPORT_PricingPath
9370        End If
9380      Else
9390        strThisPricePath = strIMPORT_PricingPath
9400      End If
          ' ** Now, after all the above, if there's a path already saved, use it.
9410      If IsNull(.priceimport_path) = False Then
9420        strTmp01 = .priceimport_path
9430        strTmp01 = Trim(strTmp01)
9440        If strTmp01 <> vbNullString Then
9450          If DirExists(strTmp01) = True Then
9460            strThisPricePath = strTmp01
9470          End If
9480        End If
9490      End If

9500      strSearchPathFile = GetOpenFileSIS(strFilter, strThisPricePath, "Find Pricing File")  ' ** Module Function: modBrowseFilesAndFolders.
9510      If strSearchPathFile = vbNullString Then
            ' ** They chose nothing, so we do nothing.
9520      Else
9530        If FileExists(strSearchPathFile) = False Then  ' ** Module Function: modFileUtilities.
9540          MsgBox "Pricing file not found!", vbInformation + vbOKOnly, "File Not Found"
9550          strSearchPathFile = vbNullString
9560        Else

              ' ** List box column constants: priceimport_id_lbx.
              ' **   LBX_PI_ID   : priceimport_id.
              ' **   LBX_PI_NAME : priceimport_name.
              ' **   LBX_PI_PUB  : priceimport_publisher.
              ' **   LBX_PI_DESC : priceimport_description.
              ' **   LBX_PI_EXT  : priceimport_extension.

9570          strIMPORT_FileType = Right(strSearchPathFile, 4)
9580          If Left(strIMPORT_FileType, 1) = "." Then
9590            If InStr(.priceimport_id_lbx.Column(LBX_PI_EXT), Mid(strIMPORT_FileType, 2)) > 0 Then  ' ** Strip the period for the InStr().
9600              blnContinue = True
9610            Else
                  ' ** Not a valid extension for the chosen import type.
9620              MsgBox "File type not recognized for this type of import:" & vbCrLf & _
                    "    " & strIMPORT_FileType, vbInformation + vbOKOnly, "Unknown File Type"
9630            End If
9640          Else
                ' ** Not a 3-character file extension!
9650            MsgBox "File type not recognized:" & vbCrLf & _
                  "    " & strIMPORT_FileType, vbInformation + vbOKOnly, "Unknown File Type"
9660          End If

9670          If blnContinue = True Then
9680            strTmp01 = Parse_Path(strSearchPathFile)  ' ** Module Function: modFileUtilities.
9690            .priceimport_path = strTmp01
9700            strTmp01 = Parse_File(strSearchPathFile)  ' ** Module Function: modFileUtilities.
9710            .priceimport_file = strTmp01
9720            strTmp01 = .priceimport_path & LNK_SEP & .priceimport_file
                ' ** Get the date and timestamp of the chosen file.
9730            If Dir(strTmp01) <> vbNullString Then
9740              .priceimport_file_lbl2.Caption = Format(FileDateTime(strTmp01), "mm/dd/yyyy hh:nn:ss AM/PM")
9750              .priceimport_file_lbl2.Visible = True
9760            End If
9770            strPath = .priceimport_path
9780            lngDiff = (15& * lngTpp)
9790            If Len(strPath) > 50 And .priceimport_path.Height = lngPriceImportPath_Height Then
9800              .cmdBrowse.Top = .cmdBrowse.Top + lngDiff
9810              .cmdBrowse_raised_focus_dots_img.Top = .cmdBrowse.Top
9820              .cmdBrowse_raised_focus_img.Top = .cmdBrowse.Top
9830              .cmdBrowse_raised_img.Top = .cmdBrowse.Top
9840              .cmdBrowse_raised_img_dis.Top = .cmdBrowse.Top
9850              .cmdBrowse_raised_semifocus_dots_img.Top = .cmdBrowse.Top
9860              .cmdBrowse_sunken_focus_dots_img.Top = .cmdBrowse.Top
9870              .priceimport_path.Height = (lngPriceImportPath_Height + lngDiff)  ' ** 510 Twips.
9880            ElseIf Len(strPath) <= 50 And .priceimport_path.Height > lngPriceImportPath_Height Then
9890              .priceimport_path.Height = lngPriceImportPath_Height
9900              .cmdBrowse.Top = .cmdBrowse.Top - lngDiff
9910              .cmdBrowse_raised_focus_dots_img.Top = .cmdBrowse.Top
9920              .cmdBrowse_raised_focus_img.Top = .cmdBrowse.Top
9930              .cmdBrowse_raised_img.Top = .cmdBrowse.Top
9940              .cmdBrowse_raised_img_dis.Top = .cmdBrowse.Top
9950              .cmdBrowse_raised_semifocus_dots_img.Top = .cmdBrowse.Top
9960              .cmdBrowse_sunken_focus_dots_img.Top = .cmdBrowse.Top
9970            End If
9980            .cmdImport.Enabled = True
9990            .ProgBar_lbl.ForeColor = CLR_VDKGRY
10000           .ProgBar_lbl_dim_hi.Visible = False
10010           .priceimport_file_lbl3.Visible = False
10020           .cmdImport.SetFocus
10030         End If

10040       End If
10050     End If

10060   End With

EXITP:
10070   Set rst = Nothing
10080   Set qdf = Nothing
10090   Set dbs = Nothing
10100   Exit Sub

ERRH:
10110   Select Case ERR.Number
        Case Else
10120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10130   End Select
10140   Resume EXITP

End Sub

Private Sub cmdBrowse_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10220   intShiftDown = (Shift And acShiftMask) > 0
10230   intAltDown = (Shift And acAltMask) > 0
10240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10260     Select Case intRetVal
          Case vbKeyTab
10270       With Me
10280         intRetVal = 0
10290         If .cmdImport.Enabled = True Then
10300           .cmdImport.SetFocus
10310         ElseIf .cmdPreview.Enabled = True Then
10320           .cmdPreview.SetFocus
10330         Else
10340           .cmdCancel.SetFocus
10350         End If
10360       End With
10370     End Select
10380   End If

        ' ** Shift keys.
10390   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10400     Select Case intRetVal
          Case vbKeyTab
10410       With Me
10420         intRetVal = 0
10430         .priceimport_file.SetFocus
10440       End With
10450     End Select
10460   End If

EXITP:
10470   KeyCode = intRetVal
10480   Exit Sub

ERRH:
10490   intRetVal = 0
10500   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
10510   Case Else
10520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10530   End Select
10540   Resume EXITP

End Sub

Private Sub cmdBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_MouseMove"

10610   With Me
10620     If blnBrowse_MouseDown = False Then
10630       Select Case blnBrowse_Focus
            Case True
10640         .cmdBrowse_raised_focus_dots_img.Visible = True
10650         .cmdBrowse_raised_focus_img.Visible = False
10660       Case False
10670         .cmdBrowse_raised_focus_img.Visible = True
10680         .cmdBrowse_raised_focus_dots_img.Visible = False
10690       End Select
10700       .cmdBrowse_raised_img.Visible = False
10710       .cmdBrowse_raised_semifocus_dots_img.Visible = False
10720       .cmdBrowse_sunken_focus_dots_img.Visible = False
10730       .cmdBrowse_raised_img_dis.Visible = False
10740     End If
10750   End With

EXITP:
10760   Exit Sub

ERRH:
10770   Select Case ERR.Number
        Case Else
10780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10790   End Select
10800   Resume EXITP

End Sub

Private Sub cmdBrowse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_MouseUp"

10910   With Me
10920     .cmdBrowse_raised_focus_dots_img.Visible = True
10930     .cmdBrowse_raised_img.Visible = False
10940     .cmdBrowse_raised_semifocus_dots_img.Visible = False
10950     .cmdBrowse_raised_focus_img.Visible = False
10960     .cmdBrowse_sunken_focus_dots_img.Visible = False
10970     .cmdBrowse_raised_img_dis.Visible = False
10980     blnBrowse_MouseDown = False
10990   End With

EXITP:
11000   Exit Sub

ERRH:
11010   Select Case ERR.Number
        Case Else
11020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11030   End Select
11040   Resume EXITP

End Sub

Private Sub cmdBrowse_LostFocus()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_LostFocus"

11110   With Me
11120     .cmdBrowse_raised_img.Visible = True
11130     .cmdBrowse_raised_semifocus_dots_img.Visible = False
11140     .cmdBrowse_raised_focus_img.Visible = False
11150     .cmdBrowse_raised_focus_dots_img.Visible = False
11160     .cmdBrowse_sunken_focus_dots_img.Visible = False
11170     .cmdBrowse_raised_img_dis.Visible = False
11180     blnBrowse_Focus = False
11190   End With

EXITP:
11200   Exit Sub

ERRH:
11210   Select Case ERR.Number
        Case Else
11220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11230   End Select
11240   Resume EXITP

End Sub

Private Sub cmdImport_Click()
' ** Import and update.

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdImport_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strPath As String, strFile As String
        Dim lngDupes As Long, arr_varDupe() As Variant
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim lngTmp05 As Long, datTmp06 As Date, datTmp07 As Date
        Dim msgResponse As VbMsgBoxResult
        Dim lngX As Long
        Dim blnContinue As Boolean

11310   blnContinue = True

11320   With Me

11330     If IsNull(.priceimport_path) = True Or IsNull(.priceimport_file) = True Then
11340       blnContinue = False
11350       MsgBox "You must choose both a path and file.", vbInformation + vbOKOnly, "Nothing To Do"
11360     Else
11370       If DirExists(.priceimport_path) = False Then
11380         blnContinue = False
11390         MsgBox "Folder not found.", vbInformation + vbOKOnly, "Folder Not Found"
11400       Else
11410         If FileExists(.priceimport_path & LNK_SEP & .priceimport_file) = False Then
11420           blnContinue = False
11430           MsgBox "File not found.", vbInformation + vbOKOnly, "File Not Found"
11440         End If
11450       End If
11460     End If

11470     If blnContinue = True Then

11480       DoCmd.Hourglass True

            ' ** If it's not in our standard Pricing directory, copy it there.
11490       strPath = .priceimport_path
11500       strFile = .priceimport_file
11510       If strIMPORT_FileType = vbNullString Then
              ' ** It won't get filled unless they use cmdBrowse.
11520         strIMPORT_FileType = Right(strFile, 4)  ' ** Include the period.
11530       End If
11540       If strPath <> strIMPORT_PricingPath Then
11550         If FileExists(strIMPORT_PricingPath & LNK_SEP & strFile) = True Then
                ' ** If the file is of the same name, that means it hasn't been processed.
11560           datTmp06 = FileDateTime(strPath & LNK_SEP & strFile)
11570           datTmp07 = FileDateTime(strIMPORT_PricingPath & LNK_SEP & strFile)
11580           If GetDateLong(datTmp06) = GetDateLong(datTmp07) Then
                  ' ** it's the same day, just copy over it.
11590             Kill (strIMPORT_PricingPath & LNK_SEP & strFile)
11600           Else
11610             msgResponse = MsgBox("An earlier unprocessed file of the same name" & vbCrLf & _
                    "already exists in the Pricing folder:" & vbCrLf & _
                    "  " & strFile & "  " & Format(datTmp07, "mm/dd/yyyy hh:nn:ss") & vbCrLf & vbCrLf & _
                    "Do you wish to save that earlier file?" & vbCrLf & vbCrLf & _
                    "Click Yes to rename the file with a 'bak' suffix." & vbCrLf & _
                    "Click No to write over the file with the new one.", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "Save Or Write Over Earlier File")
11620             If msgResponse = vbYes Then
11630               strTmp01 = (strIMPORT_PricingPath & LNK_SEP & strFile)
11640               strTmp02 = Left(strTmp01, (Len(strTmp01) - 4)) & "_bak1" & Right(strTmp01, 4)
11650               If FileExists(strTmp02) = True Then
                      ' ** Previous bak's already exist!
11660                 lngDupes = 0&
11670                 ReDim arr_varDupe(0)
                      ' ** I don't know in what order these bak's will come out!
11680                 strTmp03 = Dir((Left(strTmp01, (Len(strTmp01) - 4)) & "_bak") & "*.*")
11690                 Do While strTmp03 <> vbNullString
11700                   lngDupes = lngDupes + 1&
11710                   ReDim Preserve arr_varDupe(lngDupes - 1&)
11720                   arr_varDupe(lngDupes - 1&) = strTmp03
11730                   strTmp03 = Dir()  ' ** Ask for the next one.
11740                   If strTmp03 = vbNullString Then
11750                     Exit Do
11760                   End If
11770                 Loop
11780                 If lngDupes > 0& Then
11790                   lngTmp05 = 0&
11800                   For lngX = 0& To (lngDupes - 1&)
11810                     intPos01 = InStr(arr_varDupe(lngX), "_bak")
11820                     If Val(Mid(arr_varDupe(lngX), (intPos01 + 4))) > lngTmp05 Then
                            ' ** Val() always converts up to the first non-numeric character, unlike CLng().
11830                       lngTmp05 = Val(Mid(arr_varDupe(lngX), (intPos01 + 4)))
11840                     End If
11850                   Next
11860                   strTmp02 = Left(strTmp01, (Len(strTmp01) - 4)) & "_bak" & CStr(lngTmp05 + 1&) & Right(strTmp01, 4)
11870                   Name strTmp01 As strTmp02
11880                 End If
11890               Else
11900                 Name strTmp01 As strTmp02
11910               End If
11920             Else
11930               Kill (strIMPORT_PricingPath & LNK_SEP & strFile)
11940             End If
11950           End If

11960         End If
11970         FileCopy (strPath & LNK_SEP & strFile), (strIMPORT_PricingPath & LNK_SEP & strFile)
11980       End If

11990       ProgBar_Width_Pric True, 1  ' ** Procedure: Below.
            '.ProgBar_bar.Visible = True
12000       .ProgBar_box.Visible = True
12010       .ProgBar_box2.Visible = True
12020       .ProgBar_lbl.Visible = True
12030       .ProgBar_lbl.ForeColor = CLR_VDKGRY
12040       .ProgBar_lbl_dim_hi.Visible = False

            ' ** GetData1() updates blnUpdateSuccess variable.
12050       GetData1 Me, blnUpdateSuccess, lngFileTypes, arr_varFileType, strIMPORT_FileType, _
              dblProgBox_Width, dblProgBar_Len, intPct, dblProgBar_Incr  ' ** Module Procedure: modPricingFuncs.

            ' ** We need a place to update the other fields if chkIncludeCurrency = True.
12060       IncludeCurrency_Update  ' ** Procedure: Below.

            ' ** Update the progress bar to show complete.
12070       ProgBar_Width_Pric dblProgBox_Width, 2  ' ** Procedure: Below.
            '.ProgBar_bar.Width = dblProgBox_Width
12080       intPct = 100
12090       .ProgBar_lbl.Caption = CStr(intPct) & "%"
12100       DoEvents

12110       If blnUpdateSuccess = True Then

              'This needs to be straightened out!
              'The idea of renaming is because an exported file may always have the same name.
              'So we import it, then change its name by adding the date/time suffix.
              'Do we want to retain the original timestamp somewhere?
              'A table?
              'Within its name?
              'A backup directory would have the same problem with duplicate names.
              'If it's already got our date/time suffix, add an iteration number?
              '_01_110418_0757
              'If Right("xx", 12) is our format, Then
              '  If it has an iteration number, Then
              '    Increment it
              '  Else
              '    Add one with 02
              '  End If
              'Else
              '  Add our date/time suffix, with iteration 01.
              'End If

              ' ** Rename the file.
12120         strTmp01 = (strIMPORT_PricingPath & LNK_SEP & strFile)
12130         strTmp02 = Right(strTmp01, 4)  ' ** Just the file extension, with period.
12140         strTmp03 = Left(strTmp01, (Len(strTmp01) - 4))  ' ** Strip off the file extension.
              ' ** New procedure!
12150         intPos01 = 0
12160         If Len(strTmp03) > 12 Then
12170           strTmp04 = Right(strTmp03, 12)
12180           If Left(strTmp04, 1) = "_" And Mid(strTmp04, 8, 1) = "_" Then
12190             intPos01 = ((Len(strTmp03) - 12) + 1)  ' ** Starting position of our date/time suffix.
12200             If intPos01 > 4 Then
                    ' ** Long enough to possibly have our iteration number.
12210               If Mid(strTmp03, (intPos01 - 3), 1) = "_" Then
12220                 If Val(Mid(strTmp03, (intPos01 - 2), 2)) >= 1 Then
                        ' ** Increment iteration number.
12230                   strTmp03 = Left(strTmp03, (intPos01 - 4)) & "_" & Right("00" & CStr(Val(Mid(strTmp03, (intPos01 - 2), 2)) + 1), 2)
12240                   strTmp03 = strTmp03 & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhnn") & strTmp02
12250                 Else
                        ' ** Coincidence, no iteration number, so add one at 02.
12260                   strTmp03 = Left(strTmp03, (intPos01 - 1))
12270                   strTmp03 = strTmp03 & "_02_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhnn") & strTmp02
12280                 End If
12290               Else
                      ' ** No iteration number, so add one at 02.
12300                 strTmp03 = Left(strTmp03, (intPos01 - 1))
12310                 strTmp03 = strTmp03 & "_02_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhnn") & strTmp02
12320               End If
12330             Else
                    ' ** No iteration number, so add one at 02.
12340               strTmp03 = Left(strTmp03, (intPos01 - 1))
12350               strTmp03 = strTmp03 & "_02_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhnn") & strTmp02
12360             End If
12370           Else
                  ' ** Not ours.
12380             strTmp03 = strTmp03 & "_01_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhnn") & strTmp02
12390           End If
12400         Else
                ' ** Not ours.
12410           strTmp03 = strTmp03 & "_01_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhnn") & strTmp02
12420         End If
              'strTmp03 = strTmp03 & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhnn") & strTmp02
12430         If FileExists(strTmp03) = True Then
12440           Kill strTmp03
12450         End If
12460         Name strTmp01 As strTmp03

              ' ** Replace file name with new one.
12470         .priceimport_file = Parse_File(strTmp03)  ' ** Module Function: modFileUtilities.
12480         DoEvents

12490         cmdSave_Click  ' ** Procedure: Above.
12500         DoEvents

12510         Set dbs = CurrentDb
12520         With dbs
                ' ** Update qryPricing_20 (tblPricing_Import, linked to tblPricing_Staging, with .._new fields).
12530           Set qdf = .QueryDefs("qryPricing_21")
12540           qdf.Execute
12550           Set qdf = Nothing
12560           .Close
12570         End With
12580         Set dbs = Nothing

12590         If strIMPORT_FileType = ".prn" Then
                ' ** Update tblPricing_EstateVal2 with info needed for report.
12600           Set dbs = CurrentDb
12610           Set qdf = dbs.QueryDefs("qryPricing_EstateVal_09")
12620           With qdf.Parameters
12630             ![newdat] = Forms("frmAssetPricing").Date_New
12640           End With
12650           qdf.Execute
12660           dbs.Close
12670         End If

12680         .cmdPreview.Enabled = True
12690         .cmdPrint.Enabled = True
12700         .cmdPreviewPrint_lbl.ForeColor = CLR_DKGRY2
12710         .cmdPreviewPrint_lbl_dim_hi.Visible = False

12720         If IsLoaded("frmAssetPricing", acForm) = True Then  ' ** Module Function: modFileUtilities.
12730           Forms("frmAssetPricing").frmAssetPricing_Sub.Form.Requery
12740           Forms("frmAssetPricing").cmdUpdate.Enabled = True
12750           Forms("frmAssetPricing").cmdReset.Enabled = True
12760         End If

              ' ** Only allow one import per session.
12770         .priceimport_id_lbx.Locked = True

12780         .cmdCancel.Caption = "&Close"

12790       End If  ' ** blnUpdateSuccess.

12800       .cmdCancel.SetFocus
12810       .cmdImport.Enabled = False

12820       blnProgBarOff = True
12830       .TimerInterval = 3000&

12840     End If  ' ** blnContinue.

12850   End With

EXITP:
12860   Set rst = Nothing
12870   Set qdf = Nothing
12880   Set dbs = Nothing
12890   Exit Sub

ERRH:
12900   Select Case ERR.Number
        Case Else
12910     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12920   End Select
12930   Resume EXITP

End Sub

Private Sub cmdImport_KeyDown(KeyCode As Integer, Shift As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdImport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13020   intShiftDown = (Shift And acShiftMask) > 0
13030   intAltDown = (Shift And acAltMask) > 0
13040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13060     Select Case intRetVal
          Case vbKeyTab
13070       With Me
13080         intRetVal = 0
13090         If .cmdPreview.Enabled = True Then
13100           .cmdPreview.SetFocus
13110         Else
13120           .cmdCancel.SetFocus
13130         End If
13140       End With
13150     End Select
13160   End If

        ' ** Shift keys.
13170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13180     Select Case intRetVal
          Case vbKeyTab
13190       With Me
13200         intRetVal = 0
13210         .cmdBrowse.SetFocus
13220       End With
13230     End Select
13240   End If

EXITP:
13250   KeyCode = intRetVal
13260   Exit Sub

ERRH:
13270   intRetVal = 0
13280   Select Case ERR.Number
        Case Else
13290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13300   End Select
13310   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

13410   With Me
13420     Select Case .chkPageOf
          Case True
13430       .chkPageOf_lbl.FontBold = True
13440     Case False
13450       .chkPageOf_lbl.FontBold = False
13460     End Select
13470   End With

EXITP:
13480   Exit Sub

ERRH:
13490   Select Case ERR.Number
        Case Else
13500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13510   End Select
13520   Resume EXITP

End Sub

Private Sub cmdPreview_Click()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

13610   Select Case strIMPORT_FileType
        Case ".txt", ".rtf"
13620     strDocName = "rptPricingUpdateCusips_txt"
13630   Case ".prn"
13640     strDocName = "rptPricingUpdateCusips_prn"
13650   Case ".eut"
13660     strDocName = "rptPricingUpdateCusips_txt"
13670   Case Else
13680     strDocName = "rptPricingUpdateCusips_txt"
13690   End Select

13700   gstrReportQuerySpec = strDocName
13710   gstrReportCallingForm = THIS_NAME  ' ** When not vbNullString, sets this form .Visible = False.
13720   DoCmd.OpenReport strDocName, acViewPreview
13730   DoCmd.Maximize
13740   DoCmd.RunCommand acCmdFitToWindow

EXITP:
13750   Exit Sub

ERRH:
13760   DoCmd.Hourglass False
13770   DoCmd.Restore
13780   If Me.Visible = False Then
13790     Me.Visible = True
13800   End If
13810   Select Case ERR.Number
        Case Else
13820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13830   End Select
13840   Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13920   intShiftDown = (Shift And acShiftMask) > 0
13930   intAltDown = (Shift And acAltMask) > 0
13940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13960     Select Case intRetVal
          Case vbKeyTab
13970       With Me
13980         intRetVal = 0
13990         .cmdPrint.SetFocus
14000       End With
14010     End Select
14020   End If

        ' ** Shift keys.
14030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14040     Select Case intRetVal
          Case vbKeyTab
14050       With Me
14060         intRetVal = 0
14070         If .cmdImport.Enabled = True Then
14080           .cmdImport.SetFocus
14090         ElseIf .cmdBrowse.Enabled = True Then
14100           .cmdBrowse.SetFocus
14110         Else
14120           .priceimport_id_lbx.SetFocus
14130         End If
14140       End With
14150     End Select
14160   End If

EXITP:
14170   KeyCode = intRetVal
14180   Exit Sub

ERRH:
14190   intRetVal = 0
14200   Select Case ERR.Number
        Case Else
14210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14220   End Select
14230   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

14310   Select Case strIMPORT_FileType
        Case ".txt", ".rtf"
14320     strDocName = "rptPricingUpdateCusips_txt"
14330   Case ".prn"
14340     strDocName = "rptPricingUpdateCusips_prn"
14350   Case ".eut"
14360     strDocName = "rptPricingUpdateCusips_txt"
14370   Case Else
14380     strDocName = "rptPricingUpdateCusips_txt"
14390   End Select

14400   gstrReportQuerySpec = strDocName
14410   gstrReportCallingForm = vbNullString  ' ** When vbNullString, leaves this form .Visible = True.
14420   If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
14430     DoCmd.OpenReport strDocName, acViewPreview
14440     DoCmd.Maximize
14450     DoCmd.RunCommand acCmdFitToWindow
14460   Else
          '##GTR_Ref: rptPricingUpdateCusips_prn
          '##GTR_Ref: rptPricingUpdateCusips_txt
14470     DoCmd.OpenReport strDocName, acViewNormal
14480   End If

EXITP:
14490   Exit Sub

ERRH:
14500   Select Case ERR.Number
        Case Else
14510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14520   End Select
14530   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14620   intShiftDown = (Shift And acShiftMask) > 0
14630   intAltDown = (Shift And acAltMask) > 0
14640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14660     Select Case intRetVal
          Case vbKeyTab
14670       With Me
14680         intRetVal = 0
14690         .cmdCancel.SetFocus
14700       End With
14710     End Select
14720   End If

        ' ** Shift keys.
14730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14740     Select Case intRetVal
          Case vbKeyTab
14750       With Me
14760         intRetVal = 0
14770         .cmdPreview.SetFocus
14780       End With
14790     End Select
14800   End If

EXITP:
14810   KeyCode = intRetVal
14820   Exit Sub

ERRH:
14830   intRetVal = 0
14840   Select Case ERR.Number
        Case Else
14850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14860   End Select
14870   Resume EXITP

End Sub

Private Sub IncludeCurrency_Update()
' ** Update additional fields if chkIncludeCurrency = True.

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Update"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim strQryName As String
        Dim lngRecs As Long
        Dim dblTmp01 As Double
        Dim lngX As Long

14910   Set frm = Forms(strCallingForm)
14920   With frm

14930     Select Case .opgShow
          Case .opgShow_optAll.OptionValue
14940       strQryName = "qryPricing_01_02_04"
14950     Case .opgShow_optHeldOnly.OptionValue
14960       strQryName = "qryPricing_01_01_04"
14970     End Select

14980     Set dbs = CurrentDb
14990     Set qdf = dbs.QueryDefs(strQryName)
15000     Set rst = qdf.OpenRecordset
15010     With rst
15020       If .BOF = True And .EOF = True Then
              ' ** Shouldn't happen.
15030       Else
15040         .MoveLast
15050         lngRecs = .RecordCount
15060         .MoveFirst
15070         For lngX = 1& To lngRecs
15080           If ![marketvalue_upd] = True Then
                  ' ** This record has an updated price from here.
15090             If ![curr_id] = 150& Then
15100               .Edit
15110               ![marketvalue_usd] = ![marketvalue]
15120               .Update
15130             Else
15140               dblTmp01 = (![marketvalue] * ![curr_rate2])
15150               dblTmp01 = Round(dblTmp01, 2)
15160               .Edit
15170               ![marketvalue_usd] = dblTmp01
15180               .Update
15190             End If
15200           End If
15210           If lngX < lngRecs Then .MoveNext
15220         Next
15230       End If
15240       .Close
15250     End With  ' ** rst.
15260     Set rst = Nothing
15270     Set qdf = Nothing
15280     dbs.Close
15290     Set dbs = Nothing

15300   End With  ' ** frm.

EXITP:
15310   Set frm = Nothing
15320   Set rst = Nothing
15330   Set qdf = Nothing
15340   Set dbs = Nothing
15350   Exit Sub

ERRH:
15360   DoCmd.Hourglass False
15370   Select Case ERR.Number
        Case Else
15380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15390   End Select
15400   Resume EXITP

End Sub

Public Sub ProgBar_Width_Pric(dblWidth As Double, intMode As Integer)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "ProgBar_Width_Pric"

        Dim strCtlName As String, blnVis As Boolean
        Dim lngX As Long

15510   With Me
15520     Select Case intMode
          Case 1
15530       blnVis = CBool(dblWidth)
15540       For lngX = 1& To 6&
15550         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
15560         .Controls(strCtlName).Visible = blnVis
15570       Next
15580     Case 2
15590       For lngX = 1& To 6&
15600         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
15610         .Controls(strCtlName).Width = dblWidth
15620       Next
15630     End Select
15640   End With

EXITP:
15650   Exit Sub

ERRH:
15660   Select Case ERR.Number
        Case Else
15670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15680   End Select
15690   Resume EXITP

End Sub
