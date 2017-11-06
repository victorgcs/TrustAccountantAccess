Attribute VB_Name = "zz_mod_LTCL_STCGL_Funcs"
Option Compare Database
Option Explicit

'VGC 10/29/2014: CHANGES!

Private Const THIS_NAME As String = "zz_mod_LTCL_STCGL_Funcs"
' **

Public Function NewJrnlMap_Pos() As Boolean

  Const THIS_PROC As String = "NewJrnlMap_Pos"

  Dim dbs As DAO.Database, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngCtls As Long, arr_varCtl() As Variant
  Dim strFormName As String
  Dim blnAdd As Boolean
  Dim varTmp00 As Variant
  Dim lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCtl().
  Const C_ELEMS As Integer = 30  ' ** Array's first-element UBound().
  Const C_DID  As Integer = 0
  Const C_FID  As Integer = 1
  Const C_FNAM As Integer = 2
  Const C_CID  As Integer = 3
  Const C_CNAM As Integer = 4
  Const C_CTYP As Integer = 5
  Const C_TOP  As Integer = 6
  Const C_LFT  As Integer = 7
  Const C_WDT  As Integer = 8
  Const C_HGT  As Integer = 9
  Const C_CAP  As Integer = 10
  Const C_FCLR As Integer = 11
  Const C_BCLR As Integer = 12
  Const C_FONT As Integer = 13
  Const C_SIZE As Integer = 14
  Const C_BOLD As Integer = 15
  Const C_SPEC As Integer = 16
  Const C_BSTL As Integer = 17
  Const C_RCLR As Integer = 18
  Const C_RSTL As Integer = 19
  Const C_RWDT As Integer = 20
  Const C_SLNT As Integer = 21
  Const C_VIS  As Integer = 22
  Const C_ABLE As Integer = 23
  Const C_LOCK As Integer = 24
  Const C_DEF  As Integer = 25
  Const C_OPTV As Integer = 26
  Const C_TAB  As Integer = 27
  Const C_MODE As Integer = 28
  Const C_AUTO As Integer = 29
  Const C_UPD  As Integer = 30

  blnRetVal = True

  If Forms.Count > 0 Then
    Do While Forms.Count > 0
      DoCmd.Close acForm, Forms(0).Name
      DoEvents
    Loop
  End If

  For lngW = 1& To 10&

    lngCtls = 0&
    ReDim arr_varCtl(C_ELEMS, 0)

    strFormName = vbNullString
    Select Case lngW
    Case 1&
      strFormName = "frmMenu_Main"
    Case 2&
      strFormName = "frmMenu_Account"
    Case 3&
      strFormName = "frmMenu_Post"
    Case 4&
      strFormName = "frmMenu_Report"
    Case 5&
      strFormName = "frmMenu_Asset"
    Case 6&
      strFormName = "frmMenu_Utility"
    Case 7&
      strFormName = "frmMenu_CourtReport"
    Case 8&
      strFormName = "frmMenu_Maintenance"
    Case 9&
      strFormName = "frmMenu_Other"
    Case 10&
      strFormName = "frmMenu_ForeignExchange"
    Case 11&
      strFormName = "frmOptions"
    Case 12&
      strFormName = "frmJournal"
    Case 13&
      strFormName = "frmJournal_Sub1_Dividend"
    Case 14&
      strFormName = "frmJournal_Sub2_Interest"
    Case 15&
      strFormName = "frmJournal_Sub3_Purchase"
    Case 16&
      strFormName = "frmJournal_Sub4_Sold"
    Case 17&
      strFormName = "frmJournal_Sub5_Misc"
    End Select

    DoCmd.OpenForm strFormName, acDesign, , , , acHidden

    Set frm = Forms(strFormName)
    With frm
      For Each ctl In .Controls
        With ctl
          lngCtls = lngCtls + 1&
          lngE = lngCtls - 1&
          ReDim Preserve arr_varCtl(C_ELEMS, lngE)
          arr_varCtl(C_FNAM, lngE) = strFormName
          arr_varCtl(C_CNAM, lngE) = .Name
          arr_varCtl(C_CTYP, lngE) = .ControlType
          arr_varCtl(C_TOP, lngE) = .Top
          arr_varCtl(C_LFT, lngE) = .Left
          arr_varCtl(C_WDT, lngE) = .Width
          arr_varCtl(C_HGT, lngE) = .Height
          Select Case .ControlType
          Case acLabel
            arr_varCtl(C_CAP, lngE) = NullIfNullStr(.Caption)  ' ** Module Function: modStringFuncs.
            arr_varCtl(C_FCLR, lngE) = .ForeColor
            arr_varCtl(C_BCLR, lngE) = .BackColor
            arr_varCtl(C_FONT, lngE) = .FontName
            arr_varCtl(C_SIZE, lngE) = .FontSize
            arr_varCtl(C_BOLD, lngE) = .FontBold
            arr_varCtl(C_SPEC, lngE) = .SpecialEffect
            arr_varCtl(C_BSTL, lngE) = .BackStyle
            arr_varCtl(C_RCLR, lngE) = .BorderColor
            arr_varCtl(C_RSTL, lngE) = .BorderStyle
            arr_varCtl(C_RWDT, lngE) = .BorderWidth
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = CBool(False)
            arr_varCtl(C_LOCK, lngE) = CBool(False)
            arr_varCtl(C_DEF, lngE) = Null
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = CBool(False)
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          Case acCommandButton
            arr_varCtl(C_CAP, lngE) = NullIfNullStr(.Caption)  ' ** Module Function: modStringFuncs.
            arr_varCtl(C_FCLR, lngE) = .ForeColor
            arr_varCtl(C_BCLR, lngE) = Null
            arr_varCtl(C_FONT, lngE) = .FontName
            arr_varCtl(C_SIZE, lngE) = .FontSize
            arr_varCtl(C_BOLD, lngE) = .FontBold
            arr_varCtl(C_SPEC, lngE) = Null
            arr_varCtl(C_BSTL, lngE) = Null
            arr_varCtl(C_RCLR, lngE) = Null
            arr_varCtl(C_RSTL, lngE) = Null
            arr_varCtl(C_RWDT, lngE) = Null
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = .Enabled
            arr_varCtl(C_LOCK, lngE) = CBool(False)
            arr_varCtl(C_DEF, lngE) = Null
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = .TabStop
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          Case acTextBox
            arr_varCtl(C_CAP, lngE) = Null
            arr_varCtl(C_FCLR, lngE) = .ForeColor
            arr_varCtl(C_BCLR, lngE) = .BackColor
            arr_varCtl(C_FONT, lngE) = .FontName
            arr_varCtl(C_SIZE, lngE) = .FontSize
            arr_varCtl(C_BOLD, lngE) = .FontBold
            arr_varCtl(C_SPEC, lngE) = .SpecialEffect
            arr_varCtl(C_BSTL, lngE) = .BackStyle
            arr_varCtl(C_RCLR, lngE) = .BorderColor
            arr_varCtl(C_RSTL, lngE) = .BorderStyle
            arr_varCtl(C_RWDT, lngE) = .BorderWidth
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = .Enabled
            arr_varCtl(C_LOCK, lngE) = .Locked
            arr_varCtl(C_DEF, lngE) = NullIfNullStr(.DefaultValue)  ' ** Module Function: modStringFuncs.
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = .TabStop
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          Case acComboBox
            arr_varCtl(C_CAP, lngE) = Null
            arr_varCtl(C_FCLR, lngE) = .ForeColor
            arr_varCtl(C_BCLR, lngE) = .BackColor
            arr_varCtl(C_FONT, lngE) = .FontName
            arr_varCtl(C_SIZE, lngE) = .FontSize
            arr_varCtl(C_BOLD, lngE) = .FontBold
            arr_varCtl(C_SPEC, lngE) = .SpecialEffect
            arr_varCtl(C_BSTL, lngE) = .BackStyle
            arr_varCtl(C_RCLR, lngE) = .BorderColor
            arr_varCtl(C_RSTL, lngE) = .BorderStyle
            arr_varCtl(C_RWDT, lngE) = .BorderWidth
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = .Enabled
            arr_varCtl(C_LOCK, lngE) = .Locked
            arr_varCtl(C_DEF, lngE) = NullIfNullStr(.DefaultValue)  ' ** Module Function: modStringFuncs.
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = .TabStop
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          Case acOptionGroup
            arr_varCtl(C_CAP, lngE) = Null
            arr_varCtl(C_FCLR, lngE) = Null
            arr_varCtl(C_BCLR, lngE) = .BackColor
            arr_varCtl(C_FONT, lngE) = Null
            arr_varCtl(C_SIZE, lngE) = Null
            arr_varCtl(C_BOLD, lngE) = CBool(False)
            arr_varCtl(C_SPEC, lngE) = .SpecialEffect
            arr_varCtl(C_BSTL, lngE) = .BackStyle
            arr_varCtl(C_RCLR, lngE) = .BorderColor
            arr_varCtl(C_RSTL, lngE) = .BorderStyle
            arr_varCtl(C_RWDT, lngE) = .BorderWidth
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = .Enabled
            arr_varCtl(C_LOCK, lngE) = .Locked
            arr_varCtl(C_DEF, lngE) = NullIfNullStr(.DefaultValue)  ' ** Module Function: modStringFuncs.
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = .TabStop
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          Case acOptionButton
            arr_varCtl(C_CAP, lngE) = Null
            arr_varCtl(C_FCLR, lngE) = Null
            arr_varCtl(C_BCLR, lngE) = Null
            arr_varCtl(C_FONT, lngE) = Null
            arr_varCtl(C_SIZE, lngE) = Null
            arr_varCtl(C_BOLD, lngE) = CBool(False)
            arr_varCtl(C_SPEC, lngE) = .SpecialEffect
            arr_varCtl(C_BSTL, lngE) = Null
            arr_varCtl(C_RCLR, lngE) = .BorderColor
            arr_varCtl(C_RSTL, lngE) = .BorderStyle
            arr_varCtl(C_RWDT, lngE) = Null
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = .Enabled
            arr_varCtl(C_LOCK, lngE) = .Locked
            arr_varCtl(C_DEF, lngE) = Null
            arr_varCtl(C_OPTV, lngE) = .OptionValue
            arr_varCtl(C_TAB, lngE) = CBool(False)
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          Case acBoundObjectFrame
            arr_varCtl(C_CAP, lngE) = Null
            arr_varCtl(C_FCLR, lngE) = Null
            arr_varCtl(C_BCLR, lngE) = .BackColor
            arr_varCtl(C_FONT, lngE) = Null
            arr_varCtl(C_SIZE, lngE) = Null
            arr_varCtl(C_BOLD, lngE) = CBool(False)
            arr_varCtl(C_SPEC, lngE) = .SpecialEffect
            arr_varCtl(C_BSTL, lngE) = .BackStyle
            arr_varCtl(C_RCLR, lngE) = .BorderColor
            arr_varCtl(C_RSTL, lngE) = .BorderStyle
            arr_varCtl(C_RWDT, lngE) = .BorderWidth
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = .Enabled
            arr_varCtl(C_LOCK, lngE) = .Locked
            arr_varCtl(C_DEF, lngE) = Null
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = .TabStop
            arr_varCtl(C_MODE, lngE) = .SizeMode
            arr_varCtl(C_AUTO, lngE) = .AutoActivate
            arr_varCtl(C_UPD, lngE) = .UpdateOptions
          Case acCheckBox
            arr_varCtl(C_CAP, lngE) = Null
            arr_varCtl(C_FCLR, lngE) = Null
            arr_varCtl(C_BCLR, lngE) = Null
            arr_varCtl(C_FONT, lngE) = Null
            arr_varCtl(C_SIZE, lngE) = Null
            arr_varCtl(C_BOLD, lngE) = CBool(False)
            arr_varCtl(C_SPEC, lngE) = .SpecialEffect
            arr_varCtl(C_BSTL, lngE) = Null
            arr_varCtl(C_RCLR, lngE) = .BorderColor
            arr_varCtl(C_RSTL, lngE) = .BorderStyle
            arr_varCtl(C_RWDT, lngE) = Null
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = .Enabled
            arr_varCtl(C_LOCK, lngE) = .Locked
            arr_varCtl(C_DEF, lngE) = NullIfNullStr(.DefaultValue)  ' ** Module Function: modStringFuncs.
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = .TabStop
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          Case acLine
            arr_varCtl(C_CAP, lngE) = Null
            arr_varCtl(C_FCLR, lngE) = Null
            arr_varCtl(C_FONT, lngE) = Null
            arr_varCtl(C_SIZE, lngE) = Null
            arr_varCtl(C_BOLD, lngE) = CBool(False)
            arr_varCtl(C_SPEC, lngE) = .SpecialEffect
            arr_varCtl(C_BSTL, lngE) = Null
            arr_varCtl(C_RCLR, lngE) = .BorderColor
            arr_varCtl(C_RSTL, lngE) = .BorderStyle
            arr_varCtl(C_RWDT, lngE) = .BorderWidth
            arr_varCtl(C_SLNT, lngE) = .LineSlant
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = CBool(False)
            arr_varCtl(C_LOCK, lngE) = CBool(False)
            arr_varCtl(C_DEF, lngE) = Null
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = CBool(False)
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          Case acRectangle
            arr_varCtl(C_CAP, lngE) = Null
            arr_varCtl(C_FCLR, lngE) = Null
            arr_varCtl(C_BCLR, lngE) = .BackColor
            arr_varCtl(C_FONT, lngE) = Null
            arr_varCtl(C_SIZE, lngE) = Null
            arr_varCtl(C_BOLD, lngE) = CBool(False)
            arr_varCtl(C_SPEC, lngE) = .SpecialEffect
            arr_varCtl(C_BSTL, lngE) = .BackStyle
            arr_varCtl(C_RCLR, lngE) = .BorderColor
            arr_varCtl(C_RSTL, lngE) = .BorderStyle
            arr_varCtl(C_RWDT, lngE) = .BorderWidth
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = .Visible
            arr_varCtl(C_ABLE, lngE) = CBool(False)
            arr_varCtl(C_LOCK, lngE) = CBool(False)
            arr_varCtl(C_DEF, lngE) = Null
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = CBool(False)
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          Case Else
            arr_varCtl(C_CAP, lngE) = Null
            arr_varCtl(C_FCLR, lngE) = Null
            arr_varCtl(C_BCLR, lngE) = Null
            arr_varCtl(C_FONT, lngE) = Null
            arr_varCtl(C_SIZE, lngE) = Null
            arr_varCtl(C_BOLD, lngE) = CBool(False)
            arr_varCtl(C_SPEC, lngE) = Null
            arr_varCtl(C_BSTL, lngE) = Null
            arr_varCtl(C_RCLR, lngE) = Null
            arr_varCtl(C_RSTL, lngE) = Null
            arr_varCtl(C_SLNT, lngE) = Null
            arr_varCtl(C_VIS, lngE) = CBool(False)
            arr_varCtl(C_ABLE, lngE) = CBool(False)
            arr_varCtl(C_LOCK, lngE) = CBool(False)
            arr_varCtl(C_DEF, lngE) = Null
            arr_varCtl(C_OPTV, lngE) = Null
            arr_varCtl(C_TAB, lngE) = CBool(False)
            arr_varCtl(C_MODE, lngE) = Null
            arr_varCtl(C_AUTO, lngE) = Null
            arr_varCtl(C_UPD, lngE) = Null
          End Select
        End With
      Next
    End With

    ' ** Binary Sort arr_varCtl() array by tab name.
    For lngX = UBound(arr_varCtl, 2) To 1 Step -1
      For lngY = 0 To (lngX - 1)
        If arr_varCtl(C_CNAM, lngY) > arr_varCtl(C_CNAM, (lngY + 1)) Then
          For lngZ = 0& To C_ELEMS
            varTmp00 = arr_varCtl(lngZ, lngY)
            arr_varCtl(lngZ, lngY) = arr_varCtl(lngZ, (lngY + 1))
            arr_varCtl(lngZ, (lngY + 1)) = varTmp00
            varTmp00 = Empty
          Next  ' ** lngZ.
        End If
      Next  ' ** lngY.
    Next  ' ** lngX.

    Set dbs = CurrentDb
    With dbs
      Set rst = .OpenRecordset("zz_tbl_Form_Control_03", dbOpenDynaset, dbConsistent)
      With rst
        .MoveFirst
        For lngX = 0& To (lngCtls - 1&)
          blnAdd = False
          .FindFirst "[frm_name] = '" & arr_varCtl(C_FNAM, lngX) & "' And [ctl_name] = '" & arr_varCtl(C_CNAM, lngX) & "'"
          If .NoMatch = True Then
            blnAdd = True
          End If
          Select Case blnAdd
          Case True
            .AddNew
            ' ** ![ctltmp_id] : AutoNumber.
            ![frm_name] = arr_varCtl(C_FNAM, lngX)
            ![ctl_name] = arr_varCtl(C_CNAM, lngX)
            ![ctltype_type] = arr_varCtl(C_CTYP, lngX)
            ![ctl_top] = arr_varCtl(C_TOP, lngX)
            ![ctl_left] = arr_varCtl(C_LFT, lngX)
            ![ctl_width] = arr_varCtl(C_WDT, lngX)
            ![ctl_height] = arr_varCtl(C_HGT, lngX)
            ![ctl_found] = False
            ![ctl_caption] = arr_varCtl(C_CAP, lngX)
            ![ctl_forecolor] = arr_varCtl(C_FCLR, lngX)
            ![ctl_fontname] = arr_varCtl(C_FONT, lngX)
            ![ctl_fontsize] = arr_varCtl(C_SIZE, lngX)
            ![ctl_fontbold] = arr_varCtl(C_BOLD, lngX)
            ![ctl_backcolor] = arr_varCtl(C_BCLR, lngX)
            ![ctl_specialeffect] = arr_varCtl(C_SPEC, lngX)
            ![ctl_backstyle] = arr_varCtl(C_BSTL, lngX)
            ![ctl_bordercolor] = arr_varCtl(C_RCLR, lngX)
            ![ctl_borderstyle] = arr_varCtl(C_RSTL, lngX)
            ![ctl_borderwidth] = arr_varCtl(C_RWDT, lngX)
            ![ctl_lineslant] = arr_varCtl(C_SLNT, lngX)
            ![ctl_visible] = arr_varCtl(C_VIS, lngX)
            ![ctl_enabled] = arr_varCtl(C_ABLE, lngX)
            ![ctl_locked] = arr_varCtl(C_LOCK, lngX)
            ![ctl_defaultvalue] = arr_varCtl(C_DEF, lngX)
            ![ctl_optionvalue] = arr_varCtl(C_OPTV, lngX)
            ![ctl_tabstop] = arr_varCtl(C_TAB, lngX)
            ![ctl_sizemode] = arr_varCtl(C_MODE, lngX)
            ![ctl_autoactivate] = arr_varCtl(C_AUTO, lngX)
            ![ctl_updateoptions] = arr_varCtl(C_UPD, lngX)
            ![ctltmp_datemodified] = Now()
            .Update
          Case False
            ' ** ctltype_type.
            If ![ctltype_type] <> arr_varCtl(C_CTYP, lngX) Then
              .Edit
              ![ctltype_type] = arr_varCtl(C_CTYP, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_top.
            If ![ctl_top] <> arr_varCtl(C_TOP, lngX) Then
              .Edit
              ![ctl_top] = arr_varCtl(C_TOP, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_left.
            If ![ctl_left] <> arr_varCtl(C_LFT, lngX) Then
              .Edit
              ![ctl_left] = arr_varCtl(C_LFT, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_width.
            If ![ctl_width] <> arr_varCtl(C_WDT, lngX) Then
              .Edit
              ![ctl_width] = arr_varCtl(C_WDT, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_height.
            If ![ctl_height] <> arr_varCtl(C_HGT, lngX) Then
              .Edit
              ![ctl_height] = arr_varCtl(C_HGT, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_caption.
            Select Case IsNull(arr_varCtl(C_CAP, lngX))
            Case True
              If IsNull(![ctl_caption]) = False Then
                .Edit
                ![ctl_caption] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_caption])
              Case True
                .Edit
                ![ctl_caption] = arr_varCtl(C_CAP, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_caption] <> arr_varCtl(C_CAP, lngX) Then
                  .Edit
                  ![ctl_caption] = arr_varCtl(C_CAP, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_forecolor.
            Select Case IsNull(arr_varCtl(C_FCLR, lngX))
            Case True
              If IsNull(![ctl_forecolor]) = False Then
                .Edit
                ![ctl_forecolor] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_forecolor])
              Case True
                .Edit
                ![ctl_forecolor] = arr_varCtl(C_FCLR, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_forecolor] <> arr_varCtl(C_FCLR, lngX) Then
                  .Edit
                  ![ctl_forecolor] = arr_varCtl(C_FCLR, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_backcolor.
            Select Case IsNull(arr_varCtl(C_BCLR, lngX))
            Case True
              If IsNull(![ctl_backcolor]) = False Then
                .Edit
                ![ctl_backcolor] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_backcolor])
              Case True
                .Edit
                ![ctl_backcolor] = arr_varCtl(C_BCLR, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_backcolor] <> arr_varCtl(C_BCLR, lngX) Then
                  .Edit
                  ![ctl_backcolor] = arr_varCtl(C_BCLR, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_fontname.
            Select Case IsNull(arr_varCtl(C_FONT, lngX))
            Case True
              If IsNull(![ctl_fontname]) = False Then
                .Edit
                ![ctl_fontname] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_fontname])
              Case True
                .Edit
                ![ctl_fontname] = arr_varCtl(C_FONT, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_fontname] <> arr_varCtl(C_FONT, lngX) Then
                  .Edit
                  ![ctl_fontname] = arr_varCtl(C_FONT, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_fontsize.
            Select Case IsNull(arr_varCtl(C_SIZE, lngX))
            Case True
              If IsNull(![ctl_fontsize]) = False Then
                .Edit
                ![ctl_fontsize] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_fontsize])
              Case True
                .Edit
                ![ctl_fontsize] = arr_varCtl(C_SIZE, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_fontsize] <> arr_varCtl(C_SIZE, lngX) Then
                  .Edit
                  ![ctl_fontsize] = arr_varCtl(C_SIZE, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_fontbold.
            If ![ctl_fontbold] <> arr_varCtl(C_BOLD, lngX) Then
              .Edit
              ![ctl_fontbold] = arr_varCtl(C_BOLD, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_specialeffect.
            Select Case IsNull(arr_varCtl(C_SPEC, lngX))
            Case True
              If IsNull(![ctl_specialeffect]) = False Then
                .Edit
                ![ctl_specialeffect] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_specialeffect])
              Case True
                .Edit
                ![ctl_specialeffect] = arr_varCtl(C_SPEC, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_specialeffect] <> arr_varCtl(C_SPEC, lngX) Then
                  .Edit
                  ![ctl_specialeffect] = arr_varCtl(C_SPEC, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_backstyle.
            Select Case IsNull(arr_varCtl(C_BSTL, lngX))
            Case True
              If IsNull(![ctl_backstyle]) = False Then
                .Edit
                ![ctl_backstyle] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_backstyle])
              Case True
                .Edit
                ![ctl_backstyle] = arr_varCtl(C_BSTL, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_backstyle] <> arr_varCtl(C_BSTL, lngX) Then
                  .Edit
                  ![ctl_backstyle] = arr_varCtl(C_BSTL, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_bordercolor.
            Select Case IsNull(arr_varCtl(C_RCLR, lngX))
            Case True
              If IsNull(![ctl_bordercolor]) = False Then
                .Edit
                ![ctl_bordercolor] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_bordercolor])
              Case True
                .Edit
                ![ctl_bordercolor] = arr_varCtl(C_RCLR, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_bordercolor] <> arr_varCtl(C_RCLR, lngX) Then
                  .Edit
                  ![ctl_bordercolor] = arr_varCtl(C_RCLR, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select

            ' ** ctl_borderstyle.
            Select Case IsNull(arr_varCtl(C_RSTL, lngX))
            Case True
              If IsNull(![ctl_borderstyle]) = False Then
                .Edit
                ![ctl_borderstyle] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_borderstyle])
              Case True
                .Edit
                ![ctl_borderstyle] = arr_varCtl(C_RSTL, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_borderstyle] <> arr_varCtl(C_RSTL, lngX) Then
                  .Edit
                  ![ctl_borderstyle] = arr_varCtl(C_RSTL, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_borderwidth.
            Select Case IsNull(arr_varCtl(C_RWDT, lngX))
            Case True
              If IsNull(![ctl_borderwidth]) = False Then
                .Edit
                ![ctl_borderwidth] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_borderwidth])
              Case True
                .Edit
                ![ctl_borderwidth] = arr_varCtl(C_RWDT, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_borderwidth] <> arr_varCtl(C_RWDT, lngX) Then
                  .Edit
                  ![ctl_borderwidth] = arr_varCtl(C_RWDT, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_lineslant.
            Select Case IsNull(arr_varCtl(C_SLNT, lngX))
            Case True
              If IsNull(![ctl_lineslant]) = False Then
                .Edit
                ![ctl_lineslant] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_lineslant])
              Case True
                .Edit
                ![ctl_lineslant] = arr_varCtl(C_SLNT, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_lineslant] <> arr_varCtl(C_SLNT, lngX) Then
                  .Edit
                  ![ctl_lineslant] = arr_varCtl(C_SLNT, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_visible.
            If ![ctl_visible] <> arr_varCtl(C_VIS, lngX) Then
              .Edit
              ![ctl_visible] = arr_varCtl(C_VIS, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_enabled.
            If ![ctl_enabled] <> arr_varCtl(C_ABLE, lngX) Then
              .Edit
              ![ctl_enabled] = arr_varCtl(C_ABLE, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_locked.
            If ![ctl_locked] <> arr_varCtl(C_LOCK, lngX) Then
              .Edit
              ![ctl_locked] = arr_varCtl(C_LOCK, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_defaultvalue.
            Select Case IsNull(arr_varCtl(C_DEF, lngX))
            Case True
              If IsNull(![ctl_defaultvalue]) = False Then
                .Edit
                ![ctl_defaultvalue] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_defaultvalue])
              Case True
                .Edit
                ![ctl_defaultvalue] = arr_varCtl(C_DEF, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_defaultvalue] <> arr_varCtl(C_DEF, lngX) Then
                  .Edit
                  ![ctl_defaultvalue] = arr_varCtl(C_DEF, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_optionvalue.
            Select Case IsNull(arr_varCtl(C_OPTV, lngX))
            Case True
              If IsNull(![ctl_optionvalue]) = False Then
                .Edit
                ![ctl_optionvalue] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_optionvalue])
              Case True
                .Edit
                ![ctl_optionvalue] = arr_varCtl(C_OPTV, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_optionvalue] <> arr_varCtl(C_OPTV, lngX) Then
                  .Edit
                  ![ctl_optionvalue] = arr_varCtl(C_OPTV, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_tabstop.
            If ![ctl_tabstop] <> arr_varCtl(C_TAB, lngX) Then
              .Edit
              ![ctl_tabstop] = arr_varCtl(C_TAB, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_sizemode.
            Select Case IsNull(arr_varCtl(C_MODE, lngX))
            Case True
              If IsNull(![ctl_sizemode]) = False Then
                .Edit
                ![ctl_sizemode] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_sizemode])
              Case True
                .Edit
                ![ctl_sizemode] = arr_varCtl(C_MODE, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_sizemode] <> arr_varCtl(C_MODE, lngX) Then
                  .Edit
                  ![ctl_sizemode] = arr_varCtl(C_MODE, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_autoactivate.
            Select Case IsNull(arr_varCtl(C_AUTO, lngX))
            Case True
              If IsNull(![ctl_autoactivate]) = False Then
                .Edit
                ![ctl_autoactivate] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_autoactivate])
              Case True
                .Edit
                ![ctl_autoactivate] = arr_varCtl(C_AUTO, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_autoactivate] <> arr_varCtl(C_AUTO, lngX) Then
                  .Edit
                  ![ctl_autoactivate] = arr_varCtl(C_AUTO, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_updateoptions.
            Select Case IsNull(arr_varCtl(C_UPD, lngX))
            Case True
              If IsNull(![ctl_updateoptions]) = False Then
                .Edit
                ![ctl_updateoptions] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_updateoptions])
              Case True
                .Edit
                ![ctl_updateoptions] = arr_varCtl(C_UPD, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_updateoptions] <> arr_varCtl(C_UPD, lngX) Then
                  .Edit
                  ![ctl_updateoptions] = arr_varCtl(C_UPD, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
          End Select
        Next
        .Close
      End With
      .Close
    End With

    DoCmd.Close acForm, strFormName, acSaveNo

  Next  ' ** lngW.

  Debug.Print "'DONE!"

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set dbs = Nothing

  NewJrnlMap_Pos = blnRetVal

End Function

Public Function NewJrnlMap_Chk() As Boolean

  Const THIS_PROC As String = "NewJrnlMap_Chk"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim lngCtlsNotFound As Long, arr_varCtlNotFound() As Variant
  Dim lngCtlsHereOnly As Long, arr_varCtlHereOnly() As Variant
  Dim lngCtlsNowhere As Long, arr_varCtlNowhere() As Variant
  Dim lngCtls_Jrnl As Long, lngCtls_Div As Long, lngCtls_Int As Long, lngCtls_Purch As Long
  Dim lngCtls_Sold As Long, lngCtls_Misc As Long, lngCtls_Opts As Long
  Dim lngCtls_Main As Long, lngCtls_Account As Long, lngCtls_Post As Long, lngCtls_Report As Long, lngCtls_Asset As Long
  Dim lngCtls_Utility As Long, lngCtls_Court As Long, lngCtls_Maint As Long, lngCtls_Other As Long
  Dim lngThisDbsID As Long
  Dim strFormName As String, lngMaxWidth As Long
  Dim lngRecs As Long
  Dim blnFound As Boolean, blnAdd As Boolean, blnDelete As Boolean
  Dim varTmp00 As Variant
  Dim lngW As Long, lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCtl().
  Const C_CTID As Integer = 0
  Const C_DID  As Integer = 1
  Const C_FID  As Integer = 2
  Const C_FNAM As Integer = 3
  Const C_CID  As Integer = 4
  Const C_CNAM As Integer = 5
  Const C_CTYP As Integer = 6
  Const C_TOP  As Integer = 7
  Const C_LFT  As Integer = 8
  Const C_WDT  As Integer = 9
  Const C_HGT  As Integer = 10
  Const C_CAP  As Integer = 11
  Const C_FCLR As Integer = 12
  Const C_BCLR As Integer = 13
  Const C_FONT As Integer = 14
  Const C_SIZE As Integer = 15
  Const C_BOLD As Integer = 16
  Const C_SPEC As Integer = 17
  Const C_BSTL As Integer = 18
  Const C_RCLR As Integer = 19
  Const C_RSTL As Integer = 20
  Const C_RWDT As Integer = 21
  Const C_SLNT As Integer = 22
  Const C_VIS  As Integer = 23
  Const C_ABLE As Integer = 24
  Const C_LOCK As Integer = 25
  Const C_DEF  As Integer = 26
  Const C_OPTV As Integer = 27
  Const C_TAB  As Integer = 28
  Const C_MODE As Integer = 29
  Const C_AUTO As Integer = 30
  Const C_UPD  As Integer = 31
  Const C_FND  As Integer = 32
  Const C_HERE As Integer = 33
  Const C_DATM As Integer = 34

  ' ** Array: arr_varCtlNotFound().
  Const NF_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const NF_FNAM As Integer = 0
  Const NF_CNAM As Integer = 1

  ' ** Array: arr_varCtlHereOnly().
  Const HO_ELEMS As Integer = 32  ' ** Array's first-element UBound().
  Const HO_CTID As Integer = 0
  Const HO_DID  As Integer = 1
  Const HO_FID  As Integer = 2
  Const HO_FNAM As Integer = 3
  Const HO_CID  As Integer = 4
  Const HO_CNAM As Integer = 5
  Const HO_CTYP As Integer = 6
  Const HO_TOP  As Integer = 7
  Const HO_LFT  As Integer = 8
  Const HO_WDT  As Integer = 9
  Const HO_HGT  As Integer = 10
  Const HO_CAP  As Integer = 11
  Const HO_FCLR As Integer = 12
  Const HO_BCLR As Integer = 13
  Const HO_FONT As Integer = 14
  Const HO_SIZE As Integer = 15
  Const HO_BOLD As Integer = 16
  Const HO_SPEC As Integer = 17
  Const HO_BSTL As Integer = 18
  Const HO_RCLR As Integer = 19
  Const HO_RSTL As Integer = 20
  Const HO_RWDT As Integer = 21
  Const HO_SLNT As Integer = 22
  Const HO_VIS  As Integer = 23
  Const HO_ABLE As Integer = 24
  Const HO_LOCK As Integer = 25
  Const HO_DEF  As Integer = 26
  Const HO_OPTV As Integer = 27
  Const HO_TAB  As Integer = 28
  Const HO_MODE As Integer = 29
  Const HO_AUTO As Integer = 30
  Const HO_UPD  As Integer = 31
  Const HO_FND  As Integer = 32

  ' ** Array: arr_varCtlNowhere().
  Const NW_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const NW_CTID As Integer = 0
  Const NW_FNAM As Integer = 1
  Const NW_CNAM As Integer = 2

  blnRetVal = True

  If Forms.Count > 0 Then
    Do While Forms.Count > 0
      DoCmd.Close acForm, Forms(0).Name
      DoEvents
    Loop
  End If

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_Form_Control_03, just ctl_hereonly = False.
    'Set qdf = .QueryDefs("zzz_qry_NewJrnlMap_01")
    ' ** zz_tbl_Form_Control_03, just 'frmMenu..', ctl_hereonly = False.
    Set qdf = .QueryDefs("zzz_qry_xMenuColors_03")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' ********************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     ctltmp_id              C_CTID
      ' **     2       1     dbs_id                 C_DID
      ' **     3       2     frm_id                 C_FID
      ' **     4       3     frm_name               C_FNAM
      ' **     5       4     ctl_id                 C_CID
      ' **     6       5     ctl_name               C_CNAM
      ' **     7       6     ctltype_type           C_CTYP
      ' **     8       7     ctl_top                C_TOP
      ' **     9       8     ctl_left               C_LFT
      ' **    10       9     ctl_width              C_WDT
      ' **    11      10     ctl_height             C_HGT
      ' **    12      11     ctl_caption            C_CAP
      ' **    13      12     ctl_forecolor          C_FCLR
      ' **    14      13     ctl_backcolor          C_BCLR
      ' **    15      14     ctl_fontname           C_FONT
      ' **    16      15     ctl_fontsize           C_SIZE
      ' **    17      16     ctl_bontbold           C_BOLD
      ' **    18      17     ctl_specialeffect      C_SPEC
      ' **    19      18     ctl_backstyle          C_BSTL
      ' **    20      19     ctl_bordercolor        C_RCLR
      ' **    21      20     ctl_borderstyle        C_RSTL
      ' **    22      21     ctl_borderwidth        C_RWDT
      ' **    23      22     ctl_lineslant          C_SLNT
      ' **    24      23     ctl_visible            C_VIS
      ' **    25      24     ctl_enabled            C_ABLE
      ' **    26      25     ctl_locked             C_LOCK
      ' **    27      26     ctl_defaultvalue       C_DEF
      ' **    28      27     ctl_optionvalue        C_OPTV
      ' **    29      28     ctl_tabstop            C_TAB
      ' **    30      29     ctl_sizemode           C_MODE
      ' **    31      30     ctl_autoactivate       C_AUTO
      ' **    32      31     ctl_updateoptions      C_UPD
      ' **    33      32     ctl_found              C_FND
      ' **    34      33     ctl_hereonly           C_HERE
      ' **    35      34     ctltmp_datemodified    C_DATM
      ' **
      ' ********************************************************
      .Close
    End With  ' ** rst.
    Set rst = Nothing
    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  lngCtls_Jrnl = 0&: lngCtls_Div = 0&: lngCtls_Int = 0&: lngCtls_Purch = 0&: lngCtls_Sold = 0&: lngCtls_Misc = 0&: lngCtls_Opts = 0&
  lngCtls_Main = 0&: lngCtls_Account = 0&: lngCtls_Post = 0&: lngCtls_Report = 0&: lngCtls_Asset = 0&
  lngCtls_Utility = 0&: lngCtls_Court = 0&: lngCtls_Maint = 0&: lngCtls_Other = 0&

  For lngX = 0& To (lngCtls - 1&)
    Select Case arr_varCtl(C_FNAM, lngX)
    Case "frmMenu_Main"
      lngCtls_Main = lngCtls_Main + 1&
    Case "frmMenu_Account"
      lngCtls_Account = lngCtls_Account + 1&
    Case "frmMenu_Post"
      lngCtls_Post = lngCtls_Post + 1&
    Case "frmMenu_Report"
      lngCtls_Report = lngCtls_Report + 1&
    Case "frmMenu_Asset"
      lngCtls_Asset = lngCtls_Asset + 1&
    Case "frmMenu_Utility"
      lngCtls_Utility = lngCtls_Utility + 1&
    Case "frmMenu_CourtReport"
      lngCtls_Court = lngCtls_Court + 1&
    Case "frmMenu_Maintenance"
      lngCtls_Maint = lngCtls_Maint + 1&
    Case "frmMenu_Other"
      lngCtls_Other = lngCtls_Other + 1&
    Case "frmOptions"
      lngCtls_Opts = lngCtls_Opts + 1&
    Case "frmJournal"
      lngCtls_Jrnl = lngCtls_Jrnl + 1&
    Case "frmJournal_Sub1_Dividend"
      lngCtls_Div = lngCtls_Div + 1&
    Case "frmJournal_Sub2_Interest"
      lngCtls_Int = lngCtls_Int + 1&
    Case "frmJournal_Sub3_Purchase"
      lngCtls_Purch = lngCtls_Purch + 1&
    Case "frmJournal_Sub4_Sold"
      lngCtls_Sold = lngCtls_Sold + 1&
    Case "frmJournal_Sub5_Misc"
      lngCtls_Misc = lngCtls_Misc + 1&
    End Select
    If Len(arr_varCtl(C_FNAM, lngX)) > lngMaxWidth Then
      lngMaxWidth = Len(arr_varCtl(C_FNAM, lngX))
    End If
  Next  ' ** lngX.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Debug.Print "'TOT CTLS: " & CStr(lngCtls)
  Debug.Print

  lngCtlsNotFound = 0&
  ReDim arr_varCtlNotFound(NF_ELEMS, 0)

  lngCtlsHereOnly = 0&
  ReDim arr_varCtlHereOnly(HO_ELEMS, 0)

  strFormName = vbNullString
  For lngW = 1& To 9&

    Select Case lngW
    Case 1&
      strFormName = "frmMenu_Main"
    Case 2&
      strFormName = "frmMenu_Account"
    Case 3&
      strFormName = "frmMenu_Post"
    Case 4&
      strFormName = "frmMenu_Report"
    Case 5&
      strFormName = "frmMenu_Asset"
    Case 6&
      strFormName = "frmMenu_Utility"
    Case 7&
      strFormName = "frmMenu_CourtReport"
    Case 8&
      strFormName = "frmMenu_Maintenance"
    Case 9&
      strFormName = "frmMenu_Other"
    Case 10&
      strFormName = "frmOptions"
    Case 11&
      strFormName = "frmJournal"
    Case 12&
      strFormName = "frmJournal_Sub1_Dividend"
    Case 13&
      strFormName = "frmJournal_Sub2_Interest"
    Case 14&
      strFormName = "frmJournal_Sub3_Purchase"
    Case 15&
      strFormName = "frmJournal_Sub4_Sold"
    Case 16&
      strFormName = "frmJournal_Sub5_Misc"
    End Select

    If IsLoaded(strFormName, acForm, True) = False Then  ' ** Module Function: modFileUtilities.
      If Forms.Count > 0 Then
        DoCmd.Close acForm, Forms(0).Name
        Set frm = Nothing
      End If
      DoCmd.OpenForm strFormName, acDesign, , , , acHidden
      Set frm = Forms(0)
      If frm.Name <> strFormName Then
        Stop
      End If
    End If

    With frm

      For lngX = 0& To (lngCtls - 1&)
        If arr_varCtl(C_FNAM, lngX) = strFormName Then
          blnFound = False
          For Each ctl In .Controls
            With ctl
              If .Name = arr_varCtl(C_CNAM, lngX) Then
                blnFound = True
                arr_varCtl(C_FND, lngX) = CBool(True)
                Exit For
              End If
            End With
          Next  ' ** ctl.
          Set ctl = Nothing
          If blnFound = False Then
            lngCtlsNotFound = lngCtlsNotFound + 1&
            lngE = lngCtlsNotFound - 1&
            ReDim Preserve arr_varCtlNotFound(NF_ELEMS, lngE)
            arr_varCtlNotFound(NF_FNAM, lngE) = strFormName
            arr_varCtlNotFound(NF_CNAM, lngE) = arr_varCtl(C_CNAM, lngX)
          End If
        End If
      Next  ' ** lngX.

      Select Case lngW
      Case 1&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Main) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 2&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Account) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 3&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Post) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 4&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Report) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 5&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Asset) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 6&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Utility) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 7&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Court) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 8&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Maint) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 9&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Other) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 10&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Opts) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 11&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Jrnl) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 12&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Div) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 13&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Int) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 14&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Purch) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 15&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Sold) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      Case 16&
        Debug.Print "'" & Left(strFormName & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          "V2.2.30 CTLS: " & Left(CStr(lngCtls_Misc) & "   ", 3) & "  V2.2.20 CTLS: "; CStr(.Controls.Count)
      End Select
      DoEvents

      For Each ctl In .Controls
        blnFound = False
        With ctl
          For lngX = 0& To (lngCtls - 1&)
            If arr_varCtl(C_FNAM, lngX) = strFormName Then
              If arr_varCtl(C_CNAM, lngX) = .Name Then
                blnFound = True
                Exit For
              End If
            End If
          Next  ' ** lngX.
          If blnFound = False Then
            lngCtlsHereOnly = lngCtlsHereOnly + 1&
            lngE = lngCtlsHereOnly - 1&
            ReDim Preserve arr_varCtlHereOnly(HO_ELEMS, lngE)
            arr_varCtlHereOnly(HO_CTID, lngE) = CLng(0)
            arr_varCtlHereOnly(HO_DID, lngE) = lngThisDbsID
            arr_varCtlHereOnly(HO_FID, lngE) = CLng(0)
            arr_varCtlHereOnly(HO_FNAM, lngE) = strFormName
            arr_varCtlHereOnly(HO_CID, lngE) = CLng(0)
            arr_varCtlHereOnly(HO_CNAM, lngE) = .Name
            arr_varCtlHereOnly(HO_CTYP, lngE) = .ControlType
            arr_varCtlHereOnly(HO_TOP, lngE) = .Top
            arr_varCtlHereOnly(HO_LFT, lngE) = .Left
            arr_varCtlHereOnly(HO_WDT, lngE) = .Width
            arr_varCtlHereOnly(HO_HGT, lngE) = .Height
            Select Case .ControlType
            Case acLabel
              arr_varCtlHereOnly(HO_CAP, lngE) = .Caption
              arr_varCtlHereOnly(HO_FCLR, lngE) = .ForeColor
              arr_varCtlHereOnly(HO_BCLR, lngE) = .BackColor
              arr_varCtlHereOnly(HO_FONT, lngE) = .FontName
              arr_varCtlHereOnly(HO_SIZE, lngE) = .FontSize
              arr_varCtlHereOnly(HO_BOLD, lngE) = .FontBold
              arr_varCtlHereOnly(HO_SPEC, lngE) = .SpecialEffect
              arr_varCtlHereOnly(HO_BSTL, lngE) = .BackStyle
              arr_varCtlHereOnly(HO_RCLR, lngE) = .BorderColor
              arr_varCtlHereOnly(HO_RSTL, lngE) = .BorderStyle
              arr_varCtlHereOnly(HO_RWDT, lngE) = .BorderWidth
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_LOCK, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_DEF, lngE) = Null
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            Case acCommandButton
              arr_varCtlHereOnly(HO_CAP, lngE) = .Caption
              arr_varCtlHereOnly(HO_FCLR, lngE) = .ForeColor
              arr_varCtlHereOnly(HO_BCLR, lngE) = Null
              arr_varCtlHereOnly(HO_FONT, lngE) = .FontName
              arr_varCtlHereOnly(HO_SIZE, lngE) = .FontSize
              arr_varCtlHereOnly(HO_BOLD, lngE) = .FontBold
              arr_varCtlHereOnly(HO_SPEC, lngE) = Null
              arr_varCtlHereOnly(HO_BSTL, lngE) = Null
              arr_varCtlHereOnly(HO_RCLR, lngE) = Null
              arr_varCtlHereOnly(HO_RSTL, lngE) = Null
              arr_varCtlHereOnly(HO_RWDT, lngE) = Null
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = .Enabled
              arr_varCtlHereOnly(HO_LOCK, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_DEF, lngE) = Null
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = .TabStop
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            Case acTextBox
              arr_varCtlHereOnly(HO_CAP, lngE) = Null
              arr_varCtlHereOnly(HO_FCLR, lngE) = .ForeColor
              arr_varCtlHereOnly(HO_BCLR, lngE) = .BackColor
              arr_varCtlHereOnly(HO_FONT, lngE) = .FontName
              arr_varCtlHereOnly(HO_SIZE, lngE) = .FontSize
              arr_varCtlHereOnly(HO_BOLD, lngE) = .FontBold
              arr_varCtlHereOnly(HO_SPEC, lngE) = .SpecialEffect
              arr_varCtlHereOnly(HO_BSTL, lngE) = .BackStyle
              arr_varCtlHereOnly(HO_RCLR, lngE) = .BorderColor
              arr_varCtlHereOnly(HO_RSTL, lngE) = .BorderStyle
              arr_varCtlHereOnly(HO_RWDT, lngE) = .BorderWidth
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = .Enabled
              arr_varCtlHereOnly(HO_LOCK, lngE) = .Locked
              arr_varCtlHereOnly(HO_DEF, lngE) = NullIfNullStr(.DefaultValue)  ' ** Module Function: modStringFuncs.
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = .TabStop
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            Case acComboBox
              arr_varCtlHereOnly(HO_CAP, lngE) = Null
              arr_varCtlHereOnly(HO_FCLR, lngE) = .ForeColor
              arr_varCtlHereOnly(HO_BCLR, lngE) = .BackColor
              arr_varCtlHereOnly(HO_FONT, lngE) = .FontName
              arr_varCtlHereOnly(HO_SIZE, lngE) = .FontSize
              arr_varCtlHereOnly(HO_BOLD, lngE) = .FontBold
              arr_varCtlHereOnly(HO_SPEC, lngE) = .SpecialEffect
              arr_varCtlHereOnly(HO_BSTL, lngE) = .BackStyle
              arr_varCtlHereOnly(HO_RCLR, lngE) = .BorderColor
              arr_varCtlHereOnly(HO_RSTL, lngE) = .BorderStyle
              arr_varCtlHereOnly(HO_RWDT, lngE) = .BorderWidth
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = .Enabled
              arr_varCtlHereOnly(HO_LOCK, lngE) = .Locked
              arr_varCtlHereOnly(HO_DEF, lngE) = NullIfNullStr(.DefaultValue)  ' ** Module Function: modStringFuncs.
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = .TabStop
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            Case acCheckBox
              arr_varCtlHereOnly(HO_CAP, lngE) = Null
              arr_varCtlHereOnly(HO_FCLR, lngE) = Null
              arr_varCtlHereOnly(HO_BCLR, lngE) = Null
              arr_varCtlHereOnly(HO_FONT, lngE) = Null
              arr_varCtlHereOnly(HO_SIZE, lngE) = Null
              arr_varCtlHereOnly(HO_BOLD, lngE) = Null
              arr_varCtlHereOnly(HO_SPEC, lngE) = .SpecialEffect
              arr_varCtlHereOnly(HO_BSTL, lngE) = Null
              arr_varCtlHereOnly(HO_RCLR, lngE) = .BorderColor
              arr_varCtlHereOnly(HO_RSTL, lngE) = .BorderStyle
              arr_varCtlHereOnly(HO_RWDT, lngE) = Null
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = .Enabled
              arr_varCtlHereOnly(HO_LOCK, lngE) = .Locked
              arr_varCtlHereOnly(HO_DEF, lngE) = NullIfNullStr(.DefaultValue)  ' ** Module Function: modStringFuncs.
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = .TabStop
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            Case acOptionGroup
              arr_varCtlHereOnly(HO_CAP, lngE) = Null
              arr_varCtlHereOnly(HO_FCLR, lngE) = Null
              arr_varCtlHereOnly(HO_BCLR, lngE) = .BackColor
              arr_varCtlHereOnly(HO_FONT, lngE) = Null
              arr_varCtlHereOnly(HO_SIZE, lngE) = Null
              arr_varCtlHereOnly(HO_BOLD, lngE) = Null
              arr_varCtlHereOnly(HO_SPEC, lngE) = .SpecialEffect
              arr_varCtlHereOnly(HO_BSTL, lngE) = .BackStyle
              arr_varCtlHereOnly(HO_RCLR, lngE) = .BorderColor
              arr_varCtlHereOnly(HO_RSTL, lngE) = .BorderStyle
              arr_varCtlHereOnly(HO_RWDT, lngE) = .BorderWidth
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = .Enabled
              arr_varCtlHereOnly(HO_LOCK, lngE) = .Locked
              arr_varCtlHereOnly(HO_DEF, lngE) = NullIfNullStr(.DefaultValue)  ' ** Module Function: modStringFuncs.
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = .TabStop
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            Case acOptionButton
              arr_varCtlHereOnly(HO_CAP, lngE) = Null
              arr_varCtlHereOnly(HO_FCLR, lngE) = Null
              arr_varCtlHereOnly(HO_BCLR, lngE) = Null
              arr_varCtlHereOnly(HO_FONT, lngE) = Null
              arr_varCtlHereOnly(HO_SIZE, lngE) = Null
              arr_varCtlHereOnly(HO_BOLD, lngE) = Null
              arr_varCtlHereOnly(HO_SPEC, lngE) = .SpecialEffect
              arr_varCtlHereOnly(HO_BSTL, lngE) = Null
              arr_varCtlHereOnly(HO_RCLR, lngE) = .BorderColor
              arr_varCtlHereOnly(HO_RSTL, lngE) = .BorderStyle
              arr_varCtlHereOnly(HO_RWDT, lngE) = Null
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = .Enabled
              arr_varCtlHereOnly(HO_LOCK, lngE) = .Locked
              arr_varCtlHereOnly(HO_DEF, lngE) = Null
              arr_varCtlHereOnly(HO_OPTV, lngE) = .OptionValue
              arr_varCtlHereOnly(HO_TAB, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            Case acBoundObjectFrame
              arr_varCtlHereOnly(HO_CAP, lngE) = Null
              arr_varCtlHereOnly(HO_FCLR, lngE) = Null
              arr_varCtlHereOnly(HO_BCLR, lngE) = .BackColor
              arr_varCtlHereOnly(HO_FONT, lngE) = Null
              arr_varCtlHereOnly(HO_SIZE, lngE) = Null
              arr_varCtlHereOnly(HO_BOLD, lngE) = Null
              arr_varCtlHereOnly(HO_SPEC, lngE) = .SpecialEffect
              arr_varCtlHereOnly(HO_BSTL, lngE) = .BackStyle
              arr_varCtlHereOnly(HO_RCLR, lngE) = .BorderColor
              arr_varCtlHereOnly(HO_RSTL, lngE) = .BorderStyle
              arr_varCtlHereOnly(HO_RWDT, lngE) = .BorderWidth
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = .Enabled
              arr_varCtlHereOnly(HO_LOCK, lngE) = .Locked
              arr_varCtlHereOnly(HO_DEF, lngE) = Null
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = .TabStop
              arr_varCtlHereOnly(HO_MODE, lngE) = .SizeMode
              arr_varCtlHereOnly(HO_AUTO, lngE) = .AutoActivate
              arr_varCtlHereOnly(HO_UPD, lngE) = .UpdateOptions
            Case acLine
              arr_varCtlHereOnly(HO_CAP, lngE) = Null
              arr_varCtlHereOnly(HO_FCLR, lngE) = Null
              arr_varCtlHereOnly(HO_BCLR, lngE) = Null
              arr_varCtlHereOnly(HO_FONT, lngE) = Null
              arr_varCtlHereOnly(HO_SIZE, lngE) = Null
              arr_varCtlHereOnly(HO_BOLD, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_SPEC, lngE) = .SpecialEffect
              arr_varCtlHereOnly(HO_BSTL, lngE) = Null
              arr_varCtlHereOnly(HO_RCLR, lngE) = .BorderColor
              arr_varCtlHereOnly(HO_RSTL, lngE) = .BorderStyle
              arr_varCtlHereOnly(HO_RWDT, lngE) = .BorderWidth
              arr_varCtlHereOnly(HO_SLNT, lngE) = .LineSlant
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_LOCK, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_DEF, lngE) = Null
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            Case acRectangle
              arr_varCtlHereOnly(HO_CAP, lngE) = Null
              arr_varCtlHereOnly(HO_FCLR, lngE) = Null
              arr_varCtlHereOnly(HO_BCLR, lngE) = .BackColor
              arr_varCtlHereOnly(HO_FONT, lngE) = Null
              arr_varCtlHereOnly(HO_SIZE, lngE) = Null
              arr_varCtlHereOnly(HO_BOLD, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_SPEC, lngE) = .SpecialEffect
              arr_varCtlHereOnly(HO_BSTL, lngE) = .BackStyle
              arr_varCtlHereOnly(HO_RCLR, lngE) = .BorderColor
              arr_varCtlHereOnly(HO_RSTL, lngE) = .BorderStyle
              arr_varCtlHereOnly(HO_RWDT, lngE) = .BorderWidth
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = .Visible
              arr_varCtlHereOnly(HO_ABLE, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_LOCK, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_DEF, lngE) = Null
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            Case Else
              arr_varCtlHereOnly(HO_CAP, lngE) = Null
              arr_varCtlHereOnly(HO_FCLR, lngE) = Null
              arr_varCtlHereOnly(HO_BCLR, lngE) = Null
              arr_varCtlHereOnly(HO_FONT, lngE) = Null
              arr_varCtlHereOnly(HO_SIZE, lngE) = Null
              arr_varCtlHereOnly(HO_BOLD, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_SPEC, lngE) = Null
              arr_varCtlHereOnly(HO_BSTL, lngE) = Null
              arr_varCtlHereOnly(HO_RCLR, lngE) = Null
              arr_varCtlHereOnly(HO_RSTL, lngE) = Null
              arr_varCtlHereOnly(HO_RWDT, lngE) = Null
              arr_varCtlHereOnly(HO_SLNT, lngE) = Null
              arr_varCtlHereOnly(HO_VIS, lngE) = Null
              arr_varCtlHereOnly(HO_ABLE, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_LOCK, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_DEF, lngE) = Null
              arr_varCtlHereOnly(HO_OPTV, lngE) = Null
              arr_varCtlHereOnly(HO_TAB, lngE) = CBool(False)
              arr_varCtlHereOnly(HO_MODE, lngE) = Null
              arr_varCtlHereOnly(HO_AUTO, lngE) = Null
              arr_varCtlHereOnly(HO_UPD, lngE) = Null
            End Select
            arr_varCtlHereOnly(HO_FND, lngE) = CBool(False)
          End If
        End With  ' ** ctl.
      Next  ' ** ctl.
      Set ctl = Nothing

    End With  ' ** frm.
    DoEvents

  Next  ' ** lngW.

  DoCmd.Close acForm, Forms(0).Name
  DoEvents

  Set dbs = CurrentDb
  With dbs

    Set rst = .OpenRecordset("zz_tbl_Form_Control_03", dbOpenDynaset, dbConsistent)
    With rst
      .MoveFirst
      For lngX = 0& To (lngCtls - 1&)
        .FindFirst "[frm_name] = '" & arr_varCtl(C_FNAM, lngX) & "' And [ctl_name] = '" & arr_varCtl(C_CNAM, lngX) & "'"
        Select Case .NoMatch
        Case True
          Stop
        Case False
          If ![ctl_found] <> arr_varCtl(C_FND, lngX) Then
            .Edit
            ![ctl_found] = arr_varCtl(C_FND, lngX)
            ![ctltmp_datemodified] = Now()
            .Update
          End If
        End Select
      Next
    End With  ' ** rst.

    If lngCtlsNotFound > 0& Then
      Debug.Print
      Debug.Print "'CTLS NOT FOUND: " & CStr(lngCtlsNotFound)
      DoEvents
      strFormName = vbNullString
      For lngX = 0& To (lngCtlsNotFound - 1&)
        If arr_varCtlNotFound(NF_FNAM, lngX) <> strFormName Then
          If strFormName <> vbNullString Then
            Debug.Print
            DoEvents
          End If
          strFormName = arr_varCtlNotFound(NF_FNAM, lngX)
        End If
        Debug.Print "'" & Left(arr_varCtlNotFound(NF_FNAM, lngX) & ":" & Space(lngMaxWidth), (lngMaxWidth + 1&)) & "  " & _
          arr_varCtlNotFound(NF_CNAM, lngX)
        DoEvents
      Next  ' ** lngX.
    End If  ' ** lngCtlsNotFound.

    If lngCtlsHereOnly > 0& Then

      For lngX = 0 To (lngCtlsHereOnly - 1&)
        varTmp00 = DLookup("[frm_id]", "tblForm", "[dbs_id] = " & CStr(arr_varCtlHereOnly(HO_DID, lngX)) & " And " & _
          "[frm_name] = '" & arr_varCtlHereOnly(HO_FNAM, lngX) & "'")
        If IsNull(varTmp00) = False Then
          arr_varCtlHereOnly(HO_FID, lngX) = CLng(varTmp00)
        Else
          Stop
        End If
        varTmp00 = DLookup("[ctl_id]", "tblForm_Control", "[dbs_id] = " & CStr(arr_varCtlHereOnly(HO_DID, lngX)) & " And " & _
          "[frm_id] = " & CStr(arr_varCtlHereOnly(HO_FID, lngX)) & " And " & _
          "[ctl_name] = '" & arr_varCtlHereOnly(HO_CNAM, lngX) & "'")
        If IsNull(varTmp00) = False Then
          arr_varCtlHereOnly(HO_CID, lngX) = CLng(varTmp00)
        Else
          Stop
        End If
      Next  ' ** lngX.

      With rst
        .MoveFirst
        For lngX = 0 To (lngCtlsHereOnly - 1&)
          blnAdd = False
          .FindFirst "[frm_name] = '" & arr_varCtlHereOnly(HO_FNAM, lngX) & "' And [ctl_name] = '" & arr_varCtlHereOnly(HO_CNAM, lngX) & "'"
          If .NoMatch = True Then
            .FindFirst "[frm_name] = '" & arr_varCtlHereOnly(HO_FNAM, lngX) & "' And [ctl_id] = " & CStr(arr_varCtlHereOnly(HO_CID, lngX))
            If .NoMatch = True Then
              blnAdd = True
            End If
          End If
          Select Case blnAdd
          Case True
            .AddNew
            ' ** ![ctltmp_id] : AutoNumber.
            ![dbs_id] = arr_varCtlHereOnly(HO_DID, lngX)
            ![frm_id] = arr_varCtlHereOnly(HO_FID, lngX)
            ![frm_name] = arr_varCtlHereOnly(HO_FNAM, lngX)
            ![ctl_id] = arr_varCtlHereOnly(HO_CID, lngX)
            ![ctl_name] = arr_varCtlHereOnly(HO_CNAM, lngX)
            ![ctltype_type] = arr_varCtlHereOnly(HO_CTYP, lngX)
            ![ctl_top] = arr_varCtlHereOnly(HO_TOP, lngX)
            ![ctl_left] = arr_varCtlHereOnly(HO_LFT, lngX)
            ![ctl_width] = arr_varCtlHereOnly(HO_WDT, lngX)
            ![ctl_height] = arr_varCtlHereOnly(HO_HGT, lngX)
            ![ctl_caption] = arr_varCtlHereOnly(HO_CAP, lngX)
            ![ctl_forecolor] = arr_varCtlHereOnly(HO_FCLR, lngX)
            ![ctl_backcolor] = arr_varCtlHereOnly(HO_BCLR, lngX)
            ![ctl_fontname] = arr_varCtlHereOnly(HO_FONT, lngX)
            ![ctl_fontsize] = arr_varCtlHereOnly(HO_SIZE, lngX)
            ![ctl_fontbold] = arr_varCtlHereOnly(HO_BOLD, lngX)
            ![ctl_specialeffect] = arr_varCtlHereOnly(HO_SPEC, lngX)
            ![ctl_backstyle] = arr_varCtlHereOnly(HO_BSTL, lngX)
            ![ctl_bordercolor] = arr_varCtlHereOnly(HO_RCLR, lngX)
            ![ctl_borderstyle] = arr_varCtlHereOnly(HO_RSTL, lngX)
            ![ctl_borderwidth] = arr_varCtlHereOnly(HO_RWDT, lngX)
            ![ctl_lineslant] = arr_varCtlHereOnly(HO_SLNT, lngX)
            ![ctl_visible] = arr_varCtlHereOnly(HO_VIS, lngX)
            ![ctl_enabled] = arr_varCtlHereOnly(HO_ABLE, lngX)
            ![ctl_locked] = arr_varCtlHereOnly(HO_LOCK, lngX)
            ![ctl_defaultvalue] = arr_varCtlHereOnly(HO_DEF, lngX)
            ![ctl_optionvalue] = arr_varCtlHereOnly(HO_OPTV, lngX)
            ![ctl_tabstop] = arr_varCtlHereOnly(HO_TAB, lngX)
            ![ctl_sizemode] = arr_varCtlHereOnly(HO_MODE, lngX)
            ![ctl_autoactivate] = arr_varCtlHereOnly(HO_AUTO, lngX)
            ![ctl_updateoptions] = arr_varCtlHereOnly(HO_UPD, lngX)
            ![ctl_found] = False
            ![ctl_hereonly] = True
            ![ctltmp_datemodified] = Now()
            .Update
          Case False
            ' ** Name.
            If ![ctl_name] <> arr_varCtlHereOnly(HO_CNAM, lngX) Then
              .Edit
              ![ctl_name] = arr_varCtlHereOnly(HO_CNAM, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** Top.
            If ![ctl_top] <> arr_varCtlHereOnly(HO_TOP, lngX) Then
              .Edit
              ![ctl_top] = arr_varCtlHereOnly(HO_TOP, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** Left.
            If ![ctl_left] <> arr_varCtlHereOnly(HO_LFT, lngX) Then
              .Edit
              ![ctl_left] = arr_varCtlHereOnly(HO_LFT, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** Width.
            If ![ctl_width] <> arr_varCtlHereOnly(HO_WDT, lngX) Then
              .Edit
              ![ctl_width] = arr_varCtlHereOnly(HO_WDT, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** Height.
            If ![ctl_height] <> arr_varCtlHereOnly(HO_HGT, lngX) Then
              .Edit
              ![ctl_height] = arr_varCtlHereOnly(HO_HGT, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** Caption.
            Select Case IsNull(arr_varCtlHereOnly(HO_CAP, lngX))
            Case True
              If IsNull(![ctl_caption]) = False Then
                .Edit
                ![ctl_caption] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_caption])
              Case True
                .Edit
                ![ctl_caption] = arr_varCtlHereOnly(HO_CAP, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_caption] <> arr_varCtlHereOnly(HO_CAP, lngX) Then
                  .Edit
                  ![ctl_caption] = arr_varCtlHereOnly(HO_CAP, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ForeColor.
            Select Case IsNull(arr_varCtlHereOnly(HO_FCLR, lngX))
            Case True
              If IsNull(![ctl_forecolor]) = False Then
                .Edit
                ![ctl_forecolor] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_forecolor])
              Case True
                .Edit
                ![ctl_forecolor] = arr_varCtlHereOnly(HO_FCLR, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_forecolor] <> arr_varCtlHereOnly(HO_FCLR, lngX) Then
                  .Edit
                  ![ctl_forecolor] = arr_varCtlHereOnly(HO_FCLR, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** BackColor.
            Select Case IsNull(arr_varCtlHereOnly(HO_BCLR, lngX))
            Case True
              If IsNull(![ctl_backcolor]) = False Then
                .Edit
                ![ctl_backcolor] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_backcolor])
              Case True
                .Edit
                ![ctl_backcolor] = arr_varCtlHereOnly(HO_BCLR, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_backcolor] <> arr_varCtlHereOnly(HO_BCLR, lngX) Then
                  .Edit
                  ![ctl_backcolor] = arr_varCtlHereOnly(HO_BCLR, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** FontName.
            Select Case IsNull(arr_varCtlHereOnly(HO_FONT, lngX))
            Case True
              If IsNull(![ctl_fontname]) = False Then
                .Edit
                ![ctl_fontname] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_fontname])
              Case True
                .Edit
                ![ctl_fontname] = arr_varCtlHereOnly(HO_FONT, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_fontname] <> arr_varCtlHereOnly(HO_FONT, lngX) Then
                  .Edit
                  ![ctl_fontname] = arr_varCtlHereOnly(HO_FONT, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** FontSize.
            Select Case IsNull(arr_varCtlHereOnly(HO_SIZE, lngX))
            Case True
              If IsNull(![ctl_fontsize]) = False Then
                .Edit
                ![ctl_fontsize] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_fontsize])
              Case True
                .Edit
                ![ctl_fontsize] = arr_varCtlHereOnly(HO_SIZE, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_fontsize] <> arr_varCtlHereOnly(HO_SIZE, lngX) Then
                  .Edit
                  ![ctl_fontsize] = arr_varCtlHereOnly(HO_SIZE, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** FontBold.
            If ![ctl_fontbold] <> arr_varCtlHereOnly(HO_BOLD, lngX) Then
              .Edit
              ![ctl_fontbold] = arr_varCtlHereOnly(HO_BOLD, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_specialeffect.
            Select Case IsNull(arr_varCtlHereOnly(HO_SPEC, lngX))
            Case True
              If IsNull(![ctl_specialeffect]) = False Then
                .Edit
                ![ctl_specialeffect] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_specialeffect])
              Case True
                .Edit
                ![ctl_specialeffect] = arr_varCtlHereOnly(HO_SPEC, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_specialeffect] <> arr_varCtlHereOnly(HO_SPEC, lngX) Then
                  .Edit
                  ![ctl_specialeffect] = arr_varCtlHereOnly(HO_SPEC, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_backstyle.
            Select Case IsNull(arr_varCtlHereOnly(HO_BSTL, lngX))
            Case True
              If IsNull(![ctl_backstyle]) = False Then
                .Edit
                ![ctl_backstyle] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_backstyle])
              Case True
                .Edit
                ![ctl_backstyle] = arr_varCtlHereOnly(HO_BSTL, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_backstyle] <> arr_varCtlHereOnly(HO_BSTL, lngX) Then
                  .Edit
                  ![ctl_backstyle] = arr_varCtlHereOnly(HO_BSTL, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_bordercolor.
            Select Case IsNull(arr_varCtlHereOnly(HO_RCLR, lngX))
            Case True
              If IsNull(![ctl_bordercolor]) = False Then
                .Edit
                ![ctl_bordercolor] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_bordercolor])
              Case True
                .Edit
                ![ctl_bordercolor] = arr_varCtlHereOnly(HO_RCLR, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_bordercolor] <> arr_varCtlHereOnly(HO_RCLR, lngX) Then
                  .Edit
                  ![ctl_bordercolor] = arr_varCtlHereOnly(HO_RCLR, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_borderstyle.
            Select Case IsNull(arr_varCtlHereOnly(HO_RSTL, lngX))
            Case True
              If IsNull(![ctl_borderstyle]) = False Then
                .Edit
                ![ctl_borderstyle] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_borderstyle])
              Case True
                .Edit
                ![ctl_borderstyle] = arr_varCtlHereOnly(HO_RSTL, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_borderstyle] <> arr_varCtlHereOnly(HO_RSTL, lngX) Then
                  .Edit
                  ![ctl_borderstyle] = arr_varCtlHereOnly(HO_RSTL, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_borderwidth.
            Select Case IsNull(arr_varCtlHereOnly(HO_RWDT, lngX))
            Case True
              If IsNull(![ctl_borderwidth]) = False Then
                .Edit
                ![ctl_borderwidth] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_borderwidth])
              Case True
                .Edit
                ![ctl_borderwidth] = arr_varCtlHereOnly(HO_RWDT, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_borderwidth] <> arr_varCtlHereOnly(HO_RWDT, lngX) Then
                  .Edit
                  ![ctl_borderwidth] = arr_varCtlHereOnly(HO_RWDT, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_lineslant.
            Select Case IsNull(arr_varCtlHereOnly(HO_SLNT, lngX))
            Case True
              If IsNull(![ctl_lineslant]) = False Then
                .Edit
                ![ctl_lineslant] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_lineslant])
              Case True
                .Edit
                ![ctl_lineslant] = arr_varCtlHereOnly(HO_SLNT, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_lineslant] <> arr_varCtlHereOnly(HO_SLNT, lngX) Then
                  .Edit
                  ![ctl_lineslant] = arr_varCtlHereOnly(HO_SLNT, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_visible.
            If ![ctl_visible] <> arr_varCtlHereOnly(HO_VIS, lngX) Then
              .Edit
              ![ctl_visible] = arr_varCtlHereOnly(HO_VIS, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_enabled.
            If ![ctl_enabled] <> arr_varCtlHereOnly(HO_ABLE, lngX) Then
              .Edit
              ![ctl_enabled] = arr_varCtlHereOnly(HO_ABLE, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_locked.
            If ![ctl_locked] <> arr_varCtlHereOnly(HO_LOCK, lngX) Then
              .Edit
              ![ctl_locked] = arr_varCtlHereOnly(HO_LOCK, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_defaultvalue.
            Select Case IsNull(arr_varCtlHereOnly(HO_DEF, lngX))
            Case True
              If IsNull(![ctl_defaultvalue]) = False Then
                .Edit
                ![ctl_defaultvalue] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_defaultvalue])
              Case True
                .Edit
                ![ctl_defaultvalue] = arr_varCtlHereOnly(HO_DEF, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_defaultvalue] <> arr_varCtlHereOnly(HO_DEF, lngX) Then
                  .Edit
                  ![ctl_defaultvalue] = arr_varCtlHereOnly(HO_DEF, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_optionvalue.
            Select Case IsNull(arr_varCtlHereOnly(HO_OPTV, lngX))
            Case True
              If IsNull(![ctl_optionvalue]) = False Then
                .Edit
                ![ctl_optionvalue] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_optionvalue])
              Case True
                .Edit
                ![ctl_optionvalue] = arr_varCtlHereOnly(HO_OPTV, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_optionvalue] <> arr_varCtlHereOnly(HO_OPTV, lngX) Then
                  .Edit
                  ![ctl_optionvalue] = arr_varCtlHereOnly(HO_OPTV, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_tabstop.
            If ![ctl_tabstop] <> arr_varCtlHereOnly(HO_TAB, lngX) Then
              .Edit
              ![ctl_tabstop] = arr_varCtlHereOnly(HO_TAB, lngX)
              ![ctltmp_datemodified] = Now()
              .Update
            End If
            ' ** ctl_sizemode.
            Select Case IsNull(arr_varCtlHereOnly(HO_MODE, lngX))
            Case True
              If IsNull(![ctl_sizemode]) = False Then
                .Edit
                ![ctl_sizemode] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_sizemode])
              Case True
                .Edit
                ![ctl_sizemode] = arr_varCtlHereOnly(HO_MODE, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_sizemode] <> arr_varCtlHereOnly(HO_MODE, lngX) Then
                  .Edit
                  ![ctl_sizemode] = arr_varCtlHereOnly(HO_MODE, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_autoactivate.
            Select Case IsNull(arr_varCtlHereOnly(HO_AUTO, lngX))
            Case True
              If IsNull(![ctl_autoactivate]) = False Then
                .Edit
                ![ctl_autoactivate] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_autoactivate])
              Case True
                .Edit
                ![ctl_autoactivate] = arr_varCtlHereOnly(HO_AUTO, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_autoactivate] <> arr_varCtlHereOnly(HO_AUTO, lngX) Then
                  .Edit
                  ![ctl_autoactivate] = arr_varCtlHereOnly(HO_AUTO, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            ' ** ctl_updateoptions.
            Select Case IsNull(arr_varCtlHereOnly(HO_UPD, lngX))
            Case True
              If IsNull(![ctl_updateoptions]) = False Then
                .Edit
                ![ctl_updateoptions] = Null
                ![ctltmp_datemodified] = Now()
                .Update
              End If
            Case False
              Select Case IsNull(![ctl_updateoptions])
              Case True
                .Edit
                ![ctl_updateoptions] = arr_varCtlHereOnly(HO_UPD, lngX)
                ![ctltmp_datemodified] = Now()
                .Update
              Case False
                If ![ctl_updateoptions] <> arr_varCtlHereOnly(HO_UPD, lngX) Then
                  .Edit
                  ![ctl_updateoptions] = arr_varCtlHereOnly(HO_UPD, lngX)
                  ![ctltmp_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
          End Select  ' ** blnAdd.
        Next  ' ** lngX.
      End With  ' ** rst.

      Debug.Print
      Debug.Print "'CTLS HERE ONLY: " & CStr(lngCtlsHereOnly)
      Debug.Print

    End If  ' ** lngCtlsHereOnly.

    lngCtlsNowhere = 0&
    ReDim arr_varCtlNowhere(NW_ELEMS, 0)

    strFormName = "frmOptions"

    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        strFormName = ![frm_name]
        Select Case strFormName
        Case "frmMenu_Main", "frmMenu_Account", "frmMenu_Post", "frmMenu_Report", "frmMenu_Asset", "frmMenu_Utility", _
            "frmMenu_CourtReport", "frmMenu_Maintenance", "frmMenu_Other"
          blnFound = False
          For lngY = 0& To (lngCtls - 1&)
            If arr_varCtl(C_FNAM, lngY) = strFormName Then
              If arr_varCtl(C_CNAM, lngY) = ![ctl_name] Then
                blnFound = True
                Exit For
              End If
            End If
          Next  ' ** lngY.
          If blnFound = False Then
            For lngY = 0& To (lngCtlsHereOnly - 1&)
              If arr_varCtlHereOnly(HO_FNAM, lngY) = strFormName Then
                If arr_varCtlHereOnly(HO_CNAM, lngY) = ![ctl_name] Then
                  blnFound = True
                  Exit For
                End If
              End If
            Next  ' ** lngY.
          End If
          If blnFound = False Then
            lngCtlsNowhere = lngCtlsNowhere + 1&
            lngE = lngCtlsNowhere - 1&
            ReDim Preserve arr_varCtlNowhere(NW_ELEMS, lngE)
            arr_varCtlNowhere(NW_CTID, lngE) = ![ctltmp_id]
            arr_varCtlNowhere(NW_FNAM, lngE) = ![frm_name]
            arr_varCtlNowhere(NW_CNAM, lngE) = ![ctl_name]
          End If
        Case Else
          ' ** Skip.
        End Select
        If lngX < lngRecs Then .MoveNext
      Next  ' ** lngX.

      If lngCtlsNowhere > 0& Then
        With rst
          For lngX = 0& To (lngCtlsNowhere - 1&)
            .MoveFirst
            .FindFirst "[ctltmp_id] = " & CStr(arr_varCtlNowhere(NW_CTID, lngX))
            If .NoMatch = False Then
              blnDelete = True
              Debug.Print "'DELETE?  " & arr_varCtlNowhere(NW_CNAM, lngX) & "  ON  " & arr_varCtlNowhere(NW_FNAM, lngX)
              Stop
              If blnDelete = True Then
                .Delete
              End If
            Else
              Stop
            End If
          Next  ' ** lngX.
        End With
      End If  ' ** lngCtlsNowhere.

    End With  ' ** rst.

    rst.Close
    Set rst = Nothing

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Debug.Print "'DONE!"

'TOT CTLS: 822

'frmMenu_Main:             V2.2.30 CTLS: 81   V2.2.20 CTLS: 81
'frmMenu_Account:          V2.2.30 CTLS: 135  V2.2.20 CTLS: 135
'frmMenu_Post:             V2.2.30 CTLS: 102  V2.2.20 CTLS: 94
'frmMenu_Report:           V2.2.30 CTLS: 96   V2.2.20 CTLS: 90
'frmMenu_Asset:            V2.2.30 CTLS: 61   V2.2.20 CTLS: 61
'frmMenu_Utility:          V2.2.30 CTLS: 88   V2.2.20 CTLS: 86
'frmMenu_CourtReport:      V2.2.30 CTLS: 55   V2.2.20 CTLS: 55
'frmMenu_Maintenance:      V2.2.30 CTLS: 74   V2.2.20 CTLS: 61
'frmMenu_Other:            V2.2.30 CTLS: 67   V2.2.20 CTLS: 67

'CTLS NOT FOUND: 29
'frmMenu_Post:             chkDataCheck
'frmMenu_Post:             chkDataCheck_lbl
'frmMenu_Post:             cmdDataCheck
'frmMenu_Post:             cmdDataCheck_lbl
'frmMenu_Post:             cmdPost_l_box01_hline05
'frmMenu_Post:             cmdPost_l_box01_hline06
'frmMenu_Post:             JrnlCols
'frmMenu_Post:             JrnlPref

'frmMenu_Report:           cmdForm1099
'frmMenu_Report:           cmdSearch_arw
'frmMenu_Report:           cmdSearch_lbl1
'frmMenu_Report:           cmdSearch_lbl2
'frmMenu_Report:           GoToReport_arw_menu_r_crtrpt_img
'frmMenu_Report:           GoToReport_arw_menu_r_crtrpt_img_lbl

'frmMenu_Utility:          GoToReport_arw_menu_r_maint_img
'frmMenu_Utility:          GoToReport_arw_menu_r_maint_img_lbl

'frmMenu_Maintenance:      cmdMaint05_lbl
'frmMenu_Maintenance:      cmdMaint07
'frmMenu_Maintenance:      cmdMaint08
'frmMenu_Maintenance:      cmdMaint08_box01
'frmMenu_Maintenance:      cmdMaint08_box02
'frmMenu_Maintenance:      GoToReport_arw_menu_admin_img
'frmMenu_Maintenance:      GoToReport_arw_menu_admin_img_lbl
'frmMenu_Maintenance:      GoToReport_arw_menu_forex_img
'frmMenu_Maintenance:      GoToReport_arw_menu_forex_img_lbl
'frmMenu_Maintenance:      GoToReport_arw_menu_other_img
'frmMenu_Maintenance:      GoToReport_arw_menu_other_img_lbl
'frmMenu_Maintenance:      TPP
'frmMenu_Maintenance:      TPP_lbl
'DONE!

'#################

'TOT CTLS: 552

'frmJournal:                V2.2.30 CTLS: 149  V2.2.20 CTLS: 120
'frmJournal_Sub1_Dividend:  V2.2.30 CTLS: 61   V2.2.20 CTLS: 60
'frmJournal_Sub2_Interest:  V2.2.30 CTLS: 59   V2.2.20 CTLS: 57
'frmJournal_Sub3_Purchase:  V2.2.30 CTLS: 84   V2.2.20 CTLS: 81
'frmJournal_Sub4_Sold:      V2.2.30 CTLS: 102  V2.2.20 CTLS: 91
'frmJournal_Sub5_Misc:      V2.2.30 CTLS: 97   V2.2.20 CTLS: 87

'CTLS NOT FOUND: 61
'frmJournal:                cmdMiscMap_LTCL_raised_focus_dots_img
'frmJournal:                cmdMiscMap_LTCL_raised_focus_img
'frmJournal:                cmdMiscMap_LTCL_raised_img
'frmJournal:                cmdMiscMap_LTCL_raised_img_dis
'frmJournal:                cmdMiscMap_LTCL_raised_semifocus_dots_img
'frmJournal:                cmdMiscMap_LTCL_sunken_focus_dots_img
'frmJournal:                cmdMiscMap_STCGL_raised_focus_dots_img
'frmJournal:                cmdMiscMap_STCGL_raised_focus_img
'frmJournal:                cmdMiscMap_STCGL_raised_img
'frmJournal:                cmdMiscMap_STCGL_raised_img_dis
'frmJournal:                cmdMiscMap_STCGL_raised_semifocus_dots_img
'frmJournal:                cmdMiscMap_STCGL_sunken_focus_dots_img
'frmJournal:                DefAccountNo
'frmJournal:                DefCash
'frmJournal:                DefICash
'frmJournal:                DefPCash
'frmJournal:                DefPostingDate
'frmJournal:                DefReinvestBtn
'frmJournal:                DefShortname
'frmJournal:                DefTransdate
'frmJournal:                FromMisc
'frmJournal:                FromPurchase
'frmJournal:                GoToReport_arw_mapltcl_img
'frmJournal:                GoToReport_arw_mapstcgl_img
'frmJournal:                tglDividendReinvest_true_img_dis
'frmJournal:                tglInterestReinvest_true_img_dis
'frmJournal:                tglMiscReinvest_true_img_dis
'frmJournal:                tglPurchaseSale_true_img_dis
'frmJournal:                tglSaleReinvest_true_img_dis

'frmJournal_Sub1_Dividend:  tglDividendReinvest
'frmJournal_Sub1_Dividend:  tglDividendReinvest_true_img_dis

'frmJournal_Sub2_Interest:  Location_ID
'frmJournal_Sub2_Interest:  tglInterestReinvest
'frmJournal_Sub2_Interest:  tglInterestReinvest_true_img_dis

'frmJournal_Sub3_Purchase:  cmbLocations
'frmJournal_Sub3_Purchase:  cmbLocations_lbl
'frmJournal_Sub3_Purchase:  posted
'frmJournal_Sub3_Purchase:  posted_lbl
'frmJournal_Sub3_Purchase:  tglPurchaseSale
'frmJournal_Sub3_Purchase:  tglPurchaseSale_true_img_dis

'frmJournal_Sub4_Sold:      CurrentAssetBalance
'frmJournal_Sub4_Sold:      CurrentAssetBalance_box
'frmJournal_Sub4_Sold:      CurrentAssetBalance_chk
'frmJournal_Sub4_Sold:      CurrentAssetBalance_lbl
'frmJournal_Sub4_Sold:      CurrentAssetCash
'frmJournal_Sub4_Sold:      CurrentAssetCash_box
'frmJournal_Sub4_Sold:      CurrentAssetCash_chk
'frmJournal_Sub4_Sold:      CurrentAssetCash_lbl
'frmJournal_Sub4_Sold:      saleReinvested
'frmJournal_Sub4_Sold:      tglSaleReinvest
'frmJournal_Sub4_Sold:      tglSaleReinvest_true_img_dis

'frmJournal_Sub5_Misc:      CurrentCashBalance
'frmJournal_Sub5_Misc:      CurrentCashBalance_box
'frmJournal_Sub5_Misc:      CurrentCashBalance_chk
'frmJournal_Sub5_Misc:      CurrentCashBalance_lbl
'frmJournal_Sub5_Misc:      GoToReport_arw_mapltcl_img
'frmJournal_Sub5_Misc:      GoToReport_arw_mapltcl_img_lbl
'frmJournal_Sub5_Misc:      SpecialCapGainLoss
'frmJournal_Sub5_Misc:      SpecialCapGainLoss_lbl
'frmJournal_Sub5_Misc:      tglMiscReinvest
'frmJournal_Sub5_Misc:      tglMiscReinvest_true_img_dis

'CTLS HERE ONLY: 5
'DONE!

'#################

'TOT CTLS: 656

'frmOptions:                V2.2.30 CTLS: 104  V2.2.20 CTLS: 104
'DONE!

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NewJrnlMap_Chk = blnRetVal

End Function

Public Function NewJrnlMap_Set0_Journal() As Boolean

  Const THIS_PROC As String = "NewJrnlMap_Set0_Journal"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim strFormName As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCtl().
  Const C_CTID As Integer = 0
  Const C_FNAM As Integer = 1
  Const C_CNAM As Integer = 2
  Const C_CTYP As Integer = 3
  Const C_TOP  As Integer = 4
  Const C_LFT  As Integer = 5
  Const C_WDT  As Integer = 6
  Const C_HGT  As Integer = 7
  Const C_FND  As Integer = 8
  Const C_HERE As Integer = 9
  Const C_DATM As Integer = 10

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_Form_Control_03, just 'frmJournal', ctl_found = True.
    Set qdf = .QueryDefs("zzz_qry_NewJrnlMap_02_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' ********************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     ctltmp_id              C_CTID
      ' **     2       1     frm_name               C_FNAM
      ' **     3       2     ctl_name               C_CNAM
      ' **     4       3     ctltype_type           C_CTYP
      ' **     5       4     ctl_top                C_TOP
      ' **     6       5     ctl_left               C_LFT
      ' **     7       6     ctl_width              C_WDT
      ' **     8       7     ctl_height             C_HGT
      ' **     9       8     ctl_found              C_FND
      ' **    10       9     ctl_hereonly           C_HERE
      ' **    11      10     ctltmp_datemodified    C_DATM
      ' **
      ' ********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  strFormName = "frmJournal"
  Set frm = Forms(strFormName)
  With frm
    For lngX = 0& To (lngCtls - 1&)
      Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
      With ctl
        .Top = arr_varCtl(C_TOP, lngX)
        .Left = arr_varCtl(C_LFT, lngX)
        .Width = arr_varCtl(C_WDT, lngX)
        .Height = arr_varCtl(C_HGT, lngX)
      End With
    Next
  End With

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NewJrnlMap_Set0_Journal = blnRetVal

End Function

Public Function NewJrnlMap_Set1_Dividend() As Boolean

  Const THIS_PROC As String = "NewJrnlMap_Set1_Dividend"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim strFormName As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCtl().
  Const C_CTID As Integer = 0
  Const C_FNAM As Integer = 1
  Const C_CNAM As Integer = 2
  Const C_CTYP As Integer = 3
  Const C_TOP  As Integer = 4
  Const C_LFT  As Integer = 5
  Const C_WDT  As Integer = 6
  Const C_HGT  As Integer = 7
  Const C_FND  As Integer = 8
  Const C_HERE As Integer = 9
  Const C_DATM As Integer = 10

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_Form_Control_03, just 'frmJournal_Sub1_Dividend', ctl_found = True.
    Set qdf = .QueryDefs("zzz_qry_NewJrnlMap_03_01_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' ********************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     ctltmp_id              C_CTID
      ' **     2       1     frm_name               C_FNAM
      ' **     3       2     ctl_name               C_CNAM
      ' **     4       3     ctltype_type           C_CTYP
      ' **     5       4     ctl_top                C_TOP
      ' **     6       5     ctl_left               C_LFT
      ' **     7       6     ctl_width              C_WDT
      ' **     8       7     ctl_height             C_HGT
      ' **     9       8     ctl_found              C_FND
      ' **    10       9     ctl_hereonly           C_HERE
      ' **    11      10     ctltmp_datemodified    C_DATM
      ' **
      ' ********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  strFormName = "frmJournal_Sub1_Dividend"
  Set frm = Forms(strFormName)
  With frm
    For lngX = 0& To (lngCtls - 1&)
      Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
      With ctl
        .Top = arr_varCtl(C_TOP, lngX)
        .Left = arr_varCtl(C_LFT, lngX)
        .Width = arr_varCtl(C_WDT, lngX)
        .Height = arr_varCtl(C_HGT, lngX)
      End With
    Next
  End With

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NewJrnlMap_Set1_Dividend = blnRetVal

End Function

Public Function NewJrnlMap_Set2_Interest() As Boolean

  Const THIS_PROC As String = "NewJrnlMap_Set2_Interest"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim strFormName As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCtl().
  Const C_CTID As Integer = 0
  Const C_FNAM As Integer = 1
  Const C_CNAM As Integer = 2
  Const C_CTYP As Integer = 3
  Const C_TOP  As Integer = 4
  Const C_LFT  As Integer = 5
  Const C_WDT  As Integer = 6
  Const C_HGT  As Integer = 7
  Const C_FND  As Integer = 8
  Const C_HERE As Integer = 9
  Const C_DATM As Integer = 10

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_Form_Control_03, just 'frmJournal_Sub2_Interest', ctl_found = True.
    Set qdf = .QueryDefs("zzz_qry_NewJrnlMap_03_02_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' ********************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     ctltmp_id              C_CTID
      ' **     2       1     frm_name               C_FNAM
      ' **     3       2     ctl_name               C_CNAM
      ' **     4       3     ctltype_type           C_CTYP
      ' **     5       4     ctl_top                C_TOP
      ' **     6       5     ctl_left               C_LFT
      ' **     7       6     ctl_width              C_WDT
      ' **     8       7     ctl_height             C_HGT
      ' **     9       8     ctl_found              C_FND
      ' **    10       9     ctl_hereonly           C_HERE
      ' **    11      10     ctltmp_datemodified    C_DATM
      ' **
      ' ********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  strFormName = "frmJournal_Sub2_Interest"
  Set frm = Forms(strFormName)
  With frm
    For lngX = 0& To (lngCtls - 1&)
      Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
      With ctl
        .Top = arr_varCtl(C_TOP, lngX)
        .Left = arr_varCtl(C_LFT, lngX)
        .Width = arr_varCtl(C_WDT, lngX)
        .Height = arr_varCtl(C_HGT, lngX)
      End With
    Next
  End With

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NewJrnlMap_Set2_Interest = blnRetVal

End Function

Public Function NewJrnlMap_Set3_Purchase() As Boolean

  Const THIS_PROC As String = "NewJrnlMap_Set3_Purchase"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim strFormName As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCtl().
  Const C_CTID As Integer = 0
  Const C_FNAM As Integer = 1
  Const C_CNAM As Integer = 2
  Const C_CTYP As Integer = 3
  Const C_TOP  As Integer = 4
  Const C_LFT  As Integer = 5
  Const C_WDT  As Integer = 6
  Const C_HGT  As Integer = 7
  Const C_FND  As Integer = 8
  Const C_HERE As Integer = 9
  Const C_DATM As Integer = 10

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_Form_Control_03, just 'frmJournal_Sub3_Purchase', ctl_found = True.
    Set qdf = .QueryDefs("zzz_qry_NewJrnlMap_03_03_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' ********************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     ctltmp_id              C_CTID
      ' **     2       1     frm_name               C_FNAM
      ' **     3       2     ctl_name               C_CNAM
      ' **     4       3     ctltype_type           C_CTYP
      ' **     5       4     ctl_top                C_TOP
      ' **     6       5     ctl_left               C_LFT
      ' **     7       6     ctl_width              C_WDT
      ' **     8       7     ctl_height             C_HGT
      ' **     9       8     ctl_found              C_FND
      ' **    10       9     ctl_hereonly           C_HERE
      ' **    11      10     ctltmp_datemodified    C_DATM
      ' **
      ' ********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  strFormName = "frmJournal_Sub3_Purchase"
  Set frm = Forms(strFormName)
  With frm
    For lngX = 0& To (lngCtls - 1&)
      Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
      With ctl
        .Top = arr_varCtl(C_TOP, lngX)
        .Left = arr_varCtl(C_LFT, lngX)
        .Width = arr_varCtl(C_WDT, lngX)
        .Height = arr_varCtl(C_HGT, lngX)
      End With
    Next
  End With

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NewJrnlMap_Set3_Purchase = blnRetVal

End Function

Public Function NewJrnlMap_Set4_Sold() As Boolean

  Const THIS_PROC As String = "NewJrnlMap_Set4_Sold"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim strFormName As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCtl().
  Const C_CTID As Integer = 0
  Const C_FNAM As Integer = 1
  Const C_CNAM As Integer = 2
  Const C_CTYP As Integer = 3
  Const C_TOP  As Integer = 4
  Const C_LFT  As Integer = 5
  Const C_WDT  As Integer = 6
  Const C_HGT  As Integer = 7
  Const C_FND  As Integer = 8
  Const C_HERE As Integer = 9
  Const C_DATM As Integer = 10

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_Form_Control_03, just 'frmJournal_Sub4_Sold', ctl_found = True.
    Set qdf = .QueryDefs("zzz_qry_NewJrnlMap_03_04_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' ********************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     ctltmp_id              C_CTID
      ' **     2       1     frm_name               C_FNAM
      ' **     3       2     ctl_name               C_CNAM
      ' **     4       3     ctltype_type           C_CTYP
      ' **     5       4     ctl_top                C_TOP
      ' **     6       5     ctl_left               C_LFT
      ' **     7       6     ctl_width              C_WDT
      ' **     8       7     ctl_height             C_HGT
      ' **     9       8     ctl_found              C_FND
      ' **    10       9     ctl_hereonly           C_HERE
      ' **    11      10     ctltmp_datemodified    C_DATM
      ' **
      ' ********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  strFormName = "frmJournal_Sub4_Sold"
  Set frm = Forms(strFormName)
  With frm
    For lngX = 0& To (lngCtls - 1&)
      Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
      With ctl
        .Top = arr_varCtl(C_TOP, lngX)
        .Left = arr_varCtl(C_LFT, lngX)
        .Width = arr_varCtl(C_WDT, lngX)
        .Height = arr_varCtl(C_HGT, lngX)
      End With
    Next
  End With

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NewJrnlMap_Set4_Sold = blnRetVal

End Function

Public Function NewJrnlMap_Set5_Misc() As Boolean

  Const THIS_PROC As String = "NewJrnlMap_Set5_Misc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim strFormName As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCtl().
  Const C_CTID As Integer = 0
  Const C_FNAM As Integer = 1
  Const C_CNAM As Integer = 2
  Const C_CTYP As Integer = 3
  Const C_TOP  As Integer = 4
  Const C_LFT  As Integer = 5
  Const C_WDT  As Integer = 6
  Const C_HGT  As Integer = 7
  Const C_FND  As Integer = 8
  Const C_HERE As Integer = 9
  Const C_DATM As Integer = 10

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_Form_Control_03, just 'frmJournal_Sub5_Misc', ctl_found = True.
    Set qdf = .QueryDefs("zzz_qry_NewJrnlMap_03_05_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' ********************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     ctltmp_id              C_CTID
      ' **     2       1     frm_name               C_FNAM
      ' **     3       2     ctl_name               C_CNAM
      ' **     4       3     ctltype_type           C_CTYP
      ' **     5       4     ctl_top                C_TOP
      ' **     6       5     ctl_left               C_LFT
      ' **     7       6     ctl_width              C_WDT
      ' **     8       7     ctl_height             C_HGT
      ' **     9       8     ctl_found              C_FND
      ' **    10       9     ctl_hereonly           C_HERE
      ' **    11      10     ctltmp_datemodified    C_DATM
      ' **
      ' ********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  strFormName = "frmJournal_Sub5_Misc"
  Set frm = Forms(strFormName)
  With frm
    For lngX = 0& To (lngCtls - 1&)
      Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
      With ctl
        .Top = arr_varCtl(C_TOP, lngX)
        .Left = arr_varCtl(C_LFT, lngX)
        .Width = arr_varCtl(C_WDT, lngX)
        .Height = arr_varCtl(C_HGT, lngX)
      End With
    Next
  End With

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NewJrnlMap_Set5_Misc = blnRetVal

End Function

Public Function NewJrnlMap_Set_Options() As Boolean

  Const THIS_PROC As String = "NewJrnlMap_Set_Options"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim strFormName As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCtl().
  Const C_CTID As Integer = 0
  Const C_FNAM As Integer = 1
  Const C_CNAM As Integer = 2
  Const C_CTYP As Integer = 3
  Const C_TOP  As Integer = 4
  Const C_LFT  As Integer = 5
  Const C_WDT  As Integer = 6
  Const C_HGT  As Integer = 7
  Const C_CAP  As Integer = 8
  Const C_FCLR As Integer = 9
  Const C_BCLR As Integer = 10
  Const C_FONT As Integer = 11
  Const C_SIZE As Integer = 12
  Const C_BOLD As Integer = 13
  Const C_SPEC As Integer = 14
  Const C_BSTL As Integer = 15
  Const C_RCLR As Integer = 16
  Const C_RSTL As Integer = 17
  Const C_SLNT As Integer = 18
  Const C_VIS  As Integer = 19
  Const C_ABLE As Integer = 20
  Const C_LOCK As Integer = 21
  Const C_DEF  As Integer = 22
  Const C_OPTV As Integer = 23
  Const C_TAB  As Integer = 24
  Const C_MODE As Integer = 25
  Const C_AUTO As Integer = 26
  Const C_UPD  As Integer = 27
  'Const C_RWDT As Integer = 28
  Const C_FND  As Integer = 28
  Const C_HERE As Integer = 29
  Const C_DATM As Integer = 30

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_Form_Control_03, just 'frmOptions', ctl_found = True.
    Set qdf = .QueryDefs("zzz_qry_NewJrnlMap_06_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' ********************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     ctltmp_id              C_CTID
      ' **     2       1     frm_name               C_FNAM
      ' **     3       2     ctl_name               C_CNAM
      ' **     4       3     ctltype_type           C_CTYP
      ' **     5       4     ctl_top                C_TOP
      ' **     6       5     ctl_left               C_LFT
      ' **     7       6     ctl_width              C_WDT
      ' **     8       7     ctl_height             C_HGT
      ' **     9       8     ctl_caption            C_CAP
      ' **    10       9     ctl_forecolor          C_FCLR
      ' **    11      10     ctl_backcolor          C_BCLR
      ' **    12      11     ctl_fontname           C_FONT
      ' **    13      12     ctl_fontsize           C_SIZE
      ' **    14      13     ctl_fontbold           C_BOLD
      ' **    15      14     ctl_specialeffect      C_SPEC
      ' **    16      15     ctl_backstyle          C_BSTL
      ' **    17      16     ctl_bordercolor        C_RCLR
      ' **    18      17     ctl_borderstyle        C_RSTL
      ' **    19      18     ctl_lineslant          C_SLNT
      ' **    20      19     ctl_visible            C_VIS
      ' **    21      20     ctl_enabled            C_ABLE
      ' **    22      21     ctl_locked             C_LOCK
      ' **    23      22     ctl_defaultvalue       C_DEF
      ' **    24      23     ctl_optionvalue        C_OPTV
      ' **    25      24     ctl_tabstop            C_TAB
      ' **    26      25     ctl_sizemode           C_MODE
      ' **    27      26     ctl_autoactivate       C_AUTO
      ' **    28      27     ctl_updateoptions      C_UPD
      ' **    29      28     ctl_found              C_FND
      ' **    30      29     ctl_hereonly           C_HERE
      ' **    31      30     ctltmp_datemodified    C_DATM
      ' **
      ' ********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  strFormName = "frmOptions"
  If IsLoaded(strFormName, acForm, True) = False Then  ' ** Module Function: modFileUtilities.
    DoCmd.OpenForm strFormName, acDesign
    DoEvents
  End If

  Set frm = Forms(strFormName)
  With frm
    For lngX = 0& To (lngCtls - 1&)
      Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
      With ctl
        .Top = arr_varCtl(C_TOP, lngX)
        .Left = arr_varCtl(C_LFT, lngX)
        .Width = arr_varCtl(C_WDT, lngX)
        .Height = arr_varCtl(C_HGT, lngX)
        Select Case .ControlType
        Case acLabel
          .Caption = arr_varCtl(C_CAP, lngX)
          .ForeColor = arr_varCtl(C_FCLR, lngX)
          .BackColor = arr_varCtl(C_BCLR, lngX)
          .FontName = arr_varCtl(C_FONT, lngX)
          .FontSize = arr_varCtl(C_SIZE, lngX)
          .FontBold = arr_varCtl(C_BOLD, lngX)
          .SpecialEffect = arr_varCtl(C_SPEC, lngX)
          .BackStyle = arr_varCtl(C_BSTL, lngX)
          .BorderColor = arr_varCtl(C_RCLR, lngX)
          .BorderStyle = arr_varCtl(C_RSTL, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
        Case acCommandButton
          .Caption = arr_varCtl(C_CAP, lngX)
          .ForeColor = arr_varCtl(C_FCLR, lngX)
          .FontName = arr_varCtl(C_FONT, lngX)
          .FontSize = arr_varCtl(C_SIZE, lngX)
          .FontBold = arr_varCtl(C_BOLD, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
          .Enabled = arr_varCtl(C_ABLE, lngX)
          .TabStop = arr_varCtl(C_TAB, lngX)
        Case acTextBox
          .ForeColor = arr_varCtl(C_FCLR, lngX)
          .BackColor = arr_varCtl(C_BCLR, lngX)
          .FontName = arr_varCtl(C_FONT, lngX)
          .FontSize = arr_varCtl(C_SIZE, lngX)
          .FontBold = arr_varCtl(C_BOLD, lngX)
          .SpecialEffect = arr_varCtl(C_SPEC, lngX)
          .BackStyle = arr_varCtl(C_BSTL, lngX)
          .BorderColor = arr_varCtl(C_RCLR, lngX)
          .BorderStyle = arr_varCtl(C_RSTL, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
          .Enabled = arr_varCtl(C_ABLE, lngX)
          .Locked = arr_varCtl(C_LOCK, lngX)
          If IsNull(arr_varCtl(C_DEF, lngX)) = False Then
            .DefaultValue = arr_varCtl(C_DEF, lngX)
          End If
          .TabStop = arr_varCtl(C_TAB, lngX)
        Case acComboBox
          .ForeColor = arr_varCtl(C_FCLR, lngX)
          .BackColor = arr_varCtl(C_BCLR, lngX)
          .FontName = arr_varCtl(C_FONT, lngX)
          .FontSize = arr_varCtl(C_SIZE, lngX)
          .FontBold = arr_varCtl(C_BOLD, lngX)
          .SpecialEffect = arr_varCtl(C_SPEC, lngX)
          .BackStyle = arr_varCtl(C_BSTL, lngX)
          .BorderColor = arr_varCtl(C_RCLR, lngX)
          .BorderStyle = arr_varCtl(C_RSTL, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
          .Enabled = arr_varCtl(C_ABLE, lngX)
          .Locked = arr_varCtl(C_LOCK, lngX)
          If IsNull(arr_varCtl(C_DEF, lngX)) = False Then
            .DefaultValue = arr_varCtl(C_DEF, lngX)
          End If
          .TabStop = arr_varCtl(C_TAB, lngX)
        Case acCheckBox
          .SpecialEffect = arr_varCtl(C_SPEC, lngX)
          .BorderColor = arr_varCtl(C_RCLR, lngX)
          .BorderStyle = arr_varCtl(C_RSTL, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
          .Enabled = arr_varCtl(C_ABLE, lngX)
          .Locked = arr_varCtl(C_LOCK, lngX)
          If IsNull(arr_varCtl(C_DEF, lngX)) = False Then
            .DefaultValue = arr_varCtl(C_DEF, lngX)
          End If
          .TabStop = arr_varCtl(C_TAB, lngX)
        Case acOptionGroup
          .BackColor = arr_varCtl(C_BCLR, lngX)
          .SpecialEffect = arr_varCtl(C_SPEC, lngX)
          .BackStyle = arr_varCtl(C_BSTL, lngX)
          .BorderColor = arr_varCtl(C_RCLR, lngX)
          .BorderStyle = arr_varCtl(C_RSTL, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
          .Enabled = arr_varCtl(C_ABLE, lngX)
          .Locked = arr_varCtl(C_LOCK, lngX)
          If IsNull(arr_varCtl(C_DEF, lngX)) = False Then
            .DefaultValue = arr_varCtl(C_DEF, lngX)
          End If
          .TabStop = arr_varCtl(C_TAB, lngX)
        Case acOptionButton
          .SpecialEffect = arr_varCtl(C_SPEC, lngX)
          .BorderColor = arr_varCtl(C_RCLR, lngX)
          .BorderStyle = arr_varCtl(C_RSTL, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
          .Enabled = arr_varCtl(C_ABLE, lngX)
          .Locked = arr_varCtl(C_LOCK, lngX)
          .OptionValue = arr_varCtl(C_OPTV, lngX)
        Case acBoundObjectFrame
          .BackColor = arr_varCtl(C_BCLR, lngX)
          .SpecialEffect = arr_varCtl(C_SPEC, lngX)
          .BackStyle = arr_varCtl(C_BSTL, lngX)
          .BorderColor = arr_varCtl(C_RCLR, lngX)
          .BorderStyle = arr_varCtl(C_RSTL, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
          .Enabled = arr_varCtl(C_ABLE, lngX)
          .Locked = arr_varCtl(C_LOCK, lngX)
          .TabStop = arr_varCtl(C_TAB, lngX)
          .SizeMode = arr_varCtl(C_MODE, lngX)
          .AutoActivate = arr_varCtl(C_AUTO, lngX)
          .UpdateOptions = arr_varCtl(C_UPD, lngX)
        Case acLine
          .SpecialEffect = arr_varCtl(C_SPEC, lngX)
          .BorderColor = arr_varCtl(C_RCLR, lngX)
          .BorderStyle = arr_varCtl(C_RSTL, lngX)
          .LineSlant = arr_varCtl(C_SLNT, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
          '.BorderWidth = arr_varCtl(C_RWDT, lngX)
        Case acRectangle
          .BackColor = arr_varCtl(C_BCLR, lngX)
          .SpecialEffect = arr_varCtl(C_SPEC, lngX)
          .BackStyle = arr_varCtl(C_BSTL, lngX)
          .BorderColor = arr_varCtl(C_RCLR, lngX)
          .BorderStyle = arr_varCtl(C_RSTL, lngX)
          .Visible = arr_varCtl(C_VIS, lngX)
          '.BorderWidth = arr_varCtl(C_RWDT, lngX)
        End Select
      End With
    Next
  End With

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NewJrnlMap_Set_Options = blnRetVal

End Function

Public Function MenuColors() As Boolean

  Const THIS_PROC As String = "MenuColors"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
  Dim lngHues As Long, arr_varHue() As Variant
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim strFormName As String, strQueryName As String
  Dim lngThisDbsID As Long, lngFrmID As Long, lngRecs As Long
  Dim lngElem As Long
  Dim intPos1 As Integer
  Dim blnAdd As Boolean, blnChange As Boolean, blnSkip As Boolean
  Dim strTmp01 As String, strTmp02 As String
  Dim lngW As Long, lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  Const QRY_BASE As String = "zzz_qry_xMenuColors_05_"
  Const QRY_SFX As String = "_05"

  ' ** AcLineSlant enumeration:
  Const acLineSlantUpperLeftLowerRight As Integer = 0
  Const acLineSlantUpperRightLowerLeft As Integer = -1

  ' ** Array: arr_varHue().
  Const H_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const H_NAME As Integer = 0
  Const H_DID  As Integer = 1
  Const H_FID  As Integer = 2
  Const H_FNAM As Integer = 3
  Const H_BOX  As Integer = 4
  Const H_LIN1 As Integer = 5
  Const H_LIN2 As Integer = 6

  ' ** Array: arr_varCtl().
  Const C_CTID As Integer = 0
  Const C_DID  As Integer = 1
  Const C_FID  As Integer = 2
  Const C_FNAM As Integer = 3
  Const C_CID  As Integer = 4
  Const C_CNAM As Integer = 5
  Const C_CTYP As Integer = 6
  Const C_TOP  As Integer = 7
  Const C_LFT  As Integer = 8
  Const C_WDT  As Integer = 9
  Const C_HGT  As Integer = 10
  Const C_SPEC As Integer = 11
  Const C_BCLR As Integer = 12
  Const C_BSTL As Integer = 13
  Const C_RCLR As Integer = 14
  Const C_RSTL As Integer = 15
  Const C_RWDT As Integer = 16
  Const C_SLNT As Integer = 17
  Const C_VIS  As Integer = 18
  Const C_CHG  As Integer = 19

  blnRetVal = True

  If Forms.Count > 0 Then
    Do While Forms.Count > 0
      DoCmd.Close acForm, Forms(0).Name
      DoEvents
    Loop
  End If

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  blnSkip = True
  If blnSkip = False Then

    lngHues = 0&
    ReDim arr_varHue(H_ELEMS, 0)

    strFormName = "frmReportList"
    DoCmd.OpenForm strFormName, acDesign, , , , acHidden
    DoEvents
    Set frm = Forms(0)
    With frm
      For Each ctl In .FormHeader.Controls
        With ctl
          If .ControlType = acRectangle Then
            If Left(.Name, 5) = "Menu_" Then
              strTmp01 = Mid(.Name, 6)
              intPos1 = InStr(strTmp01, "_")
              If intPos1 > 0 Then
                strTmp02 = Mid(strTmp01, intPos1)
                strTmp01 = Left(strTmp01, (intPos1 - 1))
                Select Case strTmp01
                Case "Title", "Account", "Post", "Report", "Asset", "Utility", "Court", _
                    "Maint", "Other", "ForEx", "Admin", "Extra1", "Extra2"
                  lngHues = lngHues + 1&
                  lngE = lngHues - 1&
                  ReDim Preserve arr_varHue(H_ELEMS, lngE)
                  arr_varHue(H_NAME, lngE) = strTmp01
                  arr_varHue(H_DID, lngE) = lngThisDbsID
                  arr_varHue(H_FID, lngE) = CLng(0)
                  Select Case strTmp01
                  Case "Title"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Title"
                  Case "Account"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Account"
                  Case "Post"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Post"
                  Case "Report"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Report"
                  Case "Asset"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Asset"
                  Case "Utility"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Utility"
                  Case "Court"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_CourtReport"
                  Case "Maint"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Maintenance"
                  Case "Other"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Other"
                  Case "ForEx"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_ForeignExchange"
                  Case "Admin"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Admin"
                  Case "Extra1"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Extra1"
                  Case "Extra2"
                    arr_varHue(H_FNAM, lngE) = "frmMenu_Extra2"
                  End Select
                  Select Case strTmp02
                  Case "_box"
                    arr_varHue(H_BOX, lngE) = .BackColor
                    arr_varHue(H_LIN1, lngE) = Null
                    arr_varHue(H_LIN2, lngE) = Null
                  Case "_line1"
                    arr_varHue(H_BOX, lngE) = Null
                    arr_varHue(H_LIN1, lngE) = .BackColor
                    arr_varHue(H_LIN2, lngE) = Null
                  Case "_line2"
                    arr_varHue(H_BOX, lngE) = Null
                    arr_varHue(H_LIN1, lngE) = Null
                    arr_varHue(H_LIN2, lngE) = .BackColor
                  End Select
                End Select
              End If
            End If
          End If
        End With  ' ** ctl.
      Next  ' ** ctl.
    End With  ' ** frm.
    Set frm = Nothing
    DoCmd.Close acForm, strFormName, acSaveNo
    DoEvents

    If lngHues > 0& Then

      Set dbs = CurrentDb
      With dbs

        Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
        With rst
          .MoveFirst
          strFormName = vbNullString
          For lngX = 0& To (lngHues - 1&)
            If arr_varHue(H_FNAM, lngX) <> strFormName Then
              lngFrmID = 0&
              .FindFirst "[dbs_id] = " & CStr(arr_varHue(H_DID, lngX)) & " And [frm_name] = '" & arr_varHue(H_FNAM, lngX) & "'"
              Select Case .NoMatch
              Case True
                Select Case arr_varHue(H_FNAM, lngX)
                Case "frmMenu_Admin"
                  lngFrmID = -1&
                  strFormName = arr_varHue(H_FNAM, lngX)
                Case "frmMenu_ForeignExchange"
                  lngFrmID = -2&
                  strFormName = arr_varHue(H_FNAM, lngX)
                Case "frmMenu_Extra1"
                  lngFrmID = -3&
                  strFormName = arr_varHue(H_FNAM, lngX)
                Case "frmMenu_Extra2"
                  lngFrmID = -4&
                  strFormName = arr_varHue(H_FNAM, lngX)
                End Select
              Case False
                lngFrmID = ![frm_id]
                strFormName = ![frm_name]
              End Select
            End If
            arr_varHue(H_FID, lngX) = lngFrmID
          Next  ' ** lngX.
          .Close
        End With  ' ** rst.
        Set rst = Nothing

        'For lngX = 0& To (lngHues - 1&)
        '  If IsNull(arr_varHue(H_BOX, lngX)) = False Then
        '    Debug.Print "'" & arr_varHue(H_NAME, lngX) & "  " & CStr(arr_varHue(H_DID, lngX)) & "  " & _
        '      CStr(arr_varHue(H_FID, lngX)) & "  " & arr_varHue(H_FNAM, lngX) & "  " & _
        '      CStr(arr_varHue(H_BOX, lngX)) & "  {null}  {null}"
        '  ElseIf IsNull(arr_varHue(H_LIN1, lngX)) = False Then
        '    Debug.Print "'" & arr_varHue(H_NAME, lngX) & "  " & CStr(arr_varHue(H_DID, lngX)) & "  " & _
        '      CStr(arr_varHue(H_FID, lngX)) & "  " & arr_varHue(H_FNAM, lngX) & "  " & _
        '      "{null}  " & CStr(arr_varHue(H_LIN1, lngX)) & "  {null}"
        '  ElseIf IsNull(arr_varHue(H_LIN2, lngX)) = False Then
        '    Debug.Print "'" & arr_varHue(H_NAME, lngX) & "  " & CStr(arr_varHue(H_DID, lngX)) & "  " & _
        '      CStr(arr_varHue(H_FID, lngX)) & "  " & arr_varHue(H_FNAM, lngX) & "  " & _
        '      "{null}  {null}  " & CStr(arr_varHue(H_LIN2, lngX))
        '  End If
        'Next

        blnSkip = True
        If blnSkip = False Then

          Set rst = .OpenRecordset("tblForm_Menu_Color", dbOpenDynaset, dbConsistent)
          With rst
            For lngX = 0& To (lngHues - 1&)
              blnAdd = False
              If .BOF = True And .EOF = True Then
                blnAdd = True
              Else
                .FindFirst "[dbs_id] = " & CStr(arr_varHue(H_DID, lngX)) & " And [mnuclr_name] = '" & arr_varHue(H_NAME, lngX) & "'"
                If .NoMatch = True Then
                  blnAdd = True
                End If
              End If
              Select Case blnAdd
              Case True
                .AddNew
                ' ** ![mnuclr_id] : AutoNumber
                ![mnuclr_name] = arr_varHue(H_NAME, lngX)
                ![dbs_id] = arr_varHue(H_DID, lngX)
                ![frm_id] = arr_varHue(H_FID, lngX)
                ![frm_name] = arr_varHue(H_FNAM, lngX)
                Select Case IsNull(arr_varHue(H_BOX, lngX))
                Case True
                  ![mnuclr_box] = CLng(0)
                Case False
                  ![mnuclr_box] = arr_varHue(H_BOX, lngX)
                End Select
                Select Case IsNull(arr_varHue(H_LIN1, lngX))
                Case True
                  ![mnuclr_line1] = CLng(0)
                Case False
                  ![mnuclr_line1] = arr_varHue(H_LIN1, lngX)
                End Select
                Select Case IsNull(arr_varHue(H_LIN2, lngX))
                Case True
                  ![mnuclr_line2] = CLng(0)
                Case False
                  ![mnuclr_line2] = arr_varHue(H_LIN2, lngX)
                End Select
                ![mnuclr_datemodified] = Now()
                .Update
              Case False
                If IsNull(arr_varHue(H_BOX, lngX)) = False Then
                  .Edit
                  ![mnuclr_box] = arr_varHue(H_BOX, lngX)
                  ![mnuclr_datemodified] = Now()
                  .Update
                ElseIf IsNull(arr_varHue(H_LIN1, lngX)) = False Then
                  .Edit
                  ![mnuclr_line1] = arr_varHue(H_LIN1, lngX)
                  ![mnuclr_datemodified] = Now()
                  .Update
                ElseIf IsNull(arr_varHue(H_LIN2, lngX)) = False Then
                  .Edit
                  ![mnuclr_line2] = arr_varHue(H_LIN2, lngX)
                  ![mnuclr_datemodified] = Now()
                  .Update
                Else
                  Stop
                End If
              End Select
            Next  ' ** lngX.
            .Close
          End With  ' ** rst.
          Set rst = Nothing

        End If  ' ** blnSkip.

        .Close
      End With  ' ** dbs.
      Set dbs = Nothing

    Else
      Stop
    End If  ' ** lngHues.

  End If  ' ** blnSkip.

  blnSkip = False
  If blnSkip = False Then
    lngHues = 0&
    ReDim arr_varHue(H_ELEMS, 0)
    Set dbs = CurrentDb
    With dbs
      Set rst = .OpenRecordset("tblForm_Menu_Color", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          If ![mnuclr_name] <> "Admin" And ![mnuclr_name] <> "ForEx" And ![mnuclr_name] <> "Extra1" And ![mnuclr_name] <> "Extra2" Then
            lngHues = lngHues + 1&
            lngE = lngHues - 1&
            ReDim Preserve arr_varHue(H_ELEMS, lngE)
            arr_varHue(H_NAME, lngE) = ![mnuclr_name]
            arr_varHue(H_DID, lngE) = ![dbs_id]
            arr_varHue(H_FID, lngE) = ![frm_id]
            arr_varHue(H_FNAM, lngE) = ![frm_name]
            arr_varHue(H_BOX, lngE) = ![mnuclr_box]
            arr_varHue(H_LIN1, lngE) = ![mnuclr_line1]
            arr_varHue(H_LIN2, lngE) = ![mnuclr_line2]
          End If
          If lngX < lngRecs Then .MoveNext
        Next  ' ** lngX.
        .Close
      End With
      Set rst = Nothing
      .Close
    End With  ' ** dbs.
    Set dbs = Nothing
  End If  ' ** blnSkip.

  blnSkip = False
  If blnSkip = False Then

    Set dbs = CurrentDb
    With dbs

      blnChange = False

      For lngW = 1& To 9&

        strFormName = vbNullString

        Select Case lngW
        Case 1&
          strFormName = "frmMenu_Main"
        Case 2&
          strFormName = "frmMenu_Account"
        Case 3&
          strFormName = "frmMenu_Post"
        Case 4&
          strFormName = "frmMenu_Report"
        Case 5&
          strFormName = "frmMenu_Asset"
        Case 6&
          strFormName = "frmMenu_Utility"
        Case 7&
          strFormName = "frmMenu_CourtReport"
        Case 8&
          strFormName = "frmMenu_Maintenance"
        Case 9&
          strFormName = "frmMenu_Other"
        End Select

        ' ** zzz_qry_xMenuColors_05_01_05.
        strQueryName = QRY_BASE & Right("00" & CStr(lngW), 2) & QRY_SFX

        ' ** zz_tbl_Form_Control_03, just 'frmMenu_Main', ctl_found = True.
        Set qdf = .QueryDefs(strQueryName)
        Set rst = qdf.OpenRecordset
        With rst
          .MoveLast
          lngCtls = .RecordCount
          .MoveFirst
          arr_varCtl = .GetRows(lngCtls)
          ' ******************************************************
          ' ** Array: arr_varCtl()
          ' **
          ' **   Field  Element  Name                 Constant
          ' **   =====  =======  ===================  ==========
          ' **     1       0     ctltmp_id            C_CTID
          ' **     2       1     dbs_id               C_DID
          ' **     3       2     frm_id               C_FID
          ' **     4       3     frm_name             C_FNAM
          ' **     5       4     ctl_id               C_CID
          ' **     6       5     ctl_name             C_CNAM
          ' **     7       6     ctltype_type         C_CTYP
          ' **     8       7     ctl_top              C_TOP
          ' **     9       8     ctl_left             C_LFT
          ' **    10       9     ctl_width            C_WDT
          ' **    11      10     ctl_height           C_HGT
          ' **    12      11     ctl_specialeffect    C_SPEC
          ' **    13      12     ctl_backcolor        C_BCLR
          ' **    14      13     ctl_backstyle        C_BSTL
          ' **    15      14     ctl_bordercolor      C_RCLR
          ' **    16      15     ctl_borderstyle      C_RSTL
          ' **    17      16     ctl_borderwidth      C_RWDT
          ' **    18      17     ctl_lineslant        C_SLNT
          ' **    19      18     ctl_visible          C_VIS
          ' **    20      19     ctl_changed          C_CHG
          ' **
          ' ******************************************************
          .Close
        End With  ' ** rst.
        Set rst = Nothing
        Set qdf = Nothing

        DoCmd.OpenForm strFormName, acDesign, , , , acHidden
        Set frm = Forms(0)

        With frm
          For lngX = 0& To (lngCtls - 1&)
            Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
            With ctl
              Select Case arr_varCtl(C_CTYP, lngX)
              Case acRectangle
                If arr_varCtl(C_SPEC, lngX) = acSpecialEffectEtched Then
                  ' ** No etched boxes have identifying color, they're
                  ' ** all the main menu boxes with the beige color scheme.
                  ' **   arr_varCtl(C_BCLR, lngX) = MY_CLR_BGE Or arr_varCtl(C_BCLR, lngX) = MY_CLR_MDBGE Or
                  ' **   arr_varCtl(C_BCLR, lngX) = MY_CLR_LTBGE Or arr_varCtl(C_BCLR, lngX) = MY_CLR_VLTBGE
                Else
                  ' ** These are all buttons to the next menu.
                  If .SpecialEffect <> acSpecialEffectFlat Or arr_varCtl(C_SPEC, lngX) <> acSpecialEffectFlat Then
                    ' ** They should all be flat.
                    Stop
                    .SpecialEffect = acSpecialEffectFlat
                    arr_varCtl(C_SPEC, lngX) = acSpecialEffectFlat
                    blnChange = True
                    arr_varCtl(C_CHG, lngX) = CBool(True)
                  End If
                  If .BackStyle <> acBackStyleTransparent Or arr_varCtl(C_BSTL, lngX) <> acBackStyleTransparent Then
                    ' ** They should all be transparent.
                    Stop
                    .BackStyle = acBackStyleTransparent
                    arr_varCtl(C_BSTL, lngX) = acBackStyleTransparent
                    blnChange = True
                    arr_varCtl(C_CHG, lngX) = CBool(True)
                  End If
                  If .BorderStyle <> acBorderStyleSolid Or arr_varCtl(C_RSTL, lngX) <> acBorderStyleSolid Then
                    ' ** They should all be solid borders.
                    Stop
                    .BorderStyle = acBorderStyleSolid
                    arr_varCtl(C_RSTL, lngX) = acBorderStyleSolid
                    blnChange = True
                    arr_varCtl(C_CHG, lngX) = CBool(True)
                  End If
                  If .BorderWidth <> 0 Or arr_varCtl(C_RWDT, lngX) <> 0 Then
                    ' ** They should all be hairline width.
                    Stop
                    .BorderWidth = 0
                    arr_varCtl(C_RWDT, lngX) = 0
                    blnChange = True
                    arr_varCtl(C_CHG, lngX) = CBool(True)
                  End If
                  ' ** Identify the menu being called.
                  intPos1 = InStr(.Name, "_")
                  strTmp01 = Left(.Name, intPos1)
                  lngElem = -1&
                  For lngY = 0& To (lngHues - 1&)
                    Select Case strTmp01
                    Case "cmdAccount_"
                      If arr_varHue(H_NAME, lngY) = "Account" Then
                        lngElem = lngY
                        Exit For
                      End If
                    Case "cmdPost_"
                      If arr_varHue(H_NAME, lngY) = "Post" Then
                        lngElem = lngY
                        Exit For
                      End If
                    Case "cmdReport_"
                      If arr_varHue(H_NAME, lngY) = "Report" Then
                        lngElem = lngY
                        Exit For
                      End If
                    Case "cmdAsset_"
                      If arr_varHue(H_NAME, lngY) = "Asset" Then
                        lngElem = lngY
                        Exit For
                      End If
                    Case "cmdUtility_"
                      If arr_varHue(H_NAME, lngY) = "Utility" Then
                        lngElem = lngY
                        Exit For
                      End If
                    Case "cmdCourtReports_"
                      If arr_varHue(H_NAME, lngY) = "Court" Then
                        lngElem = lngY
                        Exit For
                      End If
                    Case "cmdMaintenance_"
                      If arr_varHue(H_NAME, lngY) = "Maint" Then
                        lngElem = lngY
                        Exit For
                      End If
                    Case "cmdMaint04_"
                      If InStr(frm.cmdMaint04.Caption, "&Other Tables") > 0 Then
                        If arr_varHue(H_NAME, lngY) = "Other" Then
                          lngElem = lngY
                          Exit For
                        End If
                      ElseIf InStr(frm.cmdMaint06.Caption, "Admini&stration") > 0 Then
                        If arr_varHue(H_NAME, lngY) = "Admin" Then
                          lngElem = lngY
                          Exit For
                        End If
                      Else
                        Stop
                      End If
                    Case "cmdMaint06_"
                      If InStr(frm.cmdMaint04.Caption, "&Other Tables") > 0 Then
                        If arr_varHue(H_NAME, lngY) = "Other" Then
                          lngElem = lngY
                          Exit For
                        End If
                      ElseIf InStr(frm.cmdMaint06.Caption, "Admini&stration") > 0 Then
                        If arr_varHue(H_NAME, lngY) = "Admin" Then
                          lngElem = lngY
                          Exit For
                        End If
                      Else
                        Stop
                      End If
                    End Select
                  Next  ' ** lngY.
                  ' ** lngElem now has element number of arr_varHues() array.
                  If Right(.Name, 6) = "_box01" Then
                    ' ** This is line 1.
                    .BorderColor = arr_varHue(H_LIN1, lngElem)
                  ElseIf Right(.Name, 6) = "_box02" Then
                    ' ** This is line 2.
                    .BorderColor = arr_varHue(H_LIN2, lngElem)
                  Else
                    Stop
                  End If
                End If
              Case acLine
                If .SpecialEffect <> acSpecialEffectFlat Or arr_varCtl(C_SPEC, lngX) <> acSpecialEffectFlat Then
                  ' ** They should all be flat.
                  Stop
                  .SpecialEffect = acSpecialEffectFlat
                  arr_varCtl(C_SPEC, lngX) = acSpecialEffectFlat
                  blnChange = True
                  arr_varCtl(C_CHG, lngX) = CBool(True)
                End If
                If .BorderStyle <> acBorderStyleSolid Or arr_varCtl(C_RSTL, lngX) <> acBorderStyleSolid Then
                  ' ** They should all be solid borders.
                  Stop
                  .BorderStyle = acBorderStyleSolid
                  arr_varCtl(C_RSTL, lngX) = acBorderStyleSolid
                  blnChange = True
                  arr_varCtl(C_CHG, lngX) = CBool(True)
                End If
                If .BorderWidth <> 0 Or arr_varCtl(C_RWDT, lngX) <> 0 Then
                  ' ** They should all be hairline width.
                  Stop
                  .BorderWidth = 0
                  arr_varCtl(C_RWDT, lngX) = 0
                  blnChange = True
                  arr_varCtl(C_CHG, lngX) = CBool(True)
                End If
                If .LineSlant <> acLineSlantUpperLeftLowerRight Or arr_varCtl(C_SLNT, lngX) <> acLineSlantUpperLeftLowerRight Then
                  ' ** They should all be backslash, '\'.
                  Stop
                  .LineSlant = acLineSlantUpperLeftLowerRight
                  arr_varCtl(C_SLNT, lngX) = acLineSlantUpperLeftLowerRight
                  blnChange = True
                  arr_varCtl(C_CHG, lngX) = CBool(True)
                End If
                If Left(.Name, 7) = "Detail_" Then
                  ' ** These are all etched line components.
                  ' ** They can be:
                  ' **   -2147483632  vbButtonShadow
                  ' **   -2147483628  vb3DHighlight
                  If (.BorderColor <> vbButtonShadow And .BorderColor <> vb3DHighlight) Or _
                      (arr_varCtl(C_RCLR, lngX) <> vbButtonShadow And arr_varCtl(C_RCLR, lngX) <> vb3DHighlight) Then
                    ' ** They shouldn't be Nav lines, some of which have other colors,
                    ' ** but if they're components of the file-folder construction,
                    ' ** the order of color for the vertical lines is different: 1, 2, 2, 1.
                    '.BorderColor = vbButtonShadow
                    'arr_varCtl(C_RCLR, lngX) = vbButtonShadow
                    '.BorderColor = vb3DHighlight
                    'arr_varCtl(C_RCLR, lngX) = vb3DHighlight
                    Stop
                  End If
                ElseIf InStr(.Name, "_box01_hline") > 0 Then
                  ' ** These are etched line components.
                  ' **   cmdPost_l_box01_hline01, cmdPost_r_box01_hline01, cmdAsset_l_box01_hline01.
                  ' ** They can be:
                  ' **   -2147483632  vbButtonShadow
                  ' **   -2147483628  vb3DHighlight
                  If (.BorderColor <> vbButtonShadow And .BorderColor <> vb3DHighlight) Or _
                      (arr_varCtl(C_RCLR, lngX) <> vbButtonShadow And arr_varCtl(C_RCLR, lngX) <> vb3DHighlight) Then
                    '.BorderColor = vbButtonShadow
                    'arr_varCtl(C_RCLR, lngX) = vbButtonShadow
                    '.BorderColor = vb3DHighlight
                    'arr_varCtl(C_RCLR, lngX) = vb3DHighlight
                    Stop
                  End If
                Else
                  lngElem = -1&
                  ' ** Identify the menu we're on.
                  For lngY = 0& To (lngHues - 1&)
                    If arr_varHue(H_FNAM, lngY) = strFormName Then
                      lngElem = lngY
                      Exit For
                    End If
                  Next  ' ** lngY.
                  strTmp01 = Right(.Name, 2)
                  If IsNumeric(strTmp01) = True Then
                    Select Case strTmp01
                    Case "01", "03"
                      .BorderColor = arr_varHue(H_LIN1, lngElem)
                    Case "02", "04"
                      .BorderColor = arr_varHue(H_LIN2, lngElem)
                    Case Else
                      Stop
                    End Select
                  Else
                    Stop
                  End If
                End If
              End Select
            End With
          Next  ' ** lngX.
        End With  ' ** frm.
        Set frm = Nothing

        DoCmd.Close acForm, strFormName, acSaveYes
        DoEvents

        If blnChange = True Then
          ' ** Update both zz_tbl_Form_Control_03 and tblForm_Control.
          Set dbs = CurrentDb
          With dbs
            Set rst = .OpenRecordset("zz_tbl_Form_Control_03", dbOpenDynaset, dbConsistent)
            With rst
              For lngY = 0& To (lngCtls - 1&)
                If arr_varCtl(C_CHG, lngY) = True Then
                  .FindFirst "[ctltmp_id] = " & CStr(arr_varCtl(C_CTID, lngY))
                  If .NoMatch = False Then
                    .Edit
                    ![ctl_specialeffect] = arr_varCtl(C_SPEC, lngY)
                    ![ctl_backcolor] = arr_varCtl(C_BCLR, lngY)
                    ![ctl_backstyle] = arr_varCtl(C_BSTL, lngY)
                    ![ctl_bordercolor] = arr_varCtl(C_RCLR, lngY)
                    ![ctl_borderstyle] = arr_varCtl(C_RSTL, lngY)
                    ![ctl_borderwidth] = arr_varCtl(C_RWDT, lngY)
                    ![ctl_lineslant] = arr_varCtl(C_SLNT, lngY)
                    ![ctltmp_datemodified] = Now()
                    .Update
                  Else
                    Stop
                  End If
                End If
              Next  ' ** lngY.
              .Close
            End With  ' ** rst.
            Set rst = Nothing
            Set rst = .OpenRecordset("tblForm_Control_Specification_A", dbOpenDynaset, dbConsistent)
            With rst
              For lngY = 0& To (lngCtls - 1&)
                If arr_varCtl(C_CHG, lngY) = True Then
                  .FindFirst "[ctl_id] = " & CStr(arr_varCtl(C_CID, lngY))
                  If .NoMatch = False Then
                    .Edit
                    ![ctlspec_backcolor] = arr_varCtl(C_BCLR, lngY)
                    ![ctlspec_backstyle] = arr_varCtl(C_BSTL, lngY)
                    ![ctlspec_bordercolor] = arr_varCtl(C_RCLR, lngY)
                    ![ctlspec_borderstyle] = arr_varCtl(C_RSTL, lngY)
                    ![ctlspec_borderwidth] = arr_varCtl(C_RWDT, lngY)
                    ![ctlspec_lineslant] = arr_varCtl(C_SLNT, lngY)
                    .Update
                  Else
                    Stop
                  End If
                End If
              Next  ' ** lngY.
            End With  ' ** rst.
            Set rst = Nothing
            Set rst = .OpenRecordset("tblForm_Control_Specification_B", dbOpenDynaset, dbConsistent)
            With rst
              For lngY = 0& To (lngCtls - 1&)
                If arr_varCtl(C_CHG, lngY) = True Then
                  .FindFirst "[ctl_id] = " & CStr(arr_varCtl(C_CID, lngY))
                  If .NoMatch = False Then
                    .Edit
                    ![ctlspec_specialeffect] = arr_varCtl(C_SPEC, lngY)
                    ![ctlspec_datemodified] = Now()
                    .Update
                  Else
                    Stop
                  End If
                End If
              Next  ' ** lngY.
            End With  ' ** rst.
            Set rst = Nothing
            .Close
          End With  ' ** dbs.
          Set dbs = Nothing
        End If

        Debug.Print "'FRM: " & strFormName
        DoEvents
        If blnChange = True Then
          Debug.Print "'  CHANGES!"
        End If

      Next  ' ** lngW.

      .Close
    End With  ' ** dbs.

  End If  ' ** blnSkip.

  Debug.Print "'DONE!"

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  MenuColors = blnRetVal

End Function
