Attribute VB_Name = "zz_mod_NickersonProbatePlusFuncs"
Option Compare Database
Option Explicit

'VGC 09/12/2015: CHANGES!

'{\rtf1\ansi\deff0\deftab720 {\fonttbl {\fonttbl {\f0\froman\fcharset0\fprq2 Times New Roman;} } }
'{\colortbl\red0\green0\blue0;}

'{\header
'  {\deftab720\margl720\margt2160\margr720\margb1440\plain\f0\fs30\b
'    \pard\qc Assets on Hand \par
'    \plain\f0\fs24\b Gowen, Edward H. Trust U/W \par
'    \plain\f0\fs24 Charles L. Nickerson, Trustee \par
'    \plain\f0\fs12 \par
'    \plain\f0\fs24\b
'    \trowd\trgaph70\trleft-70
'    \clbrdrb\brdrs\brdrw10\cellx8460\cellx8640\clbrdrb\brdrs\brdrw10\cellx10800\pard\widctlpar\intbl
'    \pard\intbl Description\cell
'    \plain\f0\fs24
'    \cell
'    \plain\f0\fs24\b
'    \pard\intbl\qc Carrying Value\cell
'    \intbl\row\pard
'  }\par
'}

'{\footer
'  {\deftab720\margl720\margt2160\margr720\margb1440\plain\f0\fs24
'    \trowd\trgaph70\trleft-70
'    \cellx4320\cellx6480\cellx10800\pard\widctlpar\intbl
'    \pard\intbl Fri Sep 04 11:50:31 2015\cell
'    \plain\f0\fs24\b
'    \pard\intbl\qc -
'      {\field
'        {\*\fldinst
'          {\cgrid0  PAGE
'          }
'        }
'      } -\cell
'    \plain\f0\fs24
'    \cell
'    \intbl\row\pard
'  }\par
'}

'{\header {\deftab720\margl720\margt2160\margr720\margb1440\plain\f0\fs30\b \pard\qc All Cash Receipts and Disbursements \par \plain\f0\fs24\b Gowen, Edward H. Trust U/W \par \plain\f0\fs24 Charles L. Nickerson, Trustee \par \plain\f0\fs12 \par \plain\f0\fs24\b \trowd\trgaph70\trleft-70 \clbrdrb\brdrs\brdrw10\cellx1260\cellx1440\clbrdrb\brdrs\brdrw10\cellx6120\cellx6300\clbrdrb\brdrs\brdrw10\cellx8460\cellx8640\clbrdrb\brdrs\brdrw10\cellx10800\pard\widctlpar\intbl \pard\intbl Date/ \par Reference\cell \plain\f0\fs24 \cell \plain\f0\fs24\b Description\cell \plain\f0\fs24 \cell \plain\f0\fs24\b \pard\intbl\qr Receipt \par Amount\cell \plain\f0\fs24 \cell \plain\f0\fs24\b Disbursement \par Amount\cell \intbl\row\pard }\par }

'{\footer {\deftab720\margl720\margt2160\margr720\margb1440\plain\f0\fs24 \trowd\trgaph70\trleft-70 \cellx4320\cellx6480\cellx10800\pard\widctlpar\intbl \pard\intbl Fri Sep 04 11:50:31 2015\cell \plain\f0\fs24\b \pard\intbl\qc - {\field{\*\fldinst {\cgrid0  PAGE }}} -\cell \plain\f0\fs24 \cell \intbl\row\pard }\par }

'{\deftab720\margl720\margt2160\margr720\margb1440\plain\f0\fs24\b \trowd\trgaph70\trleft-70 \cellx4320\pard\widctlpar\intbl \pard\intbl Bank Accounts\cell \intbl\row\pard \plain\f0\fs24 \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr \cell \cell \pard\intbl Bank Deposit Program-Smith \par Barney money funds \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 7,707.11\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr \cell \cell \pard\intbl Clearing Account \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 1,061.54\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr \cell \cell \pard\intbl Principal Cash \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab -407.91\cell \intbl\row\pard \trowd\trgaph70\trleft-70
'\cellx8460\cellx8640\clbrdrt\brdrs\brdrw10\cellx10800\pard\widctlpar\intbl \pard\intbl\qr Sub Total:\cell \cell \pard\intbl\tqr\tx1940\tx2880 $ \tab 8,360.74\cell \intbl\row\pard \plain\f0\fs24\b \trowd\trgaph70\trleft-70 \cellx4320\pard\widctlpar\intbl \pard\intbl Stocks and Funds\cell \intbl\row\pard \plain\f0\fs24\ul \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr No. of Shares\cell \plain\f0\fs24 \cell \cell \cell \cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 530.0000\cell \cell \pard\intbl AT & T Inc. [T] recd in merger \par w/ Bellsouth 01/03/07 CUSIP# \par 00206R-10-2 Bk/ Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 5,210.00\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 400.0000\cell \cell \pard\intbl Aetna Inc (New) (CT)
'\par [AET](recd as spinoff from \par Aetna, Inc.   12/19/00)100shs \par acq.6/3/96    Bk/Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 3,475.00\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 100.0000\cell \cell \pard\intbl Altria Group Inc.[MO] dtd \par 10/26/12 100shs @ $31.905/sh \par CUSIP# 02209S103 Bk/Entry: \par MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 3,190.50\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 800.0000\cell \cell \pard\intbl Automatic Data Processing, \par Inc. [ADP](acqd 7/18/90) Bk/ \par Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 4,732.76\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 164.0000\cell \cell
'\pard\intbl BP Amoco PLC Spons ADR \par [BP](recd 4/24/00 in merger w/ \par Atl Richfield Co.) Bk/Entry: \par MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880
'\tab 5,667.50\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 400.0000\cell \cell \pard\intbl Bristol Myers Squibb Co. [BMY] \par (acqd. 3/15/89) Bk/Entry: \par MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 4,435.28\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 200.0000\cell \cell \pard\intbl Broadridge Financial [BR] recd \par  in spinoff from Auto Data Pro \par cess. 4/5/07 CUSIP#11133T-10-3 \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 579.96\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 266.0000\cell \cell \pard\intbl CDK Global Inc [CDK] recd in \par spinoff from ADP 9/30/14 \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 683.05\cell \intbl\row\pard \trowd\trgaph70\trleft-70
'\cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 200.0000\cell \cell \pard\intbl Consolidated Edison Inc. [ED] \par pur 2/21/12 200shs @ \par $57.6263/sh CUSIP# 209115-10-4 \par Bk/Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 11,525.26\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 52.0000\cell \cell \pard\intbl Duke Energy Corp New [DUK] \par recd in rev. split w/ Duke \par Energy Hldg Corp. 7/5/12 \par Cusip# 26441C-20-4 Bk/Entry: \par MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 1,801.97\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 45.0000\cell \cell \pard\intbl Express Scripts Hldgc Co. \par [ESRX] recd in merger w/Medco \par Health Solutions, Inc. 4/03/12 \par 45shs CUSIP#30219G-10-8 \par Bk/Entry: MSSB \par \cell \cell
'\pard\intbl\tqr\tx1940\tx2880 \tab 409.17\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 528.0000\cell \cell \pard\intbl Exxon Mobil Corp.[XOM] common \par (recd in merger 12/1/99) \par Bk/Entry: MSSB (acq. \par 3/15/89) \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 5,010.00\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 960.0000\cell \cell \pard\intbl General Electric Co.[GE] \par common (acqd. 4/28/83) \par Bk/Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 1,937.30\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 50.0000\cell \cell \pard\intbl Hanes Brands, Inc. [HBI] recd \par in spinoff from Sara Lee Corp. \par 50shs CUSIP# \par Bk/Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 0.01\cell \intbl\row\pard
'\trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 162.0000\cell \cell \pard\intbl Hartford Financial Services \par Group, Inc. [HIG] common \par (spinoff fr ITT Corp) acqd \par 1/30/96 Bk Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 1,645.34\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 101.0000\cell \cell \pard\intbl Host Hotels & Resorts, Inc. \par [HST] recd in merger w/ \par Starwood Hotels & Resorts CL B \par dtd 4/11/2006 CUSIP# \par 44107P-10-4 Bk/Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 1,112.83\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 81.0000\cell \cell \pard\intbl ITT Industries, Inc.[ITT] \par common (spinoff fr ITT Corp. \par 1/30/96) Book Entry: MSSB \par \cell \cell
'\pard\intbl\tqr\tx1940\tx2880 \tab 146.43\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 600.0000\cell \cell \pard\intbl Intel Corp.[INTC] common, (pur \par 11/19/97 @ 79 1/4) Book Entry: \par MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 12,033.39\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 370.0000\cell \cell \pard\intbl JP Morgan Chase & Co. [JPM] \par (recd in merger w/JP Morgan & \par Co., Inc. 1/02/01)addl 270shs \par CUSIP # 46625H-10-0 Bk/Entry: \par MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 3,408.00\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 400.0000\cell \cell \pard\intbl McDonald's Corp.[MCD](acqd. \par 2/27/90) Book Entry: MSSB \par \cell \cell
'\pard\intbl\tqr\tx1940\tx2880 \tab 3,147.50\cell \intbl\row\pard \trowd\trgaph70\trleft-70
'\cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 100.0000\cell \cell \pard\intbl Plum Creek Timber Co.[PCL] \par recd in merger w/ Georgia \par Pacific     Corp. Timber \par 10/9/01 (ea.     share \par Geo.Pac.=1.37shs Plum   Creek) \par Bk/Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 1,283.93\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 78.0000\cell \cell \pard\intbl Spectra Energy Corp [SE] recd \par in spinoff from Duke Energy \par (Holding Co.) New 01/08/07 \par CUSIP# 847560-10-9 Bk/Entry: \par MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 1,299.00\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 124.0000\cell \cell \pard\intbl Starwood Hotels & \par Resorts-Worldwide, Inc.
'\par [HOT] (recd in merger-ITT \par Corp) acq.12/29/86 Bk/Entry: \par MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 826.96\cell
'\intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 5.0000\cell \cell \pard\intbl Vectrus Inc Com [VEC] recd in \par spin-off from 9/28/14 \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 6.59\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 300.0000\cell \cell \pard\intbl Verizon Communications [VZ] \par name change 7/7/00 fmly: Bell \par Atlantic Corp (rec'd in \par merger w/NYNEX 8/1/97) [acg. \par 8/1/96] Book Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 9,802.95\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 300.0000\cell \cell \pard\intbl Xerox Corp.[XRX] common (acqd \par 3/15/89) Book Entry: MSSB \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 3,085.25\cell \intbl\row\pard
'\trowd\trgaph70\trleft-70 \cellx2160\cellx2340\cellx8460\cellx8640\cellx10800\pard\widctlpar\intbl \pard\intbl\qr 140.0000\cell \cell \pard\intbl Xylem Inc. Com [XYL] \par recd 01/30/96 \par \cell \cell \pard\intbl\tqr\tx1940\tx2880 \tab 380.69\cell \intbl\row\pard \trowd\trgaph70\trleft-70 \cellx8460\cellx8640\clbrdrt\brdrs\brdrw10\cellx10800\pard\widctlpar\intbl \pard\intbl\qr Sub Total:\cell \cell \pard\intbl\tqr\tx1940\tx2880 $ \tab 86,836.62\cell \intbl\row\pard \pard \par \plain\f0\fs24\b \trowd\trgaph70\trleft-70 \cellx8460\cellx8640\clbrdrt\brdrs\brdrw10\cellx10800\pard\widctlpar\intbl \pard\intbl\qr Grand Total:\cell \plain\f0\fs24 \cell \plain\f0\fs24\b\uldb \pard\intbl\tqr\tx1940\tx2880 $ \tab 95,197.36\cell \intbl\row\pard }\margl720 \margt720 \margr360 \margb360 \pard \sect \sectd


'\margl720 \margt720 \margr360 \margb360}


'\rtf        Rich Text Format specification Version
'\ansi       ANSI character set
'\deff       Default font definition
'\deftab     Default tab width, in Twips
'
'\colortbl   Introduces color table group
'\red        Red index
'\green      Green index
'\blue       Blue index
'
'\fonttbl    Introduces font table group
'\froman     Roman, proportionally spaced serif font
'\fcharset   Character set of font in font table
'\fprq       Pitch of font in font table
'\f          Font number
'\fs         Font size
'\b          Bold
'
'\field      Introduces field destination
'\fldinst    Field instructions
'\*          New control words may be ignored if not recognized
'\cgrid
'
'\qc         Centered
'\margl      Left margin, in Twips
'\margt      Top margin, in Twips
'\margr      Right margin, in Twips
'\margb      Bottom margin, in Twips
'\par        New paragraph, end of paragraph, CrLf
'\pard       Reset to default paragraph properties
'\plain      Reset to default font properties
'\footer     Footer on all pages
'
'\brdrs      Border, Single-thickness
'\brdrw      Width of paragraph border line, in Twips
'
'\intbl      Paragraph is part of a table
'\cell       End of table cell
'\cellx      Defines right boundary of table cell, including half of space between cells
'\clbrdrl    Left table cell border
'\clbrdrt    top table cell border
'\clbrdrr    Right table cell border
'\clbrdrb    Bottom table cell border
'\row        End of table row
'\trowd      Set table row defaults
'\trgaph     Half the space between the cells of a table row in Twips
'\trleft     Position of leftmost edge of table with respect to left edge of column
'\widctlpar  Widow/orphan control for the current paragraph
' **

Private Const THIS_NAME As String = "zz_mod_NickersonProbatePlusFuncs"
' **

Public Function FindDate(varInput As Variant) As Variant

  Const THIS_PROC As String = "FindDate"

  Dim intPos1 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim intX As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  If IsNull(varInput) = False Then
    intPos1 = InStr(varInput, "acq")
    If intPos1 > 0 Then
      strTmp01 = Trim(varInput)
      strTmp02 = Trim(Mid(strTmp01, (intPos1 + 4)))
      intPos1 = InStr(strTmp02, " ")
      If intPos1 > 0 Then strTmp02 = Trim(Left(strTmp02, intPos1))
      intPos1 = InStr(strTmp02, ")")
      If intPos1 > 0 Then strTmp02 = Left(strTmp02, (intPos1 - 1))
      intPos1 = InStr(strTmp02, "]")
      If intPos1 > 0 Then strTmp02 = Left(strTmp02, (intPos1 - 1))
      If CharCnt(strTmp02, "/") = 2 Then  ' ** Module Function: modStringFuncs.
        If IsDate(strTmp02) = True Then
          strTmp03 = Format(CDate(strTmp02), "mm/dd/yyyy")
        End If
      Else
        strTmp02 = strTmp02 & "/" & CStr(year(Date))
        If IsDate(strTmp02) = True Then
          strTmp03 = Format(CDate(strTmp02), "mm/dd/yyyy")
        End If
      End If
      If strTmp03 <> vbNullString Then
        varRetVal = strTmp03
      End If
    Else
      intPos1 = InStr(varInput, "/")
      If intPos1 > 0 Then
        strTmp01 = Trim(varInput)
        strTmp02 = vbNullString: strTmp03 = vbNullString
        For intX = intPos1 To 1 Step -1
          If Mid(strTmp01, intX, 1) = " " Then
            strTmp02 = Trim(Mid(strTmp01, intX))
            Exit For
          End If
        Next
        If Left(strTmp02, 2) = "w/" Then
          intPos1 = InStr((intPos1 + 2), strTmp01, "/")
          strTmp02 = vbNullString
          If intPos1 > 0 Then
            For intX = intPos1 To 1 Step -1
              If Mid(strTmp01, intX, 1) = " " Then
                strTmp02 = Trim(Mid(strTmp01, intX))
                Exit For
              End If
            Next
          End If
        End If
        If Left(strTmp02, 2) = "Bk" Then
          intPos1 = InStr((intPos1 + 2), strTmp01, "/")
          strTmp02 = vbNullString
          If intPos1 > 0 Then
            For intX = intPos1 To 1 Step -1
              If Mid(strTmp01, intX, 1) = " " Then
                strTmp02 = Trim(Mid(strTmp01, intX))
                Exit For
              End If
            Next
          End If
        End If
        If Left(strTmp02, 4) = "acq." Then
          strTmp02 = Mid(strTmp02, 5)
          intPos1 = InStr(strTmp02, " ")
          If intPos1 > 0 Then strTmp02 = Trim(Left(strTmp02, intPos1))
        End If
        intPos1 = InStr(strTmp02, ")")
        If intPos1 > 0 Then strTmp02 = Left(strTmp02, (intPos1 - 1))  ' ** Found after date.
        intPos1 = InStr(strTmp02, ".")
        If intPos1 > 0 Then strTmp02 = Left(strTmp02, (intPos1 + 1))  ' ** Found before date.
        If strTmp02 <> vbNullString Then
          intPos1 = InStr(strTmp02, " ")
          If intPos1 > 0 Then strTmp02 = Trim(Left(strTmp02, intPos1))
          If CharCnt(strTmp02, "/") = 2 Then  ' ** Module Function: modStringFuncs.
            If IsDate(strTmp02) = True Then
              strTmp03 = Format(CDate(strTmp02), "mm/dd/yyyy")
            End If
          Else
            strTmp02 = strTmp02 & "/" & CStr(year(Date))
            If IsDate(strTmp02) = True Then
              strTmp03 = Format(CDate(strTmp02), "mm/dd/yyyy")
            End If
          End If
          If strTmp03 <> vbNullString Then
            varRetVal = strTmp03
          End If
        End If
      End If
    End If
  End If

  FindDate = varRetVal

End Function
