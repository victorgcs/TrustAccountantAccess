Attribute VB_Name = "modHelpFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modHelpFuncs"

'VGC 11/06/2011: CHANGES!

'Questions:
'1. We have a commercially sold Microsoft Access application that comes with a PDF User Manual.
'   We'd like to show and reference this Help file throughout our application.
'   We own Acrobat 9.0 Standard.
'   We'll be satisfied with either of these two options:
'   a. Open an Acrobat window within our Microsoft Access application.
'   b. Open an Acrobat Reader externally and go to the relevant page.
'
'2. Is there some type of ActiveX, Add-In, or Plug-In that can provide all
'   the functionality we want as a single component?
'
'3. I've written VBA code that opens an external Acrobat window,
'   and I'm able to move to various pages.
'   However, I have questions about this current code.
'   a. Will these same commands open whatever Reader the end-user might have?
'   b. My Acrobat remains open, even though I exit and destroy the object:
'        acro_app.Exit
'        Set acro_app = Nothing
'   c. As I've been testing it, occasionally a command will try to open and go to a page,
'      but the PDF window is frozen, and won't repaint.
'      What is that, and how do I guard against it?
'   d. What components must be distributed with our application to provide this PDF functionality?
'   e. Though the Acrobat.tlb is available, it isn't currently linked as a Reference.
'      Why do my commands still work?
'

' ** References:
' *    Application.References.Count = 11
' **   Application.VBE.ActiveVBProject.References.Count = 11
' **   Adobe Acrobat Standard:
' **     Application.References(11).Name = Acrobat
' **     Application.References("Acrobat").FullPath = C:\Program Files\Adobe\Acrobat 6.0\Acrobat\Acrobat.tlb
' **     Application.VBE.ActiveVBProject.References(11).Name = Acrobat
' **     Application.VBE.ActiveVBProject.References(11).Description = Adobe Acrobat 6.0 Type Library
' **   Adobe Acrobat Reader:
' **     Application.References(11).Name = Acrobat
' **     Application.References(11).FullPath = C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.dll
' **     Application.VBE.ActiveVBProject.References(11).Name = Acrobat
' **     Application.VBE.ActiveVBProject.References(11).Description = Adobe Acrobat 8.0 Type Library

'app is the object that represents the Acrobat application.
'Anything that has to do with application level information is handled here
'(e.g. the app user interface, or creating documents).

'AcroPDDoc represents a PDF document, this document does not necessarily have to be displayed in Acrobat
'(it may be open in the background), an AcroAVDoc is a PDF document that is actually being displayed in Acrobat.
'Every AcroAVDoc has a AcroPDFDoc (but not vice versa).
'I assume that an AcroAXDoc is a document that contains an XML form, but I've never used these.

'You need to read Adobe's API documentation, you can find that on
'http://www.adobe.com/devnet/acrobat - you want the documentation about the IAC interface.

'When you open a PDF document in the browser, the document actually gets downloaded,
'and as far as Acrobat is concerned, it gets opened as an "external" document, and is available to some API functions.

' ** These variables should be made Public if, and when, context-sensitive Help is implemented.
' ** They were to be used in frmJournal.
Private strAcro_PathFile As String
Private acro_app As Object  'As AcroApp
Private acro_pddoc As Object
Private acro_avdoc As Object  'As AcroAVDoc
Private acro_avpag As Object  'As AcroAVPageView
' **

Public Function Help_OpenPage(lngPage As Long) As Boolean
' ** This works with the Reference set to the Adobe Acrobat 8.0 Type Library, found under the Reader folder.
' ** Trust Accountant Manual page numbers:
' **   28  29  Dividend
' **   33  34  Interest
' **   34  35  Purchase/Deposit
' **   39  40  Sale/Withdrawal
' **   41  42  Miscellaneous

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Help_OpenPage"

        Dim blnRetVal As Boolean

110     blnRetVal = True

120     If strAcro_PathFile = vbNullString Then
130       blnRetVal = Help_GetPath  ' ** Function: Below.
140     End If

150     If blnRetVal = True Then
160       Set acro_app = CreateObject("AcroExch.app")
170       With acro_app
180         .Show
190         Set acro_pddoc = CreateObject("AcroExch.pddoc")
200         With acro_pddoc
210           .Open strAcro_PathFile
220           Set acro_avdoc = .OpenAVDoc(strAcro_PathFile)
230           With acro_avdoc
                '.SetTitle "Trust Accountant Manual"
250             Set acro_avpag = .GetAVPageView
260             With acro_avpag
270               .GoTo lngPage
                  'Stop
280             End With
290           End With
300         End With
310       End With
320     End If

EXITP:
330     Help_OpenPage = blnRetVal
340     Exit Function

ERRH:
350     blnRetVal = False
360     Select Case ERR.Number
        Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
380     End Select
390     Resume EXITP

End Function

Public Function Help_OpenAll() As Boolean

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Help_OpenAll"

        Dim blnRetVal As Boolean

410     blnRetVal = True

420     If strAcro_PathFile = vbNullString Then
430       blnRetVal = Help_GetPath  ' ** Function: Below.
440     End If

450     If blnRetVal = True Then
460       OpenHelp (strAcro_PathFile)  ' ** Module Function: modShellFuncs.
470     End If

EXITP:
480     Help_OpenAll = blnRetVal
490     Exit Function

ERRH:
500     blnRetVal = False
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Function

Public Function Help_GetPath() As Boolean

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Help_GetPath"

        Dim blnRetVal As Boolean

610     blnRetVal = True

        'strPathFile = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Documents\TA Manual.pdf"  '## OK
620     strAcro_PathFile = CurrentAppPath & LNK_SEP & gstrFile_Manual  ' ** Module Function: modFileUtilities.
630     If FileExists(strAcro_PathFile) = True Then  ' ** Module Function: modFileUtilities.
          ' ** It's where it belongs.
640     Else
650       blnRetVal = False
660       strAcro_PathFile = vbNullString
670       MsgBox "The file, " & gstrFile_Manual & ", could not be found in" & vbCrLf & _
            CurrentAppPath & vbCrLf & vbCrLf & "If you do not have it, contact Delta Data, Inc.", _
            vbInformation + vbOKOnly, "Trust Accountant Manual Not Found"  ' ** Module Function: modFileUtilities.
680     End If

EXITP:
690     Help_GetPath = blnRetVal
700     Exit Function

ERRH:
710     blnRetVal = False
720     Select Case ERR.Number
        Case Else
730       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
740     End Select
750     Resume EXITP

End Function
