Attribute VB_Name = "modAddrOf"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modAddrOf"

'VGC 11/05/2010: CHANGES!

' ** This code is produced by
' ** Ken Getz and Michael Kaplan.
' ** This code is of course provided as is, without any warranty.
' ** Use this unsupported/non doc'd code at your own risk!

' ******************************************************************************************************************
' ** Declarations
' **
' ** These function names were puzzled out by using DUMPBIN /exports
' ** with VBA332.DLL and then puzzling out parameter names and types
' ** through a lot of trial and error and over 100 IPFs in MSACCESS.EXE
' ** and VBA332.DLL.
' **
' ** These parameters may not be named properly but seem to be correct
' ** in light of the function names and what each parameter does.
' **
' ** EbGetExecutingProj: Gives you a handle to the current VBA project.
' ** TipGetFunctionId: Gives you a function ID given a function name.
' ** TipGetLpfnOfFunctionId: Gives you a pointer a function given its function ID.

' ******************************************************************************************************************

Private Declare Function GetCurrentVbaProject Lib "VBE6.DLL" Alias "EbGetExecutingProj" (hProject As Long) As Long
'Private Declare Function GetCurrentVbaProject Lib "C:\VictorGCS_Clients\TrustAccountant\Ver1-7-4_Setup\vba332.dll" Alias "EbGetExecutingProj" (hProject As Long) As Long

Private Declare Function GetFuncID Lib "VBE6.DLL" Alias "TipGetFunctionId" _
  (ByVal hProject As Long, ByVal strFunctionName As String, ByRef strFunctionId As String) As Long
'Private Declare Function GetFuncID Lib "C:\VictorGCS_Clients\TrustAccountant\Ver1-7-4_Setup\vba332.dll" Alias "TipGetFunctionId" _
'  (ByVal hProject As Long, ByVal strFunctionName As String, ByRef strFunctionId As String) As Long

Private Declare Function GetAddr Lib "VBE6.DLL" Alias "TipGetLpfnOfFunctionId" _
  (ByVal hProject As Long, ByVal strFunctionId As String, ByRef lpfn As Long) As Long
'Private Declare Function GetAddr Lib "C:\VictorGCS_Clients\TrustAccountant\Ver1-7-4_Setup\vba332.dll" Alias "TipGetLpfnOfFunctionId" _
'  (ByVal hProject As Long, ByVal strFunctionId As String, ByRef lpfn As Long) As Long
' **

Public Function AddrOf(strFuncName As String) As Long
' ******************************************************************************************************************
' ** AddrOf
' **
' ** Returns a function pointer of a VBA public function given its name. This function
' ** gives similar functionality to VBA as VB5 has with the AddressOf param type.
' **
' ** NOTE: This function only seems to work if the proc you are trying to get a pointer
' **   to is in the current project. This makes sense, since we are using a function
' **   named EbGetExecutingProj.
' ******************************************************************************************************************
' **
' ** From Dev Ashish's site.
' **    http://www.mvps.org/access/api/api0031.htm
' ** Warnings:
' ** (1)    AddressOf is COMPLETELY UNSUPPORTED by Microsoft in Office 97
' **        environment. Use it at your own risk!!
' ** (2)    Entering debug mode is not recommended as it is likely to
' **        cause problems (GPFs etc.).
' ** (3)    Make sure you backup your work and save before running any such code.
' **        Using this technique adds another level of instability since
' **        there are so many different ways to set up things wrong.
' **        Once you get it to work properly, everything should be OK.
' ** (4)    Make sure you enter a On Error Resume Next at the top of
' **        any callback function. This is done to ensure that any errors within
' **        a callback function are not propagated back to its caller.
' ** (5)    Be careful of ByVal or ByRef when passing arguments to the function.
' **        If you don't get this right, nothing's going to work.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "AddrOf"

        Dim lngProject As Long
        Dim lngResult As Long
        Dim strID As String
        Dim lngPFN As Long
        Dim strFuncNameUnicode As String

        ' ** The function name must be in Unicode, so convert it.
110     strFuncNameUnicode = StrConv(strFuncName, vbUnicode)

        ' ** Get the current VBA project.
        ' ** The results of GetCurrentVBAProject seemed inconsistent, in our tests,
        ' ** so now we just check the project handle when the function returns.
120     Call GetCurrentVbaProject(lngProject)

        ' ** Make sure we got a project handle... we always should, but you never know!
130     If lngProject <> 0 Then
          ' ** Get the VBA function ID (whatever that is!).
140       lngResult = GetFuncID( _
            lngProject, strFuncNameUnicode, strID)
          ' ** We have to check this because we GPF if we try to get
          ' ** a function pointer of a non-existent function.
150       If lngResult = ERROR_SUCCESS Then
            ' ** Get the function pointer.
160         lngResult = GetAddr(lngProject, strID, lngPFN)
170         If lngResult = ERROR_SUCCESS Then
180           AddrOf = lngPFN
190         End If
200       End If
210     End If

EXITP:
220     Exit Function

ERRH:
230     Select Case ERR.Number
        Case Else
240       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
250     End Select
260     Resume EXITP

End Function

Public Function GetAccessVersionX() As Integer

300   On Error GoTo ERRH

        Const THIS_PROC As String = "GetAccessVersionX"

        Dim intRetVal As Integer

310     intRetVal = CInt(SysCmd(acSysCmdAccessVer))

320     If intRetVal = 9 Then  ' ** Access 2000 ONLY.
          'lpPrevWndProc = SetWindowLong( _
          '  mhWndMDIClient, _
          '  GWL_WNDPROC, _
          '  AddressOf modMDIClient.WndProc)  ' ** API Function: modWindowFunctions.
330     ElseIf intRetVal = 8 Then  ' ** Access 97  ONLY.
          ' ** You will obviously need Michael Kaplan's and
          ' ** Ken Getz's AddrOf function in Access 97.
          'lpPrevWndProc = SetWindowLong( _
          '  mhWndMDIClient, _
          '  GWL_WNDPROC, _
          '  AddrOf("WndProc"))  ' ** API Function: modWindowFunctions.
340     End If

EXITP:
350     GetAccessVersionX = intRetVal
360     Exit Function

ERRH:
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

End Function
