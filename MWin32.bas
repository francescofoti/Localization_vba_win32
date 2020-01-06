Attribute VB_Name = "MWin32"
Option Compare Database
Option Explicit

'Windows API used only in this module
Private Const FORMAT_MESSAGE_FROM_SYSTEM  As Long = &H1000&
Private Declare Function GetLastError& Lib "kernel32" ()
Private Declare Function FormatMessageW Lib "kernel32" (ByVal pdwFlags As Long, ByVal plSource As Long, ByVal pdwMessageId As Long, ByVal pdwLanguageId As Long, ByVal plBuffer As Long, ByVal plSize As Long, plArguments As Long) As Long

'Cut string before trailing chr$(0)
Public Function CtoVB(ByRef pszString As String) As String
  Dim i   As Long
  i = InStr(pszString, Chr$(0))
  If i Then
    CtoVB = Left$(pszString, i - 1&)
  Else
    CtoVB = pszString
  End If
End Function

Public Function LastDllErrorMsg(Optional ByVal plErrCode As Long) As String
  Dim sBuffer         As String   ' Place where error description will be copied to.
  Dim lCopiedCt       As Long     ' Number of bytes copied to sBuffer
  Const BUFFER_SIZE As Long = 2048&
  
  If plErrCode = 0 Then                       ' no error code supplied
    plErrCode = Err.LastDllError              ' use the VB last known API error code
  Else
    plErrCode = Abs(plErrCode)                ' user supplied DLL error code
  End If
  If plErrCode = 0 Then Exit Function         ' bail if no error code
  ' prepare the buffer
  sBuffer = Space$(BUFFER_SIZE)
  ' translate the error code
  lCopiedCt = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, _
                         0&, plErrCode, _
                         0&, StrPtr(sBuffer), BUFFER_SIZE - 1&, 0&)
  LastDllErrorMsg = CtoVB(sBuffer)
End Function

