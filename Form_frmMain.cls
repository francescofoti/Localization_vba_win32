VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' it starts here

Private Sub Form_Load()
  ReloadAll 'show info with system defaults
End Sub

' load the listboxes

Private Sub LoadDayNames()
  Dim sRowSource    As String
  Dim i             As Integer
  
  For i = 1 To 7
    If i > 1 Then sRowSource = sRowSource & ";"
    sRowSource = sRowSource & IntlDayName(i, glCurrentLCID)
  Next i
  
  Me.lstDayName.RowSource = sRowSource
End Sub

Private Sub LoadAbbrevDayNames()
  Dim sRowSource    As String
  Dim i             As Integer
  
  For i = 1 To 7
    If i > 1 Then sRowSource = sRowSource & ";"
    sRowSource = sRowSource & IntlAbbrevDayName(i, glCurrentLCID)
  Next i
  
  Me.lstAbbrevDayName.RowSource = sRowSource
End Sub

Private Sub LoadMonthNames()
  Dim sRowSource    As String
  Dim i             As Integer
  
  For i = 1 To 12
    If i > 1 Then sRowSource = sRowSource & ";"
    sRowSource = sRowSource & IntlMonthName(i, glCurrentLCID)
  Next i
  
  Me.lstMonthName.RowSource = sRowSource
End Sub

Private Sub LoadAbbrevMonthNames()
  Dim sRowSource    As String
  Dim i             As Integer
  
  For i = 1 To 12
    If i > 1 Then sRowSource = sRowSource & ";"
    sRowSource = sRowSource & IntlAbbrevMonthName(i, glCurrentLCID)
  Next i
  
  Me.lstAbbrevMonthName.RowSource = sRowSource
End Sub

Private Function BoolString(ByVal pfFlag As Boolean)
  If pfFlag Then
    BoolString = "YES"
  Else
    BoolString = "NO"
  End If
End Function

Private Sub LoadOfficeLangInfos()
  Dim sRowSource    As String
  
  'Infos not from the Access, but from Intlxxx common wrappers around IntlLocaleInfo()
  sRowSource = sRowSource & "UseAMPM: " & BoolString(IntlUseAMPM(glCurrentLCID)) & ";"
  sRowSource = sRowSource & "IsSuffixAMPM: " & BoolString(IntlIsSuffixAMPM(glCurrentLCID)) & ";"
  sRowSource = sRowSource & "AM: " & IntlAM(glCurrentLCID) & ";"
  sRowSource = sRowSource & "PM: " & IntlPM(glCurrentLCID) & ";"
  sRowSource = sRowSource & "TwoDigitsCentury: " & BoolString(IntlTwoDigitsCentury(glCurrentLCID)) & ";"
  sRowSource = sRowSource & "DateSep: " & IntlDateSep(glCurrentLCID) & ";"
  sRowSource = sRowSource & "TimeSep: " & IntlTimeSep(glCurrentLCID) & ";"
  sRowSource = sRowSource & "ThousandSep: " & IntlThousandSep(glCurrentLCID) & ";"
  sRowSource = sRowSource & "DecimalSep: " & IntlDecimalSep(glCurrentLCID) & ";"
  
  With Application.LanguageSettings
    sRowSource = sRowSource & "---- From the Application object ----;"
    sRowSource = sRowSource & "msoLanguageIDExeMode=" & IntlLCIDToLocaleName(.LanguageID(msoLanguageIDExeMode)) & " (=" & .LanguageID(msoLanguageIDExeMode) & " / &H" & Hex$(.LanguageID(msoLanguageIDExeMode)) & ")" & ";"
    sRowSource = sRowSource & "msoLanguageIDHelp=" & IntlLCIDToLocaleName(.LanguageID(msoLanguageIDHelp)) & " (=" & .LanguageID(msoLanguageIDHelp) & " / &H" & Hex$(.LanguageID(msoLanguageIDHelp)) & ")" & ";"
    sRowSource = sRowSource & "msoLanguageIDInstall=" & IntlLCIDToLocaleName(.LanguageID(msoLanguageIDInstall)) & " (=" & .LanguageID(msoLanguageIDInstall) & " / &H" & Hex$(.LanguageID(msoLanguageIDInstall)) & ")" & ";"
    sRowSource = sRowSource & "msoLanguageIDUI=" & IntlLCIDToLocaleName(.LanguageID(msoLanguageIDUI)) & " (=" & .LanguageID(msoLanguageIDUI) & " / &H" & Hex$(.LanguageID(msoLanguageIDUI)) & ")" & ";"
    sRowSource = sRowSource & "msoLanguageIDUIPrevious=" & IntlLCIDToLocaleName(.LanguageID(msoLanguageIDUIPrevious)) & " (=" & .LanguageID(msoLanguageIDUIPrevious) & " / &H" & Hex$(.LanguageID(msoLanguageIDUIPrevious)) & ")" & ";"
    sRowSource = sRowSource & Me.txtLangCode & " is "
    If Not .LanguagePreferredForEditing(glCurrentLCID) Then
      sRowSource = sRowSource & " NOT "
    End If
    sRowSource = sRowSource & " one of the preferred editing languages"
  End With
  
  Me.lstAccessApp.RowSource = sRowSource
End Sub

' display/clear error

Private Sub ClearError()
  Me.txtError.Visible = False
  Me.txtError = ""
End Sub

Private Sub ShowError(ByVal psErrorText As String)
  Me.txtError = psErrorText
  Me.txtError.Visible = True
End Sub

Private Sub ShowMessage(ByVal psMessage As String)
  Me.txtMessage = psMessage
  Me.txtMessage.Visible = True
End Sub

Private Sub ClearMessage()
  Me.txtMessage.Visible = False
  Me.txtMessage = ""
End Sub

' reload everything

Private Sub SetLanguage(ByVal plLCID As Long)
  Dim sLangCode As String
  
  On Error Resume Next
  ClearError
  
  glCurrentLCID = plLCID
  
  sLangCode = IntlLCIDToLocaleName(glCurrentLCID)
  If Len(sLangCode) = 0 Then
    ShowError LastDllErrorMsg
  End If
  
  Me.txtLangCode = sLangCode
  LoadDayNames
  LoadAbbrevDayNames
  LoadMonthNames
  LoadAbbrevMonthNames
  LoadOfficeLangInfos
  Me.lstOtherSettings.Requery
End Sub

Private Sub ReloadAll()
  Dim vLangID     As Variant
  Dim vSubLangID  As Variant
  
  ClearMessage
  If Me.chkUserLocaleDefault Then
    SetLanguage LOCALE_USER_DEFAULT&
    ShowMessage "Currently displaying default language settings for the CURRENT USER"
  Else
    vLangID = Nz(Me.cboLanguageID, LANG_SYSTEM_DEFAULT)
    vSubLangID = Nz(Me.cboSubLanguageID.Column(1), SUBLANG_SYS_DEFAULT)
    If (vLangID = LANG_SYSTEM_DEFAULT) And (vSubLangID = SUBLANG_SYS_DEFAULT) Then
      ShowMessage "Currently displaying language settings for the SYSTEM DEFAULT"
    End If
  End If
  
  SetLanguage MAKELANGID(vLangID, vSubLangID)
End Sub

Private Sub chkUserLocaleDefault_AfterUpdate()
  Dim fUseUserDefault As Boolean
  
  fUseUserDefault = CBool(Me.chkUserLocaleDefault = True)
  Me.cboLanguageID.Enabled = Not fUseUserDefault
  Me.cboSubLanguageID.Enabled = Not fUseUserDefault
  
  ReloadAll
End Sub

' when combos change

Private Sub cboLanguageID_AfterUpdate()
  Me.cboSubLanguageID = Null
  ReloadAll
End Sub

Private Sub cboSubLanguageID_AfterUpdate()
  ReloadAll
End Sub

