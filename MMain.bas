Attribute VB_Name = "MMain"
Option Compare Database
Option Explicit

Public glCurrentLCID As Long

'Quick and dirty for this test application.
'The glCurrentLCID is updated in frmMain.
'This is a global function that is called from the SQL RowSource of lstOtherSettings in frmMain.
Public Function GetLocale(ByVal plngLCType As Long) As String
  GetLocale = IntlLocaleInfo(plngLCType, glCurrentLCID)
End Function

