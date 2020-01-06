Attribute VB_Name = "MIntl"
Option Compare Database
Option Explicit

'MIT Licence
'Copyright © 2019, Francesco Foti, devinfo.net.
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software
'and associated documentation files (the “Software”), to deal in the Software without
'restriction, including without limitation the rights to use, copy, modify, merge, publish,
'distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the
'Software is furnished to do so, subject to the following conditions: The above copyright notice
'and this permission notice shall be included in all copies or substantial portions of the
'Software. The Software is provided “as is”, without warranty of any kind, express or implied,
'including but not limited to the warranties of merchantability, fitness for a particular purpose
'and noninfringement. In no event shall the authors or copyright holders (Francesco Foti) be
'liable for any claim, damages or other liability, whether in an action of contract, tort or
'otherwise, arising from, out of or in connection with the software or the use or other dealings
'in the Software. Except as contained in this notice, the name of Francesco Foti or devinfo.net
'shall not be used in advertising or otherwise to promote the sale, use or other dealings in this
'Software without prior written authorization from Francesco Foti.

Public Const LANG_SYSTEM_DEFAULT As Long = &H800
Public Const SUBLANG_SYS_DEFAULT As Long = &H2
Public Const LOCALE_USER_DEFAULT& = &H400

' PRIMARY LANGUAGE IDENTIFIERS
'https://docs.microsoft.com/en-us/windows/win32/intl/language-identifier-constants-and-strings
Public Const LANG_AFRIKAANS As Long = &H36
Public Const LANG_ALBANIAN As Long = &H1C
Public Const LANG_ARABIC As Long = &H1
Public Const LANG_ARMENIAN As Long = &H2B
Public Const LANG_ASSAMESE As Long = &H4D
Public Const LANG_AZERI As Long = &H2C
Public Const LANG_BASQUE As Long = &H2D
Public Const LANG_BELARUSIAN As Long = &H23
Public Const LANG_BENGALI As Long = &H45
Public Const LANG_BULGARIAN As Long = &H2
Public Const LANG_CATALAN As Long = &H3
Public Const LANG_CHINESE As Long = &H4
Public Const LANG_CROATIAN As Long = &H1A
Public Const LANG_CZECH As Long = &H5
Public Const LANG_DANISH As Long = &H6
Public Const LANG_DUTCH As Long = &H13
Public Const LANG_ENGLISH As Long = &H9
Public Const LANG_ESTONIAN As Long = &H25
Public Const LANG_FAEROESE As Long = &H38
Public Const LANG_FARSI As Long = &H29
Public Const LANG_FINNISH As Long = &HB
Public Const LANG_FRENCH As Long = &HC
Public Const LANG_GEORGIAN As Long = &H37
Public Const LANG_GERMAN As Long = &H7
Public Const LANG_GREEK As Long = &H8
Public Const LANG_GUJARATI As Long = &H47
Public Const LANG_HEBREW As Long = &HD
Public Const LANG_HINDI As Long = &H39
Public Const LANG_HUNGARIAN As Long = &HE
Public Const LANG_ICELANDIC As Long = &HF
Public Const LANG_INDONESIAN As Long = &H21
Public Const LANG_ITALIAN As Long = &H10
Public Const LANG_JAPANESE As Long = &H11
Public Const LANG_KANNADA As Long = &H4B
Public Const LANG_KASHMIRI As Long = &H60
Public Const LANG_KAZAK As Long = &H3F
Public Const LANG_KONKANI As Long = &H57
Public Const LANG_KOREAN As Long = &H12
Public Const LANG_LATVIAN As Long = &H26
Public Const LANG_LITHUANIAN As Long = &H27
Public Const LANG_MACEDONIAN As Long = &H2F
Public Const LANG_MALAY As Long = &H3E
Public Const LANG_MALAYALAM As Long = &H4C
Public Const LANG_MANIPURI As Long = &H58
Public Const LANG_MARATHI As Long = &H4E
Public Const LANG_NEPALI As Long = &H61
Public Const LANG_NEUTRAL As Long = &H0
Public Const LANG_NORWEGIAN As Long = &H14
Public Const LANG_ORIYA As Long = &H48
Public Const LANG_POLISH As Long = &H15
Public Const LANG_PORTUGUESE As Long = &H16
Public Const LANG_PUNJABI As Long = &H46
Public Const LANG_ROMANIAN As Long = &H18
Public Const LANG_RUSSIAN As Long = &H19
Public Const LANG_SANSKRIT As Long = &H4F
Public Const LANG_SERBIAN As Long = &H1A
Public Const LANG_SINDHI As Long = &H59
Public Const LANG_SLOVAK As Long = &H1B
Public Const LANG_SLOVENIAN As Long = &H24
Public Const LANG_SPANISH As Long = &HA
Public Const LANG_SWAHILI As Long = &H41
Public Const LANG_SWEDISH As Long = &H1D
Public Const LANG_TAMIL As Long = &H49
Public Const LANG_TATAR As Long = &H44
Public Const LANG_TELUGU As Long = &H4A
Public Const LANG_THAI As Long = &H1E
Public Const LANG_TURKISH As Long = &H1F
Public Const LANG_UKRAINIAN As Long = &H22
Public Const LANG_URDU As Long = &H20
Public Const LANG_UZBEK As Long = &H43
Public Const LANG_VIETNAMESE As Long = &H2A
'SUBLANG ids
Public Const SUBLANG_ARABIC_ALGERIA As Long = &H5
Public Const SUBLANG_ARABIC_BAHRAIN As Long = &HF
Public Const SUBLANG_ARABIC_EGYPT As Long = &H3
Public Const SUBLANG_ARABIC_IRAQ As Long = &H2
Public Const SUBLANG_ARABIC_JORDAN As Long = &HB
Public Const SUBLANG_ARABIC_KUWAIT As Long = &HD
Public Const SUBLANG_ARABIC_LEBANON As Long = &HC
Public Const SUBLANG_ARABIC_LIBYA As Long = &H4
Public Const SUBLANG_ARABIC_MOROCCO As Long = &H6
Public Const SUBLANG_ARABIC_OMAN As Long = &H8
Public Const SUBLANG_ARABIC_QATAR As Long = &H10
Public Const SUBLANG_ARABIC_SAUDI_ARABIA As Long = &H1
Public Const SUBLANG_ARABIC_SYRIA As Long = &HA
Public Const SUBLANG_ARABIC_TUNISIA As Long = &H7
Public Const SUBLANG_ARABIC_UAE As Long = &HE
Public Const SUBLANG_ARABIC_YEMEN As Long = &H9
Public Const SUBLANG_AZERI_CYRILLIC As Long = &H2
Public Const SUBLANG_AZERI_LATIN As Long = &H1
Public Const SUBLANG_CHINESE_HONGKONG As Long = &H3
Public Const SUBLANG_CHINESE_MACAU As Long = &H5
Public Const SUBLANG_CHINESE_SIMPLIFIED As Long = &H2
Public Const SUBLANG_CHINESE_SINGAPORE As Long = &H4
Public Const SUBLANG_CHINESE_TRADITIONAL As Long = &H1
Public Const SUBLANG_DEFAULT As Long = &H1
Public Const SUBLANG_DUTCH As Long = &H1
Public Const SUBLANG_DUTCH_BELGIAN As Long = &H2
Public Const SUBLANG_ENGLISH_AUS As Long = &H3
Public Const SUBLANG_ENGLISH_BELIZE As Long = &HA
Public Const SUBLANG_ENGLISH_CAN As Long = &H4
Public Const SUBLANG_ENGLISH_CARIBBEAN As Long = &H9
Public Const SUBLANG_ENGLISH_EIRE As Long = &H6
Public Const SUBLANG_ENGLISH_JAMAICA As Long = &H8
Public Const SUBLANG_ENGLISH_NZ As Long = &H5
Public Const SUBLANG_ENGLISH_PHILIPPINES As Long = &HD
Public Const SUBLANG_ENGLISH_SOUTH_AFRICA As Long = &H7
Public Const SUBLANG_ENGLISH_TRINIDAD As Long = &HB
Public Const SUBLANG_ENGLISH_UK As Long = &H2
Public Const SUBLANG_ENGLISH_US As Long = &H1
Public Const SUBLANG_ENGLISH_ZIMBABWE As Long = &HC
Public Const SUBLANG_FRENCH As Long = &H1
Public Const SUBLANG_FRENCH_BELGIAN As Long = &H2
Public Const SUBLANG_FRENCH_CANADIAN As Long = &H3
Public Const SUBLANG_FRENCH_LUXEMBOURG As Long = &H5
Public Const SUBLANG_FRENCH_MONACO As Long = &H6
Public Const SUBLANG_FRENCH_SWISS As Long = &H4
Public Const SUBLANG_GERMAN As Long = &H1
Public Const SUBLANG_GERMAN_AUSTRIAN As Long = &H3
Public Const SUBLANG_GERMAN_LIECHTENSTEIN As Long = &H5
Public Const SUBLANG_GERMAN_LUXEMBOURG As Long = &H4
Public Const SUBLANG_GERMAN_SWISS As Long = &H2
Public Const SUBLANG_ITALIAN As Long = &H1
Public Const SUBLANG_ITALIAN_SWISS As Long = &H2
Public Const SUBLANG_KASHMIRI_INDIA As Long = &H2
Public Const SUBLANG_KOREAN As Long = &H1
Public Const SUBLANG_LITHUANIAN As Long = &H1
Public Const SUBLANG_MALAY_BRUNEI_DARUSSALAM As Long = &H2
Public Const SUBLANG_MALAY_MALAYSIA As Long = &H1
Public Const SUBLANG_NEPALI_INDIA As Long = &H2
Public Const SUBLANG_NEUTRAL As Long = &H0
Public Const SUBLANG_NORWEGIAN_BOKMAL As Long = &H1
Public Const SUBLANG_NORWEGIAN_NYNORSK As Long = &H2
Public Const SUBLANG_PORTUGUESE As Long = &H2
Public Const SUBLANG_PORTUGUESE_BRAZILIAN As Long = &H1
Public Const SUBLANG_SERBIAN_CYRILLIC As Long = &H3
Public Const SUBLANG_SERBIAN_LATIN As Long = &H2
Public Const SUBLANG_SPANISH As Long = &H1
Public Const SUBLANG_SPANISH_ARGENTINA As Long = &HB
Public Const SUBLANG_SPANISH_BOLIVIA As Long = &H10
Public Const SUBLANG_SPANISH_CHILE As Long = &HD
Public Const SUBLANG_SPANISH_COLOMBIA As Long = &H9
Public Const SUBLANG_SPANISH_COSTA_RICA As Long = &H5
Public Const SUBLANG_SPANISH_DOMINICAN_REPUBLIC As Long = &H7
Public Const SUBLANG_SPANISH_ECUADOR As Long = &HC
Public Const SUBLANG_SPANISH_EL_SALVADOR As Long = &H11
Public Const SUBLANG_SPANISH_GUATEMALA As Long = &H4
Public Const SUBLANG_SPANISH_HONDURAS As Long = &H12
Public Const SUBLANG_SPANISH_MEXICAN As Long = &H2
Public Const SUBLANG_SPANISH_MODERN As Long = &H3
Public Const SUBLANG_SPANISH_NICARAGUA As Long = &H13
Public Const SUBLANG_SPANISH_PANAMA As Long = &H6
Public Const SUBLANG_SPANISH_PARAGUAY As Long = &HF
Public Const SUBLANG_SPANISH_PERU As Long = &HA
Public Const SUBLANG_SPANISH_PUERTO_RICO As Long = &H14
Public Const SUBLANG_SPANISH_URUGUAY As Long = &HE
Public Const SUBLANG_SPANISH_VENEZUELA As Long = &H8
Public Const SUBLANG_SWEDISH As Long = &H1
Public Const SUBLANG_SWEDISH_FINLAND As Long = &H2
Public Const SUBLANG_URDU_INDIA As Long = &H2
Public Const SUBLANG_URDU_PAKISTAN As Long = &H1
Public Const SUBLANG_UZBEK_CYRILLIC As Long = &H2
Public Const SUBLANG_UZBEK_LATIN As Long = &H1

'Win32 API (get the sizes right: https://en.cppreference.com/w/cpp/language/types)
#If Win64 Then
  Private Declare PtrSafe Function apiGetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
  Private Declare PtrSafe Function LCIDToLocaleName Lib "kernel32.dll" (ByVal piLCIDlocale As Long, ByVal plpRetLangCode As LongPtr, ByVal piLangCodeLen As Long, ByVal dwFlags As Long) As Long
#Else
  Private Declare Function apiGetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
  Private Declare Function LCIDToLocaleName Lib "kernel32.dll" (ByVal piLCIDlocale As Long, ByVal plpRetLangCode As Long, ByVal piLangCodeLen As Long, ByVal dwFlags As Long) As Long
#End If

'Windows locale constants that we may retrieve with LocaleInfo() function
Public Const LOCALE_SLONGDATE As Long = &H20
Public Const LOCALE_SLANGUAGE As Long = &H2
Public Const LOCALE_FONTSIGNATURE As Long = &H58
Public Const LOCALE_ICALENDARTYPE As Long = &H1009
Public Const LOCALE_ICENTURY As Long = &H24
Public Const LOCALE_ICOUNTRY As Long = &H5
Public Const LOCALE_ICURRDIGITS As Long = &H19
Public Const LOCALE_ICURRENCY As Long = &H1B
Public Const LOCALE_IDATE As Long = &H21
Public Const LOCALE_IDAYLZERO As Long = &H26
Public Const LOCALE_IDEFAULTANSICODEPAGE As Long = &H1004
Public Const LOCALE_IDEFAULTCODEPAGE As Long = &HB
Public Const LOCALE_IDEFAULTCOUNTRY As Long = &HA
Public Const LOCALE_IDEFAULTEBCDICCODEPAGE As Long = &H1012
Public Const LOCALE_IDEFAULTLANGUAGE As Long = &H9
Public Const LOCALE_IDEFAULTMACCODEPAGE As Long = &H1011
Public Const LOCALE_IDIGITS As Long = &H11
Public Const LOCALE_IDIGITSUBSTITUTION As Long = &H1014
Public Const LOCALE_IFIRSTDAYOFWEEK As Long = &H100C
Public Const LOCALE_IFIRSTWEEKOFYEAR As Long = &H100D
Public Const LOCALE_IINTLCURRDIGITS As Long = &H1A
Public Const LOCALE_ILANGUAGE As Long = &H1
Public Const LOCALE_ILDATE As Long = &H22
Public Const LOCALE_ILZERO As Long = &H12
Public Const LOCALE_IMEASURE As Long = &HD
Public Const LOCALE_IMONLZERO As Long = &H27
Public Const LOCALE_INEGCURR As Long = &H1C
Public Const LOCALE_INEGNUMBER As Long = &H1010
Public Const LOCALE_INEGSEPBYSPACE As Long = &H57
Public Const LOCALE_INEGSIGNPOSN As Long = &H53
Public Const LOCALE_INEGSYMPRECEDES As Long = &H56
Public Const LOCALE_IOPTIONALCALENDAR As Long = &H100B
Public Const LOCALE_IPAPERSIZE As Long = &H100A
Public Const LOCALE_IPOSSEPBYSPACE As Long = &H55
Public Const LOCALE_IPOSSIGNPOSN As Long = &H52
Public Const LOCALE_IPOSSYMPRECEDES As Long = &H54
Public Const LOCALE_ITIME As Long = &H23
Public Const LOCALE_ITIMEMARKPOSN As Long = &H1005
Public Const LOCALE_ITLZERO As Long = &H25
Public Const LOCALE_NOUSEROVERRIDE As Long = &H80000000
Public Const LOCALE_RETURN_NUMBER As Long = &H20000000
Public Const LOCALE_S1159 As Long = &H28
Public Const LOCALE_S2359 As Long = &H29
Public Const LOCALE_SABBREVCTRYNAME As Long = &H7
Public Const LOCALE_SABBREVDAYNAME1 As Long = &H31
Public Const LOCALE_SABBREVDAYNAME2 As Long = &H32
Public Const LOCALE_SABBREVDAYNAME3 As Long = &H33
Public Const LOCALE_SABBREVDAYNAME4 As Long = &H34
Public Const LOCALE_SABBREVDAYNAME5 As Long = &H35
Public Const LOCALE_SABBREVDAYNAME6 As Long = &H36
Public Const LOCALE_SABBREVDAYNAME7 As Long = &H37
Public Const LOCALE_SABBREVLANGNAME As Long = &H3
Public Const LOCALE_SABBREVMONTHNAME1 As Long = &H44
Public Const LOCALE_SABBREVMONTHNAME10 As Long = &H4D
Public Const LOCALE_SABBREVMONTHNAME11 As Long = &H4E
Public Const LOCALE_SABBREVMONTHNAME12 As Long = &H4F
Public Const LOCALE_SABBREVMONTHNAME13 As Long = &H100F
Public Const LOCALE_SABBREVMONTHNAME2 As Long = &H45
Public Const LOCALE_SABBREVMONTHNAME3 As Long = &H46
Public Const LOCALE_SABBREVMONTHNAME4 As Long = &H47
Public Const LOCALE_SABBREVMONTHNAME5 As Long = &H48
Public Const LOCALE_SABBREVMONTHNAME6 As Long = &H49
Public Const LOCALE_SABBREVMONTHNAME7 As Long = &H4A
Public Const LOCALE_SABBREVMONTHNAME8 As Long = &H4B
Public Const LOCALE_SABBREVMONTHNAME9 As Long = &H4C
Public Const LOCALE_SCOUNTRY As Long = &H6
Public Const LOCALE_SCURRENCY As Long = &H14
Public Const LOCALE_SDATE As Long = &H1D
Public Const LOCALE_SDAYNAME1 As Long = &H2A
Public Const LOCALE_SDAYNAME2 As Long = &H2B
Public Const LOCALE_SDAYNAME3 As Long = &H2C
Public Const LOCALE_SDAYNAME4 As Long = &H2D
Public Const LOCALE_SDAYNAME5 As Long = &H2E
Public Const LOCALE_SDAYNAME6 As Long = &H2F
Public Const LOCALE_SDAYNAME7 As Long = &H30
Public Const LOCALE_SDECIMAL As Long = &HE
Public Const LOCALE_SENGCOUNTRY As Long = &H1002
Public Const LOCALE_SENGCURRNAME As Long = &H1007
Public Const LOCALE_SENGLANGUAGE As Long = &H1001
Public Const LOCALE_SGROUPING As Long = &H10
Public Const LOCALE_SINTLSYMBOL As Long = &H15
Public Const LOCALE_SISO3166CTRYNAME As Long = &H5A
Public Const LOCALE_SISO639LANGNAME As Long = &H59
Public Const LOCALE_SLIST As Long = &HC
Public Const LOCALE_SMONDECIMALSEP As Long = &H16
Public Const LOCALE_SMONGROUPING As Long = &H18
Public Const LOCALE_SMONTHNAME1 As Long = &H38
Public Const LOCALE_SMONTHNAME10 As Long = &H41
Public Const LOCALE_SMONTHNAME11 As Long = &H42
Public Const LOCALE_SMONTHNAME12 As Long = &H43
Public Const LOCALE_SMONTHNAME13 As Long = &H100E
Public Const LOCALE_SMONTHNAME2 As Long = &H39
Public Const LOCALE_SMONTHNAME3 As Long = &H3A
Public Const LOCALE_SMONTHNAME4 As Long = &H3B
Public Const LOCALE_SMONTHNAME5 As Long = &H3C
Public Const LOCALE_SMONTHNAME6 As Long = &H3D
Public Const LOCALE_SMONTHNAME7 As Long = &H3E
Public Const LOCALE_SMONTHNAME8 As Long = &H3F
Public Const LOCALE_SMONTHNAME9 As Long = &H40
Public Const LOCALE_SMONTHOUSANDSEP As Long = &H17
Public Const LOCALE_SNATIVECTRYNAME As Long = &H8
Public Const LOCALE_SNATIVECURRNAME As Long = &H1008
Public Const LOCALE_SNATIVEDIGITS As Long = &H13
Public Const LOCALE_SNATIVELANGNAME As Long = &H4
Public Const LOCALE_SNEGATIVESIGN As Long = &H51
Public Const LOCALE_SPOSITIVESIGN As Long = &H50
Public Const LOCALE_SSHORTDATE As Long = &H1F
Public Const LOCALE_SSORTNAME As Long = &H1013
Public Const LOCALE_STHOUSAND As Long = &HF
Public Const LOCALE_STIME As Long = &H1E
Public Const LOCALE_STIMEFORMAT As Long = &H1003
Public Const LOCALE_SYEARMONTH As Long = &H1006
Public Const LOCALE_USE_CP_ACP As Long = &H40000000

'
' Wrapping API basics for manipulating languages and locales IDs
'

Public Function PRIMARYLANGID(ByVal plLangID As Long) As Long
  '#define PRIMARYLANGID(lgid)    ((WORD  )(lgid) & 0x3ff)
  PRIMARYLANGID = plLangID And &H3FF
End Function

Public Function SUBLANGID(ByVal plLangID As Long) As Long
  '#define SUBLANGID(lgid)        ((WORD  )(lgid) >> 10)
  SUBLANGID = plLangID \ 1024&
End Function

Public Function MAKELANGID(ByVal plPrimaryLangID As Long, ByVal plSubLangID As Long) As Long
  '#define MAKELANGID(p, s) ((((WORD) (s)) << 10) | (WORD) (p))
  MAKELANGID = (plSubLangID * 1024&) Or plPrimaryLangID
End Function

'https://docs.microsoft.com/fr-fr/windows/win32/api/winnls/nf-winnls-lcidtolocalename
Public Function IntlLCIDToLocaleName(ByVal plLCID As Long, Optional ByVal pdwFlags As Long = 0&) As String
  Dim lAPIRet       As Long
  Dim sBuffer       As String
  
  lAPIRet = LCIDToLocaleName(plLCID, StrPtr(sBuffer), 0, pdwFlags)
  If lAPIRet > 0 Then
    sBuffer = Space$(lAPIRet)
    lAPIRet = LCIDToLocaleName(plLCID, StrPtr(sBuffer), lAPIRet, pdwFlags)
    If lAPIRet = 0& Then
      Debug.Print LastDllErrorMsg
    End If
    IntlLCIDToLocaleName = CtoVB(sBuffer)
  End If
End Function

Public Function IntlLocaleInfo(plngLCType As Long, Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  Dim lngLocale As Long
  Dim strLCData As String, lngData As Long
  Dim lngX As Long
  Const cMAXLEN = 255
  
  strLCData = String$(cMAXLEN, 0)
  lngData = cMAXLEN - 1
  lngX = apiGetLocaleInfo(plLCID, plngLCType, strLCData, lngData)
  If lngX <> 0 Then
    IntlLocaleInfo = Left$(strLCData, lngX - 1)
  End If
End Function

'
' Wrapping some common LOCALE infos
'

Public Function IntlMonthName(ByVal plMonthIndex As Long, Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  Debug.Assert (plMonthIndex > 0&) And (plMonthIndex < 13&) 'Month index must be [1..12]
  IntlMonthName = IntlLocaleInfo(LOCALE_SMONTHNAME1 - 1& + plMonthIndex, plLCID)
End Function

Public Function IntlAbbrevMonthName(ByVal plMonthIndex As Long, Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  Debug.Assert (plMonthIndex > 0&) And (plMonthIndex < 13&) 'Month index must be [1..12]
  IntlAbbrevMonthName = IntlLocaleInfo(LOCALE_SABBREVMONTHNAME1 - 1& + plMonthIndex, plLCID)
End Function

Public Function IntlDayName(ByVal plDayIndex As Long, Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  Debug.Assert (plDayIndex > 0&) And (plDayIndex < 8&) ' Day index must be [1..7]
  IntlDayName = IntlLocaleInfo(LOCALE_SDAYNAME1 - 1& + plDayIndex, plLCID)
End Function

Public Function IntlAbbrevDayName(ByVal plDayIndex As Long, Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  Debug.Assert (plDayIndex > 0&) And (plDayIndex < 8&) ' Day index must be [1..7]
  IntlAbbrevDayName = IntlLocaleInfo(LOCALE_SABBREVDAYNAME1 - 1& + plDayIndex, plLCID)
End Function

Public Function IntlUseAMPM(Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As Boolean
  IntlUseAMPM = CBool(IntlLocaleInfo(LOCALE_ITIME, plLCID) = "0")
End Function

Public Function IntlIsSuffixAMPM(Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As Boolean
  IntlIsSuffixAMPM = CBool(IntlLocaleInfo(LOCALE_ITIMEMARKPOSN, plLCID) = "0")
End Function

Public Function IntlAM(Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  IntlAM = IntlLocaleInfo(LOCALE_S1159, plLCID)
End Function

Public Function IntlPM(Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  IntlPM = IntlLocaleInfo(LOCALE_S2359, plLCID)
End Function

Public Function IntlTwoDigitsCentury(Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As Boolean
  IntlTwoDigitsCentury = CBool(IntlLocaleInfo(LOCALE_ICENTURY, plLCID) = "0")
End Function

Public Function IntlDateSep(Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  IntlDateSep = IntlLocaleInfo(LOCALE_SDATE, plLCID)
End Function

Public Function IntlTimeSep(Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  IntlTimeSep = IntlLocaleInfo(LOCALE_STIME)
End Function

Public Function IntlThousandSep(Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  IntlThousandSep = IntlLocaleInfo(LOCALE_STHOUSAND, plLCID)
End Function

Public Function IntlDecimalSep(Optional ByVal plLCID As Long = LOCALE_USER_DEFAULT) As String
  IntlDecimalSep = IntlLocaleInfo(LOCALE_SDECIMAL, plLCID)
End Function

