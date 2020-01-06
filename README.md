# Using national language support in Office with VBA

MIntl.bas is a VBA standard module wrapping some of the most useful NLS Windows API, immediately actionable when you add it to your VB project.

You may want to read my blog post here, where you'll find a brief explanation of the underlying concepts, to get started quickly.

Basically, the module wraps two essential Win32 API functions, GetLocaleInfo() and LCIDToLocaleName(), and provides the necessary declarations of constants for language IDs, sublanguage IDs and Locale IDs.

Some additional functions are useful shortcuts wrapping GetLocalInfo(), lile IntlDayName() or IntlUseAMPM(). They're more readable than the corresponding IntlLocaleInfo() function call they wrap, and they return an appropriately converted variable type.



