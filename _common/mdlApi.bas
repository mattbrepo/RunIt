Attribute VB_Name = "mdlAPI"
Option Explicit

Public Const RASTERCAPS As Long = 38
Public Const RC_PALETTE As Long = &H100
Public Const SIZEPALETTE As Long = 104

Public Const MEM_COMMIT As Long = &H1000
Public Const PAGE_READWRITE As Long = &H4

Public Const DIR_CURRENT_USER As String = "Control Panel\International"
Public Const DIR_USERS As String = ".DEFAULT\Control Panel\International"

Public Const LOGON_WITH_PROFILE = &H1
Public Const LOGON_NETCREDENTIALS_ONLY = &H2

Public Const LOGON32_PROVIDER_DEFAULT As Long = 0
Public Const LOGON32_PROVIDER_WINNT35 As Long = 1
Public Const LOGON32_PROVIDER_WINNT40 As Long = 2
Public Const LOGON32_PROVIDER_WINNT50 As Long = 3

Public Const LOGON32_LOGON_INTERACTIVE As Long = 2
Public Const LOGON32_LOGON_NETWORK As Long = 3
Public Const LOGON32_LOGON_BATCH As Long = 4
Public Const LOGON32_LOGON_SERVICE As Long = 5
Public Const LOGON32_LOGON_UNLOCK As Long = 7
Public Const LOGON32_LOGON_NETWORK_CLEARTEXT As Long = 8
Public Const LOGON32_LOGON_NEW_CREDENTIALS As Long = 9

Public Const PROCESS_VM_OPERATION As Long = (&H8)
Public Const PROCESS_VM_READ As Long = (&H10)
Public Const PROCESS_VM_WRITE As Long = (&H20)
Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)

Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
Public Const LVM_GETITEMPOSITION As Long = (LVM_FIRST + 16)
Public Const LVM_GETITEMRECT As Long = (LVM_FIRST + 14)
Public Const LVM_GETITEMSPACING As Long = (LVM_FIRST + 51)
Public Const LVM_GETITEMA As Long = (LVM_FIRST + 5)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXTA As Long = (LVM_FIRST + 45)
Public Const LVM_GETITEMTEXTW As Long = (LVM_FIRST + 115)
Public Const LVM_GETITEMW As Long = (LVM_FIRST + 75)
Public Const LVM_GETNEXTITEM As Long = (LVM_FIRST + 12)
Public Const LVM_GETNUMBEROFWORKAREAS As Long = (LVM_FIRST + 73)
Public Const LVM_GETORIGIN As Long = (LVM_FIRST + 41)
Public Const LVM_GETOUTLINECOLOR As Long = (LVM_FIRST + 176)
Public Const LVM_GETSELECTEDCOLUMN As Long = (LVM_FIRST + 174)
Public Const LVM_GETSELECTEDCOUNT As Long = (LVM_FIRST + 50)
Public Const LVM_GETSELECTIONMARK As Long = (LVM_FIRST + 66)
Public Const LVM_GETSTRINGWIDTHA As Long = (LVM_FIRST + 17)
Public Const LVM_GETSTRINGWIDTHW As Long = (LVM_FIRST + 87)
Public Const LVM_GETSUBITEMRECT As Long = (LVM_FIRST + 56)
Public Const LVM_GETTEXTBKCOLOR As Long = (LVM_FIRST + 37)
Public Const LVM_GETTEXTCOLOR As Long = (LVM_FIRST + 35)
Public Const LVM_GETTILEINFO As Long = (LVM_FIRST + 165)
Public Const LVM_GETTILEVIEWINFO As Long = (LVM_FIRST + 163)
Public Const LVM_GETTOOLTIPS As Long = (LVM_FIRST + 78)
Public Const LVM_GETTOPINDEX As Long = (LVM_FIRST + 39)
Public Const LVM_GETVIEW As Long = (LVM_FIRST + 143)
Public Const LVM_GETVIEWRECT As Long = (LVM_FIRST + 34)
Public Const LVM_GETWORKAREAS As Long = (LVM_FIRST + 70)
Public Const LVM_HASGROUP As Long = (LVM_FIRST + 161)
Public Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Public Const LVM_INSERTCOLUMNA As Long = (LVM_FIRST + 27)
Public Const LVM_INSERTCOLUMNW As Long = (LVM_FIRST + 97)
Public Const LVM_INSERTGROUP As Long = (LVM_FIRST + 145)
Public Const LVM_INSERTGROUPSORTED As Long = (LVM_FIRST + 159)
Public Const LVM_INSERTITEMA As Long = (LVM_FIRST + 7)
Public Const LVM_INSERTITEMW As Long = (LVM_FIRST + 77)
Public Const LVM_INSERTMARKHITTEST As Long = (LVM_FIRST + 168)
Public Const LVM_ISGROUPVIEWENABLED As Long = (LVM_FIRST + 175)
Public Const LVM_MOVEGROUP As Long = (LVM_FIRST + 151)
Public Const LVM_MOVEITEMTOGROUP As Long = (LVM_FIRST + 154)
Public Const LVM_REDRAWITEMS As Long = (LVM_FIRST + 21)
Public Const LVM_REMOVEALLGROUPS As Long = (LVM_FIRST + 160)
Public Const LVM_REMOVEGROUP As Long = (LVM_FIRST + 150)
Public Const LVM_SCROLL As Long = (LVM_FIRST + 20)
Public Const LVM_SETBKCOLOR As Long = (LVM_FIRST + 1)
Public Const LVM_SETBKIMAGEA As Long = (LVM_FIRST + 68)
Public Const LVM_SETBKIMAGEW As Long = (LVM_FIRST + 138)
Public Const LVM_SETCALLBACKMASK As Long = (LVM_FIRST + 11)
Public Const LVM_SETCOLUMNA As Long = (LVM_FIRST + 26)
Public Const LVM_SETCOLUMNORDERARRAY As Long = (LVM_FIRST + 58)
Public Const LVM_SETCOLUMNW As Long = (LVM_FIRST + 96)
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Public Const LVM_SETGROUPINFO As Long = (LVM_FIRST + 147)
Public Const LVM_SETGROUPMETRICS As Long = (LVM_FIRST + 155)
Public Const LVM_SETHOTCURSOR As Long = (LVM_FIRST + 62)
Public Const LVM_SETHOTITEM As Long = (LVM_FIRST + 60)
Public Const LVM_SETHOVERTIME As Long = (LVM_FIRST + 71)
Public Const LVM_SETICONSPACING As Long = (LVM_FIRST + 53)
Public Const LVM_SETIMAGELIST As Long = (LVM_FIRST + 3)
Public Const LVM_SETINFOTIP As Long = (LVM_FIRST + 173)
Public Const LVM_SETINSERTMARK As Long = (LVM_FIRST + 166)
Public Const LVM_SETINSERTMARKCOLOR As Long = (LVM_FIRST + 170)
Public Const LVM_SETITEMA As Long = (LVM_FIRST + 6)
Public Const LVM_SETITEMCOUNT As Long = (LVM_FIRST + 47)
Public Const LVM_SETITEMPOSITION As Long = (LVM_FIRST + 15)
Public Const LVM_SETITEMPOSITION32 As Long = (LVM_FIRST + 49)
Public Const LVM_SETITEMSTATE As Long = (LVM_FIRST + 43)
Public Const LVM_SETITEMTEXTA As Long = (LVM_FIRST + 46)
Public Const LVM_SETITEMTEXTW As Long = (LVM_FIRST + 116)
Public Const LVM_SETITEMW As Long = (LVM_FIRST + 76)
Public Const LVM_SETOUTLINECOLOR As Long = (LVM_FIRST + 177)
Public Const LVM_SETSELECTEDCOLUMN As Long = (LVM_FIRST + 140)
Public Const LVM_SETSELECTIONMARK As Long = (LVM_FIRST + 67)
Public Const LVM_SETTEXTBKCOLOR As Long = (LVM_FIRST + 38)
Public Const LVM_SETTEXTCOLOR As Long = (LVM_FIRST + 36)
Public Const LVM_SETTILEINFO As Long = (LVM_FIRST + 164)
Public Const LVM_SETTILEVIEWINFO As Long = (LVM_FIRST + 162)
Public Const LVM_SETTILEWIDTH As Long = (LVM_FIRST + 141)
Public Const LVM_SETTOOLTIPS As Long = (LVM_FIRST + 74)
Public Const LVM_SETVIEW As Long = (LVM_FIRST + 142)
Public Const LVM_SETWORKAREAS As Long = (LVM_FIRST + 65)
Public Const LVM_SORTGROUPS As Long = (LVM_FIRST + 158)
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
Public Const LVM_SORTITEMSEX As Long = (LVM_FIRST + 81)
Public Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
Public Const LVM_UPDATE As Long = (LVM_FIRST + 42)
Public Const LVM_APPROXIMATEVIEWRECT As Long = (LVM_FIRST + 64)
Public Const LVM_ARRANGE As Long = (LVM_FIRST + 22)
Public Const LVM_CREATEDRAGIMAGE As Long = (LVM_FIRST + 33)
Public Const LVM_DELETEALLITEMS As Long = (LVM_FIRST + 9)
Public Const LVM_DELETECOLUMN As Long = (LVM_FIRST + 28)
Public Const LVM_DELETEITEM As Long = (LVM_FIRST + 8)
Public Const LVM_EDITLABELA As Long = (LVM_FIRST + 23)
Public Const LVM_EDITLABELW As Long = (LVM_FIRST + 118)
Public Const LVM_ENABLEGROUPVIEW As Long = (LVM_FIRST + 157)
Public Const LVM_ENSUREVISIBLE As Long = (LVM_FIRST + 19)
Public Const LVM_FINDITEMA As Long = (LVM_FIRST + 13)
Public Const LVM_FINDITEMW As Long = (LVM_FIRST + 83)
Public Const LVM_GETBKCOLOR As Long = (LVM_FIRST + 0)
Public Const LVM_GETBKIMAGEA As Long = (LVM_FIRST + 69)
Public Const LVM_GETBKIMAGEW As Long = (LVM_FIRST + 139)
Public Const LVM_GETCALLBACKMASK As Long = (LVM_FIRST + 10)
Public Const LVM_GETCOLUMNA As Long = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMNORDERARRAY As Long = (LVM_FIRST + 59)
Public Const LVM_GETCOLUMNW As Long = (LVM_FIRST + 95)
Public Const LVM_GETCOLUMNWIDTH As Long = (LVM_FIRST + 29)
Public Const LVM_GETCOUNTPERPAGE As Long = (LVM_FIRST + 40)
Public Const LVM_GETEDITCONTROL As Long = (LVM_FIRST + 24)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Public Const LVM_GETGROUPINFO As Long = (LVM_FIRST + 149)
Public Const LVM_GETGROUPMETRICS As Long = (LVM_FIRST + 156)
Public Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Public Const LVM_GETHOTCURSOR As Long = (LVM_FIRST + 63)
Public Const LVM_GETHOTITEM As Long = (LVM_FIRST + 61)
Public Const LVM_GETHOVERTIME As Long = (LVM_FIRST + 72)
Public Const LVM_GETIMAGELIST As Long = (LVM_FIRST + 2)
Public Const LVM_GETINSERTMARK As Long = (LVM_FIRST + 167)
Public Const LVM_GETINSERTMARKCOLOR As Long = (LVM_FIRST + 171)
Public Const LVM_GETINSERTMARKRECT As Long = (LVM_FIRST + 169)
Public Const LVM_GETISEARCHSTRINGA As Long = (LVM_FIRST + 52)
Public Const LVM_GETISEARCHSTRINGW As Long = (LVM_FIRST + 117)

Public Const READ_CONTROL As Long = &H20000
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Public Const STANDARD_RIGHTS_EXECUTE As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const STANDARD_RIGHTS_WRITE As Long = (READ_CONTROL)

Public Const TOKEN_ADJUST_DEFAULT As Long = &H80
Public Const TOKEN_ADJUST_GROUPS As Long = &H40
Public Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Public Const TOKEN_ADJUST_SESSIONID As Long = &H100
Public Const TOKEN_AND As Long = &H3
Public Const TOKEN_ASSIGN_PRIMARY As Long = &H1
Public Const TOKEN_CLOSEPAREN As Long = &H5
Public Const TOKEN_DUPLICATE As Long = &H2
Public Const TOKEN_IMPERSONATE As Long = &H4
Public Const TOKEN_NOTIN As Long = &H20
Public Const TOKEN_OPENPAREN As Long = &H4
Public Const TOKEN_OPERATOR As Long = &H2
Public Const TOKEN_OR As Long = &H2
Public Const TOKEN_PAREN As Long = &H4
Public Const TOKEN_QUERY As Long = &H8
Public Const TOKEN_QUERY_SOURCE As Long = &H10
Public Const TOKEN_SOURCE_LENGTH As Long = 8
Public Const TOKEN_STRING_SIZE As Long = 4608
Public Const TOKEN_USER As Long = &H8
Public Const TOKEN_READ As Long = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)
Public Const TOKEN_EXECUTE As Long = STANDARD_RIGHTS_EXECUTE
Public Const TOKEN_WRITE As Long = (STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
Public Const TOKEN_ALL_ACCESS_P As Long = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
Public Const TOKEN_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_SESSIONID Or TOKEN_ADJUST_DEFAULT)

Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const CREATE_DEFAULT_ERROR_MODE As Long = &H4000000
Public Const CREATE_NEW_CONSOLE As Long = &H10
Public Const CREATE_NEW_PROCESS_GROUP As Long = &H200
Public Const CREATE_SEPARATE_WOW_VDM As Long = &H800
Public Const CREATE_SHARED_WOW_VDM As Long = &H1000
Public Const DETACHED_PROCESS As Long = &H8

Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Public Const FILE_ATTRIBUTE_DEVICE As Long = &H40
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_ENCRYPTED As Long = &H4000
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED As Long = &H2000
Public Const FILE_ATTRIBUTE_OFFLINE As Long = &H1000
Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400
Public Const FILE_ATTRIBUTE_SPARSE_FILE As Long = &H200
Public Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100

Public Const SHGFI_ADDOVERLAYS As Long = &H20
Public Const SHGFI_ATTR_SPECIFIED As Long = &H20000
Public Const SHGFI_ATTRIBUTES As Long = &H800
Public Const SHGFI_DISPLAYNAME As Long = &H200
Public Const SHGFI_EXETYPE As Long = &H2000
Public Const SHGFI_ICON As Long = &H100
Public Const SHGFI_ICONLOCATION As Long = &H1000
Public Const SHGFI_LARGEICON As Long = &H0
Public Const SHGFI_LINKOVERLAY As Long = &H8000
Public Const SHGFI_OPENICON As Long = &H2
Public Const SHGFI_OVERLAYINDEX As Long = &H40
Public Const SHGFI_PIDL As Long = &H8
Public Const SHGFI_SELECTED As Long = &H10000
Public Const SHGFI_SHELLICONSIZE As Long = &H4
Public Const SHGFI_SMALLICON As Long = &H1
Public Const SHGFI_SYSICONINDEX As Long = &H4000
Public Const SHGFI_TYPENAME As Long = &H400
Public Const SHGFI_USEFILEATTRIBUTES As Long = &H10

Public Const WINDING As Long = 2
Public Const RGN_DIFF As Long = 4

Public Const SM_ARRANGE As Long = 56
Public Const SM_CLEANBOOT As Long = 67
Public Const SM_CMETRICS As Long = 44
Public Const SM_CMONITORS As Long = 80
Public Const SM_CMOUSEBUTTONS As Long = 43
Public Const SM_CXBORDER As Long = 5
Public Const SM_CXCURSOR As Long = 13
Public Const SM_CXDLGFRAME As Long = 7
Public Const SM_CXDOUBLECLK As Long = 36
Public Const SM_CXDRAG As Long = 68
Public Const SM_CXEDGE As Long = 45
Public Const SM_CXFIXEDFRAME As Long = SM_CXDLGFRAME
Public Const SM_CXFRAME As Long = 32
Public Const SM_CXFULLSCREEN As Long = 16
Public Const SM_CXHSCROLL As Long = 21
Public Const SM_CXHTHUMB As Long = 10
Public Const SM_CXICON As Long = 11
Public Const SM_CXICONSPACING As Long = 38
Public Const SM_CXMAXIMIZED As Long = 61
Public Const SM_CXMAXTRACK As Long = 59
Public Const SM_CXMENUCHECK As Long = 71
Public Const SM_CXMENUSIZE As Long = 54
Public Const SM_CXMIN As Long = 28
Public Const SM_CXMINIMIZED As Long = 57
Public Const SM_CXMINSPACING As Long = 47
Public Const SM_CXMINTRACK As Long = 34
Public Const SM_CXSCREEN As Long = 0
Public Const SM_CXSIZE As Long = 30
Public Const SM_CXSIZEFRAME As Long = SM_CXFRAME
Public Const SM_CXSMICON As Long = 49
Public Const SM_CXSMSIZE As Long = 52
Public Const SM_CXVIRTUALSCREEN As Long = 78
Public Const SM_CXVSCROLL As Long = 2
Public Const SM_CYBORDER As Long = 6
Public Const SM_CYCAPTION As Long = 4
Public Const SM_CYCURSOR As Long = 14
Public Const SM_CYDLGFRAME As Long = 8
Public Const SM_CYDOUBLECLK As Long = 37
Public Const SM_CYDRAG As Long = 69
Public Const SM_CYEDGE As Long = 46
Public Const SM_CYFIXEDFRAME As Long = SM_CYDLGFRAME
Public Const SM_CYFRAME As Long = 33
Public Const SM_CYFULLSCREEN As Long = 17
Public Const SM_CYHSCROLL As Long = 3
Public Const SM_CYICON As Long = 12
Public Const SM_CYICONSPACING As Long = 39
Public Const SM_CYKANJIWINDOW As Long = 18
Public Const SM_CYMAXIMIZED As Long = 62
Public Const SM_CYMAXTRACK As Long = 60
Public Const SM_CYMENU As Long = 15
Public Const SM_CYMENUCHECK As Long = 72
Public Const SM_CYMENUSIZE As Long = 55
Public Const SM_CYMIN As Long = 29
Public Const SM_CYMINIMIZED As Long = 58
Public Const SM_CYMINSPACING As Long = 48
Public Const SM_CYMINTRACK As Long = 35
Public Const SM_CYSCREEN As Long = 1
Public Const SM_CYSIZE As Long = 31
Public Const SM_CYSIZEFRAME As Long = SM_CYFRAME
Public Const SM_CYSMCAPTION As Long = 51
Public Const SM_CYSMICON As Long = 50
Public Const SM_CYSMSIZE As Long = 53
Public Const SM_CYVIRTUALSCREEN As Long = 79
Public Const SM_CYVSCROLL As Long = 20
Public Const SM_CYVTHUMB As Long = 9
Public Const SM_DBCSENABLED As Long = 42
Public Const SM_DEBUG As Long = 22
Public Const SM_FOCUS_TYPE_LM_DOMAIN As Long = 2
Public Const SM_FOCUS_TYPE_LM_SERVER As Long = 5
Public Const SM_FOCUS_TYPE_NT_DOMAIN As Long = 1
Public Const SM_FOCUS_TYPE_NT_SERVER As Long = 4
Public Const SM_FOCUS_TYPE_UNKNOWN_DOMAIN As Long = 3
Public Const SM_FOCUS_TYPE_UNKNOWN_SERVER As Long = 7
Public Const SM_FOCUS_TYPE_WFW_SERVER As Long = 6
Public Const SM_IMMENABLED As Long = 82
Public Const SM_MENUDROPALIGNMENT As Long = 40
Public Const SM_MIDEASTENABLED As Long = 74
Public Const SM_MOUSEPRESENT As Long = 19
Public Const SM_MOUSEWHEELPRESENT As Long = 75
Public Const SM_NETWORK As Long = 63
Public Const SM_PENWINDOWS As Long = 41
Public Const SM_REMOTESESSION As Long = &H1000
Public Const SM_RESERVED1 As Long = 24
Public Const SM_RESERVED2 As Long = 25
Public Const SM_RESERVED3 As Long = 26
Public Const SM_RESERVED4 As Long = 27
Public Const SM_SAMEDISPLAYFORMAT As Long = 81
Public Const SM_SECURE As Long = 44
Public Const SM_SHOWSOUNDS As Long = 70
Public Const SM_SLOWMACHINE As Long = 73
Public Const SM_SWAPBUTTON As Long = 23
Public Const SM_XVIRTUALSCREEN As Long = 76
Public Const SM_YVIRTUALSCREEN As Long = 77

Public Const CSIDL_ADMINTOOLS As Long = &H30
Public Const CSIDL_ALTSTARTUP As Long = &H1D
Public Const CSIDL_APPDATA As Long = &H1A
Public Const CSIDL_BITBUCKET As Long = &HA
Public Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F
Public Const CSIDL_COMMON_ALTSTARTUP As Long = &H1E
Public Const CSIDL_COMMON_APPDATA As Long = &H23
Public Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19
Public Const CSIDL_COMMON_DOCUMENTS As Long = &H2E
Public Const CSIDL_COMMON_FAVORITES As Long = &H1F
Public Const CSIDL_COMMON_PROGRAMS As Long = &H17
Public Const CSIDL_COMMON_STARTMENU As Long = &H16
Public Const CSIDL_COMMON_STARTUP As Long = &H18
Public Const CSIDL_COMMON_TEMPLATES As Long = &H2D
Public Const CSIDL_CONNECTIONS As Long = &H31
Public Const CSIDL_CONTROLS As Long = &H3
Public Const CSIDL_COOKIES As Long = &H21
Public Const CSIDL_DESKTOP As Long = &H0
Public Const CSIDL_DESKTOPDIRECTORY As Long = &H10
Public Const CSIDL_DRIVES As Long = &H11
Public Const CSIDL_FAVORITES As Long = &H6
Public Const CSIDL_FLAG_CREATE As Long = &H8000
Public Const CSIDL_FLAG_DONT_VERIFY As Long = &H4000
Public Const CSIDL_FLAG_MASK As Long = &HFF00&
Public Const CSIDL_FLAG_PFTI_TRACKTARGET As Long = CSIDL_FLAG_DONT_VERIFY
Public Const CSIDL_FONTS As Long = &H14
Public Const CSIDL_HISTORY As Long = &H22
Public Const CSIDL_INTERNET As Long = &H1
Public Const CSIDL_INTERNET_CACHE As Long = &H20
Public Const CSIDL_LOCAL_APPDATA As Long = &H1C
Public Const CSIDL_MYPICTURES As Long = &H27
Public Const CSIDL_NETHOOD As Long = &H13
Public Const CSIDL_NETWORK As Long = &H12
Public Const CSIDL_PERSONAL As Long = &H5
Public Const CSIDL_PRINTERS As Long = &H4
Public Const CSIDL_PRINTHOOD As Long = &H1B
Public Const CSIDL_PROFILE As Long = &H28
Public Const CSIDL_PROGRAM_FILES As Long = &H26
Public Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B
Public Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C
Public Const CSIDL_PROGRAM_FILESX86 As Long = &H2A
Public Const CSIDL_PROGRAMS As Long = &H2
Public Const CSIDL_RECENT As Long = &H8
Public Const CSIDL_SENDTO As Long = &H9
Public Const CSIDL_STARTMENU As Long = &HB
Public Const CSIDL_STARTUP As Long = &H7
Public Const CSIDL_SYSTEM As Long = &H25
Public Const CSIDL_SYSTEMX86 As Long = &H29
Public Const CSIDL_TEMPLATES As Long = &H15
Public Const CSIDL_WINDOWS As Long = &H24

Public Const AW_SLIDE As Long = &H40000
Public Const AW_ACTIVATE As Long = &H20000

Public Const SPI_GETWORKAREA As Integer = 48

Public Const GW_CHILD As Long = 5
Public Const GW_HWNDNEXT As Long = 2

Public Const INPUT_MOUSE As Long = 0

Public Const MOUSEEVENTF_MOVE As Long = &H1
Public Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Public Const MOUSEEVENTF_LEFTUP As Long = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN As Long = &H20
Public Const MOUSEEVENTF_MIDDLEUP As Long = &H40
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10

Public Const SEM_FAILCRITICALERRORS = &H1
Public Const SEM_NOGPFAULTERRORBOX = &H2
Public Const SEM_NOOPENFILEERRORBOX = &H8000

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_ASYNCWINDOWPOS As Long = &H4000

Public Const PROCESS_TERMINATE As Long = (&H1)
Public Const SYNCHRONIZE As Long = &H100000

Public Const HWND_TOPMOST As Long = -1
Public Const HWND_TOP As Long = 0
Public Const HWND_NOTOPMOST As Long = -2
Public Const HWND_BOTTOM As Long = 1

Public Const IMAGE_BITMAP As Long = 0
Public Const IMAGE_ICON As Long = 1
Public Const BS_BITMAP As Long = &H80&
Public Const BS_ICON As Long = &H40&
Public Const BS_DEFPUSHBUTTON As Long = &H1&
Public Const BM_SETIMAGE As Long = &HF7&
Public Const GWL_HINSTANCE As Long = -6

Public Const WH_MOUSE As Long = 7
Public Const WH_CALLWNDPROC As Long = 4
Public Const WH_JOURNALRECORD As Long = 0
Public Const WH_MOUSE_LL As Long = 14
Public Const WH_KEYBOARD_LL As Long = 13

Public Const DT_CENTER             As Long = &H1
Public Const DT_WORDBREAK          As Long = &H10
Public Const DT_CALCRECT           As Long = &H400

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

'-------------costanti per il registry
Public Const HKEY_USERS As Long = &H80000003
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REGKEYSVGP_CT = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VGP3D"
Public Const REGKEYSVGP_AST = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VGP2D"
Public Const REGVALSUNINSTALL = "UninstallString"
'-------------

'HtmlHelp
Public Const HH_HELP_CONTEXT As Long = &HF&
Public Const HH_DISPLAY_TOPIC As Long = 0
Public Const HH_DISPLAY_TEXT_POPUP As Long = &HE

'CreateWindow
Public Const WS_ACTIVECAPTION As Long = &H1
Public Const WS_BORDER As Long = &H800000
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_CHILD As Long = &H40000000
Public Const WS_CLIPCHILDREN As Long = &H2000000
Public Const WS_CLIPSIBLINGS As Long = &H4000000
Public Const WS_DISABLED As Long = &H8000000
Public Const WS_DLGFRAME As Long = &H400000
Public Const WS_GROUP As Long = &H20000
Public Const WS_HSCROLL As Long = &H100000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_OVERLAPPED As Long = &H0&
Public Const WS_POPUP As Long = &H80000000
Public Const WS_SYSMENU As Long = &H80000
Public Const WS_TABSTOP As Long = &H10000
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_VSCROLL As Long = &H200000

Public Const WS_EX_ACCEPTFILES As Long = &H10&
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_CLIENTEDGE As Long = &H200&
Public Const WS_EX_CONTEXTHELP As Long = &H400&
Public Const WS_EX_CONTROLPARENT As Long = &H10000
Public Const WS_EX_DLGMODALFRAME As Long = &H1&
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_LAYOUTRTL As Long = &H400000
Public Const WS_EX_LEFT As Long = &H0&
Public Const WS_EX_LEFTSCROLLBAR As Long = &H4000&
Public Const WS_EX_LTRREADING As Long = &H0&
Public Const WS_EX_MDICHILD As Long = &H40&
Public Const WS_EX_NOACTIVATE As Long = &H8000000
Public Const WS_EX_NOINHERITLAYOUT As Long = &H100000
Public Const WS_EX_NOPARENTNOTIFY As Long = &H4&
Public Const WS_EX_WINDOWEDGE As Long = &H100&
Public Const WS_EX_TOOLWINDOW As Long = &H80&
Public Const WS_EX_RIGHT As Long = &H1000&
Public Const WS_EX_RIGHTSCROLLBAR As Long = &H0&
Public Const WS_EX_RTLREADING As Long = &H2000&
Public Const WS_EX_STATICEDGE As Long = &H20000
Public Const WS_EX_TOPMOST As Long = &H8&
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const WS_EX_OVERLAPPEDWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

'ShowWindow
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_FORCEMINIMIZE = 11

Public Const WM_ACTIVATE As Long = &H6
Public Const WM_ACTIVATEAPP As Long = &H1C
Public Const WM_AFXFIRST As Long = &H360
Public Const WM_AFXLAST As Long = &H37F
Public Const WM_APP As Long = &H8000
Public Const WM_APPCOMMAND As Long = &H319
Public Const WM_ASKCBFORMATNAME As Long = &H30C
Public Const WM_CANCELJOURNAL As Long = &H4B
Public Const WM_CANCELMODE As Long = &H1F
Public Const WM_CHANGECBCHAIN As Long = &H30D
Public Const WM_CHANGEUISTATE As Long = &H127
Public Const WM_CHAR As Long = &H102
Public Const WM_CHARTOITEM As Long = &H2F
Public Const WM_CHILDACTIVATE As Long = &H22
Public Const WM_CLEAR As Long = &H303
Public Const WM_CLOSE As Long = &H10
Public Const WM_COMMAND As Long = &H111
Public Const WM_COMMNOTIFY As Long = &H44
Public Const WM_COMPACTING As Long = &H41
Public Const WM_COMPAREITEM As Long = &H39
Public Const WM_CONTEXTMENU As Long = &H7B
Public Const WM_CONVERTREQUEST As Long = &H10A
Public Const WM_CONVERTREQUESTEX As Long = &H108
Public Const WM_CONVERTRESULT As Long = &H10B
Public Const WM_COPY As Long = &H301
Public Const WM_COPYDATA As Long = &H4A
Public Const WM_CREATE As Long = &H1
Public Const WM_CTLCOLOR As Long = &H19
Public Const WM_CTLCOLORBTN As Long = &H135
Public Const WM_CTLCOLORDLG As Long = &H136
Public Const WM_CTLCOLOREDIT As Long = &H133
Public Const WM_CTLCOLORLISTBOX As Long = &H134
Public Const WM_CTLCOLORMSGBOX As Long = &H132
Public Const WM_CTLCOLORSCROLLBAR As Long = &H137
Public Const WM_CTLCOLORSTATIC As Long = &H138
Public Const WM_CUT As Long = &H300
Public Const WM_DEADCHAR As Long = &H103
Public Const WM_DELETEITEM As Long = &H2D
Public Const WM_DESTROY As Long = &H2
Public Const WM_DESTROYCLIPBOARD As Long = &H307
Public Const WM_DEVICECHANGE As Long = &H219
Public Const WM_DEVMODECHANGE As Long = &H1B
Public Const WM_DISPLAYCHANGE As Long = &H7E
Public Const WM_DRAWCLIPBOARD As Long = &H308
Public Const WM_DRAWITEM As Long = &H2B
Public Const WM_DROPFILES As Long = &H233
Public Const WM_ENABLE As Long = &HA
Public Const WM_ENDSESSION As Long = &H16
Public Const WM_ENTERIDLE As Long = &H121
Public Const WM_ENTERMENULOOP As Long = &H211
Public Const WM_ENTERSIZEMOVE As Long = &H231
Public Const WM_ERASEBKGND As Long = &H14
Public Const WM_EXITMENULOOP As Long = &H212
Public Const WM_EXITSIZEMOVE As Long = &H232
Public Const WM_FONTCHANGE As Long = &H1D
Public Const WM_FORWARDMSG As Long = &H37F
Public Const WM_GETDLGCODE As Long = &H87
Public Const WM_GETFONT As Long = &H31
Public Const WM_GETHOTKEY As Long = &H33
Public Const WM_GETICON As Long = &H7F
Public Const WM_GETMINMAXINFO As Long = &H24
Public Const WM_GETOBJECT As Long = &H3D
Public Const WM_GETTEXT As Long = &HD
Public Const WM_GETTEXTLENGTH As Long = &HE
Public Const WM_HANDHELDFIRST As Long = &H358
Public Const WM_HANDHELDLAST As Long = &H35F
Public Const WM_HELP As Long = &H53
Public Const WM_HOTKEY As Long = &H312
Public Const WM_HSCROLL As Long = &H114
Public Const WM_HSCROLLCLIPBOARD As Long = &H30E
Public Const WM_ICONERASEBKGND As Long = &H27
Public Const WM_IME_CHAR As Long = &H286
Public Const WM_IME_COMPOSITION As Long = &H10F
Public Const WM_IME_COMPOSITIONFULL As Long = &H284
Public Const WM_IME_CONTROL As Long = &H283
Public Const WM_IME_ENDCOMPOSITION As Long = &H10E
Public Const WM_IME_KEYDOWN As Long = &H290
Public Const WM_IME_KEYLAST As Long = &H10F
Public Const WM_IME_KEYUP As Long = &H291
Public Const WM_IME_NOTIFY As Long = &H282
Public Const WM_IME_REPORT As Long = &H280
Public Const WM_IME_REQUEST As Long = &H288
Public Const WM_IME_SELECT As Long = &H285
Public Const WM_IME_SETCONTEXT As Long = &H281
Public Const WM_IME_STARTCOMPOSITION As Long = &H10D
Public Const WM_IMEKEYDOWN As Long = &H290
Public Const WM_IMEKEYUP As Long = &H291
Public Const WM_INITDIALOG As Long = &H110
Public Const WM_INITMENU As Long = &H116
Public Const WM_INITMENUPOPUP As Long = &H117
Public Const WM_INPUTLANGCHANGE As Long = &H51
Public Const WM_INPUTLANGCHANGEREQUEST As Long = &H50
Public Const WM_INTERIM As Long = &H10C
Public Const WM_KEYDOWN As Long = &H100
Public Const WM_KEYFIRST As Long = &H100
Public Const WM_KEYLAST As Long = &H108
Public Const WM_KEYUP As Long = &H101
Public Const WM_KILLFOCUS As Long = &H8
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_MBUTTONDBLCLK As Long = &H209
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MDIACTIVATE As Long = &H222
Public Const WM_MDICASCADE As Long = &H227
Public Const WM_MDICREATE As Long = &H220
Public Const WM_MDIDESTROY As Long = &H221
Public Const WM_MDIGETACTIVE As Long = &H229
Public Const WM_MDIICONARRANGE As Long = &H228
Public Const WM_MDIMAXIMIZE As Long = &H225
Public Const WM_MDINEXT As Long = &H224
Public Const WM_MDIREFRESHMENU As Long = &H234
Public Const WM_MDIRESTORE As Long = &H223
Public Const WM_MDISETMENU As Long = &H230
Public Const WM_MDITILE As Long = &H226
Public Const WM_MEASUREITEM As Long = &H2C
Public Const WM_MENUCHAR As Long = &H120
Public Const WM_MENUCOMMAND As Long = &H126
Public Const WM_MENUDRAG As Long = &H123
Public Const WM_MENUGETOBJECT As Long = &H124
Public Const WM_MENURBUTTONUP As Long = &H122
Public Const WM_MENUSELECT As Long = &H11F
Public Const WM_MOUSEACTIVATE As Long = &H21
Public Const WM_MOUSEFIRST As Long = &H200
Public Const WM_MOUSEHOVER As Long = &H2A1
Public Const WM_MOUSELAST As Long = &H209
Public Const WM_MOUSELEAVE As Long = &H2A3
Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_MOUSEWHEEL As Long = &H20A
Public Const WM_MOVE As Long = &H3
Public Const WM_MOVING As Long = &H216
Public Const WM_NCACTIVATE As Long = &H86
Public Const WM_NCCALCSIZE As Long = &H83
Public Const WM_NCCREATE As Long = &H81
Public Const WM_NCDESTROY As Long = &H82
Public Const WM_NCHITTEST As Long = &H84
Public Const WM_NCLBUTTONDBLCLK As Long = &HA3
Public Const WM_NCLBUTTONDOWN As Long = &HA1
Public Const WM_NCLBUTTONUP As Long = &HA2
Public Const WM_NCMBUTTONDBLCLK As Long = &HA9
Public Const WM_NCMBUTTONDOWN As Long = &HA7
Public Const WM_NCMBUTTONUP As Long = &HA8
Public Const WM_NCMOUSEHOVER As Long = &H2A0
Public Const WM_NCMOUSELEAVE As Long = &H2A2
Public Const WM_NCMOUSEMOVE As Long = &HA0
Public Const WM_NCPAINT As Long = &H85
Public Const WM_NCRBUTTONDBLCLK As Long = &HA6
Public Const WM_NCRBUTTONDOWN As Long = &HA4
Public Const WM_NCRBUTTONUP As Long = &HA5
Public Const WM_NCXBUTTONDBLCLK As Long = &HAD
Public Const WM_NCXBUTTONDOWN As Long = &HAB
Public Const WM_NCXBUTTONUP As Long = &HAC
Public Const WM_NEXTDLGCTL As Long = &H28
Public Const WM_NEXTMENU As Long = &H213
Public Const WM_NOTIFY As Long = &H4E
Public Const WM_NOTIFYFORMAT As Long = &H55
Public Const WM_NULL As Long = &H0
Public Const WM_OTHERWINDOWCREATED As Long = &H42
Public Const WM_OTHERWINDOWDESTROYED As Long = &H43
Public Const WM_PAINT As Long = &HF&
Public Const WM_PAINTCLIPBOARD As Long = &H309
Public Const WM_PAINTICON As Long = &H26
Public Const WM_PALETTECHANGED As Long = &H311
Public Const WM_PALETTEISCHANGING As Long = &H310
Public Const WM_PARENTNOTIFY As Long = &H210
Public Const WM_PASTE As Long = &H302
Public Const WM_PENWINFIRST As Long = &H380
Public Const WM_PENWINLAST As Long = &H38F
Public Const WM_POWER As Long = &H48
Public Const WM_POWERBROADCAST As Long = &H218
Public Const WM_PRINT As Long = &H317
Public Const WM_PRINTCLIENT As Long = &H318
Public Const WM_QUERYDRAGICON As Long = &H37
Public Const WM_QUERYENDSESSION As Long = &H11
Public Const WM_QUERYNEWPALETTE As Long = &H30F
Public Const WM_QUERYOPEN As Long = &H13
Public Const WM_QUERYUISTATE As Long = &H129
Public Const WM_QUEUESYNC As Long = &H23
Public Const WM_QUIT As Long = &H12
Public Const WM_RASDIALEVENT As Long = &HCCCD
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RENDERALLFORMATS As Long = &H306
Public Const WM_RENDERFORMAT As Long = &H305
Public Const WM_SETCURSOR As Long = &H20
Public Const WM_SETFOCUS As Long = &H7
Public Const WM_SETFONT As Long = &H30
Public Const WM_SETHOTKEY As Long = &H32
Public Const WM_SETICON As Long = &H80
Public Const WM_SETREDRAW As Long = &HB
Public Const WM_SETTEXT As Long = &HC
Public Const WM_SHOWWINDOW As Long = &H18
Public Const WM_SIZE As Long = &H5
Public Const WM_SIZECLIPBOARD As Long = &H30B
Public Const WM_SIZING As Long = &H214
Public Const WM_SPOOLERSTATUS As Long = &H2A
Public Const WM_STYLECHANGED As Long = &H7D
Public Const WM_STYLECHANGING As Long = &H7C
Public Const WM_SYNCPAINT As Long = &H88
Public Const WM_SYSCHAR As Long = &H106
Public Const WM_SYSCOLORCHANGE As Long = &H15
Public Const WM_SYSCOMMAND As Long = &H112
Public Const WM_SYSDEADCHAR As Long = &H107
Public Const WM_SYSKEYDOWN As Long = &H104
Public Const WM_SYSKEYUP As Long = &H105
Public Const WM_TCARD As Long = &H52
Public Const WM_TIMECHANGE As Long = &H1E
Public Const WM_TIMER As Long = &H113
Public Const WM_UNDO As Long = &H304
Public Const WM_UNINITMENUPOPUP As Long = &H125
Public Const WM_UPDATEUISTATE As Long = &H128
Public Const WM_USER As Long = &H400
Public Const WM_USERCHANGED As Long = &H54
Public Const WM_VKEYTOITEM As Long = &H2E
Public Const WM_VSCROLL As Long = &H115
Public Const WM_VSCROLLCLIPBOARD As Long = &H30A
Public Const WM_WINDOWPOSCHANGED As Long = &H47
Public Const WM_WINDOWPOSCHANGING As Long = &H46
Public Const WM_WININICHANGE As Long = &H1A
Public Const WM_WNT_CONVERTREQUESTEX As Long = &H109
Public Const WM_XBUTTONDBLCLK As Long = &H20D
Public Const WM_XBUTTONDOWN As Long = &H20B
Public Const WM_XBUTTONUP As Long = &H20C

Public Const BM_CLICK As Long = &HF5&

Public Const F0_DELETE = &H3
Public Const F0F_ALLOWUNDO = &H40
Public Const F0F_CREATEPROGRESSDLG As Long = &H0
Public Const FOF_NOCONFIRMATION As Long = &H10

Public Const GWL_STYLE As Long = (-16)
Public Const GWL_WNDPROC As Long = (-4)
Public Const GWL_EXSTYLE As Long = (-20)
Public Const LWA_ALPHA As Long = &H2

Public Const ICC_USEREX_CLASSES As Long = &H200

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const VK_ESCAPE As Long = &H1B
Public Const VK_ACCEPT As Long = &H1E
Public Const VK_ADD As Long = &H6B
Public Const VK_APPS As Long = &H5D
Public Const VK_ATTN As Long = &HF6&
Public Const VK_BACK As Long = &H8
Public Const VK_BROWSER_BACK As Long = &HA6
Public Const VK_BROWSER_FAVORITES As Long = &HAB
Public Const VK_BROWSER_FORWARD As Long = &HA7
Public Const VK_BROWSER_HOME As Long = &HAC
Public Const VK_BROWSER_REFRESH As Long = &HA8
Public Const VK_BROWSER_SEARCH As Long = &HAA
Public Const VK_BROWSER_STOP As Long = &HA9
Public Const VK_CANCEL As Long = &H3
Public Const VK_CAPITAL As Long = &H14
Public Const VK_CLEAR As Long = &HC
Public Const VK_CONTROL As Long = &H11
Public Const VK_CONVERT As Long = &H1C
Public Const VK_CRSEL As Long = &HF7&
Public Const VK_DBE_ALPHANUMERIC As Long = &HF0
Public Const VK_DBE_CODEINPUT As Long = &HFA
Public Const VK_DBE_DBCSCHAR As Long = &HF4
Public Const VK_DBE_DETERMINESTRING As Long = &HFC
Public Const VK_DBE_ENTERDLGCONVERSIONMODE As Long = &HFD
Public Const VK_DBE_ENTERIMECONFIGMODE As Long = &HF8
Public Const VK_DBE_ENTERWORDREGISTERMODE As Long = &HF7
Public Const VK_DBE_FLUSHSTRING As Long = &HF9
Public Const VK_DBE_HIRAGANA As Long = &HF2
Public Const VK_DBE_KATAKANA As Long = &HF1
Public Const VK_DBE_NOCODEINPUT As Long = &HFB
Public Const VK_DBE_NOROMAN As Long = &HF6
Public Const VK_DBE_ROMAN As Long = &HF5
Public Const VK_DBE_SBCSCHAR As Long = &HF3
Public Const VK_DECIMAL As Long = &H6E
Public Const VK_DELETE As Long = &H2E
Public Const VK_DIVIDE As Long = &H6F
Public Const VK_DOWN As Long = &H28
Public Const VK_END As Long = &H23
Public Const VK_EREOF As Long = &HF9&
Public Const VK_EXECUTE As Long = &H2B
Public Const VK_EXSEL As Long = &HF8&
Public Const VK_F1 As Long = &H70
Public Const VK_F10 As Long = &H79
Public Const VK_F11 As Long = &H7A
Public Const VK_F12 As Long = &H7B
Public Const VK_F13 As Long = &H7C
Public Const VK_F14 As Long = &H7D
Public Const VK_F15 As Long = &H7E
Public Const VK_F16 As Long = &H7F
Public Const VK_F17 As Long = &H80
Public Const VK_F18 As Long = &H81
Public Const VK_F19 As Long = &H82
Public Const VK_F2 As Long = &H71
Public Const VK_F20 As Long = &H83
Public Const VK_F21 As Long = &H84
Public Const VK_F22 As Long = &H85
Public Const VK_F23 As Long = &H86
Public Const VK_F24 As Long = &H87
Public Const VK_F3 As Long = &H72
Public Const VK_F4 As Long = &H73
Public Const VK_F5 As Long = &H74
Public Const VK_F6 As Long = &H75
Public Const VK_F7 As Long = &H76
Public Const VK_F8 As Long = &H77
Public Const VK_F9 As Long = &H78
Public Const VK_FINAL As Long = &H18
Public Const VK_HANGEUL As Long = &H15
Public Const VK_HANGUL As Long = &H15
Public Const VK_HANJA As Long = &H19
Public Const VK_HELP As Long = &H2F
Public Const VK_HOME As Long = &H24
Public Const VK_ICO_00 As Long = &HE4
Public Const VK_ICO_CLEAR As Long = &HE6
Public Const VK_ICO_HELP As Long = &HE3
Public Const VK_INSERT As Long = &H2D
Public Const VK_JUNJA As Long = &H17
Public Const VK_KANA As Long = &H15
Public Const VK_KANJI As Long = &H19
Public Const VK_LAUNCH_APP1 As Long = &HB6
Public Const VK_LAUNCH_APP2 As Long = &HB7
Public Const VK_LAUNCH_MAIL As Long = &HB4
Public Const VK_LAUNCH_MEDIA_SELECT As Long = &HB5
Public Const VK_LBUTTON As Long = &H1
Public Const VK_LCONTROL As Long = &HA2
Public Const VK_LEFT As Long = &H25
Public Const VK_LMENU As Long = &HA4
Public Const VK_LSHIFT As Long = &HA0
Public Const VK_LWIN As Long = &H5B
Public Const VK_MBUTTON As Long = &H4
Public Const VK_MEDIA_NEXT_TRACK As Long = &HB0
Public Const VK_MEDIA_PLAY_PAUSE As Long = &HB3
Public Const VK_MEDIA_PREV_TRACK As Long = &HB1
Public Const VK_MEDIA_STOP As Long = &HB2
Public Const VK_MENU As Long = &H12
Public Const VK_MODECHANGE As Long = &H1F
Public Const VK_MULTIPLY As Long = &H6A
Public Const VK_NEXT As Long = &H22
Public Const VK_NONAME As Long = &HFC&
Public Const VK_NONCONVERT As Long = &H1D
Public Const VK_NUMLOCK As Long = &H90
Public Const VK_NUMPAD0 As Long = &H60
Public Const VK_NUMPAD1 As Long = &H61
Public Const VK_NUMPAD2 As Long = &H62
Public Const VK_NUMPAD3 As Long = &H63
Public Const VK_NUMPAD4 As Long = &H64
Public Const VK_NUMPAD5 As Long = &H65
Public Const VK_NUMPAD6 As Long = &H66
Public Const VK_NUMPAD7 As Long = &H67
Public Const VK_NUMPAD8 As Long = &H68
Public Const VK_NUMPAD9 As Long = &H69
Public Const VK_OEM_1 As Long = &HBA
Public Const VK_OEM_102 As Long = &HE2
Public Const VK_OEM_2 As Long = &HBF
Public Const VK_OEM_3 As Long = &HC0
Public Const VK_OEM_4 As Long = &HDB
Public Const VK_OEM_5 As Long = &HDC
Public Const VK_OEM_6 As Long = &HDD
Public Const VK_OEM_7 As Long = &HDE
Public Const VK_OEM_8 As Long = &HDF
Public Const VK_OEM_ATTN As Long = &HF0&
Public Const VK_OEM_AUTO As Long = &HF3&
Public Const VK_OEM_AX As Long = &HE1
Public Const VK_OEM_BACKTAB As Long = &HF5&
Public Const VK_OEM_CLEAR As Long = &HFE&
Public Const VK_OEM_COMMA As Long = &HBC
Public Const VK_OEM_COPY As Long = &HF2&
Public Const VK_OEM_CUSEL As Long = &HEF
Public Const VK_OEM_ENLW As Long = &HF4&
Public Const VK_OEM_FINISH As Long = &HF1&
Public Const VK_OEM_FJ_JISHO As Long = &H92
Public Const VK_OEM_FJ_LOYA As Long = &H95
Public Const VK_OEM_FJ_MASSHOU As Long = &H93
Public Const VK_OEM_FJ_ROYA As Long = &H96
Public Const VK_OEM_FJ_TOUROKU As Long = &H94
Public Const VK_OEM_JUMP As Long = &HEA
Public Const VK_OEM_MINUS As Long = &HBD
Public Const VK_OEM_NEC_EQUAL As Long = &H92
Public Const VK_OEM_PA1 As Long = &HEB
Public Const VK_OEM_PA2 As Long = &HEC
Public Const VK_OEM_PA3 As Long = &HED
Public Const VK_OEM_PERIOD As Long = &HBE
Public Const VK_OEM_PLUS As Long = &HBB
Public Const VK_OEM_RESET As Long = &HE9
Public Const VK_OEM_WSCTRL As Long = &HEE
Public Const VK_PA1 As Long = &HFD&
Public Const VK_PACKET As Long = &HE7
Public Const VK_PAUSE As Long = &H13
Public Const VK_PLAY As Long = &HFA&
Public Const VK_PRINT As Long = &H2A
Public Const VK_PRIOR As Long = &H21
Public Const VK_PROCESSKEY As Long = &HE5
Public Const VK_RBUTTON As Long = &H2
Public Const VK_RCONTROL As Long = &HA3
Public Const VK_RETURN As Long = &HD
Public Const VK_RIGHT As Long = &H27
Public Const VK_RMENU As Long = &HA5
Public Const VK_RSHIFT As Long = &HA1
Public Const VK_RWIN As Long = &H5C
Public Const VK_SCROLL As Long = &H91
Public Const VK_SELECT As Long = &H29
Public Const VK_SEPARATOR As Long = &H6C
Public Const VK_SHIFT As Long = &H10
Public Const VK_SLEEP As Long = &H5F
Public Const VK_SNAPSHOT As Long = &H2C
Public Const VK_SPACE As Long = &H20
Public Const VK_SUBTRACT As Long = &H6D
Public Const VK_TAB As Long = &H9
Public Const VK_UP As Long = &H26
Public Const VK_VOLUME_DOWN As Long = &HAE
Public Const VK_VOLUME_MUTE As Long = &HAD
Public Const VK_VOLUME_UP As Long = &HAF
Public Const VK_XBUTTON1 As Long = &H5
Public Const VK_XBUTTON2 As Long = &H6
Public Const VK_ZOOM As Long = &HFB&

Public Const MAX_PATH = 260

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------

Public Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Public Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Public Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Type ConvertPOINTAPI  ' Used by WM_SYSCOMMAND - converts mouse location.
    xy As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type

'per la funzione SHFileOperation
Public Type SHFILEOPSTRUCT
   hwnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Integer
   fAnyOperationsAborted As Boolean
   hNameMappings As Long
   lpszProgressTitle As String
End Type

'Per la funzione print-screen (screen-shot)
Public Type bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

Public Type HH_POPUP
    cbStruct As Long
    hinst As Long
    idString As Long
    pszText As String
    pt As POINTAPI
    clrForeground As Long
    clrBackground As Long
    rcMargins As RECT
    pszFont As String
End Type

Public Type EVENTMSG
    message As Long
    paramL As Long
    paramH As Long
    time As Long
    hwnd As Long
End Type

Public Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Public Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public nid As NOTIFYICONDATA

Public Type tMOUSEINPUT
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Type tINPUT
    type As Long
    mi As tMOUSEINPUT
End Type

Public Type ShortItemId
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As ShortItemId
End Type

Public Type SHFILEINFO
    hIcon As Long ' : icon
    iIcon As Long ' : icondex
    dwAttributes As Long ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80 ' : type name
End Type

Public Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------

Public Declare Function GetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
Public Declare Function CloseWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long                                                                                                                           'Esempio: SetWindowPos FormTest.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRECT As RECT) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRECT As RECT) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function UnLoadLibrary Lib "kernel32" Alias "FreeLibrary" (ByVal hDLL As OLE_HANDLE) As Long
Public Declare Function GetDLLHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long
Public Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal Hkey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal strClassName As String, ByVal lpWindowName As Any) As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Integer) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Integer, ByVal dwData As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetProcessVersion Lib "kernel32.dll" (ByVal ProcessId As Long) As Long
Public Declare Function GetWindowContextHelpId Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ChildWindowFromPoint Lib "user32.dll" (ByVal hWndParent As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef iccex As tagInitCommonControlsEx) As Boolean
Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetMessageExtraInfo Lib "user32.dll" () As Long
Public Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, ByRef lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRECT As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal N As Long, lpRECT As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
Public Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function SendInput Lib "user32.dll" (ByVal cInputs As Long, ByRef pInputs As tINPUT, ByVal cbSize As Long) As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function VkKeyScanEx Lib "user32.dll" Alias "VkKeyScanExA" (ByVal ch As Byte, ByVal dwhkl As Long) As Integer
Public Declare Function VkKeyScan Lib "user32.dll" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Public Declare Function MapVirtualKey Lib "user32.dll" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function OemKeyScan Lib "user32.dll" (ByVal wOemChar As Long) As Long
Public Declare Function GetFocus Lib "user32.dll" () As Long
Public Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function AttachThreadInput Lib "user32.dll" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function AnimateWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
Public Declare Function BringWindowToTop Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Public Declare Function LoadCursorFromFile Lib "user32.dll" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetCursor Lib "user32.dll" () As Long
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function CreateEllipticRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreatePolygonRgn Lib "GDI32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal CombineMode As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, ipic As IPicture) As Long
Public Declare Sub CoCreateInstance Lib "ole32.dll" (ByVal rclsid As Long, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByVal riid As Long, ByRef ppv As Any)
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SelectPalette Lib "GDI32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function GetLastActivePopup Lib "user32.dll" (ByVal hwndOwnder As Long) As Long
Public Declare Function ToAscii Lib "user32.dll" (ByVal uVirtKey As Long, ByVal uScanCode As Long, ByRef lpbKeyState As Byte, ByRef lpwTransKey As Long, ByVal fuState As Long) As Long
Public Declare Function ToAsciiEx Lib "user32.dll" (ByVal uVirtKey As Long, ByVal uScanCode As Long, ByRef lpKeyState As Byte, ByRef lpChar As Long, ByVal uFlags As Long, ByVal dwhkl As Long) As Long
Public Declare Function GetKeyboardState Lib "user32.dll" (ByRef pbKeyState As Byte) As Long
Public Declare Function GetKeyboardLayout Lib "user32.dll" (ByVal dwLayout As Long) As Long
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpAddress As Any, ByRef dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualAlloc Lib "kernel32.dll" (ByRef lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Public Declare Function CreateProcessWithLogonW Lib "advapi32.dll" (ByVal lpUsername As String, ByVal lpDomain As String, ByVal lpPassword As String, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32.dll" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function SetKeyboardState Lib "user32.dll" (ByRef lppbKeyState As Byte) As Long
Public Declare Sub GetKeyboardStateByString Lib "user32" Alias "GetKeyboardState" (ByVal pbKeyState As String)
Public Declare Sub SetKeyboardStateByString Lib "user32" Alias "SetKeyboardState" (ByVal lppbKeyState As String)
