Attribute VB_Name = "mTextUTF"
Option Explicit
'https://www.cnblogs.com/shenhaocn/archive/2011/10/23/2221572.html
'mTextUTF.bas
'Ä£¿é£ºUTFÎÄ±¾ÎÄ¼ş·ÃÎÊ
'×÷Õß£ºzyl910
'°æ±¾£º1.0
'ÈÕÆÚ£º2006-1-23

'ĞŞ¸Ä: shenhao(shenhaocn@qq.com)
'°æ±¾: 1.1
'ÈÕÆÚ: 2011-10-23

'== ËµÃ÷ ===================================================
'Ôö¼ÓÖ§³ÖUTF-8µÄÎŞBOM¸ñÊ½±àÂë shenhao
'Ö§³ÖUnicode±àÂëµÄÎÄ±¾ÎÄ¼ş¶ÁĞ´¡£ÔİÊ±Ö§³ÖANSI¡¢UTF-8¡¢UTF-16LE¡¢UTF-16BEÕâ¼¸ÖÖ±àÂëÎÄ±¾

'== ¸üĞÂ¼ÇÂ¼ ===============================================
'[V1.0] 2006-1-23
'1.Ö§³Ö×î³£¼ûµÄANSI¡¢UTF-8¡¢UTF-16LE¡¢UTF-16BEÕâ¼¸ÖÖ±àÂëÎÄ±¾
'
'[V1.1] 2011-10-23 by shenhao (shenhaocn@qq.com)
'1.Ö§³ÖUTF-8µÄÎŞBOM¸ñÊ½±àÂë

'## ±àÒëÔ¤´¦Àí³£Êı #########################################
'== È«¾Ö³£Êı ===============================================
'IncludeAPILib£ºÒıÓÃÁËAPI¿â£¬´ËÊ±²»ĞèÒªÊÖ¶¯Ğ´APIÉùÃ÷

'## API ####################################################
#If IncludeAPILib = 0 Then
'== File ===================================================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Const INVALID_HANDLE_VALUE = -1

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const Create_NEW = 1
Private Const Create_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const TRUNCATE_EXISTING = 5

Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

'== Unicode ================================================

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 As Long = 65001

#End If

'###########################################################

'Unicode±àÂë¸ñÊ½
Public Enum UnicodeEncodeFormat
UEF_ANSI = 0 'ANSI+DBCS
UEF_UTF8     'UTF-8
uef_utf8NB   'UTF-8 No BOM
UEF_UTF16LE  'UTF-16LE
UEF_UTF16BE  'UTF-16BE
UEF_UTF32LE  'UTF-32LE
UEF_UTF32BE  'UTF-32BE

UEF_AUTO = -1 '×Ô¶¯Ê¶±ğ±àÂë

'Òş²ØÏîÄ¿
[_UEF_Min] = UEF_ANSI
[_UEF_Max] = UEF_UTF32BE

End Enum

'ANSI+DBCS·½Ê½µÄÎÄ±¾ËùÊ¹ÓÃµÄ´úÂëÒ³¡£Ä¬ÈÏÎª0£¬±íÊ¾Ê¹ÓÃÏµÍ³µ±Ç°´úÂëÒ³¡£
'¿ÉÒÔÀûÓÃ¸Ã²ÎÊıÊµÏÖ¶ÁÈ¡ÆäËû´úÂë±àÂëµÄÎÄ±¾£¬±ÈÈçÏëÔÚ ¼òÌåÖĞÎÄÆ½Ì¨ÏÂ ¶ÁÈ¡ ·±ÌåÖĞÎÄÆ½Ì¨Éú³ÉµÄtxt£¬¾Í½«ËüÉèÎª950
Public UEFCodePage As Long


'°ÑBYTEÀàĞÍ±äÁ¿×óÒÆ1Î»µÄº¯Êı
'·µ»ØÖµ£ºÒÆÎ»½á¹û
'Byt£º´ıÒÆÎ»µÄ×Ö½Ú
Private Function ShLB_By1Bit(ByVal Byt As Byte) As Byte
    
    '(Byt And &H7F)µÄ×÷ÓÃÊÇÆÁ±Î×î¸ßÎ»¡£ *2£º×óÒÆÒ»Î»
    ShLB_By1Bit = (Byt And &H7F) * 2

End Function

'ÅĞ¶ÏBOM
'·µ»ØÖµ£ºBOMËùÕ¼×Ö½Ú
'dwFirst£º[in]ÎÄ¼ş×î¿ªÊ¼µÄ4¸ö×Ö½Ú
'fmt£º[out]·µ»Ø±àÂëÀàĞÍ
Public Function UEFCheckBOM(ByVal dwFirst As Long, ByRef fmt As UnicodeEncodeFormat) As Long
    If dwFirst = &HFEFF& Then
        fmt = UEF_UTF32LE
        UEFCheckBOM = 4
    ElseIf dwFirst = &HFFFE0000 Then
        fmt = UEF_UTF32BE
        UEFCheckBOM = 4
    ElseIf (dwFirst And &HFFFF&) = &HFEFF& Then
        fmt = UEF_UTF16LE
        UEFCheckBOM = 2
    ElseIf (dwFirst And &HFFFF&) = &HFFFE& Then
        fmt = UEF_UTF16BE
        UEFCheckBOM = 2
    ElseIf (dwFirst And &HFFFFFF) = &HBFBBEF Then
        fmt = UEF_UTF8
        UEFCheckBOM = 3
    Else 'ÏÈÔİ¶¨ÎªUEF_ANSI ºóĞø»áÔÙÇø·ÖUEF_ANSI ºÍ UEF_UTF8NB
        fmt = UEF_ANSI
        UEFCheckBOM = 0
    End If
End Function

'==========================================================================================
'UTF-8×î´óµÄÒ»¸öÌØµã£¬¾ÍÊÇËüÊÇÒ»ÖÖ±ä³¤µÄ±àÂë·½Ê½¡£
'Ëü¿ÉÒÔÊ¹ÓÃ1~4¸ö×Ö½Ú±íÊ¾Ò»¸ö·ûºÅ£¬¸ù¾İ²»Í¬µÄ·ûºÅ¶ø±ä»¯×Ö½Ú³¤¶È¡£
'ÿ
'UTF-8µÄ±àÂë¹æÔòºÜ¼òµ¥£¬Ö»ÓĞ¶şÌõ£º
'ÿ
'1£©¶ÔÓÚµ¥×Ö½ÚµÄ·ûºÅ£¬×Ö½ÚµÄµÚÒ»Î»ÉèÎª0£¬ºóÃæ7Î»ÎªÕâ¸ö·ûºÅµÄunicodeÂë¡£
'   Òò´Ë¶ÔÓÚÓ¢Óï×ÖÄ¸£¬UTF-8±àÂëºÍASCIIÂëÊÇÏàÍ¬µÄ¡£
'ÿ
'2£©¶ÔÓÚn×Ö½ÚµÄ·ûºÅ£¨n>1£©£¬µÚÒ»¸ö×Ö½ÚµÄÇ°nÎ»¶¼ÉèÎª1£¬µÚn+1Î»ÉèÎª0£¬ºóÃæ×Ö½ÚµÄÇ°Á½Î»Ò»ÂÉÉèÎª10¡£
'   Ê£ÏÂµÄÃ»ÓĞÌá¼°µÄ¶ş½øÖÆÎ»£¬È«²¿ÎªÕâ¸ö·ûºÅµÄunicodeÂë¡£
'ÿ
'ÏÂ±í×Ü½áÁË±àÂë¹æÔò£¬×ÖÄ¸x±íÊ¾¿ÉÓÃ±àÂëµÄÎ»¡£
'Unicode·ûºÅ·¶Î§     | UTF-8±àÂë·½Ê½
'(Ê®Áù½øÖÆ)          | £¨¶ş½øÖÆ£©
'--------------------+---------------------------------------------
'0000 0000-0000 007F | 0xxxxxxx
'0000 0080-0000 07FF | 110xxxxx 10xxxxxx
'0000 0800-0000 FFFF | 1110xxxx 10xxxxxx 10xxxxxx
'0001 0000-0010 FFFF | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
'==========================================================================================ÿÿ


'Çø·ÖUEF_ANSI ºÍ UEF_UTF8NB
'ÿ
'¹²Í¬µã:¶şÕß¾ùÎŞBOM Òò´Ëµ÷ÓÃÇ°ĞèÒª½«ÎÄ¼şÖ¸ÕëÖØÖÃµ½×î¿ªÊ¼µÄÎ»ÖÃ
'ÿ
'bufAll£º[in]ÎÄ¼şËùÓĞ×Ö½Ú
'fmt£º[out]·µ»Ø±àÂëÀàĞÍ
Public Function UEFCheckUTF8NoBom(ByRef bufAll() As Byte, ByRef fmt As UnicodeEncodeFormat)
    
    Dim i As Long               '¿ÉÄÜ»áÒç³ö
    Dim cOctets As Long         '¿ÉÒÔÈİÄÉUTF-8±àÂë×Ö·ûµÄ×Ö½Ú´óĞ¡ 4bytes
    Dim bAllAscii As Boolean    'Èç¹ûÈ«²¿ÎªASCII£¬ËµÃ÷²»ÊÇUTF-8
    
    bAllAscii = True
    cOctets = 0
    
    'Debug.Print Hex(bufAll(0)) & "-" & Hex(bufAll(1)) & "-" & Hex(bufAll(2))
    
    For i = 0 To UBound(bufAll)
        If (bufAll(i) And &H80) <> 0 Then
            'ASCIIÓÃ7Î»´¢´æ£¬×î¸ßÎ»Îª0£¬Èç¹ûÕâÀïÏàÓë·Ç0£¬¾Í²»ÊÇASCII
            '¶ÔÓÚµ¥×Ö½ÚµÄ·ûºÅ£¬×Ö½ÚµÄµÚÒ»Î»ÉèÎª0£¬ºóÃæ7Î»ÎªÕâ¸ö·ûºÅµÄunicodeÂë¡£
            'Òò´Ë¶ÔÓÚÓ¢Óï×ÖÄ¸£¬UTF-8±àÂëºÍASCIIÂëÊÇÏàÍ¬µÄ
            bAllAscii = False
        End If
        
        '¶ÔÓÚn×Ö½ÚµÄ·ûºÅ£¨n>1£©£¬µÚÒ»¸ö×Ö½ÚµÄÇ°nÎ»¶¼ÉèÎª1£¬µÚn+1Î»ÉèÎª0£¬ºóÃæ×Ö½ÚµÄÇ°Á½Î»Ò»ÂÉÉèÎª10
        'cOctets = 0 ±íÊ¾±¾×Ö½ÚÊÇleading byte
        If cOctets = 0 Then
            If bufAll(i) >= &H80 Then
                '¼ÆÊı£ºÊÇcOctets×Ö½ÚµÄ·ûºÅ
                Do While (bufAll(i) And &H80) <> 0
                    'bufAll(i)×óÒÆÒ»Î»
                    bufAll(i) = ShLB_By1Bit(bufAll(i))
                    cOctets = cOctets + 1
                Loop
                
                'leading byteÖÁÉÙÓ¦Îª110x xxxx
                cOctets = cOctets - 1
                If cOctets = 0 Then
                    '·µ»ØÄ¬ÈÏ±àÂë
                    fmt = UEF_ANSI
                    Exit Function
                End If
            End If
        Else
            '·Çleading byteĞÎÊ½±ØĞëÊÇ 10xxxxxx
            If (bufAll(i) And &HC0) <> &H80 Then
                '·µ»ØÄ¬ÈÏ±àÂë
                fmt = UEF_ANSI
                Exit Function
            End If
            '×¼±¸ÏÂÒ»¸öbyte
            cOctets = cOctets - 1
        End If
    
    Next i
    
    'ÎÄ±¾½áÊø.  ²»Ó¦ÓĞÈÎºÎ¶àÓàµÄbyte ÓĞ¼´Îª´íÎó ·µ»ØÄ¬ÈÏ±àÂë
    If cOctets > 0 Then
        fmt = UEF_ANSI
        Exit Function
    End If
    
    'Èç¹ûÈ«ÊÇascii.  ĞèÒª×¢ÒâµÄÊÇÊ¹ÓÃÏàÓ¦µÄcode pages×ö×ª»»
    If bAllAscii = True Then
        fmt = UEF_ANSI
        Exit Function
    End If
    
    'ĞŞ³ÉÕı¹û ÖÕÓÚ¸ñÊ½È«²¿ÕıÈ· ·µ»ØUTF8 No BOM±àÂë¸ñÊ½
    fmt = uef_utf8NB
    
End Function

'Éú³ÉBOM
'·µ»ØÖµ£ºBOMËùÕ¼×Ö½Ú
'fmt£º[in]±àÂëÀàĞÍ
'dwFirst£º[out]ÎÄ¼ş×î¿ªÊ¼µÄ4¸ö×Ö½Ú
Public Function UEFMakeBOM(ByVal fmt As UnicodeEncodeFormat, ByRef dwFirst As Long) As Long
    Select Case fmt
    Case UEF_UTF8
        dwFirst = &HBFBBEF
        UEFMakeBOM = 3
    Case UEF_UTF16LE
        dwFirst = &HFEFF&
        UEFMakeBOM = 2
    Case UEF_UTF16BE
        dwFirst = &HFFFE&
        UEFMakeBOM = 2
    Case UEF_UTF32LE
        dwFirst = &HFEFF&
        UEFMakeBOM = 4
    Case UEF_UTF32BE
        dwFirst = &HFFFE0000
        UEFMakeBOM = 4
    Case Else 'UEF_UTF8NBºÍUEF_ANSI
        dwFirst = 0
        UEFMakeBOM = 0
    End Select
End Function

'ÅĞ¶ÏÎÄ±¾ÎÄ¼şµÄ±àÂëÀàĞÍ
'·µ»ØÖµ£º±àÂëÀàĞÍ¡£ÎÄ¼şÎŞ·¨´ò¿ªÊ±£¬·µ»ØUEF_Auto
'FileName£ºÎÄ¼şÃû
Public Function UEFCheckTextFileFormat(ByVal FileName As String) As UnicodeEncodeFormat
    Dim hFile     As Long
    Dim dwFirst   As Long
    Dim nNumRead  As Long
    
    Dim nFileSize As Long
    Dim bufAll()  As Byte

    '´ò¿ªÎÄ¼ş
    hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If INVALID_HANDLE_VALUE = hFile Then 'ÎÄ¼şÎŞ·¨´ò¿ª
        UEFCheckTextFileFormat = UEF_AUTO
        Exit Function
    End If

    'ÅĞ¶ÏBOM
    dwFirst = 0
    Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
    nNumRead = UEFCheckBOM(dwFirst, UEFCheckTextFileFormat) '·µ»ØBOMËùÕ¼×Ö½Ú
    'Debug.Print nNumRead
    
    'Èç¹ûÊÇÅĞ¶Ï½á¹ûÊÇUEF_ANSI ÔòĞè¼ÌĞøÇø·ÖUEF_ANSI ºÍ UEF_UTF8NB
    If UEFCheckTextFileFormat = UEF_ANSI Then
        nFileSize = GetFileSize(hFile, nNumRead)
        ReDim bufAll(0 To nFileSize - 1)
        
        nNumRead = 0
        'UEF_ANSI UEF_UTF8NB µÄcbBOM¾ùÎª0
        Call SetFilePointer(hFile, 0, ByVal 0&, FILE_BEGIN) '»Ö¸´ÎÄ¼şÖ¸Õë
        Call ReadFile(hFile, bufAll(0), nFileSize, nNumRead, ByVal 0&)
        UEFCheckUTF8NoBom bufAll, UEFCheckTextFileFormat
        
    End If

    'Debug.Print UEFCheckTextFileFormat
    
    '¹Ø±ÕÎÄ¼ş
    Call CloseHandle(hFile)

End Function

'¶ÁÈ¡ÎÄ±¾ÎÄ¼ş
'·µ»ØÖµ£º¶ÁÈ¡µÄÎÄ±¾¡£·µ»ØvbNullString±íÊ¾ÎÄ¼şÎŞ·¨´ò¿ª
'FileName£º[in]ÎÄ¼şÃû
'fmt£º[in,out]Ê¹ÓÃºÎÖÖÎÄ±¾±àÂë¸ñÊ½À´¶ÁÈ¡ÎÄ±¾¡£ÎªUEF_AutoÊ±±íÊ¾×Ô¶¯ÅĞ¶Ï£¬ÇÒÔÚfmt²ÎÊı·µ»ØÎÄ±¾ËùÓÃ±àÂë¸ñÊ½
Public Function UEFLoadTextFile(ByVal FileName As String, Optional ByRef fmt As UnicodeEncodeFormat = UEF_AUTO) As String
    Dim hFile As Long
    Dim nFileSize As Long
    Dim nNumRead As Long
    Dim dwFirst As Long
    Dim CurFmt As UnicodeEncodeFormat
    Dim cbBOM As Long
    Dim cbTextData As Long
    Dim CurCP As Long
    Dim byBuf() As Byte
    Dim byBufDiff() As Byte
    Dim cchStr As Long
    Dim i As Long
    Dim byTemp As Byte
    
    'ÅĞ¶Ïfmt·¶Î§
    If fmt <> UEF_AUTO Then
        If fmt < [_UEF_Min] Or fmt > [_UEF_Max] Then
            GoTo FunEnd
        End If
    End If
    
    '´ò¿ªÎÄ¼ş
    hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If INVALID_HANDLE_VALUE = hFile Then 'ÎÄ¼şÎŞ·¨´ò¿ª
        GoTo FunEnd
    End If
    
    'ÅĞ¶ÏÎÄ¼ş´óĞ¡
    nFileSize = GetFileSize(hFile, nNumRead)
    If nNumRead <> 0 Then '³¬¹ı4GB
        GoTo FreeHandle
    End If
    If nFileSize < 0 Then '³¬¹ı2GB
        GoTo FreeHandle
    End If
    
    'ÅĞ¶ÏBOM
    dwFirst = 0
    Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
    cbBOM = UEFCheckBOM(dwFirst, CurFmt)
    '¼ÌĞøÇø·ÖUEF_ANSI ºÍ UEF_UTF8NB cbBOM¶şÕßÏàÍ¬ by shenhao
    If CurFmt = UEF_ANSI Then
        ReDim byBufDiff(0 To nFileSize - 1)
        'UEF_ANSI UEF_UTF8NB µÄcbBOM¾ùÎª0
        Call SetFilePointer(hFile, 0, ByVal 0&, FILE_BEGIN) '»Ö¸´ÎÄ¼şÖ¸Õë
        Call ReadFile(hFile, byBufDiff(0), nFileSize, nNumRead, ByVal 0&)
        UEFCheckUTF8NoBom byBufDiff, CurFmt
    End If
    
    
    '»Ö¸´ÎÄ¼şÖ¸Õë
    If fmt = UEF_AUTO Then '×Ô¶¯ÅĞ¶Ï
        fmt = CurFmt
        'cbBOM = cbBOM
    Else 'ÊÖ¶¯ÉèÖÃ±àÂë
        If fmt = CurFmt Then 'Èô±àÂëÏàÍ¬£¬ÔòºöÂÔBOM±ê¼Ç
            'cbBOM = cbBOM
        Else '±àÂë²»Í¬£¬ÄÇÃ´¶¼ÊÇÊı¾İ
            cbBOM = 0
        End If
    End If
    Call SetFilePointer(hFile, cbBOM, ByVal 0&, FILE_BEGIN)
    cbTextData = nFileSize - cbBOM
    
    '¶ÁÈ¡Êı¾İ
    UEFLoadTextFile = ""
    Select Case fmt
        Case UEF_ANSI, UEF_UTF8, uef_utf8NB
            'ÅĞ¶ÏÓ¦Ê¹ÓÃµÄCodePage
            CurCP = IIf((fmt = UEF_UTF8) Or (fmt = uef_utf8NB), CP_UTF8, UEFCodePage)
            
            '·ÖÅä»º³åÇø
            On Error GoTo FreeHandle
            ReDim byBuf(0 To cbTextData - 1)
            On Error GoTo 0
            
            '¶ÁÈ¡Êı¾İ
            nNumRead = 0
            Call ReadFile(hFile, byBuf(0), cbTextData, nNumRead, ByVal 0&)
            
            'È¡µÃUnicodeÎÄ±¾³¤¶È
            cchStr = MultiByteToWideChar(CurCP, 0, byBuf(0), nNumRead, ByVal 0&, ByVal 0&)
            If cchStr > 0 Then
                '·ÖÅä×Ö·û´®¿Õ¼ä
                On Error GoTo FreeHandle
                UEFLoadTextFile = String$(cchStr, 0)
                On Error GoTo 0
                
                'È¡µÃÎÄ±¾
                cchStr = MultiByteToWideChar(CurCP, 0, byBuf(0), nNumRead, ByVal StrPtr(UEFLoadTextFile), cchStr + 1)
            End If
            
        Case UEF_UTF16LE
            cchStr = (cbTextData + 1) / 2
            
            '·ÖÅä×Ö·û´®¿Õ¼ä
            On Error GoTo FreeHandle
            UEFLoadTextFile = String$(cchStr, 0)
            On Error GoTo 0
            
            'È¡µÃÎÄ±¾
            nNumRead = 0
            Call ReadFile(hFile, ByVal StrPtr(UEFLoadTextFile), cbTextData, nNumRead, ByVal 0&)
            
            'ĞŞÕıÎÄ±¾³¤¶È
            cchStr = (nNumRead + 1) / 2
            If cchStr > 0 Then
                If Len(UEFLoadTextFile) > cchStr Then
                    UEFLoadTextFile = Left$(UEFLoadTextFile, cchStr)
                End If
            Else
                UEFLoadTextFile = ""
            End If
            
        Case UEF_UTF16BE
            '·ÖÅä»º³åÇø
            On Error GoTo FreeHandle
            ReDim byBuf(0 To cbTextData - 1)
            On Error GoTo 0
            
            '¶ÁÈ¡Êı¾İ
            nNumRead = 0
            Call ReadFile(hFile, byBuf(0), cbTextData, nNumRead, ByVal 0&)
            
            If nNumRead > 0 Then
                '¸ôÁ½×Ö½Ú·­×ªÏàÁÚ×Ö½Ú
                For i = 0 To nNumRead - 1 - 1 Step 2 'ÔÙ-1ÊÇÎªÁË±ÜÃâ×îºó¶à³öµÄÄÇ¸ö×Ö½Ú
                    byTemp = byBuf(i)
                    byBuf(i) = byBuf(i + 1)
                    byBuf(i + 1) = byTemp
                Next i
                
                'È¡µÃÎÄ±¾
                UEFLoadTextFile = byBuf 'VBÔÊĞíStringÖĞµÄ×Ö·û´®Êı¾İÓëByteÊı×éÖ±½Ó×ª»»
            End If
            
        Case UEF_UTF32LE
            UEFLoadTextFile = vbNullString 'ÔİÊ±²»Ö§³Ö
        Case UEF_UTF32BE
            UEFLoadTextFile = vbNullString 'ÔİÊ±²»Ö§³Ö
        Case Else
            Debug.Assert False
    End Select
    
FreeHandle:
    '¹Ø±ÕÎÄ¼ş
    Call CloseHandle(hFile)
    
FunEnd:

End Function

'±£´æÎÄ±¾ÎÄ¼ş
'·µ»ØÖµ£ºÊÇ·ñ³É¹¦
'FileName£º[in]ÎÄ¼şÃû
'sText£º[in]ÓûÊä³öµÄÎÄ±¾
'IsAppend£º[in]ÊÇ·ñÊÇÌí¼Ó·½Ê½
'fmt£º[in,out]Ê¹ÓÃºÎÖÖÎÄ±¾±àÂë¸ñÊ½À´´æ´¢ÎÄ±¾¡£µ±IsAppend=TrueÊ±ÔÊĞíUEF_Auto×Ô¶¯ÅĞ¶Ï£¬ÇÒÔÚfmt²ÎÊı·µ»ØÎÄ±¾ËùÓÃ±àÂë¸ñÊ½
'DefFmt£º[in]µ±Ê¹ÓÃÌí¼ÓÄ£Ê½Ê±£¬ÈôÎÄ¼ş²»´æÔÚÇÒfmt = UEF_AutoÊ±Ó¦Ê¹ÓÃµÄ±àÂë¸ñÊ½
Public Function UEFSaveTextFile(ByVal FileName As String, _
                                ByRef sText As String, _
                                Optional ByVal IsAppend As Boolean = False, _
                                Optional ByRef fmt As UnicodeEncodeFormat = UEF_AUTO, _
                                Optional ByVal DefFmt As UnicodeEncodeFormat = UEF_ANSI) As Boolean
                                
    Dim hFile As Long
    Dim nFileSize As Long
    Dim nNumRead As Long
    Dim dwFirst As Long
    Dim cbBOM As Long
    Dim CurCP As Long
    Dim byBuf() As Byte
    Dim byBufDiff() As Byte
    Dim cbBuf As Long
    Dim i As Long
    Dim byTemp As Byte
    
    'ÅĞ¶Ïfmt·¶Î§
    If IsAppend And (fmt = UEF_AUTO) Then
    Else
        If fmt < [_UEF_Min] Or fmt > [_UEF_Max] Then
            GoTo FunEnd
        End If
    End If
    
    '´ò¿ªÎÄ¼ş
    hFile = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, IIf(IsAppend, OPEN_ALWAYS, Create_ALWAYS), FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If INVALID_HANDLE_VALUE = hFile Then 'ÎÄ¼şÎŞ·¨´ò¿ª
            GoTo FunEnd
    End If
    
    'ÅĞ¶ÏÎÄ¼ş´óĞ¡
    nFileSize = GetFileSize(hFile, nNumRead)
    If nFileSize = 0 And nNumRead = 0 Then 'ÎÄ¼ş´óĞ¡Îª0×Ö½Ú
         IsAppend = False '´ËÊ±ĞèÒªĞ´BOM±êÖ¾
    End If
    If fmt = UEF_AUTO Then
        fmt = DefFmt
    End If
    
    'ÅĞ¶ÏBOM
    If IsAppend And (fmt = UEF_AUTO) Then
        dwFirst = 0
        Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
        cbBOM = UEFCheckBOM(dwFirst, fmt)
        '¼ÌĞøÇø·ÖUEF_ANSI ºÍ UEF_UTF8NB cbBOM¶şÕßÏàÍ¬ by shenhao
        If fmt = UEF_ANSI Then
            ReDim byBufDiff(0 To nFileSize - 1)
            'UEF_ANSI UEF_UTF8NB µÄcbBOM¾ùÎª0
            Call SetFilePointer(hFile, 0, ByVal 0&, FILE_BEGIN) '»Ö¸´ÎÄ¼şÖ¸Õë
            Call ReadFile(hFile, byBufDiff(0), nFileSize, nNumRead, ByVal 0&)
            UEFCheckUTF8NoBom byBufDiff, fmt
        End If
        
    ElseIf IsAppend = False Then
        cbBOM = UEFMakeBOM(fmt, dwFirst)
    End If
    
    'ÎÄ¼şÖ¸Õë¶¨Î»
    Call SetFilePointer(hFile, 0, ByVal 0&, IIf(IsAppend, FILE_END, FILE_BEGIN))
    
    'Ğ´BOM
    If IsAppend = False Then
        If cbBOM > 0 Then
            Call WriteFile(hFile, dwFirst, cbBOM, nNumRead, ByVal 0&)
        End If
    End If
    
    'Ğ´ÎÄ±¾Êı¾İ
    If Len(sText) > 0 Then
        Select Case fmt
            Case UEF_ANSI, UEF_UTF8, uef_utf8NB
                'ÅĞ¶ÏÓ¦Ê¹ÓÃµÄCodePage
                CurCP = IIf((fmt = UEF_UTF8) Or (fmt = uef_utf8NB), CP_UTF8, UEFCodePage)
                
                'È¡µÃ»º³åÇø´óĞ¡
                cbBuf = WideCharToMultiByte(CurCP, 0, ByVal StrPtr(sText), Len(sText), ByVal 0&, 0, ByVal 0&, ByVal 0&)
                If cbBuf > 0 Then
                    '·ÖÅä»º³åÇø
                    On Error GoTo FreeHandle
                    ReDim byBuf(0 To cbBuf)
                    On Error GoTo 0
                
                    '×ª»»ÎÄ±¾
                    cbBuf = WideCharToMultiByte(CurCP, 0, ByVal StrPtr(sText), Len(sText), byBuf(0), cbBuf + 1, ByVal 0&, ByVal 0&)
                
                    'Ğ´ÎÄ¼ş
                    Call WriteFile(hFile, byBuf(0), cbBuf, nNumRead, ByVal 0&)
                
                    UEFSaveTextFile = True
                End If
                
            Case UEF_UTF16LE
                'Ğ´ÎÄ¼ş
                Call WriteFile(hFile, ByVal StrPtr(sText), LenB(sText), nNumRead, ByVal 0&)
            
                UEFSaveTextFile = True
            
            Case UEF_UTF16BE
                '½«×Ö·û´®ÖĞµÄÊı¾İ¸´ÖÆµ½byBuf
                On Error GoTo FreeHandle
                byBuf = sText
                On Error GoTo 0
                cbBuf = UBound(byBuf) - LBound(byBuf) + 1
            
                '¸ôÁ½×Ö½Ú·­×ªÏàÁÚ×Ö½Ú
                For i = 0 To cbBuf - 1 - 1 Step 2 'ÔÙ-1ÊÇÎªÁË±ÜÃâ×îºó¶à³öµÄÄÇ¸ö×Ö½Ú
                    byTemp = byBuf(i)
                    byBuf(i) = byBuf(i + 1)
                    byBuf(i + 1) = byTemp
                Next i
            
                'Ğ´ÎÄ¼ş
                Call WriteFile(hFile, byBuf(0), cbBuf, nNumRead, ByVal 0&)
            
                UEFSaveTextFile = True
            
            Case UEF_UTF32LE
                UEFSaveTextFile = False 'ÔİÊ±²»Ö§³Ö
            Case UEF_UTF32BE
                UEFSaveTextFile = False 'ÔİÊ±²»Ö§³Ö
            Case Else
                Debug.Assert False
        End Select
    Else
        UEFSaveTextFile = True
    End If
    
FreeHandle:
    '¹Ø±ÕÎÄ¼ş
    Call CloseHandle(hFile)
    
FunEnd:
End Function


