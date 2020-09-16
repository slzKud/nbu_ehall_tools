Attribute VB_Name = "mTextUTF"
Option Explicit
'https://www.cnblogs.com/shenhaocn/archive/2011/10/23/2221572.html
'mTextUTF.bas
'ģ�飺UTF�ı��ļ�����
'���ߣ�zyl910
'�汾��1.0
'���ڣ�2006-1-23

'�޸�: shenhao(shenhaocn@qq.com)
'�汾: 1.1
'����: 2011-10-23

'== ˵�� ===================================================
'����֧��UTF-8����BOM��ʽ���� shenhao
'֧��Unicode������ı��ļ���д����ʱ֧��ANSI��UTF-8��UTF-16LE��UTF-16BE�⼸�ֱ����ı�

'== ���¼�¼ ===============================================
'[V1.0] 2006-1-23
'1.֧�������ANSI��UTF-8��UTF-16LE��UTF-16BE�⼸�ֱ����ı�
'
'[V1.1] 2011-10-23 by shenhao (shenhaocn@qq.com)
'1.֧��UTF-8����BOM��ʽ����

'## ����Ԥ������ #########################################
'== ȫ�ֳ��� ===============================================
'IncludeAPILib��������API�⣬��ʱ����Ҫ�ֶ�дAPI����

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

'Unicode�����ʽ
Public Enum UnicodeEncodeFormat
UEF_ANSI = 0 'ANSI+DBCS
UEF_UTF8     'UTF-8
uef_utf8NB   'UTF-8 No BOM
UEF_UTF16LE  'UTF-16LE
UEF_UTF16BE  'UTF-16BE
UEF_UTF32LE  'UTF-32LE
UEF_UTF32BE  'UTF-32BE

UEF_AUTO = -1 '�Զ�ʶ�����

'������Ŀ
[_UEF_Min] = UEF_ANSI
[_UEF_Max] = UEF_UTF32BE

End Enum

'ANSI+DBCS��ʽ���ı���ʹ�õĴ���ҳ��Ĭ��Ϊ0����ʾʹ��ϵͳ��ǰ����ҳ��
'�������øò���ʵ�ֶ�ȡ�������������ı����������� ��������ƽ̨�� ��ȡ ��������ƽ̨���ɵ�txt���ͽ�����Ϊ950
Public UEFCodePage As Long


'��BYTE���ͱ�������1λ�ĺ���
'����ֵ����λ���
'Byt������λ���ֽ�
Private Function ShLB_By1Bit(ByVal Byt As Byte) As Byte
    
    '(Byt And &H7F)���������������λ�� *2������һλ
    ShLB_By1Bit = (Byt And &H7F) * 2

End Function

'�ж�BOM
'����ֵ��BOM��ռ�ֽ�
'dwFirst��[in]�ļ��ʼ��4���ֽ�
'fmt��[out]���ر�������
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
    Else '���ݶ�ΪUEF_ANSI ������������UEF_ANSI �� UEF_UTF8NB
        fmt = UEF_ANSI
        UEFCheckBOM = 0
    End If
End Function

'==========================================================================================
'UTF-8����һ���ص㣬��������һ�ֱ䳤�ı��뷽ʽ��
'������ʹ��1~4���ֽڱ�ʾһ�����ţ����ݲ�ͬ�ķ��Ŷ��仯�ֽڳ��ȡ�
'�
'UTF-8�ı������ܼ򵥣�ֻ�ж�����
'�
'1�����ڵ��ֽڵķ��ţ��ֽڵĵ�һλ��Ϊ0������7λΪ������ŵ�unicode�롣
'   ��˶���Ӣ����ĸ��UTF-8�����ASCII������ͬ�ġ�
'�
'2������n�ֽڵķ��ţ�n>1������һ���ֽڵ�ǰnλ����Ϊ1����n+1λ��Ϊ0�������ֽڵ�ǰ��λһ����Ϊ10��
'   ʣ�µ�û���ἰ�Ķ�����λ��ȫ��Ϊ������ŵ�unicode�롣
'�
'�±��ܽ��˱��������ĸx��ʾ���ñ����λ��
'Unicode���ŷ�Χ     | UTF-8���뷽ʽ
'(ʮ������)          | �������ƣ�
'--------------------+---------------------------------------------
'0000 0000-0000 007F | 0xxxxxxx
'0000 0080-0000 07FF | 110xxxxx 10xxxxxx
'0000 0800-0000 FFFF | 1110xxxx 10xxxxxx 10xxxxxx
'0001 0000-0010 FFFF | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
'==========================================================================================��


'����UEF_ANSI �� UEF_UTF8NB
'�
'��ͬ��:���߾���BOM ��˵���ǰ��Ҫ���ļ�ָ�����õ��ʼ��λ��
'�
'bufAll��[in]�ļ������ֽ�
'fmt��[out]���ر�������
Public Function UEFCheckUTF8NoBom(ByRef bufAll() As Byte, ByRef fmt As UnicodeEncodeFormat)
    
    Dim i As Long               '���ܻ����
    Dim cOctets As Long         '��������UTF-8�����ַ����ֽڴ�С 4bytes
    Dim bAllAscii As Boolean    '���ȫ��ΪASCII��˵������UTF-8
    
    bAllAscii = True
    cOctets = 0
    
    'Debug.Print Hex(bufAll(0)) & "-" & Hex(bufAll(1)) & "-" & Hex(bufAll(2))
    
    For i = 0 To UBound(bufAll)
        If (bufAll(i) And &H80) <> 0 Then
            'ASCII��7λ���棬���λΪ0��������������0���Ͳ���ASCII
            '���ڵ��ֽڵķ��ţ��ֽڵĵ�һλ��Ϊ0������7λΪ������ŵ�unicode�롣
            '��˶���Ӣ����ĸ��UTF-8�����ASCII������ͬ��
            bAllAscii = False
        End If
        
        '����n�ֽڵķ��ţ�n>1������һ���ֽڵ�ǰnλ����Ϊ1����n+1λ��Ϊ0�������ֽڵ�ǰ��λһ����Ϊ10
        'cOctets = 0 ��ʾ���ֽ���leading byte
        If cOctets = 0 Then
            If bufAll(i) >= &H80 Then
                '��������cOctets�ֽڵķ���
                Do While (bufAll(i) And &H80) <> 0
                    'bufAll(i)����һλ
                    bufAll(i) = ShLB_By1Bit(bufAll(i))
                    cOctets = cOctets + 1
                Loop
                
                'leading byte����ӦΪ110x xxxx
                cOctets = cOctets - 1
                If cOctets = 0 Then
                    '����Ĭ�ϱ���
                    fmt = UEF_ANSI
                    Exit Function
                End If
            End If
        Else
            '��leading byte��ʽ������ 10xxxxxx
            If (bufAll(i) And &HC0) <> &H80 Then
                '����Ĭ�ϱ���
                fmt = UEF_ANSI
                Exit Function
            End If
            '׼����һ��byte
            cOctets = cOctets - 1
        End If
    
    Next i
    
    '�ı�����.  ��Ӧ���κζ����byte �м�Ϊ���� ����Ĭ�ϱ���
    If cOctets > 0 Then
        fmt = UEF_ANSI
        Exit Function
    End If
    
    '���ȫ��ascii.  ��Ҫע�����ʹ����Ӧ��code pages��ת��
    If bAllAscii = True Then
        fmt = UEF_ANSI
        Exit Function
    End If
    
    '�޳����� ���ڸ�ʽȫ����ȷ ����UTF8 No BOM�����ʽ
    fmt = uef_utf8NB
    
End Function

'����BOM
'����ֵ��BOM��ռ�ֽ�
'fmt��[in]��������
'dwFirst��[out]�ļ��ʼ��4���ֽ�
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
    Case Else 'UEF_UTF8NB��UEF_ANSI
        dwFirst = 0
        UEFMakeBOM = 0
    End Select
End Function

'�ж��ı��ļ��ı�������
'����ֵ���������͡��ļ��޷���ʱ������UEF_Auto
'FileName���ļ���
Public Function UEFCheckTextFileFormat(ByVal FileName As String) As UnicodeEncodeFormat
    Dim hFile     As Long
    Dim dwFirst   As Long
    Dim nNumRead  As Long
    
    Dim nFileSize As Long
    Dim bufAll()  As Byte

    '���ļ�
    hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If INVALID_HANDLE_VALUE = hFile Then '�ļ��޷���
        UEFCheckTextFileFormat = UEF_AUTO
        Exit Function
    End If

    '�ж�BOM
    dwFirst = 0
    Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
    nNumRead = UEFCheckBOM(dwFirst, UEFCheckTextFileFormat) '����BOM��ռ�ֽ�
    'Debug.Print nNumRead
    
    '������жϽ����UEF_ANSI �����������UEF_ANSI �� UEF_UTF8NB
    If UEFCheckTextFileFormat = UEF_ANSI Then
        nFileSize = GetFileSize(hFile, nNumRead)
        ReDim bufAll(0 To nFileSize - 1)
        
        nNumRead = 0
        'UEF_ANSI UEF_UTF8NB ��cbBOM��Ϊ0
        Call SetFilePointer(hFile, 0, ByVal 0&, FILE_BEGIN) '�ָ��ļ�ָ��
        Call ReadFile(hFile, bufAll(0), nFileSize, nNumRead, ByVal 0&)
        UEFCheckUTF8NoBom bufAll, UEFCheckTextFileFormat
        
    End If

    'Debug.Print UEFCheckTextFileFormat
    
    '�ر��ļ�
    Call CloseHandle(hFile)

End Function

'��ȡ�ı��ļ�
'����ֵ����ȡ���ı�������vbNullString��ʾ�ļ��޷���
'FileName��[in]�ļ���
'fmt��[in,out]ʹ�ú����ı������ʽ����ȡ�ı���ΪUEF_Autoʱ��ʾ�Զ��жϣ�����fmt���������ı����ñ����ʽ
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
    
    '�ж�fmt��Χ
    If fmt <> UEF_AUTO Then
        If fmt < [_UEF_Min] Or fmt > [_UEF_Max] Then
            GoTo FunEnd
        End If
    End If
    
    '���ļ�
    hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If INVALID_HANDLE_VALUE = hFile Then '�ļ��޷���
        GoTo FunEnd
    End If
    
    '�ж��ļ���С
    nFileSize = GetFileSize(hFile, nNumRead)
    If nNumRead <> 0 Then '����4GB
        GoTo FreeHandle
    End If
    If nFileSize < 0 Then '����2GB
        GoTo FreeHandle
    End If
    
    '�ж�BOM
    dwFirst = 0
    Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
    cbBOM = UEFCheckBOM(dwFirst, CurFmt)
    '��������UEF_ANSI �� UEF_UTF8NB cbBOM������ͬ by shenhao
    If CurFmt = UEF_ANSI Then
        ReDim byBufDiff(0 To nFileSize - 1)
        'UEF_ANSI UEF_UTF8NB ��cbBOM��Ϊ0
        Call SetFilePointer(hFile, 0, ByVal 0&, FILE_BEGIN) '�ָ��ļ�ָ��
        Call ReadFile(hFile, byBufDiff(0), nFileSize, nNumRead, ByVal 0&)
        UEFCheckUTF8NoBom byBufDiff, CurFmt
    End If
    
    
    '�ָ��ļ�ָ��
    If fmt = UEF_AUTO Then '�Զ��ж�
        fmt = CurFmt
        'cbBOM = cbBOM
    Else '�ֶ����ñ���
        If fmt = CurFmt Then '��������ͬ�������BOM���
            'cbBOM = cbBOM
        Else '���벻ͬ����ô��������
            cbBOM = 0
        End If
    End If
    Call SetFilePointer(hFile, cbBOM, ByVal 0&, FILE_BEGIN)
    cbTextData = nFileSize - cbBOM
    
    '��ȡ����
    UEFLoadTextFile = ""
    Select Case fmt
        Case UEF_ANSI, UEF_UTF8, uef_utf8NB
            '�ж�Ӧʹ�õ�CodePage
            CurCP = IIf((fmt = UEF_UTF8) Or (fmt = uef_utf8NB), CP_UTF8, UEFCodePage)
            
            '���仺����
            On Error GoTo FreeHandle
            ReDim byBuf(0 To cbTextData - 1)
            On Error GoTo 0
            
            '��ȡ����
            nNumRead = 0
            Call ReadFile(hFile, byBuf(0), cbTextData, nNumRead, ByVal 0&)
            
            'ȡ��Unicode�ı�����
            cchStr = MultiByteToWideChar(CurCP, 0, byBuf(0), nNumRead, ByVal 0&, ByVal 0&)
            If cchStr > 0 Then
                '�����ַ����ռ�
                On Error GoTo FreeHandle
                UEFLoadTextFile = String$(cchStr, 0)
                On Error GoTo 0
                
                'ȡ���ı�
                cchStr = MultiByteToWideChar(CurCP, 0, byBuf(0), nNumRead, ByVal StrPtr(UEFLoadTextFile), cchStr + 1)
            End If
            
        Case UEF_UTF16LE
            cchStr = (cbTextData + 1) / 2
            
            '�����ַ����ռ�
            On Error GoTo FreeHandle
            UEFLoadTextFile = String$(cchStr, 0)
            On Error GoTo 0
            
            'ȡ���ı�
            nNumRead = 0
            Call ReadFile(hFile, ByVal StrPtr(UEFLoadTextFile), cbTextData, nNumRead, ByVal 0&)
            
            '�����ı�����
            cchStr = (nNumRead + 1) / 2
            If cchStr > 0 Then
                If Len(UEFLoadTextFile) > cchStr Then
                    UEFLoadTextFile = Left$(UEFLoadTextFile, cchStr)
                End If
            Else
                UEFLoadTextFile = ""
            End If
            
        Case UEF_UTF16BE
            '���仺����
            On Error GoTo FreeHandle
            ReDim byBuf(0 To cbTextData - 1)
            On Error GoTo 0
            
            '��ȡ����
            nNumRead = 0
            Call ReadFile(hFile, byBuf(0), cbTextData, nNumRead, ByVal 0&)
            
            If nNumRead > 0 Then
                '�����ֽڷ�ת�����ֽ�
                For i = 0 To nNumRead - 1 - 1 Step 2 '��-1��Ϊ�˱�����������Ǹ��ֽ�
                    byTemp = byBuf(i)
                    byBuf(i) = byBuf(i + 1)
                    byBuf(i + 1) = byTemp
                Next i
                
                'ȡ���ı�
                UEFLoadTextFile = byBuf 'VB����String�е��ַ���������Byte����ֱ��ת��
            End If
            
        Case UEF_UTF32LE
            UEFLoadTextFile = vbNullString '��ʱ��֧��
        Case UEF_UTF32BE
            UEFLoadTextFile = vbNullString '��ʱ��֧��
        Case Else
            Debug.Assert False
    End Select
    
FreeHandle:
    '�ر��ļ�
    Call CloseHandle(hFile)
    
FunEnd:

End Function

'�����ı��ļ�
'����ֵ���Ƿ�ɹ�
'FileName��[in]�ļ���
'sText��[in]��������ı�
'IsAppend��[in]�Ƿ�����ӷ�ʽ
'fmt��[in,out]ʹ�ú����ı������ʽ���洢�ı�����IsAppend=Trueʱ����UEF_Auto�Զ��жϣ�����fmt���������ı����ñ����ʽ
'DefFmt��[in]��ʹ�����ģʽʱ�����ļ���������fmt = UEF_AutoʱӦʹ�õı����ʽ
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
    
    '�ж�fmt��Χ
    If IsAppend And (fmt = UEF_AUTO) Then
    Else
        If fmt < [_UEF_Min] Or fmt > [_UEF_Max] Then
            GoTo FunEnd
        End If
    End If
    
    '���ļ�
    hFile = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, IIf(IsAppend, OPEN_ALWAYS, Create_ALWAYS), FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If INVALID_HANDLE_VALUE = hFile Then '�ļ��޷���
            GoTo FunEnd
    End If
    
    '�ж��ļ���С
    nFileSize = GetFileSize(hFile, nNumRead)
    If nFileSize = 0 And nNumRead = 0 Then '�ļ���СΪ0�ֽ�
         IsAppend = False '��ʱ��ҪдBOM��־
    End If
    If fmt = UEF_AUTO Then
        fmt = DefFmt
    End If
    
    '�ж�BOM
    If IsAppend And (fmt = UEF_AUTO) Then
        dwFirst = 0
        Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
        cbBOM = UEFCheckBOM(dwFirst, fmt)
        '��������UEF_ANSI �� UEF_UTF8NB cbBOM������ͬ by shenhao
        If fmt = UEF_ANSI Then
            ReDim byBufDiff(0 To nFileSize - 1)
            'UEF_ANSI UEF_UTF8NB ��cbBOM��Ϊ0
            Call SetFilePointer(hFile, 0, ByVal 0&, FILE_BEGIN) '�ָ��ļ�ָ��
            Call ReadFile(hFile, byBufDiff(0), nFileSize, nNumRead, ByVal 0&)
            UEFCheckUTF8NoBom byBufDiff, fmt
        End If
        
    ElseIf IsAppend = False Then
        cbBOM = UEFMakeBOM(fmt, dwFirst)
    End If
    
    '�ļ�ָ�붨λ
    Call SetFilePointer(hFile, 0, ByVal 0&, IIf(IsAppend, FILE_END, FILE_BEGIN))
    
    'дBOM
    If IsAppend = False Then
        If cbBOM > 0 Then
            Call WriteFile(hFile, dwFirst, cbBOM, nNumRead, ByVal 0&)
        End If
    End If
    
    'д�ı�����
    If Len(sText) > 0 Then
        Select Case fmt
            Case UEF_ANSI, UEF_UTF8, uef_utf8NB
                '�ж�Ӧʹ�õ�CodePage
                CurCP = IIf((fmt = UEF_UTF8) Or (fmt = uef_utf8NB), CP_UTF8, UEFCodePage)
                
                'ȡ�û�������С
                cbBuf = WideCharToMultiByte(CurCP, 0, ByVal StrPtr(sText), Len(sText), ByVal 0&, 0, ByVal 0&, ByVal 0&)
                If cbBuf > 0 Then
                    '���仺����
                    On Error GoTo FreeHandle
                    ReDim byBuf(0 To cbBuf)
                    On Error GoTo 0
                
                    'ת���ı�
                    cbBuf = WideCharToMultiByte(CurCP, 0, ByVal StrPtr(sText), Len(sText), byBuf(0), cbBuf + 1, ByVal 0&, ByVal 0&)
                
                    'д�ļ�
                    Call WriteFile(hFile, byBuf(0), cbBuf, nNumRead, ByVal 0&)
                
                    UEFSaveTextFile = True
                End If
                
            Case UEF_UTF16LE
                'д�ļ�
                Call WriteFile(hFile, ByVal StrPtr(sText), LenB(sText), nNumRead, ByVal 0&)
            
                UEFSaveTextFile = True
            
            Case UEF_UTF16BE
                '���ַ����е����ݸ��Ƶ�byBuf
                On Error GoTo FreeHandle
                byBuf = sText
                On Error GoTo 0
                cbBuf = UBound(byBuf) - LBound(byBuf) + 1
            
                '�����ֽڷ�ת�����ֽ�
                For i = 0 To cbBuf - 1 - 1 Step 2 '��-1��Ϊ�˱�����������Ǹ��ֽ�
                    byTemp = byBuf(i)
                    byBuf(i) = byBuf(i + 1)
                    byBuf(i + 1) = byTemp
                Next i
            
                'д�ļ�
                Call WriteFile(hFile, byBuf(0), cbBuf, nNumRead, ByVal 0&)
            
                UEFSaveTextFile = True
            
            Case UEF_UTF32LE
                UEFSaveTextFile = False '��ʱ��֧��
            Case UEF_UTF32BE
                UEFSaveTextFile = False '��ʱ��֧��
            Case Else
                Debug.Assert False
        End Select
    Else
        UEFSaveTextFile = True
    End If
    
FreeHandle:
    '�ر��ļ�
    Call CloseHandle(hFile)
    
FunEnd:
End Function


