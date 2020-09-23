Attribute VB_Name = "modAlbumArt"

' ********************************************************************
'  This is MP3 Album Art tagging module by OTIV.
'  It reads, writes and deletes Album Art in ID3 v2.2, v2.3 and v2.4.
' ********************************************************************

Option Explicit
Private Type v2TagHeader
    Identifier(2)                             As Byte
    Version(1)                                As Byte
    flags                                     As Byte
    Size(3)                                   As Byte
End Type
Private Type v2_2FrameHeader
    Ident(2)                                  As Byte
    Size(2)                                   As Byte
End Type
Private Type v2_34FrameHeader
    Ident(3)                                  As Byte
    Size(3)                                   As Byte
    flags(1)                                  As Byte
End Type
Private Enum v2_StrEncoding
    ENC_ISO = 0
    ENC_UNICODE_UTF16_BOM = 1
    ENC_UNICODE_UTF16 = 2
    ENC_UNICODE_UTF8 = 3
End Enum
#If False Then
Private ENC_ISO, ENC_UNICODE_UTF16_BOM, ENC_UNICODE_UTF16, ENC_UNICODE_UTF8
#End If
Private Const GENERIC_READ                As Long = &H80000000
Private Const FILE_SHARE_READ             As Long = &H1
Private Const GENERIC_WRITE               As Long = &H40000000
Private Const OPEN_EXISTING               As Long = &H3
Private Const FILE_BEGIN                  As Long = 0
Private Const FILE_ATTRIBUTE_READONLY     As Long = &H1
Private Const INVALID_HANDLE_VALUE        As Long = -1
Private TFrameSize                        As Long
Private TData                             As Long
Public TVersion                           As Long
Private rsize                             As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, _
                                                            ByVal fDeleteOnRelease As Long, _
                                                            ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, _
                                                     ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, _
                                                        ByVal lSize As Long, _
                                                        ByVal fRunmode As Long, _
                                                        riid As Any, _
                                                        ppvObj As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
                                                                        ByVal dwDesiredAccess As Long, _
                                                                        ByVal dwShareMode As Long, _
                                                                        lpSecurityAttributes As Any, _
                                                                        ByVal dwCreationDisposition As Long, _
                                                                        ByVal dwFlagsAndAttributes As Long, _
                                                                        ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, _
                                                     lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, _
                                                        ByVal lDistanceToMove As Long, _
                                                        lpDistanceToMoveHigh As Long, _
                                                        ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
                                                  lpBuffer As Any, _
                                                  ByVal nNumberOfBytesToRead As Long, _
                                                  lpNumberOfBytesRead As Long, _
                                                  lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, _
                                                   lpBuffer As Any, _
                                                   ByVal nNumberOfBytesToWrite As Long, _
                                                   lpNumberOfBytesWritten As Long, _
                                                   lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, _
                                                                                      ByVal dwFileAttributes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)

Private Function ArrayToPicture(theArray() As Byte) As IPicture

Dim o_hMem         As Long
Dim o_lpMem        As Long
Dim o_lngByteCount As Long
Dim aGUID(0 To 3)  As Long
Dim IIStream       As IUnknown

    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    o_lngByteCount = UBound(theArray) - LBound(theArray) + 1
    o_hMem = GlobalAlloc(&H2&, o_lngByteCount)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, theArray(LBound(theArray)), o_lngByteCount
            GlobalUnlock o_hMem
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                OleLoadPicture ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture
            End If
        End If
    End If

End Function

Private Function Data2Long(ByVal pData As Long, _
                           ByVal bSynchSafe As Boolean) As Long

Dim i       As Integer
Dim Data(3) As Byte

    CopyMemory Data(0), ByVal pData, 4
    For i = 0 To 3
        If Data(i) And &H80& Then
            bSynchSafe = False
        End If
    Next i
    If bSynchSafe Then
        Data2Long = (Data(0) * &H200000) Or (Data(1) * &H4000&) Or (Data(2) * &H80&) Or Data(3)
    Else
        Data2Long = (Data(0) * &H1000000) Or (Data(1) * &H10000) Or (Data(2) * &H100&) Or Data(3)
    End If

End Function

Private Function Data2String(ByVal pData As Long, _
                             ByVal Length As Long, _
                             ByVal EncFormat As v2_StrEncoding, _
                             Optional ByVal BreakOnNull As Boolean = True) As String

Dim i       As Long
Dim cursize As Byte
Dim curSign As String

    For i = 0 To Length - 1
        CopyMemory cursize, ByVal pData + i, 1
        If cursize = 13 Then
            curSign = vbNullString
        ElseIf cursize = 10 Then
            curSign = vbNewLine
        Else
            curSign = Chr$(cursize)
        End If
        If EncFormat = ENC_ISO Or EncFormat = ENC_UNICODE_UTF8 Then
            If cursize = 0 And BreakOnNull Then
                Exit Function
            Else
                Data2String = Data2String & curSign
            End If
        ElseIf EncFormat = ENC_UNICODE_UTF16_BOM Then
            If i >= 2 Then
                If i Mod 2 = 0 Then
                    If cursize = 0 And BreakOnNull Then
                        Exit Function
                    Else
                        Data2String = Data2String & curSign
                    End If
                End If
            End If
        ElseIf EncFormat = ENC_UNICODE_UTF16 Then
            If i Mod 2 = 0 Then
                If cursize = 0 And BreakOnNull Then
                    Exit Function
                Else
                    Data2String = Data2String & curSign
                End If
            End If
        End If
    Next i

End Function

Public Function DeleteAlbumArt(ByVal Filename As String, _
                               ByVal FrameIndex As Long, _
                               Optional ByVal bIgnoreReadOnly As Boolean = True) As Boolean

Dim fh        As Long
Dim fsize     As Long
Dim tagsize   As Long
Dim TagHeader As v2TagHeader
Dim WrBuf()   As Byte
Dim pbuffer   As Long

    If bIgnoreReadOnly Then
        IgnoreReadOnly Filename
    End If
    ReadAlbumArt Filename, FrameIndex
    fh = CreateFile(Filename, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then
        Exit Function
    End If
    If TFrameSize > 0 Then
        tagsize = GetTagSize(fh, 0) - TFrameSize - Len(TagHeader)
        fsize = GetFileSize(fh, 0)
        ReDim WrBuf(fsize - TFrameSize) As Byte
        pbuffer = VarPtr(WrBuf(0))
        SetFilePointer fh, 0, 0, FILE_BEGIN
        ReadFile fh, ByVal pbuffer, TData, 0, ByVal 0
        pbuffer = VarPtr(WrBuf(TData))
        SetFilePointer fh, TData + TFrameSize, 0, FILE_BEGIN
        ReadFile fh, ByVal pbuffer, fsize - (TData + TFrameSize - 1), 0, ByVal 0
        Long2Data tagsize, IIf(TVersion = 2, True, True), VarPtr(WrBuf(6))
        fsize = fsize - TFrameSize
        SetFilePointer fh, fsize, 0, FILE_BEGIN
        SetEndOfFile fh
        pbuffer = VarPtr(WrBuf(0))
        SetFilePointer fh, 0, 0, FILE_BEGIN
        WriteFile fh, ByVal pbuffer, fsize, 0, ByVal 0
        DeleteAlbumArt = True
    End If
    CloseHandle fh

End Function

Public Function GetAlbumArtCount(ByVal Filename As String) As Long

Dim fh         As Long
Dim ExistSize  As Long
Dim pbuffer    As Long
Dim id3v2tag() As Byte
Dim tp         As Long
Dim SizeBuf(3) As Byte
Dim tp2        As Long

    fh = CreateFile(Filename, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If Not fh = INVALID_HANDLE_VALUE Then
        ExistSize = GetTagSize(fh, 0)
        If ExistSize <> 0 Then
            ReDim id3v2tag(ExistSize - 1) As Byte
            pbuffer = VarPtr(id3v2tag(0))
            SetFilePointer fh, 10, 0, FILE_BEGIN
            ReadFile fh, ByVal pbuffer, ExistSize - 10, 0, ByVal 0
            If TVersion > 2 Then
                Do Until tp >= ExistSize - 11
                    If Data2String(VarPtr(id3v2tag(tp)), 4, ENC_ISO) = "APIC" Then
                        tp2 = tp2 + 1
                    End If
                    tp = tp + Data2Long(VarPtr(id3v2tag(tp + 4)), IIf(TVersion = 3, False, True)) + 10
                Loop
            Else
                Do Until tp >= ExistSize - 7
                    CopyMemory SizeBuf(1), id3v2tag(tp + 3), 3
                    If Data2String(VarPtr(id3v2tag(tp)), 3, ENC_ISO, False) = "PIC" Then
                        tp2 = tp2 + 1
                    End If
                    tp = tp + Data2Long(VarPtr(SizeBuf(0)), False) + 6
                Loop
            End If
        End If
        CloseHandle fh
        GetAlbumArtCount = tp2
    End If

End Function

Private Function GetFrameSize(ByVal AlbumArt As IPictureDisp, _
                              ByVal bIncludeHeader As Boolean) As Long

Dim PicBits() As Byte

    If bIncludeHeader Then
        If TVersion > 2 Then
            GetFrameSize = 10
        Else
            GetFrameSize = 6
        End If
    End If
    SavePicture AlbumArt, App.Path & "\Temp.bmp"
    ReDim PicBits(1 To FileLen(App.Path & "\Temp.bmp")) As Byte
    Open App.Path & "\Temp.bmp" For Binary As #1
    Get #1, 1, PicBits
    Close #1
    Kill App.Path & "\Temp.bmp"
    If TVersion > 2 Then
        GetFrameSize = GetFrameSize + UBound(PicBits) + 14
    Else
        GetFrameSize = GetFrameSize + UBound(PicBits) + 6
    End If

End Function

Private Function GetTagSize(ByVal fh As Long, _
                            ByVal fp As Long) As Long

Dim TagHeader As v2TagHeader

    SetFilePointer fh, fp, 0, FILE_BEGIN
    ReadFile fh, TagHeader, Len(TagHeader), 0, ByVal 0
    If Data2String(VarPtr(TagHeader.Identifier(0)), 3, ENC_ISO) = "ID3" Then
        GetTagSize = Data2Long(VarPtr(TagHeader.Size(0)), True) + Len(TagHeader)
        If TagHeader.Version(0) >= 4 Then
            If TagHeader.flags And &H10& Then
                GetTagSize = GetTagSize + Len(TagHeader)
            End If
        End If
    End If
    TVersion = TagHeader.Version(0)

End Function

Public Function ID3Exist(ByVal Filename As String) As Boolean

Dim fh         As Long

    fh = CreateFile(Filename, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If Not fh = INVALID_HANDLE_VALUE Then
        If GetTagSize(fh, 0) > 0 Then
            ID3Exist = True
        End If
        CloseHandle fh
    End If

End Function

Private Sub IgnoreReadOnly(ByVal Filename As String)

Dim rdAttr As Long

    rdAttr = GetFileAttributes(Filename)
    If Not rdAttr = -1 Then
        If rdAttr And FILE_ATTRIBUTE_READONLY Then
            rdAttr = rdAttr Xor FILE_ATTRIBUTE_READONLY
            SetFileAttributes Filename, rdAttr
        End If
    End If

End Sub

Private Sub Long2Data(ByVal SrcValue As Long, _
                      ByVal bSynchSafe As Boolean, _
                      ByVal pData As Long)

Dim Data(3) As Byte

    If bSynchSafe Then
        Data(0) = (SrcValue And &HFE00000) \ &H200000
        Data(1) = (SrcValue And &H1FC000) \ &H4000&
        Data(2) = (SrcValue And &H3F80&) \ &H80&
        Data(3) = SrcValue And &H7F&
    Else
        Data(0) = SrcValue \ &H1000000
        Data(1) = (SrcValue And &HFF0000) \ &H10000
        Data(2) = (SrcValue And &HFF00&) \ &H100&
        Data(3) = SrcValue And &HFF&
    End If
    CopyMemory ByVal pData, Data(0), 4

End Sub

Public Function ReadAlbumArt(ByVal Filename As String, _
                             ByVal FrameIndex As Long, _
                             Optional ByRef AlbumArt As IPictureDisp, _
                             Optional ByRef PictureType As Long) As Boolean

Dim fh         As Long
Dim ExistSize  As Long
Dim pbuffer    As Long
Dim id3v2tag() As Byte
Dim tp         As Long
Dim SizeBuf(3) As Byte
Dim i          As Long
Dim j          As Long
Dim tp2        As Long

    TFrameSize = 0
    TData = 0
    fh = CreateFile(Filename, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then
        Exit Function
    End If
    ExistSize = GetTagSize(fh, 0)
    If ExistSize > 0 Then
        ReDim id3v2tag(ExistSize - 1) As Byte
        pbuffer = VarPtr(id3v2tag(0))
        SetFilePointer fh, 10, 0, FILE_BEGIN
        ReadFile fh, ByVal pbuffer, ExistSize - 10, 0, ByVal 0
        If TVersion > 2 Then
            Do Until tp >= ExistSize - 11
                If Data2String(VarPtr(id3v2tag(tp)), 4, ENC_ISO) = "APIC" Then
                    tp2 = tp2 + 1
                    If tp2 = FrameIndex Then
                        TFrameSize = Data2Long(VarPtr(id3v2tag(tp + 4)), IIf(TVersion = 3, False, True))
                        TData = tp + 10
                        Exit Do
                    End If
                End If
                tp = tp + Data2Long(VarPtr(id3v2tag(tp + 4)), IIf(TVersion = 3, False, True)) + 10
            Loop
            If TFrameSize > 0 Then
                v4_Check tp, id3v2tag
                If TVersion = 4 And rsize < TFrameSize Then
                    For i = tp + 19 To tp + TFrameSize
                        If id3v2tag(i) = 0 Then
                            Exit For
                        End If
                    Next i
                Else
                    For i = tp + 11 To tp + TFrameSize
                        If id3v2tag(i) = 0 Then
                            Exit For
                        End If
                    Next i
                End If
                PictureType = id3v2tag(i + 1)
                For j = i + 2 To tp + TFrameSize Step 2
                    If id3v2tag(j) = 0 Then
                        Exit For
                    End If
                Next j
                If j = i + 2 Then
                    i = j - tp - 9
                Else
                    i = j - tp - 8
                End If
                If TVersion = 4 Then
                    If rsize < TFrameSize Then
                        rsize = rsize - i + 4
                    End If
                End If
                ReadAlbumArt = ReadFrame(VarPtr(id3v2tag(TData + i)), TFrameSize - i, AlbumArt)
                TFrameSize = TFrameSize + 10
            End If
        Else
            Do Until tp >= ExistSize - 7
                CopyMemory SizeBuf(1), id3v2tag(tp + 3), 3
                If Data2String(VarPtr(id3v2tag(tp)), 3, ENC_ISO, False) = "PIC" Then
                    tp2 = tp2 + 1
                    If tp2 = FrameIndex Then
                        TFrameSize = Data2Long(VarPtr(SizeBuf(0)), False)
                        TData = tp + 10
                        Exit Do
                    End If
                End If
                tp = tp + Data2Long(VarPtr(SizeBuf(0)), False) + 6
            Loop
            If TFrameSize > 0 Then
                PictureType = id3v2tag(tp + 10)
                For i = tp + 11 To tp + TFrameSize Step 2
                    If id3v2tag(i) = 0 Then
                        Exit For
                    End If
                Next i
                If i = tp + 11 Then
                    i = i - tp - 11
                Else
                    i = i - tp - 10
                End If
                ReadAlbumArt = ReadFrame(VarPtr(id3v2tag(tp + i + 6)), TFrameSize - i, AlbumArt)
                TFrameSize = TFrameSize + 6
            End If
        End If
    End If
    CloseHandle fh

End Function

Private Function ReadFrame(ByVal pData As Long, _
                           ByVal Framesize As Long, _
                           ByRef AlbumArt As IPictureDisp) As Boolean

Dim Data()    As Byte
Dim TmpData() As Byte
Dim i         As Long
Dim j         As Long

    If TVersion > 2 Then
        If TVersion = 4 And rsize < TFrameSize Then
            ReDim TmpData(Framesize - 1) As Byte
            CopyMemory TmpData(0), ByVal pData, Framesize
            ReDim Data(rsize - 1) As Byte
            Do Until j > UBound(TmpData)
                If j < UBound(TmpData) Then
                    If TmpData(j) = 255 And TmpData(j + 1) = 0 Then
                        Data(i) = 255
                        j = j + 2
                    Else
                        Data(i) = TmpData(j)
                        j = j + 1
                    End If
                Else
                    Data(i) = TmpData(j)
                    j = j + 1
                End If
                i = i + 1
            Loop
        Else
            ReDim Data(Framesize - 1) As Byte
            CopyMemory Data(0), ByVal pData, Framesize
        End If
    Else
        ReDim Data(Framesize - 7) As Byte
        CopyMemory Data(0), ByVal pData + 6, Framesize - 6
    End If
    Set AlbumArt = ArrayToPicture(Data)
    ReadFrame = True

End Function

Private Function String2Data(ByVal SrcStr As String, _
                             ByVal MaxLength As Long, _
                             ByVal pData As Long, _
                             ByVal bTerminate As Boolean) As Long

Dim i      As Long
Dim SrcLen As Long
Dim curAsc As Byte

    SrcLen = Len(SrcStr)
    If MaxLength <= 0 Then
        MaxLength = SrcLen
    ElseIf MaxLength > SrcLen Then
        MaxLength = SrcLen
    End If
    For i = 0 To MaxLength - 1
        curAsc = Asc(Mid$(SrcStr, i + 1, 1))
        CopyMemory ByVal pData + i, curAsc, 1
    Next i
    If bTerminate Then
        CopyMemory ByVal pData + MaxLength, curAsc, 1
        String2Data = MaxLength + 1
    Else
        String2Data = MaxLength
    End If

End Function

Private Sub v4_Check(ByVal tp As Long, _
                     ByRef TagData() As Byte)

Dim TmpData As Long
Dim TmpPos  As Long
Dim i       As Long
Dim j       As Long

    rsize = Data2Long(VarPtr(TagData(TData)), IIf(TVersion = 3, False, True))
    If TVersion = 4 Then
        If rsize < TFrameSize Then
            For i = tp + 19 To tp + TFrameSize
                If TagData(i) = 0 Then
                    Exit For
                End If
            Next i
            For j = i + 2 To tp + TFrameSize Step 2
                If TagData(j) = 0 Then
                    Exit For
                End If
            Next j
            If j = i + 2 Then
                i = j - tp - 9
            Else
                i = j - tp - 8
            End If
            TmpData = TFrameSize - i - 1
            TmpPos = TData + i
            rsize = rsize - i + 4
            i = 0
            j = 0
            Do Until j > TmpData
                If j < TmpData Then
                    If TagData(TmpPos + j) = 255 And TagData(TmpPos + 1 + j) = 0 Then
                        j = j + 2
                    Else
                        j = j + 1
                    End If
                Else
                    j = j + 1
                End If
                i = i + 1
            Loop
        End If
    End If
    If rsize <> i Then
        rsize = TFrameSize
    Else
        rsize = Data2Long(VarPtr(TagData(TData)), IIf(TVersion = 3, False, True))
    End If

End Sub

Public Function WriteAlbumArt(ByVal Filename As String, _
                              ByVal FrameIndex As Long, _
                              ByRef AlbumArt As IPictureDisp, _
                              Optional ByVal PictureType As Long = 3, _
                              Optional ByVal bIgnoreReadOnly As Boolean = True) As Boolean

Dim fh         As Long
Dim fsize      As Long
Dim tagsize    As Long
Dim TagHeader  As v2TagHeader
Dim ExistSize2 As Long
Dim ExistSize  As Long
Dim WrBuf()    As Byte
Dim pbuffer    As Long
Dim tp2        As Long

    ReadAlbumArt Filename, FrameIndex
    If TData = 0 Then
        tp2 = Len(TagHeader)
    Else
        tp2 = TData + TFrameSize
    End If
    If bIgnoreReadOnly Then
        IgnoreReadOnly Filename
    End If
    fh = CreateFile(Filename, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then
        Exit Function
    End If
    ExistSize = GetTagSize(fh, 0)
    If ExistSize = 0 Then
        fsize = GetFileSize(fh, 0)
        ReDim WrBuf(fsize + Len(TagHeader)) As Byte
        With TagHeader
            String2Data "ID3", 3, VarPtr(.Identifier(0)), False
            .Version(0) = 3
            Long2Data 0, True, VarPtr(.Size(0))
        End With
        CopyMemory WrBuf(0), TagHeader, Len(TagHeader)
        pbuffer = VarPtr(WrBuf(Len(TagHeader)))
        SetFilePointer fh, 0, 0, FILE_BEGIN
        ReadFile fh, ByVal pbuffer, fsize, 0, ByVal 0
        SetFilePointer fh, fsize, 0, FILE_BEGIN
        SetEndOfFile fh
        pbuffer = VarPtr(WrBuf(0))
        SetFilePointer fh, 0, 0, FILE_BEGIN
        WriteFile fh, ByVal pbuffer, fsize + Len(TagHeader), 0, ByVal 0
        CloseHandle fh
        If bIgnoreReadOnly Then
            IgnoreReadOnly Filename
        End If
        fh = CreateFile(Filename, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
        If fh = INVALID_HANDLE_VALUE Then
            Exit Function
        End If
    End If
    ExistSize = GetTagSize(fh, 0)
    fsize = GetFileSize(fh, 0)
    ExistSize2 = GetFrameSize(AlbumArt, True)
    tagsize = ExistSize + ExistSize2 - Len(TagHeader)
    ReDim WrBuf(fsize + ExistSize2) As Byte
    pbuffer = VarPtr(WrBuf(0))
    SetFilePointer fh, 0, 0, FILE_BEGIN
    ReadFile fh, ByVal pbuffer, tp2, 0, ByVal 0
    WriteFrame VarPtr(WrBuf(tp2)), AlbumArt, PictureType
    pbuffer = VarPtr(WrBuf(tp2 + ExistSize2))
    SetFilePointer fh, tp2, 0, FILE_BEGIN
    ReadFile fh, ByVal pbuffer, fsize - tp2, 0, ByVal 0
    Long2Data tagsize, True, VarPtr(WrBuf(6))
    fsize = fsize + ExistSize2
    SetFilePointer fh, fsize, 0, FILE_BEGIN
    SetEndOfFile fh
    pbuffer = VarPtr(WrBuf(0))
    SetFilePointer fh, 0, 0, FILE_BEGIN
    WriteFile fh, ByVal pbuffer, fsize, 0, ByVal 0
    CloseHandle fh
    WriteAlbumArt = True

End Function

Private Sub WriteFrame(ByVal pData As Long, _
                       Optional ByVal AlbumArt As IPictureDisp, _
                       Optional ByVal PictureType As Long = 3)

Dim fraHeader_34 As v2_34FrameHeader
Dim fraHeader_2  As v2_2FrameHeader
Dim PicBits()    As Byte
Dim ttemp(14)    As Byte
Dim SizeBuf(3)   As Byte

    If TVersion > 2 Then
        String2Data "APIC", 4, VarPtr(fraHeader_34.Ident(0)), False
        Long2Data GetFrameSize(AlbumArt, False), IIf(TVersion = 3, False, True), VarPtr(fraHeader_34.Size(0))
        CopyMemory ByVal pData, fraHeader_34, Len(fraHeader_34)
        SavePicture AlbumArt, App.Path & "\Temp.bmp"
        ReDim PicBits(1 To FileLen(App.Path & "\Temp.bmp")) As Byte
        Open App.Path & "\Temp.bmp" For Binary As #1
        Get #1, 1, PicBits
        Close #1
        Kill App.Path & "\Temp.bmp"
        String2Data "image/jpeg", 10, VarPtr(ttemp(1)), False
        If PictureType < 0 Or PictureType > 20 Then
            PictureType = 3
        End If
        ttemp(12) = PictureType
        CopyMemory ByVal pData + 11, ttemp(1), 14
        CopyMemory ByVal pData + 24, PicBits(1), UBound(PicBits)
    Else
        String2Data "PIC", 3, VarPtr(fraHeader_2.Ident(0)), False
        Long2Data GetFrameSize(AlbumArt, False), False, VarPtr(SizeBuf(0))
        CopyMemory fraHeader_2.Size(0), SizeBuf(1), 3
        CopyMemory ByVal pData, fraHeader_2, Len(fraHeader_2)
        SavePicture AlbumArt, App.Path & "\Temp.bmp"
        ReDim PicBits(1 To FileLen(App.Path & "\Temp.bmp")) As Byte
        Open App.Path & "\Temp.bmp" For Binary As #1
        Get #1, 1, PicBits
        Close #1
        Kill App.Path & "\Temp.bmp"
        String2Data "JPG", 3, VarPtr(ttemp(2)), False
        If PictureType < 0 Or PictureType > 20 Then
            PictureType = 3
        End If
        ttemp(5) = PictureType
        CopyMemory ByVal pData + 6, ttemp(1), 6
        CopyMemory ByVal pData + 12, PicBits(1), UBound(PicBits)
    End If

End Sub
