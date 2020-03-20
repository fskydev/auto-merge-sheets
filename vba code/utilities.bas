Attribute VB_Name = "utilities"
' $Id: basUtf8FromString $

' Written by David Ireland DI Management Services Pty Limited 2015
' <http://www.di-mgt.com.au> <http://www.cryptosys.net>
' @license MIT license <http://opensource.org/licenses/MIT>
' [2015-06-30] First published.
' [2018-07-27] Updated to cope with empty or null input string.
' [2018-08-15] Added Utf8BytesToString and BytesLength functions.
' [2018-11-06] Changed vbNull argument to 0&

Option Explicit

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
    
''' Maps a character string to a UTF-16 (wide character) string
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long _
    ) As Long

' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

''' Return length of byte array or zero if uninitialized
Private Function BytesLength(abBytes() As Byte) As Long
    ' Trap error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
End Function


''' Return byte array with VBA "Unicode" string encoded in UTF-8
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

''' Return VBA "Unicode" string from byte array encoded in UTF-8
Public Function Utf8BytesToString(abUtf8Array() As Byte) As String
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
    Utf8BytesToString = ""
    ' Catch uninitialized input array
    nBytes = BytesLength(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)
End Function
Public Function Utf8StringFromUtf8Bytes(indexHeader As Integer) As String
    Dim i As Integer
    
    Select Case indexHeader
        Case 1:
            Dim colHeader1(11) As Byte
            For i = 0 To 11
                colHeader1(i) = Choose(i + 1, &HE3, &H83, &H96, &HE3, &H83, &H83, &HE3, &H82, &HAF, &HE5, &H90, &H8D)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader1)
            
        Case 2:
            Dim colHeader35(11) As Byte
            For i = 0 To 11
                colHeader35(i) = Choose(i + 1, &HE3, &H82, &HB7, &HE3, &H83, &HBC, &HE3, &H83, &H88, &HE5, &H90, &H8D)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader35)
            
        
        Case 3:
            Dim colHeader2(8) As Byte
            For i = 0 To 8
                colHeader2(i) = Choose(i + 1, &HE4, &HBC, &H81, &HE6, &HA5, &HAD, &HE5, &H90, &H8D)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader2)
        Case 4:
            Dim colHeader3(14) As Byte
            For i = 0 To 14
                colHeader3(i) = Choose(i + 1, &HE6, &H9C, &HAC, &HE7, &HA4, &HBE, &HE6, &H89, &H80, &HE5, &H9C, &HA8, &HE5, &H9C, &HB0)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader3)
        Case 5:
            Utf8StringFromUtf8Bytes = "URL"
        Case 6:
            Dim colHeader5(11) As Byte
            For i = 0 To 11
                colHeader5(i) = Choose(i + 1, &HE8, &HA8, &HAD, &HE7, &HAB, &H8B, &HE5, &HB9, &HB4, &HE6, &H9C, &H88)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader5)
        Case 7:
            Dim colHeader6(8) As Byte
            For i = 0 To 8
                colHeader6(i) = Choose(i + 1, &HE8, &HB3, &H87, &HE6, &H9C, &HAC, &HE9, &H87, &H91)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader6)
        Case 8:
            Dim colHeader7(11) As Byte
            For i = 0 To 11
                colHeader7(i) = Choose(i + 1, &HE5, &HBE, &H93, &HE6, &HA5, &HAD, &HE5, &H93, &HA1, &HE6, &H95, &HB0)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader7)
        Case 9:
            Dim colHeader8(11) As Byte
            For i = 0 To 11
                colHeader8(i) = Choose(i + 1, &HE6, &HA0, &HAA, &HE5, &HBC, &H8F, &HE5, &H85, &HAC, &HE9, &H96, &H8B)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader8)
        Case 10:
            Dim colHeader9(11) As Byte
            For i = 0 To 11
                colHeader9(i) = Choose(i + 1, &HE4, &HBA, &H8B, &HE6, &HA5, &HAD, &HE5, &H86, &H85, &HE5, &HAE, &HB9)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader9)
        Case 11:
            Dim colHeader10(5) As Byte
            For i = 0 To 5
                colHeader10(i) = Choose(i + 1, &HE6, &HA5, &HAD, &HE7, &HA8, &HAE)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader10)
        Case 12:
            Dim colHeader11(8) As Byte
            For i = 0 To 8
                colHeader11(i) = Choose(i + 1, &HE4, &HBB, &HA3, &HE8, &HA1, &HA8, &HE8, &H80, &H85)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader11)
        Case 13:
            Dim colHeader12(8) As Byte
            For i = 0 To 8
                colHeader12(i) = Choose(i + 1, &HE5, &HA3, &HB2, &HE4, &HB8, &H8A, &HE9, &HAB, &H98)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader12)
        Case 14:
            Dim colHeader13(14) As Byte
            For i = 0 To 14
                colHeader13(i) = Choose(i + 1, &HE3, &H83, &H9D, &HE3, &H82, &HB8, &HE3, &H82, &HB7, &HE3, &H83, &HA7, &HE3, &H83, &HB3)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader13)
        Case 15:
            Dim colHeader14(11) As Byte
            For i = 0 To 11
                colHeader14(i) = Choose(i + 1, &HE9, &H85, &H8D, &HE5, &HB1, &H9E, &HE9, &H83, &HA8, &HE7, &HBD, &HB2)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader14)
        Case 16:
            Dim colHeader15(17) As Byte
            For i = 0 To 17
                colHeader15(i) = Choose(i + 1, &HE9, &H85, &H8D, &HE5, &HB1, &H9E, &HE9, &H83, &HA8, &HE7, &HBD, &HB2, &HE8, &HA9, &HB3, &HE7, &HB4, &HB0)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader15)
        Case 17:
            Dim colHeader16(11) As Byte
            For i = 0 To 11
                colHeader16(i) = Choose(i + 1, &HE9, &H9B, &H87, &HE7, &H94, &HA8, &HE5, &HBD, &HA2, &HE6, &H85, &H8B)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader16)
        Case 18:
            Dim colHeader17(11) As Byte
            For i = 0 To 11
                colHeader17(i) = Choose(i + 1, &HE8, &HA9, &HA6, &HE7, &H94, &HA8, &HE6, &H9C, &H9F, &HE9, &H96, &H93)
            Next  'byte array
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader17)
        Case 19:
            Dim colHeader18(11) As Byte
            For i = 0 To 11
                colHeader18(i) = Choose(i + 1, &HE5, &H8B, &H9F, &HE9, &H9B, &H86, &HE8, &H83, &H8C, &HE6, &H99, &HAF)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader18)
        Case 20:
            Dim colHeader19(11) As Byte
            For i = 0 To 11
                colHeader19(i) = Choose(i + 1, &HE6, &H8E, &HA1, &HE7, &H94, &HA8, &HE4, &HBA, &HBA, &HE6, &H95, &HB0)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader19)
        Case 21:
            Dim colHeader20(11) As Byte
            For i = 0 To 11
                colHeader20(i) = Choose(i + 1, &HE4, &HBB, &H95, &HE4, &HBA, &H8B, &HE5, &H86, &H85, &HE5, &HAE, &HB9)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader20)
        Case 22:
            Dim colHeader21(11) As Byte
            For i = 0 To 11
                colHeader21(i) = Choose(i + 1, &HE5, &HBF, &H85, &HE9, &HA0, &H88, &HE8, &HA6, &H81, &HE4, &HBB, &HB6)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader21)
        Case 23:
            Dim colHeader22(12) As Byte
            For i = 0 To 12
                colHeader22(i) = Choose(i + 1, &HE6, &HAD, &H93, &HE8, &HBF, &H8E, &H2F, &HE5, &HB0, &H9A, &HE5, &H8F, &HAF)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader22)
        Case 24:
            Dim colHeader23(20) As Byte
            For i = 0 To 20
                colHeader23(i) = Choose(i + 1, &HE5, &H85, &HA5, &HE7, &HA4, &HBE, &HE6, &H99, &H82, &HE6, &H83, &HB3, &HE5, &HAE, &H9A, &HE5, &HB9, &HB4, &HE5, &H8F, &H8E)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader23)
        Case 25:
            Dim colHeader24(11) As Byte
            For i = 0 To 11
                colHeader24(i) = Choose(i + 1, &HE5, &HB0, &HB1, &HE6, &HA5, &HAD, &HE6, &H99, &H82, &HE9, &H96, &H93)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader24)
        Case 26:
            Dim colHeader25(11) As Byte
            For i = 0 To 11
                colHeader25(i) = Choose(i + 1, &HE8, &HB3, &H83, &HE9, &H87, &H91, &HE5, &H88, &HB6, &HE5, &HBA, &HA6)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader25)
        Case 27:
            Dim colHeader26(14) As Byte
            For i = 0 To 14
                colHeader26(i) = Choose(i + 1, &HE6, &H99, &H82, &HE9, &H96, &H93, &HE5, &HA4, &H96, &HE5, &H8A, &HB4, &HE5, &H83, &H8D)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader26)
        Case 28:
            Dim colHeader27(14) As Byte
            For i = 0 To 14
                colHeader27(i) = Choose(i + 1, &HE8, &HA3, &H81, &HE9, &H87, &H8F, &HE5, &H8A, &HB4, &HE5, &H83, &H8D, &HE5, &H88, &HB6)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader27)
        Case 29:
            Dim colHeader28(11) As Byte
            For i = 0 To 11
                colHeader28(i) = Choose(i + 1, &HE5, &HBE, &H85, &HE9, &H81, &H87, &HE6, &H9D, &HA1, &HE4, &HBB, &HB6)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader28)
        Case 30:
            Dim colHeader29(11) As Byte
            For i = 0 To 11
                colHeader29(i) = Choose(i + 1, &HE7, &HA6, &H8F, &HE5, &H88, &HA9, &HE5, &H8E, &H9A, &HE7, &H94, &H9F)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader29)
        Case 31:
            Dim colHeader30(17) As Byte
            For i = 0 To 17
                colHeader30(i) = Choose(i + 1, &HE9, &H81, &HB8, &HE8, &H80, &H83, &HE3, &H83, &H97, &HE3, &H83, &HAD, &HE3, &H82, &HBB, &HE3, &H82, &HB9)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader30)
        Case 32:
            Dim colHeader31(14) As Byte
            For i = 0 To 14
                colHeader31(i) = Choose(i + 1, &HE5, &H8B, &HA4, &HE5, &H8B, &H99, &HE5, &H9C, &HB0, &HE4, &HBD, &H8F, &HE6, &H89, &H80)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader31)
        Case 33:
            Dim colHeader32(14) As Byte
            For i = 0 To 14
                colHeader32(i) = Choose(i + 1, &HE8, &HBB, &HA2, &HE5, &H8B, &HA4, &HE3, &H81, &HAE, &HE6, &H9C, &H89, &HE7, &H84, &HA1)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader32)
        Case 34:
            Dim colHeader33(11) As Byte
            For i = 0 To 11
                colHeader33(i) = Choose(i + 1, &HE4, &HBC, &H91, &HE6, &H97, &HA5, &HE4, &HBC, &H91, &HE6, &H9A, &H87)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader33)
        Case 35:
            Dim colHeader34(11) As Byte
            For i = 0 To 11
                colHeader34(i) = Choose(i + 1, &HE4, &HBC, &H91, &HE6, &H86, &HA9, &HE6, &H99, &H82, &HE9, &H96, &H93)
            Next
            Utf8StringFromUtf8Bytes = Utf8BytesToString(colHeader34)
    End Select
End Function
Public Function Utf8YearLetter() As String
    Dim year(2) As Byte
    Dim i As Integer
    For i = 0 To 2
        year(i) = Choose(i + 1, &HE5, &HB9, &HB4)
    Next
    Utf8YearLetter = Utf8BytesToString(year)
End Function
Public Function Utf8MonthLetter() As String
    Dim month(2) As Byte
    Dim i As Integer
    For i = 0 To 2
        month(i) = Choose(i + 1, &HE6, &H9C, &H88)
    Next
    Utf8MonthLetter = Utf8BytesToString(month)
End Function
Function GetFileNames(ByVal FolderPath As String) As Variant
    Dim result As Variant
    Dim i As Integer
    Dim MyFile As Object
    Dim MyFSO As Object
    Dim MyFolder As Object
    Dim MyFiles As Object
    Set MyFSO = CreateObject("Scripting.FileSystemObject")
    Set MyFolder = MyFSO.GetFolder(FolderPath)
    Set MyFiles = MyFolder.Files
    If MyFiles.Count > 0 Then
    
        ReDim result(1 To MyFiles.Count)
        i = 1
        For Each MyFile In MyFiles
            If ((InStr(1, MyFile.Name, "xlsx") <> 0) Or (InStr(1, MyFile.Name, "xlsm") <> 0) Or (InStr(1, MyFile.Name, "xls") <> 0)) And (Mid(MyFile.Name, 1, 1) <> "~") Then
                result(i) = FolderPath + "\" + MyFile.Name
                i = i + 1
            End If
        Next MyFile
        If i > 1 Then
            ReDim Preserve result(1 To i - 1)
            GetFileNames = result
        End If
    End If
End Function
