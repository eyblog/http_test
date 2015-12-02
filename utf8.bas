Attribute VB_Name = "utf8"
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001

'Purpose:Convert Utf8 to Unicode
Public Function UTF8_Decode(ByVal sUTF8 As String) As String
   Dim lngUtf8Size      As Long
   Dim strBuffer        As String
   Dim lngBufferSize    As Long
   Dim lngResult        As Long
   Dim bytUtf8()        As Byte
   Dim n                As Long
   If LenB(sUTF8) = 0 Then Exit Function
      On Error GoTo EndFunction
      bytUtf8 = StrConv(sUTF8, vbFromUnicode)
      lngUtf8Size = UBound(bytUtf8) + 1
      On Error GoTo 0
      lngBufferSize = lngUtf8Size * 2
      strBuffer = String$(lngBufferSize, vbNullChar)
      'Translate using code page 65001(UTF-8)
      lngResult = MultiByteToWideChar(CP_UTF8, 0, bytUtf8(0), _
         lngUtf8Size, StrPtr(strBuffer), lngBufferSize)
      'Trim result to actual length
      If lngResult Then
         UTF8_Decode = Left$(strBuffer, lngResult)
      End If
EndFunction:
End Function

