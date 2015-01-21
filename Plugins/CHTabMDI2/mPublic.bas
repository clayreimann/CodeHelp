Attribute VB_Name = "mPublic"
Option Explicit

Public Enum enWinVersion
    enWin95 = 1
    enWinNT = 2
    enWin98 = 3
    enWin2000 = 4
    enWinXP = 5
End Enum

Function LoWord(lDWord As Long) As Integer

    If lDWord And &H8000& Then
        LoWord = lDWord Or &HFFFF0000
    Else
        LoWord = lDWord And &HFFFF&
    End If

End Function

Function HiWord(lDWord As Long) As Integer
    HiWord = (lDWord And &HFFFF0000) \ &H10000
End Function

Function GetWinText(hWnd As Long, Optional className As Boolean = False) As String
    'some static vars to speed up things, this func will be called many times
    Static sBuffer As String * 128& 'is it safe to use 128 bytes? should be enough..
    Static textLength As Long
  
    If className Then
        textLength = A_GetClassName(hWnd, sBuffer, 129&)
    Else
        textLength = A_GetWindowText(hWnd, sBuffer, 129&)
    End If
  
    If textLength > 0 Then
        GetWinText = Left$(sBuffer, textLength)
    End If

End Function

Function GetOSVersion() As enWinVersion
    'Get Windows version
    Dim tOS As A_OSVERSIONINFO
  
    tOS.dwOSVersionInfoSize = Len(tOS)
    A_GetVersionEx tOS
  
    If tOS.dwMajorVersion > 4& Then
        If tOS.dwMinorVersion > 0& Then
            GetOSVersion = enWinXP
        ElseIf tOS.dwMinorVersion = 0& Then
            GetOSVersion = enWin2000
        End If

    Else

        If tOS.dwPlatformId = 1& Then
            If tOS.dwMinorVersion > 0& Then
                GetOSVersion = enWin98
            Else
                GetOSVersion = enWin95
            End If

        ElseIf tOS.dwPlatformId = 2& Then
            GetOSVersion = enWinNT 'Should be check for NT 3.5 but we're not going that far
        End If
    End If

End Function

Public Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
' by Karl E. Peterson, http://www.mvps.org/vb, 20001207
  ' High word is coerced to Long to allow it to
  ' overflow limits of multiplication which shifts
  ' it left.
  MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function
