Attribute VB_Name = "modUtils"
' //
' // modUtils.bas - utility functions
' // by The trick, 2022
' //

Option Explicit

Private m_tBytesParse(63)   As Byte
Private m_tParseNum         As NUMPARSE

Public Function FAR_PROC( _
                ByVal pfn As PTR) As PTR
    FAR_PROC = pfn
End Function

Public Function GetOpenFile( _
                ByVal hWnd As Long, _
                ByRef sTitle As String, _
                ByRef sFilter As String) As String
    Dim tOFN            As OPENFILENAME
    Dim strInputFile    As String

    With tOFN
    
        .nMaxFile = 260
        strInputFile = String$(.nMaxFile, vbNullChar)
        .hwndOwner = hWnd
        .lpstrTitle = StrPtr(sTitle)
        .lpstrFile = StrPtr(strInputFile)
        .lStructSize = Len(tOFN)
        .lpstrFilter = StrPtr(sFilter)
        
        If GetOpenFileName(tOFN) = 0 Then Exit Function
        
        GetOpenFile = Left$(strInputFile, InStr(1, strInputFile, vbNullChar) - 1)
        
    End With

End Function

Public Function GetSaveFile( _
                ByVal hWnd As Long, _
                ByRef sTitle As String, _
                ByRef sFilter As String, _
                ByRef sDefExtension As String) As String
    Dim tOFN            As OPENFILENAME
    Dim strOutputFile   As String
    
    With tOFN
    
        .nMaxFile = 260
        strOutputFile = String$(.nMaxFile, vbNullChar)
        .hwndOwner = hWnd
        .lpstrTitle = StrPtr(sTitle)
        .lpstrFile = StrPtr(strOutputFile)
        .lStructSize = Len(tOFN)
        .lpstrFilter = StrPtr(sFilter)
        .lpstrDefExt = StrPtr(sDefExtension)
        .Flags = OFN_EXPLORER Or _
                 OFN_ENABLESIZING Or OFN_NOREADONLYRETURN Or OFN_PATHMUSTEXIST Or _
                 OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
                 
        If GetSaveFileName(tOFN) = 0 Then Exit Function
        
        GetSaveFile = Left$(strOutputFile, InStr(1, strOutputFile, vbNullChar) - 1)
        
    End With

End Function

Public Function StrToDbl( _
                ByRef sValue As String, _
                ByRef dOut As Double) As Long
    Dim hr      As Long
    Dim vRet    As Variant
    
    m_tParseNum.cDig = UBound(m_tBytesParse) + 1
    m_tParseNum.dwInFlags = NUMPRS_LEADING_MINUS Or NUMPRS_DECIMAL
    
    hr = VarParseNumFromStr(sValue, GetUserDefaultLCID, 0, m_tParseNum, m_tBytesParse(0))
    If hr < 0 Then
        StrToDbl = hr
        Exit Function
    ElseIf m_tParseNum.cchUsed <> Len(sValue) Then
        StrToDbl = E_FAIL
        Exit Function
    End If
    
    hr = VarNumFromParseNum(m_tParseNum, m_tBytesParse(0), VTBIT_R8, vRet)
    If hr < 0 Then
        StrToDbl = hr
        Exit Function
    End If
    
    dOut = vRet
    
End Function

Public Function StrToLng( _
                ByRef sValue As String, _
                ByVal bAcceptDecimalPt As Boolean, _
                ByRef lOut As Long) As Long
    Dim hr      As Long
    Dim vRet    As Variant
    
    m_tParseNum.cDig = UBound(m_tBytesParse) + 1
    m_tParseNum.dwInFlags = NUMPRS_LEADING_MINUS Or (bAcceptDecimalPt And NUMPRS_DECIMAL)
    
    hr = VarParseNumFromStr(sValue, GetUserDefaultLCID, 0, m_tParseNum, m_tBytesParse(0))
    If hr < 0 Then
        StrToLng = hr
        Exit Function
    ElseIf m_tParseNum.cchUsed <> Len(sValue) Then
        StrToLng = E_FAIL
        Exit Function
    End If
    
    hr = VarNumFromParseNum(m_tParseNum, m_tBytesParse(0), VTBIT_I4, vRet)
    If hr < 0 Then
        StrToLng = hr
        Exit Function
    End If
    
    lOut = vRet

End Function

