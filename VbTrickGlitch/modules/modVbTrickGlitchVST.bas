Attribute VB_Name = "modVbTrickGlitchVST"

' //
' // modVbTrickGlitchVST.bas - global functions for VbTrickGlitch plugin usage
' // You can call this function from initialized state (STA thread / initialized runtime / context).
' // by The trick, 2022
' //

Option Explicit
Option Base 0

' // VST_PLUGIN_CLASS_NAME - name of plugin class.
' // Code uses it to create class from dll
Public Const VST_PLUGIN_CLASS_NAME  As String = "CVbTrickGlitchVST"

' // This is used for display info about plugin in host
Public Const VST_PLUGIN_NAME        As String = "VbTrickGlitchVST"
Public Const VST_PRODUCT_STRING     As String = "VbTrickGlitchVST VST plugin by The trick"
Public Const VST_VENDOR_STRING      As String = "The trick"
Public Const VST_VENDOR_VERSION     As Long = 1

Public Const NUM_OF_PARAMETERS      As Long = 4             ' // Number of parameters
Public Const NUM_OF_PROGRAMS        As Long = 32            ' // Number of programs
Public Const UNIQUE_ID              As Long = &HED9A34B0    ' // "registered unique identifier (register it at Steinberg 3rd party support Web).
                                                            ' // This is used to identify a plug-in during save+load of preset and project."
                                                            ' // I didn't register ;)
Public Const VST_PLUGIN_VERSION     As Long = 1

' // List of parameters
Public Enum eParameterType
    PT_SPEED = 0
    PT_PITCH = 1
    PT_SLOT = 2
    PT_SMOOTH = 3
End Enum

' // When a parameter has been changed it sets specified flags to update GUI and vice versa
Public Enum eStateChangedMask
    SCM_SPEED = 1
    SCM_PITCH = 2
    SCM_SMOOTH = 4
    SCM_PROGRAM = 8
    SCM_SLOT_PLAYBACK = &H10
    SCM_SLOT_CURRENT = &H20
    SCM_SLOT_ACTIVE = &H40
End Enum

' // Piano roll key (note) item
Public Type tKeyItem
    lValue      As Long
    dPos        As Double       ' // Position in quarter-notes
    dLength     As Double       ' // Length in quarter-notes
End Type

' // Piano roll pattern
Public Type tPattern
    lLengthPerBeats As Long
    lNumOfKeys      As Long
    tKeys()         As tKeyItem
End Type

' // Slot description
Public Type tSlot
    tPattern    As tPattern
    fPitch      As Single
    fSmooth     As Single
    fSpeed      As Single
End Type

' // Preset
Public Type tPreset
    sName           As String
    tSlots(39)      As tSlot
End Type

' // Shared data. Class and GUI use the same copy of this data
Public Type tSharedData
    lNumOfPresets   As Long
    lCurPreset      As Long
    tPresets()      As tPreset
    fPitchBend      As Single               ' // Current MIDI pitch-bend value
    lPlaybackSlot   As Long                 ' // Current playing slot
    lCurrentSlot    As Long                 ' // Current slot which is value of automation
    lActiveSlot     As Long                 ' // Active slot (which user sees and edits)
    eChStateEffect  As eStateChangedMask    ' // Changing mask class->UI
    eChStateUI      As eStateChangedMask    ' // Changing mask UI->class (automation writing)
    dPlaybackPos    As Double
    bRecordMode     As Boolean              ' // True - recording automation
End Type

' // Simple byte stream
Private Type tByteStream
    bData() As Byte
    lPos    As Long
    lSize   As Long
End Type

' // Serialize preset(s). It's used when host save state
Public Function SerializePresets( _
                ByRef tPresets() As tPreset, _
                ByVal lStartIndex As Long, _
                ByVal lCount As Long, _
                ByRef bOut() As Byte) As Boolean
    Dim tStm    As tByteStream
    Dim lIndex  As Long
    Dim lSIndex As Long
    Dim lKIndex As Long
    
    If Not StmWrite(tStm, VarPtr(lCount), 4) Then
        Exit Function
    End If
    
    For lIndex = lStartIndex To lStartIndex + lCount - 1
        
        With tPresets(lIndex)
        
            If Not StmWrite(tStm, StrPtr(.sName), LenB(.sName) + 2) Then
                Exit Function
            End If
            
            For lSIndex = 0 To UBound(.tSlots)
                
                With .tSlots(lSIndex)
                    
                    If Not StmWrite(tStm, VarPtr(.fPitch), 4) Then
                        Exit Function
                    End If
                    
                    If Not StmWrite(tStm, VarPtr(.fSpeed), 4) Then
                        Exit Function
                    End If
                    
                    If Not StmWrite(tStm, VarPtr(.fSmooth), 4) Then
                        Exit Function
                    End If
                    
                    If Not StmWrite(tStm, VarPtr(.tPattern.lLengthPerBeats), 4) Then
                        Exit Function
                    End If
                    
                    If Not StmWrite(tStm, VarPtr(.tPattern.lNumOfKeys), 4) Then
                        Exit Function
                    End If
                    
                    For lKIndex = 0 To .tPattern.lNumOfKeys - 1
                        If Not StmWrite(tStm, VarPtr(.tPattern.tKeys(lKIndex)), Len(.tPattern.tKeys(lKIndex))) Then
                            Exit Function
                        End If
                    Next
                    
                End With
            
            Next
        
        End With
        
    Next
    
    If tStm.lSize > 0 Then
        ReDim Preserve tStm.bData(tStm.lSize - 1)
    Else
        Erase tStm.bData
    End If
    
    bOut = tStm.bData
    
    SerializePresets = True
    
End Function

' // Deserialize presets. It's used to restore saved state
Public Function DeserializePresets( _
                ByVal pData As PTR, _
                ByVal lSize As Long, _
                ByRef tOut() As tPreset) As Boolean
    Dim tSADesc As SAFEARRAY1D
    Dim tStm    As tByteStream
    Dim lPCount As Long
    Dim lKCount As Long
    Dim lPIndex As Long
    Dim lSIndex As Long
    Dim lKIndex As Long
    
    With tSADesc
        .cbElements = 1
        .cDims = 1
        .fFeatures = FADF_AUTO
        .rgsabound(0).cElements = lSize
        .pvData = pData
    End With
    
    PutMemPtr ByVal ArrPtr(tStm.bData), VarPtr(tSADesc)

    tStm.lSize = lSize
    
    If Not StmRead(tStm, VarPtr(lPCount), 4) Then
        Exit Function
    End If
    
    If lPCount <= 0 Then
        Exit Function
    End If
    
    ReDim tOut(lPCount - 1)
    
    For lPIndex = 0 To lPCount - 1
        
        With tOut(lPIndex)
            
            If Not StmReadString(tStm, .sName) Then
                Exit Function
            End If
            
            For lSIndex = 0 To UBound(.tSlots)
                
                With .tSlots(lSIndex)
                    
                    If Not StmRead(tStm, VarPtr(.fPitch), 4) Then
                        Exit Function
                    End If
                    
                    If Not StmRead(tStm, VarPtr(.fSpeed), 4) Then
                        Exit Function
                    End If
                    
                    If Not StmRead(tStm, VarPtr(.fSmooth), 4) Then
                        Exit Function
                    End If
                    
                    If Not StmRead(tStm, VarPtr(.tPattern.lLengthPerBeats), 4) Then
                        Exit Function
                    End If
                    
                    If Not StmRead(tStm, VarPtr(.tPattern.lNumOfKeys), 4) Then
                        Exit Function
                    End If
                    
                    If .tPattern.lNumOfKeys > 0 Then
                        ReDim .tPattern.tKeys(.tPattern.lNumOfKeys - 1)
                    End If
                    
                    For lKIndex = 0 To .tPattern.lNumOfKeys - 1
                        If Not StmRead(tStm, VarPtr(.tPattern.tKeys(lKIndex)), Len(.tPattern.tKeys(lKIndex))) Then
                            Exit Function
                        End If
                    Next
                    
                End With
                
            Next
        
        End With
        
    Next
    
    DeserializePresets = True
    
End Function

' // Stream functions //

Private Function StmRead( _
                 ByRef tStm As tByteStream, _
                 ByVal pData As PTR, _
                 ByVal lSize As Long) As Boolean
                     
    If tStm.lPos + lSize > tStm.lSize Then
        Exit Function
    End If
    
    memcpy ByVal pData, tStm.bData(tStm.lPos), lSize
    
    tStm.lPos = tStm.lPos + lSize
    
    StmRead = True
                     
End Function

Private Function StmReadString( _
                 ByRef tStm As tByteStream, _
                 ByRef sOut As String) As Boolean
    Dim lSize   As Long
    Dim lIndex  As Long
    
    For lIndex = tStm.lPos To tStm.lSize - 2 Step 2
        
        If tStm.bData(lIndex) = 0 And tStm.bData(lIndex + 1) = 0 Then
            Exit For
        End If
        
        lSize = lSize + 1
        
    Next
    
    If lIndex > tStm.lSize - 2 Then
        Exit Function
    End If
    
    If lSize > 0 Then
        sOut = Space$(lSize)
        memcpy ByVal StrPtr(sOut), tStm.bData(tStm.lPos), lSize * 2
    Else
        sOut = vbNullString
    End If
    
    tStm.lPos = tStm.lPos + lSize * 2 + 2
    
    StmReadString = True
    
End Function

Private Function StmWrite( _
                 ByRef tStm As tByteStream, _
                 ByVal pData As PTR, _
                 ByVal lSize As Long) As Boolean
    
    On Error GoTo error_handler
    
    If tStm.lPos Then
        If tStm.lPos + lSize > UBound(tStm.bData) Then
            ReDim Preserve tStm.bData((UBound(tStm.bData) + 1) * 2 - 1)
        End If
    Else
        ReDim tStm.bData(255)
    End If
                     
    memcpy tStm.bData(tStm.lPos), ByVal pData, lSize
    
    tStm.lPos = tStm.lPos + lSize
    tStm.lSize = tStm.lSize + lSize
    
    StmWrite = True
    
error_handler:
    
End Function

