VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVbTrickGlitchVST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' //
' // CVbTrickGlitchVST.cls - main VST plugin class
' // by The trick, 2022
' //

Option Explicit

' // Pattern events
Private Enum eEventType
    ET_START_PATTERN
    ET_END_PATTERN
    ET_START_NOTE
    ET_END_NOTE
End Enum

Private Type tEvent
    eType   As eEventType
    lIndex  As Long
    dPos    As Double   ' // Event position
End Type

' // Current plugin state
Private Type tState
    dPhase          As Double   ' // Per samples
    lNote           As Long     ' // Current note
    bPlaying        As Boolean  ' // Is key down
    dSpeed          As Double   ' // Current speed
    dPitch          As Double   ' // Current pitch
    dSmooth         As Double   ' // Current smooth
    
    ' // Precalculations
    bIsFirst        As Boolean
    dFrequency      As Double   ' // Current frequency
    dSampleOffset   As Double   ' // Offset in samples of next wave
    dSamplesPerWave As Double   ' // Number of samples per waveform
    lSmoothSamples  As Long     ' // Number of samples to smooth
    lSmoothBufOfst  As Long     ' // Offset in buffer
    lCurveBufSize   As Long
    dSmoothCurve()  As Double   ' // Smooth curve
    
End Type

' // Plugin uses buffer to save samples (because you can change Speed value)
Private Type tBuffer
    lSize           As Long
    fSamplesLeft()  As Single
    fSamplesRight() As Single
    lWriteCursor    As Long
    dReadCursor     As Double
End Type

' // Always implement it
Implements IVBVstEffect

Private m_pAEffect          As PTR              ' // Pointer to AEffect structure
Private m_pfnHostCallback   As PTR              ' // Pointer to Audio master callback function
Private m_bEnabled          As Boolean          ' // Suspend/Resume
Private m_frmMain           As frmMain          ' // Form reference
Private m_ctlUI             As ctlUI            ' // GUI reference
Private m_tSharedData       As tSharedData      ' // Shared data between class and GUI
Private m_bEditorVisible    As Boolean
Private m_fSampleRate       As Single           ' // Current sample rate
Private m_tBuffer           As tBuffer          ' // Buffer
Private m_tState            As tState           ' // Current state
Private m_bChunkedProg()    As Byte             ' // Saved state (required by VST)

Private Sub Class_Initialize()
    Dim lPresetIndex    As Long
    Dim lSlotIndex      As Long
    Dim sAnsi           As String
    
    m_fSampleRate = 44100

    Set m_frmMain = New frmMain
    Set m_ctlUI = m_frmMain.ctlUI.object
    
    ReDim m_tSharedData.tPresets(NUM_OF_PROGRAMS - 1)
    m_tSharedData.lNumOfPresets = NUM_OF_PROGRAMS
    
    For lPresetIndex = 0 To NUM_OF_PROGRAMS - 1
    
        m_tSharedData.tPresets(lPresetIndex).sName = "Preset " & CStr(lPresetIndex)
            
        For lSlotIndex = 0 To UBound(m_tSharedData.tPresets(lPresetIndex).tSlots)
            m_tSharedData.tPresets(lPresetIndex).tSlots(lSlotIndex).tPattern.lLengthPerBeats = 4
        Next
        
    Next
    
    ' // Share variable
    m_ctlUI.SetBankShared VarPtr(m_tSharedData)

End Sub

Private Sub Class_Terminate()
    
    ' // Restore
    SetParent m_ctlUI.hWnd, m_frmMain.hWnd
    
    Unload m_frmMain
    
    Set m_ctlUI = Nothing
    Set m_frmMain = Nothing

End Sub

' // Save pointer to AEffect. Use it for GetTimeInfo call
Private Property Let IVBVstEffect_AEffectPtr( _
                     ByVal RHS As PTR)
    m_pAEffect = RHS
End Property
Private Property Get IVBVstEffect_AEffectPtr() As PTR
    IVBVstEffect_AEffectPtr = m_pAEffect
End Property
 
' // Save pointer to AudioMasterCallback. Use it for GetTimeInfo call
Private Property Get IVBVstEffect_AudioMasterCallback() As PTR
    IVBVstEffect_AudioMasterCallback = m_pfnHostCallback
End Property
Private Property Let IVBVstEffect_AudioMasterCallback( _
                     ByVal RHS As PTR)
    m_pfnHostCallback = RHS
End Property

' // You can use it for preallocating data
Private Property Let IVBVstEffect_BlockSize( _
                     ByVal RHS As Long)

End Property

' // Set the sample rate
Private Property Let IVBVstEffect_sampleRate( _
                     ByVal RHS As Single)
    Dim lMaxBufferLen   As Long
    
    m_fSampleRate = RHS
    
    ' // Update buffers (max 8 seconds)
    lMaxBufferLen = -Int(-m_fSampleRate * 8)
    
    m_tBuffer.lSize = lMaxBufferLen
    m_tBuffer.dReadCursor = 0
    m_tBuffer.lWriteCursor = 0
    
    ReDim m_tBuffer.fSamplesLeft(lMaxBufferLen - 1)
    ReDim m_tBuffer.fSamplesRight(lMaxBufferLen - 1)
    
End Property
Private Property Get IVBVstEffect_sampleRate() As Single
    IVBVstEffect_sampleRate = m_fSampleRate
End Property

Private Function IVBVstEffect_StartProcess() As Boolean
    IVBVstEffect_StartProcess = True
End Function

Private Function IVBVstEffect_StopProcess() As Boolean
    IVBVstEffect_StopProcess = True
End Function

' // Suspend plugin
Private Sub IVBVstEffect_Suspend()
    
    m_bEnabled = False
    ' // Key up
    m_tState.bPlaying = False
    
End Sub

' // Restore plugin
Private Sub IVBVstEffect_Resume()
    m_bEnabled = True
End Sub

' // Requests abilities
Private Property Get IVBVstEffect_CanDo( _
                     ByRef ppszRequest As String) As Boolean
    
    Select Case ppszRequest
    Case canDoReceiveVstTimeInfo, canDoReceiveVstEvents, canDoReceiveVstMidiEvent
        IVBVstEffect_CanDo = True
    End Select
    
End Property

' // Set bypass
Private Function IVBVstEffect_SetBypass( _
                 ByVal bValue As Boolean) As Boolean

End Function

' // Return true if your plugin uses ProcessEvents calls
Private Property Get IVBVstEffect_SupportsVSTEvents() As Boolean
    IVBVstEffect_SupportsVSTEvents = True
End Property

Private Property Get IVBVstEffect_CanMono() As Boolean
    IVBVstEffect_CanMono = False
End Property

' // Close editor (GUI window)
Private Sub IVBVstEffect_EditorClose()
    SetParent m_ctlUI.hWnd, m_frmMain.hWnd
    m_bEditorVisible = False
End Sub

' // Idle processing. You can return recorded events here
Private Function IVBVstEffect_EditorIdle( _
                 ByRef tRecords() As tAutomationRecord) As Long
    Dim lNumOfRecords   As Long
    
    ' // Update GUI
    m_ctlUI.StateChanged
    m_tSharedData.eChStateEffect = 0  ' // All changes applied
    
    ' // Check changed parameters
    With m_tSharedData
               
        If .eChStateUI = 0 Then
            Exit Function
        End If
        
        lNumOfRecords = 0
        ReDim tRecords(NUM_OF_PARAMETERS - 1)

        ' // Specify changed events
        If .eChStateUI And SCM_PITCH Then
        
            tRecords(lNumOfRecords).lParamIndex = PT_PITCH
            tRecords(lNumOfRecords).fParamValue = (.tPresets(.lCurPreset).tSlots(.lActiveSlot).fPitch + 12) / 24
            lNumOfRecords = lNumOfRecords + 1
            
        End If
        
        If .eChStateUI And SCM_SPEED Then
        
            tRecords(lNumOfRecords).lParamIndex = PT_SPEED
            tRecords(lNumOfRecords).fParamValue = .tPresets(.lCurPreset).tSlots(.lActiveSlot).fSpeed
            lNumOfRecords = lNumOfRecords + 1

        End If
        
        If .eChStateUI And SCM_SMOOTH Then
        
            tRecords(lNumOfRecords).lParamIndex = PT_SMOOTH
            tRecords(lNumOfRecords).fParamValue = .tPresets(.lCurPreset).tSlots(.lActiveSlot).fSmooth
            lNumOfRecords = lNumOfRecords + 1

        End If
        
        If .eChStateUI And SCM_SLOT_ACTIVE Then
            
            tRecords(lNumOfRecords).lParamIndex = PT_SLOT
            tRecords(lNumOfRecords).fParamValue = .lActiveSlot / 39
            lNumOfRecords = lNumOfRecords + 1

        End If
        
        ' // Applied
        .eChStateUI = 0
        
        IVBVstEffect_EditorIdle = lNumOfRecords
        
    End With
    
End Function

' // Editor window has beed opened
Private Function IVBVstEffect_EditorOpen( _
                 ByVal hWnd As Handle) As Boolean
                 
    SetParent m_ctlUI.hWnd, hWnd
    m_bEditorVisible = True
    IVBVstEffect_EditorOpen = True
    
End Function

' // Get GUI size. You should return pointer to ERect structure
Private Property Get IVBVstEffect_EditorRect() As ERect
        
    IVBVstEffect_EditorRect.wRight = m_frmMain.ctlUI.Width
    IVBVstEffect_EditorRect.wBottom = m_frmMain.ctlUI.Height

End Property

' // Get plugin name
Private Property Get IVBVstEffect_EffectName() As String
    IVBVstEffect_EffectName = VST_PLUGIN_NAME
End Property

Private Property Get IVBVstEffect_ProductString() As String
    IVBVstEffect_ProductString = VST_PRODUCT_STRING
End Property

' // You should register it
Private Property Get IVBVstEffect_UniqueId() As Long
    IVBVstEffect_UniqueId = UNIQUE_ID
End Property

Private Sub IVBVstEffect_VendorSpecific( _
            ByVal lArg1 As Long, _
            ByVal lArg2 As Long, _
            ByVal lpPtr As PTR, _
            ByVal fArg3 As Single)
End Sub

Private Property Get IVBVstEffect_VendorString() As String
    IVBVstEffect_VendorString = VST_VENDOR_STRING
End Property

Private Property Get IVBVstEffect_VendorVersion() As Long
    IVBVstEffect_VendorVersion = VST_VENDOR_VERSION
End Property

Private Property Get IVBVstEffect_Version() As Long
    IVBVstEffect_Version = VST_PLUGIN_VERSION
End Property

Private Property Get IVBVstEffect_VstVersion() As Long
    IVBVstEffect_VstVersion = 2400  ' // VST 2.4
End Property

' // Get plugin category (VstPlugCategory)
Private Property Get IVBVstEffect_PlugCategory() As Long
    IVBVstEffect_PlugCategory = kPlugCategEffect
End Property

' // Host wants to save state. You should save pointer to data to pData and size to lSize
Private Function IVBVstEffect_GetStateChunk( _
                 ByVal bIsPreset As Boolean, _
                 ByRef pData As PTR) As Long
    Dim lStartIndex As Long
    Dim lCount      As Long
    
    If bIsPreset Then
        lStartIndex = m_tSharedData.lCurPreset
        lCount = 1
    Else
        lStartIndex = 0
        lCount = m_tSharedData.lNumOfPresets
    End If
    
    If Not SerializePresets(m_tSharedData.tPresets(), lStartIndex, lCount, m_bChunkedProg) Then
        Exit Function
    End If
    
    pData = VarPtr(m_bChunkedProg(0))

    IVBVstEffect_GetStateChunk = UBound(m_bChunkedProg) + 1
    
End Function

' // Host want to restore state
Private Function IVBVstEffect_SetStateChunk( _
                 ByVal bIsPreset As Boolean, _
                 ByVal pData As PTR, _
                 ByVal lSize As Long) As Boolean
    Dim tPresets()  As tPreset
    Dim lIndex      As Long
    
    If Not DeserializePresets(pData, lSize, tPresets) Then
        Exit Function
    End If
    
    If bIsPreset Then
        m_tSharedData.tPresets(m_tSharedData.lCurPreset) = tPresets(0)
    Else
        For lIndex = 0 To UBound(tPresets)
            m_tSharedData.tPresets(lIndex) = tPresets(lIndex)
        Next
    End If
    
    m_tSharedData.eChStateEffect = m_tSharedData.eChStateEffect Or SCM_PROGRAM
    
    IVBVstEffect_SetStateChunk = True
    
End Function

' // If plugin has GUI return True
Private Property Get IVBVstEffect_HasEditor() As Boolean
    IVBVstEffect_HasEditor = True
End Property

' // Number of inputs/outputs. 2 for stereo.
Private Property Get IVBVstEffect_NumOfInputs() As Long
    IVBVstEffect_NumOfInputs = 2
End Property
Private Property Get IVBVstEffect_NumOfOutputs() As Long
    IVBVstEffect_NumOfOutputs = 2
End Property

' // Number of parameters
Private Property Get IVBVstEffect_NumOfParam() As Long
    IVBVstEffect_NumOfParam = NUM_OF_PARAMETERS
End Property

' // Number of programs
Private Property Get IVBVstEffect_NumOfPrograms() As Long
    IVBVstEffect_NumOfPrograms = NUM_OF_PROGRAMS
End Property

' // The following 3 properties is used for default GUI or automation displaying values
' // Returns display name of parameter
' // Text with a string representation ("0.5", "-3", "PLATE", etc...) of the value of parameter index. Limited to kVstMaxParamStrLen.
Private Property Get IVBVstEffect_ParamDisplay( _
                     ByVal lIndex As Long) As String
    With m_tSharedData
        Select Case lIndex
        Case PT_SPEED:  IVBVstEffect_ParamDisplay = Format$(.tPresets(.lCurPreset).tSlots(.lCurrentSlot).fSpeed * 100, "0.00")
        Case PT_PITCH:  IVBVstEffect_ParamDisplay = Format$(.tPresets(.lCurPreset).tSlots(.lCurrentSlot).fPitch, "0.00")
        Case PT_SLOT:   IVBVstEffect_ParamDisplay = .lCurrentSlot
        Case PT_SMOOTH: IVBVstEffect_ParamDisplay = Format$(.tPresets(.lCurPreset).tSlots(.lCurrentSlot).fSmooth * 100, "0.00")
        End Select
    End With
End Property

' // Returns parameter label
' // Label with the units in which parameter index is displayed (i.e. "sec", "dB", "type", etc...). Limited to kVstMaxParamStrLen.
Private Property Get IVBVstEffect_ParamLabel( _
                     ByVal lIndex As Long) As String
                     
    Select Case lIndex
    Case PT_SPEED:  IVBVstEffect_ParamLabel = "%"
    Case PT_PITCH:  IVBVstEffect_ParamLabel = "sem."
    Case PT_SLOT:   IVBVstEffect_ParamLabel = "id"
    Case PT_SMOOTH: IVBVstEffect_ParamLabel = "%"
    End Select
    
End Property

' // Return parameter name
' // Text with the name ("Time", "Gain", "RoomType", etc...) of parameter index. Limited to kVstMaxParamStrLen.
Private Property Get IVBVstEffect_ParamName( _
                     ByVal lIndex As Long) As String
                     
    Select Case lIndex
    Case PT_SPEED:  IVBVstEffect_ParamName = "Speed"
    Case PT_PITCH:  IVBVstEffect_ParamName = "Pitch"
    Case PT_SLOT:   IVBVstEffect_ParamName = "Slot"
    Case PT_SMOOTH: IVBVstEffect_ParamName = "Smooth"
    End Select
    
End Property

' // Parameters properties. You can return pointer to VstParameterProperties array
Private Property Get IVBVstEffect_ParameterProperties( _
                     ByRef pProperties As PTR) As Boolean
End Property

' // All the parameters can be automated
Private Property Get IVBVstEffect_CanParameterBeAutomated( _
                     ByVal lIndex As Long) As Boolean
    IVBVstEffect_CanParameterBeAutomated = True
End Property


' // Host sets value of parameter. The value should be 0..1 float (Single) range
' // When you don't use chunked programs host obtains / sets all the parameters during saving / loading plugin
' // Also it used during automation
Private Property Let IVBVstEffect_ParamValue( _
                     ByVal lIndex As Long, _
                     ByVal RHS As Single)

    With m_tSharedData

        Select Case lIndex
        Case PT_SPEED
            
            ' // Check if there is writing events
            If .bRecordMode And (.eChStateUI And SCM_SPEED) Then
                Exit Property
            End If
            
            If RHS < 0 Then
                RHS = 0
            ElseIf RHS > 1 Then
                RHS = 1
            End If
        
            .tPresets(.lCurPreset).tSlots(.lCurrentSlot).fSpeed = RHS
            .eChStateEffect = .eChStateEffect Or SCM_SPEED
            
        Case PT_PITCH
        
            ' // Check if there is writing events
            If .bRecordMode And (.eChStateUI And SCM_PITCH) Then
                Exit Property
            End If
            
            RHS = RHS * 24 - 12
            
            If RHS > 12 Then
                RHS = 12
            ElseIf RHS < -12 Then
                RHS = -12
            End If
            
            .tPresets(.lCurPreset).tSlots(.lCurrentSlot).fPitch = RHS
            .eChStateEffect = .eChStateEffect Or SCM_PITCH
            
        Case PT_SLOT
            
            ' // Check if there is writing events
            If .bRecordMode And (.eChStateUI And SCM_SLOT_ACTIVE) Then
                Exit Property
            End If
            
            RHS = RHS * 39
            
            If RHS < 0 Then
                RHS = 0
            ElseIf RHS > 39 Then
                RHS = 39
            End If
            
            .lCurrentSlot = RHS
            .eChStateEffect = .eChStateEffect Or SCM_SLOT_CURRENT
            
        Case PT_SMOOTH
            
            ' // Check if there is writing events
            If .bRecordMode And (.eChStateUI And SCM_SMOOTH) Then
                Exit Property
            End If
            
            If RHS < 0 Then
                RHS = 0
            ElseIf RHS > 1 Then
                RHS = 1
            End If
        
            .tPresets(.lCurPreset).tSlots(.lCurrentSlot).fSmooth = RHS
            .eChStateEffect = .eChStateEffect Or SCM_SMOOTH
            
        End Select

    End With
    
End Property

' // Get parameter value
Private Property Get IVBVstEffect_ParamValue( _
                     ByVal lIndex As Long) As Single
    With m_tSharedData
        Select Case lIndex
        Case PT_SPEED:  IVBVstEffect_ParamValue = .tPresets(.lCurPreset).tSlots(.lCurrentSlot).fSpeed
        Case PT_PITCH:  IVBVstEffect_ParamValue = (.tPresets(.lCurPreset).tSlots(.lCurrentSlot).fPitch + 12) / 24
        Case PT_SLOT:   IVBVstEffect_ParamValue = .lCurrentSlot / 39
        Case PT_SMOOTH: IVBVstEffect_ParamValue = .tPresets(.lCurPreset).tSlots(.lCurrentSlot).fSmooth
        End Select
    End With
End Property

' // Process samples. This method should accumulate result ie each output sample = output sample + fn(input sample)
Private Sub IVBVstEffect_Process( _
            ByVal pInput As PTR, _
            ByVal pOutput As PTR, _
            ByVal lSampleFrames As Long)
    Process pInput, pOutput, lSampleFrames, False
End Sub

' // Process samples. This method should replace result ie each output sample = fn(input sample)
Private Sub IVBVstEffect_ProcessReplacing( _
            ByVal pInput As PTR, _
            ByVal pOutput As PTR, _
            ByVal lSampleFrames As Long)
    Process pInput, pOutput, lSampleFrames, True
End Sub

' // Process events. Use it to track PitchBend MIDI messages
Private Function IVBVstEffect_ProcessEvents( _
                 ByRef tEvents As VstEvents) As Boolean
    Dim lType       As Long
    Dim tMidiEvent  As VstMidiEvent
    Dim lPitchVal   As Long
    Dim pEvent      As PTR
    Dim lIndex      As Long
    
    For lIndex = 0 To tEvents.numEvents - 1
        
        pEvent = tEvents.pEvents + lIndex * 4
        
        GetMem4 ByVal pEvent, lType

        ' // Process only MIDI events
        If lType = kVstMidiType Then
            
            memcpy tMidiEvent, ByVal pEvent, Len(tMidiEvent)

            ' // Check only Pitch events (ignore channel)
            If tMidiEvent.midiData(0) >= &HE0 And tMidiEvent.midiData(0) <= &HEF Then
                
                lPitchVal = (tMidiEvent.midiData(1) And &H7F) Or _
                            (tMidiEvent.midiData(2) And &H7F) * &H80&
                
                ' // -24..+24 sem
                m_tSharedData.fPitchBend = (lPitchVal - &H2000) / 341
                
            End If
            
        End If
        
    Next
    
    IVBVstEffect_ProcessEvents = True
    
End Function

' // Set current program
Private Property Let IVBVstEffect_Program( _
                     ByVal RHS As Long)
    If RHS < NUM_OF_PROGRAMS And m_tSharedData.lCurPreset <> RHS Then
        m_tSharedData.lCurPreset = RHS
        m_tSharedData.eChStateEffect = m_tSharedData.eChStateEffect Or SCM_PROGRAM
    End If
End Property

' // Get current program
Private Property Get IVBVstEffect_Program() As Long
    IVBVstEffect_Program = m_tSharedData.lCurPreset
End Property

' // Set current program name
Private Property Let IVBVstEffect_ProgramName( _
                     ByVal RHS As String)
    m_tSharedData.tPresets(m_tSharedData.lCurPreset).sName = RHS
End Property

' // Get current program name
Private Property Get IVBVstEffect_ProgramName() As String
    IVBVstEffect_ProgramName = m_tSharedData.tPresets(m_tSharedData.lCurPreset).sName
End Property

' // Get specified program name
Private Property Get IVBVstEffect_ProgramNameIndexed( _
                     ByVal lCategory As Long, _
                     ByVal lIndex As Long) As String
    If lIndex >= 0 And lIndex < m_tSharedData.lNumOfPresets Then
        IVBVstEffect_ProgramNameIndexed = m_tSharedData.tPresets(lIndex).sName
    End If
End Property

' // Copy program (seems deprecated)
Private Sub IVBVstEffect_CopyProgram( _
            ByVal lDestination As Long)
    If lDestination > 0 And lDestination < m_tSharedData.lNumOfPresets Then
        m_tSharedData.tPresets(lDestination) = m_tSharedData.tPresets(m_tSharedData.lCurPreset)
    End If
End Sub

' // If you return true host will call GetStateChunk / SetStateChunk to save / restore state.
Private Property Get IVBVstEffect_ProgramsAreChunks() As Boolean
    IVBVstEffect_ProgramsAreChunks = True
End Property

' // Return tail size in samples. For example for FIR filter it return the kernel size.
' // This plugin tail size depends on pattern so we don't use it
Private Property Get IVBVstEffect_TailSize() As Long
End Property

Private Property Get IVBVstEffect_ThreadId() As Long
    IVBVstEffect_ThreadId = App.ThreadID
End Property

' // Get next event in pattern
Private Sub GetNextEvent( _
            ByRef tPattern As tPattern, _
            ByRef tEvent As tEvent)
    Dim lIndex  As Long
    
    With tEvent
    
        Select Case .eType
        Case ET_START_NOTE
                    
            ' // Return note releasing
            .dPos = tPattern.tKeys(.lIndex).dPos + tPattern.tKeys(.lIndex).dLength
            
            If .dPos > tPattern.lLengthPerBeats * 4 Then
                .dPos = tPattern.lLengthPerBeats * 4
            End If
            
            .eType = ET_END_NOTE
        
        Case ET_END_NOTE
        
            If .lIndex < tPattern.lNumOfKeys - 1 Then
            
                ' // Return next note
                .eType = ET_START_NOTE
                .lIndex = .lIndex + 1
                .dPos = tPattern.tKeys(.lIndex).dPos
                
                If .dPos >= tPattern.lLengthPerBeats * 4 Then
                
                    .eType = ET_END_PATTERN
                    .dPos = tPattern.lLengthPerBeats * 4
                    lIndex = -1
                    
                End If
                
            Else
            
                ' // End of pattern
                .eType = ET_END_PATTERN
                .dPos = tPattern.lLengthPerBeats * 4
                .lIndex = -1
                
            End If
            
        Case ET_END_PATTERN

            .eType = ET_START_PATTERN
            .dPos = 0
            .lIndex = -1

        Case ET_START_PATTERN
        
            If tPattern.lNumOfKeys Then
                
                ' // Search for first suited note
                For lIndex = 0 To tPattern.lNumOfKeys - 1
                    
                    If .dPos > tPattern.tKeys(lIndex).dPos And .dPos <= tPattern.tKeys(lIndex).dPos + _
                        tPattern.tKeys(lIndex).dLength Then
                    
                        .eType = ET_END_NOTE
                        .dPos = tPattern.tKeys(lIndex).dPos + tPattern.tKeys(lIndex).dLength
                        
                        If .dPos > tPattern.lLengthPerBeats * 4 Then
                            .dPos = tPattern.lLengthPerBeats * 4
                        End If
            
                        .lIndex = lIndex
                        
                        Exit Sub
                        
                    ElseIf .dPos <= tPattern.tKeys(lIndex).dPos Then
                    
                        .eType = ET_START_NOTE
                        .dPos = tPattern.tKeys(lIndex).dPos
                        .lIndex = lIndex
                        
                        Exit Sub

                    End If
                    
                Next
                
                .eType = ET_END_PATTERN
                .dPos = tPattern.lLengthPerBeats * 4
                .lIndex = -1
                
            Else
            
                ' // End of pattern
                .eType = ET_END_PATTERN
                .dPos = tPattern.lLengthPerBeats * 4
                .lIndex = -1
                
            End If
        
        End Select
        
    End With
    
End Sub

' // Calculate precalculations
Private Sub RecalcStatePrecalculations()
    Dim lIndex      As Long
    Dim bChanged    As Boolean
    Dim lSamples    As Long
    
    With m_tState
    
        .dFrequency = 523.251 * (2 ^ ((.lNote - 60 + .dPitch + m_tSharedData.fPitchBend) / 12))
        .dSamplesPerWave = m_fSampleRate / .dFrequency
        .dSampleOffset = .dSamplesPerWave * .dSpeed
        .lSmoothBufOfst = Int(.dSamplesPerWave - .dSampleOffset)
        
        lSamples = .dSamplesPerWave * .dSmooth / 2
        
        If lSamples < 2 Then
            ' // No smooth on high frequencies
            lSamples = 0
        End If
        
        If lSamples <> .lSmoothSamples Then
            
            .lSmoothSamples = lSamples
            
            If lSamples > .lCurveBufSize Then
                ReDim .dSmoothCurve(lSamples - 1)
                .lCurveBufSize = lSamples
            End If
            
            For lIndex = 0 To lSamples - 1
                .dSmoothCurve(lIndex) = Sin(1.5707963267949 * lIndex / lSamples)
            Next
            
        End If
        
   End With
   
End Sub

Private Sub Process( _
            ByVal pInput As PTR, _
            ByVal pOutput As PTR, _
            ByVal lSampleFrames As Long, _
            ByVal bReplacing As Boolean)
    Static s_tSADescs(3)    As SAFEARRAY1D
    
    Dim lSampleIndex    As Long:    Dim tTimeInfo       As VstTimeInfo
    Dim tNextEvent      As tEvent:  Dim lNumOfSamples   As Long
    Dim dCurPos         As Double:  Dim lCountSamples   As Long
    Dim lNextEvtSample  As Long:    Dim fInLeft()       As Single
    Dim fInRight()      As Single:  Dim fOutLeft()      As Single
    Dim fOutRight()     As Single
    
    If Not m_bEnabled Then
        Exit Sub
    End If
    
    On Error GoTo exit_proc
    
    GetChannelData pInput, 0, lSampleFrames, fInLeft, s_tSADescs(0)
    GetChannelData pInput, 1, lSampleFrames, fInRight, s_tSADescs(1)
    GetChannelData pOutput, 0, lSampleFrames, fOutLeft, s_tSADescs(2)
    GetChannelData pOutput, 1, lSampleFrames, fOutRight, s_tSADescs(3)
    
    ' // Get current timeslice info
    tTimeInfo = GetTimeInfo(m_pfnHostCallback, m_pAEffect, kVstPpqPosValid Or kVstTempoValid)
    
    If tTimeInfo.Flags And kVstTempoValid Then
        If tTimeInfo.tempo <= 0 Then
            tTimeInfo.tempo = 120
        End If
    Else
        tTimeInfo.tempo = 120
    End If
    
    If tTimeInfo.Flags And kVstPpqPosValid Then
        m_tSharedData.dPlaybackPos = tTimeInfo.ppqPos
    End If
    
    ' // Something changed
    If tTimeInfo.Flags And kVstTransportChanged Then
    
        If (tTimeInfo.Flags And kVstTransportPlaying) = 0 Then
            ' // Key up
            m_tState.bPlaying = False
        End If
        
    End If

    If tTimeInfo.Flags And kVstAutomationWriting Then
        m_tSharedData.bRecordMode = True
    Else
        m_tSharedData.bRecordMode = False
    End If
    
    With m_tSharedData.tPresets(m_tSharedData.lCurPreset).tSlots(m_tSharedData.lPlaybackSlot)
        
        ' // Calculate offset from begin pattern
        dCurPos = (tTimeInfo.ppqPos - Int(tTimeInfo.ppqPos / .tPattern.lLengthPerBeats) * .tPattern.lLengthPerBeats) * 4

        tNextEvent.dPos = dCurPos
        tNextEvent.eType = ET_START_PATTERN
        
        If tNextEvent.dPos = 0 Then
        
            ' // Update slot
            If m_tSharedData.lCurrentSlot <> m_tSharedData.lPlaybackSlot Then
                m_tSharedData.eChStateEffect = m_tSharedData.eChStateEffect Or SCM_SLOT_PLAYBACK
                m_tSharedData.lPlaybackSlot = m_tSharedData.lCurrentSlot
            End If
        
            ' // Key up
            m_tState.bPlaying = False
 
        End If
        
    End With

    Do
        
        ' // Because of playback slot may be changed use different With statement
        With m_tSharedData.tPresets(m_tSharedData.lCurPreset).tSlots(m_tSharedData.lPlaybackSlot)
            
            ' // Get next event in pattern
            GetNextEvent .tPattern, tNextEvent
    
            ' // Get number of samples to next event pos
            lNextEvtSample = 15 / tTimeInfo.tempo * (tNextEvent.dPos - dCurPos) * m_fSampleRate
            
            ' // Process samples to next event
            If lSampleIndex + lNextEvtSample > lSampleFrames Then
                lNumOfSamples = lSampleFrames - lSampleIndex
            Else
                lNumOfSamples = lNextEvtSample
            End If
    
            If m_tState.bPlaying And lNumOfSamples > 0 Then
                
                ' // Save samples to buffer. We should save all the samples because note can play for long time with 100% speed
                lCountSamples = lNumOfSamples
                
                If m_tBuffer.lWriteCursor + lCountSamples > m_tBuffer.lSize Then
                    lCountSamples = m_tBuffer.lSize - m_tBuffer.lWriteCursor
                End If
                
                If lCountSamples > 0 Then
                
                    memcpy m_tBuffer.fSamplesLeft(m_tBuffer.lWriteCursor), fInLeft(lSampleIndex), lCountSamples * 4
                    memcpy m_tBuffer.fSamplesRight(m_tBuffer.lWriteCursor), fInRight(lSampleIndex), lCountSamples * 4
                    m_tBuffer.lWriteCursor = m_tBuffer.lWriteCursor + lCountSamples
                    
                End If

                m_tState.dPitch = .fPitch
                m_tState.dSpeed = .fSpeed
                m_tState.dSmooth = .fSmooth
                
                RecalcStatePrecalculations
                
            End If
            
            If lNumOfSamples > 0 Then
                ProcessSamples fInLeft, fInRight, fOutLeft, fOutRight, lSampleIndex, lNumOfSamples, bReplacing
            End If
            
            ' // If next event behind current time slice - exit
            If lSampleIndex + lNextEvtSample > lSampleFrames Then
                Exit Do
            End If
            
            ' // Increase number of processed samples
            lSampleIndex = lSampleIndex + lNumOfSamples
            
            dCurPos = tNextEvent.dPos
            
            ' // Analyse next event
            Select Case tNextEvent.eType
            Case ET_START_NOTE

                ' // Key down
                m_tState.bPlaying = True
                m_tState.bIsFirst = True    ' // To smooth
                m_tState.dPhase = 0
                m_tState.lNote = .tPattern.tKeys(tNextEvent.lIndex).lValue
   
                m_tBuffer.lWriteCursor = 0
                m_tBuffer.dReadCursor = 0
                
            Case ET_END_NOTE
                
                ' // Key up
                m_tState.bPlaying = False
                
            Case ET_END_PATTERN
    
                dCurPos = 0

                ' // Update slot
                If m_tSharedData.lCurrentSlot <> m_tSharedData.lPlaybackSlot Then
                    m_tSharedData.eChStateEffect = m_tSharedData.eChStateEffect Or SCM_SLOT_PLAYBACK
                    m_tSharedData.lPlaybackSlot = m_tSharedData.lCurrentSlot
                End If
                
                ' // Key up
                m_tState.bPlaying = False
                    
            End Select
            
        End With
        
    Loop While True
                
exit_proc:
    
    PutMemPtr ByVal ArrPtr(fInLeft), NULL_PTR
    PutMemPtr ByVal ArrPtr(fInRight), NULL_PTR
    PutMemPtr ByVal ArrPtr(fOutLeft), NULL_PTR
    PutMemPtr ByVal ArrPtr(fOutRight), NULL_PTR
    
End Sub

' // Process specified samples (accumulating)
Private Sub ProcessSamples( _
            ByRef fInLeft() As Single, _
            ByRef fInRight() As Single, _
            ByRef fOutLeft() As Single, _
            ByRef fOutRight() As Single, _
            ByVal lStartSample As Long, _
            ByVal lSampleFrames As Long, _
            ByVal bReplacing As Boolean)
    Dim lPhase          As Long
    Dim lBufIndex       As Long
    Dim lIndex          As Long
    Dim fSampleValueL   As Single
    Dim fSampleValueR   As Single
    Dim dSmoothValue    As Double
    
    If lSampleFrames <= 0 Then
        Exit Sub
    End If
    
    If m_tState.bPlaying Then

        For lIndex = lStartSample To lStartSample + lSampleFrames - 1
            
            lBufIndex = m_tBuffer.dReadCursor + m_tState.dPhase
            lPhase = m_tState.dPhase
            
            If lBufIndex >= m_tBuffer.lSize Then
                lBufIndex = 0
                m_tBuffer.dReadCursor = 0
            End If
            
            If lPhase < m_tState.lSmoothSamples Then
            
                ' // Left smooth
                dSmoothValue = m_tState.dSmoothCurve(lPhase)
                fSampleValueL = m_tBuffer.fSamplesLeft(lBufIndex) * dSmoothValue
                fSampleValueR = m_tBuffer.fSamplesRight(lBufIndex) * dSmoothValue
                
                 ' // Add right smooth
                If Not m_tState.bIsFirst Then
                
                    ' // Because of bIsFirst flag we don't access behind buffer
                    dSmoothValue = m_tState.dSmoothCurve(m_tState.lSmoothSamples - lPhase - 1)
                    fSampleValueL = fSampleValueL + m_tBuffer.fSamplesLeft(lBufIndex + m_tState.lSmoothBufOfst) * dSmoothValue
                    fSampleValueR = fSampleValueR + m_tBuffer.fSamplesRight(lBufIndex + m_tState.lSmoothBufOfst) * dSmoothValue
                  
                End If
                
            Else
                fSampleValueL = m_tBuffer.fSamplesLeft(lBufIndex)
                fSampleValueR = m_tBuffer.fSamplesRight(lBufIndex)
            End If
            
            If bReplacing Then
                fOutLeft(lIndex) = fSampleValueL
                fOutRight(lIndex) = fSampleValueR
            Else
                fOutLeft(lIndex) = fOutLeft(lIndex) + fSampleValueL
                fOutRight(lIndex) = fOutRight(lIndex) + fSampleValueR
            End If
            
            m_tState.dPhase = m_tState.dPhase + 1
            
            If m_tState.dPhase >= m_tState.dSamplesPerWave Then
            
                m_tState.dPhase = m_tState.dPhase - m_tState.dSamplesPerWave
                m_tBuffer.dReadCursor = m_tBuffer.dReadCursor + m_tState.dSampleOffset
                m_tState.bIsFirst = False
                
            End If
            
        Next
    
    Else
        If bReplacing Then
            memcpy fOutLeft(lStartSample), fInLeft(lStartSample), lSampleFrames * 4
            memcpy fOutRight(lStartSample), fInRight(lStartSample), lSampleFrames * 4
        Else
            For lIndex = lStartSample To lStartSample + lSampleFrames - 1
                fOutLeft(lIndex) = fOutLeft(lIndex) + fInLeft(lIndex)
                fOutRight(lIndex) = fOutRight(lIndex) + fInRight(lIndex)
            Next
        End If
    End If
       
End Sub

