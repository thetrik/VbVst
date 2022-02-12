VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVbTrickCrusher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' //
' // VST-class template
' //

Option Explicit

Private Const NUM_OF_PARAMETERS As Long = 2
Private Const NUM_OF_PROGRAMS   As Long = 0
Private Const NUM_OF_INPUTS     As Long = 2
Private Const NUM_OF_OUTPUTS    As Long = 2
Private Const PLUGIN_UNIQUE_ID  As Long = &H32120843    ' // Update it always!
Private Const VST_EFFECT_NAME   As String = "VbTrickCrusher"

Private m_pAEffect          As PTR
Private m_pfnHostCallBack   As PTR
Private m_fSampleRate       As Single
Private m_fFrequency        As Single
Private m_fQuantize         As Single
Private m_fPhase            As Single
Private m_fCurSampleL       As Single
Private m_fCurSampleR       As Single

Implements IVBVstEffect

Private Sub Class_Initialize()
    m_fFrequency = 1
    m_fQuantize = 1
End Sub

' // Process samples here
Private Sub IVBVstEffect_Process( _
            ByVal pInput As PTR, _
            ByVal pOutput As PTR, _
            ByVal sampleFrames As Long)
    Static s_tDescs(3)   As SAFEARRAY1D
    
    Dim fInLeft()   As Single:      Dim fInRight()  As Single
    Dim fOutLeft()  As Single:      Dim fOutRight() As Single
    Dim lIndex      As Long:        Dim fSmpWidth   As Single
    Dim fQuant      As Single
    
    GetChannelData pInput, 0, sampleFrames, fInLeft, s_tDescs(0)
    GetChannelData pInput, 1, sampleFrames, fInRight, s_tDescs(1)
    GetChannelData pOutput, 0, sampleFrames, fOutLeft, s_tDescs(2)
    GetChannelData pOutput, 1, sampleFrames, fOutRight, s_tDescs(3)
    
    fQuant = 2 ^ (m_fQuantize * 16)
    
    If m_fFrequency > 0 Then
        fSmpWidth = 1 / (m_fFrequency ^ 4)
    Else
        fSmpWidth = sampleFrames + 1
        m_fPhase = 1
    End If
    
    If m_fPhase = 0 Then
        m_fCurSampleL = CLng(fInLeft(0) * fQuant) / fQuant
        m_fCurSampleR = CLng(fInRight(0) * fQuant) / fQuant
    End If

    For lIndex = 0 To sampleFrames - 1
            
        If m_fPhase >= fSmpWidth Then
        
            m_fPhase = m_fPhase - fSmpWidth
            m_fCurSampleL = CLng(fInLeft(lIndex) * fQuant) / fQuant
            m_fCurSampleR = CLng(fInRight(lIndex) * fQuant) / fQuant
        
        End If
        
        fOutLeft(lIndex) = fOutLeft(lIndex) + m_fCurSampleL
        fOutRight(lIndex) = fOutRight(lIndex) + m_fCurSampleR
        
        m_fPhase = m_fPhase + 1
        
    Next
    
    PutMemPtr ByVal ArrPtr(fInLeft), NULL_PTR
    PutMemPtr ByVal ArrPtr(fInRight), NULL_PTR
    PutMemPtr ByVal ArrPtr(fOutLeft), NULL_PTR
    PutMemPtr ByVal ArrPtr(fOutRight), NULL_PTR
    
End Sub

' // Process samples here
Private Sub IVBVstEffect_ProcessReplacing( _
            ByVal pInput As PTR, _
            ByVal pOutput As PTR, _
            ByVal sampleFrames As Long)
    Static s_tDescs(3)   As SAFEARRAY1D
    
    Dim fInLeft()   As Single:      Dim fInRight()  As Single
    Dim fOutLeft()  As Single:      Dim fOutRight() As Single
    Dim lIndex      As Long:        Dim fSmpWidth   As Single
    Dim fQuant      As Single
    
    GetChannelData pInput, 0, sampleFrames, fInLeft, s_tDescs(0)
    GetChannelData pInput, 1, sampleFrames, fInRight, s_tDescs(1)
    GetChannelData pOutput, 0, sampleFrames, fOutLeft, s_tDescs(2)
    GetChannelData pOutput, 1, sampleFrames, fOutRight, s_tDescs(3)
    
    fQuant = 2 ^ (m_fQuantize * 16)
    
    If m_fFrequency > 0 Then
        fSmpWidth = 1 / (m_fFrequency ^ 4)
    Else
        fSmpWidth = sampleFrames + 1
        m_fPhase = 1
    End If
    
    If m_fPhase = 0 Then
        m_fCurSampleL = CLng(fInLeft(0) * fQuant) / fQuant
        m_fCurSampleR = CLng(fInRight(0) * fQuant) / fQuant
    End If

    For lIndex = 0 To sampleFrames - 1
            
        If m_fPhase >= fSmpWidth Then
        
            m_fPhase = m_fPhase - fSmpWidth
            m_fCurSampleL = CLng(fInLeft(lIndex) * fQuant) / fQuant
            m_fCurSampleR = CLng(fInRight(lIndex) * fQuant) / fQuant
        
        End If
        
        fOutLeft(lIndex) = m_fCurSampleL
        fOutRight(lIndex) = m_fCurSampleR
        
        m_fPhase = m_fPhase + 1
        
    Next
    
    PutMemPtr ByVal ArrPtr(fInLeft), NULL_PTR
    PutMemPtr ByVal ArrPtr(fInRight), NULL_PTR
    PutMemPtr ByVal ArrPtr(fOutLeft), NULL_PTR
    PutMemPtr ByVal ArrPtr(fOutRight), NULL_PTR
    
End Sub

' // Process MIDI-events here
Private Function IVBVstEffect_ProcessEvents( _
                 pEvents As VBVST2X.VstEvents) As Boolean

End Function

' // Use following properties to work with parameters

Private Property Get IVBVstEffect_NumOfParam() As Long
    IVBVstEffect_NumOfParam = NUM_OF_PARAMETERS
End Property

Private Property Get IVBVstEffect_ParamLabel( _
                     ByVal lIndex As Long) As String
    Select Case lIndex
    Case 0: IVBVstEffect_ParamLabel = "Hz"
    Case 1: IVBVstEffect_ParamLabel = "bits"
    End Select
End Property

Private Property Get IVBVstEffect_ParamName( _
                     ByVal lIndex As Long) As String
    Select Case lIndex
    Case 0: IVBVstEffect_ParamName = "Freq"
    Case 1: IVBVstEffect_ParamName = "Quant"
    End Select
End Property

Private Property Let IVBVstEffect_ParamValue( _
                     ByVal lIndex As Long, _
                     ByVal RHS As Single)
    Select Case lIndex
    Case 0: m_fFrequency = RHS
    Case 1: m_fQuantize = RHS
    End Select
End Property

Private Property Get IVBVstEffect_ParamValue( _
                     ByVal lIndex As Long) As Single
    Select Case lIndex
    Case 0: IVBVstEffect_ParamValue = m_fFrequency
    Case 1: IVBVstEffect_ParamValue = m_fQuantize
    End Select
End Property

Private Property Get IVBVstEffect_ParamDisplay( _
                     ByVal lIndex As Long) As String
    Select Case lIndex
    Case 0: IVBVstEffect_ParamDisplay = Format$((m_fFrequency ^ 4) * m_fSampleRate / 2, "0")
    Case 1: IVBVstEffect_ParamDisplay = Format$(m_fQuantize * 24 + 1, "0")
    End Select
End Property

Private Property Get IVBVstEffect_ParameterProperties( _
                     ByRef pProperties As VBVST2X.PTR) As Boolean

End Property

Private Property Get IVBVstEffect_CanParameterBeAutomated( _
                     ByVal lIndex As Long) As Boolean

End Property

' // Plugin information

Private Property Get IVBVstEffect_UniqueId() As Long
    IVBVstEffect_UniqueId = PLUGIN_UNIQUE_ID
End Property

Private Property Get IVBVstEffect_PlugCategory() As Long
    IVBVstEffect_PlugCategory = kPlugCategEffect
End Property

Private Property Get IVBVstEffect_EffectName() As String
    IVBVstEffect_EffectName = VST_EFFECT_NAME
End Property

Private Property Get IVBVstEffect_ProductString() As String

End Property

Private Sub IVBVstEffect_VendorSpecific( _
            ByVal lArg1 As Long, _
            ByVal lArg2 As Long, _
            ByVal lpPtr As VBVST2X.PTR, _
            ByVal fArg3 As Single)

End Sub

Private Property Get IVBVstEffect_VendorString() As String

End Property

Private Property Get IVBVstEffect_VendorVersion() As Long

End Property

Private Property Get IVBVstEffect_Version() As Long

End Property

Private Property Get IVBVstEffect_VstVersion() As Long
    IVBVstEffect_VstVersion = 0
End Property

' // Persistence (for chunked programs)

Private Function IVBVstEffect_GetStateChunk( _
                 ByVal bIsPreset As Boolean, _
                 ByRef pData As VBVST2X.PTR) As Long

End Function

Private Function IVBVstEffect_SetStateChunk( _
                 ByVal bIsPreset As Boolean, _
                 ByVal pData As VBVST2X.PTR, _
                 ByVal lSize As Long) As Boolean

End Function

' // Abilities

Private Property Get IVBVstEffect_CanDo( _
                     ByRef ppszRequest As String) As Boolean

End Property

Private Property Get IVBVstEffect_CanMono() As Boolean

End Property

Private Property Get IVBVstEffect_HasEditor() As Boolean

End Property

Private Property Get IVBVstEffect_ProgramsAreChunks() As Boolean

End Property

Private Property Get IVBVstEffect_SupportsVSTEvents() As Boolean

End Property

' // Programs (presets)

Private Property Get IVBVstEffect_NumOfPrograms() As Long
    IVBVstEffect_NumOfPrograms = NUM_OF_PROGRAMS
End Property

Private Property Let IVBVstEffect_Program( _
                     ByVal RHS As Long)

End Property

Private Property Get IVBVstEffect_Program() As Long

End Property

Private Property Let IVBVstEffect_ProgramName( _
                     ByVal RHS As String)

End Property

Private Property Get IVBVstEffect_ProgramName() As String

End Property

Private Property Get IVBVstEffect_ProgramNameIndexed( _
                     ByVal lCategory As Long, _
                     ByVal lIndex As Long) As String

End Property

Private Sub IVBVstEffect_CopyProgram( _
            ByVal lDestination As Long)

End Sub

' // Editor

Private Function IVBVstEffect_EditorOpen( _
                 ByVal hWnd As VBVST2X.Handle) As Boolean

End Function

Private Sub IVBVstEffect_EditorClose()

End Sub

Private Function IVBVstEffect_EditorIdle( _
                 tRecords() As VBVST2X.tAutomationRecord) As Long

End Function

Private Property Get IVBVstEffect_EditorRect() As VBVST2X.ERect

End Property

' // Misc

Private Sub IVBVstEffect_Suspend()
    m_fPhase = 0
End Sub

Private Sub IVBVstEffect_Resume()
    
End Sub

Private Function IVBVstEffect_StartProcess() As Boolean
    m_fPhase = 0
    IVBVstEffect_StartProcess = True
End Function

Private Function IVBVstEffect_StopProcess() As Boolean
    m_fPhase = 0
    IVBVstEffect_StopProcess = True
End Function

Private Property Let IVBVstEffect_SampleRate( _
                     ByVal RHS As Single)
    m_fSampleRate = RHS
End Property

Private Property Let IVBVstEffect_BlockSize( _
                     ByVal RHS As Long)

End Property

Private Property Get IVBVstEffect_SampleRate() As Single
    IVBVstEffect_SampleRate = m_fSampleRate
End Property

Private Property Let IVBVstEffect_AEffectPtr( _
                     ByVal RHS As VBVST2X.PTR)
    m_pAEffect = RHS
End Property

Private Property Get IVBVstEffect_AEffectPtr() As VBVST2X.PTR
    IVBVstEffect_AEffectPtr = m_pAEffect
End Property

Private Property Let IVBVstEffect_AudioMasterCallback( _
                     ByVal RHS As VBVST2X.PTR)
    m_pfnHostCallBack = RHS
End Property

Private Property Get IVBVstEffect_AudioMasterCallback() As VBVST2X.PTR
    IVBVstEffect_AudioMasterCallback = m_pfnHostCallBack
End Property

Private Property Get IVBVstEffect_NumOfInputs() As Long
    IVBVstEffect_NumOfInputs = NUM_OF_INPUTS
End Property

Private Property Get IVBVstEffect_NumOfOutputs() As Long
    IVBVstEffect_NumOfOutputs = NUM_OF_OUTPUTS
End Property

Private Function IVBVstEffect_SetBypass( _
                 ByVal bValue As Boolean) As Boolean

End Function

Private Property Get IVBVstEffect_TailSize() As Long

End Property

Private Property Get IVBVstEffect_ThreadId() As Long
    IVBVstEffect_ThreadId = App.ThreadID
End Property
