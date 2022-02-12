struct ERect {
	short wTop;		///< top coordinate
	short wLeft;		///< left coordinate
	short wBottom;	///< bottom coordinate
	short wRight;		///< right coordinate
};

struct AEEffect {
    long magic;    // kEffectMagic
    PTR Dispatcher;
    PTR Process;
    PTR SetParameter;
    PTR GetParameter;
    long numPrograms;
    long numParams;
    long numInputs;
    long numOutputs;
    long flags;
    long resvd1;
    long resvd2;
    long initialDelay;
    long realQualities;
    long offQualities;
    float ioRatio;
    SAFEARRAY(CVBVstDispatcher) object;
    long user;
    long UniqueId;
    long Version;
    PTR ProcessReplacing;
    PTR ProcessDoubleReplacing;
    unsigned char future[56];
};

typedef enum VstAEffectFlags {
    effFlagsHasEditor = 1,
    effFlagsCanReplacing = 16,
    effFlagsProgramChunks = 32
} VstAEffectFlags;

typedef enum VstTimeInfoFlags {
    kVstTransportChanged = 1,            ///< indicates that play, cycle or record state has changed
    kVstTransportPlaying = 2,            ///< set if Host sequencer is currently playing
    kVstTransportCycleActive = 4,        ///< set if Host sequencer is in cycle mode
    kVstTransportRecording = 0x8,        ///< set if automation write mode active (record parameter changes)
	kVstAutomationWriting = 0x40,
    kVstAutomationReading = 0x80,        ///< set if automation read mode active (play parameter changes)
    kVstNanosValid = 0x100,              ///< VstTimeInfo::nanoSeconds valid
    kVstPpqPosValid = 0x200,             ///< VstTimeInfo::ppqPos valid
    kVstTempoValid = 0x400,              ///< VstTimeInfo::tempo valid
    kVstBarsValid = 0x800,               ///< VstTimeInfo::barStartPos valid
    kVstCyclePosValid = 0x1000,          ///< VstTimeInfo::cycleStartPos and VstTimeInfo::cycleEndPos valid
    kVstTimeSigValid = 0x2000,           ///< VstTimeInfo::timeSigNumerator and VstTimeInfo::timeSigDenominator valid
    kVstSmpteValid = 0x4000,             ///< VstTimeInfo::smpteOffset and VstTimeInfo::smpteFrameRate valid
    kVstClockValid = 0x00008000          ///< VstTimeInfo::samplesToNextClock valid
} VstTimeInfoFlags;

typedef enum AudioMasterOpcodesX {
    audioMasterAutomate = 0, ///< [index]: parameter index [opt]: parameter value  @see AudioEffect::setParameterAutomated
    audioMasterVersion = 1,
    audioMasterGetTime = 7   ///< [return value]: #VstTimeInfo* or null if not supported [value]: request mask  @see VstTimeInfoFlags @see AudioEffectX::getTimeInfo
} AudioMasterOpcodesX;

typedef enum AEffectOpcodes {

    effOpen = 0,
    effClose = 1,

    effSetProgram = 2,
    effGetProgram = 3,
    effSetProgramName = 4,
    effGetProgramName = 5,
    
    effGetParamLabel = 6,
    effGetParamDisplay = 7,
    effGetParamName = 8,
    
    effGetVu = 9,

    effSetSampleRate = 10,
    effSetBlockSize = 11,
    effMainsChanged = 12,

    effEditGetRect = 13,
    effEditOpen = 14,
    effEditClose = 15,

    effEditDraw = 16,
    effEditMouse = 17,
    effEditKey = 18,

    effEditIdle = 19,
    
    effEditTop = 20,
    effEditSleep = 21,
    effIdentify = 22,
    
    effGetChunk = 23,
    effSetChunk = 24,
 
    effProcessEvents = 25,
    
    effCanDo = 51,
    
    effGetVstVersion = 58,
    
    effCanBeAutomated = 26,
    effGetParameterProperties = 56,
    effGetPlugCategory = 35,
    effGetProgramNameIndexed = 29,
    
    effGetEffectName = 45,
    effGetVendorString = 47,
    effGetProductString = 48,
    effGetVendorVersion = 49,
    effVendorSpecific = 50,
    effSetBypass = 44,
    
    effGetTailSize = 52,
    effStartProcess = 71,
	effStopProcess = 72

} AEffectOpcodes;


typedef struct VstTimeInfo
{
//-------------------------------------------------------------------------------------------------------
	double samplePos;				///< current Position in audio samples (always valid)
	double sampleRate;				///< current Sample Rate in Herz (always valid)
	double nanoSeconds;				///< System Time in nanoseconds (10^-9 second)
	double ppqPos;					///< Musical Position, in Quarter Note (1.0 equals 1 Quarter Note)
	double tempo;					///< current Tempo in BPM (Beats Per Minute)
	double barStartPos;				///< last Bar Start Position, in Quarter Note
	double cycleStartPos;			///< Cycle Start (left locator), in Quarter Note
	double cycleEndPos;				///< Cycle End (right locator), in Quarter Note
	long timeSigNumerator;			///< Time Signature Numerator (e.g. 3 for 3/4)
	long timeSigDenominator;		///< Time Signature Denominator (e.g. 4 for 3/4)
	long smpteOffset;				///< SMPTE offset (in SMPTE subframes (bits; 1/80 of a frame)). The current SMPTE position can be calculated using #samplePos, #sampleRate, and #smpteFrameRate.
	long smpteFrameRate;			///< @see VstSmpteFrameRate
	long samplesToNextClock;		///< MIDI Clock Resolution (24 Per Quarter Note), can be negative (nearest clock)
	long flags;						///< @see VstTimeInfoFlags
//-------------------------------------------------------------------------------------------------------
} VstTimeInfo;

//-------------------------------------------------------------------------------------------------------
/** Flags used in #VstParameterProperties. */
//-------------------------------------------------------------------------------------------------------
typedef enum VstParameterFlags
{
//-------------------------------------------------------------------------------------------------------
	kVstParameterIsSwitch				 = 1,	///< parameter is a switch (on/off)
	kVstParameterUsesIntegerMinMax		 = 2,	///< minInteger, maxInteger valid
	kVstParameterUsesFloatStep			 = 4,	///< stepFloat, smallStepFloat, largeStepFloat valid
	kVstParameterUsesIntStep			 = 8,	///< stepInteger, largeStepInteger valid
	kVstParameterSupportsDisplayIndex 	 = 0x10,	///< displayIndex valid
	kVstParameterSupportsDisplayCategory = 0x20,	///< category, etc. valid
	kVstParameterCanRamp				 = 0x40	///< set if parameter value can ramp up/down
//-------------------------------------------------------------------------------------------------------
} VstParameterFlags;

typedef struct VstParameterProperties
{
//-------------------------------------------------------------------------------------------------------
	float stepFloat;			///< float step
	float smallStepFloat;		///< small float step
	float largeStepFloat;		///< large float step
	unsigned char label[64];///< parameter label
	long flags;				///< @see VstParameterFlags
	long minInteger;		///< integer minimum
	long maxInteger;		///< integer maximum
	long stepInteger;		///< integer step
	long largeStepInteger;	///< large integer step
	unsigned char shortLabel[8];	///< short label, recommended: 6 + delimiter

	// The following are for remote controller display purposes.
	// Note that the kVstParameterSupportsDisplayIndex flag must be set.
	// Host can scan all parameters, and find out in what order
	// to display them:

	short displayIndex;		///< index where this parameter should be displayed (starting with 0)

	// Host can also possibly display the parameter group (category), such as...
	// ---------------------------
	// Osc 1
	// Wave  Detune  Octave  Mod
	// ---------------------------
	// ...if the plug-in supports it (flag #kVstParameterSupportsDisplayCategory)

	short category;			///< 0: no category, else group index + 1
	short numParametersInCategory;			///< number of parameters in category
	short reserved;			///< zero
	unsigned char categoryLabel[24];	///< category label, e.g. "Osc 1" 

	unsigned char future[16];			///< reserved for future use
//-------------------------------------------------------------------------------------------------------
} VstParameterProperties;

//-------------------------------------------------------------------------------------------------------
/** A block of events for the current processed audio block. */
//-------------------------------------------------------------------------------------------------------
struct VstEvents
{
//-------------------------------------------------------------------------------------------------------
	long numEvents;		///< number of Events in array
	PTR reserved;		///< zero (Reserved for future use)
	PTR pEvents;	///< event pointer array, variable size
//-------------------------------------------------------------------------------------------------------
};

typedef enum VstEventTypes
{
//-------------------------------------------------------------------------------------------------------
	kVstMidiType = 1,		///< MIDI event  @see VstMidiEvent
	kVstAudioType,		///< \deprecated unused event type
	kVstVideoType,		///< \deprecated unused event type
	kVstParameterType,	///< \deprecated unused event type
	kVstTriggerType,	///< \deprecated unused event type
	kVstSysExType			///< MIDI system exclusive  @see VstMidiSysexEvent
//-------------------------------------------------------------------------------------------------------
} VstEventTypes;

typedef struct VstEvent
{
//-------------------------------------------------------------------------------------------------------
	long type;			///< @see VstEventTypes
	long byteSize;		///< size of this event, excl. type and byteSize
	long deltaFrames;	///< sample frames related to the current block start sample position
	long flags;			///< generic flags, none defined yet

	unsigned char data[16];			///< data size may vary, depending on event type
//-------------------------------------------------------------------------------------------------------
} VstEvent;

//-------------------------------------------------------------------------------------------------------
/** MIDI Event (to be casted from VstEvent). */
//-------------------------------------------------------------------------------------------------------
typedef struct VstMidiEvent
{
//-------------------------------------------------------------------------------------------------------
	long type;			///< #kVstMidiType
	long byteSize;		///< sizeof (VstMidiEvent)
	long deltaFrames;	///< sample frames related to the current block start sample position
	long flags;			///< @see VstMidiEventFlags
	long noteLength;	///< (in sample frames) of entire note, if available, else 0
	long noteOffset;	///< offset (in sample frames) into note from note start if available, else 0
	unsigned char midiData[4];		///< 1 to 3 MIDI bytes; midiData[3] is reserved (zero)
	unsigned char detune;			///< -64 to +63 cents; for scales other than 'well-tempered' ('microtuning')
	unsigned char noteOffVelocity;	///< Note Off Velocity [0, 127]
	unsigned char reserved1;		///< zero (Reserved for future use)
	unsigned char reserved2;		///< zero (Reserved for future use)
//-------------------------------------------------------------------------------------------------------
} VstMidiEvent;

//-------------------------------------------------------------------------------------------------------
/** Flags used in #VstMidiEvent. */
//-------------------------------------------------------------------------------------------------------
typedef enum VstMidiEventFlags
{
//-------------------------------------------------------------------------------------------------------
	kVstMidiEventIsRealtime = 1 << 0	///< means that this event is played life (not in playback from a sequencer track).\n This allows the Plug-In to handle these flagged events with higher priority, especially when the Plug-In has a big latency (AEffect::initialDelay)
//-------------------------------------------------------------------------------------------------------
} VstMidiEventFlags;

//-------------------------------------------------------------------------------------------------------
/** MIDI Sysex Event (to be casted from #VstEvent). */
//-------------------------------------------------------------------------------------------------------
typedef struct VstMidiSysexEvent
{
//-------------------------------------------------------------------------------------------------------
	long type;			///< #kVstSysexType
	long byteSize;		///< sizeof (VstMidiSysexEvent)
	long deltaFrames;	///< sample frames related to the current block start sample position
	long flags;			///< none defined yet (should be zero)
	long dumpBytes;		///< byte size of sysexDump
	long resvd1;		///< zero (Reserved for future use)
	PTR sysexDump;		///< sysex dump
	PTR resvd2;		///< zero (Reserved for future use)
//-------------------------------------------------------------------------------------------------------
} VstMidiSysexEvent;

typedef enum VstPlugCategory
{
//-------------------------------------------------------------------------------------------------------
    kPlugCategUnknown = 0,		///< Unknown, category not implemented
    kPlugCategEffect,			///< Simple Effect
    kPlugCategSynth,			///< VST Instrument (Synths, samplers,...)
    kPlugCategAnalysis,			///< Scope, Tuner, ...
    kPlugCategMastering,		///< Dynamics, ...
	kPlugCategSpacializer,		///< Panners, ...
	kPlugCategRoomFx,			///< Delays and Reverbs
	kPlugSurroundFx,			///< Dedicated surround processor
	kPlugCategRestoration,		///< Denoiser, ...
	kPlugCategOfflineProcess,	///< Offline Process
	kPlugCategShell,			///< Plug-in is container of other plug-ins  @see effShellGetNextPlugin
	kPlugCategGenerator,		///< ToneGenerator, ...

	kPlugCategMaxCount			///< Marker to count the categories
//-------------------------------------------------------------------------------------------------------
} VstPlugCategory;



[dllname("*")]
module vst2x {

	const BSTR canDoSendVstEvents = "sendVstEvents" ; ///< plug-in will send Vst events to Host
	const BSTR canDoSendVstMidiEvent = "sendVstMidiEvent" ; ///< plug-in will send MIDI events to Host
	const BSTR canDoReceiveVstEvents = "receiveVstEvents" ; ///< plug-in can receive MIDI events from Host
	const BSTR canDoReceiveVstMidiEvent = "receiveVstMidiEvent"; ///< plug-in can receive MIDI events from Host
	const BSTR canDoReceiveVstTimeInfo = "receiveVstTimeInfo"; ///< plug-in can receive Time info from Host
	const BSTR canDoOffline = "offline"; ///< plug-in supports offline functions (#offlineNotify, #offlinePrepare, #offlineRun)
	const BSTR canDoMidiProgramNames = "midiProgramNames"; ///< plug-in supports function #getMidiProgramName ()
	const BSTR canDoBypass = "bypass";

	const BSTR canDoSendVstTimeInfo = "sendVstTimeInfo"; ///< Host supports send of VstTimeInfo to plug-in
	const BSTR canDoReportConnectionChanges = "reportConnectionChanges"; ///< Host will indicates the plug-in when something change in plug-in´s routing/connections with #suspend/#resume/#setSpeakerArrangement 
	const BSTR canDoAcceptIOChanges = "acceptIOChanges"; ///< Host supports #ioChanged ()
	const BSTR canDoSizeWindow = "sizeWindow"; ///< used by VSTGUI
	const BSTR canDoOpenFileSelector = "openFileSelector"; ///< Host supports function #openFileSelector ()
	const BSTR canDoCloseFileSelector = "closeFileSelector"; ///< Host supports function #closeFileSelector ()
	const BSTR canDoStartStopProcess = "startStopProcess"; ///< Host supports functions #startProcess () and #stopProcess ()
	const BSTR canDoShellCategory = "shellCategory"; ///< 'shell' handling via uniqueID. If supported by the Host and the Plug-in has the category #kPlugCategShell
	const BSTR canDoSendVstMidiEventFlagIsRealtime = "sendVstMidiEventFlagIsRealtime"; ///< Host supports flags for #VstMidiEvent

	const int kEffectMagic = 0x56737450;
	const int SIZEOF_VstParameterProperties = 0x98;
	const int kVstMaxProductStrLen = 64;
	const int kVstMaxVendorStrLen = 64;
	const int kVstMaxEffectNameLen = 32;
	const int kVstMaxProgNameLen = 24;
	const int kVstMaxParamStrLen = 8;
}