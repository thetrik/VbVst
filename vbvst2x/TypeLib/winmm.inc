
typedef struct WAVEFORMATEX {
  short  wFormatTag;
  short  nChannels;
  long nSamplesPerSec;
  long nAvgBytesPerSec;
  short  nBlockAlign;
  short  wBitsPerSample;
  short  cbSize;
} WAVEFORMATEX;

typedef struct MMCKINFO {
  long ckid;
  long  cksize;
  long fccType;
  long  dwDataOffset;
  long  dwFlags;
} MMCKINFO;

typedef struct MMIOINFO{
  long dwFlags;
  long fccIOProc;
  PTR pIOProc;
  long wErrorRet;
  HANDLE hTask;
  long cchBuffer;
  PTR pchBuffer;
  PTR pchNext;
  PTR pchEndRead;
  PTR pchEndWrite;
  long lBufOffset;
  long lDiskOffset;
  long adwInfo[4];
  long dwReserved1;
  long dwReserved2;
  HANDLE hmmio;
} MMIOINFO;

typedef struct WAVEHDR {
	PTR lpData;
	long dwBufferLength;
	long dwBytesRecorded;
	PTR dwUser;
	long dwFlags;
	long dwLoops;
	PTR lpNext;
	PTR reserved;
} WAVEHDR;

[dllname("WINMM.DLL")] 
module winmm{
	
	const int WAVE_FORMAT_PCM = 1;
	const int MMIO_READWRITE = 2;
	const int MMIO_FINDCHUNK = 0x10;
	const int MMIO_FINDRIFF = 0x20;
	const int FOURCC_RIFF = 0x46464952;
	const int FOURCC_WAVE = 0x45564157;
	const int FOURCC_FMT = 0x20746D66;
	const int FOURCC_DATA = 0x61746164;
	const int CALLBACK_WINDOW = 0x10000;
	const int MM_WOM_DONE = 0x3BD;
	const int WAVE_MAPPER = -1;

	[entry("mmioOpenW")]
	HANDLE mmioOpen(
	   [in] LPWSTR szFilename,
	   [in] void* lpmmioinfo,
	   [in] long dwOpenFlags
	);

	[entry("mmioClose")]
	long mmioClose(
	   [in] HANDLE hmmio,
	   [in] long wFlags
	);

	[entry("mmioDescend")]
	long mmioDescend(
	   [in] HANDLE hmmio,
	   [in] MMCKINFO* pmmcki,
	   [in] void* pmmckiParent,
	   [in] long fuDescend
	);

	[entry("mmioRead")]
	long mmioRead(
	   [in] HANDLE hmmio,
	   [in] void* pch,
	   [in] long cch
	);

	[entry("mmioStringToFOURCCW")]
	long mmioStringToFOURCC(
	   [in] LPWSTR sz,
	   [in] long uFlags
	);

	[entry("mmioAscend")]
	long mmioAscend(
	   [in] HANDLE hmmio,
	   [in] MMCKINFO* pmmcki,
	   [in] long fuAscend
	);

	[entry("waveOutOpen")]
	long waveOutOpen(
		[in, out] HANDLE* phwo,
		[in] long uDeviceID,
		[in] WAVEFORMATEX* pwfx,
		[in] PTR dwCallback,
		[in] PTR dwInstance,
		[in] long fdwOpen
	);

	[entry("waveOutPrepareHeader")]
	long waveOutPrepareHeader(
		[in] HANDLE hwo,
		[in] WAVEHDR* pwh,
		[in] long cbwh
	);

	[entry("waveOutUnprepareHeader")]
	long waveOutUnprepareHeader(
		[in] HANDLE hwo,
		[in] WAVEHDR* pwh,
		[in] long cbwh
	);

	[entry("waveOutWrite")]
	long waveOutWrite(
		[in] HANDLE hwo,
		[in] WAVEHDR* pwh,
		[in] long cbwh
	);

	[entry("waveOutClose")]
	long waveOutClose(
		[in] HANDLE hwo
	);

	[entry("waveOutReset")]
	long waveOutReset(
		[in] HANDLE hwo
	);

	[entry("waveOutPause")]
	long waveOutPause(
		[in] HANDLE hwo
	);

	[entry("waveOutRestart")]
	long waveOutRestart(
		[in] HANDLE hwo
	);

}
