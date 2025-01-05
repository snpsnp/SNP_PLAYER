Imports System
Imports System.Runtime.InteropServices
Imports System.Security

Module ModLAVAudioInterfaces


    ' Define the LAVAudioCodec enum
    Public Enum LAVAudioCodec
        Codec_AAC
        Codec_AC3
        Codec_EAC3
        Codec_DTS
        Codec_MP2
        Codec_MP3
        Codec_TRUEHD
        Codec_FLAC
        Codec_VORBIS
        Codec_LPCM
        Codec_PCM
        Codec_WAVPACK
        Codec_TTA
        Codec_WMA2
        Codec_WMAPRO
        Codec_Cook
        Codec_RealAudio
        Codec_WMALL
        Codec_ALAC
        Codec_Opus
        Codec_AMR
        Codec_Nellymoser
        Codec_MSPCM
        Codec_Truespeech
        Codec_TAK
        Codec_ATRAC
        Codec_AudioNB
    End Enum

    ' Define the LAVBitstreamCodec enum
    Public Enum LAVBitstreamCodec
        Bitstream_AC3
        Bitstream_EAC3
        Bitstream_TRUEHD
        Bitstream_DTS
        Bitstream_DTSHD
        Bitstream_NB
    End Enum

    ' Define the LAVAudioSampleFormat enum
    Public Enum LAVAudioSampleFormat
        SampleFormat_None = -1
        SampleFormat_16
        SampleFormat_24
        SampleFormat_32
        SampleFormat_U8
        SampleFormat_FP32
        SampleFormat_Bitstream
        SampleFormat_NB
    End Enum

    ' Define the LAVAudioMixingMode enum
    Public Enum LAVAudioMixingMode
        MatrixEncoding_None
        MatrixEncoding_Dolby
        MatrixEncoding_DPLII
        MatrixEncoding_NB
    End Enum

    ' LAV Audio configuration interface
    <Guid("4158A22B-6553-45D0-8069-24716F8FF171")>
    <ComImport()>
    <SuppressUnmanagedCodeSecurity()>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ILAVAudioSettings
        ' Switch to Runtime Config mode
        <PreserveSig>
        Function SetRuntimeConfig(ByVal bRuntimeConfig As Boolean) As Integer

        ' Get and set Dynamic Range Compression
        <PreserveSig>
        Function GetDRC(ByRef pbDRCEnabled As Boolean, ByRef piDRCLevel As Integer) As Integer
        <PreserveSig>
        Function SetDRC(ByVal bDRCEnabled As Boolean, ByVal iDRCLevel As Integer) As Integer

        ' Get and set format configuration for specific codec
        <PreserveSig>
        Function GetFormatConfiguration(ByVal aCodec As LAVAudioCodec) As Boolean
        <PreserveSig>
        Function SetFormatConfiguration(ByVal aCodec As LAVAudioCodec, ByVal bEnabled As Boolean) As Integer

        ' Get and set bitstream configuration for specific codec
        <PreserveSig>
        Function GetBitstreamConfig(ByVal bsCodec As LAVBitstreamCodec) As Boolean
        <PreserveSig>
        Function SetBitstreamConfig(ByVal bsCodec As LAVBitstreamCodec, ByVal bEnabled As Boolean) As Integer

        ' Get and set DTS-HD framing
        <PreserveSig>
        Function GetDTSHDFraming() As Boolean
        <PreserveSig>
        Function SetDTSHDFraming(ByVal bHDFraming As Boolean) As Integer

        ' Get and set Auto A/V sync
        <PreserveSig>
        Function GetAutoAVSync() As Boolean
        <PreserveSig>
        Function SetAutoAVSync(ByVal bAutoSync As Boolean) As Integer

        ' Get and set output standard layout
        <PreserveSig>
        Function GetOutputStandardLayout() As Boolean
        <PreserveSig>
        Function SetOutputStandardLayout(ByVal bStdLayout As Boolean) As Integer

        ' Get and set expand mono to stereo
        <PreserveSig>
        Function GetExpandMono() As Boolean
        <PreserveSig>
        Function SetExpandMono(ByVal bExpandMono As Boolean) As Integer

        ' Get and set expand 6.1 to 7.1
        <PreserveSig>
        Function GetExpand61() As Boolean
        <PreserveSig>
        Function SetExpand61(ByVal bExpand61 As Boolean) As Integer

        ' Get and set allow raw PCM and SPDIF input
        <PreserveSig>
        Function GetAllowRawSPDIFInput() As Boolean
        <PreserveSig>
        Function SetAllowRawSPDIFInput(ByVal bAllow As Boolean) As Integer

        ' Get and set sample formats
        <PreserveSig>
        Function GetSampleFormat(ByVal format As LAVAudioSampleFormat) As Boolean
        <PreserveSig>
        Function SetSampleFormat(ByVal format As LAVAudioSampleFormat, ByVal bEnabled As Boolean) As Integer

        ' Get and set audio delay
        <PreserveSig>
        Function GetAudioDelay(ByRef pbEnabled As Boolean, ByRef pDelay As Integer) As Integer
        <PreserveSig>
        Function SetAudioDelay(ByVal bEnabled As Boolean, ByVal delay As Integer) As Integer

        ' Enable/Disable mixing
        <PreserveSig>
        Function SetMixingEnabled(ByVal bEnabled As Boolean) As Integer
        <PreserveSig>
        Function GetMixingEnabled() As Boolean

        ' Get and set mixing layout
        <PreserveSig>
        Function SetMixingLayout(ByVal dwLayout As UInteger) As Integer
        <PreserveSig>
        Function GetMixingLayout() As UInteger

        ' Set and get mixing flags
        <PreserveSig>
        Function SetMixingFlags(ByVal dwFlags As UInteger) As Integer
        <PreserveSig>
        Function GetMixingFlags() As UInteger

        ' Set and get mixing mode
        <PreserveSig>
        Function SetMixingMode(ByVal mixingMode As LAVAudioMixingMode) As Integer
        <PreserveSig>
        Function GetMixingMode() As LAVAudioMixingMode

        ' Set and get mixing levels
        <PreserveSig>
        Function SetMixingLevels(ByVal dwCenterLevel As UInteger, ByVal dwSurroundLevel As UInteger, ByVal dwLFELevel As UInteger) As Integer
        <PreserveSig>
        Function GetMixingLevels(ByRef dwCenterLevel As UInteger, ByRef dwSurroundLevel As UInteger, ByRef dwLFELevel As UInteger) As Integer

        ' Enable/Disable tray icon
        <PreserveSig>
        Function SetTrayIcon(ByVal bEnabled As Boolean) As Integer
        <PreserveSig>
        Function GetTrayIcon() As Boolean

        ' Enable/Disable sample format conversion dithering
        <PreserveSig>
        Function SetSampleConvertDithering(ByVal bEnabled As Boolean) As Integer
        <PreserveSig>
        Function GetSampleConvertDithering() As Boolean

        ' Enable/Disable suppress format changes
        <PreserveSig>
        Function SetSuppressFormatChanges(ByVal bEnabled As Boolean) As Integer
        <PreserveSig>
        Function GetSuppressFormatChanges() As Boolean

        ' Get and set output 5.1 legacy layout
        <PreserveSig>
        Function GetOutput51LegacyLayout() As Boolean
        <PreserveSig>
        Function SetOutput51LegacyLayout(ByVal b51Legacy As Boolean) As Integer

        ' Get and set bitstreaming fallback
        <PreserveSig>
        Function GetBitstreamingFallback() As Boolean
        <PreserveSig>
        Function SetBitstreamingFallback(ByVal bBitstreamingFallback As Boolean) As Integer
    End Interface

    ' LAV Audio Status Interface
    <Guid("A668B8F2-BA87-4F63-9D41-768F7DE9C50E")>
    <ComImport()>
    <SuppressUnmanagedCodeSecurity()>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ILAVAudioStatus
        ' Check if the given sample format is supported by the current playback chain
        <PreserveSig>
        Function IsSampleFormatSupported(ByVal sfCheck As LAVAudioSampleFormat) As Boolean

        ' Get details about the current decoding format
        <PreserveSig>
        Function GetDecodeDetails(ByRef pCodec As IntPtr, ByRef pDecodeFormat As IntPtr, ByRef pnChannels As Integer, ByRef pSampleRate As Integer, ByRef pChannelMask As UInteger) As Integer

        ' Get details about the current output format
        <PreserveSig>
        Function GetOutputDetails(ByRef pOutputFormat As IntPtr, ByRef pnChannels As Integer, ByRef pSampleRate As Integer, ByRef pChannelMask As UInteger) As Integer

        ' Enable Volume measurements
        <PreserveSig>
        Function EnableVolumeStats() As Integer

        ' Disable Volume measurements
        <PreserveSig>
        Function DisableVolumeStats() As Integer

        ' Get Volume Average for the given channel
        <PreserveSig>
        Function GetChannelVolumeAverage(ByVal nChannel As UShort, ByRef pfDb As Single) As Integer
    End Interface
End Module
