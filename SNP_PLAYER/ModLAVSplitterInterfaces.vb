Imports System
Imports System.Runtime.InteropServices
Imports System.Security

Module ModLAVSplitterInterfaces

    ' Define the LAVSubtitleMode enum
    Public Enum LAVSubtitleMode
        LAVSubtitleMode_NoSubs
        LAVSubtitleMode_ForcedOnly
        LAVSubtitleMode_Default
        LAVSubtitleMode_Advanced
    End Enum

    ' LAVF Settings interface
    <Guid("774A919D-EA95-4A87-8A1E-F48ABE8499C7")>
    <ComImport()>
    <SuppressUnmanagedCodeSecurity()>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ILAVFSettings
        ' Switch to Runtime Config mode
        <PreserveSig>
        Function SetRuntimeConfig(ByVal bRuntimeConfig As Boolean) As Integer

        ' Retrieve preferred languages as ISO 639-2 language codes, comma separated
        ' Memory for the string will be allocated, and has to be freed by the caller with CoTaskMemFree
        <PreserveSig>
        Function GetPreferredLanguages(<Out> ByRef ppLanguages As IntPtr) As Integer

        ' Set preferred languages as ISO 639-2 language codes, comma separated
        ' To reset to no preferred language, pass NULL or the empty string
        <PreserveSig>
        Function SetPreferredLanguages(ByVal pLanguages As String) As Integer

        ' Retrieve preferred subtitle languages as ISO 639-2 language codes, comma separated
        ' Memory for the string will be allocated, and has to be freed by the caller with CoTaskMemFree
        <PreserveSig>
        Function GetPreferredSubtitleLanguages(<Out> ByRef ppLanguages As IntPtr) As Integer

        ' Set preferred subtitle languages as ISO 639-2 language codes, comma separated
        ' To reset to no preferred language, pass NULL or the empty string
        <PreserveSig>
        Function SetPreferredSubtitleLanguages(ByVal pLanguages As String) As Integer

        ' Get the current subtitle mode
        ' See enum for possible values
        <PreserveSig>
        Function GetSubtitleMode() As LAVSubtitleMode

        ' Set the current subtitle mode
        ' See enum for possible values
        <PreserveSig>
        Function SetSubtitleMode(ByVal mode As LAVSubtitleMode) As Integer

        ' Get the subtitle matching language flag
        ' Deprecated
        <PreserveSig>
        Function GetSubtitleMatchingLanguage() As Boolean

        ' Set the subtitle matching language flag
        ' Deprecated
        <PreserveSig>
        Function SetSubtitleMatchingLanguage(ByVal dwMode As Boolean) As Integer

        ' Control whether a special "Forced Subtitles" stream will be created for PGS subs
        <PreserveSig>
        Function GetPGSForcedStream() As Boolean

        ' Control whether a special "Forced Subtitles" stream will be created for PGS subs
        <PreserveSig>
        Function SetPGSForcedStream(ByVal bFlag As Boolean) As Integer

        ' Get the PGS forced subs config
        <PreserveSig>
        Function GetPGSOnlyForced() As Boolean

        ' Set the PGS forced subs config
        <PreserveSig>
        Function SetPGSOnlyForced(ByVal bForced As Boolean) As Integer

        ' Get the VC-1 Timestamp Processing mode
        <PreserveSig>
        Function GetVC1TimestampMode() As Integer

        ' Set the VC-1 Timestamp Processing mode
        <PreserveSig>
        Function SetVC1TimestampMode(ByVal iMode As Integer) As Integer

        ' Set whether substreams should be shown as a separate stream
        <PreserveSig>
        Function SetSubstreamsEnabled(ByVal bSubStreams As Boolean) As Integer

        ' Check whether substreams should be shown as a separate stream
        <PreserveSig>
        Function GetSubstreamsEnabled() As Boolean

        ' Deprecated - no longer required
        <PreserveSig>
        Function SetVideoParsingEnabled(ByVal bEnabled As Boolean) As Integer

        ' Deprecated - no longer required
        <PreserveSig>
        Function GetVideoParsingEnabled() As Boolean

        ' Deprecated - no longer required
        <PreserveSig>
        Function SetFixBrokenHDPVR(ByVal bEnabled As Boolean) As Integer

        ' Deprecated - no longer required
        <PreserveSig>
        Function GetFixBrokenHDPVR() As Boolean

        ' Control whether the given format is enabled
        <PreserveSig>
        Function SetFormatEnabled(ByVal strFormat As String, ByVal bEnabled As Boolean) As Integer

        ' Check if the given format is enabled
        <PreserveSig>
        Function IsFormatEnabled(ByVal strFormat As String) As Boolean

        ' Set whether LAV Splitter should remove the filter connected to its Audio Pin when the audio stream is changed
        <PreserveSig>
        Function SetStreamSwitchRemoveAudio(ByVal bEnabled As Boolean) As Integer

        ' Query if LAV Splitter should remove the filter connected to its Audio Pin when the audio stream is changed
        <PreserveSig>
        Function GetStreamSwitchRemoveAudio() As Boolean

        ' Advanced Subtitle configuration. Refer to the documentation for details.
        ' Memory for the string will be allocated, and has to be freed by the caller with CoTaskMemFree
        <PreserveSig>
        Function GetAdvancedSubtitleConfig(<Out> ByRef ppAdvancedConfig As IntPtr) As Integer

        ' Advanced Subtitle configuration. Refer to the documentation for details.
        ' To reset the config, pass NULL or the empty string.
        <PreserveSig>
        Function SetAdvancedSubtitleConfig(ByVal pAdvancedConfig As String) As Integer

        ' Set if LAV Splitter should prefer audio streams for the hearing or visually impaired
        <PreserveSig>
        Function SetUseAudioForHearingVisuallyImpaired(ByVal bEnabled As Boolean) As Integer

        ' Get if LAV Splitter should prefer audio streams for the hearing or visually impaired
        <PreserveSig>
        Function GetUseAudioForHearingVisuallyImpaired() As Boolean

        ' Set the maximum queue size, in megabytes
        <PreserveSig>
        Function SetMaxQueueMemSize(ByVal dwMaxSize As UInteger) As Integer

        ' Get the maximum queue size, in megabytes
        <PreserveSig>
        Function GetMaxQueueMemSize() As UInteger

        ' Toggle Tray Icon
        <PreserveSig>
        Function SetTrayIcon(ByVal bEnabled As Boolean) As Integer

        ' Get Tray Icon
        <PreserveSig>
        Function GetTrayIcon() As Boolean

        ' Toggle whether higher quality audio streams are preferred
        <PreserveSig>
        Function SetPreferHighQualityAudioStreams(ByVal bEnabled As Boolean) As Integer

        ' Get whether higher quality audio streams are preferred
        <PreserveSig>
        Function GetPreferHighQualityAudioStreams() As Boolean

        ' Toggle whether Matroska Linked Segments should be loaded from other files
        <PreserveSig>
        Function SetLoadMatroskaExternalSegments(ByVal bEnabled As Boolean) As Integer

        ' Get whether Matroska Linked Segments should be loaded from other files
        <PreserveSig>
        Function GetLoadMatroskaExternalSegments() As Boolean

        ' Get the list of available formats
        ' Memory for the string array will be allocated, and has to be freed by the caller with CoTaskMemFree
        <PreserveSig>
        Function GetFormats(<Out> ByRef formats As IntPtr, <Out> ByRef nFormats As UInteger) As Integer

        ' Set the duration (in ms) of analysis for network streams
        <PreserveSig>
        Function SetNetworkStreamAnalysisDuration(ByVal dwDuration As UInteger) As Integer

        ' Get the duration (in ms) of analysis for network streams
        <PreserveSig>
        Function GetNetworkStreamAnalysisDuration() As UInteger

        ' Set the maximum queue size, in number of packets
        <PreserveSig>
        Function SetMaxQueueSize(ByVal dwMaxSize As UInteger) As Integer

        ' Get the maximum queue size, in number of packets
        <PreserveSig>
        Function GetMaxQueueSize() As UInteger

        ' Set if LAV Splitter should reselect subs based on given rules when audio stream is changed
        <PreserveSig>
        Function SetStreamSwitchReselectSubtitles(ByVal bEnabled As Boolean) As Integer

        ' Query if LAV Splitter should reselect subs based on given rules when audio stream is changed
        <PreserveSig>
        Function GetStreamSwitchReselectSubtitles() As Boolean
    End Interface
End Module
