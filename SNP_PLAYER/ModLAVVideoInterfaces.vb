Imports System
Imports System.Runtime.InteropServices
Imports System.Security

Module ModLAVVideoInterfaces

    ' Define the LAVVideoCodec enum
    Public Enum LAVVideoCodec
        Codec_H264
        Codec_VC1
        Codec_MPEG1
        Codec_MPEG2
        Codec_MPEG4
        Codec_MSMPEG4
        Codec_VP8
        Codec_WMV3
        Codec_WMV12
        Codec_MJPEG
        Codec_Theora
        Codec_FLV1
        Codec_VP6
        Codec_SVQ
        Codec_H261
        Codec_H263
        Codec_Indeo
        Codec_TSCC
        Codec_Fraps
        Codec_HuffYUV
        Codec_QTRle
        Codec_DV
        Codec_Bink
        Codec_Smacker
        Codec_RV12
        Codec_RV34
        Codec_Lagarith
        Codec_Cinepak
        Codec_Camstudio
        Codec_QPEG
        Codec_ZLIB
        Codec_QTRpza
        Codec_PNG
        Codec_MSRLE
        Codec_ProRes
        Codec_UtVideo
        Codec_Dirac
        Codec_DNxHD
        Codec_MSVideo1
        Codec_8BPS
        Codec_LOCO
        Codec_ZMBV
        Codec_VCR1
        Codec_Snow
        Codec_FFV1
        Codec_v210
        Codec_JPEG2000
        Codec_VMNC
        Codec_FLIC
        Codec_G2M
        Codec_ICOD
        Codec_THP
        Codec_HEVC
        Codec_VP9
        Codec_TrueMotion
        Codec_VP7
        Codec_H264MVC
        Codec_CineformHD
        Codec_MagicYUV
        Codec_AV1
        Codec_VVC
        Codec_VideoNB
    End Enum

    ' Define the LAVVideoHWCodec enum
    Public Enum LAVVideoHWCodec
        HWCodec_H264 = LAVVideoCodec.Codec_H264
        HWCodec_VC1 = LAVVideoCodec.Codec_VC1
        HWCodec_MPEG2 = LAVVideoCodec.Codec_MPEG2
        HWCodec_MPEG4 = LAVVideoCodec.Codec_MPEG4
        HWCodec_MPEG2DVD
        HWCodec_HEVC
        HWCodec_VP9
        HWCodec_H264MVC
        HWCodec_AV1
        HWCodec_NB
    End Enum

    ' Define the LAVHWAccel enum
    Public Enum LAVHWAccel
        HWAccel_None
        HWAccel_CUDA
        HWAccel_QuickSync
        HWAccel_DXVA2
        HWAccel_DXVA2CopyBack = LAVHWAccel.HWAccel_DXVA2
        HWAccel_DXVA2Native
        HWAccel_D3D11
        HWAccel_NB
    End Enum

    ' Define the LAVHWDeintModes enum
    Public Enum LAVHWDeintModes
        HWDeintMode_Weave
        HWDeintMode_BOB ' Deprecated
        HWDeintMode_Hardware
    End Enum

    ' Define the LAVSWDeintModes enum
    Public Enum LAVSWDeintModes
        SWDeintMode_None
        SWDeintMode_YADIF
        SWDeintMode_W3FDIF_Simple
        SWDeintMode_W3FDIF_Complex
        SWDeintMode_BWDIF
    End Enum

    ' Define the LAVDeintMode enum
    Public Enum LAVDeintMode
        DeintMode_Auto
        DeintMode_Aggressive
        DeintMode_Force
        DeintMode_Disable
    End Enum

    ' Define the LAVDeintOutput enum
    Public Enum LAVDeintOutput
        DeintOutput_FramePerField
        DeintOutput_FramePer2Field
    End Enum

    ' Define the LAVDeintFieldOrder enum
    Public Enum LAVDeintFieldOrder
        DeintFieldOrder_Auto
        DeintFieldOrder_TopFieldFirst
        DeintFieldOrder_BottomFieldFirst
    End Enum

    ' Define the LAVOutPixFmts enum
    Public Enum LAVOutPixFmts
        LAVOutPixFmt_None = -1
        LAVOutPixFmt_YV12
        LAVOutPixFmt_NV12
        LAVOutPixFmt_YUY2
        LAVOutPixFmt_UYVY
        LAVOutPixFmt_AYUV
        LAVOutPixFmt_P010
        LAVOutPixFmt_P210
        LAVOutPixFmt_Y410
        LAVOutPixFmt_P016
        LAVOutPixFmt_P216
        LAVOutPixFmt_Y416
        LAVOutPixFmt_RGB32
        LAVOutPixFmt_RGB24
        LAVOutPixFmt_v210
        LAVOutPixFmt_v410
        LAVOutPixFmt_YV16
        LAVOutPixFmt_YV24
        LAVOutPixFmt_RGB48
        LAVOutPixFmt_NB
    End Enum

    ' Define the LAVDitherMode enum
    Public Enum LAVDitherMode
        LAVDither_Ordered
        LAVDither_Random
    End Enum

    ' Flags for HW Resolution support
    Public Const LAVHWResFlag_SD As Integer = &H1
    Public Const LAVHWResFlag_HD As Integer = &H2
    Public Const LAVHWResFlag_UHD As Integer = &H4

    ' LAV Video configuration interface
    <Guid("FA40D6E9-4D38-4761-ADD2-71A9EC5FD32F")>
    <ComImport()>
    <SuppressUnmanagedCodeSecurity()>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ILAVVideoSettings
        ' Switch to Runtime Config mode
        <PreserveSig>
        Function SetRuntimeConfig(ByVal bRuntimeConfig As Boolean) As Integer

        ' Configure which codecs are enabled
        <PreserveSig>
        Function GetFormatConfiguration(ByVal vCodec As LAVVideoCodec) As Boolean
        <PreserveSig>
        Function SetFormatConfiguration(ByVal vCodec As LAVVideoCodec, ByVal bEnabled As Boolean) As Integer

        ' Set the number of threads for Multi-Threaded decoding
        <PreserveSig>
        Function SetNumThreads(ByVal dwNum As UInteger) As Integer

        ' Get the number of threads for Multi-Threaded decoding
        <PreserveSig>
        Function GetNumThreads() As UInteger

        ' Set the aspect ratio encoded in the stream
        <PreserveSig>
        Function SetStreamAR(ByVal bStreamAR As UInteger) As Integer

        ' Get the aspect ratio encoded in the stream
        <PreserveSig>
        Function GetStreamAR() As UInteger

        ' Configure which pixel formats are enabled for output
        <PreserveSig>
        Function GetPixelFormat(ByVal pixFmt As LAVOutPixFmts) As Boolean
        <PreserveSig>
        Function SetPixelFormat(ByVal pixFmt As LAVOutPixFmts, ByVal bEnabled As Boolean) As Integer

        ' Set the RGB output range for the YUV->RGB conversion
        <PreserveSig>
        Function SetRGBOutputRange(ByVal dwRange As UInteger) As Integer

        ' Get the RGB output range for the YUV->RGB conversion
        <PreserveSig>
        Function GetRGBOutputRange() As UInteger

        ' Set the deinterlacing field order of the hardware decoder
        <PreserveSig>
        Function SetDeintFieldOrder(ByVal fieldOrder As LAVDeintFieldOrder) As Integer

        ' Get the deinterlacing field order of the hardware decoder
        <PreserveSig>
        Function GetDeintFieldOrder() As LAVDeintFieldOrder

        ' DEPRECATED, use SetDeinterlacingMode
        <PreserveSig>
        Function SetDeintAggressive(ByVal bAggressive As Boolean) As Integer

        ' DEPRECATED, use GetDeinterlacingMode
        <PreserveSig>
        Function GetDeintAggressive() As Boolean

        ' DEPRECATED, use SetDeinterlacingMode
        <PreserveSig>
        Function SetDeintForce(ByVal bForce As Boolean) As Integer

        ' DEPRECATED, use GetDeinterlacingMode
        <PreserveSig>
        Function GetDeintForce() As Boolean

        ' Check if the specified HWAccel is supported
        <PreserveSig>
        Function CheckHWAccelSupport(ByVal hwAccel As LAVHWAccel) As UInteger

        ' Set which HW Accel method is used
        <PreserveSig>
        Function SetHWAccel(ByVal hwAccel As LAVHWAccel) As Integer

        ' Get which HW Accel method is active
        <PreserveSig>
        Function GetHWAccel() As LAVHWAccel

        ' Set which codecs should use HW Acceleration
        <PreserveSig>
        Function SetHWAccelCodec(ByVal hwAccelCodec As LAVVideoHWCodec, ByVal bEnabled As Boolean) As Integer

        ' Get which codecs should use HW Acceleration
        <PreserveSig>
        Function GetHWAccelCodec(ByVal hwAccelCodec As LAVVideoHWCodec) As Boolean

        ' Set the deinterlacing mode used by the hardware decoder
        <PreserveSig>
        Function SetHWAccelDeintMode(ByVal deintMode As LAVHWDeintModes) As Integer

        ' Get the deinterlacing mode used by the hardware decoder
        <PreserveSig>
        Function GetHWAccelDeintMode() As LAVHWDeintModes

        ' Set the deinterlacing output for the hardware decoder
        <PreserveSig>
        Function SetHWAccelDeintOutput(ByVal deintOutput As LAVDeintOutput) As Integer

        ' Get the deinterlacing output for the hardware decoder
        <PreserveSig>
        Function GetHWAccelDeintOutput() As LAVDeintOutput

        ' deprecated. HQ mode is only supported by NVIDIA CUVID/NVDEC and officially deprecated by NVIDIA
        <PreserveSig>
        Function SetHWAccelDeintHQ(ByVal bHQ As Boolean) As Integer

        <PreserveSig>
        Function GetHWAccelDeintHQ() As Boolean

        ' Set the software deinterlacing mode used
        <PreserveSig>
        Function SetSWDeintMode(ByVal deintMode As LAVSWDeintModes) As Integer

        ' Get the software deinterlacing mode used
        <PreserveSig>
        Function GetSWDeintMode() As LAVSWDeintModes

        ' Set the software deinterlacing output
        <PreserveSig>
        Function SetSWDeintOutput(ByVal deintOutput As LAVDeintOutput) As Integer

        ' Get the software deinterlacing output
        <PreserveSig>
        Function GetSWDeintOutput() As LAVDeintOutput

        ' DEPRECATED, use SetDeinterlacingMode
        <PreserveSig>
        Function SetDeintTreatAsProgressive(ByVal bEnabled As Boolean) As Integer

        ' DEPRECATED, use GetDeinterlacingMode
        <PreserveSig>
        Function GetDeintTreatAsProgressive() As Boolean

        ' Set the dithering mode used
        <PreserveSig>
        Function SetDitherMode(ByVal ditherMode As LAVDitherMode) As Integer

        ' Get the dithering mode used
        <PreserveSig>
        Function GetDitherMode() As LAVDitherMode

        ' Set if the MS WMV9 DMO Decoder should be used for VC-1/WMV3
        <PreserveSig>
        Function SetUseMSWMV9Decoder(ByVal bEnabled As Boolean) As Integer

        ' Get if the MS WMV9 DMO Decoder should be used for VC-1/WMV3
        <PreserveSig>
        Function GetUseMSWMV9Decoder() As Boolean

        ' Set if DVD Video support is enabled
        <PreserveSig>
        Function SetDVDVideoSupport(ByVal bEnabled As Boolean) As Integer

        ' Get if DVD Video support is enabled
        <PreserveSig>
        Function GetDVDVideoSupport() As Boolean

        ' Set the HW Accel Resolution Flags
        <PreserveSig>
        Function SetHWAccelResolutionFlags(ByVal dwResFlags As UInteger) As Integer

        ' Get the HW Accel Resolution Flags
        <PreserveSig>
        Function GetHWAccelResolutionFlags() As UInteger

        ' Toggle Tray Icon
        <PreserveSig>
        Function SetTrayIcon(ByVal bEnabled As Boolean) As Integer

        ' Get Tray Icon
        <PreserveSig>
        Function GetTrayIcon() As Boolean

        ' Set the Deint Mode
        <PreserveSig>
        Function SetDeinterlacingMode(ByVal deintMode As LAVDeintMode) As Integer

        ' Get the Deint Mode
        <PreserveSig>
        Function GetDeinterlacingMode() As LAVDeintMode

        ' Set the index of the GPU to be used for hardware decoding
        <PreserveSig>
        Function SetGPUDeviceIndex(ByVal dwDevice As UInteger) As Integer

        ' Get the number of available devices for the specified HWAccel
        <PreserveSig>
        Function GetHWAccelNumDevices(ByVal hwAccel As LAVHWAccel) As UInteger

        ' Get a list of available HWAccel devices for the specified HWAccel
        <PreserveSig>
        Function GetHWAccelDeviceInfo(ByVal hwAccel As LAVHWAccel, ByVal dwIndex As UInteger, <Out> ByRef pstrDeviceName As String, <Out> ByRef pdwDeviceIdentifier As UInteger) As Integer

        ' Get/Set the device for a specified HWAccel
        <PreserveSig>
        Function GetHWAccelDeviceIndex(ByVal hwAccel As LAVHWAccel, <Out> ByRef pdwDeviceIdentifier As UInteger) As UInteger

        <PreserveSig>
        Function SetHWAccelDeviceIndex(ByVal hwAccel As LAVHWAccel, ByVal dwIndex As UInteger, ByVal dwDeviceIdentifier As UInteger) As Integer

        ' Temporary Override for players to disable H.264 MVC decoding
        <PreserveSig>
        Function SetH264MVCDecodingOverride(ByVal bEnabled As Boolean) As Integer

        '  Enable the creation of the Closed Caption output pin
        <PreserveSig>
        Function SetEnableCCOutputPin(ByVal bEnabled As Boolean) As Integer
    End Interface

    ' LAV Video status interface
    <Guid("1CC2385F-36FA-41B1-9942-5024CE0235DC")>
    <ComImport()>
    <SuppressUnmanagedCodeSecurity()>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ILAVVideoStatus
        ' Get the name of the active decoder (can return NULL if none is active)
        <PreserveSig>
        Function GetActiveDecoderName() As IntPtr

        ' Get the name of the currently active hwaccel device
        <PreserveSig>
        Function GetHWAccelActiveDevice(<Out> ByRef pstrDeviceName As String) As Integer
    End Interface

End Module
