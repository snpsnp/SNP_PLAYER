﻿Imports System.Runtime.InteropServices

Public Enum ESTADO
    ABIERTO
    CERRADO
    REPRODUCIENDO
    PAUSADO
    DETENIDO
    SEEKING
End Enum

Public Enum GET_PIN_TYPE
    PT_NAME
    PT_NUMBER
End Enum

Public Class DSHOW_PLAYER
    Private mFG As DirectShowLib.IFilterGraph2

    Private bHasAudio As Boolean = False
    Private bHasVideo As Boolean = False

    'MEDIA MANIPULATION INTERFACES, TO BE RETRIEVED FROM THE FILTERGRAPH
    Private mMedia As DirectShowLib.IMediaControl
    Private mSound As DirectShowLib.IBasicAudio
    Private mVideo As DirectShowLib.IBasicVideo
    Private mEvent As DirectShowLib.IMediaEventEx
    Private mSeek As DirectShowLib.IMediaSeeking

    'SOURCE FILTER OBJECT
    Private mSourceFilter As DirectShowLib.IFileSourceFilter

    'LAV SPLITTER FILTER OBJECT
    Private mLAV_Splitter As DirectShowLib.IBaseFilter

    'LAV AUDIO DECODER FILTER OBJECT
    Private mLAV_Audio As DirectShowLib.IBaseFilter

    'LAV VIDEO DECODER FILTER
    Private mLAV_Video As DirectShowLib.IBaseFilter

    'DSOUND RENDERER
    Private mAudioRenderer As DirectShowLib.IBaseFilter

    'VMR9 RENDERER
    Private mVideoRenderer As DirectShowLib.IBaseFilter

    'VMR9 OBJECTS
    Private mWindowlessControl As DirectShowLib.IVMRWindowlessControl9
    Private mVMRSurfaceAlloc As DirectShowLib.IVMRSurfaceAllocator9
    Private mVMRSurfaceAllocNotify As DirectShowLib.IVMRSurfaceAllocatorNotify9
    Private mVMRStreamControl As DirectShowLib.IVMRVideoStreamControl9
    Private mVMRWindowlessControl As DirectShowLib.IVMRWindowlessControl9

    'FILENAME OF THE FILE TO BE RENDERED
    Private mFileName As String

    'WINDOW THAT WILL HOLD (IF THERE IS ANY) THE VIDEO TO BE PLAYED
    Private mVideoWindowHandle As IntPtr = IntPtr.Zero


#Region "WINAPI"
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True, ExactSpelling:=True)>
    Private Shared Function IsWindow(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
#End Region

#Region "PROPERTIES"
    Public Property VideoWindowWND() As IntPtr
        Get
            Return Me.mVideoWindowHandle
        End Get
        Set(value As IntPtr)
            If Not IsWindow(value) Then Throw New ArgumentException("the specified handle does not correspond to a vaid window")

            Me.mVideoWindowHandle = value
        End Set
    End Property
#End Region

    'CONSTRUCTORS

    Public Sub New(ByVal szFileName As String)
        If System.IO.File.Exists(szFileName) Then
            Me.mFileName = szFileName
        Else
            Throw New System.IO.FileNotFoundException("No se puede encontrar el archivo especificado", szFileName)
        End If
    End Sub

    Public Function CreateCustomFG() As Boolean
        'file source filter only output pin
        Dim FileSourceOutputPin As DirectShowLib.IPin = Nothing

        'lav splitter input pin and output audio and video pins
        Dim LAVSplitterInputPin As DirectShowLib.IPin = Nothing
        Dim LAVSplitterAudioPin As DirectShowLib.IPin = Nothing
        Dim LAVSplitterVideoPin As DirectShowLib.IPin = Nothing

        'lav audio decoder input and output pins
        Dim LAVAudioDecInputPin As DirectShowLib.IPin = Nothing
        Dim LAVAudioDecOutputPin As DirectShowLib.IPin = Nothing

        'lav video decoder input and output pins
        Dim LAVVideoDecInputPin As DirectShowLib.IPin = Nothing
        Dim LAVVideoDecOutputPin As DirectShowLib.IPin = Nothing

        'dsound renderer input pins
        Dim DSoundInputPin As DirectShowLib.IPin = Nothing

        'VMR9 input0 pin 
        Dim VMR9InputPin As DirectShowLib.IPin = Nothing

        Try


            If Not System.IO.File.Exists(Me.mFileName) Then Return False

            Me.GetFilterPointer("E436EBB5-524F-11CE-9F53-0020AF0BA770", Me.mSourceFilter)
            Me.GetFilterPointer("171252A0-8820-4AFE-9DF8-5C92B2D66B04", Me.mLAV_Splitter)

            Me.AddCustomFilter(Me.mSourceFilter, "File Source (Async.)", True, Me.mFileName)
            Me.AddCustomFilter(Me.mLAV_Splitter, "LAV Splitter", False)

            FileSourceOutputPin = Me.GetPin(Me.mSourceFilter, GET_PIN_TYPE.PT_NAME, "Output")
            LAVSplitterInputPin = Me.GetPin(Me.mLAV_Splitter, GET_PIN_TYPE.PT_NAME, "Input")

            If FileSourceOutputPin Is Nothing OrElse LAVSplitterInputPin Is Nothing Then
                GoTo SafeRelease
            End If

            Me.ConnectPins(FileSourceOutputPin, LAVSplitterInputPin, True)

            'LAV SPLITTER WILL CREATE VIDEO AND AUDIO OUTPUT PINS ONLY IF NECESARY

            LAVSplitterAudioPin = Me.GetPin(Me.mLAV_Splitter, GET_PIN_TYPE.PT_NAME, "Audio")

            If LAVSplitterAudioPin IsNot Nothing Then
                'stream has audio pin, it means lav detected and audio stream within the file
                Me.bHasAudio = True
                'instanciamos el decoder de audio de LAV
                Me.GetFilterPointer("E8E73B6B-4CB3-44A4-BE99-4F7BCB96E491", Me.mLAV_Audio)
                Me.AddCustomFilter(Me.mLAV_Audio, "LAV Audio Decoder", False)

                'and we bring both its pins (input and output to connect to the splitter and renderer
                LAVAudioDecInputPin = Me.GetPin(Me.mLAV_Audio, GET_PIN_TYPE.PT_NAME, "Input")
                LAVAudioDecOutputPin = Me.GetPin(Me.mLAV_Audio, GET_PIN_TYPE.PT_NAME, "Output")

                'connect the decoder's input pin to the output pin of the splitter
                Me.ConnectPins(LAVSplitterAudioPin, LAVAudioDecInputPin, False)

                'instantiate directsound renderer for audio rendering 
                Me.GetFilterPointer("79376820-07D0-11CF-A24D-0020AFD79767", Me.mAudioRenderer)

                'add the sound renderer to the filter graph
                Me.AddCustomFilter(Me.mAudioRenderer, "DirectSound Audio Renderer", False)

                'and we bring its input pin
                DSoundInputPin = Me.GetPin(Me.mAudioRenderer, GET_PIN_TYPE.PT_NAME, "Audio Input pin (rendered)")

                'connect the audio decoder output pin to the renderer input pin
                Me.ConnectPins(LAVAudioDecOutputPin, DSoundInputPin, True)

            End If

            LAVSplitterVideoPin = Me.GetPin(Me.mLAV_Splitter, GET_PIN_TYPE.PT_NAME, "Video")

            If LAVSplitterVideoPin IsNot Nothing Then
                'LAV Splitter only exposes Video pin if the stream has video

                'we instantiate the video decoder
                Me.GetFilterPointer("EE30215D-164F-4A92-A4EB-9D4C13390F9F", Me.mLAV_Video)

                'we bring the video decoder input and output pins
                LAVVideoDecInputPin = Me.GetPin(Me.mLAV_Video, GET_PIN_TYPE.PT_NAME, "Input")
                LAVVideoDecOutputPin = Me.GetPin(Me.mLAV_Video, GET_PIN_TYPE.PT_NAME, "Output")

                'we connect the splitter video output pin to the decoder input pin
                Me.ConnectPins(LAVSplitterVideoPin, LAVVideoDecInputPin, True)

                'we instantiate the Video Mixing Renderer 9 codec
                Me.GetFilterPointer("51B4ABF3-748F-4E3B-A276-C828330E926A", Me.mVideoRenderer)

                'configure VMR9 renderer properties before connecting it
                If Not Me.ConfigureVMR9Windowless() Then GoTo SafeRelease

                'we get the vmr9 first input pin
                VMR9InputPin = Me.GetPin(Me.mVideoRenderer, GET_PIN_TYPE.PT_NAME, "VMR Input0")

                'and we finally connect them
                Me.ConnectPins(LAVVideoDecOutputPin, VMR9InputPin, True)
            End If

            'now the filtergraph should be completely assembled and connected, when we call the Run() method, it should play
            Return True

        Catch ex As Exception

            Return False

        Finally
            'release all the pins
            If Not FileSourceOutputPin Is Nothing Then Marshal.ReleaseComObject(FileSourceOutputPin)
            If Not LAVSplitterInputPin Is Nothing Then Marshal.ReleaseComObject(LAVSplitterInputPin)
            If Not LAVSplitterAudioPin Is Nothing Then Marshal.ReleaseComObject(LAVSplitterAudioPin)
            If Not LAVSplitterVideoPin Is Nothing Then Marshal.ReleaseComObject(LAVSplitterVideoPin)
            If Not LAVAudioDecInputPin Is Nothing Then Marshal.ReleaseComObject(LAVAudioDecInputPin)
            If Not LAVAudioDecOutputPin Is Nothing Then Marshal.ReleaseComObject(LAVAudioDecOutputPin)
            If Not DSoundInputPin Is Nothing Then Marshal.ReleaseComObject(DSoundInputPin)
            If Not LAVVideoDecInputPin Is Nothing Then Marshal.ReleaseComObject(LAVVideoDecInputPin)
            If Not LAVVideoDecOutputPin Is Nothing Then Marshal.ReleaseComObject(LAVVideoDecOutputPin)
            If Not VMR9InputPin Is Nothing Then Marshal.ReleaseComObject(VMR9InputPin)
        End Try


SafeRelease:

    End Function


#Region "METODOS_PRIVADOS"
    Private Function GetFilterPointer(ByVal filterCLSiD As String, ByRef Filter As DirectShowLib.IBaseFilter) As Boolean
        Try
            Dim _guid As New Guid(filterCLSiD)

            'CREAMOS INSTANCIA DE OBJETO COM Y LA AGREGAMOS AL REFERENCE COUNTER
            Dim _type As Type = Type.GetTypeFromCLSID(_guid)

            'si el type devuelto no es un objecto COM, entonces no es un filtro
            If Not _type.IsCOMObject Then
                Return False
            End If

            Filter = DirectCast(Activator.CreateInstance(_type), DirectShowLib.IBaseFilter)

            Return True
        Catch ex As Exception
            'RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
            'Me.ReleaseCustomFG()
            Return False
        End Try
    End Function

    Private Function AddCustomFilter(ByVal iFilter As DirectShowLib.IBaseFilter, ByVal FriendlyName As String, ByVal bIsSource As Boolean, Optional ByVal szFile As String = "") As Boolean
        If iFilter Is Nothing Then Return False

        If bIsSource Then
            DirectShowLib.DsError.ThrowExceptionForHR(Me.mFG.AddSourceFilter(mFileName, FriendlyName, iFilter))
        Else
            DirectShowLib.DsError.ThrowExceptionForHR(Me.mFG.AddFilter(iFilter, FriendlyName))
        End If

        Return True
    End Function

    Private Function GetPin(ByVal filter As DirectShowLib.IBaseFilter, ByVal searchType As GET_PIN_TYPE, ByVal searchCriteria As Object) As DirectShowLib.IPin
        If filter Is Nothing Then Throw New ArgumentNullException(NameOf(filter), "The filter must be specified.")

        Dim pinArray(0) As DirectShowLib.IPin
        Dim pinEnumerator As DirectShowLib.IEnumPins = Nothing
        Dim pinInfo As DirectShowLib.PinInfo = Nothing

        ' Enumerate pins
        DirectShowLib.DsError.ThrowExceptionForHR(filter.EnumPins(pinEnumerator))

        Try
            While pinEnumerator.Next(1, pinArray, IntPtr.Zero) = 0 ' S_OK
                Dim currentPin As DirectShowLib.IPin = pinArray(0)

                Try
                    ' Check pin based on the criteria
                    If TypeOf searchCriteria Is String Then
                        currentPin.QueryPinInfo(pinInfo)

                        ' If the pin matches the search criteria, return it to the caller
                        If pinInfo.name = CType(searchCriteria, String) Then
                            Return currentPin
                        End If
                    End If

                Catch ex As Exception
                    ' If QueryPinInfo fails, release the pin immediately to avoid memory leaks
                    Marshal.ReleaseComObject(currentPin)
                    ' Optionally log the exception here
                Finally
                    ' Always release pinInfo.filter to avoid memory leaks
                    If pinInfo.filter IsNot Nothing Then
                        Marshal.FinalReleaseComObject(pinInfo.filter)
                        pinInfo.filter = Nothing
                    End If
                End Try

                ' If pin is not needed, release it here to avoid memory leaks
                Marshal.ReleaseComObject(currentPin)

            End While

            ' Pin not found
            Return Nothing

        Catch ex As Exception
            ' Handle or log the exception if needed
            Return Nothing

        Finally
            ' Clean up the enumerator if needed
            If pinEnumerator IsNot Nothing Then Marshal.FinalReleaseComObject(pinEnumerator)
        End Try
    End Function

    Private Function ConnectPins(ByVal outputPin As DirectShowLib.IPin, ByVal inputPin As DirectShowLib.IPin, ByVal bUseIC As Boolean) As Boolean

        If outputPin Is Nothing Then Throw New ArgumentNullException(NameOf(outputPin), "The output pin must be specified.")
        If inputPin Is Nothing Then Throw New ArgumentNullException(NameOf(inputPin), "The input pin must be specified.")

        'query the pin direction to ensure they are ok
        Dim outputPinDirection As DirectShowLib.PinDirection
        Dim inputPinDirection As DirectShowLib.PinDirection

        Try
            DirectShowLib.DsError.ThrowExceptionForHR(outputPin.QueryDirection(outputPinDirection))
            DirectShowLib.DsError.ThrowExceptionForHR(inputPin.QueryDirection(inputPinDirection))

            If (Not outputPinDirection = DirectShowLib.PinDirection.Output) OrElse (Not inputPinDirection = DirectShowLib.PinDirection.Input) Then
                Return False
            End If

            'try to connect both pins using ic (intelligent connect) or dc (direct connection)
            Try
                DirectShowLib.DsError.ThrowExceptionForHR(Me.mFG.ConnectDirect(outputPin, inputPin, Nothing))
            Catch exDirectConnect As Exception
                'direct connection failed, try to use intelligent connect
                Try
                    'if the dev wants to try ic then we try it, otherwise we return with the error
                    If bUseIC Then
                        DirectShowLib.DsError.ThrowExceptionForHR(Me.mFG.Connect(outputPin, inputPin))
                    Else
                        'direct connect failed, and the user doesn't want to use ic, so we return
                        Return False
                    End If

                Catch exIntelligentConnect As Exception
                    'man connection failed, so did the ic connection, return false
                    Return False
                End Try
            End Try

            'wether it was intelligent connect or direct connect we succeded
            Return True
        Catch ex As Exception
            'fail, we return
            Return False
        End Try
    End Function

    Private Function ConfigureVMR9Windowless() As Boolean
        If Me.mVideoRenderer Is Nothing Then Return False

        Dim mVMRConfig As DirectShowLib.IVMRFilterConfig9 = Nothing

        Try

            mVMRConfig = DirectCast(Me.mVideoRenderer, DirectShowLib.IVMRFilterConfig9)

            mVMRConfig.SetNumberOfStreams(1)
            mVMRConfig.SetRenderingMode(DirectShowLib.VMR9Mode.Windowless)

            Me.mVMRWindowlessControl = DirectCast(Me.mVideoRenderer, DirectShowLib.IVMRWindowlessControl9)

            Me.mVMRWindowlessControl.SetVideoClippingWindow(Me.mVideoWindowHandle)
            Me.mVMRWindowlessControl.SetAspectRatioMode(DirectShowLib.VMR9AspectRatioMode.None)
            Return True
        Catch ex As Exception
            If Not Me.mVMRWindowlessControl Is Nothing Then Marshal.ReleaseComObject(Me.mVMRWindowlessControl)
            Return False

        Finally
            If Not mVMRConfig Is Nothing Then Marshal.ReleaseComObject(mVMRConfig)
        End Try


    End Function

    Private Sub SetFiltersToNothing()
        If Not Me.mSourceFilter Is Nothing Then Marshal.ReleaseComObject(Me.mSourceFilter)
        If Not Me.mLAV_Splitter Is Nothing Then Marshal.ReleaseComObject(Me.mLAV_Splitter)
        If Not Me.mLAV_Audio Is Nothing Then Marshal.ReleaseComObject(Me.mLAV_Audio)
        If Not Me.mLAV_Video Is Nothing Then Marshal.ReleaseComObject(Me.mLAV_Video)
        If Not Me.mAudioRenderer Is Nothing Then Marshal.ReleaseComObject(Me.mAudioRenderer)
        If Not Me.mVideoRenderer Is Nothing Then Marshal.ReleaseComObject(Me.mVideoRenderer)
    End Sub
#End Region

End Class
