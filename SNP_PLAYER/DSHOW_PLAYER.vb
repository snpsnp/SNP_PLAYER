Imports System.Runtime.InteropServices
Imports SNP_PLAYER.ENUMS
Imports DirectShowLib

Public Class DSHOW_PLAYER
    Implements IDisposable

    Private mFG As DirectShowLib.IFilterGraph2

    Private bHasAudio As Boolean = False
    Private bHasVideo As Boolean = False
    Private bRenderUsingIC As Boolean = False

    'MEDIA MANIPULATION INTERFACES, TO BE RETRIEVED FROM THE FILTERGRAPH
    Private mMedia As DirectShowLib.IMediaControl
    Private mSound As DirectShowLib.IBasicAudio
    Private mVideo As DirectShowLib.IBasicVideo
    Private mEvent As DirectShowLib.IMediaEventEx
    Private mSeek As DirectShowLib.IMediaSeeking
    Private mROTEntry As DirectShowLib.DsROTEntry

    'SOURCE FILTER OBJECT
    Private mSourceFilter As DirectShowLib.IBaseFilter

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

    'FIELD TO HOLD THE CURRENT TIME FORMAT 
    Private mTimeFormat As ENUMS.TIME_FORMAT

    'FIELD TO HOLD CURRENT STREAM STATE
    Private mState As ENUMS.ESTADO

#Region "WINAPI"
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True, ExactSpelling:=True)>
    Private Shared Function IsWindow(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("ole32.dll")>
    Private Shared Function GetRunningObjectTable(ByVal reserved As UInteger, ByRef pprot As UCOMIRunningObjectTable) As Integer
    End Function

    <DllImport("ole32.dll")>
    Private Shared Function CreateItemMoniker(ByVal lpszDelim As String, ByVal lpszItem As String, ByRef ppmk As UCOMIMoniker) As Integer
    End Function
#End Region

#Region "PROPERTIES"
    Public Property VideoWindowWND() As IntPtr
        Get
            Return Me.mVideoWindowHandle
        End Get
        Set(value As IntPtr)
            If Not IsWindow(value) Then Throw New ArgumentException("the specified handle does not correspond to a valid window")

            Me.mVideoWindowHandle = value
        End Set
    End Property

    Public Property Volume() As Long
        Get
            If Me.mSound Is Nothing Then
                Throw New InvalidOperationException("IBasicAudio is not available.")
            End If

            ' Retrieve the volume in dB from DirectShow
            Dim dBVolume As Integer
            Me.mSound.get_Volume(dBVolume)

            ' Convert dBVolume (-10000 to 0) back to linear scale (0 to 100)
            Dim mVolume As Integer
            If dBVolume <= -10000 Then
                ' Completely silent
                mVolume = 0
            Else
                ' Inverse of logarithmic scaling
                mVolume = CInt(Math.Pow(10, dBVolume / 2000.0) * 100)
            End If

            Return mVolume
        End Get

        Set(value As Long)
            If Me.mSound Is Nothing Then
                Throw New InvalidOperationException("IBasicAudio is not available.")
            End If

            If value < 0 OrElse value > 100 Then
                Throw New ArgumentOutOfRangeException(NameOf(value), "Volume must be between 0 and 100.")
            End If

            ' Convert linear slider value to logarithmic dB scale
            Dim scaledVolume As Integer
            If value = 0 Then
                ' Completely silent
                scaledVolume = -10000
            Else
                ' Logarithmic scaling: Maps 1-100 to -10000 to 0 dB
                scaledVolume = CInt(2000.0 * Math.Log10(value / 100.0))
            End If

            Me.mSound.put_Volume(scaledVolume)
        End Set
    End Property

    Public Property Balance() As Integer
        Get
            If Me.mSound Is Nothing Then Throw New InvalidOperationException("IBasicAudio is not available")

            Dim mBal As Integer
            Me.mSound.get_Balance(mBal)
            Return mBal / 100

        End Get
        Set(value As Integer)
            If Me.mSound Is Nothing Then Throw New InvalidOperationException("IBasicAudio is not available")

            ' Validate the input value to ensure it is within the range [-100, 100]
            If value < -100 OrElse value > 100 Then
                Throw New ArgumentOutOfRangeException(NameOf(value), "Balance must be between -100 and 100.")
            End If

            Me.mSound.put_Balance(value * 100)
        End Set
    End Property

    Public Property TimeFormat() As ENUMS.TIME_FORMAT
        Get
            Return Me.mTimeFormat
        End Get
        Set(value As ENUMS.TIME_FORMAT)
            If Me.mSeek Is Nothing Then Throw New InvalidOperationException("IMediaSeeking is not set")
            Select Case value
                Case ENUMS.TIME_FORMAT.TF_FRAMES
                    If Me.mSeek.IsFormatSupported(DirectShowLib.TimeFormat.Frame) Then
                        Me.mSeek.SetTimeFormat(DirectShowLib.TimeFormat.Frame)
                    Else
                        Throw New ArgumentException("Time Format not supported")
                    End If
                Case TIME_FORMAT.TF_MEDIATIME
                Case TIME_FORMAT.TF_MILLISECONDS
                Case TIME_FORMAT.TF_SECONDS
                    If Me.mSeek.IsFormatSupported(DirectShowLib.TimeFormat.MediaTime) Then
                        Me.mSeek.SetTimeFormat(DirectShowLib.TimeFormat.MediaTime)
                    Else
                        Throw New ArgumentException("Time Format not supported")
                    End If
                Case TIME_FORMAT.TF_SAMPLES
                    If Me.mSeek.IsFormatSupported(DirectShowLib.TimeFormat.Sample) Then
                        Me.mSeek.SetTimeFormat(DirectShowLib.TimeFormat.Sample)
                    Else
                        Throw New ArgumentException("Time Format not supported")
                    End If
            End Select
            Me.mTimeFormat = value
        End Set
    End Property

    Public ReadOnly Property Duration() As Long
        Get
            If Me.mSeek Is Nothing Then Throw New InvalidOperationException("IMediaSeeking is not set")

            Dim mDuration As Long
            Me.mSeek.GetDuration(mDuration)

            If Me.mTimeFormat = TIME_FORMAT.TF_MILLISECONDS Then
                mDuration /= 10000
            End If

            If Me.mTimeFormat = TIME_FORMAT.TF_SECONDS Then
                mDuration /= 10000 / 1000
            End If

            Return mDuration
        End Get
    End Property

    Public Property Position() As Long
        Get
            Dim mPosition As Long, mStop As Long
            If Me.mSeek Is Nothing Then Throw New InvalidOperationException("IMediaSeeking is not set")

            Me.mSeek.GetPositions(mPosition, mStop)

            If Me.mTimeFormat = TIME_FORMAT.TF_MILLISECONDS Then mPosition /= 10000
            If Me.mTimeFormat = TIME_FORMAT.TF_SECONDS Then mPosition /= 10000 / 1000

            Return mPosition
        End Get
        Set(value As Long)
            If Me.mSeek Is Nothing Then Throw New InvalidOperationException("IMediaSeeking is not set")

            Dim mPosition As Long

            If mTimeFormat = TIME_FORMAT.TF_MILLISECONDS Then
                mPosition = value * 10000
            ElseIf mTimeFormat = TIME_FORMAT.TF_SECONDS Then
                mPosition = value * 10000 * 1000
            Else
                mPosition = value
            End If

            Me.mSeek.SetPositions(mPosition, DirectShowLib.AMSeekingSeekingFlags.AbsolutePositioning, 0, DirectShowLib.AMSeekingSeekingFlags.NoPositioning)

        End Set
    End Property

    Public Property UseIntelligentConnect() As Boolean
        Get
            Return Me.bRenderUsingIC
        End Get
        Set(value As Boolean)
            Me.bRenderUsingIC = value
        End Set
    End Property

    Public Property FileName() As String
        Get
            Return Me.mFileName
        End Get
        Set(value As String)
            If Not System.IO.File.Exists(value) Then Throw New ArgumentException("the specified file can not be accessed")
            Me.mFileName = value
        End Set
    End Property
#End Region

    'CONSTRUCTORS

    Public Sub New(ByVal szFileName As String)
        Me.SetFiltersToNothing()
        Me.SetInterfacesToNothing()
        Me.SetVMRInterfacesToNothing()

        If System.IO.File.Exists(szFileName) Then
            Me.mFileName = szFileName
            Me.Open()
        Else
            Throw New System.IO.FileNotFoundException("No se puede encontrar el archivo especificado", szFileName)
        End If
    End Sub

    Public Sub New()
        Me.SetFiltersToNothing()
        Me.SetInterfacesToNothing()
        Me.SetVMRInterfacesToNothing()
    End Sub

    Public Sub Play()
        If Me.mMedia Is Nothing Then Throw New InvalidOperationException("IMediaControl is not set")
        Me.mMedia.Run()
        Me.mState = ESTADO.REPRODUCIENDO
    End Sub

    Public Sub Pause()
        If Me.mMedia Is Nothing Then Throw New InvalidOperationException("IMediaControl is not set")
        Me.mMedia.Pause()
        Me.mState = ESTADO.PAUSADO
    End Sub

    Public Sub [Stop]()
        If Me.mMedia Is Nothing Then Throw New InvalidOperationException("IMediaControl is not set")
        Me.mMedia.Stop()
        Position = 0
        Me.mState = ESTADO.DETENIDO
    End Sub

    Public Function Open() As Boolean

        'try to open a file creating a custom filter graph 
        If Me.CreateCustomFG() Then

            Me.mROTEntry = New DsROTEntry(Me.mFG)

            'graph creation succeded, we query all the necessary interfaces
            Me.mMedia = DirectCast(Me.mFG, DirectShowLib.IMediaControl)
            If bHasAudio Then Me.mSound = DirectCast(Me.mFG, DirectShowLib.IBasicAudio)
            If bHasVideo Then Me.mVideo = DirectCast(Me.mFG, DirectShowLib.IBasicVideo)
            Me.mEvent = DirectCast(Me.mFG, DirectShowLib.IMediaEventEx)
            Me.mSeek = DirectCast(Me.mFG, DirectShowLib.IMediaSeeking)
            'set the stream state to open
            Me.mState = ESTADO.ABIERTO
            'return succesfully
            Return True

        Else
            'failed to open the stream, release and re-set all the interfaces and filter objects

            Me.SafeReleaseVMRInterfaces()
            Me.SafeReleaseInterfaces()
            Me.SafeReleaseFilters()
            Me.SetInterfacesToNothing()
            Me.SetVMRInterfacesToNothing()
            Me.SetFiltersToNothing()
            'return error!
            Return False
        End If

    End Function

    Public Sub Close()
        Me.Stop()
        Me.SafeReleaseVMRInterfaces()
        Me.SafeReleaseInterfaces()
        Me.SafeReleaseFilters()
        Me.SetInterfacesToNothing()
        Me.SetVMRInterfacesToNothing()
        Me.SetFiltersToNothing()
    End Sub
#Region "METODOS_PRIVADOS"

    Private Function CreateCustomFG() As Boolean
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

            Me.mFG = DirectCast(New DirectShowLib.FilterGraph, DirectShowLib.IFilterGraph2)

            Me.GetFilterPointer("E436EBB5-524F-11CE-9F53-0020AF0BA770", Me.mSourceFilter)

            Me.GetFilterPointer("171252A0-8820-4AFE-9DF8-5C92B2D66B04", Me.mLAV_Splitter)

            Me.AddCustomFilter(Me.mSourceFilter, "File Source (Async.)", True, Me.mFileName)

            Me.AddCustomFilter(Me.mLAV_Splitter, "LAV Splitter", False)

            Me.mSourceFilter.FindPin("Output", FileSourceOutputPin)

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
            'error, limpiamos todo
            Me.SafeReleaseFilters()
            Me.SetFiltersToNothing()
            Me.SafeReleaseVMRInterfaces()
            Me.SetVMRInterfacesToNothing()
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
        Me.SafeReleaseFilters()
        Me.SetFiltersToNothing()
        Me.SafeReleaseVMRInterfaces()
        Me.SetVMRInterfacesToNothing()
        Return False
    End Function

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

    Private Function AddCustomFilter(ByRef iFilter As DirectShowLib.IBaseFilter, ByVal FriendlyName As String, ByVal bIsSource As Boolean, Optional ByVal szFile As String = "") As Boolean
        If iFilter Is Nothing Then Return False

        If bIsSource Then
            DirectShowLib.DsError.ThrowExceptionForHR(Me.mFG.AddSourceFilter(szFile, FriendlyName, iFilter))
        Else
            DirectShowLib.DsError.ThrowExceptionForHR(Me.mFG.AddFilter(iFilter, FriendlyName))
        End If

        Return True
    End Function

    Private Function GetPin(ByRef filter As DirectShowLib.IBaseFilter, ByVal searchType As GET_PIN_TYPE, ByVal searchCriteria As Object) As DirectShowLib.IPin
        If filter Is Nothing Then Throw New ArgumentNullException(NameOf(filter), "The filter must be specified.")

        Dim pinArray(0) As DirectShowLib.IPin
        Dim pinEnumerator As DirectShowLib.IEnumPins = Nothing
        Dim pinInfo As DirectShowLib.PinInfo = Nothing

        ' Enumerate pins
        DirectShowLib.DsError.ThrowExceptionForHR(filter.EnumPins(pinEnumerator))

        Try
            Dim pinIndex As Integer = 0

            While pinEnumerator.Next(1, pinArray, IntPtr.Zero) = 0 ' S_OK
                'Dim currentPin As DirectShowLib.IPin = pinArray(0)

                Try
                    DirectShowLib.DsError.ThrowExceptionForHR(pinArray(0).QueryPinInfo(pinInfo))

                    Select Case searchType
                        Case GET_PIN_TYPE.PT_NAME
                            ' Check pin based on name
                            If pinInfo.name = CType(searchCriteria, String) Then
                                Return pinArray(0)
                            End If

                        Case GET_PIN_TYPE.PT_NUMBER
                            ' Check pin based on number
                            If pinIndex = CType(searchCriteria, Integer) Then
                                Return pinArray(0)
                            End If
                            pinIndex += 1

                    End Select

                Catch ex As Exception
                    ' If QueryPinInfo fails, release the pin immediately to avoid memory leaks
                    Marshal.ReleaseComObject(pinArray(0))
                    ' Optionally log the exception here
                Finally
                    ' Always release pinInfo.filter to avoid memory leaks
                    'If pinInfo Is Nothing Then
                    DsUtils.FreePinInfo(pinInfo)
                    '    pinInfo = Nothing
                    'End If
                End Try

                ' If pin is not needed, release it here to avoid memory leaks
                Marshal.ReleaseComObject(pinArray(0))

            End While

            ' Pin not found
            Return Nothing

        Catch ex As Exception
            ' Handle or log the exception if needed
            Return Nothing

        Finally
            ' Clean up the enumerator if needed
            If Not pinEnumerator Is Nothing Then Marshal.FinalReleaseComObject(pinEnumerator)
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

    Private Sub SafeReleaseInterfaces()
        If Not Me.mMedia Is Nothing Then Marshal.ReleaseComObject(Me.mMedia)
        If Not Me.mSound Is Nothing Then Marshal.ReleaseComObject(Me.mSound)
        If Not Me.mVideo Is Nothing Then Marshal.ReleaseComObject(Me.mVideo)
        If Not Me.mEvent Is Nothing Then Marshal.ReleaseComObject(Me.mEvent)
        If Not Me.mSeek Is Nothing Then Marshal.ReleaseComObject(Me.mSeek)
    End Sub

    Private Sub SetInterfacesToNothing()
        Me.mMedia = Nothing
        Me.mSound = Nothing
        Me.mVideo = Nothing
        Me.mEvent = Nothing
        Me.mSeek = Nothing
    End Sub

    Private Sub SafeReleaseVMRInterfaces()
        If Not Me.mWindowlessControl Is Nothing Then Marshal.ReleaseComObject(Me.mWindowlessControl)
        If Not Me.mVMRStreamControl Is Nothing Then Marshal.ReleaseComObject(Me.mVMRStreamControl)
        If Not Me.mVMRSurfaceAlloc Is Nothing Then Marshal.ReleaseComObject(Me.mVMRSurfaceAlloc)
        If Not Me.mVMRSurfaceAllocNotify Is Nothing Then Marshal.ReleaseComObject(Me.mVMRSurfaceAllocNotify)
    End Sub

    Private Sub SetVMRInterfacesToNothing()
        Me.mWindowlessControl = Nothing
        Me.mVMRStreamControl = Nothing
        Me.mVMRSurfaceAlloc = Nothing
        Me.mVMRSurfaceAllocNotify = Nothing
    End Sub

    Private Sub SafeReleaseFilters()
        If Me.mROTEntry IsNot Nothing Then Me.mROTEntry.Dispose()

        If Not Me.mSourceFilter Is Nothing Then Marshal.ReleaseComObject(Me.mSourceFilter)
        If Not Me.mLAV_Splitter Is Nothing Then Marshal.ReleaseComObject(Me.mLAV_Splitter)
        If Not Me.mLAV_Audio Is Nothing Then Marshal.ReleaseComObject(Me.mLAV_Audio)
        If Not Me.mLAV_Video Is Nothing Then Marshal.ReleaseComObject(Me.mLAV_Video)
        If Not Me.mAudioRenderer Is Nothing Then Marshal.ReleaseComObject(Me.mAudioRenderer)
        If Not Me.mVideoRenderer Is Nothing Then Marshal.ReleaseComObject(Me.mVideoRenderer)
        If Not Me.mROTEntry Is Nothing Then Me.mROTEntry.Dispose()

        If Not Me.mFG Is Nothing Then Marshal.ReleaseComObject(Me.mFG)
    End Sub

    Private Sub SetFiltersToNothing()
        Me.mSourceFilter = Nothing
        Me.mLAV_Splitter = Nothing
        Me.mLAV_Audio = Nothing
        Me.mLAV_Video = Nothing
        Me.mROTEntry = Nothing
        Me.mAudioRenderer = Nothing
        Me.mVideoRenderer = Nothing
        mFG = Nothing
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            Me.Close()
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
#End Region

End Class
