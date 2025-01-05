Imports System.Runtime.InteropServices

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
    Private mVMRConfig As DirectShowLib.IVMRFilterConfig9
    Private mVMRSurfaceAlloc As DirectShowLib.IVMRSurfaceAllocator9
    Private mVMRSurfaceAllocNotify As DirectShowLib.IVMRSurfaceAllocatorNotify9
    Private mVMRStreamControl As DirectShowLib.IVMRVideoStreamControl9

    'FILENAME OF THE FILE TO BE RENDERED
    Private mFileName As String

    'CONSTRUCTORS

    Public Sub New(ByVal szFileName As String)
        If System.IO.File.Exists(szFileName) Then
            Me.mFileName = szFileName
        Else
            Throw New System.IO.FileNotFoundException("No se puede encontrar el archivo especificado", szFileName)
        End If
    End Sub


#Region "METODOS_PRIVADOS"
    Private Function GetFilterPointer(ByVal filterCLSiD As String, ByVal filterFriendlyName As String, ByRef Filter As DirectShowLib.IBaseFilter) As Boolean
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


#End Region

End Class
