Imports System.Runtime.InteropServices

Namespace ENUMS
    Public Enum ESTADO
        ABIERTO
        CERRADO
        REPRODUCIENDO
        PAUSADO
        DETENIDO
        SEEKING
    End Enum

    Public Enum TIME_FORMAT
        TF_NONE
        TF_MILLISECONDS
        TF_FRAMES
        TF_SECONDS
        TF_MEDIATIME
        TF_SAMPLES
    End Enum

    Public Enum GET_PIN_TYPE
        PT_NAME
        PT_NUMBER
    End Enum

    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000010-0000-0000-C000-000000000046")>
    Public Interface IRunningObjectTable
        Sub Register(ByVal grfFlags As Integer, <[In]()> ByVal punkObject As Object, <[In]()> ByVal pmkObjectName As IMoniker, ByRef pdwRegister As Integer)
        Sub Revoke(ByVal dwRegister As Integer)
        Sub IsRunning(<[In]()> ByVal pmkObjectName As IMoniker)
        Sub GetObject(<[In]()> ByVal pmkObjectName As IMoniker, <MarshalAs(UnmanagedType.[Interface])> ByRef ppunkObject As Object)
        Sub NoteChangeTime(ByVal dwRegister As Integer, <[In]()> ByRef pfiletime As FILETIME)
        Sub GetTimeOfLastChange(<[In]()> ByVal pmkObjectName As IMoniker, <Out()> ByRef pfiletime As FILETIME)
        Sub EnumRunning(<MarshalAs(UnmanagedType.[Interface])> ByRef ppenumMoniker As IEnumMoniker)
    End Interface

    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000f-0000-0000-C000-000000000046")>
    Public Interface IMoniker
        ' Methods of the IMoniker interface
    End Interface

    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000102-0000-0000-C000-000000000046")>
    Public Interface IEnumMoniker
        Sub [Next](ByVal celt As Integer, <Out(), MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal rgelt() As IMoniker, ByRef pceltFetched As Integer)
        Sub Skip(ByVal celt As Integer)
        Sub Reset()
        Sub Clone(ByRef ppenum As IEnumMoniker)
    End Interface

End Namespace
