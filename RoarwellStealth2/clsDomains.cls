VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDomains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// This class will provide the functionality
'// to enumerate all domains in your network.

'//To begin the enumeration of a network container resource,
'//your application should perform the following steps:

'//1.)Pass the address of a NETRESOURCE structure that represents _
the resource to the WNetOpenEnum function.

'//2.)Allocate a buffer large enough to hold the array of NETRESOURCE _
structures that the WNetEnumResource function returns, _
plus the strings to which their members point.

'//3.)Pass the resource handle returned by WNetOpenEnum to the _
WNetEnumResource function.

'//4.)Close the resource handle when it is no longer needed by _
calling the WNetCloseEnum function.

Private cDomains As New Collection

Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" _
(ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, _
lpNetResource As Any, lppEnumHwnd As Long) As Long
''''''''''''''''''''''''''''''PARAMETERS-WNetOpenEnum(1,2,3,4,5)''''''''''''''''''''''
'1'''dwScope'''''=[in] Specifies the scope of the enumeration. _
This parameter can be one of the following values: _
RESOURCE_CONNECTED,RESOURCE_GLOBALNET,RESOURCE_REMEMBERED

'2'''dwType''''''=[in] Specifies the resource types to enumerate. _
This parameter can be a combination of the following values: _
RESOURCETYPE_ANY,RESOURCETYPE_DISK,RESOURCETYPE_PRINT

'3'''dwUsage'''''=[in] Specifies the resource usage to be enumerated. _
This parameter can be a combination of the following values: _
RESOURCEUSAGE_CONNECTABLE,RESOURCEUSAGE_CONTAINER

'4'lpNetResource'=[in] Long pointer to a NETRESOURCE structure _
specifying the container to enumerate.If the dwScope parameter _
is not RESOURCE_GLOBALNET, this parameter must be NULL.

'5''''lphEnum''''=[out] Long pointer to a variable filled with _
an enumeration handle that can be used in a subsequent call to WNetEnumResource.


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'NOTE:-To enumerate all network resources, an application can begin _
the enumeration by calling WNetOpenEnum with lpNetResource _
pointing to a NETRESOURCE structure specifying the root of _
the network, then using the returned handle with WNetEnumResource _
to enumerate resources. If one of the resources in the NETRESOURCE _
array returned by the WNetEnumResource function is a container _
resource, WNetOpenEnum can be used to open the resource for further enumeration.



Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal pEnumHwnd As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal p_lngEnumHwnd As Long) As Long
Private Declare Function StrLenA Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Private Declare Function StrCopyA Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long

'RESOURCE ENUMERATION
Private Const RESOURCE_CONNECTED As Long = &H1 'All currently connected _
resources (the dwUsage parameter is ignored).
Private Const RESOURCE_GLOBALNET As Long = &H2& 'All resources on the network.
Private Const RESOURCE_REMEMBERED As Long = &H3 'All remembered (persistent) _
connections (the dwUsage parameter is ignored). These connections may or may not currently be connected.

Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCETYPE_DISK As Long = &H1
Private Const RESOURCETYPE_PRINT As Long = &H2
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF

Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1
Private Const RESOURCEUSAGE_CONTAINER As Long = &H2
Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000

Private Const RESOURCE_ENUM_ALL As Long = &HFFFF

Private Const NO_ERROR As Long = 0&
Private Const MAX_RESOURCES As Long = 256

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type

Public Function GetCount() As Long
    '// Return the number of domains in
    '// the collection.
    GetCount = cDomains.Count
End Function


Public Function GetItem(Index As Long) As String
    '// Check for legal value
    If Index < 1 Or Index > cDomains.Count Then
        '// Raise an error
        Err.Raise 7771, , "The index is not within valid range!"
        Exit Function
    End If
    
    '// Return an item in the domains collection
    GetItem = cDomains.Item(Index)
End Function


Public Sub Refresh()
    Dim tNetAPI(0 To MAX_RESOURCES) As NETRESOURCE
    Dim lBufferSize As Long
    Dim lEnumHwnd As Long
    Dim lReturn As Long
    Dim lCount As Long
    Dim lLoop As Long
    
    '// Clear the domains collection
    Do While cDomains.Count > 0
        cDomains.Remove 1
    Loop
    
    '// Refresh the domains collection
    lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, tNetAPI(0), lEnumHwnd)
    
    If lReturn = NO_ERROR Then
        lCount = RESOURCE_ENUM_ALL
        
        lBufferSize = UBound(tNetAPI) * Len(tNetAPI(0))
        lReturn = WNetEnumResource(lEnumHwnd, lCount, tNetAPI(0), lBufferSize)
        
        If lCount > 0 Then
            For lLoop = 0 To lCount - 1
                '// Add domain to collection
                cDomains.Add PointerToAsciiStr(tNetAPI(lLoop).pRemoteName)
            Next lLoop
        End If
    End If
    
    If lEnumHwnd <> 0 Then
        Call WNetCloseEnum(lEnumHwnd)
    End If
End Sub

Private Function PointerToAsciiStr(ByVal lPtrToString As Long) As String
    On Local Error Resume Next
    Dim lLength As Long
    Dim sStringValue As String
    Dim lNullPos As Long
    Dim lReturn As Long
    
    lLength = StrLenA(lPtrToString)
    
    If lPtrToString > 0 And lLength > 0 Then
        '// Copy Pointer to String
        sStringValue = Space$(lLength + 1)
        lReturn = StrCopyA(sStringValue, lPtrToString)
        
        '// Find null terminator
        lNullPos = InStr(1, sStringValue, Chr$(0))
        
        If lNullPos > 0 Then
            '// Lose the null terminator
            PointerToAsciiStr = Left$(sStringValue, lNullPos - 1)
        Else
            '// Just pass the string
            PointerToAsciiStr = sStringValue
        End If
    Else
        PointerToAsciiStr = ""
    End If
End Function



Private Sub Class_Initialize()
    '// Refresh the domains
    Call Refresh
End Sub


Private Sub Class_Terminate()
    '// Clear the collection
    Do While cDomains.Count > 0
        cDomains.Remove 1
    Loop
End Sub



