Attribute VB_Name = "modLan"
Option Explicit

Public Function getLanInfo() As String
    Dim cComputers As New clsComputers
    Dim cDomains As New clsDomains
    Dim domCount As Long
    Dim compCount As Long
    Dim compNames As String
    '// Enumerate Domains
    cDomains.Refresh
    'MsgBox cDomains.GetCount
    For domCount = 1 To cDomains.GetCount
        ''get the domain name
        cComputers.Domain = cDomains.GetItem(domCount)
        ''get the computers in the domain
        'MsgBox cComputers.Domain
        cComputers.Refresh
        For compCount = 1 To cComputers.GetCount
            If Not cComputers.GetItem(compCount) = "" Then
                compNames = compNames & cComputers.GetItem(compCount) & "|"
                'MsgBox compNames
            Else
                'do something if there are no computers in this domain
                Exit For
            End If
        Next
        'MsgBox "computer iteration over for a domain::" & compNames
    Next
    If compNames <> "" Then
        compNames = Mid(compNames, 1, InStrRev(compNames, "|") - 1)
    End If
    'MsgBox "computer1|computer2=" & compNames
    getLanInfo = compNames
End Function
