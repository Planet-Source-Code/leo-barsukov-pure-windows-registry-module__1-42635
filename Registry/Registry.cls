VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The Only Object variable
Dim WSHShell

'Enum for Reg ROOT Folders
Public Enum RegRoots

    'Current User
    HKEY_CURRENT_USER
    'Classes Root
    HKEY_CLASSES_ROOT
    
End Enum


'Get RootFolder Symbol
Private Function Rut(Num)
    
    '#0 Is the HKEY_CURRENT_USER
    If Num = 0 Then Rut = "HKCU"
    '#1 Is the HKEY_CLASSES_ROOT
    If Num = 1 Then Rut = "HKCR"

End Function

'Write to Registry
Public Sub WriteKey(Root As RegRoots, Folder As String, Value As String)
    
    On Error GoTo Err:
    'Create the Shell Object
    Set WSHShell = CreateObject("WScript.Shell")
    'Write to registry through the Shell Object
    WSHShell.regwrite Rut(Root) & "\" & Folder, Value
    
    Exit Sub

Err:
End Sub

'Read From Registry
Public Function ReadKey(Root As RegRoots, Folder As String)
    
    On Error GoTo Err:
    'Create th Shell Object
    Set WSHShell = CreateObject("WScript.Shell")
    'Read from Registry
    ReadKey = WSHShell.regread(Rut(Root) & "\" & Folder)

    Exit Function

Err:
End Function

'Delete From Registry
Public Sub DeleteKey(Root As RegRoots, Folder As String)
    
    On Error GoTo Err:
    'Create Shell Object
    Set WSHShell = CreateObject("WScript.Shell")
    'Delete from Registry
    WSHShell.regdelete Rut(Root) & "\" & Folder

    Exit Sub

Err:
End Sub

'Assiciate Files
Public Sub AssociateFile(ext As String, OpenWith As String, ProgName As String, Optional ico As String)
    
    Set WSHShell = CreateObject("WScript.Shell")

On Error GoTo Nxt:
    
    WSHShell.regwrite "HKCR\AsocBak\" & ext & "\Prog\", WSHShell.regread("HKCR\" & "." & ext & "\")

Nxt:
    
    WSHShell.regwrite "HKCR\" & "." & ext, ""
    WSHShell.regwrite "HKCR\" & "." & ext & "\", ProgName
    WSHShell.regwrite "HKCR\" & ProgName & "\Shell\Open\Command\", OpenWith
    WSHShell.regwrite "HKCR\" & ProgName & "\DefaultIcon\", ico
    
End Sub

'Roll Back Files
Public Sub RollBackFile(ext As String, ProgName As String)
    
    Set WSHShell = CreateObject("WScript.Shell")

On Error GoTo Nxt:
    
    WSHShell.regwrite ("HKCR\" & "." & ext & "\"), WSHShell.regread("HKCR\AsocBak\" & ext & "\Prog\")
    WSHShell.regdelete "HKCR\" & ProgName
    
Nxt:
        
End Sub

