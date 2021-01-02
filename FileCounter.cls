VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private FS As FileSystemObject
Private Carpeta As folder
Private Archivo As file
Private mCancel As Boolean


Public Property Get Cancel() As Boolean
    On Error GoTo errorHandler
    Cancel = mCancel
    Exit Property
errorHandler:
    MsgBox "Error en CFileCounter.Cancel (Property Get)" & vbCrLf & Err.Number & " : " & Err.Description
End Property

Public Property Let Cancel(aCancel As Boolean)
    On Error GoTo errorHandler
    mCancel = aCancel
    Exit Property
errorHandler:
    MsgBox "Error en CFileCounter.Cancel (Property Let)" & vbCrLf & Err.Number & " : " & Err.Description
End Property

Public Function ContarArchivos(path As String, pattern As String, _
            subfolders As Long, Size As Double) As Long
           
    On Error GoTo errorHandler
    Set Carpeta = FS.GetFolder(path)
    Dim file As file
    Dim CantArchivos As Long
    Dim CantSubFolders As Long
    Dim ext As String
    Dim totalBytes As Long
    
    DoEvents

    If mCancel = True Then
        ContarArchivos = 0
        subfolders = 0
        Size = 0
        Exit Function
    End If
    ' Initial amount of subfolders
    CantSubFolders = CantSubFolders + Carpeta.subfolders.Count
    ' Start the analisis with the actual folder file collection.
    For Each file In Carpeta.Files
        Dim ext1 As Long
        Dim ext2 As Long
        
        ext1 = Len(file.Name) - InStrRev(file.Name, ".")
        ext2 = Len(pattern) - InStrRev(pattern, ".")
        
        ' The counter increases with each file which extension
        ' is coincident with the required pattern.
        ' At the same time, the total size (bytes) of files gets
        ' computed.
        If Right(UCase(file.Name), ext1) = Right(UCase(pattern), ext2) Then
            CantArchivos = CantArchivos + 1
            totalBytes = totalBytes + file.Size
        ElseIf pattern = "*.*" Then
            CantArchivos = CantArchivos + 1
            totalBytes = totalBytes + file.Size
        End If
    Next
    
    Dim folder As folder
    ' Now for each subfolder i search its child subfolders.
    ' Now the recursive phase ...
    For Each folder In Carpeta.subfolders
        ' Change the path to that of the respective
        ' subfolder...
        path = folder.path
        ' Increase the amount of files and call
        ' recursively this same function passing
        ' the corresponding arguments.
        CantArchivos = CantArchivos + ContarArchivos(path, pattern, subfolders, Size)
        CantSubFolders = CantSubFolders + Carpeta.subfolders.Count
    Next
    
    subfolders = subfolders + CantSubFolders
    Size = Size + totalBytes
    ContarArchivos = CantArchivos
    
    Exit Function
errorHandler:
    MsgBox "error en ContadorDeArchivos.ContarArchivos() ; " & Err.Description
    
End Function

Private Sub Class_Initialize()
    Set FS = New FileSystemObject
    mCancel = False
End Sub
