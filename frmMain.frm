VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4950
      Left            =   5565
      TabIndex        =   7
      Top             =   210
      Width           =   4320
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   2940
         TabIndex        =   19
         Top             =   4410
         Width           =   1170
      End
      Begin VB.CommandButton cmdContar 
         BackColor       =   &H8000000C&
         Caption         =   "Count"
         Height          =   435
         Left            =   1470
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3045
         Width           =   2640
      End
      Begin VB.TextBox txtFiles 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1470
         TabIndex        =   12
         Top             =   840
         Width           =   1065
      End
      Begin VB.TextBox txtDirs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1470
         TabIndex        =   11
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox txtTotalBytes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1470
         TabIndex        =   10
         Top             =   1365
         Width           =   2640
      End
      Begin VB.TextBox txtTotalKBytes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1470
         TabIndex        =   9
         Top             =   1890
         Width           =   2640
      End
      Begin VB.TextBox txtTotalMBytes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1470
         TabIndex        =   8
         Top             =   2415
         Width           =   2640
      End
      Begin VB.Label Label2 
         Caption         =   "Folders"
         Height          =   330
         Left            =   210
         TabIndex        =   17
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label3 
         Caption         =   "Files"
         Height          =   330
         Left            =   210
         TabIndex        =   16
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label4 
         Caption         =   "Total Bytes"
         Height          =   330
         Left            =   210
         TabIndex        =   15
         Top             =   1365
         Width           =   1170
      End
      Begin VB.Label Label5 
         Caption         =   "Total KBytes"
         Height          =   330
         Left            =   210
         TabIndex        =   14
         Top             =   1890
         Width           =   1170
      End
      Begin VB.Label Label6 
         Caption         =   "Total MBytes"
         Height          =   330
         Left            =   210
         TabIndex        =   13
         Top             =   2415
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4950
      Left            =   105
      TabIndex        =   1
      Top             =   210
      Width           =   5475
      Begin VB.FileListBox File1 
         Height          =   4185
         Left            =   2940
         TabIndex        =   5
         Top             =   630
         Width           =   2430
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   4305
         TabIndex        =   4
         Top             =   210
         Width           =   1065
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Width           =   2745
      End
      Begin VB.DirListBox Dir1 
         Height          =   4140
         Left            =   105
         TabIndex        =   2
         Top             =   630
         Width           =   2745
      End
      Begin VB.Label Label1 
         Caption         =   "Filter"
         Height          =   225
         Left            =   3780
         TabIndex        =   6
         Top             =   210
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   435
      Left            =   9555
      TabIndex        =   0
      Top             =   6195
      Width           =   1065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================================================================
' Project      : FileCounter
'
' Module       : frmMain
'
' Author       : Ramon Antonio Gimenez (ramtonio@yahoo.com)
'                      Formosa
'                     Argentina
'
' Created On   : Nov 27, 2001
'
' Description
' -----------
'=====================================================================================
'    DECLARATIONS...
'=====================================================================================

Option Explicit

Private FileSelected           As String
Private FS                     As New FileSystemObject
Private filer                  As CFileCounter

'=====================================================================================
'    PROPERTIES AND METHODS...
'=====================================================================================

 


'=====================================================================================
'    Sub frmMain.cboFiltro_Change()...
'=====================================================================================
Private Sub cboFiltro_Change()
    On Error GoTo errorHandler
    
    File1.pattern = cboFiltro.Text
    File1.Refresh

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.cboFiltro_Change ; " & Err.Number & vbCrLf & Err.Description
End Sub

'=====================================================================================
'    Sub frmMain.cboFiltro_Click()...
'=====================================================================================
Private Sub cboFiltro_Click()
    On Error GoTo errorHandler
    
    File1.pattern = cboFiltro.Text
    File1.Refresh

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.cboFiltro_Click ; " & Err.Number & vbCrLf & Err.Description
End Sub

'=====================================================================================
'    Sub frmMain.cmdContar_Click()...
'=====================================================================================
Private Sub cmdContar_Click()
    On Error GoTo errorHandler
    
    Dim TotalDirs As Long
    Dim TotalFiles As Long
    Dim totalBytes As Double
    
    txtFiles.Text = ""
    txtDirs.Text = ""
    txtTotalBytes.Text = ""
    txtTotalKBytes.Text = ""
    txtTotalMBytes.Text = ""
    
    Me.MousePointer = vbHourglass
    Set frmInProcess.mFC = filer
    frmInProcess.Show
    TotalFiles = filer.ContarArchivos(File1.path, File1.pattern, TotalDirs, totalBytes)
    If filer.Cancel = True Then
        filer.Cancel = False
        MsgBox "Action Canceled by User", vbOKOnly, "Cancel Action"
        TotalFiles = 0
        TotalDirs = 0
        totalBytes = 0
    End If
        
    Me.Refresh
    Me.MousePointer = vbNormal
    Unload frmInProcess
    
    txtFiles.Text = Format(TotalFiles, "#,##0")
    txtDirs.Text = Format(TotalDirs, "#,##0")
    txtTotalBytes.Text = Format(totalBytes, "#,##0")
    txtTotalKBytes.Text = Format(totalBytes / 1024, "#,##0.0")
    txtTotalMBytes.Text = Format(totalBytes / (1024 ^ 2), "#,##0.0")

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.cmdContar_Click ; " & Err.Number & vbCrLf & Err.Description
End Sub


'=====================================================================================
'    Sub frmMain.cmdSalir_Click()...
'=====================================================================================
Private Sub cmdExit_Click()
    Unload Me
End Sub


'=====================================================================================
'    Sub frmMain.Dir1_Change()...
'=====================================================================================
Private Sub Dir1_Change()
    On Error GoTo errorHandler
    
    ' Cuando cambia el dirlist, sincronizo el FileList
    File1.path = Dir1.path

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.Dir1_Change ; " & Err.Number & vbCrLf & Err.Description
End Sub

'=====================================================================================
'    Sub frmMain.Drive1_Change()...
'=====================================================================================
Private Sub Drive1_Change()
    On Error GoTo errorHandler
    
    Dir1.path = Drive1.Drive

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.Drive1_Change ; " & Err.Number & vbCrLf & Err.Description
End Sub





'=====================================================================================
'    Sub frmMain.Form_Load()...
'=====================================================================================
Private Sub Form_Load()
    On Error GoTo errorHandler
    
    Set filer = New CFileCounter
    Dir1.path = App.path
    Drive1.Drive = Left(App.path, 3)
    File1.path = Dir1.path
    
    With cboFiltro
        .AddItem "*.*", 0
        .Text = .List(0)
    End With
    
    File1.pattern = cboFiltro.Text

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.Form_Load ; " & Err.Number & vbCrLf & Err.Description
End Sub



'=====================================================================================
'    Sub frmMain.Form_Unload()...
'=====================================================================================
Private Sub Form_Unload(Cancel As Integer)
        Unload frmInProcess
End Sub

