VERSION 5.00
Begin VB.Form frmInProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processing...."
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2205
      TabIndex        =   1
      Top             =   930
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Processing...please wait"
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   420
      Width           =   3060
   End
End
Attribute VB_Name = "frmInProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mFC As CFileCounter

Private Sub cmdCancel_Click()
    mFC.Cancel = True
End Sub
