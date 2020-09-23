VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveX Register"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Go!"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.Frame fmeMain 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4935
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   2655
         TabIndex        =   7
         Top             =   600
         Width           =   2655
         Begin VB.OptionButton optUnregister 
            Caption         =   "Unregister"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optRegister 
            Caption         =   "Register"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   4920
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblOpt 
         Caption         =   "Commands:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DLL/OCX Path:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   285
         Width           =   1290
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ActiveX Registration Program"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   5400
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
Dim objReg As New VBRegSvr
    CD.Filter = "DLL Files (*.DLL)|*.DLL|OCX Files (*.OCX)|*.OCX"
    CD.ShowOpen
    If CD.FileName = "" Then
        Exit Sub
    Else
        If objReg.IsDLLActiveX(CD.FileName, False) = False Then
            MsgBox "The file you have selected is not a valid ActiveX file.", vbInformation, "Input Error"
        Else
            txtPath.Text = CD.FileName
            lblOpt.Enabled = True
            optRegister.Enabled = True
            optUnregister.Enabled = True
        End If
    End If
End Sub

Private Sub cmdRegister_Click()
Dim objReg As New VBRegSvr
    If objReg.fVBRegServer(txtPath.Text, IIf(optRegister.Value, True, False)) Then
        MsgBox "Your ActiveX file has successfully been (Un)registered. Congradulations!", vbInformation, "Operation Successful"
    Else
        MsgBox "Cry me a river, you file registration was not successful.", vbCritical, "Too Bad!"
    End If
End Sub

Private Sub txtPath_Change()
    If Not txtPath.Text = CD.FileName Then
        If txtPath.Text = "" Then
            Exit Sub
        Else
            lblOpt.Enabled = False
            optRegister.Enabled = False
            optUnregister.Enabled = False
            cmdRegister.Enabled = False
        End If
        Else
            cmdRegister.Enabled = True
            lblOpt.Enabled = True
            optRegister.Enabled = True
            optUnregister.Enabled = True
        End If
End Sub
