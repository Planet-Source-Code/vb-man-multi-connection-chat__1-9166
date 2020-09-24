VERSION 5.00
Begin VB.Form frmServerOptions 
   Caption         =   "Options"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1260
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      Top             =   1260
      Width           =   1035
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   435
   End
End
Attribute VB_Name = "frmServerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
If IsNumeric(Me.txtPort) Then
    i = modIni.SetINIValue("Server", "Port", Me.txtPort, App.Path & "\chat.ini")
    iPort = Me.txtPort
    frmServer.StatusBar1.Panels.Item(2) = "Port: " & iPort
    Unload Me
Else
    i = MsgBox("You must enter a number", 64, "Error")
End If
End Sub

Private Sub Form_Load()

Me.txtPort = iPort
End Sub

Private Sub txtPort_GotFocus()
Me.txtPort.SelStart = 0
Me.txtPort.SelLength = Len(Me.txtPort)
End Sub

