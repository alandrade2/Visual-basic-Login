VERSION 5.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Clique para Sair"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7320
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   5415
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
  Me.Caption = "Exemplo de Login"
  Me.Label1.Caption = "BEM VINDO AO FORMULARIO PRINCIPAL"
End Sub


