VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton Command2 
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Entrar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   30
         PasswordChar    =   "?"
         TabIndex        =   3
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label labform 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SENHA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label labform 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3285
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Prefira criar functions ou procedures separadas e colocar no evento do objeto somente
'esses procedimentos ou funções para serem chamados
    
    If Logon = True Then
       Principal.Show 1
    Else
       MsgBox "Usuário ou senha invalidos", vbCritical, "Validação de Login e senha"
       Me.Text1.SetFocus
    End If
End Sub

'Function criada para fazer a validação, fora do evento do botao
Private Function Logon() As Boolean
Dim bool As Boolean
  If UCase(Me.Text1.Text) = "ADMIN" And Me.Text2.Text = 123456 Then
      bool = True
  Else
      bool = False
  End If

   Logon = bool

End Function



Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
  Me.Caption = "Exemplo de Login"
End Sub

'Efeito de ativar o botao de entrar, somente quando os requisitos forem atendidos
'Nesse caso, enquanto campo de login E o campo de Senha nao estiverem mais de 4 digitos, o botao de entrar não é liberado
Private Sub Text1_Change()
  Me.Command1.Enabled = (Len(Me.Text1.Text) > 4 And _
                          Len(Me.Text2.Text) > 4)
End Sub

Private Sub Text2_Change()
  Me.Command1.Enabled = (Len(Me.Text1.Text) > 4 And _
                          Len(Me.Text2.Text) > 4)
End Sub

'Evento que provoca um efeito visual no formulario, quando o campo recebe o foco, ele seleciona todo o texto dentro dele
Private Sub Text1_GotFocus()
    With Me.Text1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub Text2_GotFocus()
    With Me.Text2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


