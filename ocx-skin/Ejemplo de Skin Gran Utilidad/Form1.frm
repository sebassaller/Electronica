VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form1 
   Caption         =   "Skin "
   ClientHeight    =   3870
   ClientLeft      =   4845
   ClientTop       =   2910
   ClientWidth     =   6990
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   1605
      Left            =   3840
      TabIndex        =   4
      Text            =   "Hola [ www.recursosvisualbasic.com.ar ]"
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1605
      Left            =   120
      TabIndex        =   2
      Text            =   "KandelaMerengue@Hotmail.Com"
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3720
      OleObjectBlob   =   "Form1.frx":0000
      Top             =   4080
   End
   Begin VB.Menu MNUSkin 
      Caption         =   "Skin"
      Index           =   0
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "B-Studio"
         Index           =   0
      End
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "Galaxy"
         Index           =   1
      End
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "Green"
         Index           =   2
      End
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "Mac"
         Index           =   3
      End
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "Media"
         Index           =   4
      End
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "Metallic"
         Index           =   5
      End
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "Paper"
         Index           =   6
      End
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "TopSecret"
         Index           =   7
      End
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "Web-II"
         Index           =   8
      End
      Begin VB.Menu MNUCambiarSkin 
         Caption         =   "Zhelezo"
         Index           =   9
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Aplicar_skin Me

End Sub

Private Sub MNUArchivo_Click(Index As Integer)
Unload Me
End Sub


Private Sub MNUCambiarSkin_Click(Index As Integer)

Select Case Index

Case Is = 0
Skin1.LoadSkin App.Path & "\Skins\B-Studio.skn"
Skin1.ApplySkin Form1.hWnd


Case Is = 1
Skin1.LoadSkin App.Path & "\Skins\galaxy.skn"
Skin1.ApplySkin Form1.hWnd


Case Is = 2
Form1.Skin1.LoadSkin App.Path & "\Skins\green.skn"
Skin1.ApplySkin Form1.hWnd


Case Is = 3
Form1.Skin1.LoadSkin App.Path & "\Skins\Mac.skn"
Skin1.ApplySkin Form1.hWnd


Case Is = 4
Form1.Skin1.LoadSkin App.Path & "\Skins\Media.skn"
Skin1.ApplySkin Form1.hWnd


Case Is = 5
Form1.Skin1.LoadSkin App.Path & "\Skins\metallic.skn"
Skin1.ApplySkin Form1.hWnd


Case Is = 6
Form1.Skin1.LoadSkin App.Path & "\Skins\Paper.skn"
Skin1.ApplySkin Form1.hWnd

Case Is = 7
Form1.Skin1.LoadSkin App.Path & "\Skins\TopSecret.skn"
Skin1.ApplySkin Form1.hWnd


Case Is = 8
Form1.Skin1.LoadSkin App.Path & "\Skins\Web-II.skn"
Skin1.ApplySkin Form1.hWnd


Case Is = 9
Form1.Skin1.LoadSkin App.Path & "\Skins\Zhelezo.skn"
Skin1.ApplySkin Form1.hWnd

End Select

End Sub
