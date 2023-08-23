VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Costo_Vencimiento_Cantidad 
   ClientHeight    =   4110
   ClientLeft      =   5340
   ClientTop       =   2085
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   8370
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   5400
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CheckBox Check1 
         Caption         =   "Crear precio por porcentaje"
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "Costo_Vencimiento_Cantidad.frx":0000
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "Costo_Vencimiento_Cantidad.frx":0072
      Top             =   240
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Costo_Vencimiento_Cantidad.frx":02A6
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   1680
      OleObjectBlob   =   "Costo_Vencimiento_Cantidad.frx":0324
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   495
      Left            =   1680
      OleObjectBlob   =   "Costo_Vencimiento_Cantidad.frx":038C
      TabIndex        =   6
      Top             =   240
      Width           =   3135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "Costo_Vencimiento_Cantidad.frx":0408
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "Costo_Vencimiento_Cantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim Asciiaux As Integer
Dim textboxaux As String

'------------------------------------------------------------------------------------------------

Private Sub Solo_Numeros()
If InStr("0123456789", Chr(Asciiaux)) = 0 And Asciiaux <> 8 Then
    Asciiaux = 0
    MsgBox ("Solo ingresar numeros")
End If
End Sub



'------------------------------------------------------------------------------------------------
'--------------------------------------SUBS-PROPIAS----------------------------------------------
'------------------------------------------------------------------------------------------------

Private Sub Command1_Click()


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click(Index As Integer)

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin1.ApplySkin Me.hWnd

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
RS.Source = "Productos"
RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
RS.Open "select * from Productos", cn
RS.MoveFirst
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

'Solo_Numeros (KeyAscii)
'KeyAscii = Asciiaux
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Asciiaux = KeyAscii
Solo_Numeros
KeyAscii = Asciiaux
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
textboxaux = Text2.Text
Asciiaux = KeyAscii
'Solo_numeros_y_Un_punto (KeyAscii)
KeyAscii = Asciiaux
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
textboxaux = Text2.Text
Asciiaux = KeyAscii
'Solo_numeros_y_Un_punto (KeyAscii)
KeyAscii = Asciiaux
End Sub

Private Sub Text4_Change()
Text3.Text = Str(Round(((Val(Text4.Text) * 0.01) * Val(Text2.Text)) + Val(Text2.Text), 3))
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Asciiaux = KeyAscii
Solo_Numeros
KeyAscii = Asciiaux
End Sub
