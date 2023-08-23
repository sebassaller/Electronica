VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form USUARIO_NUEVO 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc CONEXION 
      Height          =   330
      Left            =   480
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3840
      OleObjectBlob   =   "USUARIO_NUEVO.frx":0000
      Top             =   360
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "USUARIO_NUEVO.frx":0234
      TabIndex        =   18
      Top             =   6000
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   375
      Left            =   720
      OleObjectBlob   =   "USUARIO_NUEVO.frx":029C
      TabIndex        =   17
      Top             =   6000
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   375
      Left            =   600
      OleObjectBlob   =   "USUARIO_NUEVO.frx":0306
      TabIndex        =   15
      Top             =   3720
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "USUARIO_NUEVO.frx":0374
      TabIndex        =   14
      Top             =   3120
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "USUARIO_NUEVO.frx":03DC
      TabIndex        =   13
      Top             =   2640
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   1920
      OleObjectBlob   =   "USUARIO_NUEVO.frx":0440
      TabIndex        =   12
      Top             =   2160
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   1080
      OleObjectBlob   =   "USUARIO_NUEVO.frx":04A0
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "USUARIO_NUEVO.frx":050A
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
      Height          =   855
      Left            =   960
      TabIndex        =   9
      Top             =   5040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   13
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   4210752
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "USUARIO_NUEVO.frx":0570
      PICN            =   "USUARIO_NUEVO.frx":058C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
      Height          =   855
      Left            =   6120
      TabIndex        =   8
      Top             =   5040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   13
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   4210752
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "USUARIO_NUEVO.frx":0E66
      PICN            =   "USUARIO_NUEVO.frx":0E82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check1 
      Caption         =   "MOSTRAR CONTRASEÑA"
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   4320
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3480
      Width           =   4215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "USUARIO_NUEVO.frx":175C
      Left            =   2760
      List            =   "USUARIO_NUEVO.frx":1769
      TabIndex        =   4
      Text            =   "ELEGIR NIVEL"
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   480
      TabIndex        =   16
      Top             =   600
      Width           =   7215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11245
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "INGRESAR O MODIFICAR /USUARIO"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   8880
      Top             =   1320
      Width           =   615
   End
   Begin VB.Menu mn_volver 
      Caption         =   "vover"
      NegotiatePosition=   2  'Middle
   End
End
Attribute VB_Name = "USUARIO_NUEVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChameleonBtn1_Click()
Dim puesto As Integer
If DNI_USUARIO <> "" Then
Select Case [puesto]
    Case Combo1.ListIndex = 1
    puesto = 2

    Case Combo1.ListIndex = 2
    puesto = 3

 End Select
 MsgBox (puesto)
 Exit Sub



CONEXION.Recordset("nombre") = Text1.Text
CONEXION.Recordset("apellido") = Text2.Text
CONEXION.Recordset("dni") = Text3.Text
CONEXION.Recordset("usuario") = Text4.Text
CONEXION.Recordset("contraseña") = Text5.Text



End If
End Sub

Private Sub Form_Load()
Dim i, a As Integer

Skin1.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin1.ApplySkin USUARIO_NUEVO.hWnd
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "USUARIOS"
CONEXION.Refresh
If DNI_USUARIO <> "" Then
CONEXION.Recordset.MoveFirst
    While Not UCase(CONEXION.Recordset.EOF = True)
    If DNI_USUARIO = CONEXION.Recordset("dni") Then
        Text1.Text = CONEXION.Recordset("nombre")
        Text2.Text = CONEXION.Recordset("apellido")
        Text3.Text = CONEXION.Recordset("dni")
        Text4.Text = CONEXION.Recordset("usuario")
        Text5.Text = CONEXION.Recordset("contraseña")
        a = CONEXION.Recordset("niveldeacceso")
        For i = 1 To Combo1.ListCount
            If a = 2 Then
                Combo1.ListIndex = 1
            End If
            If a = 1 Then
                Combo1.ListIndex = 0
            End If
            If a = 3 Then
               Combo1.ListIndex = 2
            End If
        Next i
        Exit Sub
    End If
    CONEXION.Recordset.MoveNext
Wend
End If

End Sub

Private Sub MN_VOLVER_Click()
usuario.Show
USUARIO_NUEVO.Hide
Unload Me
End Sub
