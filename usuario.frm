VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form usuario 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   15315
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc CONEXION 
      Height          =   495
      Left            =   2280
      Top             =   7800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
   Begin MSComctlLib.ListView LISTA_USUARIO 
      DragIcon        =   "usuario.frx":0000
      Height          =   7215
      Left            =   5040
      TabIndex        =   3
      Top             =   720
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12726
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Audiowide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NOMBRE"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "APELLIDO"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DNI"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NIVEL"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   7935
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   13996
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "LISTA DE USUARIOS"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4455
      Begin VB.TextBox Text5 
         BackColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   6960
         Width           =   855
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H0000FF00&
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   6600
         Width           =   855
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn4 
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   5760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "SALIR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Audiowide"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "usuario.frx":08CA
         PICN            =   "usuario.frx":08E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
         Height          =   735
         Left            =   240
         TabIndex        =   12
         Top             =   5040
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "ELIMINAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Audiowide"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "usuario.frx":11C0
         PICN            =   "usuario.frx":11DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
         Height          =   735
         Left            =   240
         TabIndex        =   11
         Top             =   4320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "EDITAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Audiowide"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "usuario.frx":1AB6
         PICN            =   "usuario.frx":1AD2
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
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "NUEVO"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Audiowide"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "usuario.frx":23AC
         PICN            =   "usuario.frx":23C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   2880
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "usuario.frx":2CA2
         TabIndex        =   8
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1800
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "usuario.frx":2D02
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   720
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "usuario.frx":2D6C
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "usuario.frx":2DD2
         TabIndex        =   14
         Top             =   6600
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "usuario.frx":2E48
         TabIndex        =   16
         Top             =   6960
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   13996
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "INFORMACION DEL USUARIO"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "usuario.frx":2EB2
      Top             =   3480
   End
End
Attribute VB_Name = "usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public li As ListItem
Public DNI As String

Private Sub ChameleonBtn1_Click()
DNI_USUARIO = ""
USUARIO_NUEVO.Show
usuario.Hide
Unload Me
End Sub

Private Sub ChameleonBtn2_Click()
USUARIO_NUEVO.Show
usuario.Hide
Unload Me
End Sub

Private Sub ChameleonBtn4_Click()
 Login_y_Menu.Show
    Login_y_Menu.SkinLabel1.Visible = False
    Login_y_Menu.SkinLabel2.Visible = False
    Login_y_Menu.Text1.Visible = False
    Login_y_Menu.Text2.Visible = False
    Login_y_Menu.ChameleonBtn4.Caption = "Atras"
    Login_y_Menu.ChameleonBtn4.Visible = True
    Login_y_Menu.ChameleonBtn3.Visible = False
    Login_y_Menu.Picture1.Visible = True
    Login_y_Menu.TEXTO.Visible = True
    Login_y_Menu.Width = 14025
    Login_y_Menu.Height = 5685
    Login_y_Menu.ChameleonBtn4.Left = 12000
    Login_y_Menu.ChameleonBtn4.Top = 3360
    Login_y_Menu.CENTRO_BAR.Visible = True
    usuario.Hide
    Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin1.ApplySkin usuario.hWnd
'CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
'CONEXION.CursorType = adOpenDynamic
'CONEXION.RecordSource = "USUARIOS"
'CONEXION.Refresh
'CONEXION.Recordset.Close
CARGARLISTA
End Sub
Private Sub CARGARLISTA()
Dim NIVEL As String
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "USUARIOS"
CONEXION.Refresh
LISTA_USUARIO.ListItems.Clear
If CONEXION.Recordset.EOF = False Then
CONEXION.Recordset.MoveFirst
While CONEXION.Recordset.EOF = False
        Set li = LISTA_USUARIO.ListItems.Add(, , CONEXION.Recordset("NOMBRE"))
            li.ListSubItems.Add , , CONEXION.Recordset("APELLIDO")
            li.ListSubItems.Add , , CONEXION.Recordset("DNI")
            NIVEL = CONEXION.Recordset("NivelDeAcceso")
                If NIVEL = 1 Then
                    li.ListSubItems.Add , , "ADMINISTRADOR"
                Else
                If NIVEL = 2 Then
                    li.ListSubItems.Add , , "VENDEDOR"
                 Else
                    li.ListSubItems.Add , , "REPARADOR"
                End If
                End If
                Colocar_Color
    CONEXION.Recordset.MoveNext
Wend
CONEXION.Recordset.Close

End If
'Dim ADO As New ADODB.Connection
'Dim RS As New ADODB.Recordset






End Sub
Private Sub Colocar_Color()
If CONEXION.Recordset("NivelDeAcceso") = 1 Then
    li.ForeColor = &HC000&
    For i = 1 To 3
        li.ListSubItems(i).ForeColor = &HC000&
    Next i
End If
If CONEXION.Recordset("NivelDeAcceso") = 2 Then
    li.ForeColor = &HFF0000
    For i = 1 To 3
        li.ListSubItems(i).ForeColor = &HFF0000
    Next i
End If
End Sub

Private Sub LISTA_USUARIO_Click()
If Not LISTA_USUARIO.SelectedItem Is Nothing Then
    Dim NOMBRE, APELLIDO As String
    DNI = LISTA_USUARIO.SelectedItem.SubItems(2)
    DNI_USUARIO = LISTA_USUARIO.SelectedItem.SubItems(2)
    APELLIDO = LISTA_USUARIO.SelectedItem.SubItems(1)
    NOMBRE = LISTA_USUARIO.SelectedItem
    Text3.Text = DNI
    Text2.Text = APELLIDO
    Text1.Text = NOMBRE
End If
End Sub
