VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form RECEPCION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEPCION DE MERCADERIA"
   ClientHeight    =   6075
   ClientLeft      =   705
   ClientTop       =   2250
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8445
   Begin VB.Frame Frame4 
      Caption         =   "ENTRADA DE CLIENTE"
      Height          =   3495
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   6255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   615
         Left            =   1080
         OleObjectBlob   =   "RECEPCION.frx":0000
         TabIndex        =   31
         Top             =   600
         Width           =   3975
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
         Height          =   1335
         Left            =   3480
         TabIndex        =   28
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2355
         BTYPE           =   14
         TX              =   "NUEVO CLIENTE"
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
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "RECEPCION.frx":00AA
         PICN            =   "RECEPCION.frx":00C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
         Height          =   1335
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2355
         BTYPE           =   14
         TX              =   "CLIENTE REGISTRADO"
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
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "RECEPCION.frx":09A0
         PICN            =   "RECEPCION.frx":09BC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RECEPCION.frx":1296
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   "PROVEEDOR"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   32
      Top             =   600
      Width           =   5295
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "RECEPCION.frx":1B70
         TabIndex        =   35
         Top             =   480
         Width           =   3135
      End
      Begin VB.ListBox LISTA_PROVEEDOR 
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         IntegralHeight  =   0   'False
         ItemData        =   "RECEPCION.frx":1BEE
         Left            =   240
         List            =   "RECEPCION.frx":1BF0
         TabIndex        =   34
         Top             =   720
         Width           =   4695
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "RECEPCION.frx":1BF2
         Left            =   2280
         List            =   "RECEPCION.frx":1BFC
         TabIndex        =   33
         Text            =   "CON  GARANTIA"
         Top             =   2640
         Width           =   2655
      End
   End
   Begin MSAdodcLib.Adodc CONEXION 
      Height          =   375
      Left            =   9840
      Top             =   5640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "CLIENTE"
      Caption         =   "CONEXION"
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H000080FF&
      Caption         =   "INGRESAR NUEVO CLIENTE"
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5280
      Width           =   3375
   End
   Begin ChamaleonButton.ChameleonBtn GUARDAR 
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "GUARDAR"
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
      BCOL            =   8421504
      BCOLO           =   8421504
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "RECEPCION.frx":1C1C
      PICN            =   "RECEPCION.frx":1C38
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "RECEPCION.frx":2512
      Left            =   13080
      List            =   "RECEPCION.frx":251C
      TabIndex        =   17
      Text            =   "DNI"
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Caption         =   "BUSCAR CLIENTE"
      Height          =   1215
      Left            =   8520
      TabIndex        =   16
      Top             =   240
      Width           =   7455
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "RECEPCION.frx":252F
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   18
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATOS DE ARTICULO RECIBIDO"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7695
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         TabIndex        =   39
         Text            =   "INGRESE LA MARCA"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "INGRE. NUEVA MARCA"
         Height          =   375
         Left            =   1440
         TabIndex        =   38
         Top             =   1680
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   375
         Left            =   3600
         OleObjectBlob   =   "RECEPCION.frx":259B
         TabIndex        =   37
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         TabIndex        =   36
         Top             =   3360
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   615
         Left            =   5160
         OleObjectBlob   =   "RECEPCION.frx":2623
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "RECEPCION.frx":26A1
         TabIndex        =   23
         Top             =   2880
         Width           =   3255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   375
         Left            =   4800
         OleObjectBlob   =   "RECEPCION.frx":271F
         TabIndex        =   22
         Top             =   1680
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "RECEPCION.frx":2793
         TabIndex        =   21
         Top             =   1680
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "RECEPCION.frx":27F7
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3360
         Width           =   3255
      End
      Begin ACTIVESKINLibCtl.SkinLabel FECHA 
         Height          =   375
         Left            =   5040
         OleObjectBlob   =   "RECEPCION.frx":2861
         TabIndex        =   13
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   12
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATOS DEL CLIENTE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6375
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "RECEPCION.frx":28D8
         TabIndex        =   9
         Top             =   2400
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   1080
         OleObjectBlob   =   "RECEPCION.frx":2942
         TabIndex        =   8
         Top             =   1680
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "RECEPCION.frx":29A2
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "RECEPCION.frx":2A0C
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   5
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   3735
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1080
      OleObjectBlob   =   "RECEPCION.frx":2A74
      Top             =   1320
   End
   Begin MSComctlLib.ListView LISTA_CLIENTE 
      Height          =   3615
      Left            =   8520
      TabIndex        =   15
      Top             =   1560
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483643
      BackColor       =   -2147483640
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CLIENTE"
         Object.Width           =   4411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "APELLIDO"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DNI"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.TabStrip MOSTRAR 
      Height          =   5055
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CLIENTE"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ARTICULO RECIBIDO"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "INGRESAR PROVEEDOR"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu NM_VOLVER 
      Caption         =   "VOLVER"
   End
   Begin VB.Menu MN_SALIR 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "RECEPCION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NUEVOART As String
Public MISMO_CLIENTE As Boolean
Public PASARUNAVEZ As Integer
Public COD_CLIENTE As Integer
Public MARCA As Integer
Public li As ListItem
Private Sub ChameleonBtn2_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = True
RECEPCION.Width = 16470
LISTA_CLIENTE.Visible = True
GUARDAR.Visible = True
Combo1.Visible = True
Frame4.Visible = False
MOSTRAR.Visible = True
Check1.Visible = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub

Private Sub ChameleonBtn3_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
RECEPCION.Width = 8900
LISTA_CLIENTE.Visible = False
GUARDAR.Visible = True
Combo1.Visible = False
Frame4.Visible = False
MOSTRAR.Visible = True
Check1.Visible = False

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Frame5.Visible = False
    RECEPCION.Width = 9900
    LISTA_CLIENTE.Visible = False
    GUARDAR.Visible = True
    Combo1.Visible = False
    Frame4.Visible = False
    MOSTRAR.Visible = True
    Check1.Visible = True
    Check1.BackColor = &HC000&
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    If MOSTRAR.SelectedItem.Index = 2 And Frame1.Visible = True Then
        Frame2.Visible = True
        Frame1.Visible = False
        Frame5.Visible = False
    End If
    If MOSTRAR.SelectedItem.Index = 3 And Frame1.Visible = True Then
        Frame2.Visible = False
        Frame1.Visible = False
        Frame5.Visible = True
    End If
    Else
    Frame3.Visible = True
    RECEPCION.Width = 16470
    LISTA_CLIENTE.Visible = True
    GUARDAR.Visible = True
    Combo1.Visible = True
    Frame4.Visible = False
    MOSTRAR.Visible = True
    Check1.Visible = True
    Check1.BackColor = &H80FF&
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then
    Text6.Visible = True
    Combo3.Visible = False
Else
    Text6.Visible = False
    Combo3.Visible = True
End If
End Sub
Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin1.ApplySkin RECEPCION.hWnd
PASARUNAVEZ = 0
Text6.Visible = False
CARGAR_LISTA_a_COMBO
If EDITAR = False Then
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame5.Visible = False
    RECEPCION.Width = 8900
    LISTA_CLIENTE.Visible = False
    GUARDAR.Visible = False
    Combo1.Visible = False
    MOSTRAR.Width = 6855
    MOSTRAR.Height = 4095
    MOSTRAR.Visible = False
    Check1.Visible = False
    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "CONSULTA_CLIENTE"
    CONEXION.Refresh
    If CONEXION.Recordset.EOF = False Then
    CONEXION.Recordset.MoveFirst
    While CONEXION.Recordset.EOF = False
           Set li = LISTA_CLIENTE.ListItems.Add(, , CONEXION.Recordset("NOMBRE"), , 1)
                 li.ListSubItems.Add , , CONEXION.Recordset("APELLIDO")
                 li.ListSubItems.Add , , CONEXION.Recordset("DNI")
            CONEXION.Recordset.MoveNext
    Wend
    End If
    CONEXION.Recordset.Close
    FECHA.Caption = Format(Now, "DD/MM/YYYY")
    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "ARTICULO_RECIBIDO"
    CONEXION.Refresh
    CONEXION.Recordset.Close
    CARGAR_PROVEEDOR
End If
If EDITAR = True Then
    Frame4.Visible = False
    Frame2.Visible = False
    Frame5.Visible = False
    RECEPCION.Width = 8700
    MOSTRAR.Width = 6855
    MOSTRAR.Height = 4095
    Check1.Visible = False
    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "CLIENTE"
    CONEXION.Refresh
    CONEXION.Recordset.MoveFirst
    Do While CONEXION.Recordset.EOF = False
         If CLIENTE = CONEXION.Recordset("ID") Then
              Text1.Text = CONEXION.Recordset("NOMBRE")
              Text2.Text = CONEXION.Recordset("APELLIDO")
              Text3.Text = CONEXION.Recordset("DNI")
              Text4.Text = CONEXION.Recordset("TELEFONO")
              Exit Do
         End If
    CONEXION.Recordset.MoveNext
   Loop
    CONEXION.Recordset.Close

    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "ARTICULO_RECIBIDO"
    CONEXION.Refresh
    CONEXION.Recordset.MoveFirst
    Do While CONEXION.Recordset.EOF = False
        If Val(ARTICULO) = CONEXION.Recordset("ID") Then
                    Text5.Text = CONEXION.Recordset("ARTICULO")
                    'Text6.Text = CONEXION.Recordset("MARCA")AQUI LA MARCA
                    Text7.Text = CONEXION.Recordset("DESCRIPCION")
                    Text8.Text = CONEXION.Recordset("APORTE")
                    FECHA.Caption = CONEXION.Recordset("FECHA")
                    Exit Do
        End If
    CONEXION.Recordset.MoveNext
   Loop
   CONEXION.Recordset.Close
   CARGAR_PROVEEDOR
End If
End Sub
Private Sub GUARDAR_Click()
CARGAR_MARCAS
If EDITAR = False Then
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "CLIENTE"
CONEXION.Refresh
Dim COD_CLIENTE As Integer
If PASARUNAVEZ = 0 Then
    CONEXION.Recordset.AddNew
    CONEXION.Recordset("NOMBRE") = Text1.Text
    CONEXION.Recordset("APELLIDO") = Text2.Text
    CONEXION.Recordset("DNI") = Text3.Text
    CONEXION.Recordset("TELEFONO") = Val(Text4.Text)
    CONEXION.Recordset.Update
End If

CONEXION.Recordset.MoveLast
COD_CLIENTE = CONEXION.Recordset("ID")
CONEXION.Recordset.Close
If COD_CLIENTE <> 0 Then
    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "ARTICULO_RECIBIDO"
    CONEXION.Refresh
    CONEXION.Recordset.AddNew
    CONEXION.Recordset("NRO_CLIENTE") = COD_CLIENTE
    CONEXION.Recordset("ARTICULO") = Text5.Text
    CONEXION.Recordset("MARCA") = MARCA ' LA MARCA
    CONEXION.Recordset(3) = Text7.Text
    CONEXION.Recordset("DESCRIPCION") = Text10.Text
    CONEXION.Recordset("APORTE") = Text8.Text
    CONEXION.Recordset("ESTADO") = "RECIVIDO"
    CONEXION.Recordset("FECHA") = FECHA.Caption
    CONEXION.Recordset("GARANTIA") = Combo2.Text
    CONEXION.Recordset("PROVEEDOR") = LISTA_PROVEEDOR.Text
    CONEXION.Recordset.Update
    'CONEXION.Recordset.Close
End If
NUEVOART = MsgBox("¿DECEA REGISTRAR OTRO ARTICULO DEL CLIENTE?", vbYesNo)
If NUEVOART = vbYes Then
     PASARUNAVEZ = PASARUNAVEZ + 1
     MISMO_CLIENTE = True
Else
    MISMO_CLIENTE = False
    PASARUNAVEZ = 0

End If

If MISMO_CLIENTE = True Then
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
Else
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
End If
End If

If EDITAR = True Then
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "CLIENTE"
CONEXION.Refresh

Dim GARANTINA As String
CONEXION.Recordset.MoveFirst
    Do While CONEXION.Recordset.EOF = False
         If CLIENTE = CONEXION.Recordset("ID") Then
               CONEXION.Recordset("NOMBRE") = Text1.Text
               CONEXION.Recordset("APELLIDO") = Text2.Text
               CONEXION.Recordset("DNI") = Text3.Text
               CONEXION.Recordset("TELEFONO") = Text4.Text
               CONEXION.Recordset.Update
              Exit Do
         End If
    CONEXION.Recordset.MoveNext
   Loop
 CONEXION.Recordset.Close
  CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
  CONEXION.CursorType = adOpenDynamic
  CONEXION.RecordSource = "ARTICULO_RECIBIDO"
  CONEXION.Refresh
  CONEXION.Recordset.MoveFirst
    Do While CONEXION.Recordset.EOF = False
        If Val(ARTICULO) = CONEXION.Recordset("ID") Then
                     CONEXION.Recordset("ARTICULO") = Text5.Text
                     CONEXION.Recordset("MARCA") = MARCA ' AQUI LA MARCA
                     CONEXION.Recordset(3) = Text7.Text
                     CONEXION.Recordset("DESCRIPCION") = Text10.Text
                     CONEXION.Recordset("APORTE") = Text8.Text
                     GARANTIA = Combo2.Text
                     CONEXION.Recordset("GARANTIA") = GARANTIA
                     CONEXION.Recordset("PROVEEDOR") = LISTA_PROVEEDOR.Text
                     CONEXION.Recordset.Update
                    Exit Do
        End If
    CONEXION.Recordset.MoveNext
   Loop
   CONEXION.Recordset.Close
EDITAR = False
CLIENTE = 0
ARTICULO = 0
REGISTRO_DE_LLEGADAS.Show
RECEPCION.Hide
Unload Me
End If
End Sub

Private Sub LISTA_CLIENTE_Click()
Dim SELECCIONAR As String
SELECCIONAR = LISTA_CLIENTE.SelectedItem.SubItems(2)
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "CLIENTE"
CONEXION.Refresh
CONEXION.Recordset.MoveFirst

While CONEXION.Recordset.EOF = False
    If SELECCIONAR = CONEXION.Recordset("DNI") Then
    Text1.Text = CONEXION.Recordset("NOMBRE")
    Text2.Text = CONEXION.Recordset("APELLIDO")
    Text3.Text = CONEXION.Recordset("DNI")
    Text4.Text = CONEXION.Recordset("TELEFONO")
    Exit Sub
    End If
    CONEXION.Recordset.MoveNext
Wend
End Sub
Private Sub MN_SALIR_Click()
End
End Sub
Private Sub MOSTRAR_Click()
Dim CONTROL As Boolean

If MOSTRAR.SelectedItem.Caption = "CLIENTE" Then
    Frame2.Visible = False
    Frame1.Visible = True
    Frame5.Visible = False
    MOSTRAR.Width = 6855
    MOSTRAR.Height = 4095
End If


If MOSTRAR.SelectedItem.Caption = "ARTICULO RECIBIDO" Then
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" Then
    Frame2.Visible = True
    Frame1.Visible = False
    Frame5.Visible = False
    MOSTRAR.Height = 4935
    MOSTRAR.Width = 8175
    Else
    MsgBox ("COMPLETE LOS DATOS DEL CLIENTES")
    Text1.SetFocus
    Exit Sub
End If
End If
    

If MOSTRAR.SelectedItem.Caption = "INGRESAR PROVEEDOR" Then
If Text5.Text <> "" And Combo3.Text <> "" Or Text6.Text <> "" And Text7.Text <> "" Then
    Frame2.Visible = False
    Frame1.Visible = False
    Frame5.Visible = True
    MOSTRAR.Width = 6000
    MOSTRAR.Height = 4595
    Else
    MsgBox ("COMPLETE LOS DATOS DEL ARTICULO")
    Text5.SetFocus
    Exit Sub
End If
End If

End Sub

Private Sub MOSTRAR_Validate(Cancel As Boolean)
If MOSTRAR.SelectedItem.Caption = "ARTICULO RECIBIDO" And CONTROL = False Then
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" Then
CONTROL = True
MsgBox ("COMPLETE LOS DATOS DEL CLIENTES")
Cancel = True
End If
End If
'If Text5.Text = "" And Combo3.Text = "" Or Text6.Text = "" And Text7.Text = "" Then
'CONTROL = True
'End If
End Sub

Private Sub NM_VOLVER_Click()
EDITAR = False
REGISTRO_DE_LLEGADAS.Show
RECEPCION.Hide
Unload Me
End Sub
Private Sub BUSQUEDA()
CONEXION.ConnectionString = "provider=microsoft.jet.oledb.4.0;" & "data source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "CLIENTE"
CONEXION.Refresh
Dim cantaux As String
Dim rsaux As String

cantaux = 1
LISTA_CLIENTE.ListItems.Clear
CONEXION.Recordset.MoveFirst

For i = 0 To (CONEXION.Recordset.RecordCount - 1)
    rsaux = CONEXION.Recordset(Combo1.Text)
    If Mid(rsaux, 1, 1) = Mid(Text9.Text, 1, 1) Then
        For j = 2 To Len(Text9.Text)
            If Mid(CONEXION.Recordset(Combo1.Text), 1, j) = Mid(Text9.Text, 1, j) Then
                cantaux = cantaux + 1
            End If
        Next j
        If cantaux = Len(Text9.Text) Then
           Set li = LISTA_CLIENTE.ListItems.Add(, , CONEXION.Recordset("NOMBRE"))
                 li.ListSubItems.Add , , CONEXION.Recordset("APELLIDO")
                 li.ListSubItems.Add , , CONEXION.Recordset("DNI")
        End If
        cantaux = 1
    End If
   CONEXION.Recordset.MoveNext
Next i
CONEXION.Recordset.Close
End Sub

Private Sub Text9_Change()
If Text9.Text <> "" Then
    BUSQUEDA
Else
    CARGARLISTA
End If
End Sub
Private Sub CARGARLISTA()
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "CLIENTE"
CONEXION.Refresh
LISTA_CLIENTE.ListItems.Clear
CONEXION.Recordset.MoveFirst
While CONEXION.Recordset.EOF = False
        Set li = LISTA_CLIENTE.ListItems.Add(, , CONEXION.Recordset("NOMBRE"))
                 li.ListSubItems.Add , , CONEXION.Recordset("APELLIDO")
                 li.ListSubItems.Add , , CONEXION.Recordset("DNI")
    CONEXION.Recordset.MoveNext
Wend
CONEXION.Recordset.Close
End Sub
Sub CARGAR_PROVEEDOR()
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "PROVEEDOR"
CONEXION.Refresh
CONEXION.Recordset.MoveFirst
While UCase(CONEXION.Recordset.EOF = False)
    LISTA_PROVEEDOR.AddItem (CONEXION.Recordset("NOMBRE"))
    CONEXION.Recordset.MoveNext
Wend
CONEXION.Recordset.Close
End Sub
Sub CARGAR_MARCAS()
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "MARCAS"
CONEXION.Refresh
If Check2.Value = 1 Then
    CONEXION.Recordset.AddNew
    CONEXION.Recordset("MARCA") = Text6.Text
    CONEXION.Recordset("PROVEEDOR") = LISTA_PROVEEDOR.Text
    CONEXION.Recordset.Update
    CONEXION.Refresh
    CONEXION.Recordset.MoveLast
    MARCA = CONEXION.Recordset("COD_MARCA")
    Exit Sub
End If
If Check2.Value = 0 Then
    Do While (CONEXION.Recordset.EOF = False)
    If Combo3.Text = CONEXION.Recordset("MARCA") Then
    MARCA = CONEXION.Recordset("COD_MARCA")
    Exit Sub
    End If
    CONEXION.Recordset.MoveNext
Loop
End If
End Sub
Sub CARGAR_LISTA_a_COMBO()
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "MARCAS"
CONEXION.Refresh
CONEXION.Recordset.MoveFirst
Do While (CONEXION.Recordset.EOF = False)
    Combo3.AddItem (CONEXION.Recordset("marca"))
    CONEXION.Recordset.MoveNext
Loop
End Sub
