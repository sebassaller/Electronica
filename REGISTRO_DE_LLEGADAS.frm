VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#2.0#0"; "Toolbar2.ocx"
Begin VB.Form REGISTRO_DE_LLEGADAS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTROS DE LLEGADAS"
   ClientHeight    =   9315
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   19845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   19845
   Begin VB.Frame Frame6 
      Caption         =   "OPCIONES DE BUSQUEDA"
      Height          =   1215
      Left            =   15000
      TabIndex        =   28
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton Option4 
         Caption         =   "BUSCAR ARTICULO"
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "BUSCAR CLIENTE"
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.ComboBox Combo4 
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
      ItemData        =   "REGISTRO_DE_LLEGADAS.frx":0000
      Left            =   15240
      List            =   "REGISTRO_DE_LLEGADAS.frx":000D
      TabIndex        =   24
      Text            =   "OPCIONES DE REPUESTOS"
      Top             =   6600
      Width           =   4215
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
      Height          =   975
      Left            =   15000
      TabIndex        =   18
      Top             =   7920
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1720
      BTYPE           =   14
      TX              =   "INGRESAR REPUESTO"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "REGISTRO_DE_LLEGADAS.frx":0041
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   10920
      TabIndex        =   19
      Text            =   "ELIGA EL PROVEEDOR"
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Caption         =   "AGREGAR REPUESTO"
      Height          =   2175
      Left            =   10680
      TabIndex        =   20
      Top             =   7080
      Width           =   8895
      Begin VB.OptionButton Option2 
         Caption         =   "INGR. NUEVO REPUESTO"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   360
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "REPUESTO EN STOCK"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   2055
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   4095
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4680
      Top             =   6720
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
            Picture         =   "REGISTRO_DE_LLEGADAS.frx":005D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LISTA_RESPUESTOS 
      Height          =   2055
      Left            =   120
      TabIndex        =   16
      Top             =   7080
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList2"
      ColHdrIcons     =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "REPUESTO"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PRECIO"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PROVEEDOR"
         Object.Width           =   5292
      EndProperty
   End
   Begin AIFCmp1.asxToolbar asxToolbar1 
      Height          =   1095
      Left            =   0
      Top             =   120
      Width           =   14115
      _ExtentX        =   24315
      _ExtentY        =   1931
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   -2147483643
      BorderStyle     =   0
      BorderLeft      =   0   'False
      BorderTop       =   0   'False
      BorderRight     =   0   'False
      BorderBottom    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Audiowide"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   5
      SolidChecked    =   -1  'True
      ShowSeparators  =   -1  'True
      BoldOnChecked   =   -1  'True
      CaptionAlignment=   1
      AutoSize        =   -1  'True
      BackStyle       =   0
      ButtonCaption1  =   "NUEVO RECIVIDO"
      ButtonDescription1=   "INGRESAR LLEGADA DE REPARACION"
      ButtonKey1      =   "RECIVIDO"
      ButtonPicture1  =   "REGISTRO_DE_LLEGADAS.frx":0937
      ButtonToolTipText1=   "add"
      ButtonCaption2  =   "EDITAR RECIVIDO"
      ButtonDescription2=   "EDITAR ARTICULO RECIVIDO"
      ButtonKey2      =   "EDITAR"
      ButtonPicture2  =   "REGISTRO_DE_LLEGADAS.frx":1589
      ButtonToolTipText2=   "app_preferences"
      ButtonCaption3  =   "ELIMINAR "
      ButtonDescription3=   "ELIMINAR REGISTRO"
      ButtonKey3      =   "ELIMINARR"
      ButtonPicture3  =   "REGISTRO_DE_LLEGADAS.frx":21DB
      ButtonToolTipText3=   "delete"
      ButtonCaption4  =   "IMPRIMIR"
      ButtonDescription4=   "IMPRIMIR COMPROBANTE"
      ButtonKey4      =   "IMP_COMPROBANTE"
      ButtonPicture4  =   "REGISTRO_DE_LLEGADAS.frx":2E2D
      ButtonToolTipText4=   "print"
      ButtonCaption5  =   "GUARDAR"
      ButtonDescription5=   "GUARDAR"
      ButtonKey5      =   "GUARDAR_ESTA"
      ButtonPicture5  =   "REGISTRO_DE_LLEGADAS.frx":3A7F
      ButtonToolTipText5=   "save"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "REGISTRO_DE_LLEGADAS.frx":46D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "REGISTRO_DE_LLEGADAS.frx":4FAB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H000000FF&
      Caption         =   "CAMBIAR ESTADO?¿"
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Frame Frame4 
      Caption         =   "CAMBIAR ESTADO"
      Height          =   1455
      Left            =   15000
      TabIndex        =   11
      Top             =   2880
      Width           =   4695
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "REGISTRO_DE_LLEGADAS.frx":5885
         Left            =   120
         List            =   "REGISTRO_DE_LLEGADAS.frx":588F
         TabIndex        =   12
         Text            =   "EN PROCESO"
         Top             =   840
         Width           =   4455
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "REGISTRO_DE_LLEGADAS.frx":58A8
      TabIndex        =   10
      Top             =   4080
      Width           =   8055
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "REGISTRO_DE_LLEGADAS.frx":5960
      TabIndex        =   9
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Frame Frame3 
      Caption         =   "DATOS Y APORTES AL DETALLE"
      Height          =   2175
      Left            =   10680
      TabIndex        =   6
      Top             =   7080
      Width           =   8895
      Begin VB.TextBox Text4 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Krona One"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   4680
         TabIndex        =   26
         Text            =   "DESCRIPCION DEL ARTICULO"
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Krona One"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Text            =   "DATOS DE LA FALLA"
         Top             =   840
         Width           =   3975
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
         Height          =   855
         Left            =   4680
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel COD_PRODUCTO_LAB 
         Height          =   375
         Left            =   3000
         OleObjectBlob   =   "REGISTRO_DE_LLEGADAS.frx":59F8
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "REGISTRO_DE_LLEGADAS.frx":5A6E
         TabIndex        =   8
         Top             =   360
         Width           =   2415
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
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "BUSQUEDA PERSONALIZADA"
      Height          =   1335
      Left            =   15000
      TabIndex        =   1
      Top             =   1560
      Width           =   4695
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
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   2415
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
         ItemData        =   "REGISTRO_DE_LLEGADAS.frx":5AE0
         Left            =   2040
         List            =   "REGISTRO_DE_LLEGADAS.frx":5AEA
         TabIndex        =   2
         Text            =   "DNI"
         Top             =   720
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   615
         Left            =   120
         OleObjectBlob   =   "REGISTRO_DE_LLEGADAS.frx":5AFD
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "REGISTRO_DE_LLEGADAS.frx":5B77
         TabIndex        =   27
         Top             =   840
         Width           =   735
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
         Left            =   2040
         TabIndex        =   31
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox Combo5 
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
         ItemData        =   "REGISTRO_DE_LLEGADAS.frx":5BD9
         Left            =   2040
         List            =   "REGISTRO_DE_LLEGADAS.frx":5BE3
         TabIndex        =   32
         Text            =   "ARTICULO"
         Top             =   720
         Width           =   2415
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "REGISTRO_DE_LLEGADAS.frx":5BF8
      Top             =   840
   End
   Begin MSComctlLib.ListView LISTA_CLIENTE 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   3836
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CLIENTE"
         Object.Width           =   8821
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "APELLIDO"
         Object.Width           =   5293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DNI"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "TELEFONO"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID_LLEGADA"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSAdodcLib.Adodc CONEXION 
      Height          =   375
      Left            =   13560
      Top             =   8160
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
   Begin MSComctlLib.ListView LISTA_ARTICULOS 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   19275
      _ExtentX        =   33999
      _ExtentY        =   3625
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
      ForeColor       =   -2147483647
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ARTICULO"
         Object.Width           =   7939
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MARCA"
         Object.Width           =   4411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ESTADO"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "FECHA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "N° DE ARTICULO"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "GARANTIA"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "PROVEEDOR"
         Object.Width           =   5292
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "REGISTRO_DE_LLEGADAS.frx":5E2C
      TabIndex        =   17
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Menu MN_VOLVER 
      Caption         =   "VOLVER"
   End
   Begin VB.Menu MN_SALIR 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "REGISTRO_DE_LLEGADAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public li As ListItem
Public SELECCIONAR, AGREGAR_REPUESTO As Integer
Public SELECCIONARARTICULO As String
Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
Select Case ButtonIndex
Case 1
    RECEPCION.Show
    REGISTRO_DE_LLEGADAS.Hide
Unload Me
Case 2
    CLIENTE = SELECCIONAR
    ARTICULO = SELECCIONARARTICULO
        If CLIENTE <> 0 And ARTICULO <> 0 Then
            EDITAR = True
            RECEPCION.Show
            REGISTRO_DE_LLEGADAS.Hide
            Unload Me
        Else
            MsgBox ("DEBE SELECCIONAR UN CLIENTE Y SU ARTICULO")
        End If
Case 3
Case 4
    compro = SELECCIONAR
        If compro <> 0 Then
            comprobante.Show
            REGISTRO_DE_LLEGADAS.Hide
            Unload Me
        Else
            MsgBox ("NO AHY CLIENTE SELECCIONADO")
        End If
Case 5
If SELECCIONARARTICULO <> 0 Then
    LISTA_ARTICULOS.ListItems.Clear
    CONEXION.RecordSource = "ARTICULO_RECIBIDO"
    CONEXION.Refresh
    CONEXION.Recordset.MoveFirst
        While CONEXION.Recordset.EOF = False
            If Val(SELECCIONARARTICULO) = CONEXION.Recordset(3) Then
                CONEXION.Recordset("ESTADO") = Combo2.Text
                CONEXION.Recordset.Update
                CONEXION.Refresh
                MsgBox ("LOS DATOS SE GUARDARON CORRECTAMENTE")
                Exit Sub
            End If
            CONEXION.Recordset.MoveNext
        Wend
    CONEXION.Recordset.Close
    Else
    MsgBox ("DEBE SELECCIONAR UN ARTICULO PARA CAMBIAR SU ESTADO")
    End If
End Select
End Sub
Private Sub ChameleonBtn1_Click()
Dim REPUESTO As String
Dim PRECIO As Double
Dim proveedor As String
If AGREGAR_REPUESTO = 0 Then
    MsgBox ("DEBE SELECCIONAR UN ARTICULO PARA AGREGAR SU REPUESTO")
    Exit Sub
End If
If AGREGAR_REPUESTO <> 0 Then
If Option1.Value = True Then
 CONEXION.RecordSource = "PRODUCTOS"
 CONEXION.Refresh
 CONEXION.Recordset.MoveFirst
 Do While (CONEXION.Recordset.EOF = False)
    If List1.Text = CONEXION.Recordset(1) Then
        REPUESTO = List1.Text
        PRECIO = CONEXION.Recordset("PrecioDeVenta")
        proveedor = CONEXION.Recordset("PROVEEDOR")
        Exit Do
    End If
    CONEXION.Recordset.MoveNext
 Loop
CONEXION.RecordSource = "REPUESTOS"
CONEXION.Refresh
CONEXION.Recordset.AddNew
CONEXION.Recordset("RESPUESTO") = REPUESTO
CONEXION.Recordset("PRECIO") = PRECIO
CONEXION.Recordset("PROVEEDOR") = proveedor
CONEXION.Recordset("COD_ARTICULO") = AGREGAR_REPUESTO
CONEXION.Recordset.Update
CONEXION.Refresh
'CONEXION.Recordset.Close
CARGARLISTAS_ARTICULOS
End If
If Option2.Value = True Then
REPUESTO = InputBox("INGRESE EL NOMBRE DEL REPUESTO")
PRECIO = Val(InputBox("INGRESEEL PRECIO DEL REPUESTO"))
If PRECIO = 0 Or REPUESTO = "" Then
   MsgBox ("debe rellenar las ventanas de textos para guardar los cambios")
    Exit Sub
End If
If REPUESTO <> "" And PRECIO <> 0 Then
    CONEXION.RecordSource = "REPUESTOS"
     CONEXION.Refresh
    CONEXION.Recordset.AddNew
    CONEXION.Recordset("RESPUESTO") = REPUESTO
    CONEXION.Recordset("PRECIO") = PRECIO
    CONEXION.Recordset("PROVEEDOR") = Combo3.Text
    CONEXION.Recordset("COD_ARTICULO") = AGREGAR_REPUESTO
    CONEXION.Recordset.Update
    CONEXION.Refresh
    CONEXION.Recordset.Close
End If
CARGARLISTAS_ARTICULOS
    End If
End If
AGREGAR_REPUESTO = 0
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check1.BackColor = &HFF00&
    Combo2.Enabled = True
Else
    Combo2.Enabled = False
    Check1.BackColor = &HFF&
End If
End Sub
Private Sub Combo4_Click()
If Combo4.Text = "INGRESAR REPUESTO" Then
    Frame2.Visible = True
    List1.Visible = False
End If
If Combo4.Text = "CANCELAR" Then
    Frame2.Visible = False
    Combo3.Visible = False
    ChameleonBtn1.Visible = False
End If
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin1.ApplySkin REGISTRO_DE_LLEGADAS.hWnd
CARGARLISTA
Combo2.Enabled = False
Text2.Enabled = False
Text1.Enabled = False
Frame2.Visible = False
Combo3.Visible = False
ChameleonBtn1.Visible = False
Text9.Visible = False
Combo1.Visible = False
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "PROVEEDOR"
CONEXION.Refresh
CONEXION.Recordset.MoveFirst
While UCase(CONEXION.Recordset.EOF = False)
    Combo3.AddItem (CONEXION.Recordset("NOMBRE"))
    CONEXION.Recordset.MoveNext
Wend
CONEXION.RecordSource = "PRODUCTOS"
CONEXION.Refresh
CONEXION.Recordset.MoveFirst
While UCase(CONEXION.Recordset.EOF = False)
    If CONEXION.Recordset("TIPO") = "REPUESTO" Or CONEXION.Recordset("TIPO") = "PRODUCTO-REPUESTO" Then
        List1.AddItem (CONEXION.Recordset("PRODUCTO"))
    End If
    CONEXION.Recordset.MoveNext
Wend
End Sub
Private Sub LISTA_ARTICULOS_Click()
Dim BUSCA_REPU As Integer
If Not LISTA_ARTICULOS.SelectedItem Is Nothing Then
    SELECCIONARARTICULO = LISTA_ARTICULOS.SelectedItem.SubItems(4)
    COD_PRODUCTO_LAB.Caption = "N°" & LISTA_ARTICULOS.SelectedItem.SubItems(4)
    CONEXION.RecordSource = "ARTICULO_RECIBIDO"
    CONEXION.Refresh
    If SELECCIONAR <> 0 And SELECCIONARARTICULO <> "" Then
        'ChameleonBtn2.BackColor = &HC000& CABIAR COLOR
    End If

    CONEXION.Recordset.MoveFirst
   Do While CONEXION.Recordset.EOF = False
                 If Val(SELECCIONARARTICULO) = CONEXION.Recordset(3) Then
                     Text2.Text = CONEXION.Recordset("APORTE")
                     Text1.Text = CONEXION.Recordset("DESCRIPCION")
                        AGREGAR_REPUESTO = CONEXION.Recordset(0)
                       
                      Exit Do
                End If
    CONEXION.Recordset.MoveNext
    Loop
    If LISTA_ARTICULOS.SelectedItem.SubItems(2) = "ENTREGADO" Then
    Frame4.Enabled = False
    Check1.Enabled = False
    Check1.Value = 0
    Else
    Frame4.Enabled = True
    Check1.Enabled = True
    Check1.Value = 0
    End If
    CONEXION.RecordSource = "REPUESTOS"
    CONEXION.Refresh
    LISTA_RESPUESTOS.ListItems.Clear
    CONEXION.Recordset.MoveFirst
    While CONEXION.Recordset.EOF = False
        If AGREGAR_REPUESTO = CONEXION.Recordset("COD_ARTICULO") Then
             Set li = LISTA_RESPUESTOS.ListItems.Add(, , CONEXION.Recordset("RESPUESTO"), , 1)
                 li.ListSubItems.Add , , "$" & CONEXION.Recordset("PRECIO")
                 li.ListSubItems.Add , , CONEXION.Recordset("PROVEEDOR")
        End If
    CONEXION.Recordset.MoveNext
    Wend
End If
CONEXION.Recordset.Close
End Sub
Private Sub LISTA_CLIENTE_Click()
If Not LISTA_CLIENTE.SelectedItem Is Nothing Then
SELECCIONAR = LISTA_CLIENTE.SelectedItem.SubItems(4)
compro = LISTA_CLIENTE.SelectedItem.SubItems(4)
LISTA_ARTICULOS.ListItems.Clear
LISTA_RESPUESTOS.ListItems.Clear
CONEXION.RecordSource = "C1"
CONEXION.Refresh
CONEXION.Recordset.MoveFirst
While CONEXION.Recordset.EOF = False
If Val(SELECCIONAR) = CONEXION.Recordset("NRO_CLIENTE") Then
        Set li = LISTA_ARTICULOS.ListItems.Add(, , CONEXION.Recordset("ARTICULO"), , 2)
            li.ListSubItems.Add , , CONEXION.Recordset("MARCAS.MARCA")
            li.ListSubItems.Add , , CONEXION.Recordset("ESTADO")
            li.ListSubItems.Add , , CONEXION.Recordset("FECHA")
            li.ListSubItems.Add , , CONEXION.Recordset(3)
            li.ListSubItems.Add , , CONEXION.Recordset("GARANTIA")
            li.ListSubItems.Add , , CONEXION.Recordset("NOMBRE")
            End If
    CONEXION.Recordset.MoveNext
Wend
Text2.Text = ""
SELECCIONARARTICULO = 0
If SELECCIONARARTICULO = 0 Then
    'ChameleonBtn2.BackColor = &H808080 CAMBIAR COLOR DE ALGO
End If
End If
End Sub
Private Sub BUSQUEDA()
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
           Set li = LISTA_CLIENTE.ListItems.Add(, , CONEXION.Recordset("NOMBRE"), , 1)
            li.ListSubItems.Add , , CONEXION.Recordset("APELLIDO")
            li.ListSubItems.Add , , CONEXION.Recordset("DNI")
            li.ListSubItems.Add , , CONEXION.Recordset("TELEFONO")
            li.ListSubItems.Add , , CONEXION.Recordset("ID")
        End If
        cantaux = 1
    End If
   CONEXION.Recordset.MoveNext
Next i
CONEXION.Recordset.Close
End Sub
Private Sub MN_SALIR_Click()
End
End Sub
Private Sub MN_VOLVER_Click()
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
    REGISTRO_DE_LLEGADAS.Hide
    Unload Me
End Sub
Private Sub Option1_Click()
If Option1.Value = True Then
List1.Visible = True
ChameleonBtn1.Visible = True
Combo3.Visible = False
End If
End Sub
Private Sub Option2_Click()
If Option2.Value = True Then
List1.Visible = False
Combo3.Visible = True
ChameleonBtn1.Visible = True
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
    Text9.Visible = True
    Combo1.Visible = True
    Text9.ZOrder (0)
    Combo1.ZOrder (0)
    SkinLabel6.Caption = "BUSCAR CLIENTE"
Else
    Text9.Visible = False
    Combo1.Visible = False
    Text9.ZOrder (1)
    Combo1.ZOrder (1)
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
Text5.Visible = True
Combo5.Visible = True
Text5.ZOrder (0)
Combo5.ZOrder (0)
   SkinLabel6.Caption = "BUSCAR ARTICULO"
Else
Text5.Visible = False
Combo5.Visible = False
Combo5.ZOrder (1)
Text5.ZOrder (1)
End If
End Sub

Private Sub Text5_Change()
If Text5.Text <> "" Then
BUSQUEDA_ARTICULO
Else
LISTA_ARTICULOS.ListItems.Clear
End If
End Sub
Private Sub Text9_Change()
If Text9.Text <> "" Then
BUSQUEDA
Else
CARGARLISTA
End If
End Sub
Private Sub CARGARLISTA()
Dim SQL As String
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "Cliente"
CONEXION.Refresh
LISTA_CLIENTE.ListItems.Clear
If CONEXION.Recordset.EOF = False Then
CONEXION.Recordset.MoveFirst
While CONEXION.Recordset.EOF = False
        Set li = LISTA_CLIENTE.ListItems.Add(, , CONEXION.Recordset("NOMBRE"), , 1)
            li.ListSubItems.Add , , CONEXION.Recordset("APELLIDO")
            li.ListSubItems.Add , , CONEXION.Recordset("DNI")
            li.ListSubItems.Add , , CONEXION.Recordset("TELEFONO")
            li.ListSubItems.Add , , CONEXION.Recordset("ID")
    CONEXION.Recordset.MoveNext
Wend
CONEXION.Recordset.Close

End If
End Sub
Sub CARGARLISTAS_ARTICULOS()
    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "ULTIMA_REPUESTO"
    CONEXION.Refresh
    LISTA_RESPUESTOS.ListItems.Clear
    CONEXION.Recordset.MoveFirst
    While CONEXION.Recordset.EOF = False
        If AGREGAR_REPUESTO = CONEXION.Recordset(0) Then
             Set li = LISTA_RESPUESTOS.ListItems.Add(, , CONEXION.Recordset("RESPUESTO"), , 1)
                 li.ListSubItems.Add , , "$" & CONEXION.Recordset("PRECIO")
                 li.ListSubItems.Add , , CONEXION.Recordset("REPUESTOS.PROVEEDOR")
        End If
    CONEXION.Recordset.MoveNext
    Wend
End Sub
Sub BUSQUEDA_ARTICULO()
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "ARTICULO_RECIBIDO_MARCA_PROVEEDOR"
CONEXION.Refresh
Dim cantaux As String
Dim rsaux As String
cantaux = 1
LISTA_ARTICULOS.ListItems.Clear
CONEXION.Recordset.MoveFirst
For i = 0 To (CONEXION.Recordset.RecordCount - 1)
    rsaux = CONEXION.Recordset(Combo5.Text)
    If Mid(rsaux, 1, 1) = Mid(Text5.Text, 1, 1) Then
        For j = 2 To Len(Text5.Text)
            If Mid(CONEXION.Recordset(Combo5.Text), 1, j) = Mid(Text5.Text, 1, j) Then
                cantaux = cantaux + 1
            End If
        Next j
        If cantaux = Len(Text5.Text) Then
           Set li = LISTA_ARTICULOS.ListItems.Add(, , CONEXION.Recordset("ARTICULO"), , 2)
            li.ListSubItems.Add , , CONEXION.Recordset("MARCA")
            li.ListSubItems.Add , , CONEXION.Recordset("ESTADO")
            li.ListSubItems.Add , , CONEXION.Recordset("FECHA")
            li.ListSubItems.Add , , CONEXION.Recordset(2)
            li.ListSubItems.Add , , CONEXION.Recordset("GARANTIA")
            li.ListSubItems.Add , , CONEXION.Recordset("NOMBRE")
        End If
        cantaux = 1
    End If
   CONEXION.Recordset.MoveNext
Next i
CONEXION.Recordset.Close
End Sub
