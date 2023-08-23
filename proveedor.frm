VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#2.0#0"; "Toolbar2.ocx"
Begin VB.Form proveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROVEEDORES"
   ClientHeight    =   8415
   ClientLeft      =   195
   ClientTop       =   615
   ClientWidth     =   15840
   BeginProperty Font 
      Name            =   "Audiowide"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   15840
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   1440
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
            Picture         =   "proveedor.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "proveedor.frx":08DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   375
      Left            =   0
      OleObjectBlob   =   "proveedor.frx":11B4
      TabIndex        =   15
      Top             =   6000
      Width           =   9135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   9840
      OleObjectBlob   =   "proveedor.frx":127C
      TabIndex        =   14
      Top             =   1920
      Width           =   5775
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   3360
      OleObjectBlob   =   "proveedor.frx":1316
      TabIndex        =   13
      Top             =   2040
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   3480
      Left            =   9840
      TabIndex        =   12
      Top             =   2400
      Width           =   5775
   End
   Begin MSComctlLib.ListView LISTA_CLIENTE 
      Height          =   1935
      Left            =   0
      TabIndex        =   11
      Top             =   6360
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° CLIENTE"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   " ARTICULO"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "N°ARTICULO"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "FECHA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ESTADO"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "GARANTIA"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "MARCA"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "ORDENAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   3255
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
         ItemData        =   "proveedor.frx":1398
         Left            =   240
         List            =   "proveedor.frx":13A5
         TabIndex        =   8
         Text            =   "NOMBRE"
         Top             =   360
         Width           =   2775
      End
   End
   Begin AIFCmp1.asxToolbar BAR_PROVEEDOR 
      Height          =   1200
      Left            =   0
      Top             =   0
      Width           =   15450
      _ExtentX        =   27252
      _ExtentY        =   2117
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
      DoubleTopBorder =   -1  'True
      DoubleBottomBorder=   -1  'True
      BackColor       =   0
      HighlightDarkColor=   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Audiowide"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      ButtonCaption1  =   "NUEVO PROVEEDOR"
      ButtonKey1      =   "NUEVO"
      ButtonPicture1  =   "proveedor.frx":13BE
      ButtonToolTipText1=   "documents"
      ButtonCaption2  =   "EDITAR PROVEEDOR"
      ButtonKey2      =   "EDITAR"
      ButtonPicture2  =   "proveedor.frx":2010
      ButtonToolTipText2=   "edit"
      ButtonCaption3  =   "ELIMINAR PROVEEDOR"
      ButtonKey3      =   "ELIMINAR"
      ButtonPicture3  =   "proveedor.frx":2C62
      ButtonToolTipText3=   "cancel"
      ButtonCaption4  =   "BUSCAR"
      ButtonKey4      =   "BUSCARR"
      ButtonPicture4  =   "proveedor.frx":38B4
      ButtonToolTipText4=   "BUSCAR_PROVEEDOR"
      ButtonCaption5  =   "GUARDAR"
      ButtonKey5      =   "GUARDAR_CAMBIOS"
      ButtonPicture5  =   "proveedor.frx":4506
      ButtonToolTipText5=   "GUARDAR"
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   7800
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc CONEXION 
      Height          =   495
      Left            =   600
      Top             =   6360
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
   Begin VB.Frame Frame1 
      Caption         =   "DATOS DEL PROVEEDOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   9840
      TabIndex        =   1
      Top             =   2400
      Width           =   3615
      Begin ACTIVESKINLibCtl.SkinLabel COD_LBL 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "proveedor.frx":5158
         TabIndex        =   10
         Top             =   2520
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "proveedor.frx":51BA
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "proveedor.frx":521E
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "proveedor.frx":529A
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   240
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Audiowide"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   2895
      End
   End
   Begin MSComctlLib.ListView LISTA_PROVEEDOR 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
         Text            =   "PROVEEDOR"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PRECIO DOLAR"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "COD_PROVEEDOR"
         Object.Width           =   4410
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "proveedor.frx":530A
      Top             =   7320
   End
   Begin VB.Menu NM_VOLVER 
      Caption         =   "VOLVER"
   End
End
Attribute VB_Name = "proveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public li As ListItem
Public DOLAR As Double
Public nuevo As Boolean
Private Sub BAR_PROVEEDOR_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
Select Case ButtonIndex
Case 1
    Frame1.Visible = True
    Text1.Text = ""
    Text2.Text = ""
    COD_LBL.Caption = ""
    proveedor.Height = 9500
    nuevo = True
Case 2
    Frame1.Visible = True
    'proveedor.Height = 9500
    List1.Visible = False
Case 3
    MsgBox ("¿ESTA SEGURO QUE DESEA ELIMINAR ?"), vbYesNo
Case 4
    InputBox ("INGRESE EL PROVEEDOR A BUSCAR")
Case 5
    Dim cod_del_proveedor As Integer
    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "SELECT * FROM proveedor ORDER BY COD_PROVEEDOR asc"
    CONEXION.Refresh
If nuevo = True Then
    CONEXION.Recordset.MoveLast
    cod_del_proveedor = CONEXION.Recordset("cod_proveedor")
    CONEXION.Recordset.AddNew
    CONEXION.Recordset("nombre") = Text1.Text
    CONEXION.Recordset("precio_dolar") = Text2.Text
    CONEXION.Recordset("cod_proveedor") = cod_del_proveedor + 1
    CONEXION.Recordset.Update
    CONEXION.Refresh
    LISTA_PROVEEDOR.ListItems.Clear
    CARGAR_LISTA
    List1.Visible = True
    Text1.Text = ""
    Text2.Text = ""
    COD_LBL.Caption = ""
    Frame1.Visible = False
    nuevo = False
    'proveedor.Height = 7500
    Exit Sub
End If
If nuevo = False Then
    CONEXION.Recordset.MoveFirst
    If CONEXION.Recordset.EOF = False Then
        While Not (CONEXION.Recordset.EOF = True)
            If COD_LBL.Caption = CONEXION.Recordset("cod_proveedor") Then
                CONEXION.Recordset("nombre") = Text1.Text
                CONEXION.Recordset("precio_dolar") = Text2.Text
                CONEXION.Recordset.Update
                CONEXION.Refresh
                LISTA_PROVEEDOR.ListItems.Clear
                CARGAR_LISTA
                List1.Visible = True
                Text1.Text = ""
                Text2.Text = ""
                COD_LBL.Caption = ""
                Frame1.Visible = False
                'proveedor.Height = 7500
                Exit Sub
            End If
            CONEXION.Recordset.MoveNext
        Wend
    End If
End If
 CONEXION.Recordset.Close
End Select

End Sub
Private Sub Combo1_Click()
If Combo1.Text = "NOMBRE" Then
    LISTA_PROVEEDOR.ListItems.Clear
    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "SELECT * FROM proveedor ORDER BY NOMBRE DESC"
    CONEXION.Refresh
    CARGAR_LISTA
CONEXION.Recordset.Close
End If
If Combo1.Text = "PRECIO" Then
    LISTA_PROVEEDOR.ListItems.Clear
    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "SELECT * FROM proveedor ORDER BY PRECIO_DOLAR DESC"
    CONEXION.Refresh
    CARGAR_LISTA
CONEXION.Recordset.Close
End If
If Combo1.Text = "COD" Then
    LISTA_PROVEEDOR.ListItems.Clear
    CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
    CONEXION.CursorType = adOpenDynamic
    CONEXION.RecordSource = "SELECT * FROM proveedor ORDER BY COD_PROVEEDOR DESC"
    CONEXION.Refresh
    CARGAR_LISTA
CONEXION.Recordset.Close
End If
End Sub
Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin1.ApplySkin proveedor.hWnd
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "SELECT * FROM proveedor ORDER BY COD_PROVEEDOR ASC"
CONEXION.Refresh
CARGAR_LISTA
'proveedor.Height = 7500
'Frame1.Visible = False
'List1.Enabled = False
'LISTA_CLIENTE.Enabled = False
End Sub
Sub CARGAR_LISTA()
LISTA_PROVEEDOR.ListItems.Clear
If CONEXION.Recordset.EOF = False Then
CONEXION.Recordset.MoveFirst
While CONEXION.Recordset.EOF = False
        Set li = LISTA_PROVEEDOR.ListItems.Add(, , CONEXION.Recordset("NOMBRE"), , 2)
            DOLAR = Val(CONEXION.Recordset("PRECIO_DOLAR"))
            li.ListSubItems.Add , , "$" & Format(DOLAR, "#,###.#0")
            li.ListSubItems.Add , , CONEXION.Recordset("COD_PROVEEDOR")
    CONEXION.Recordset.MoveNext
Wend

End If

End Sub

Private Sub LISTA_PROVEEDOR_Click()
Dim NUMERO_DE_CLEINTE As Integer
Text1.Text = LISTA_PROVEEDOR.SelectedItem
Text2.Text = LISTA_PROVEEDOR.SelectedItem.SubItems(1)
COD_LBL.Caption = LISTA_PROVEEDOR.SelectedItem.SubItems(2)
CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "PRODUCTOS"
CONEXION.Refresh
 List1.Clear
While CONEXION.Recordset.EOF = False
    If CONEXION.Recordset("PROVEEDOR") = LISTA_PROVEEDOR.SelectedItem Then
       If CONEXION.Recordset("TIPO") = "REPUESTO" Or CONEXION.Recordset("TIPO") = "PRODUCTO-REPUESTO" Then
        List1.AddItem (CONEXION.Recordset("PRODUCTO"))
    End If
    End If
    CONEXION.Recordset.MoveNext
Wend
'CONEXION.Recordset.Close
'CONEXION.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
'CONEXION.CursorType = adOpenDynamic
CONEXION.RecordSource = "ARTICULO_Y_MARCA"
CONEXION.Refresh
LISTA_CLIENTE.ListItems.Clear
While CONEXION.Recordset.EOF = False
If CONEXION.Recordset("PROVEEDOR") = LISTA_PROVEEDOR.SelectedItem Then 'And NUMERO_dE_CLIENTE <> CONEXION.Recordset("CLIENTE.ID") Then
  Set li = LISTA_CLIENTE.ListItems.Add(, , CONEXION.Recordset("NRO_CLIENTE"), , 1)
          li.ListSubItems.Add , , CONEXION.Recordset("ARTICULO")
          li.ListSubItems.Add , , CONEXION.Recordset(3)
          li.ListSubItems.Add , , CONEXION.Recordset("FECHA")
          li.ListSubItems.Add , , CONEXION.Recordset("ESTADO")
           li.ListSubItems.Add , , CONEXION.Recordset("GARANTIA")
           li.ListSubItems.Add , , CONEXION.Recordset("MARCAS.MARCA")
  End If
  CONEXION.Recordset.MoveNext
  Wend


End Sub

Private Sub NM_VOLVER_Click()
    Login_y_Menu.Show
    Login_y_Menu.Show
    Login_y_Menu.SkinLabel1.Visible = False
    Login_y_Menu.SkinLabel2.Visible = False
    Login_y_Menu.Text1.Visible = False
    Login_y_Menu.Text2.Visible = False
    Login_y_Menu.ChameleonBtn4.Caption = "Atras"
    Login_y_Menu.ChameleonBtn4.Visible = True
    Login_y_Menu.ChameleonBtn3.Visible = False
    Login_y_Menu.Picture1.Visible = True
    Login_y_Menu.Picture2.Visible = True
    Login_y_Menu.TEXTO.Visible = True
    Login_y_Menu.Width = 14025
    Login_y_Menu.Height = 5685
    Login_y_Menu.ChameleonBtn4.Left = 12000
    Login_y_Menu.ChameleonBtn4.Top = 3360
    Login_y_Menu.CENTRO_BAR.Visible = True
    proveedor.Hide
    Unload Me
End Sub
