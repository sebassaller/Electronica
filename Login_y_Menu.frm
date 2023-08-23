VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#2.0#0"; "Toolbar2.ocx"
Begin VB.Form Login_y_Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INICIO DE SECCION"
   ClientHeight    =   3300
   ClientLeft      =   1080
   ClientTop       =   1425
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6465
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   2640
      Width           =   495
      Begin VB.Image Image2 
         Height          =   480
         Left            =   0
         Picture         =   "Login_y_Menu.frx":0000
         Top             =   0
         Width           =   480
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel TEXTO 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "Login_y_Menu.frx":08CA
      TabIndex        =   7
      Top             =   4200
      Width           =   8775
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn4 
      Height          =   1095
      Left            =   4440
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      BTYPE           =   14
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Audiowide"
         Size            =   9.75
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
      FCOLO           =   16777152
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Login_y_Menu.frx":0989
      PICN            =   "Login_y_Menu.frx":09A5
      PICH            =   "Login_y_Menu.frx":127F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
      Height          =   1095
      Left            =   4440
      TabIndex        =   4
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      BTYPE           =   14
      TX              =   "ENTRAR "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Audiowide"
         Size            =   9.75
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
      FCOLO           =   16777152
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Login_y_Menu.frx":1B59
      PICN            =   "Login_y_Menu.frx":1B75
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2040
      OleObjectBlob   =   "Login_y_Menu.frx":244F
      Top             =   1320
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   615
      Left            =   840
      OleObjectBlob   =   "Login_y_Menu.frx":1EA9E
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9840
      Top             =   5760
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
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   615
      Left            =   360
      OleObjectBlob   =   "Login_y_Menu.frx":1EB06
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   3480
      Width           =   735
      Begin VB.Image Image1 
         Height          =   975
         Left            =   0
         Picture         =   "Login_y_Menu.frx":1EB74
         Stretch         =   -1  'True
         Top             =   0
         Width           =   675
      End
   End
   Begin AIFCmp1.asxToolbar CENTRO_BAR 
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   2778
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Audiowide"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   -2147483641
      ButtonGap       =   10
      BorderStyle     =   1
      BackColor       =   -2147483645
      HighlightColor  =   -2147483648
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
      ButtonCount     =   4
      ButtonCaption1  =   "RECEPCION DE ARTICULOS"
      ButtonKey1      =   "recepcion"
      ButtonPicture1  =   "Login_y_Menu.frx":219C2
      ButtonToolTipText1=   "RECEPCION DE MERCADERIA ARRIBADA"
      ButtonCaption2  =   "PRODUCTOS"
      ButtonKey2      =   "productos"
      ButtonPicture2  =   "Login_y_Menu.frx":22614
      ButtonToolTipText2=   "PRODUCTOS EN ALMACEN"
      ButtonCaption3  =   "REPOCICION"
      ButtonKey3      =   "repocicion"
      ButtonPicture3  =   "Login_y_Menu.frx":23266
      ButtonToolTipText3=   "MANEJO DE STOCK"
      ButtonCaption4  =   "PROVEEDOR"
      ButtonKey4      =   "proveedor"
      ButtonPicture4  =   "Login_y_Menu.frx":23EB8
      ButtonToolTipText4=   "PROVEEDORES"
   End
End
Attribute VB_Name = "Login_y_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Private WithEvents RS As ADODB.Recordset
Attribute RS.VB_VarHelpID = -1
'--------------------------------------------------------------------------------------------------------
Private Sub Mensaje_de_error()
MsgBox ("Usuario o Contraseña incorrecta")
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub CENTRO_BAR_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
Select Case ButtonIndex
Case 1
REGISTRO_DE_LLEGADAS.Show
Login_y_Menu.Hide
Unload Me
Case 2
Productos_y_Reposicion.Show
Productos_y_Reposicion.Command5.Visible = False
Productos_y_Reposicion.Command7.Visible = False
Productos_y_Reposicion.Frame4.Height = 6015
Productos_y_Reposicion.ListView1.Height = 5200
Unload Me
Case 3
Productos_y_Reposicion.Frame2.Visible = False
Productos_y_Reposicion.Frame7.Visible = True
Productos_y_Reposicion.Command6.Top = 6000
Productos_y_Reposicion.Show
Productos_y_Reposicion.Frame4.Visible = False
Unload Me
Case 4
proveedor.Show
Login_y_Menu.Hide
Unload Me
End Select
End Sub



Private Sub ChameleonBtn3_Click()
RS.MoveFirst
RS.Find "Usuario = '" & Text1.Text & "'", , , 1
If RS.BOF = False And RS.EOF = False Then
    If RS.Fields("Contraseña") = Text2.Text Then
        Text1.Text = ""
        Text2.Text = ""
        SkinLabel1.Visible = False
        SkinLabel2.Visible = False
        Text1.Visible = False
        Text2.Visible = False
        ChameleonBtn3.Visible = False
        ChameleonBtn4.Caption = "Atras"
        Picture1.Visible = True
        Picture2.Visible = True
        CENTRO_BAR.Visible = True
        Login_y_Menu.Caption = "MENU PREDETERMINADO"
        TEXTO.Visible = True
        Login_y_Menu.Width = 14655
        Login_y_Menu.Height = 5685
        ChameleonBtn4.Left = 12000
        ChameleonBtn4.Top = 3360
    Else
        Mensaje_de_error
    End If
Else
    Mensaje_de_error
End If

End Sub

'--------------------------------------------------------------------------------------------------------
'---------------------------------SUBS-PROPIAS-----------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
'If Command3.Visible = True Then
 '   SkinLabel1.Visible = True
 '   SkinLabel2.Visible = True
 '   Text1.Visible = True
 '   Text2.Visible = True
 '   Command1.Visible = True
 '   Command2.Caption = "Salir"
 '   Command3.Visible = False
 '   Command4.Visible = False
    
'Else
   End
'End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub ChameleonBtn4_Click()
End
End Sub


Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin1.ApplySkin Me.hWnd

Set RS = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
RS.Source = "Usuarios"
RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
RS.Open "select * from Usuarios", cn

RS.MoveFirst
Unload Productos_y_Reposicion


Picture2.Visible = False
Picture1.Visible = False
TEXTO.Visible = False
CENTRO_BAR.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
End Sub


Private Sub Image1_Click()
usuario.Show
Login_y_Menu.Hide
Unload Me
End Sub

Private Sub Image2_Click()
REGISTROS_EN_EXEL.Show
Login_y_Menu.Hide
Unload Me
End Sub
