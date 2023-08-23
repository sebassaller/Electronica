VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form comprobante 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   13335
   StartUpPosition =   3  'Windows Default
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
      Height          =   855
      Left            =   11040
      TabIndex        =   16
      Top             =   6240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "IMPRIMIR"
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
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "comprobante.frx":0000
      PICN            =   "comprobante.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog IMPRIMIR 
      Left            =   4920
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "comprobante.frx":08F6
      TabIndex        =   13
      Top             =   3120
      Width           =   5055
   End
   Begin MSAdodcLib.Adodc CONEXION_COMPRO 
      Height          =   375
      Left            =   3120
      Top             =   7080
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
      Caption         =   "CONEXION_COMPRO"
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
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
      Height          =   855
      Left            =   5760
      TabIndex        =   3
      Top             =   6240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "ABRIR WORD"
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
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "comprobante.frx":0988
      PICN            =   "comprobante.frx":09A4
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
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "VOLVER"
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
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "comprobante.frx":127E
      PICN            =   "comprobante.frx":129A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView LISTA_COMPROBANTE 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   " DETALLE"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MARCA"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FECHA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "GARANTIA"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATOS DEL CLIENTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   13095
      Begin ACTIVESKINLibCtl.SkinLabel FECHA_LB 
         Height          =   375
         Left            =   11040
         OleObjectBlob   =   "comprobante.frx":1B74
         TabIndex        =   22
         Top             =   960
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   9840
         OleObjectBlob   =   "comprobante.frx":1BF0
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel PROVEEDOR_LB 
         Height          =   375
         Left            =   1680
         OleObjectBlob   =   "comprobante.frx":1C56
         TabIndex        =   20
         Top             =   1440
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel ESTADO_LB 
         Height          =   375
         Left            =   5760
         OleObjectBlob   =   "comprobante.frx":1CEA
         TabIndex        =   19
         Top             =   1560
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "comprobante.frx":1D8C
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "comprobante.frx":1DF4
         TabIndex        =   17
         Top             =   1560
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   9000
         OleObjectBlob   =   "comprobante.frx":1E62
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel id_lb 
         Height          =   375
         Left            =   11040
         OleObjectBlob   =   "comprobante.frx":1ED8
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "comprobante.frx":1F4C
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "comprobante.frx":1FB8
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "comprobante.frx":201A
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "comprobante.frx":2086
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel tel_lb 
         Height          =   375
         Left            =   5760
         OleObjectBlob   =   "comprobante.frx":20EE
         TabIndex        =   7
         Top             =   360
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel apellido_lb 
         Height          =   375
         Left            =   1680
         OleObjectBlob   =   "comprobante.frx":2192
         TabIndex        =   6
         Top             =   840
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel dni_lb 
         Height          =   375
         Left            =   5760
         OleObjectBlob   =   "comprobante.frx":2226
         TabIndex        =   5
         Top             =   840
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel nombre_lb 
         Height          =   375
         Left            =   1680
         OleObjectBlob   =   "comprobante.frx":22C4
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "comprobante.frx":235A
      Top             =   7560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "comprobante.frx":258E
      TabIndex        =   14
      Top             =   240
      Width           =   5175
   End
   Begin VB.Menu MN_AYUDA 
      Caption         =   "AYUDA"
   End
   Begin VB.Menu MN_SALIR 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "comprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Proceso As Double
Public li As ListItem
Private Sub ChameleonBtn1_Click()
REGISTRO_DE_LLEGADAS.Show
comprobante.Hide
Unload Me
End Sub

Private Sub ChameleonBtn2_Click()
'Dim WordDoc As Word.Document
'Dim WordApp As Word.Application
'Set WordApp = New Word.Application
'WordApp.Visible = True

'With WordApp
'WordApp.Selection.TypeText "hola mundo"
' WordApp.Documents.Open "C:\Users\saller\Desktop\Programa para controlar stock\comprobante.docx"
' WordApp.ActiveDocument.PrintOut
'  WordApp.Quit
 'End With
'
'IMPRIMIR.ShowPrinter

''anda en exel_____________________________________________________________________________________
    'Dim objExcel As New Excel.Application
    'Dim bkWorkBook As Workbook
    'Dim shWorkSheet As Worksheet
    'Dim i As Integer
    'Dim j As Integer

    'Set objExcel = New Excel.Application
    'Set bkWorkBook = objExcel.Workbooks.Add
    'Set shWorkSheet = bkWorkBook.ActiveSheet
    'For i = 1 To LISTA_COMPROBANTE.ColumnHeaders.Count
     '   shWorkSheet.Cells(1, Chr(64 + i)) = LISTA_COMPROBANTE.ColumnHeaders(i)
    'Next
    'For i = 1 To LISTA_COMPROBANTE.ListItems.Count
    '    shWorkSheet.Cells(i + 2, "A") = LISTA_COMPROBANTE.ListItems(i).Text
    '    For j = 2 To LISTA_COMPROBANTE.ColumnHeaders.Count
    '        shWorkSheet.Cells(i + 2, Chr(64 + j)) = LISTA_COMPROBANTE.ListItems(i).SubItems(j - 1)
    '    Next
    'Next
    
   ' objExcel.Visible = True
    'layervisiblelity 'set visible lity
    'Set objExcel = Nothing
    'Set bkWorkBook = Nothing
    'Set shWorkSheet = Nothing
    ''____________________________________________________________________________________________________
    
    
    
Dim objword As Word.Application
Dim odoc As Word.Document
Dim otable As Word.Table
Dim i As Integer
Dim j As Integer







Set objword = New Word.Application
Set odoc = objword.Documents.Add("C:\Users\saller\Desktop\Programa para controlar stock\comprobante.docx")






odoc.Activate
odoc.Bookmarks.Item("Nombre").Range.Text = nombre_lb.Caption
odoc.Bookmarks.Item("Apellido").Range.Text = apellido_lb.Caption
odoc.Bookmarks.Item("dni").Range.Text = dni_lb.Caption
odoc.Bookmarks.Item("telefono").Range.Text = tel_lb.Caption
odoc.Bookmarks.Item("PROVEEDOR").Range.Text = PROVEEDOR_LB.Caption
odoc.Bookmarks.Item("ESTADO").Range.Text = ESTADO_LB.Caption
odoc.Bookmarks.Item("NRO_RECIVO").Range.Text = id_lb.Caption
odoc.Bookmarks.Item("FECHA").Range.Text = FECHA_LB.Caption

objword.Selection.Move Unit:=wdStory
objword.Selection.TypeParagraph
objword.Selection.TypeParagraph

'wdAutoFitWindow

 
Set otable = odoc.Tables.Add(odoc.Paragraphs(1).Range, LISTA_COMPROBANTE.ListItems.Count + 1, LISTA_COMPROBANTE.ColumnHeaders.Count, wdWord9TableBehavior, wdAutoFitWindow)
For i = 1 To LISTA_COMPROBANTE.ColumnHeaders.Count
    otable.Rows(1).Cells(i).Range.Text = LISTA_COMPROBANTE.ColumnHeaders(i)
Next
For i = 1 To LISTA_COMPROBANTE.ListItems.Count
    otable.Columns(1).Cells(i + 1).Range.Text = LISTA_COMPROBANTE.ListItems(i).Text
For j = 2 To LISTA_COMPROBANTE.ColumnHeaders.Count
    otable.Rows(i + 1).Cells(j).Range.Text = LISTA_COMPROBANTE.ListItems(i).SubItems(j - 1)
Next
Next

objword.Visible = False
objword.ActiveDocument.PrintOut
'odoc.Application.Quit
Set objword = Nothing
Set otable = Nothing
Set odoc = Nothing
 


Proceso = Shell("taskkill /IM WINWORD.EXE /F") 'mata el proceso de word



'Set odoc.Close(odoc.Save, objword.Documents, otable.Tables) = Nothingg
'Set odoc.Close (wdDoNotSaveChanges,wdOriginalDocumentFormat)



'(wdDoNotSaveChanges)
'wdDoNotSaveChanges
'SaveChanges
End Sub

Private Sub ChameleonBtn3_Click()
Dim objword As Word.Application
Dim odoc As Word.Document
Dim otable As Word.Table
Dim i As Integer
Dim j As Integer




Set objword = New Word.Application
Set odoc = objword.Documents.Add("C:\Users\saller\Desktop\Programa para controlar stock\RECIBO.docx")






odoc.Activate
odoc.Bookmarks.Item("Nombre").Range.Text = nombre_lb.Caption
odoc.Bookmarks.Item("Apellido").Range.Text = apellido_lb.Caption
odoc.Bookmarks.Item("dni").Range.Text = dni_lb.Caption
odoc.Bookmarks.Item("telefono").Range.Text = tel_lb.Caption
odoc.Bookmarks.Item("PROVEEDOR").Range.Text = PROVEEDOR_LB.Caption
odoc.Bookmarks.Item("ESTADO").Range.Text = ESTADO_LB.Caption
odoc.Bookmarks.Item("NRO_RECIVO").Range.Text = id_lb.Caption
odoc.Bookmarks.Item("FECHA").Range.Text = FECHA_LB.Caption

objword.Selection.Move Unit:=wdStory
objword.Selection.TypeParagraph
objword.Selection.TypeParagraph

'wdAutoFitWindow

 
Set otable = odoc.Tables.Add(odoc.Paragraphs(1).Range, LISTA_COMPROBANTE.ListItems.Count + 1, LISTA_COMPROBANTE.ColumnHeaders.Count, wdWord9TableBehavior, wdAutoFitWindow)
For i = 1 To LISTA_COMPROBANTE.ColumnHeaders.Count
    otable.Rows(1).Cells(i).Range.Text = LISTA_COMPROBANTE.ColumnHeaders(i)
Next
For i = 1 To LISTA_COMPROBANTE.ListItems.Count
    otable.Columns(1).Cells(i + 1).Range.Text = LISTA_COMPROBANTE.ListItems(i).Text
For j = 2 To LISTA_COMPROBANTE.ColumnHeaders.Count
    otable.Rows(i + 1).Cells(j).Range.Text = LISTA_COMPROBANTE.ListItems(i).SubItems(j - 1)
Next
Next

objword.Visible = False
objword.ActiveDocument.PrintOut
'odoc.Application.Quit
Set objword = Nothing
Set otable = Nothing
Set odoc = Nothing
Proceso = Shell("taskkill /IM WINWORD.EXE /F")
    
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin1.ApplySkin comprobante.hWnd

CONEXION_COMPRO.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION_COMPRO.CursorType = adOpenDynamic
CONEXION_COMPRO.RecordSource = "cliente"
CONEXION_COMPRO.Refresh

CONEXION_COMPRO.Recordset.MoveFirst
Do While Not UCase(CONEXION_COMPRO.Recordset.EOF = True)
    If compro = CONEXION_COMPRO.Recordset("id") Then
    id_lb = CONEXION_COMPRO.Recordset("id")
    nombre_lb.Caption = CONEXION_COMPRO.Recordset("nombre")
    apellido_lb.Caption = CONEXION_COMPRO.Recordset("apellido")
    dni_lb.Caption = CONEXION_COMPRO.Recordset("dni")
    tel_lb.Caption = CONEXION_COMPRO.Recordset("telefono")
    
    Exit Do
End If
CONEXION_COMPRO.Recordset.MoveNext
Loop
CONEXION_COMPRO.Recordset.Close


CONEXION_COMPRO.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
CONEXION_COMPRO.CursorType = adOpenKeyset
CONEXION_COMPRO.RecordSource = " articulo_recibido  "
CONEXION_COMPRO.Refresh


If CONEXION_COMPRO.Recordset.EOF = False Then

CONEXION_COMPRO.Recordset.MoveFirst
While Not UCase(CONEXION_COMPRO.Recordset.EOF = True)
If id_lb = CONEXION_COMPRO.Recordset("NRO_CLIENTE") Then
Set li = LISTA_COMPROBANTE.ListItems.Add(, , CONEXION_COMPRO.Recordset("ARTICULO"))
            li.ListSubItems.Add , , CONEXION_COMPRO.Recordset("MARCA")
            li.ListSubItems.Add , , CONEXION_COMPRO.Recordset("FECHA")
            li.ListSubItems.Add , , CONEXION_COMPRO.Recordset("GARANTIA")
            FECHA_LB.Caption = CONEXION_COMPRO.Recordset("FECHA")
            PROVEEDOR_LB.Caption = CONEXION_COMPRO.Recordset("PROVEEDOR")
            ESTADO_LB.Caption = CONEXION_COMPRO.Recordset("ESTADO")
    
End If
CONEXION_COMPRO.Recordset.MoveNext

Wend



End If

CONEXION_COMPRO.Recordset.Close


End Sub

Private Sub MN_SALIR_Click()
End
End Sub
