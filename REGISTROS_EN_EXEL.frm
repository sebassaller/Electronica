VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form REGISTROS_EN_EXEL 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   12420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "SELECCIONAR  ARCHIVO"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   12255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   6360
         OleObjectBlob   =   "REGISTROS_EN_EXEL.frx":0000
         TabIndex        =   9
         Top             =   1080
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "REGISTROS_EN_EXEL.frx":0080
         TabIndex        =   8
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Frame Frame2 
         Caption         =   "RUTA DEL ARCHIVO"
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   2880
         Width           =   11775
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
            Left            =   360
            TabIndex        =   7
            Top             =   360
            Width           =   10575
         End
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H80000001&
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
         Height          =   1365
         Left            =   6240
         TabIndex        =   5
         Top             =   1320
         Width           =   5775
      End
      Begin ChamaleonButton.ChameleonBtn CERRAR 
         Height          =   975
         Left            =   9240
         TabIndex        =   4
         Top             =   4200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
         BTYPE           =   14
         TX              =   "CERRAR_EXEL"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Krona One"
            Size            =   9.75
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
         MICON           =   "REGISTROS_EN_EXEL.frx":0100
         PICN            =   "REGISTROS_EN_EXEL.frx":011C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ABRIR 
         Height          =   975
         Left            =   240
         TabIndex        =   3
         Top             =   4200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1720
         BTYPE           =   14
         TX              =   "ABRIR_EXEL"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Krona One"
            Size            =   9.75
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
         MICON           =   "REGISTROS_EN_EXEL.frx":09F6
         PICN            =   "REGISTROS_EN_EXEL.frx":0A12
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Krona One"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   1515
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   5895
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Krona One"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "REGISTROS_EN_EXEL.frx":12EC
      Top             =   120
   End
End
Attribute VB_Name = "REGISTROS_EN_EXEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RUTA1, RUTA2, RUTA3 As String
Public XL As New Excel.Application 'Crea el objeto excel
Private Sub ABRIR_Click()
XL.Workbooks.Open (RUTA3), , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
XL.Visible = True
XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada.End Sub
End Sub
Private Sub CERRAR_Click()
XL.Quit
Set XL = Nothing
Unload Me
End Sub
Private Sub Dir1_Change()
' -- Cada vez que cambiamos de directorio, le indicamos al
    ' -- control FileListBox que muestre los archivos de ese directorio
    File1.Path = Dir1.Path
    RUTA1 = Dir1.Path
End Sub
Private Sub Drive1_Change()
On Error GoTo error_handler
    ' -- Cada vez que cambiamos de unidad, indicamos al control
    ' -- Dir Que muestre los directorios de esa unidad
    Dir1.Path = Drive1.List(Drive1.ListIndex)
    ' -- Rutina de error en caso de que se seleccione una unidad no disponible
    ' -- O que se produzca cualquier otro tipo de error
    Exit Sub
error_handler:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub File1_Click()
On Error GoTo error_handler
    ' -- Mostramos en la barra de título del formulario el nombre del
    ' -- archivo seleccionado en el control File1
    Me.Caption = "Archivo Actual: " & File1.FileName
    'Image1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
      RUTA2 = File1.FileName
    ' -- Rutina de error en caso de que no se pueda cargar la imagen en el Image
      RUTA3 = RUTA1 & "\" & RUTA2
      Text2.Text = RUTA3
    Exit Sub
error_handler:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin1.ApplySkin REGISTROS_EN_EXEL.hWnd
 ' -- Para indicarle al control File que liste y filtre solo Bmp
    File1.Pattern = "*.xlsx"
End Sub
