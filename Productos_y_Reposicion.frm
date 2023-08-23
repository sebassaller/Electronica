VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Productos_y_Reposicion 
   Caption         =   "Form2"
   ClientHeight    =   10500
   ClientLeft      =   720
   ClientTop       =   300
   ClientWidth     =   15930
   LinkTopic       =   "Form2"
   ScaleHeight     =   10500
   ScaleWidth      =   15930
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   6360
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
            Picture         =   "Productos_y_Reposicion.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "Editar Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10215
      Left            =   720
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   13455
      Begin VB.CommandButton Command11 
         Caption         =   "Editar"
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
         Index           =   6
         Left            =   9480
         TabIndex        =   55
         Top             =   7320
         Width           =   1815
      End
      Begin VB.CommandButton Command12 
         DisabledPicture =   "Productos_y_Reposicion.frx":08DA
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   7440
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         DisabledPicture =   "Productos_y_Reposicion.frx":0C8C
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   5760
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         DisabledPicture =   "Productos_y_Reposicion.frx":103E
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4800
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         DisabledPicture =   "Productos_y_Reposicion.frx":13F0
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3840
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         DisabledPicture =   "Productos_y_Reposicion.frx":17A2
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         DisabledPicture =   "Productos_y_Reposicion.frx":1B54
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Editar"
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
         Index           =   5
         Left            =   9480
         TabIndex        =   43
         Top             =   5640
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Editar"
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
         Index           =   4
         Left            =   9480
         TabIndex        =   42
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Editar"
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
         Index           =   3
         Left            =   9480
         TabIndex        =   41
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Editar"
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
         Index           =   2
         Left            =   9480
         TabIndex        =   40
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton Command12 
         DisabledPicture =   "Productos_y_Reposicion.frx":1F06
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Editar"
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
         Index           =   1
         Left            =   9480
         TabIndex        =   38
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Editar"
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
         Index           =   0
         Left            =   9480
         TabIndex        =   37
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7800
         TabIndex        =   36
         Top             =   8760
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3960
         TabIndex        =   35
         Top             =   8760
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   0
         Left            =   2280
         OleObjectBlob   =   "Productos_y_Reposicion.frx":22B8
         TabIndex        =   30
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   29
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   28
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   27
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   5640
         TabIndex        =   26
         Top             =   5640
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   5640
         TabIndex        =   25
         Top             =   4680
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   5640
         TabIndex        =   24
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   5640
         TabIndex        =   23
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   5640
         TabIndex        =   22
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   5640
         TabIndex        =   21
         Top             =   840
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   1
         Left            =   2280
         OleObjectBlob   =   "Productos_y_Reposicion.frx":2326
         TabIndex        =   31
         Top             =   2040
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   2
         Left            =   2280
         OleObjectBlob   =   "Productos_y_Reposicion.frx":239A
         TabIndex        =   32
         Top             =   3000
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   3
         Left            =   2280
         OleObjectBlob   =   "Productos_y_Reposicion.frx":2402
         TabIndex        =   33
         Top             =   3960
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   4
         Left            =   2280
         OleObjectBlob   =   "Productos_y_Reposicion.frx":246A
         TabIndex        =   34
         Top             =   4920
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   5
         Left            =   2280
         OleObjectBlob   =   "Productos_y_Reposicion.frx":24E8
         TabIndex        =   44
         Top             =   5880
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   6
         Left            =   4800
         OleObjectBlob   =   "Productos_y_Reposicion.frx":2564
         TabIndex        =   45
         Top             =   6600
         Width           =   4215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   7
         Left            =   3000
         OleObjectBlob   =   "Productos_y_Reposicion.frx":25F0
         TabIndex        =   46
         Top             =   7560
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   8
         Left            =   4920
         OleObjectBlob   =   "Productos_y_Reposicion.frx":265A
         TabIndex        =   47
         Top             =   7560
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Index           =   9
         Left            =   6360
         OleObjectBlob   =   "Productos_y_Reposicion.frx":26C0
         TabIndex        =   48
         Top             =   7560
         Width           =   1575
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Agregar Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   3120
      TabIndex        =   61
      Top             =   2280
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox Text7 
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
         Left            =   3000
         TabIndex        =   73
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CommandButton Command16 
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
         Left            =   5640
         TabIndex        =   72
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command15 
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
         Left            =   2400
         TabIndex        =   71
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text4 
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
         Left            =   3000
         TabIndex        =   67
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text5 
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
         Left            =   3000
         TabIndex        =   66
         Top             =   720
         Width           =   2415
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   5760
         TabIndex        =   62
         Top             =   600
         Width           =   3495
         Begin VB.CheckBox Check1 
            Caption         =   "Crear precio de venta por porcentaje"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   240
            TabIndex        =   64
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox Text6 
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
            Left            =   2040
            TabIndex        =   63
            Text            =   "30"
            Top             =   1200
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   375
            Left            =   240
            OleObjectBlob   =   "Productos_y_Reposicion.frx":272C
            TabIndex        =   65
            Top             =   1320
            Width           =   1455
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   480
         OleObjectBlob   =   "Productos_y_Reposicion.frx":279E
         TabIndex        =   68
         Top             =   840
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   375
         Left            =   480
         OleObjectBlob   =   "Productos_y_Reposicion.frx":2806
         TabIndex        =   69
         Top             =   1560
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   480
         OleObjectBlob   =   "Productos_y_Reposicion.frx":2884
         TabIndex        =   70
         Top             =   2280
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Producto seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   8160
      Width           =   11535
      Begin VB.CommandButton Command14 
         Caption         =   "Nuevo Producto "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6000
         TabIndex        =   60
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Editar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cambiar precio "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Borrar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9120
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   480
         OleObjectBlob   =   "Productos_y_Reposicion.frx":2900
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   495
         Left            =   480
         OleObjectBlob   =   "Productos_y_Reposicion.frx":296A
         TabIndex        =   11
         Top             =   480
         Width           =   3735
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
         Height          =   495
         ItemData        =   "Productos_y_Reposicion.frx":29F0
         Left            =   4440
         List            =   "Productos_y_Reposicion.frx":29FA
         TabIndex        =   2
         Text            =   "Producto"
         Top             =   360
         Width           =   2535
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
         Left            =   2640
         TabIndex        =   1
         Top             =   960
         Width           =   4335
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Aumentar o disminuir stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2880
      TabIndex        =   57
      Top             =   6600
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CommandButton Command13 
         Caption         =   "Disminuir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   5400
         TabIndex        =   59
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Aumentar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   1440
         TabIndex        =   58
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Agregar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   56
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Productos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   15735
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   15200
         _ExtentX        =   26802
         _ExtentY        =   5741
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
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Producto"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Stock"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Precio de compra"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Precio de venta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Posicion en deposito"
            Object.Width           =   3881
         EndProperty
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Atras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13320
      TabIndex        =   3
      Top             =   9000
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   12120
      OleObjectBlob   =   "Productos_y_Reposicion.frx":2A0F
      Top             =   9360
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Significado de los colores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8160
      TabIndex        =   8
      Top             =   120
      Width           =   7695
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   600
         OleObjectBlob   =   "Productos_y_Reposicion.frx":BFC14
         TabIndex        =   10
         Top             =   720
         Width           =   4695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Productos a reponer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   15735
      Begin VB.CommandButton Command8 
         Caption         =   "Nuevo producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   13800
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Guardar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   13800
         TabIndex        =   18
         Top             =   2280
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2775
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483643
         BackColor       =   -2147483640
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
            Text            =   "Producto"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Precio de compra"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Precio de venta"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Codigo"
            Object.Width           =   2999
         EndProperty
      End
   End
End
Attribute VB_Name = "Productos_y_Reposicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim RS As New ADODB.Recordset
Public li As ListItem
Dim AumentarDisminuir As String
Dim auxVacio As Boolean
Dim Asciiaux As Integer
Dim textboxaux As String


'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

Private Sub Solo_numeros_y_Un_punto()
If InStr("0123456789.", Chr(Asciiaux)) = 0 And Asciiaux <> 8 Then
    Asciiaux = 0
    MsgBox ("Solo ingresar numeros y un punto")
Else
If Asciiaux <> 46 Then
    Else
        For i = 1 To Len(textboxaux)
            If Mid(textboxaux, i, 1) = "." Then
                MsgBox ("Solo se permite un punto")
                 Asciiaux = 0
            End If
        Next i
    End If
End If
End Sub

Private Sub Solo_Numeros()
If InStr("0123456789", Chr(Asciiaux)) = 0 And Asciiaux <> 8 Then
    Asciiaux = 0
    MsgBox ("Solo ingresar numeros")
End If
End Sub

Private Sub Colocar_Color()
If RS("Stock") <= 3 Then
    li.ForeColor = vbRed
    For i = 1 To 7
        li.ListSubItems(i).ForeColor = vbRed
    Next i
End If
End Sub
Private Sub Recargar_listview1()
If Text1.Text = "" Then
    Colocar_todos_los_productos
Else
    BUSQUEDA
End If
End Sub
Private Sub Buscar_por_Codigo_listview1()
RS.MoveFirst
While RS("Id") <> ListView1.SelectedItem.ListSubItems(6) And RS.EOF = False
    RS.MoveNext
Wend
End Sub
Private Sub Bloquear_Text()
    For i = 0 To 6
        Command11(i).Caption = "Editar"
        Command12(i).DisabledPicture = LoadPicture(App.Path & "\Datos\Candado_Cerrado.jpg")
        Text3(i).Enabled = False
    Next i
        Text3(7).Enabled = False
        Text3(8).Enabled = False
End Sub
Private Sub Verificar_Text()
auxVacio = False
For i = 0 To 8
    If Text3(i).Text = "" Then
        auxVacio = True
    End If
Next i
If auxVacio = True Then
    MsgBox ("Hay uno o mas datos sin completar, por favor complete los datos")
End If
End Sub
Private Sub BUSQUEDA()
Dim cantaux As String
Dim rsaux As String

cantaux = 1
ListView1.ListItems.Clear
RS.MoveFirst

For i = 0 To (RS.RecordCount - 1)
    rsaux = RS(Combo1.Text)
    If RS("Estado") <> "baja" And RS("Estado") <> "x" Then
        If Mid(rsaux, 1, 1) = Mid(Text1.Text, 1, 1) Then
            For j = 2 To Len(Text1.Text)
                If Mid(RS(Combo1.Text), 1, j) = Mid(Text1.Text, 1, j) Then
                    cantaux = cantaux + 1
                End If
            Next j
            If cantaux = Len(Text1.Text) Then
                Set li = ListView1.ListItems.Add(, , RS("Producto"))
                    li.ListSubItems.Add , , RS("Descripcion")
                    li.ListSubItems.Add , , RS("Marca")
                    li.ListSubItems.Add , , RS("Stock")
                    li.ListSubItems.Add , , RS("PrecioDeCompra")
                    li.ListSubItems.Add , , RS("PrecioDeVenta")
                    li.ListSubItems.Add , , RS("Id")
                    li.ListSubItems.Add , , RS("PosicionEnDeposito")
                Colocar_Color
            End If
            cantaux = 1
        End If
    End If
    RS.MoveNext
Next i
End Sub

Private Sub Colocar_todos_los_productos()

ListView1.ListItems.Clear

RS.MoveFirst
While RS.EOF = False
    If RS("Estado") <> "baja" And RS("Estado") <> "x" Then
        Set li = ListView1.ListItems.Add(, , RS("Producto"), , 1)
            li.ListSubItems.Add , , RS("Descripcion")
            li.ListSubItems.Add , , RS("Marca")
            li.ListSubItems.Add , , RS("Stock")
            li.ListSubItems.Add , , RS("PrecioDeCompra")
            li.ListSubItems.Add , , RS("PrecioDeVenta")
            li.ListSubItems.Add , , RS("Id")
            li.ListSubItems.Add , , RS("PosicionEnDeposito")
        Colocar_Color
    End If
    RS.MoveNext
Wend
End Sub
'--------------------------------------------------------------------------------------------------------
'---------------------------------SUBS-PROPIAS-----------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

Private Sub Check1_Click()
If Check1.Value = "0" Then
   Text6.Enabled = False
Else
    Text6.Enabled = True
    Dim aux As Integer
    Text7.Text = Str(Round(((Val(Text6.Text) * 0.01) * Val(Text4.Text)) + Val(Text4.Text), 3))
    aux = Str(Int(Len(Text7.Text)) - 1)
    If Mid(Text7.Text, aux, 1) = "." Then
        Text7.Text = Text7.Text + "0"
    End If
End If
End Sub

Private Sub Combo1_GotFocus()
Command5.Enabled = False
Command7.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()

Frame6.Visible = True
Frame6.Caption = "Editar Producto"
Frame4.Visible = False
Frame1.Visible = False
Frame3.Visible = False
Frame2.Visible = False
Command6.Visible = False
Text3(0).Text = ListView1.SelectedItem
Text3(1).Text = ListView1.SelectedItem.SubItems(1)
Text3(2).Text = ListView1.SelectedItem.SubItems(2)
Text3(3).Text = ListView1.SelectedItem.SubItems(3)
Text3(4).Text = ListView1.SelectedItem.SubItems(4)
Text3(5).Text = ListView1.SelectedItem.SubItems(5)
Text3(6).Text = Mid(ListView1.SelectedItem.SubItems(7), 5, 1)
Text3(7).Text = Mid(ListView1.SelectedItem.SubItems(7), 8, 1)
Text3(8).Text = Mid(ListView1.SelectedItem.SubItems(7), 11, 1)


End Sub

Private Sub Command10_Click()
If Frame6.Caption = "Editar Producto" Then
    Frame1.Visible = True
    Frame4.Visible = True
    Frame3.Visible = True
    Frame2.Visible = True
    Command6.Visible = True
    Frame6.Visible = False
    Bloquear_Text
End If
If Frame6.Caption = "Nuevo Producto" Then
    If Command6.Visible = True Then
        Frame2.Visible = True
    Else
        Frame5.Visible = True
        Command6.Visible = True
        Command5.Visible = True
        Command7.Visible = True
    End If
    Frame3.Visible = True
    Frame1.Visible = True
    Frame6.Visible = False
    Frame4.Visible = True
End If
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command11_Click(Index As Integer)
If Command11(Index).Caption = "Editar" Then
    Command12(Index).DisabledPicture = LoadPicture(App.Path & "\Datos\Candado_Abierto.jpg")
    Text3(Index).Enabled = True
    Text3(Index).SetFocus
    If Index = 6 Then
        Text3(Index + 1).Enabled = True
        Text3(Index + 2).Enabled = True
    End If
    Command11(Index).Caption = "Listo"
Else
    Command12(Index).DisabledPicture = LoadPicture(App.Path & "\Datos\Candado_Cerrado.jpg")
    Text3(Index).Enabled = False
    If Index = 6 Then
        Text3(Index + 1).Enabled = False
        Text3(Index + 2).Enabled = False
    End If
    Command11(Index).Caption = "Editar"
End If
End Sub

Private Sub Command13_Click(Index As Integer)
Frame7.Visible = False
Frame5.Visible = True
Frame4.Visible = True
Command5.Visible = True
Command7.Visible = True
If Index = 0 Then
    AumentarDisminuir = "Aumentar"
Else
    AumentarDisminuir = "Disminuir"
End If
End Sub

Private Sub Command14_Click()

Frame6.Visible = True
Frame6.Caption = "Nuevo Producto"
Frame4.Visible = False
Frame2.Visible = False
Frame5.Visible = False
Frame3.Visible = False
Frame1.Visible = False
Command5.Visible = False
Command7.Visible = False
End Sub

Private Sub Command15_Click()
If Text5.Visible = True Then
    Set li = ListView2.ListItems.Add(, , ListView1.SelectedItem)
        li.ListSubItems.Add , , ListView1.SelectedItem.SubItems(1)
        li.ListSubItems.Add , , ListView1.SelectedItem.SubItems(2)
    If Text5.Text <> "" Then
        li.ListSubItems.Add , , Text5.Text
    Else
        li.ListSubItems.Add , , "1"
    End If
    If Text4.Text <> "" Then
        li.ListSubItems.Add , , Text4.Text
    Else
        li.ListSubItems.Add , , ListView1.SelectedItem.SubItems(4)
    End If
    If Text7.Text <> "" Then
        li.ListSubItems.Add , , Text7.Text
    Else
        li.ListSubItems.Add , , ListView1.SelectedItem.SubItems(5)
    End If
    li.ListSubItems.Add , , ListView1.SelectedItem.SubItems(6)
    li.ListSubItems.Add , , ListView1.SelectedItem.SubItems(7)
    If ListView2.ListItems.Count > 0 Then
        Command9.Enabled = True
    End If
Else
    RS.MoveFirst
    While RS.EOF = False And RS("Id") <> ListView1.SelectedItem.SubItems(6)
        RS.MoveNext
    Wend
    RS.Update
    If Text4.Text <> "" Then
        RS("PrecioDeCompra") = Text4.Text
    End If
    If Text7.Text <> "" Then
        RS("PrecioDeVenta") = Text7.Text
    End If
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
End If
Text4.Text = ""
Text7.Text = ""
Text5.Text = ""
Recargar_listview1
Frame8.Visible = False
Frame4.Visible = True
Frame2.Enabled = True
Frame1.Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
Frame5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command16_Click()
Frame8.Visible = False
Frame4.Visible = True
Frame2.Enabled = True
Frame1.Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
Frame5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command5.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command2_Click()
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False

Dim respuesta As String
If MsgBox("¿Desea borrar el producto " + ListView1.SelectedItem + " de la marca " + ListView1.SelectedItem.ListSubItems(2) + " " + ListView1.SelectedItem.ListSubItems(1) + " ?", vbYesNo, "Confirmacion para Borrar Producto") = vbYes Then
    Buscar_por_Codigo_listview1
    RS("Estado") = "baja"
    RS.Update
    Recargar_listview1
End If
End Sub

Private Sub Command3_Click()

Text5.Visible = False
SkinLabel5.Visible = False
Frame8.Visible = True
Frame4.Visible = False
Frame2.Enabled = False
Frame1.Enabled = False
Frame3.Enabled = False
Command6.Enabled = False
Text4.Text = ListView1.SelectedItem.ListSubItems(4)
Text7.Text = ListView1.SelectedItem.ListSubItems(5)
End Sub

Private Sub Command4_Click()
If Frame6.Caption = "Editar Producto" Then
    Verificar_Text
    If auxVacio = False Then
        RS.MoveFirst
        Buscar_por_Codigo_listview1
        If RS.EOF = False Then
            RS("Producto") = Text3(0).Text
            RS("Descripcion") = Text3(1).Text
            RS("Marca") = Text3(2).Text
            RS("Stock") = Text3(3).Text
            RS("PrecioDeCompra") = Text3(4).Text
            RS("PrecioDeVenta") = Text3(5).Text
            RS("PosicionEnDeposito") = ("Sec." + Text3(6).Text + "/F" + Text3(7).Text + "/C" + Text3(8).Text)
            RS.Update
            
            Recargar_listview1
            
            auxVacio = False
            Frame1.Visible = True
            Frame4.Visible = True
            Frame3.Visible = True
            Frame2.Visible = True
            Command6.Visible = True
            Frame6.Visible = False
            Bloquear_Text
        End If
    End If
End If
If Frame6.Caption = "Nuevo Producto" Then
    Verificar_Text
    If auxVacio = False Then
        RS.MoveFirst
        While RS.EOF = False And RS("Estado") <> "baja"
            RS.MoveNext
        Wend
        If RS.EOF = True Then
            RS.MoveLast
            RS.AddNew
        End If
        
        RS("Producto") = Text3(0).Text
        RS("Descripcion") = Text3(1).Text
        RS("Marca") = Text3(2).Text
        RS("Stock") = Text3(3).Text
        RS("PrecioDeCompra") = Text3(4).Text
        RS("PrecioDeVenta") = Text3(5).Text
        RS("PosicionEnDeposito") = ("Sec." + Text3(6).Text + "/F" + Text3(7).Text + "/C" + Text3(8).Text)
        RS("Estado") = "alta"
        RS.Update
        
        If Command6.Visible = True Then
            Frame2.Visible = True
        Else
            Frame5.Visible = True
            Command6.Visible = True
            Command5.Visible = True
            Command7.Visible = True
        End If
        Frame3.Visible = True
        Frame1.Visible = True
        Frame6.Visible = False
        Frame4.Visible = True
            
        Bloquear_Text
        Recargar_listview1
    End If
End If
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command5_Click()

Frame8.Visible = True
Frame4.Enabled = False
Frame2.Enabled = False
Frame1.Enabled = False
Frame3.Enabled = False
Frame5.Enabled = False
Command5.Enabled = False
Command7.Enabled = False
Command6.Enabled = False
End Sub


Private Sub Command5_GotFocus()
Command7.Enabled = False
End Sub

Private Sub Command6_Click()
If Frame5.Visible = True Then
    Frame5.Visible = False
    Command5.Visible = False
    Command7.Visible = False
    Frame7.Visible = True
    ListView2.ListItems.Clear
    Frame4.Visible = False
Else
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
     
    
    Unload Me
End If
End Sub

Private Sub Command7_Click()
If ListView2.ListItems.Count <> 0 Then
    ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
    ListView2.SetFocus
End If
End Sub

Private Sub Command7_GotFocus()
Command5.Enabled = False
End Sub

Private Sub Command7_LostFocus()
If ListView2.ListItems.Count = 0 Then
    Command7.Enabled = False
    Command9.Enabled = False
End If

End Sub

Private Sub Command8_Click()
Frame6.Visible = True
Frame6.Caption = "Nuevo Producto"
Frame4.Visible = False
Frame5.Visible = False
Frame3.Visible = False
Frame1.Visible = False
Command6.Visible = False
Command5.Visible = False
Command7.Visible = False
End Sub

Private Sub Command8_GotFocus()
Command5.Enabled = False
Command7.Enabled = False
End Sub

Private Sub Command9_Click()
If ListView2.ListItems.Count > 0 Then
    For i = 1 To ListView2.ListItems.Count
        RS.MoveFirst
        While RS("Id") <> ListView2.ListItems(i).ListSubItems(6) And RS.EOF = False
            RS.MoveNext
        Wend
        If RS("PrecioDeCompra") <> ListView2.ListItems(i).ListSubItems(4) Then
            RS("PrecioDeCompra") = ListView2.ListItems(i).ListSubItems(4)
        End If
        If AumentarDisminuir = "Aumentar" Then
            RS("Stock") = Str(Val(RS("Stock")) + Val(ListView2.ListItems(i).ListSubItems(3)))
        Else
            If Val(RS("Stock")) - Val(ListView2.ListItems(i).ListSubItems(3)) < 0 Then
                RS("Stock") = "0"
            Else
                RS("Stock") = Str(Val(RS("Stock")) - Val(ListView2.ListItems(i).ListSubItems(3)))
            End If
        End If
        RS.Update
    Next i
Recargar_listview1
ListView2.ListItems.Clear
End If

End Sub

Private Sub Command9_GotFocus()
Command5.Enabled = False
End Sub

Private Sub Form_Load()

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Datos\Datos.mdb"
RS.Source = "Productos"
RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
RS.Open "select * from Productos", cn
RS.MoveFirst

ListView1.ListItems.Clear

Colocar_todos_los_productos

Skin2.LoadSkin App.Path & "\Datos\SKN\GT3.skn"
Skin2.ApplySkin Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
End Sub

Private Sub ListView1_Click()
Command1.Enabled = True
End Sub

Private Sub ListView1_GotFocus()
Command5.Enabled = True
Command7.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub ListView2_GotFocus()
Command5.Enabled = False
If ListView2.ListItems.Count > 0 Then
    Command7.Enabled = True
End If
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
    BUSQUEDA
Else
    ListView1.ListItems.Clear
    Colocar_todos_los_productos
End If
End Sub

Private Sub Text1_GotFocus()
Command5.Enabled = False
Command7.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.Text = "" Then
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = (Int(KeyAscii) - 32)
    End If
End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Or Index >= 7 Then
    Asciiaux = KeyAscii
    Solo_Numeros
    KeyAscii = Asciiaux
End If
If Index = 4 Or Index = 5 Then
    textboxaux = Text3(Index).Text
    Asciiaux = KeyAscii
    Solo_numeros_y_Un_punto
    KeyAscii = Asciiaux
End If
If Text3(Index).Text = "" Then
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = (Int(KeyAscii) - 32)
    End If
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
textboxaux = Text4.Text
Asciiaux = KeyAscii
Solo_numeros_y_Un_punto
KeyAscii = Asciiaux
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Asciiaux = KeyAscii
Solo_Numeros
KeyAscii = Asciiaux
End Sub

Private Sub Text6_Change()
Dim aux As Integer
Text7.Text = Str(Round(((Val(Text6.Text) * 0.01) * Val(Text4.Text)) + Val(Text4.Text), 3))
aux = Str(Int(Len(Text7.Text)) - 1)
If Mid(Text7.Text, aux, 1) = "." Then
    Text7.Text = Text7.Text + "0"
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
textboxaux = Text7.Text
Asciiaux = KeyAscii
Solo_numeros_y_Un_punto
KeyAscii = Asciiaux
End Sub
