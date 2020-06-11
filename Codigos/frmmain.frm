VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compresor WinterAO"
   ClientHeight    =   5160
   ClientLeft      =   -15
   ClientTop       =   255
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraRecursos 
      Caption         =   "Recursos"
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4575
      Begin VB.OptionButton OptRecursos 
         Caption         =   "Mapas"
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton OptRecursos 
         Caption         =   "Fuentes"
         Height          =   255
         Index           =   7
         Left            =   2400
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton OptRecursos 
         Caption         =   "Sonidos Ambientales"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton OptRecursos 
         Caption         =   "Interfaces"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton OptRecursos 
         Caption         =   "Inits"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OptRecursos 
         Caption         =   "Sonidos"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptRecursos 
         Caption         =   "Musica"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton OptRecursos 
         Caption         =   "Graficos"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin vbalProgBarLib6.vbalProgressBar Barrita 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      Picture         =   "frmmain.frx":000C
      ForeColor       =   0
      Appearance      =   2
      BarPicture      =   "frmmain.frx":0028
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parches"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   4575
      Begin VB.CommandButton cmdComprimirParche 
         Caption         =   "Comprimir Parche"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Descomprimir Parche"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comprension de recursos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4575
      Begin VB.CommandButton Command15 
         Caption         =   "Descomprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Comprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MaskColor       =   &H0000FFFF&
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1770
      TabIndex        =   0
      Top             =   3720
      Width           =   1245
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Tipo As resource_file_type

Private Sub cmdComprimirParche_Click()
    Call Compress_Files(Patch, App.Path, App.Path & "\COMPRIMIDOS\")
End Sub

Private Sub Command1_Click()
    Label1.Caption = "ESPERE!!Trabajando"
    Label1.BackColor = &HFF&
    
    Call Compress_Files(Tipo, App.Path, App.Path & "\COMPRIMIDOS\")
    
    Label1.Caption = "Compresion terminada"
    Label1.BackColor = &HFF00&
End Sub
Private Sub Command11_Click()
    End
End Sub
Private Sub Command15_Click()
    Label1.Caption = "ESPERE!!Trabajando"
    Label1.BackColor = &HFF&
    
    Call Extract_All_Files(Tipo, App.Path & "\COMPRIMIDOS")
    
    Label1.Caption = "Extracción terminada"
    Label1.BackColor = &HFF00&
End Sub

Private Sub Command17_Click()
    If General_File_Exists(App.Path & "\tmp.WAO", vbNormal) Then
    '    'Instalamos el Parche
        Extract_Patch App.Path & "\COMPRIMIDOS", App.Path & "\tmp.WAO"
    '
    '    'Esperamos a que termine
        DoEvents
    '
    '    'Borramos el Parche
        Kill App.Path & "\tmp.WAO"
    Else
        MsgBox "No se encuentro el archivo de parche."
        MsgBox App.Path & "\tmp.WAO"
    End If
    
End Sub

Public Function General_File_Exists(ByVal file_path As String, ByVal file_type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************
    If Dir(file_path, file_type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Private Sub Form_Load()
    Call GenerateContra
End Sub

Private Sub OptRecursos_Click(Index As Integer)

    Tipo = Index
        
End Sub
