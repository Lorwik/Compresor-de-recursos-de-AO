VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00424242&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compresor ComunidadWinter"
   ClientHeight    =   5370
   ClientLeft      =   -15
   ClientTop       =   255
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7470
   StartUpPosition =   1  'CenterOwner
   Begin Compresor.lvButtons_H LvBCerrar 
      Height          =   405
      Left            =   120
      TabIndex        =   17
      Top             =   4830
      Width           =   7245
      _extentx        =   12779
      _extenty        =   714
      caption         =   "Cerrar"
      capalign        =   2
      backstyle       =   2
      font            =   "frmmain.frx":10CA
      cfore           =   16777215
      cfhover         =   16777215
      cbhover         =   0
      cgradient       =   0
      gradient        =   3
      mode            =   0
      value           =   0   'False
      cback           =   255
   End
   Begin VB.Frame FraRecursos 
      BackColor       =   &H00535353&
      Caption         =   "Recursos"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   7215
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Minimapa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   3
         Top             =   2100
         Width           =   1575
      End
      Begin VB.TextBox txtSkinName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   16
         Text            =   "Winter"
         Top             =   1750
         Width           =   1335
      End
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Skin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   15
         Top             =   1760
         Width           =   735
      End
      Begin VB.TextBox txtVersion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Text            =   "0"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Mapas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Fuentes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Sonidos Ambientales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Interfaces"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Inits"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Sonidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Musica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton OptRecursos 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Graficos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label lblVersión 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Versión:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   585
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00535353&
      Caption         =   "Parches"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3510
      Width           =   7215
      Begin Compresor.lvButtons_H cmdComprmirParche 
         Height          =   345
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Width           =   3705
         _extentx        =   6535
         _extenty        =   609
         caption         =   "Comprimir Parche"
         capalign        =   2
         backstyle       =   2
         shape           =   2
         font            =   "frmmain.frx":10F2
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
      Begin Compresor.lvButtons_H cmdDesComprmirParche 
         Height          =   345
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   3465
         _extentx        =   6112
         _extenty        =   609
         caption         =   "Descomprimir Parche"
         capalign        =   2
         backstyle       =   2
         shape           =   1
         font            =   "frmmain.frx":111A
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00535353&
      Caption         =   "Comprension de recursos"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2670
      Width           =   7215
      Begin Compresor.lvButtons_H cmdDescompresion 
         Height          =   345
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   3615
         _extentx        =   6376
         _extenty        =   609
         caption         =   "Descomprimir"
         capalign        =   2
         backstyle       =   2
         shape           =   1
         font            =   "frmmain.frx":1142
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
      Begin Compresor.lvButtons_H cmdCompresion 
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   270
         Width           =   3495
         _extentx        =   6165
         _extenty        =   609
         caption         =   "Comprimir"
         capalign        =   2
         backstyle       =   2
         shape           =   2
         font            =   "frmmain.frx":116A
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   4350
      Width           =   7215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Tipo As srcFileType

Private Sub cmdComprimirParche_Click()

    If Not IsNumeric(txtVersion.Text) Then
        MsgBox "La versión debe ser un valor numerico"
        Exit Sub
    End If
    
    Label1.Caption = "Creando parche..."
    Label1.BackColor = &HFF&
    
    If compressFiles(Patch) Then
        Label1.Caption = "Compresion terminada"
        Label1.BackColor = &HFF00&
    Else
        Label1.Caption = "Error al comprimir"
        Label1.BackColor = &HFF&
    End If
End Sub

Private Sub cmdCompresion_Click()

    If Not IsNumeric(txtVersion.Text) Then
        MsgBox "La versión debe ser un valor numerico"
        Exit Sub
    End If

    Label1.Caption = "Trabajando..."
    Label1.BackColor = &HFF&
    
    If compressFiles(Tipo) Then
        Label1.Caption = "Compresion terminada"
        Label1.BackColor = &HFF00&
    Else
        Label1.Caption = "Error al comprimir"
        Label1.BackColor = &HFF&
    End If
    
End Sub

Private Sub cmdDescompresion_Click()
    Label1.Caption = "Trabajando..."
    Label1.BackColor = &HFF&
    
    If extractFiles(Tipo) Then
        Label1.Caption = "Extracción terminada"
        Label1.BackColor = &HFF00&
    Else
        Label1.Caption = "Error al extraer"
        Label1.BackColor = &HFF&
    End If
End Sub

Public Function General_File_Exists(ByVal file_path As String, ByVal File_Type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************
    If Dir(file_path, File_Type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Private Sub LvBCerrar_Click()
    End
End Sub

Private Sub OptRecursos_Click(Index As Integer)

    Tipo = Index
        
End Sub
