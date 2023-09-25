VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00424242&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraDescomprimidos 
      BackColor       =   &H00424242&
      Caption         =   "Descomprimidos"
      Height          =   1335
      Left            =   150
      TabIndex        =   4
      Top             =   1590
      Width           =   7815
      Begin Compresor.lvButtons_H cmdSelecDescom 
         Height          =   465
         Left            =   270
         TabIndex        =   5
         Top             =   660
         Width           =   7245
         _extentx        =   12779
         _extenty        =   820
         caption         =   "Seleccionar carpeta"
         capalign        =   2
         backstyle       =   2
         font            =   "frmConfig.frx":0000
         cfore           =   16777215
         cfhover         =   16777215
         cbhover         =   0
         cgradient       =   0
         gradient        =   3
         mode            =   0
         value           =   0   'False
         cback           =   32768
      End
      Begin VB.Label lblDescomprimidos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "./"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   990
         TabIndex        =   7
         Top             =   180
         Width           =   135
      End
      Begin VB.Label lblCarpeta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   180
         Width           =   645
      End
   End
   Begin VB.Frame FraRecursos 
      BackColor       =   &H00424242&
      Caption         =   "Recursos"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7815
      Begin Compresor.lvButtons_H cmdSelecRecursos 
         Height          =   465
         Left            =   270
         TabIndex        =   1
         Top             =   660
         Width           =   7245
         _extentx        =   12779
         _extenty        =   820
         caption         =   "Seleccionar carpeta"
         capalign        =   2
         backstyle       =   2
         font            =   "frmConfig.frx":0028
         cfore           =   16777215
         cfhover         =   16777215
         cbhover         =   0
         cgradient       =   0
         gradient        =   3
         mode            =   0
         value           =   0   'False
         cback           =   32768
      End
      Begin VB.Label lblCarpeta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   180
         Width           =   645
      End
      Begin VB.Label lblRecursos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "./"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   990
         TabIndex        =   2
         Top             =   180
         Width           =   135
      End
   End
   Begin Compresor.lvButtons_H LvBAceptar 
      Height          =   405
      Left            =   420
      TabIndex        =   8
      Top             =   3210
      Width           =   7245
      _extentx        =   12779
      _extenty        =   714
      caption         =   "Aceptar"
      capalign        =   2
      backstyle       =   2
      font            =   "frmConfig.frx":0050
      cfore           =   16777215
      cfhover         =   16777215
      cbhover         =   0
      cgradient       =   0
      gradient        =   3
      mode            =   0
      value           =   0   'False
      cback           =   255
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSelecRecursos_Click()
    SrcPath = Buscar_Carpeta("Seleccione la carpeta donde se encontraran los archivos .WAO", "")
    lblRecursos.Caption = OutPath
End Sub

Private Sub cmdSelecDescom_Click()
    OutPath = Buscar_Carpeta("Seleccione la carpeta donde se encontraran los archivos de recursos (graficos, musicas, mapas, etc)", "")
    lblDescomprimidos.Caption = SrcPath
End Sub

Private Sub LvBAceptar_Click()
    Call WriteVar(App.Path & "\Configuracion.ini", "MAIN", "DirEncriptados", SrcPath)
    Call WriteVar(App.Path & "\Configuracion.ini", "MAIN", "DirDesencriptados", OutPath)
    
    Unload Me
    FrmMain.Show
End Sub
