Attribute VB_Name = "modGeneral"
Option Explicit

Public Const Extension = ".WAO"
Public Const PasswordResources = "$FlLrjB3JoliHdAPKA8&YaJR5"

Public SrcPath As String
Public OutPath As String

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Sub Main()
    Call CargarDirectorios

    Call GenerateContra
    
    If SrcPath = "" Or OutPath = "" Then
        frmConfig.Show
        
    Else
        FrmMain.Show
        
    End If
End Sub

Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String
                        
'******************************************************************
' Funcción que abre el cuadro de dialogo y retorna la ruta
'******************************************************************
  
On Local Error GoTo errFunction
      
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
      
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
      
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
      
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
      
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path
  
Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString

End Function

Private Function CargarDirectorios() As Boolean
'*****************************
'Autor: Lorwik
'Fecha: 25/09/2023
'*****************************

    On Error GoTo hErr

    Dim ConfigFile As String
    
    ConfigFile = App.Path & "\Configuracion.ini"
    
    SrcPath = GetVar(ConfigFile, "MAIN", "DirEncriptados")
    OutPath = GetVar(ConfigFile, "MAIN", "DirDesencriptados")
    
    CargarDirectorios = True
    
    Exit Function
    
hErr:

    CargarDirectorios = False

End Function

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, value, file
    
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean

    FileExist = (Dir$(file, FileType) <> "")
    
End Function
