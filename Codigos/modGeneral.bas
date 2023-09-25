Attribute VB_Name = "modGeneral"
Option Explicit

Public Const Extension = ".WAO"
Public Const PasswordResources = "$FlLrjB3JoliHdAPKA8&YaJR5"


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

