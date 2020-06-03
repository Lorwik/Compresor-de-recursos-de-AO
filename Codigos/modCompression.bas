Attribute VB_Name = "modCompression"
Option Explicit

Public Windows_Temp_Dir As String 'Normalmente va en Declaraciones, lo pongo aqui para no mezclar.
'Public Formato As String * 6
Public Const Formato As String * 6 = ".WAO"
Public Const PswPk As Long = 274504515

Public PkContra() As Byte

'This structure will describe our binary file's
'size and number of contained files
Public Type FILEHEADER
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    intNumFiles As Integer              'How many files are inside?
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    
    lngFileStart As Long            'Where does the chunk start?
    lngFileSize As Long             'How big is this chunk of stored data?
    strFileName As String * 32      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Public Enum resource_file_type
    Graphics
    Ambient
    Music
    Wav
    Scripts
    Patch
    Map
    Interface
    Fuentes
End Enum

Private Const GRAPHIC_PATH As String = "\GRAFICOS\"
Private Const AMBIENT_PATH As String = "\Ambient\"
Private Const Music_PATH As String = "\Music\"
Private Const WAV_PATH As String = "\Wav\"
Private Const MAP_PATH As String = "\Mapas\"
Private Const INTERFACE_PATH As String = "\Interface\"
Private Const FUENTES_PATH As String = "\Fuentes\"
Private Const SCRIPT_PATH As String = "\Init\"
Private Const OUTPUT_PATH As String = "\Output\"
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Public Sub GenerateContra()
'***************************************************
'Author: ^[GS]^
'Last Modification: 17/06/2012 - ^[GS]^
'
'***************************************************

'on error resume next
    Dim Contra As String
    Dim loopC As Byte
    
    Contra = "T5lTCWm2m1rR7#SMgq!cazNv"
    
    Erase PkContra
    
    If LenB(Contra) <> 0 Then
        ReDim PkContra(Len(Contra) - 1)
        For loopC = 0 To UBound(PkContra)
            PkContra(loopC) = Asc(mid(Contra, loopC + 1, 1))
        Next loopC
    End If
    
End Sub

Public Sub Compress_Data(ByRef Data() As Byte)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Compresses binary data avoiding data loses
'*****************************************************************
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim BufTemp2() As Byte
    Dim loopC As Long
    
    'Get size of the uncompressed data
    Dimensions = UBound(Data)
    
    'Prepare a buffer 1.06 times as big as the original size
    DimBuffer = Dimensions * 1.06
    ReDim BufTemp(DimBuffer)
    
    'Compress data using zlib
    Compress BufTemp(0), DimBuffer, Data(0), Dimensions
    
    'Deallocate memory used by uncompressed data
    Erase Data
    
    'Get rid of unused bytes in the compressed data buffer
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    'Copy the compressed data buffer to the original data array so it will return to caller
    Data = BufTemp
    
    'Deallocate memory used by the temp buffer
    Erase BufTemp
    
    'Encrypt the first byte of the compressed data for extra security
    If UBound(PkContra) <= UBound(Data) And UBound(PkContra) <> 0 Then
        For loopC = 0 To UBound(PkContra)
            Data(loopC) = Data(loopC) Xor PkContra(loopC)
        Next loopC
    End If
End Sub

Public Sub Decompress_Data(ByRef Data() As Byte, ByVal OrigSize As Long)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Decompresses binary data
'*****************************************************************
    Dim BufTemp() As Byte
    Dim loopC As Integer
    
    ReDim BufTemp(OrigSize - 1)
    
    'Des-encrypt the first byte of the compressed data
    If UBound(PkContra) <= UBound(Data) And UBound(PkContra) <> 0 Then
        For loopC = 0 To UBound(PkContra)
            Data(loopC) = Data(loopC) Xor PkContra(loopC)
        Next loopC
    End If
    
    UnCompress BufTemp(0), OrigSize, Data(0), UBound(Data) + 1
    
    ReDim Data(OrigSize - 1)
    
    Data = BufTemp
    
    Erase BufTemp
End Sub

Public Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************
    Dim RetVal As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    RetVal = GetDiskFreeSpace(Left(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function

Public Function General_Get_Temp_Dir() As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'Gets windows temporary directory
'**************************************************************
 Const MAX_LENGTH = 512
   Dim s As String
   Dim C As Long
   s = Space$(MAX_LENGTH)
   C = GetTempPath(MAX_LENGTH, s)
   If C > 0 Then
       If C > Len(s) Then
           s = Space$(C + 1)
           C = GetTempPath(MAX_LENGTH, s)
       End If
   End If
   General_Get_Temp_Dir = IIf(C > 0, Left$(s, C), "")
End Function

Public Sub Extract_All_Files(ByVal file_type As resource_file_type, ByVal resource_path As String, Optional ByVal UseOutputFolder As Boolean = False)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopC As Long
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim Handle As Integer

On Local Error GoTo errhandler
    Select Case file_type
        Case Graphics
            If UseOutputFolder Then
                SourceFilePath = resource_path & "\Graficos" & Formato ' & OUTPUT2_PATH & "Graficos." & Formato
            Else
                SourceFilePath = resource_path & "\Graficos" & Formato
            End If
            OutputFilePath = App.Path & OUTPUT_PATH
            
        Case Ambient
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Ambient" & Formato
            Else
                SourceFilePath = resource_path & "\Ambient" & Formato
            End If
            OutputFilePath = App.Path & OUTPUT_PATH
        
        Case Music
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Music" & Formato
            Else
                SourceFilePath = resource_path & "\Music" & Formato
            End If
            OutputFilePath = App.Path & OUTPUT_PATH
        
        Case Wav
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Sounds" & Formato
            Else
                SourceFilePath = resource_path & "\Sounds" & Formato
            End If
            OutputFilePath = App.Path & OUTPUT_PATH
            
        Case Map
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Maps" & Formato
            Else
                SourceFilePath = resource_path & "\Maps" & Formato
            End If
            OutputFilePath = App.Path & OUTPUT_PATH
        
        Case Scripts
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Init" & Formato
            Else
                SourceFilePath = resource_path & "\Init" & Formato
            End If
            OutputFilePath = App.Path & OUTPUT_PATH
        
        Case Interface
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Interface" & Formato
            Else
                SourceFilePath = resource_path & "\Interface" & Formato
            End If
            OutputFilePath = App.Path & OUTPUT_PATH
            
        Case Fuentes
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Fuentes" & Formato
            Else
                SourceFilePath = resource_path & "\Fuentes" & Formato
            End If
            OutputFilePath = App.Path & OUTPUT_PATH
        
        Case Else
            Exit Sub
    End Select
        
    
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead

    'If PasswordFile2 <> Password3 Then Exit Sub
    FileHead.lngFileSize = FileHead.lngFileSize - (PswPk * 2)
    
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then ' - Pass1.lngFileSize - 1 Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Close SourceFile
        Erase InfoHead
        Exit Sub
    End If
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)

    'Extract the INFOHEADER
    Get SourceFile, , InfoHead

    'Extract all of the files from the binary file
    For loopC = 0 To UBound(InfoHead)
        
        'Check if there is enough memory
        If InfoHead(loopC).lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left(App.Path, 3)) Then
            MsgBox "No tienes suficiente espacio en el disco para seguir descomprimiendo archivos."
            Exit Sub
        End If
        
        'Resize the byte data array
        ReDim SourceData(InfoHead(loopC).lngFileSize - 1)
        
        'Get the data
        Get SourceFile, InfoHead(loopC).lngFileStart, SourceData
        
        'Decompress all data
        Decompress_Data SourceData, InfoHead(loopC).lngFileSizeUncompressed
        
        'Get a free handler
        Handle = FreeFile

        'Create a new file and put in the data
        Open OutputFilePath & InfoHead(loopC).strFileName For Binary As Handle
        
        Put Handle, , SourceData
        
        Close Handle
        
        Erase SourceData
        
        DoEvents
    Next loopC
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    MsgBox "Complete"
Exit Sub

errhandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Sub

Public Function Extract_Patch(ByVal resource_path As String, ByVal file_name As String) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim loopC As Long
    Dim LoopC2 As Long
    Dim LoopC3 As Long
    Dim OutputFile As Integer
    Dim UpdatedFile As Integer
    Dim SourceFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim ResFileHead As FILEHEADER
    Dim ResInfoHead() As INFOHEADER
    Dim UpdatedInfoHead As INFOHEADER
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency
    Dim FileExtension As String
    Dim DataOffset As Long
    Dim OutputFilePath As String
    
    'Done flags
    Dim png_done As Boolean
    Dim wav_done As Boolean
    Dim mid_done As Boolean
    Dim Music_done As Boolean
    Dim exe_done As Boolean
    Dim gui_done As Boolean
    Dim ind_done As Boolean
    Dim dat_done As Boolean
    Dim bmp_done As Boolean
    
    '************************************************************************************************
    'This is similar to Extract, but has some small differences to make sure what is being updated
    '************************************************************************************************
'Set up the error handler
On Local Error GoTo errhandler
    
    'Open the binary file
    SourceFile = FreeFile
    SourceFilePath = file_name
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
        
    'Check the file for validity
    'If LOF(SourceFile) <> FileHead.lngFileSize Then
    '    MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
    '    Exit Function
    'End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
    
    'Check if there is enough hard drive space to extract all files
    For loopC = 0 To UBound(InfoHead)
        RequiredSpace = RequiredSpace + InfoHead(loopC).lngFileSizeUncompressed
    Next loopC
    
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left(App.Path, 3)) Then
        Erase InfoHead
        MsgBox "¡No hay espacio suficiente para extraer el archivo!", , "Error"
        Exit Function
    End If
    
    'Extract all of the files from the binary file
    For loopC = 0 To UBound(InfoHead())
        'Check the extension of the file
        Select Case LCase(Right(Trim(InfoHead(loopC).strFileName), 3))
            Case Is = "png"
                If png_done Then GoTo EndMainLoop
                FileExtension = "png"
                OutputFilePath = resource_path & "\Graficos" & Formato
                png_done = True
            Case Is = "amb"
                If mid_done Then GoTo EndMainLoop
                FileExtension = "amb"
                OutputFilePath = resource_path & "\Ambient" & Formato
                mid_done = True
            Case Is = "Music"
                If Music_done Then GoTo EndMainLoop
                FileExtension = "mp3"
                OutputFilePath = resource_path & "\Music" & Formato
                Music_done = True
            Case Is = "wav"
                If wav_done Then GoTo EndMainLoop
                FileExtension = "wav"
                OutputFilePath = resource_path & "\Sounds" & Formato
                wav_done = True
            Case Is = "jpg"
                If gui_done Then GoTo EndMainLoop
                FileExtension = "jpg"
                OutputFilePath = resource_path & "\Interface" & Formato
                gui_done = True
            Case Is = "ind"
                If bmp_done Then GoTo EndMainLoop
                FileExtension = "bmp"
                OutputFilePath = resource_path & "\Fuentes" & Formato
                ind_done = True
            Case Is = "bmp"
                If bmp_done Then GoTo EndMainLoop
                FileExtension = "bmp"
                OutputFilePath = resource_path & "\Fuentes" & Formato
                bmp_done = True
        End Select
        
        OutputFile = FreeFile
        Open OutputFilePath For Binary Access Read Lock Write As OutputFile
        
        'Get file header
        Get OutputFile, 1, ResFileHead
                
        'Resize the Info Header array
        ReDim ResInfoHead(ResFileHead.intNumFiles - 1)
        
        'Load the info header
        Get OutputFile, , ResInfoHead
                
        'Check how many of the files are new, and how many are replacements
        For LoopC2 = loopC To UBound(InfoHead())
            If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                'Look for same name in the resource file
                For LoopC3 = 0 To UBound(ResInfoHead())
                    If ResInfoHead(LoopC3).strFileName = InfoHead(LoopC2).strFileName Then
                        Exit For
                    End If
                Next LoopC3
                
                'Update the File Head
                If LoopC3 > UBound(ResInfoHead()) Then
                    'Update number of files and size
                    ResFileHead.intNumFiles = ResFileHead.intNumFiles + 1
                    ResFileHead.lngFileSize = ResFileHead.lngFileSize + Len(InfoHead(0)) + InfoHead(LoopC2).lngFileSize
                Else
                    'We substract the size of the old file and add the one of the new one
                    ResFileHead.lngFileSize = ResFileHead.lngFileSize - ResInfoHead(LoopC3).lngFileSize + InfoHead(LoopC2).lngFileSize
                End If
            End If
        Next LoopC2
        
        'Get the offset of the compressed data
        DataOffset = CLng(ResFileHead.intNumFiles) * Len(ResInfoHead(0)) + Len(FileHead) + 1
                
        'Now we start saving the updated file
        UpdatedFile = FreeFile
        Open OutputFilePath & "2" For Binary Access Write Lock Read As UpdatedFile
        
        'Store the filehead
        Put UpdatedFile, 1, ResFileHead
        
        'Start storing the Info Heads
        LoopC2 = loopC
        For LoopC3 = 0 To UBound(ResInfoHead())
            Do While LoopC2 <= UBound(InfoHead())
                If LCase$(ResInfoHead(LoopC3).strFileName) < LCase$(InfoHead(LoopC2).strFileName) Then Exit Do
                If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                    'Copy the info head data
                    UpdatedInfoHead = InfoHead(LoopC2)
                    
                    'Set the file start pos and update the offset
                    UpdatedInfoHead.lngFileStart = DataOffset
                    DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                                        
                    Put UpdatedFile, , UpdatedInfoHead
                    
                    DoEvents
                    
                End If
                LoopC2 = LoopC2 + 1
            Loop
            
            'If the file was replaced in the patch, we skip it
            If LoopC2 Then
                If LCase$(ResInfoHead(LoopC3).strFileName) <= LCase$(InfoHead(LoopC2 - 1).strFileName) Then GoTo EndLoop
            End If
            
            'Copy the info head data
            UpdatedInfoHead = ResInfoHead(LoopC3)
            
            'Set the file start pos and update the offset
            UpdatedInfoHead.lngFileStart = DataOffset
            DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                        
            Put UpdatedFile, , UpdatedInfoHead
EndLoop:
        Next LoopC3
        
        'If there was any file in the patch that would go in the bottom of the list we put it now
        For LoopC2 = LoopC2 To UBound(InfoHead())
            If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                'Copy the info head data
                UpdatedInfoHead = InfoHead(LoopC2)
                
                'Set the file start pos and update the offset
                UpdatedInfoHead.lngFileStart = DataOffset
                DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                                
                Put UpdatedFile, , UpdatedInfoHead
            End If
        Next LoopC2
        
        'Now we start adding the compressed data
        LoopC2 = loopC
        For LoopC3 = 0 To UBound(ResInfoHead())
            Do While LoopC2 <= UBound(InfoHead())
                If LCase$(ResInfoHead(LoopC3).strFileName) < LCase$(InfoHead(LoopC2).strFileName) Then Exit Do
                If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                    'Get the compressed data
                    ReDim SourceData(InfoHead(LoopC2).lngFileSize - 1)
                    
                    Get SourceFile, InfoHead(LoopC2).lngFileStart, SourceData
                    
                    Put UpdatedFile, , SourceData
                End If
                LoopC2 = LoopC2 + 1
            Loop
            
            'If the file was replaced in the patch, we skip it
            If LoopC2 Then
                If LCase$(ResInfoHead(LoopC3).strFileName) <= LCase$(InfoHead(LoopC2 - 1).strFileName) Then GoTo EndLoop2
            End If
            
            'Get the compressed data
            ReDim SourceData(ResInfoHead(LoopC3).lngFileSize - 1)
            
            Get OutputFile, ResInfoHead(LoopC3).lngFileStart, SourceData
            
            Put UpdatedFile, , SourceData
EndLoop2:
        Next LoopC3
        
        'If there was any file in the patch that would go in the bottom of the lsit we put it now
        For LoopC2 = LoopC2 To UBound(InfoHead())
            If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                'Get the compressed data
                ReDim SourceData(InfoHead(LoopC2).lngFileSize - 1)
                
                Get SourceFile, InfoHead(LoopC2).lngFileStart, SourceData
                
                Put UpdatedFile, , SourceData
            End If
        Next LoopC2
        
        'We are done updating the file
        Close UpdatedFile
        
        'Close and delete the old resource file
        Close OutputFile
        Kill OutputFilePath
        
        'Rename the new one
        Name OutputFilePath & "2" As OutputFilePath
        
        'Deallocate the memory used by the data array
        Erase SourceData
EndMainLoop:
    Next loopC
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    Erase ResInfoHead
    
    Extract_Patch = True
Exit Function

errhandler:
    Erase SourceData
    Erase InfoHead

End Function

Public Sub Compress_Files(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal dest_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim SourceFilePath As String
    Dim SourceFileExtension As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceFileName As String
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim FileNames() As String
    Dim lngFileStart As Long
    Dim loopC As Long
    Dim tmplng As Long
    
'Set up the error handler
On Local Error GoTo errhandler
    Select Case file_type
        Case Graphics
            SourceFilePath = resource_path & GRAPHIC_PATH
            SourceFileExtension = ".png"
            OutputFilePath = dest_path & "Graficos" & Formato
        
        Case Ambient
            SourceFilePath = resource_path & AMBIENT_PATH
            SourceFileExtension = ".amb"
            OutputFilePath = dest_path & "Ambient" & Formato
        
        Case Music
            SourceFilePath = resource_path & Music_PATH
            SourceFileExtension = ".mp3"
            OutputFilePath = dest_path & "Music" & Formato
        
        Case Wav
            SourceFilePath = resource_path & WAV_PATH
            SourceFileExtension = ".wav"
            OutputFilePath = dest_path & "Sounds" & Formato
            
        Case Map
            SourceFilePath = resource_path & MAP_PATH
            SourceFileExtension = ".csm"
            OutputFilePath = dest_path & "Maps" & Formato
            
        Case Scripts
            SourceFilePath = resource_path & SCRIPT_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Init" & Formato
            
        Case Interface
            SourceFilePath = resource_path & INTERFACE_PATH
            SourceFileExtension = ".jpg"
            OutputFilePath = dest_path & "Interface" & Formato
            
        Case Fuentes
            SourceFilePath = resource_path & INTERFACE_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Fuentes" & Formato
    
    End Select
    
    'Get first file in the directoy
    SourceFileName = Dir$(SourceFilePath & "*" & SourceFileExtension, vbNormal)
    
    SourceFile = FreeFile
    
    'Get all other files i nthe directory
    While SourceFileName <> ""
        FileHead.intNumFiles = FileHead.intNumFiles + 1
        
        ReDim Preserve FileNames(FileHead.intNumFiles - 1)
        FileNames(FileHead.intNumFiles - 1) = LCase(SourceFileName)
        
        'Search new file
        SourceFileName = Dir$()
    Wend
    
    'If we found none, be can't compress a thing, so we exit
    If FileHead.intNumFiles = 0 Then
        MsgBox "There are no files of extension " & SourceFileExtension & " in " & SourceFilePath & ".", , "Error"
        Exit Sub
    End If
    
    'Sort file names alphabetically (this will make patching much easier).
    General_Quick_Sort FileNames(), 0, UBound(FileNames)
    
    'Resize InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
        
    'Destroy file if it previuosly existed
    If Dir(OutputFilePath, vbNormal) <> "" Then
        Kill OutputFilePath
    End If
    
    'Open a new file
    OutputFile = FreeFile
    Open OutputFilePath For Binary Access Read Write As OutputFile
    
    For loopC = 0 To FileHead.intNumFiles - 1
        'Find a free file number to use and open the file
        
        SourceFile = FreeFile
        Open SourceFilePath & FileNames(loopC) For Binary Access Read Lock Write As SourceFile
        
        'Store file name
        InfoHead(loopC).strFileName = FileNames(loopC)
        
        'Find out how large the file is and resize the data array appropriately
        ReDim SourceData(LOF(SourceFile) - 1)
        
        'Store the value so we can decompress it later on
        InfoHead(loopC).lngFileSizeUncompressed = LOF(SourceFile)
        
        'Get the data from the file
        Get SourceFile, , SourceData
        'If loopc = 0 Then SourceData = "115792!"
        'Compress it
        Compress_Data SourceData
        'Save it to a temp file
        Put OutputFile, , SourceData
        
        'Set up the file header
        FileHead.lngFileSize = FileHead.lngFileSize + UBound(SourceData) + 1
        
        'Set up the info headers
        InfoHead(loopC).lngFileSize = UBound(SourceData) + 1
        
        Erase SourceData
        
        'Close temp file
        Close SourceFile
        
        DoEvents
    Next loopC
    
    'Finish setting the FileHeader data
    FileHead.lngFileSize = FileHead.lngFileSize + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + Len(FileHead)
    
    'Set InfoHead data
    lngFileStart = Len(FileHead) + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + 1
    For loopC = 0 To FileHead.intNumFiles - 1
        InfoHead(loopC).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(loopC).lngFileSize
    Next loopC
        
    '************ Write Data
    FileHead.lngFileSize = FileHead.lngFileSize + (PswPk * 2)
    
    'InfoHead.lngFileSize = InfoHead.lngFileSize + 15
    'Get all data stored so far
    ReDim SourceData(LOF(OutputFile) - 1)
    Seek OutputFile, 1
    Get OutputFile, , SourceData
    
    Seek OutputFile, 1
    
    'Store the data in the file
    
    'Put OutputFile, , Pass1
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
    Put OutputFile, , SourceData
    
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
    
    MsgBox "complete"
Exit Sub

errhandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to create binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Sub

Public Function Extract_File(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal file_name As String, ByVal OutputFilePath As String, Optional ByVal UseOutputFolder As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopC As Long
    Dim SourceFilePath As String
    Dim SourceData() As Byte
    Dim InfoHead As INFOHEADER
    Dim Handle As Integer
    
'Set up the error handler
On Local Error GoTo errhandler
    
    Select Case file_type
        Case Graphics
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Graficos" & Formato
            Else
                SourceFilePath = resource_path & "\Graficos" & Formato
            End If
            
        Case Ambient
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Ambient" & Formato
            Else
                SourceFilePath = resource_path & "\Ambient" & Formato
            End If
        
        Case Music
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Music" & Formato
            Else
                SourceFilePath = resource_path & "\Music" & Formato
            End If
        
        Case Wav
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Sounds" & Formato
            Else
                SourceFilePath = resource_path & "\Sounds" & Formato
            End If
        
        Case Scripts
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Init" & Formato
            Else
                SourceFilePath = resource_path & "\Init" & Formato
            End If
        
        Case Interface
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Interface" & Formato
            Else
                SourceFilePath = resource_path & "\Interface" & Formato
            End If
            
        Case Fuentes
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Fuentes" & Formato
            Else
                SourceFilePath = resource_path & "\Fuentes" & Formato
            End If
        
        Case Else
            Exit Function
    End Select
    
    'Find the Info Head of the desired file
    InfoHead = File_Find(SourceFilePath, file_name)
    
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    Handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As Handle
    
    'Check the file for validity
    'If LOF(handle) <> InfoHead.lngFileSize Then
    '    Close handle
    '    MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
    '    Exit Function
    'End If
    
    'Make sure there is enough space in the HD
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close Handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim SourceData(InfoHead.lngFileSize - 1)
    
    'Get the data
    Get Handle, InfoHead.lngFileStart, SourceData
    
    'Decompress all data
    Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    
    'Close the binary file
    Close Handle
    
    'Get a free handler
    Handle = FreeFile
    
    Open OutputFilePath & InfoHead.strFileName For Binary As Handle
    
    Put Handle, 1, SourceData
    
    Close Handle
    
    Erase SourceData
        
    Extract_File = True
Exit Function

errhandler:
    Close Handle
    Erase SourceData
    'Display an error message if it didn't work
    'MsgBox "Unable to decode binary file. Reason: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_File_Ex(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal file_name As String, ByRef bytArr() As Byte) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopC As Long
    Dim SourceFilePath As String
    Dim InfoHead As INFOHEADER
    Dim Handle As Integer
    
'Set up the error handler
On Local Error GoTo errhandler
    
    Select Case file_type
        Case Graphics
                SourceFilePath = resource_path & "\Graphics" & Formato
            
        Case Music
                SourceFilePath = resource_path & "\Musics" & Formato
        
        Case Wav
                SourceFilePath = resource_path & "\Sounds" & Formato

        Case Scripts
                SourceFilePath = resource_path & "\Scripts" & Formato

        Case Interface
                SourceFilePath = resource_path & "\Interface" & Formato

        Case Map
                SourceFilePath = resource_path & "\Maps" & Formato

        Case Ambient
                SourceFilePath = resource_path & "\Ambient" & Formato
                
        Case Fuentes
                SourceFilePath = resource_path & "\Fuentes" & Formato
                
        Case Else
            Exit Function
    End Select
    
    'Find the Info Head of the desired file
    InfoHead = File_Find(SourceFilePath, file_name)
    
    If InfoHead.strFileName = vbNullString Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    Handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As Handle
        
    'Make sure there is enough space in the HD
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close Handle
        MsgBox "¡Espacio insuficiente!", , "Error"
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim bytArr(InfoHead.lngFileSize - 1)
    
    'Get the data
    Get Handle, InfoHead.lngFileStart, bytArr
    
    'Decompress all data
    Decompress_Data bytArr, InfoHead.lngFileSizeUncompressed
        
    'Close the binary file
    Close Handle
            
    Extract_File_Ex = True
Exit Function

errhandler:
    Close Handle
    Erase bytArr
    
End Function

Public Sub Delete_File(ByVal file_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 3/03/2005
'Deletes a resource files
'*****************************************************************
    Dim Handle As Integer
    Dim Data() As Byte
    
    On Error GoTo Error_Handler
    
    'We open the file to delete
    Handle = FreeFile
    Open file_path For Binary Access Write Lock Read As Handle
    
    'We replace all the bytes in it with 0s
    ReDim Data(LOF(Handle) - 1)
    Put Handle, 1, Data
    
    'We close the file
    Close Handle
    
    'Now we delete it, knowing that if they retrieve it (some antivirus may create backup copies of deleted files), it will be useless
    Kill file_path
    
    Exit Sub
    
Error_Handler:
    Kill file_path
        
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

Public Sub Parchear(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal dest_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim SourceFilePath As String
    Dim SourceFileExtension As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceFileName As String
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim FileNames() As String
    Dim lngFileStart As Long
    Dim loopC As Long
'Set up the error handler
On Local Error GoTo errhandler
    Select Case file_type
        Case Graphics
            SourceFilePath = resource_path & GRAPHIC_PATH
            SourceFileExtension = ".png"
            OutputFilePath = dest_path & "Graficos" & Formato
        
        Case Ambient
            SourceFilePath = resource_path & AMBIENT_PATH
            SourceFileExtension = ".amb"
            OutputFilePath = dest_path & "Ambient" & Formato
        
        Case Music
            SourceFilePath = resource_path & Music_PATH
            SourceFileExtension = ".mp3"
            OutputFilePath = dest_path & "Music" & Formato
        
        Case Wav
            SourceFilePath = resource_path & WAV_PATH
            SourceFileExtension = ".wav"
            OutputFilePath = dest_path & "Sounds" & Formato
                
        Case Map
            SourceFilePath = resource_path & MAP_PATH
            SourceFileExtension = ".csm"
            OutputFilePath = dest_path & "Maps" & Formato
            
        Case Scripts
            SourceFilePath = resource_path & SCRIPT_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Init" & Formato
            
        Case Interface
            SourceFilePath = resource_path & INTERFACE_PATH
            SourceFileExtension = ".jpg"
            OutputFilePath = dest_path & "Interface" & Formato
            
        Case Fuentes
            SourceFilePath = resource_path & FUENTES_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Fuentes" & Formato
    
    End Select
    
    'Get first file in the directoy
    SourceFileName = Dir$(SourceFilePath & "*" & SourceFileExtension, vbNormal)
    
    SourceFile = FreeFile
    
    'Get all other files i nthe directory
    While SourceFileName <> ""
        FileHead.intNumFiles = FileHead.intNumFiles + 1
        
        ReDim Preserve FileNames(FileHead.intNumFiles - 1)
        FileNames(FileHead.intNumFiles - 1) = LCase(SourceFileName)
        
        'Search new file
        SourceFileName = Dir$()
    Wend
    
    'If we found none, be can't compress a thing, so we exit
    If FileHead.intNumFiles = 0 Then
        MsgBox "There are no files of extension " & SourceFileExtension & " in " & SourceFilePath & ".", , "Error"
        Exit Sub
    End If
    
    'Sort file names alphabetically (this will make patching much easier).
    General_Quick_Sort FileNames(), 0, UBound(FileNames)
    
    'Resize InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
        
    'Destroy file if it previuosly existed
    'If Dir(OutputFilePath, vbNormal) <> "" Then
        'Kill OutputFilePath
    'End If
    
    'Open a new file
    OutputFile = FreeFile
    Open OutputFilePath For Binary Access Read Write As OutputFile
    
    For loopC = 0 To FileHead.intNumFiles - 1
        'Find a free file number to use and open the file
        
        SourceFile = FreeFile
        Open SourceFilePath & FileNames(loopC) For Binary Access Read Lock Write As SourceFile
        
        'Store file name
        InfoHead(loopC).strFileName = FileNames(loopC)
        
        'Find out how large the file is and resize the data array appropriately
        ReDim SourceData(LOF(SourceFile) - 1)
        
        'Store the value so we can decompress it later on
        InfoHead(loopC).lngFileSizeUncompressed = LOF(SourceFile)
        
        'Get the data from the file
        Get SourceFile, , SourceData
        'If loopc = 0 Then SourceData = "115792!"
        'Compress it
        Compress_Data SourceData
        'Save it to a temp file
        Put OutputFile, , SourceData
        
        'Set up the file header
        FileHead.lngFileSize = FileHead.lngFileSize + UBound(SourceData) + 1
        
        'Set up the info headers
        InfoHead(loopC).lngFileSize = UBound(SourceData) + 1
        
        Erase SourceData
        
        'Close temp file
        Close SourceFile
        
        DoEvents
    Next loopC
    
    'Finish setting the FileHeader data
    FileHead.lngFileSize = FileHead.lngFileSize + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + Len(FileHead)
    
    'Set InfoHead data
    lngFileStart = Len(FileHead) + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + 1
    For loopC = 0 To FileHead.intNumFiles - 1
        InfoHead(loopC).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(loopC).lngFileSize
    Next loopC
        
    '************ Write Data
    FileHead.lngFileSize = FileHead.lngFileSize + (PswPk * 2)
    
    'Get all data stored so far
    ReDim SourceData(LOF(OutputFile) - 1)
    Seek OutputFile, 1
    Get OutputFile, , SourceData
    
    Seek OutputFile, 1
    
    'Store the data in the file
    
    'Put OutputFile, , Pass1
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
    Put OutputFile, , SourceData
    
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
Exit Sub

errhandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to create binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Sub

Public Function Extract_File_Memory(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal file_name As String, ByRef SourceData() As Byte) As Boolean
 
    ' Parra was here (;
    Dim loopC As Long
    Dim SourceFilePath As String
    Dim InfoHead As INFOHEADER
    Dim Handle As Integer
   
On Local Error GoTo errhandler
   
    Select Case file_type
        Case Graphics
                SourceFilePath = resource_path & "\Graphics" & Formato
            
        Case Music
                SourceFilePath = resource_path & "\Musics" & Formato
        
        Case Wav
                SourceFilePath = resource_path & "\Sounds" & Formato

        Case Scripts
                SourceFilePath = resource_path & "\Scripts" & Formato

        Case Interface
                SourceFilePath = resource_path & "\Interface" & Formato

        Case Map
                SourceFilePath = resource_path & "\Maps" & Formato

        Case Ambient
                SourceFilePath = resource_path & "\Ambient" & Formato
                
        Case Fuentes
                SourceFilePath = resource_path & "\Fuentes" & Formato
                
        Case Else
            Exit Function
    End Select
   
    InfoHead = File_Find(SourceFilePath, file_name)
   
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function
 
    Handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As Handle
   
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close Handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
   
   
    ReDim SourceData(InfoHead.lngFileSize - 1)
   
    Get Handle, InfoHead.lngFileStart, SourceData
        Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    Close Handle
       
    Extract_File_Memory = True
Exit Function
 
errhandler:
    Close Handle
    Erase SourceData
End Function
Public Sub General_Quick_Sort(ByRef SortArray As Variant, ByVal first As Long, ByVal last As Long)
    Dim Low As Long, High As Long
    Dim temp As Variant
    Dim List_Separator As Variant
   
    Low = first
    High = last
    List_Separator = SortArray((first + last) / 2)
    Do While (Low <= High)
        Do While SortArray(Low) < List_Separator
            Low = Low + 1
        Loop
        Do While SortArray(High) > List_Separator
            High = High - 1
        Loop
        If Low <= High Then
            temp = SortArray(Low)
            SortArray(Low) = SortArray(High)
            SortArray(High) = temp
            Low = Low + 1
            High = High - 1
        End If
    Loop
    If first < High Then General_Quick_Sort SortArray, first, High
    If Low < last Then General_Quick_Sort SortArray, Low, last
End Sub
Public Function Get_Extract(ByVal file_type As resource_file_type, ByVal file_name As String) As String
    Extract_File file_type, App.Path & "\Graficos", LCase$(file_name), App.Path & "\amdInData\"
    Get_Extract = App.Path & "\amdInData\" & LCase$(file_name)
End Function
Public Function Get_Interface(ByVal file_type As resource_file_type, ByVal file_name As String) As String
    Extract_File file_type, App.Path & "\Interface", LCase$(file_name), App.Path & "\Interface\"
    Get_Interface = App.Path & "\Interface\" & LCase$(file_name)
End Function

Public Function File_Find(ByVal resource_file_path As String, ByVal file_name As String) As INFOHEADER
 
On Error GoTo errhandler
 
    Dim Max As Integer
    Dim Min As Integer
    Dim mid As Integer
    Dim file_handler As Integer
   
    Dim file_head As FILEHEADER
    Dim info_head As INFOHEADER
   
    If Len(file_name) < Len(info_head.strFileName) Then _
        file_name = file_name & Space$(Len(info_head.strFileName) - Len(file_name))
   
    file_handler = FreeFile
    Open resource_file_path For Binary Access Read Lock Write As file_handler
   
    Get file_handler, 1, file_head
   
    Min = 1
    Max = file_head.intNumFiles
   
    Do While Min <= Max
        mid = (Min + Max) / 2
       
        Get file_handler, CLng(Len(file_head) + CLng(Len(info_head)) * CLng((mid - 1)) + 1), info_head
               
        If file_name < info_head.strFileName Then
            If Max = mid Then
                Max = Max - 1
            Else
                Max = mid
            End If
        ElseIf file_name > info_head.strFileName Then
            If Min = mid Then
                Min = Min + 1
            Else
                Min = mid
            End If
        Else
            File_Find = info_head
           
            Close file_handler
            Exit Function
        End If
    Loop
   
errhandler:
    Close file_handler
    File_Find.strFileName = ""
    File_Find.lngFileSize = 0
End Function


