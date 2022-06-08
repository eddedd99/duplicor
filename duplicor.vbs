'**********************************************************************************
' Objetivo: Ordernar la salida de un DIR para integrar en tabla y buscar duplicados
'    Fecha:07/Jun/2022
'    Autor: edcruces99
'
'**********************************************************************************

'Crear Objeto
Set fso = CreateObject("Scripting.FileSystemObject")

'Crear Array
Set arrCampos = CreateObject("System.Collections.ArrayList")

'Leer archivo entrada
Filename = WScript.Arguments.Item(0)
'msgbox "Arg0: " & Filename

arrCampos.Add "Archivo,Tamaño,Fecha_modif"

'Leer nombre archivo sin extensión
arrFilename=Split(Filename,".")
for each fName in arrFilename
    FilenameJustName=fName
	Exit For
next
'msgBox FilenameJustName

'Abrir archivo
Set f = fso.OpenTextFile(Filename)

'Parseo del Contenido
bEncontroTexto = 0
linea = f.ReadLine
'msgBox linea
Do Until f.AtEndOfStream

   'Parseo
   If InStr(linea," Directorio de ") > 0 Then
      bEncontroTexto = 1
      i=16
      While i <= Len(linea)
         strC = strC & Mid(linea, i, 1)
         i=i+1
      Wend
	  sRutaArch = strC
	  strC=""
   ElseIf InStr(linea,"a. m.") > 0 or InStr(linea,"p. m.") > 0 Then
      bEncontroTexto = 1
      i=1
      While i <= 23 'Fecha
         strC = strC & Mid(linea, i, 1)
         i=i+1
      Wend
	  sFecha = strC
      strC=""
	  
      i=24
      While Mid(linea, i, 1) = Chr(32) 'Leo Espacios
         i=i+1
      Wend

      While IsNumeric(Mid(linea, i, 1)) 'Leo Numeros (Tamaño)
         strC = strC & Mid(linea, i, 1)
         i=i+1
      Wend
	  sFileSize = strC
      strC=""
	  
      While Mid(linea, i, 1) <> Chr(32) 'Leo Espacios
         i=i+1
      Wend
	  
	  While i <= Len(linea) 'Leo NombreArchivo
	     strC = strC & Mid(linea, i, 1)
         i=i+1
      Wend
	  sFileName = strC
      strC=""
	  
   End If

   If bEncontroTexto = 1 Then

    'arrF(x,y) = sRutaArch
    'arrF(x+1,y) = sFileName
    'arrF(x+2,y) = sFecha
    'arrF(x+3,y) = sFileSize
 
    arrCampos.Add trim(sRutaArch) & "\" & trim(sFileName) & "," & trim(sFileSize) & "," & trim(sFecha)

    End If

   bEncontroTexto = 0
   linea = f.ReadLine
Loop

f.Close
'Set fso = Nothing

'Archivo con resultados
'msgbox FilenameJustName & "_aaa.txt"
Set f = FSO.OpenTextFile(FilenameJustName & "_out.csv" , 2, True)

'Archivo con Resultados
'Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
'    Dim fs, f1
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set f1 = fs.OpenTextFile("sale_out.txt", ForWriting, True)

'Set objFSO=CreateObject("Scripting.FileSystemObject")
'outFile="salesale23.txt"
'Set objFile = objFSO.CreateTextFile(outFile,2,True)
'objFile.Write "test string" & vbCrLf

'Escribe Resultados
For Each campo In arrCampos
    f.Write campo & vbCrLf
    'objFile.Write campo & vbCrLf
	'msgbox campo & vbCrLf
Next

'objFile.Close
'f1.Close
Set objFSO= Nothing
'Wscript.Quit

'Chr(34) = "
'Chr(39) = '
'Chr(32) = espacio
'Chr(44) = ,
