VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Unistore Creator, Add App/Game"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   ScaleHeight     =   7905
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Finalizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   25
      Top             =   6720
      Width           =   1935
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Archivo .FIRM"
      Height          =   735
      Left            =   7080
      TabIndex        =   24
      Top             =   5520
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Archivo .3dsx"
      Height          =   735
      Left            =   5400
      TabIndex        =   23
      Top             =   5520
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Archivo CIA"
      Height          =   735
      Left            =   2280
      TabIndex        =   22
      Top             =   5520
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Archivo NDS"
      Height          =   735
      Left            =   720
      TabIndex        =   21
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Text            =   "0"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Text            =   "0"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Text            =   "v0.0.1"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar Enlace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Creador del programa: Extintor Incendiandose"
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Detalles de la aplicación o videojuego"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   26
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label10 
      Caption         =   "licencia de la app o videojuego"
      Height          =   495
      Left            =   5280
      TabIndex        =   20
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "ultima vez que se actualizo MM-DD-YY"
      Height          =   495
      Left            =   5160
      TabIndex        =   19
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "descripción de la app o videojuego"
      Height          =   495
      Left            =   5160
      TabIndex        =   18
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "sheet index (leer documentación)"
      Height          =   495
      Left            =   5280
      TabIndex        =   17
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "icon index (leer documentación)"
      Height          =   495
      Left            =   5280
      TabIndex        =   16
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "consola"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "categoria"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "autor de la app o videojuego"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "version de la app o videojuego"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de la app o videojuego"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fileName As String
Public listo As Boolean
Dim title As String
Dim ver As String
Dim author As String
Dim category As String
Dim console As String
Dim iconIndex As Integer
Dim sheetIndex As Integer
Dim description As String
Dim updated As String
Dim license As String
Dim addCia As Boolean

Private Sub Command1_Click()
title = Text1.Text
ver = Text2.Text
author = Text3.Text
category = Text4.Text
console = Text5.Text
iconIndex = Text6.Text
sheetIndex = Text7.Text
description = Text8.Text
updated = Text9.Text
license = Text10.Text

' Crear la cadena que se va a anexar al archivo
Dim texto As String
texto = texto & "{" & Chr$(34) & "info" & Chr$(34) & ":{"
texto = texto & Chr$(34) & "title" & Chr$(34) & ":" & Chr$(34) & title & Chr$(34) & ","
texto = texto & Chr$(34) & "version" & Chr$(34) & ":" & Chr$(34) & ver & Chr$(34) & ","
texto = texto & Chr$(34) & "author" & Chr$(34) & ":" & Chr$(34) & author & Chr$(34) & ","
texto = texto & Chr$(34) & "category" & Chr$(34) & ":" & Chr$(34) & category & Chr$(34) & ","
texto = texto & Chr$(34) & "console" & Chr$(34) & ":[" & Chr$(34) & console & Chr$(34) & "],"
texto = texto & Chr$(34) & "icon_index" & Chr$(34) & ":" & iconIndex & ","
texto = texto & Chr$(34) & "sheet_index" & Chr$(34) & ":" & sheetIndex & ","
texto = texto & Chr$(34) & "description" & Chr$(34) & ":" & Chr$(34) & description & Chr$(34) & ","
texto = texto & Chr$(34) & "last_updated" & Chr$(34) & ":" & Chr$(34) & updated & Chr$(34) & ","
texto = texto & Chr$(34) & "license" & Chr$(34) & ":" & Chr$(34) & license & Chr$(34) & "},"

' Crear un objeto FileSystemObject
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

' Abrir el archivo prueba.txt en modo de anexar
Dim ts As Object
Set ts = fso.OpenTextFile(fileName & ".unistore", 8, True)

' Escribir la cadena en el archivo
ts.Write texto

' Cerrar el archivo
ts.Close

' Liberar los objetos
Set ts = Nothing
Set fso = Nothing

If (Option2.Value = True) Then
' Crear una instancia de Form2
Dim result As VbMsgBoxResult
result = MsgBox("Cambios guardados" & Chr(13) & "toca agregar un cia", vbOKOnly + vbInformation, "Confirmación")
If result = vbOK Then
Dim frm As New Form3
' Mostrar Form2
frm.fileName = fileName
frm.Show
' Ocultar Form1
Me.Hide
End If
End If

End Sub

Private Sub Command2_Click()
If (listo = False) Then
MsgBox "No haz agregado ningun enlace!" & Chr(13) & "sin enlaces no se puede finalizar", vbOKOnly + vbCritical, "PELIGRO"
Else

    Dim fso As Object
    Dim ts As Object
    Dim txt As String
    
    'Crear un objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Abrir el archivo de texto en modo lectura
    Set ts = fso.OpenTextFile(fileName & ".unistore", 1)
    
    'Leer el contenido del archivo y guardarlo en una variable
    txt = ts.ReadAll
    
    'Cerrar el archivo de texto
    ts.Close
    
    'Borrar el último carácter del texto usando la función Left
    txt = Left(txt, Len(txt) - 1)
    
    'Agregar la cadena de texto al final del texto usando el operador &
    txt = txt & "]}"
    
    'Abrir el archivo de texto en modo escritura
    Set ts = fso.OpenTextFile(fileName & ".unistore", 2)
    
    'Escribir el nuevo texto en el archivo de texto
    ts.Write txt
    
    'Cerrar el archivo de texto
    ts.Close
    
    'Mostrar el nuevo texto en un cuadro de mensaje
    MsgBox "Archivo guardado con exito", vbOKOnly + vbInformation, "Gracias :3"
    End
    
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
