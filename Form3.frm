VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Unistore Creator, Add link and file"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9180
   LinkTopic       =   "Form3"
   ScaleHeight     =   5805
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar y Agregar otro archivo de la misma app o juego (ejmplo: juego.nds)"
      Height          =   975
      Left            =   6000
      TabIndex        =   10
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar y agregar Actualización"
      Height          =   975
      Left            =   6240
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar y agregar otro Juego o App"
      Height          =   975
      Left            =   6240
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Creador del programa: Extintor Incendiandose"
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Configurar Descarga e instalación de Archivo.CIA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   11
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre del archivo.cia tal y como aparece en el enlace, sin espacios ni parentesis"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Mensaje a mostrar en Universal Updater"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Enlace directo permanente (que nunca cambie o expire)"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del cia a descargar (ejemplo: FBI.cia) pdt: se permiten espacios"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fileName As String
Dim nomScrp As String
Dim zelda As String
Dim msgUU As String
Dim ciaNom As String
Dim listo As Boolean

Private Sub Command1_Click()
nomScrp = Text1.Text
zelda = Text2.Text
msgUU = Text3.Text
ciaNom = Text4.Text

' Crear la cadena que se va a anexar al archivo
Dim texto As String
texto = texto & Chr$(34) & nomScrp & Chr$(34) & ":[{"
texto = texto & Chr$(34) & "file" & Chr$(34) & ":" & Chr$(34) & zelda & Chr$(34) & ","
texto = texto & Chr$(34) & "message" & Chr$(34) & ":" & Chr$(34) & msgUU & Chr$(34) & ","
texto = texto & Chr$(34) & "output" & Chr$(34) & ":" & Chr$(34) & "sdmc:/" & ciaNom & Chr$(34) & ","
texto = texto & Chr$(34) & "type" & Chr$(34) & ":" & Chr$(34) & "downloadFile" & Chr$(34) & "},"
texto = texto & "{" & Chr$(34) & "type" & Chr$(34) & ":" & Chr$(34) & "installCia" & Chr$(34) & ","
texto = texto & Chr$(34) & "file" & Chr$(34) & ":" & Chr$(34) & "/" & ciaNom & Chr$(34) & "},"

texto = texto & "{" & Chr$(34) & "type" & Chr$(34) & ":" & Chr$(34) & "deleteFile" & Chr$(34) & ","
texto = texto & Chr$(34) & "file" & Chr$(34) & ":" & Chr$(34) & "sdmc:/" & ciaNom & Chr$(34) & "}]},"


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
listo = True

Dim frm As New Form2
frm.fileName = fileName
frm.listo = listo
frm.Show
Me.Hide

End Sub

Private Sub Command2_Click()
MsgBox "No disponible"
End Sub

Private Sub Command3_Click()
MsgBox "No disponible"
End Sub
