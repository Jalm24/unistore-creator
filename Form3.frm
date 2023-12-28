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
      Caption         =   "save and add another file of this package (example: .3dsx .nds .zip .7z)"
      Height          =   855
      Left            =   6120
      TabIndex        =   10
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save and add an update"
      Height          =   615
      Left            =   6480
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
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
      Caption         =   "save and add another app"
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Credits: Extintor Incendiandose"
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "CiaFile: Script Config"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Name of the .cia file as it appears in the download link, without spaces or parentheses"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Message"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Permanent direct download link (that never changes or expires)"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "file name, you can use spaces or parentheses"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
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
' Declarar las variables para almacenar el ancho y el alto de la pantalla
Dim anchoPantalla As Long
Dim altoPantalla As Long

' Crear un procedimiento para el evento Load del formulario secundario
Private Sub Form_Load()
    ' Obtener el ancho y el alto de la pantalla en twips
    anchoPantalla = Screen.Width
    altoPantalla = Screen.Height
    
    ' Centrar el formulario usando las propiedades Left y Top
    Me.Left = (anchoPantalla - Me.Width) / 2
    Me.Top = (altoPantalla - Me.Height) / 2
End Sub


Private Sub Command1_Click()
nomScrp = Text1.Text
zelda = Text2.Text
msgUU = Text3.Text
ciaNom = Text4.Text

' Crear la cadena que se va a anexar al archivo
Dim texto As String
texto = texto & Chr$(34) & "Download " & nomScrp & Chr$(34) & ":[{"
texto = texto & Chr$(34) & "file" & Chr$(34) & ":" & Chr$(34) & zelda & Chr$(34) & ","
texto = texto & Chr$(34) & "message" & Chr$(34) & ":" & Chr$(34) & msgUU & Chr$(34) & ","
texto = texto & Chr$(34) & "output" & Chr$(34) & ":" & Chr$(34) & "sdmc:/" & ciaNom & Chr$(34) & ","
texto = texto & Chr$(34) & "type" & Chr$(34) & ":" & Chr$(34) & "downloadFile" & Chr$(34) & "},"
texto = texto & "{" & Chr$(34) & "type" & Chr$(34) & ":" & Chr$(34) & "installCia" & Chr$(34) & ","
texto = texto & Chr$(34) & "file" & Chr$(34) & ":" & Chr$(34) & "/" & ciaNom & Chr$(34) & "},"

texto = texto & "{" & Chr$(34) & "type" & Chr$(34) & ":" & Chr$(34) & "deleteFile" & Chr$(34) & ","
texto = texto & Chr$(34) & "file" & Chr$(34) & ":" & Chr$(34) & "sdmc:/" & ciaNom & Chr$(34) & "}]},"


' Crear un objeto ADODB.Stream
Dim fsT As Object
Set fsT = CreateObject("ADODB.Stream")

' Especificar el tipo y el conjunto de caracteres
fsT.Type = 2 'Texto
fsT.Charset = "ascii" 'UTF-8

' Abrir el stream en modo de anexar
fsT.Open
fsT.LoadFromFile fileName & ".unistore"
fsT.Position = fsT.Size

' Escribir la cadena en el stream
fsT.WriteText texto

' Guardar el stream en el archivo
fsT.SaveToFile fileName & ".unistore", 2

' Cerrar el stream
fsT.Close

' Liberar los objetos
Set fsT = Nothing
listo = True

Dim frm As New Form2
frm.fileName = fileName
frm.listo = listo
frm.Show
Me.Hide

End Sub

Private Sub Command2_Click()
MsgBox "Not available"
End Sub

Private Sub Command3_Click()
MsgBox "Not available"
End Sub
