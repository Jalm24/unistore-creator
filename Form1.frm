VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Unistore Creator"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar apps o juegos"
      Height          =   855
      Left            =   3960
      TabIndex        =   19
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      Text            =   "0"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Text            =   "0"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear archivo base"
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Creador del programa: Extintor Incendiandose"
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Detalles de la Tienda"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   20
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label9 
      Caption         =   "revision (solo numeros enteros)"
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   " version (solo numeros enteros)"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "descripción de la tienda"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "url del archivo de imagenes.t3x"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "archivo de imagenes.t3x"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "nombre del archivo.unistore"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "url"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "autor"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de la tienda"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nomTienda As String
Dim autor As String
Dim url As String
Dim fileName As String
Dim imgName As String
Dim urlImg As String
Dim descTienda As String
Dim ver As Integer
Dim rev As Integer


Private Sub Command1_Click()
nomTienda = Text1.Text
autor = Text2.Text
url = Text3.Text
fileName = Text4.Text
imgName = Text5.Text
urlImg = Text6.Text
descTienda = Text7.Text
ver = Text8.Text
rev = Text9.Text

Open fileName & ".unistore" For Output As #1
Print #1, "{"
Print #1, Chr$(34) & "storeInfo" & Chr$(34) & ":{"
Print #1, Chr$(34) & "title" & Chr$(34) & ":" & Chr$(34) & nomTienda & Chr$(34) & ","
Print #1, Chr$(34) & "author" & Chr$(34) & ":" & Chr$(34) & autor & Chr$(34) & ","
Print #1, Chr$(34) & "url" & Chr$(34) & ":" & Chr$(34) & url & Chr$(34) & ","
Print #1, Chr$(34) & "file" & Chr$(34) & ":" & Chr$(34) & fileName & Chr$(34) & ","
Print #1, Chr$(34) & "sheet" & Chr$(34) & ":" & Chr$(34) & imgName & Chr$(34) & ","
Print #1, Chr$(34) & "sheetURL" & Chr$(34) & ":" & Chr$(34) & urlImg & Chr$(34) & ","
Print #1, Chr$(34) & "description" & Chr$(34) & ":" & Chr$(34) & descTienda & Chr$(34) & ","
Print #1, Chr$(34) & "version" & Chr$(34) & ":" & Chr$(34) & ver & Chr$(34) & ","
Print #1, Chr$(34) & "revision" & Chr$(34) & ":" & Chr$(34) & rev & Chr$(34) & "},"
Print #1, Chr$(34) & "storeContent" & Chr$(34) & ": ["
Close #1

MsgBox "Ya se ha creado el archivo base," & Chr(13) & "por favor agrega un juego", vbInformation + vbOKOnly, "Ïnformación Importante"
End Sub

Private Sub Command2_Click()
Dim frm As New Form2
frm.fileName = fileName
frm.Show
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
