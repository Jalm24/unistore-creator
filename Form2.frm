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
      Caption         =   "Finish"
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
      Caption         =   " .FIRM"
      Height          =   735
      Left            =   7080
      TabIndex        =   24
      Top             =   5520
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   ".3dsx"
      Height          =   735
      Left            =   5400
      TabIndex        =   23
      Top             =   5520
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "CIA"
      Height          =   735
      Left            =   2280
      TabIndex        =   22
      Top             =   5520
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   " NDS"
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
      Caption         =   "add download link"
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
      Width           =   2295
   End
   Begin VB.Label Label12 
      Caption         =   "credits: Extintor Incendiandose"
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Application details"
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
      Caption         =   "licence"
      Height          =   255
      Left            =   6000
      TabIndex        =   20
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "last update  MM-DD-YY"
      Height          =   495
      Left            =   5880
      TabIndex        =   19
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "description"
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "sheet index (pls read documentation)"
      Height          =   495
      Left            =   5280
      TabIndex        =   17
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "icon index (pls read documentation)"
      Height          =   495
      Left            =   5280
      TabIndex        =   16
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "console"
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "category"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "author"
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "version"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "App name"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
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
texto = texto & "{" & Chr$(34) & "info" & Chr$(34) & ": {"
texto = texto & Chr$(34) & "title" & Chr$(34) & ": " & Chr$(34) & title & Chr$(34) & ","
texto = texto & Chr$(34) & "version" & Chr$(34) & ": " & Chr$(34) & ver & Chr$(34) & ","
texto = texto & Chr$(34) & "author" & Chr$(34) & ": " & Chr$(34) & author & Chr$(34) & ","
texto = texto & Chr$(34) & "category" & Chr$(34) & ": " & Chr$(34) & category & Chr$(34) & ","
texto = texto & Chr$(34) & "console" & Chr$(34) & ": [" & Chr$(34) & console & Chr$(34) & "  ],"
texto = texto & Chr$(34) & "icon_index" & Chr$(34) & ": " & iconIndex & ","
texto = texto & Chr$(34) & "sheet_index" & Chr$(34) & ": " & sheetIndex & ","
texto = texto & Chr$(34) & "description" & Chr$(34) & ": " & Chr$(34) & description & Chr$(34) & ","
texto = texto & Chr$(34) & "last_updated" & Chr$(34) & ": " & Chr$(34) & updated & Chr$(34) & ","
texto = texto & Chr$(34) & "license" & Chr$(34) & ": " & Chr$(34) & license & Chr$(34) & "  },"

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

If (Option2.Value = True) Then
' Crear una instancia de Form2
Dim result As VbMsgBoxResult
result = MsgBox("Changes saved" & Chr(13) & "It's time to add a .cia", vbOKOnly + vbInformation, "Confirm")
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
MsgBox "You haven't added any links!!" & Chr(13) & "without links you can't finish", vbOKOnly + vbCritical, "DANGER"
Else

    Dim fsT As Object
    Dim txt As String
    
    'Crear un objeto ADODB.Stream
    Set fsT = CreateObject("ADODB.Stream")
    
    'Especificar el tipo y el conjunto de caracteres
    fsT.Type = 2 'Texto
    fsT.Charset = "ascii" 'UTF-8
    
    'Especificar el separador de líneas
    fsT.LineSeparator = -1 'CR LF
    
    'Abrir el stream y leer el contenido del archivo
    fsT.Open
    fsT.LoadFromFile fileName & ".unistore"
    
    'Guardar el contenido del stream en una variable
    txt = fsT.ReadText
    
    'Cerrar el stream
    fsT.Close
    
    'Borrar el último carácter del texto usando la función Left
    txt = Left(txt, Len(txt) - 1)
    
    'Agregar la cadena de texto al final del texto usando el operador &
    txt = txt & "]}"
    
    'Abrir el stream en modo de escritura
    fsT.Open
    
    'Escribir el nuevo texto en el stream
    fsT.WriteText txt
    
    'Guardar el stream en el archivo
    fsT.SaveToFile fileName & ".unistore", 2
    
    'Cerrar el stream
    fsT.Close
    
    'Mostrar el nuevo texto en un cuadro de mensaje
    MsgBox "File saved successfully", vbOKOnly + vbInformation, "thanks :3"


    End
    
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
