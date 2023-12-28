VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Unistore Creator"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9180
   LinkTopic       =   "Form4"
   ScaleHeight     =   6030
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Universal Updater Discord Server"
      Height          =   855
      Left            =   5760
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Universal Updater Github"
      Height          =   855
      Left            =   3360
      TabIndex        =   6
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Universal Updater Webpage"
      Height          =   855
      Left            =   960
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Unistore Creator Docs"
      Height          =   855
      Left            =   5760
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update Unistore"
      Height          =   855
      Left            =   3360
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Unistore"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "This software is not maintained by Universal Team, any problem please report it to jalm24"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "make a unistore in minutes!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Welcome to Unistore Creator! "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   4575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declarar la función ShellExecute de la API de Windows
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Asignar una constante para el modo de mostrar la ventana
Private Const SW_SHOWNORMAL As Long = 1

' Declarar las variables para almacenar el ancho y el alto de la pantalla
Dim anchoPantalla As Long
Dim altoPantalla As Long

' Crear un procedimiento para el evento Resize del formulario
Private Sub Form_Resize()
    ' Obtener el ancho y el alto de la pantalla en twips
    anchoPantalla = Screen.Width
    altoPantalla = Screen.Height
    
    ' Centrar el formulario usando las propiedades Left y Top
    Me.Left = (anchoPantalla - Me.Width) / 2
    Me.Top = (altoPantalla - Me.Height) / 2
End Sub

Private Sub Command1_Click()
Dim frm As New Form1
frm.Show
Me.Hide
End Sub

Private Sub Command2_Click()
MsgBox "work in progress", vbOKOnly + vbInformation, "sorry :c"
End Sub

Private Sub Command3_Click()
    ' Obtener la dirección de la página web que se quiere abrir
    Dim url As String
    url = "https://github.com/Jalm24/unistore-creator/wiki"
    
    ' Llamar a la función ShellExecute para abrir la página web en el navegador predeterminado
    ShellExecute Me.hwnd, "open", url, "", "C:\", SW_SHOWNORMAL
End Sub

Private Sub Command4_Click()
    ' Obtener la dirección de la página web que se quiere abrir
    Dim url As String
    url = "https://universal-team.net/projects/universal-updater.html"
    
    ' Llamar a la función ShellExecute para abrir la página web en el navegador predeterminado
    ShellExecute Me.hwnd, "open", url, "", "C:\", SW_SHOWNORMAL
End Sub

Private Sub Command5_Click()
    ' Obtener la dirección de la página web que se quiere abrir
    Dim url As String
    url = "https://github.com/Universal-Team/Universal-Updater"
    
    ' Llamar a la función ShellExecute para abrir la página web en el navegador predeterminado
    ShellExecute Me.hwnd, "open", url, "", "C:\", SW_SHOWNORMAL
End Sub

Private Sub Command6_Click()
    ' Obtener la dirección de la página web que se quiere abrir
    Dim url As String
    url = "https://universal-team.net/discord"
    
    ' Llamar a la función ShellExecute para abrir la página web en el navegador predeterminado
    ShellExecute Me.hwnd, "open", url, "", "C:\", SW_SHOWNORMAL
End Sub
