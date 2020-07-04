VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16215
   LinkTopic       =   "Form3"
   ScaleHeight     =   8835
   ScaleWidth      =   16215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C000&
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "MODIFICAR"
      Height          =   615
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "GUARDAR"
      Height          =   615
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "NUEVO"
      Height          =   615
      Left            =   13080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "fotografia"
      Height          =   2295
      Left            =   600
      TabIndex        =   14
      Top             =   4920
      Width           =   3135
      Begin VB.Image Image2 
         Height          =   1920
         Left            =   360
         Picture         =   "Form3.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2400
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   13200
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form3.frx":5C094
      OLEDBString     =   $"Form3.frx":5C133
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "empleados"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text6 
      DataField       =   "fecha_inicio"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      DataField       =   "sueldo"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataField       =   "cargo"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "edad"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre_completo"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "CUI"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      DataField       =   "foto"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   13800
      TabIndex        =   13
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA DE INICIO"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUELDO"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CARGO"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EDAD"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUI"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLEADOS"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   9135
      Left            =   -120
      Picture         =   "Form3.frx":5C1D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.MoveNext

If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveFirst
End If

X = App.Path
Image2.Picture = LoadPicture(X & "\" & Label8.Caption)
 
End Sub


Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious

If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveLast
End If

Image2.Picture = LoadPicture(X & "\" & Label8.Caption)

End Sub

Private Sub Command3_Click()
Adodc1.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False

End Sub

Private Sub Form_Load()
 X = App.Path
 Image2.Picture = LoadPicture(X & "\" & Label8.Caption)
End Sub

