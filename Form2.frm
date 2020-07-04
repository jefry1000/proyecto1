VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15870
   LinkTopic       =   "Form2"
   ScaleHeight     =   8460
   ScaleWidth      =   15870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "MODIFICAR"
      Height          =   615
      Left            =   12960
      TabIndex        =   25
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "NUEVO"
      Height          =   615
      Left            =   12960
      TabIndex        =   24
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GUARDAR"
      Height          =   615
      Left            =   12960
      TabIndex        =   23
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   12960
      TabIndex        =   22
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   21
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   20
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "FOTOGRAFIA"
      Height          =   2775
      Left            =   1080
      TabIndex        =   19
      Top             =   5160
      Width           =   3855
      Begin VB.Image Image2 
         Height          =   2295
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1095
      Left            =   5160
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1931
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
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":009F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "animales"
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
   Begin VB.TextBox Text8 
      DataField       =   "lugar_de_origen"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      DataField       =   "edad"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      DataField       =   "alimentacion"
      DataSource      =   "Adodc1"
      Height          =   525
      Left            =   2280
      TabIndex        =   12
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      DataField       =   "peso"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2280
      TabIndex        =   10
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "cantidad"
      DataSource      =   "Adodc1"
      Height          =   525
      Left            =   2280
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "especies"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "codigo"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label11 
      DataField       =   "foto"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   13320
      TabIndex        =   18
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "KG"
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LUGAR DE ORIGEN"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EDAD"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALIMENTACION"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PESO"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ESPECIES"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE ANIMAL"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ANIMALES"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   11160
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image Image1 
      DataSource      =   "Adodc1"
      Height          =   10260
      Left            =   0
      Picture         =   "Form2.frx":013E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15600
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.MoveNext
 
 If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveFirst
Image2.Picture = LoadPicture(X & "\" & Label11.Caption)
X = App.Path
      
    End If
X = App.Path
 Image2.Picture = LoadPicture(X & "\" & Label11.Caption)
    
  
    
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious
'''x = App.Path
 '''Image2.Picture = LoadPicture(x & "\" & Label11.Caption)
 
 If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveLast
    Image2.Picture = LoadPicture(X & "\" & Label11.Caption)
 End If
 
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveFirst
 '''x = App.Path
 '''Image2.Picture = LoadPicture(x & "\" & Label11.Caption)
End Sub

Private Sub Command4_Click()

Adodc1.Recordset.MoveFirst
 ''' x = App.Path
 '''Image2.Picture = LoadPicture(x & "\" & Label11.Caption)
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
 Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command6.Enabled = True

End Sub

Private Sub Command5_Click()
Adodc1.Recordset.AddNew
 Text1.Enabled = True
 Text2.Enabled = True
 Text3.Enabled = True
 Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Command3.Enabled = False
Command2.Enabled = False
Command6.Enabled = False
Command1.Enabled = False
Command5.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Command6_Click()
 Text1.Enabled = True
 Text2.Enabled = True
 Text3.Enabled = True
 Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Command3.Enabled = False
Command2.Enabled = False
Command6.Enabled = False
Command1.Enabled = False

End Sub

Private Sub Form_Load()
 X = App.Path
 Image2.Picture = LoadPicture(X & "\" & Label11.Caption)
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
 Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command6.Enabled = True

End Sub


