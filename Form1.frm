VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   14340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_eliminar 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmd_editar 
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton btn_fav 
      Caption         =   "Libros Favoritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton btn_genfav 
      Caption         =   "Generos Favoritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton btn_no_gustar 
      Caption         =   "No me gustó"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton btn_quiero 
      Caption         =   "Quiero Leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton btn_agregar 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   6360
      Width           =   2175
   End
   Begin MSComctlLib.ListView List_libros 
      Height          =   5775
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton btn_leiste 
      Caption         =   "Ya Leiste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton btn_catalogo 
         Caption         =   "Catalogo MEGA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CargarLibros(filtroSQL As String)
    Dim rs As ADODB.Recordset
    Dim sql As String
    sql = "SELECT L.LibroID, L.Titulo, L.Autor, G.Nombre AS Genero, L.PrestadoA, L.Calificacion, L.Prestado " & _
      "FROM Libros L INNER JOIN Generos G ON L.GeneroID = G.GeneroID "

If filtroSQL <> "" Then
    sql = sql & "WHERE " & filtroSQL
End If

    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    List_libros.ListItems.Clear
    
    If Not rs.EOF Then
        Dim item As ListItem
        Do Until rs.EOF
        
        Set item = List_libros.ListItems.Add(, , rs!titulo)
        item.SubItems(1) = rs!autor
        item.SubItems(2) = rs!Genero
        item.SubItems(3) = IIf(IsNull(rs!Calificacion), "", rs!Calificacion)
        If rs!prestado = True Then
            item.SubItems(4) = rs!prestadoA
        Else
            item.SubItems(4) = ""
        End If
        item.Tag = rs!libroID
        rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    
End Sub

Private Sub btn_agregar_Click()
    Form2.EditandoID = 0
    Form2.Show vbModal
End Sub

Private Sub btn_catalogo_Click()
    CargarLibros ""
End Sub

Private Sub btn_fav_Click()
    CargarLibros "L.Recomendado=1"
End Sub

Private Sub btn_genfav_Click()
    CargarLibros "G.EsFavorito=1"
End Sub

Private Sub btn_leiste_Click()
    
    CargarLibros "L.Leido=1"
End Sub

Private Sub btn_no_gustar_Click()
    CargarLibros "L.Leido=1 AND L.Calificacion <=2"
End Sub

Private Sub btn_quiero_Click()
    CargarLibros "L.PorLeer=1"
End Sub

Private Sub cmd_editar_Click()
    Form2.EditandoID = List_libros.SelectedItem.Tag
    Form2.Show vbModal
    
End Sub

Private Sub cmd_eliminar_Click()
    Dim item As ListItem
    Set item = List_libros.SelectedItem
    
    If item Is Nothing Then
        MsgBox "Selecciona el libro a eliminar", vbExclamation
        Exit Sub
    End If
    
    Dim titulo As String
    titulo = item.Text
    Dim resp As Integer
    resp = MsgBox("¿Estas seguro de eliminar el libro?" & titulo, vbYesNo + vbQuestion, "Confirmar eliminacion")
    
    If resp = vbYes Then
        Dim libroID As Long
        libroID = item.Tag
        On Error GoTo ErrorDelete
        conn.Execute "DELETE FROM Libros WHERE LibroID=" & CStr(libroID)
        MsgBox "Libro Eliminado.", vbInformation
        CargarLibros ""
    End If
    Exit Sub
ErrorDelete:
    MsgBox "Error eliminando libro: " & Err.Description, vbCritical
    
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient
    Dim connString As String
    
    connString = "Provider=SQLOLEDB.1;Data Source=LAPTOP-RCH3MGG7;Initial Catalog=MegaLibreria;Integrated Security=SSPI;"
        
    conn.Open connString
    
    With List_libros
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Titulo", 2000
        .ColumnHeaders.Add , , "Autor", 1500
        .ColumnHeaders.Add , , "Genero", 1000
        .ColumnHeaders.Add , , "Calif", 800
        .ColumnHeaders.Add , , "Prestado a", 1500
    End With
End Sub

