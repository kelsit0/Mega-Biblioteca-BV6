VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   15
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmd_guardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   14
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox txt_prestado_a 
      Height          =   525
      Left            =   3960
      TabIndex        =   13
      Top             =   4560
      Width           =   4215
   End
   Begin VB.CheckBox ch_prestado 
      Caption         =   "Prestado actualmente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2040
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CheckBox ch_recomendado 
      Caption         =   "Recomendado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   11
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CheckBox ch_quiero_leer 
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
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CheckBox ch_leido 
      Caption         =   "Ya leido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txt_calificacion 
      Height          =   525
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   615
   End
   Begin VB.ComboBox cmb_genero 
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox txt_autor 
      Height          =   525
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.TextBox txt_titulo 
      Height          =   525
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label8 
      Caption         =   "Calif"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label T 
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EditandoID As Integer

Private Sub ch_leido_Click()
    If ch_leido.Value = 1 Then
        ch_quiero_leer.Value = 0
        txt_calificacion.Enabled = True
    Else
        txt_calificacion.Enabled = False
    End If
End Sub

Private Sub ch_prestado_Click()
    If ch_prestado.Value = 1 Then
        txt_prestado_a.Enabled = True
    Else
        txt_prestado_a.Enabled = False
        txt_prestado_a.Text = ""
    End If
End Sub

Private Sub ch_quiero_leer_Click()
    If ch_quiero_leer.Value = 1 Then
        ch_leido.Value = 0
    End If
End Sub

Private Sub cmd_cancelar_Click()
    Unload Me
End Sub

Private Sub cmd_guardar_Click()
    If Trim(txt_titulo.Text) = "" Or Trim(txt_autor.Text) = "" Then
        MsgBox "El titulo y el autor son obligatorios", vbExclamation, "Datos Incompletos"
        Exit Sub
    End If
    
    If cmb_genero.ListIndex = -1 Then
        MsgBox "Seleccione un genero", vbExclamation, "Datos Incompletos"
        Exit Sub
    End If
    
    If ch_leido.Value = 1 And Trim(txt_calificacion.Text) = "" Then
        MsgBox "Por favor ingrese una calificacion(1-5)", vbInformation
    End If
    
    'Calif 1-5
    Dim calif As Variant
    If Trim(txt_calificacion.Text) <> "" Then
        calif = Val(txt_calificacion.Text)
        If (calif < 1 Or calif > 5) Then
            MsgBox "Calificacion debe ser un numero del 1 al 5.", vbExclamation
        End If
    Else
        calif = "NULL"
    End If
    
    Dim titulo As String, autor As String, generoID As Long
    titulo = Replace(txt_titulo.Text, "'", "''")
    autor = Replace(txt_autor.Text, "'", "''")
    generoID = cmb_genero.ItemData(cmb_genero.ListIndex)
    
    Dim leido As Integer, porLeer As Integer, recom As Integer, prestado As Integer
    leido = IIf(ch_leido.Value = 1, 1, 0)
    porLeer = IIf(ch_quiero_leer.Value = 1, 1, 0)
    recom = IIf(ch_recomendado.Value = 1, 1, 0)
    prestado = IIf(ch_prestado.Value = 1, 1, 0)
    
    Dim prestadoA As String, FechaPrestamo As String
    If prestado = 1 Then
        prestadoA = Replace(txt_prestado_a.Text, "'", "''")
        FechaPrestamo = Format$(Now, "yyyy-mm-dd")
    Else
        prestadoA = ""
        FechaPrestamo = ""
    End If

    On Error GoTo ErrSave
    
        Dim sqlInsert As String
        If EditandoID = 0 Then
    ' AGREGAR
    sqlInsert = "INSERT INTO Libros (Titulo, Autor, GeneroID, Leido, PorLeer, Recomendado, Prestado, PrestadoA, FechaPrestamo, Calificacion) VALUES (" & _
        "'" & titulo & "', '" & autor & "', " & CStr(generoID) & ", " & _
        CStr(leido) & ", " & CStr(porLeer) & ", " & CStr(recom) & ", " & CStr(prestado) & ", "
    
    If prestado = 1 Then
        sqlInsert = sqlInsert & "'" & prestadoA & "', '" & FechaPrestamo & "', "
    Else
        sqlInsert = sqlInsert & "NULL, NULL, "
    End If

    If calif = "NULL" Then
        sqlInsert = sqlInsert & "NULL)"
    Else
        sqlInsert = sqlInsert & CStr(calif) & ")"
    End If

    conn.Execute sqlInsert
    MsgBox "Libro agregado exitosamente", vbInformation
    Exit Sub
Else
    ' EDITAR
    sqlInsert = "UPDATE Libros SET " & _
        "Titulo = '" & titulo & "', " & _
        "Autor = '" & autor & "', " & _
        "GeneroID = " & CStr(generoID) & ", " & _
        "Leido = " & CStr(leido) & ", " & _
        "PorLeer = " & CStr(porLeer) & ", " & _
        "Recomendado = " & CStr(recom) & ", " & _
        "Prestado = " & CStr(prestado) & ", "

    If prestado = 1 Then
        sqlInsert = sqlInsert & "PrestadoA = '" & prestadoA & "', FechaPrestamo = '" & FechaPrestamo & "', "
    Else
        sqlInsert = sqlInsert & "PrestadoA = NULL, FechaPrestamo = NULL, "
    End If

    If calif = "NULL" Then
        sqlInsert = sqlInsert & "Calificacion = NULL "
    Else
        sqlInsert = sqlInsert & "Calificacion = " & CStr(calif) & " "
    End If

    sqlInsert = sqlInsert & "WHERE LibroID = " & EditandoID

    conn.Execute sqlInsert
    MsgBox "Libro editado exitosamente", vbInformation
    Exit Sub
End If

        
ErrSave:
    MsgBox "Ocurrio un Error al guardar: " & Err.Description, vbCritical
    
End Sub

Private Sub Form_Load()
    
    Dim rsG As ADODB.Recordset
    Set rsG = New ADODB.Recordset
    rsG.Open "SELECT GeneroID, Nombre FROM Generos ORDER BY Nombre", conn, adOpenStatic, adLockReadOnly
    cmb_genero.Clear
    Do Until rsG.EOF
        cmb_genero.AddItem rsG!Nombre
        cmb_genero.ItemData(cmb_genero.NewIndex) = rsG!generoID
        rsG.MoveNext
    Loop
    rsG.Close: Set rsG = Nothing
    
    If EditandoID = 0 Then
        ' Modo agregar
        txt_titulo.Text = ""
        txt_autor.Text = ""
        cmb_genero.ListIndex = -1
        txt_calificacion = ""
        ch_leido.Value = 0
        txt_prestado_a.Enabled = False
        Me.Caption = "Agregar Libro"
    Else
     Me.Caption = "Editar Libro"
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM Libros WHERE LibroID = " & EditandoID, conn, adOpenStatic, adLockReadOnly

        If Not rs.EOF Then
            txt_titulo.Text = rs!titulo
            txt_autor.Text = rs!autor

            ' Buscar el índice correcto del género en el ComboBox
            Dim i As Integer
            For i = 0 To cmb_genero.ListCount - 1
                If cmb_genero.ItemData(i) = rs!generoID Then
                    cmb_genero.ListIndex = i
                    Exit For
                End If
            Next i
            If IsNull(rs!leido) Then
                ch_leido.Value = 0
            Else
                ch_leido.Value = IIf(rs!leido, 1, 0)
            End If

        If IsNull(rs!porLeer) Then
            ch_quiero_leer.Value = 0
        Else
            ch_quiero_leer.Value = IIf(rs!porLeer, 1, 0)
        End If

        If IsNull(rs!Recomendado) Then
            ch_recomendado.Value = 0
        Else
            ch_recomendado.Value = IIf(rs!Recomendado, 1, 0)
        End If

        If IsNull(rs!prestado) Then
            ch_prestado.Value = 0
        Else
        ch_prestado.Value = IIf(rs!prestado, 1, 0)
        End If

            If rs!prestado = True Then
                txt_prestado_a.Text = rs!prestadoA
                txt_prestado_a.Enabled = True
            End If

            If Not IsNull(rs!Calificacion) Then
                txt_calificacion.Text = rs!Calificacion
                txt_calificacion.Enabled = True
            End If
        End If
        rs.Close: Set rs = Nothing
    End If
    
End Sub
