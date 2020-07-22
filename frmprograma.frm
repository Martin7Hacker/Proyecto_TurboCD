VERSION 5.00
Begin VB.Form frmprograma 
   Caption         =   "TurboCD/ROM. RW v1.0"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6360
   Icon            =   "frmprograma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command13 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "&Acerca:"
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Cerrar todo"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Abrir todo "
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "&Ascendente."
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   4920
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Descendente."
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Abrir"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Control"
      Height          =   1095
      Left            =   1440
      TabIndex        =   19
      Top             =   4680
      Width           =   4695
      Begin VB.Line Line1 
         BorderColor     =   &H80000006&
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   1080
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Eliminar Todo"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&renombrar"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Añadir:"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Edición"
      Height          =   1815
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   6135
   End
   Begin VB.ListBox misUnidades 
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Abrir"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   75
   End
   Begin VB.Label labinfo 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   14
      Top             =   3360
      Width           =   75
   End
End
Attribute VB_Name = "frmprograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
' Software escrito por:
' Martin Grasso Castrillo
' donar 1 dolar.
' WhatsApp: +598096922232
' WhatsApp: +598097254018
'**************************
'Función Api getLogicalDrives para recuperar las unidades
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
' Función Api GetDriveType para obtener el tipo y clase de unidad
Private Declare Function GetDriveType _
    Lib "kernel32" _
    Alias "GetDriveTypeA" ( _
    ByVal nDrive As String) As Long
    
    Dim cdrom As New cd_rom
    Dim es As New escripta
    Dim op(7) As String
    Dim cf As Boolean
    
Private Sub cargarCDROM()
Dim LDs As Long, Cnt As Long, sDrives As String
    List1.Clear
    LDs = GetLogicalDrives
    sDrives = "Drives disponibles: "
    For Cnt = 0 To Drive1.ListCount * 2
        If (LDs And 2 ^ Cnt) <> 0 Then
            sDrives = sDrives + " " + Chr$(65 + Cnt)
         If Not (Drive1.List(Cnt) = "") Then
            List1.AddItem UCase(Drive1.List(Cnt))
            End If
            End If
      Next Cnt
End Sub

Private Sub Command10_Click()
If List1.ListIndex = -1 Or List2.ListIndex = -1 Then
 If List1.ListIndex = -1 Then
    MsgBox "Selecionar unidad de CD/ROM disponibles.", vbInformation
  ElseIf List2.ListIndex = -1 Then
    MsgBox "Selecionar unidad de CD/ROM añadidas.", vbInformation
 End If
End If
If Not (List1.ListIndex = -1) And Not (List2.ListIndex = -1) Then
   If (Text2.Text = "") Then
      MsgBox "Ingrese nombre identificativo para " & "( " & List1.List(List1.ListIndex) & " )" & " unidad de CD/ROM.", vbInformation
      Else
      Select Case MsgBox("¿Seguro que Quieres renombrar la unidad " & "( " & List2.List(List2.ListIndex) & " unidad de CD/ROM?.", vbInformation + vbYesNo)
  Case vbYes
  List2.List(List2.ListIndex) = UCase(List1.List(List1.ListIndex) & " ) " & Text2.Text)
  misUnidades.List(List2.ListIndex) = List1.List(List1.ListIndex)
  Text2.Text = ""
  End Select
   End If
End If
End Sub

Private Sub Command11_Click()
If (Text2.Text = "") Then
   MsgBox "No existe descripción que limpiar unidad de CD/ROM.", vbInformation
   Else
   Text2.Text = ""
End If
End Sub

Private Sub Command12_Click()
MsgBox "Programa escrito por: Martin Grasso Castrillo," & vbNewLine & _
       "Colaborar mediante Western Union / Internacional," & vbNewLine & _
       "envio y amistad,regalos contactos Donation (Donación) $1 y proyectos, en general para seguir en mi pasión, http://www.planet-source-code.com/," & vbNewLine & _
       "WhatsApp: +598096922232" & vbNewLine & _
       "WhatsApp: +598097254018" & vbNewLine & _
       "COVID-19 nos ayudamos entre todos bajo la crisis. por nuestros sueños" & vbNewLine & _
       "Echo en Uruguay 2020.", vbInformation
       
End Sub
Private Sub Command13_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Select Case MsgBox("¿Seguro que quieres eliminar todas las unidades creadas de CD/ROM?.", vbYesNo + vbInformation)
  Case (vbYes)
  List2.Clear
  misUnidades.Clear
End Select
End Sub


Private Sub Command1_Click()
 If Not (List2.ListIndex = -1) Then
 Select Case MsgBox("¿Seguro que quieres eliminar esta unidad " & List2.List(List2.ListIndex) & "?.", vbYesNo + vbInformation)
  Case (vbYes)
    misUnidades.RemoveItem (List2.ListIndex)
    List2.RemoveItem (List2.ListIndex)
   End Select
 End If
End Sub

Private Sub Command3_Click()
If Not (TypeOf Screen.ActiveControl Is ListBox) Then
  If Not (List1.ListIndex = -1) Then
   Text2.Text = "(" & List1.List(List1.ListIndex) & " )."
   Label1.Caption = "Abriendo..." & "(" & List1.List(List1.ListIndex) & " )."
   cdrom.abrir List1.List(List1.ListIndex)
   Label1.Caption = "Trabajo Echo."
   labinfo.Caption = ""
   Text2.Text = ""
   Else
   If List1.ListCount = 0 Then
     MsgBox "Imposible utilizar no se encontraron unidades de CD/ROM.", vbInformation
   Else
     mensajecdrom
   End If
   End If
End If
End Sub

Private Sub Command4_Click()
If Not (TypeOf Screen.ActiveControl Is ListBox) Then
 If Not (List1.ListIndex = -1) Then
   Text2.Text = "(" & List1.List(List1.ListIndex) & " )."
   Label1.Caption = "Cerrando..." & "(" & List1.List(List1.ListIndex) & " )."
   cdrom.Cerrar List1.List(List1.ListIndex)
   Label1.Caption = "Trabajo Echo."
   Text2.Text = ""
   labinfo.Caption = ""
   Else
   If List1.ListCount = 0 Then
     MsgBox "Imposible utilizar no se encontraron unidades de CD/ROM.", vbInformation
   Else
     mensajecdrom
   End If
   End If
End If
End Sub

Private Sub Command5_Click()

If List1.ListIndex = -1 Then

MsgBox "Selecionar unidad de CD/ROM.", vbInformation

ElseIf Not (Text2.Text = "") Then


List2.AddItem UCase(List1.List(List1.ListIndex) & ")       " & Text2.Text)
misUnidades.AddItem List1.List(List1.ListIndex)
Text2.Text = ""
List1.ListIndex = -1
Me.Text2.SetFocus
Else
MsgBox "Ingrese nombre Identificativo para " & "( " & List1.List(List1.ListIndex) & " )" & " unidad de CD/ROM.", vbInformation



End If
End Sub



Private Sub Command6_Click()
If Not (TypeOf Screen.ActiveControl Is ListBox) Then
   If Not (List2.ListIndex = -1) Then
      'cdrom.Abrir List2.List(List2.ListIndex)
      Text2.Text = Mid(List2.List(List2.ListIndex), 6, 77)
      info "Abriendo..." & List2.List(List2.ListIndex) & " ."
      cdrom.abrir misUnidades.List(List2.ListIndex)
      Text2.Text = ""
      trabajo_realizado
      Else
  If List2.ListCount = 0 Then
     MsgBox "Imposible utilizar sin Añadir unidades de CD/ROM.", vbInformation
   Else
     mensajecdrom
   End If
End If
End If
End Sub

Private Sub Command7_Click()
If Not (TypeOf Screen.ActiveControl Is ListBox) Then
   If Not (List2.ListIndex = -1) Then
      'cdrom.Cerrar List2.List(List2.ListIndex)
       Text2.Text = Mid(List2.List(List2.ListIndex), 6, 77)
       info "Cerrando..." & List2.List(List2.ListIndex) & " ."
       cdrom.Cerrar misUnidades.List(List2.ListIndex)
       trabajo_realizado
       Text2.Text = ""
       Else
   If List2.ListCount = 0 Then
     MsgBox "Imposible utilizar sin Añadir unidades de CD/ROM.", vbInformation
   Else
     mensajecdrom
   End If
   End If
End If
End Sub

Private Sub Command8_Click()
If Not (List2.ListCount = 0) Then
 AbrirCerrarTodo True
 Else
 mensajecrear
End If
End Sub

Private Sub Command9_Click()
If Not (List2.ListCount = 0) Then
 AbrirCerrarTodo False
 Else
 mensajecrear
End If
End Sub

Private Sub Form_Load()
cargarCDROM
abrir
End Sub

Private Sub Form_Unload(Cancel As Integer)
guardar
End Sub


Private Sub List1_Click()
enfocar List1

End Sub
Private Sub enfocar(ByRef control As ListBox)
control.SetFocus
End Sub

Private Sub List2_Click()
enfocar List2
Text2.Text = Mid(List2.List(List2.ListIndex), 11, 77)
End Sub

Private Sub AbrirCerrarTodo(ByVal abrirOcerrar As Boolean)
On Error GoTo nose
Dim udx As Integer
Dim opx As Integer
Dim infox As Integer

info ""
infox = List2.ListCount
Select Case abrirOcerrar
    Case True
    
    Select Case Option1.Value
    Case (True)
    If Not (TypeOf Screen.ActiveControl Is ListBox) Then
     List2.ListIndex = 1
     If Not (List2.ListCount = 0) Then
        'cdrom.Abrir List2.List(List2.ListIndex)
        For opx = 0 To List2.ListCount
        
       'Me.Caption = "desendente abrir"
        
        If List2.ListIndex = 2 Then
         opx = 2
         Exit For
          Else
         List2.ListIndex = opx
         info "Abriendo..." & List2.List(opx) & " ."
         cdrom.abrir misUnidades.List(opx)
       End If
        
        Next
     End If
    End If
    End Select
    
    Select Case Option2.Value
    Case True
    opx = 0
    udx = List2.ListCount
    If Not (TypeOf Screen.ActiveControl Is ListBox) Then
      If Not (List2.ListCount = 0) Then
        'cdrom.Abrir List2.List(List2.ListIndex)
        udx = List2.ListCount
       
       For opx = 0 To List2.ListCount
           
        
          
         'List2.ListIndex = udx
         
         info "Abriendo..." & List2.List(udx) & " ."
         cdrom.abrir misUnidades.List(udx)
         
         'Me.Caption = 2
        
       
       
       
       udx = udx - 1
       List2.ListIndex = udx
       'info "Abriendo... :" & List2.List(udx) & " ."
        Next
        List2.ListIndex = 0
     End If
    End If
     End Select
     
    Case False
    
     Select Case Option1.Value
     
    Case (True)
    If Not (TypeOf Screen.ActiveControl Is ListBox) Then
     List2.ListIndex = 1
     If Not (List2.ListCount = 0) Then
        'cdrom.Abrir List2.List(List2.ListIndex)
        For opx = 0 To List2.ListCount
        
         'Me.Caption = "desendente cerrar"
        If List2.ListIndex = 2 Then
         opx = 2
         Exit For
          Else
         List2.ListIndex = opx
         info "Cerrando..." & List2.List(opx) & " ."
         cdrom.Cerrar misUnidades.List(opx)
        End If
        
        
        Next
     End If
    End If
    End Select
    
    
    Select Case Option2.Value
    Case True
    opx = 0
    udx = List2.ListCount
    If Not (TypeOf Screen.ActiveControl Is ListBox) Then
      If Not (List2.ListCount = 0) Then
        'cdrom.Abrir List2.List(List2.ListIndex)
        udx = List2.ListCount
       
       For opx = 0 To List2.ListCount
           
        
          
         'List2.ListIndex = udx
         
         info "Cerrando..." & List2.List(udx) & " ."
         cdrom.Cerrar misUnidades.List(udx)
         
         'Me.Caption = 2
        
       
       
       
       udx = udx - 1
       List2.ListIndex = udx
       'info "Abriendo... :" & List2.List(udx) & " ."
        Next
        List2.ListIndex = 0
     End If
    End If

End Select
End Select
trabajo_realizado
nose:
End Sub

Private Sub info(ByVal info As String)
labinfo.Caption = info
Label1.Caption = ""
End Sub

Private Sub mensajecdrom()
MsgBox "Selecionar unidad de CD/ROM.", vbInformation
End Sub

Private Sub mensajecrear()
MsgBox "Imposible de utilizar sin añadir unidades de CD/ROM.", vbInformation
End Sub

Private Sub trabajo_realizado()
 info "Trabajo Echo."
 Text2.Text = ""
End Sub

Private Sub guardar()
On Error GoTo nose:
    
Dim x As Integer ' variable para el for que carga los datos
 Open "configuro.ini" For Output As 1
 For x = 0 To List2.ListCount - 1
 op(0) = es.escriptar(List2.List(x))
 Print #1, op(0)
 op(1) = es.escriptar(misUnidades.List(x))
 Print #1, op(1)
 
 op(2) = es.escriptar(Option1.Value)
 Print #1, op(2)
Next x
Close #1
nose:
End Sub

Private Sub abrir()
On Error GoTo nose
Open "configuro.ini" For Input As 1

Do While Not EOF(1)

Line Input #1, op(3)
                      op(0) = es.desescriptar(op(3))
                      List2.AddItem op(0)

Line Input #1, op(4)
                      op(1) = es.desescriptar(op(4))
                      misUnidades.AddItem op(1)

Line Input #1, op(5)
                      op(2) = es.desescriptar(op(5))
                      Option1.Value = op(2)
                      
                      If op(2) = True Then
                         Option1.Value = True
                       Else
                         Option2.Value = True
                      End If
Loop
Close #1
nose:
End Sub
