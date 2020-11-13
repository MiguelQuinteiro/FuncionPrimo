VERSION 5.00
Begin VB.Form frmFuncionPrimo 
   AutoRedraw      =   -1  'True
   Caption         =   "Función Números Primos"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReducir 
      Caption         =   "Reducir"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtN 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Text            =   "100"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmFuncionPrimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declaración de variables
Dim miN As Long
Dim miMayorGap As Long
Dim miNumeros() As Long
Dim miGaps() As Long

' Reducir la lista de primos al primero de cada orbita
Private Sub cmdReducir_Click()
  Dim i As Long
  Dim j As Long
  Dim miBusqueda As Long

  Open "ArregloGaps.txt" For Output As #1
  Cls
  For i = 0 To miMayorGap
    For j = 1 To miN
      miBusqueda = miGaps(j)
      If miBusqueda = i Then
        Print miNumeros(j), "  es primo, y salió luego de ", miGaps(j)
        Print #1, miNumeros(j), miGaps(j)
        j = miN + 1
      End If
    Next j
  Next i
  Close #1
End Sub

' Calcular los numeros primos hasta cierto valor
Private Sub Command1_Click()
  Dim i As Long
  miN = Val(txtN.Text)
  ReDim miNumeros(miN)
  ReDim miGaps(miN)
  Cls
  miMayorGap = 0
  For i = 1 To miN
    If Primo(i) Then
      miNumeros(i) = i
      miGaps(i) = GapAnterior(i)
      If miMayorGap < miGaps(i) Then
        miMayorGap = miGaps(i)
      End If
      'Print miNumeros(i); "  es primo, y salió luego de "; miGaps(i)
    End If
  Next i
  Print ""
  Print " Listo **************"
End Sub

' FUNCION PARA CALCULAR SI EL NUMERO ES PRIMO
'*****************************************************************************************
' Nombre   : Primo
' Objetivo : Indica si un número dado es primo
' Entradas : pN
' Salida   : Primo (Boolean)
' Sintaxis : Var = Primo(7)
'*****************************************************************************************
Public Function Primo(ByVal pN As Long) As Boolean
  Dim i As Long
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function

' FUNCION PARA CALCULAR EL GAP ANTERIOR
'*****************************************************************************************
' Nombre   : GapAnterior
' Objetivo : Busca al primo posterior e indica a que distancia está
' Entradas : pN
' Salida   : GapAnterior (Long)
' Sintaxis : Var = GapAnterior(7)
'*****************************************************************************************
Function GapAnterior(ByVal pN As Long) As Long
  Dim i As Long
  If Primo(pN) Then
    GapAnterior = 0
    For i = (pN - 1) To 2 Step -1
      GapAnterior = GapAnterior + 1
      If Primo(i) Then
        i = 1
      End If
    Next i
  Else
    GapAnterior = 0
  End If
End Function

' FUNCION PARA CALCULAR EL GAP POSTERIOR
'*****************************************************************************************
' Nombre   : GapPosterior
' Objetivo : Busca al primo posterior e indica a que distancia está
' Entradas : pN
' Salida   : GapPosterior (Long)
' Sintaxis : Var = GapPosterior(7)
'*****************************************************************************************
Function GapPosterior(ByVal pN As Long) As Long
  Dim i As Long
  If Primo(pN) Then
    GapPosterior = 1
    i = pN + 1
    While Not Primo(i)
      GapPosterior = GapPosterior + 1
      i = i + 1
    Wend
  Else
    GapPosterior = 0
  End If
End Function

' FUNCION PARA CALCULAR GAP ENTRE PRIMOS
'*****************************************************************************************
' Nombre   : GapPimos
' Objetivo : Indica si hay un primo a cierta distancia
' Entradas : pN, pG
' Salida   : GapPrimos (Boolean)
' Sintaxis : Var = GapPrimos(7,4)
'*****************************************************************************************
Public Function GapPrimos(ByVal pN As Long, ByVal pG As Long) As Boolean
  Dim i As Long
  If Primo(pN) Then
    If Primo(pN + pG) Then
      GapPrimos = True
    End If

    For i = (pN + 1) To (pN + pG - 1)
      If Primo(i) Then
        GapPrimos = False
      End If
    Next
  Else
    GapPrimos = False
  End If
End Function


'**********************************************************
' Uso de las funciones
'**********************************************************
'    Dim x As Long
'    Dim y As Long
'
'    x = 79
'    Print x
'
'    If Primo(x) Then
'        Print " x es primo"
'    Else
'        Print " x no es primo"
'    End If
'
'    Print
'    y = GapAnterior(x)
'    Print y
'    Print x - y
'
'    Print
'    y = GapPosterior(x)
'    Print y
'    Print x + y
'
'    Print
'    If GapPrimos(x, y) Then
'        Print " tiene un primo a esa distancia "; y
'    Else
'        Print " no tiene un primo a esa distancia "; y
'    End If

