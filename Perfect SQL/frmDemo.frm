VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo SQL Structure - HACKPRO TM"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd8 
      Caption         =   "&BETWEEN"
      Height          =   390
      Left            =   3720
      TabIndex        =   7
      Top             =   615
      Width           =   1365
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "SELECT &JOIN"
      Height          =   390
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "&MULTI SELECT"
      Height          =   390
      Left            =   2265
      TabIndex        =   6
      Top             =   615
      Width           =   1365
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "&COMMAS"
      Height          =   390
      Left            =   1185
      TabIndex        =   5
      Top             =   615
      Width           =   990
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "&INSERT INTO"
      Height          =   390
      Left            =   2265
      TabIndex        =   2
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "&UPDATE"
      Height          =   390
      Left            =   1185
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "&DELETE"
      Height          =   390
      Left            =   105
      TabIndex        =   4
      Top             =   615
      Width           =   990
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "&SELECT"
      Height          =   390
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1515
      Width           =   4965
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5070
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result SQL"
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
      TabIndex        =   8
      Top             =   1275
      Width           =   975
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************'
'* Programmed by HACKPRO TM © Copyright 2005  *'
'* Programado por HACKPRO TM © Copyright 2005 *'
'**********************************************'
Option Explicit

 Private xText As String, yText

Private Sub cmd1_Click()
 '* This a Simple SELECT sentence.
 xText = modSQL.Get_Commas(False, "Name", "Company")
 yText = modSQL.Get_Commas(True, "Heriberto", "HACKPRO TM")
 txtSQL.Text = modSQL.Get_Select("Empleados", "*", modSQL.Get_Mult_Set(xText, yText, "and", "like", False), "Nombre")
End Sub

Private Sub cmd2_Click()
 '* UPDATE sentence.
 xText = "Name, Company"
 yText = "Heriberto, HACKPRO TM"
 txtSQL.Text = modSQL.Get_Update("Clients", modSQL.Get_Mult_Set(xText, yText), modSQL.Get_Simp_Set("id", 125))
End Sub

Private Sub cmd3_Click()
 '* INSERT INTO sentence.
 xText = modSQL.Get_Commas(False, "Name", "Company")
 yText = modSQL.Get_Commas(True, "Heriberto", "HACKPRO TM")
 txtSQL.Text = modSQL.Get_Insert("Clients", xText, yText)
End Sub

Private Sub cmd4_Click()
 '* SELECT two charts for their id's.
 xText = "Name"
 yText = "!!Company"
 txtSQL.Text = modSQL.Get_Select_Join("Clients", "Personal", "id_client", "id_personal", modSQL.Get_Mult_Set(xText, yText), modSQL.Get_Simp_Set("id_client", "25"))
End Sub

Private Sub cmd5_Click()
 '* DELETE sentence.
 txtSQL.Text = modSQL.Get_Delete("Clients", , modSQL.Get_Simp_Set("id", 125))
End Sub

Private Sub cmd6_Click()
 '* Example to set commas.
 xText = modSQL.Get_Commas(False, "Name", "Last Name", "Phone")
 yText = modSQL.Get_Commas(True, "Heriberto", "Mantilla", "!!6211277")
 txtSQL.Text = xText & vbCrLf & yText
End Sub

Private Sub cmd7_Click()
 '* MULTIPLE SELECT sentence.
 xText = "Name, Age, Date"
 yText = "Jose, !!30, !!now()"
 xText = modSQL.Get_Mult_Set(xText, yText, "and")
 txtSQL.Text = modSQL.Get_Select("Clients", "*", xText)
End Sub

Private Sub cmd8_Click()
 '* BETWEEN sentence.
 xText = modSQL.Get_Between("#04/1/95#", "#07/1/95#")
 yText = "[ID de cliente] IN (SELECT [ID de cliente] FROM Pedidos WHERE FechaPedido"
 txtSQL.Text = modSQL.Get_Select("Clientes", "NombreContacto, Compañía, CargoContacto, Teléfono", yText & xText & ")")
End Sub

Private Sub Form_Load()
 modSQL.Delimiter = ","
End Sub
