VERSION 5.00
Begin VB.Form frmPRUEBA 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "OLAP"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Normal"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "ADODB.Connection"
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmPRUEBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFORM() As ExploradorActiveXForm
Private intCONT As Integer

Private Sub Command1_Click()
Dim objS
Dim objM
Dim objX

    intCONT = intCONT + 1
    ReDim Preserve clsFORM(1 To intCONT)
    Set clsFORM(intCONT) = New ExploradorActiveXForm
    Set clsFORM(intCONT).Objeto = CreateObject(Text1.Text)
    If Not clsFORM(intCONT).Muestra Then
        clsFORM(intCONT).Descarga
        Set clsFORM(intCONT) = Nothing
        intCONT = intCONT - 1
        MsgBox "IMPOSIBLE MOSTRAR EL OBJETO"
    End If
End Sub

Private Sub Command2_Click()
    frmOLAP.Show
End Sub

Private Sub Form_Load()
    intCONT = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim a
    If intCONT > 0 Then
        For a = 1 To intCONT Step 1
            clsFORM(a).Descarga
            Set clsFORM(a) = Nothing
        Next
    End If
End Sub
