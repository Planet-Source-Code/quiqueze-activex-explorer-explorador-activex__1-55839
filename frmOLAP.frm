VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOLAP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OLAP Objects"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSHOW 
      Caption         =   "Mostrar"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.TreeView tvLISTA 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10186
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox txtSERVER 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdCONNECT 
      Caption         =   "Conectar"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmOLAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'******************************************
'* Referencias:
'*    ExploradorActivex.dll
'*
'* Objeto de prueba usado: DSO.Server (Analysis Services)
'******************************************

' Variables
Public flagCONN As Boolean
Public intNUM As Integer

' Objetos
Private dsoSERVER
Private dsoBASE
Private dsoCUBE
Private dsoDIM
Private dsoLEVEL
Private dsoPROP

Private clsFORM() As ExploradorActiveX.ExploradorActiveXForm

' Apuntadores
Private Bases As Collection
Private Cubos As Collection
Private Dimensiones As Collection
Private Niveles As Collection
Private Propiedades As Collection
Private Objetos As Collection

Private Sub cmdCONNECT_Click()
' Carga
Dim xFOLDER As Node
Dim xNODO As Node

    If flagCONN Then
        If Trim(txtSERVER.Text) = vbNullString Then
            Exit Sub
        End If

        'Set dsoSERVER = New DSO.Server
        Set dsoSERVER = CreateObject("DSO.Server")
        dsoSERVER.Connect txtSERVER.Text
        
        tvLISTA.Nodes.Clear
        Set xNODO = tvLISTA.Nodes.Add(, , dsoSERVER.Name, dsoSERVER.Name)
        xNODO.Expanded = True
        
        Objetos.Add dsoSERVER, dsoSERVER.Name
        
        ' Agrega la bases
        AddBases dsoSERVER.MDStores, xNODO.Key
        
        flagCONN = False
        cmdCONNECT.Caption = "Desconectar"
    Else
        dsoSERVER.CloseServer
        Set dsoSERVER = Nothing
        
        flagCONN = True
        cmdCONNECT.Caption = "Conectar"
    End If
End Sub

Private Sub AddBases(Coleccion, Nodo As String)
' Agrega una coleccion debajo del nodo
Dim xNODO As Node
Dim yNODO As Node
    
    Set xNODO = tvLISTA.Nodes.Add(Nodo, tvwChild, tvLISTA.Nodes(Nodo).FullPath & "\FolderBases", "Bases")
    Objetos.Add Coleccion, xNODO.Key
    For Each dsoBASE In Coleccion
        Set yNODO = tvLISTA.Nodes.Add(xNODO.Key, tvwChild, xNODO.FullPath & "\" & dsoBASE.Name, dsoBASE.Name)
        
        Objetos.Add dsoBASE, yNODO.Key
        
        AddCubes dsoBASE.MDStores, yNODO.Key
        AddDims dsoBASE.Dimensions, yNODO.Key
        AddProps dsoBASE.CustomProperties, yNODO.Key
    Next
End Sub

Private Sub AddCubes(Coleccion, Nodo As String)
' Agrega un folder con la coleccion debajo del nodo
Dim xNODO As Node
Dim yNODO As Node

    Set xNODO = tvLISTA.Nodes.Add(Nodo, tvwChild, tvLISTA.Nodes(Nodo).FullPath & "\FolderCubes", "Cubos")
    Objetos.Add Coleccion, xNODO.Key
    For Each dsoCUBE In Coleccion
        Set yNODO = tvLISTA.Nodes.Add(xNODO.Key, tvwChild, xNODO.FullPath & "\" & dsoCUBE.Name, dsoCUBE.Name)
        
        Objetos.Add dsoCUBE, yNODO.Key
        
        AddDims dsoCUBE.Dimensions, yNODO.Key
        AddProps dsoCUBE.CustomProperties, yNODO.Key
    Next
End Sub

Private Sub AddDims(Coleccion, Nodo As String)
' Agrega un folder con las dimensiones debajo del nodo
Dim xNODO As Node
Dim yNODO As Node

    Set xNODO = tvLISTA.Nodes.Add(Nodo, tvwChild, tvLISTA.Nodes(Nodo).FullPath & "\FolderDims", "Dimensiones")
    Objetos.Add Coleccion, xNODO.Key
    For Each dsoDIM In Coleccion
        Set yNODO = tvLISTA.Nodes.Add(xNODO.Key, tvwChild, xNODO.FullPath & "\" & dsoDIM.Name, dsoDIM.Name)
        
        Objetos.Add dsoDIM, yNODO.Key
        
        AddLevels dsoDIM.Levels, yNODO.Key
        AddProps dsoDIM.CustomProperties, yNODO.Key
    Next
End Sub


Private Sub AddLevels(Coleccion, Nodo As String)
' Agrega un folder con los niveles debajo del nodo
Dim xNODO As Node
Dim yNODO As Node

    Set xNODO = tvLISTA.Nodes.Add(Nodo, tvwChild, tvLISTA.Nodes(Nodo).FullPath & "\FolderLevels", "Niveles")
    Objetos.Add Coleccion, xNODO.Key
    For Each dsoLEVEL In Coleccion
        Set yNODO = tvLISTA.Nodes.Add(xNODO.Key, tvwChild, xNODO.FullPath & "\" & dsoLEVEL.Name, dsoLEVEL.Name)
        
        Objetos.Add dsoLEVEL, yNODO.Key
        
        AddProps dsoLEVEL.CustomProperties, yNODO.Key
    Next
End Sub

Private Sub AddProps(Coleccion, Nodo As String)
' Agrega un folder con las propiedades debajo del nodo
Dim xNODO As Node
Dim yNODO As Node

    Set xNODO = tvLISTA.Nodes.Add(Nodo, tvwChild, tvLISTA.Nodes(Nodo).FullPath & "\FolderProperties", "Propiedades Personalizadas")
    Objetos.Add Coleccion, xNODO.Key
    For Each dsoPROP In Coleccion
        Set yNODO = tvLISTA.Nodes.Add(xNODO.Key, tvwChild, xNODO.FullPath & "\" & dsoPROP.Name, dsoPROP.Name)
        
        Objetos.Add dsoPROP, yNODO.Key
    Next
End Sub

Private Sub cmdSHOW_Click()
' Abre la ventana de browse del objeto
Dim xNODO As Node
Dim xOBJETO
Dim strKEY As String

    'If Button = vbKeyRButton Then
        If Not (tvLISTA.SelectedItem Is Nothing) Then
            Set xNODO = tvLISTA.SelectedItem
            
            Set xOBJETO = Nothing
            strKEY = xNODO.Key
            On Error Resume Next
            Set xOBJETO = Objetos(strKEY)
            On Error GoTo 0
            
            If Not (xOBJETO Is Nothing) Then
                intNUM = intNUM + 1
                
                ReDim Preserve clsFORM(1 To intNUM)
                Set clsFORM(intNUM) = New ExploradorActiveX.ExploradorActiveXForm
                Set clsFORM(intNUM).Objeto = xOBJETO
                clsFORM(intNUM).Muestra
            End If
        End If
    'End If
End Sub

Private Sub Form_Load()
' Inicio

    Set Bases = New Collection
    Set Cubos = New Collection
    Set Dimensiones = New Collection
    Set Niveles = New Collection
    Set Propiedades = New Collection
    Set Objetos = New Collection
    
    flagCONN = True
    
    intNUM = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Termina
Dim a As Integer

    If Not flagCONN Then
        dsoSERVER.CloseServer
        Set dsoSERVER = Nothing
        
        Set Bases = Nothing
        Set Cubos = Nothing
        Set Dimensiones = Nothing
        Set Niveles = Nothing
        Set Propiedades = Nothing
        Set Objetos = Nothing
        
        If intNUM > 0 Then
            On Error Resume Next
            For a = 1 To intNUM Step 1
                clsFORM(a).Descarga
                Set clsFORM(a) = Nothing
            Next a
            On Error GoTo 0
        End If
        
        flagCONN = True
        cmdCONNECT.Caption = "Conectar"
        
        End
    End If
End Sub
