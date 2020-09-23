VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExploradorActiveXForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private fPRINCIPAL As Form

Private mvarObjeto

Public Function Muestra(Optional ByVal Modalidad As FormShowConstants = vbModeless) As Boolean
Attribute Muestra.VB_UserMemId = 0
' Muestra el formulario de exploracion del control
Dim flagRETURN As Boolean

    flagRETURN = False
    If IsObject(mvarObjeto) Then
        If Not (mvarObjeto Is Nothing) Then
            Set fPRINCIPAL = New frmPRINCIPAL
            Load fPRINCIPAL
            fPRINCIPAL.CargaObjeto mvarObjeto
            fPRINCIPAL.Show Modalidad
            
            flagRETURN = True
        End If
    End If
    
    Muestra = flagRETURN
End Function

Public Function Descarga() As Boolean
' Descarga el formulario
Dim flagRETURN As Boolean

    flagRETURN = False
    If Not (fPRINCIPAL Is Nothing) Then
        Unload fPRINCIPAL
        
        flagRETURN = True
    End If
    
    Descarga = flagRETURN
End Function

Public Property Let Objeto(ByVal vData)
    mvarObjeto = vData
End Property

Public Property Set Objeto(ByVal vData)
    Set mvarObjeto = vData
End Property

Public Property Get Objeto()
    If IsObject(mvarObjeto) Then
        Set Objeto = mvarObjeto
    Else
        Objeto = mvarObjeto
    End If
End Property

Private Sub Class_Initialize()
' Rutinas de inicio de la clase

    Set mvarObjeto = Nothing
End Sub
