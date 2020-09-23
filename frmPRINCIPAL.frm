VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPRINCIPAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador ActiveX"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdSAVE 
      Left            =   5880
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Guardar informacion de ActiveX"
      Filter          =   "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
   End
   Begin VB.TextBox txtVALOR 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5520
      Width           =   6975
   End
   Begin MSComctlLib.TreeView tvLISTA 
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9340
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilIMAGEN 
      Left            =   5280
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPRINCIPAL.frx":0000
            Key             =   "Clase"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPRINCIPAL.frx":0A12
            Key             =   "Evento"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPRINCIPAL.frx":1424
            Key             =   "Metodo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPRINCIPAL.frx":1E36
            Key             =   "Funcion"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPRINCIPAL.frx":2848
            Key             =   "Propiedad"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPRINCIPAL.frx":325A
            Key             =   "Constante"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPRINCIPAL.frx":3C6C
            Key             =   "Indefinido"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstMIEMBROS 
      Height          =   450
      Left            =   5880
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstCONTROL 
      Height          =   450
      Left            =   5280
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdGUARDAR 
      Caption         =   "&Guardar"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdACTUALIZAR 
      Caption         =   "&Actualizar"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmPRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Referencias API
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

' Objetos solo del formulario
Private objCONTROL
Private m_TLInf As TLI.TypeLibInfo

Public Sub CargaObjeto(ByVal Objeto)
' Carga el objeto especificado

    Set objCONTROL = Objeto
    
    Actualizar
End Sub

Private Sub Actualizar()
' Actualiza el formulario
Dim enumKINDS As TLI.InvokeKinds
Dim enumTYPE As TLI.TypeKinds
Dim oMI As TLI.MemberInfo
Dim oP As TLI.ParameterInfo
Dim flagErr As Boolean
Dim IsFunc As Boolean
Dim strICON As String
Dim strAUX
Dim Valor
Dim sCH
Dim a

    tvLISTA.Nodes.Clear
    txtVALOR.Text = vbNullString
    
    If objCONTROL Is Nothing Then
        tvLISTA.Nodes.Add , , "Root|1", "Indefinido", "Propiedad", "Propiedad"
        Exit Sub
    End If
    
    Referencia

    With lstMIEMBROS
        For a = 0 To (.ListCount - 1) Step 1
            enumKINDS = .ItemData(a)
            
            Err.Clear
            On Error Resume Next
            Set oMI = m_TLInf.GetMemberInfo(lstCONTROL.ItemData(lstCONTROL.ListIndex), enumKINDS, , .List(a))
            flagErr = IIf(Err, True, False)
            On Error GoTo 0
            
            If flagErr Then
                tvLISTA.Nodes.Add "Root", tvwChild, "Evento|" & (tvLISTA.Nodes.Count + 1), .List(a), "Evento", "Evento"
            Else
                IsFunc = False
                
                strAUX = oMI.Name
                
                If oMI.Parameters.Count > 0 Then
                    strAUX = strAUX & "("
                    sCH = ""
                    For Each oP In oMI.Parameters
                        strAUX = strAUX & sCH
                        If oP.Default Or oP.Optional Then
                            strAUX = strAUX & "["
                        End If
                        strAUX = strAUX & oP.Name
                        If (oP.VarTypeInfo.VarType = VT_ARRAY) Or (oP.VarTypeInfo.VarType = VT_VECTOR) Then
                            strAUX = strAUX & "() As "
                        Else
                            strAUX = strAUX & " As "
                        End If
                        If oP.VarTypeInfo.VarType <> vbVariant Then
                            strAUX = strAUX & TypeName(oP.VarTypeInfo.TypedVariant)
                        Else
                            strAUX = strAUX & "Variant"
                        End If
                        If oP.Optional Then
                            If oP.Default Then
                                strAUX = strAUX & " = " & ProduceDefaultValue(oP.DefaultValue, oP.VarTypeInfo.TypeInfo)
                            End If
                        End If
                        If oP.Default Or oP.Optional Then
                            strAUX = strAUX & "]"
                        End If
                        sCH = ", "
                    Next
                    strAUX = strAUX & ")"
                Else
                    If .ItemData(a) = TLI.InvokeKinds.INVOKE_EVENTFUNC Or .ItemData(a) = TLI.InvokeKinds.INVOKE_FUNC Then
                        strAUX = strAUX & "()"
                    End If
                End If
                
                If Not (oMI.ReturnType.TypeInfo Is Nothing) Then
                    strAUX = strAUX & " As "
                    If oMI.ReturnType.IsExternalType Then
                        strICON = oMI.ReturnType.TypeLibInfoExternal.Name
                    Else
                        strICON = oMI.ReturnType.TypeInfo.Name
                    End If
                    While Left(strICON, 1) = "_"
                        strICON = Right(strICON, Len(strICON) - 1)
                    Wend
                    strAUX = strAUX & strICON
                    
                    IsFunc = True
                End If
                
                If .ItemData(a) <> TLI.InvokeKinds.INVOKE_EVENTFUNC And .ItemData(a) <> TLI.InvokeKinds.INVOKE_FUNC Then
                    strAUX = strAUX & " <Type = " & TypeName(oMI.ReturnType.TypedVariant) & ">"
                End If
                
                Err.Clear
                On Error Resume Next
                Valor = TLI.InvokeHook(objCONTROL, .List(a), INVOKE_PROPERTYGET)
                flagErr = IIf(Err, True, False)
                On Error GoTo 0
                
                If Not flagErr Then
                    If CStr(Valor) <> "" Then
                        strAUX = strAUX & " <Value = " & CStr(Valor) & ">"
                    End If
                End If
                
                strICON = "Indefinido"
                Select Case .ItemData(a)
                    Case TLI.InvokeKinds.INVOKE_CONST
                        strICON = "Constante"
                    Case TLI.InvokeKinds.INVOKE_EVENTFUNC
                        strICON = "Evento"
                    Case TLI.InvokeKinds.INVOKE_FUNC
                        If IsFunc Then
                            strICON = "Funcion"
                        Else
                            strICON = "Metodo"
                        End If
                    Case TLI.InvokeKinds.INVOKE_PROPERTYGET, TLI.InvokeKinds.INVOKE_PROPERTYPUT, TLI.InvokeKinds.INVOKE_PROPERTYPUTREF, TLI.InvokeKinds.INVOKE_PROPERTYGET Or TLI.InvokeKinds.INVOKE_PROPERTYPUT, TLI.InvokeKinds.INVOKE_PROPERTYGET Or TLI.InvokeKinds.INVOKE_PROPERTYPUTREF, TLI.InvokeKinds.INVOKE_PROPERTYPUT Or TLI.InvokeKinds.INVOKE_PROPERTYPUTREF, TLI.InvokeKinds.INVOKE_PROPERTYGET Or TLI.InvokeKinds.INVOKE_PROPERTYPUT Or TLI.InvokeKinds.INVOKE_PROPERTYPUTREF
                        strICON = "Propiedad"
                    Case Else
                        strICON = "Indefinido"
                End Select
                
                tvLISTA.Nodes.Add "Root|1", tvwChild, strICON & "|" & (tvLISTA.Nodes.Count + 1), strAUX, strICON, strICON
            End If
        Next
    End With
End Sub

Private Sub Referencia()
Dim oII As TLI.InterfaceInfo
Dim oTLI As TLI.TypeLibInfo
Dim objPROP
Dim strAUX
Dim Valor
Dim a

    Set oII = TLI.InterfaceInfoFromObject(objCONTROL)
    Set oTLI = oII.Parent
    
    If Not (oTLI Is Nothing) Then
        Set oTLI = TLI.TypeLibInfoFromRegistry(oTLI.Guid, oTLI.MajorVersion, oTLI.MinorVersion, oTLI.LCID)
        
        oTLI.GetTypesWithSubStringDirect TypeName(objCONTROL), lstCONTROL.hWnd, tliWtListBox
        
        For a = 0 To (lstCONTROL.ListCount - 1) Step 1
            If lstCONTROL.List(a) = TypeName(objCONTROL) Then
                lstCONTROL.ListIndex = a
                oTLI.SetMemberFilters FUNCFLAG_NONE, VARFLAG_NONE
                oTLI.GetMembersDirect lstCONTROL.ItemData(a), lstMIEMBROS.hWnd, tliWtListBox, tliIdtInvokeKinds, False
            End If
        Next
        
        tvLISTA.Nodes.Add , , "Root|1", oTLI.Name & "." & TypeName(objCONTROL), "Clase", "Clase"
        tvLISTA.Nodes(1).Expanded = True
        
        Set m_TLInf = oTLI
    End If
End Sub

Private Function GUIDToString(ByRef uGUID As String) As String
' Convierte un guid en un guis string
Dim lngSize As Long
Dim strBuffer As String
Dim strReturn As String
Dim lngRetVal As Long
    
    lngSize = 50
    
    strBuffer = String(lngSize, Chr(0))
    lngRetVal = StringFromGUID2(uGUID, StrPtr(strBuffer), lngSize)
    strReturn = vbNullString
    If lngRetVal > 0 Then
        strReturn = Mid(strBuffer, 1, lngRetVal - 1)
    End If
    
    GUIDToString = strReturn
End Function

Private Sub Guardar()
' Guardar la informacion recopilada
Dim strAUX() As String
Dim strFILE As String
Dim intFILE As String
Dim flagErr As Boolean
Dim xNODO As Node

    If objCONTROL Is Nothing Then
        Exit Sub
    End If

    cdSAVE.DefaultExt = "txt"
    cdSAVE.FileName = "Informacion (" & tvLISTA.Nodes(1).Text & ").txt"
    cdSAVE.Flags = cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    
    Err.Clear
    On Error Resume Next
    cdSAVE.ShowSave
    flagErr = IIf(Err, True, False)
    On Error GoTo 0
    
    If Not flagErr Then
        intFILE = FreeFile()
        strFILE = cdSAVE.FileName
        Open strFILE For Output As #intFILE
        For Each xNODO In tvLISTA.Nodes
            strAUX() = Split(xNODO.Key, "|")
            Select Case strAUX(LBound(strAUX))
                Case "Root"
                    Print #intFILE, "- " & Trim(xNODO.Text)
                Case Else
                    Print #intFILE, "{" & strAUX(LBound(strAUX)) & "} " & xNODO.Text
            End Select
        Next
        Close #intFILE
        
        MsgBox "Archivo guardado con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Guardar informacion"
    End If
End Sub

Private Sub cmdACTUALIZAR_Click()
' Actualiza el formulario

    Actualizar
End Sub

Private Sub cmdGUARDAR_Click()
' Guarda la informacion recopilada

    Guardar
End Sub

Private Sub Form_Load()
' Rutinas de inicio del formulario

    Set tvLISTA.ImageList = ilIMAGEN
End Sub

Private Sub tvLISTA_NodeClick(ByVal Node As MSComctlLib.Node)
' Click en nodo

    If Not (Node Is Nothing) Then
        txtVALOR.Text = Replace(tvLISTA.SelectedItem.Text, " <", vbCrLf & "<") & vbCrLf
    Else
        txtVALOR.Text = vbNullString
    End If
End Sub
