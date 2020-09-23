Attribute VB_Name = "Funciones"
Option Explicit

' Set de funciones obtenidas de Microsoft

Public Function ProduceDefaultValue(DefVal As Variant, ByVal TI As TypeInfo) As String
Dim lTrackVal As Long
Dim MI As MemberInfo
Dim TKind As TypeKinds
    If TI Is Nothing Then
        Select Case VarType(DefVal)
            Case vbString
                If Len(DefVal) Then
                    ProduceDefaultValue = """" & DefVal & """"
                End If
            Case vbBoolean 'Always show for Boolean
                ProduceDefaultValue = DefVal
            Case vbDate
                If DefVal Then
                    ProduceDefaultValue = "#" & DefVal & "#"
                End If
            Case Else 'Numeric Values
                If DefVal <> 0 Then
                    ProduceDefaultValue = DefVal
                End If
        End Select
    Else
        'See if we have an enum and track the matching member
        'If the type is an object, then there will never be a
        'default value other than Nothing
        TKind = TI.TypeKind
        Do While TKind = TKIND_ALIAS
            TKind = TKIND_MAX
            On Error Resume Next
            Set TI = TI.ResolvedType
            If Err = 0 Then TKind = TI.TypeKind
            On Error GoTo 0
        Loop
        If TI.TypeKind = TKIND_ENUM Then
            lTrackVal = DefVal
            For Each MI In TI.Members
                If MI.Value = lTrackVal Then
                    ProduceDefaultValue = MI.Name
                    Exit For
                End If
            Next
        End If
    End If
End Function
