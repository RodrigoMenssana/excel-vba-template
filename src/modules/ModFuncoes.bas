Attribute VB_Name = "ModFuncoes"
'A função AcSQL é uma rotina de sanitização e normalização de
'string voltada para uso em consultas SQL (principalmente Access com LIKE).
'Ela percorre o texto caractere a caractere e substitui determinados
'caracteres por classes de equivalência, permitindo busca acento-insensível.

Public Function AcSQL(Valor As String) As String

    Dim N, t, v
    t = ""

    For N = 1 To VBA.Len(Valor)

        v = VBA.Asc(VBA.Mid(Valor, N, 1))

        Select Case v
            Case 39: t = t & "''"
            Case 65: t = t & "[ÁÀÂÃÄÅ]"
            Case 67: t = t & "[ÇC]"
            Case 69: t = t & "[ÉÈÊËE]"
            Case 73: t = t & "[ÍÌÎÏI]"
            Case 79: t = t & "[ÓÒÔÕÖO]"
            Case 85: t = t & "[ÚÙÛÜU]"
            Case 97: t = t & "[áàâãäåa]"
            Case 99: t = t & "[çc]"
            Case 101: t = t & "[éèêëe]"
            Case 105: t = t & "[íìîïi]"
            Case 111: t = t & "[óòôõöo]"
            Case 117: t = t & "[úùûüu]"

            Case Else
                If v > 31 And v < 127 Then
                    t = t & VBA.Chr(v)
                Else
                    t = t & "_"
                End If
        End Select

    Next

    AcSQL = t

End Function
