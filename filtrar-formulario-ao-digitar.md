# Filtrar formulário no Access ao digitar em um controle de texto

1 - Criar um campo na consulta com o nome 'fieldFilter' incluindo todos os outros campos separados por um espaço " "
2 - Inserir no formulário uma Caixa de texto com o nome 'txtFilter'
3 - Na propriedade 'Ao liberar tecla' dessa Caixa de texto incluir a segunte função:

```csharp
On Error GoTo errHandler
Dim txtFilterOrigin, likeText As String, word As Variant
txtFilterOrigin = Me.txtFilter.Text
txtFilter = (txtFilterOrigin)
likeText = ""
If Len(txtFilter) > 0 Then
    For Each word In Split(txtFilter)
        If Len(word) > 0 Then
            likeText = likeText & "fieldFilter like '*" & word & "*' and "
        End If
    Next
   Me.Form.Filter = Left(likeText, Len(likeText) - 5)
   Me.FilterOn = True
   Me.txtFilter.SelStart = Len(Me.txtFilter.Text)
Else
   Me.Filter = ""
   Me.FilterOn = False
   Me.txtFilter.SetFocus
End If
Exit Sub
errHandler:
MsgBox "Nenhum registro encontrado com a referência informada", vbInformation, "Filtro"
Me.Filter = ""
Me.FilterOn = False
Me.txtFilter.SetFocus
Me.txtFilter.Text = Left(txtFilterOrigin, Len(txtFilterOrigin) - 1)
```

Agora o formulário deve ser filtrado à medida que o usuário digita alguma coisa
