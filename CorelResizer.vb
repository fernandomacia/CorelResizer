Sub SS06()

    'Macro que redimensiona círculos a una medida dada por teclado
    '@autor: Fernando Maciá
    '@version: 1.03 27/10/21
    
    Dim sr As ShapeRange, nr As ShapeRange, s As Shape, size, Pregunta
    
    Set sr = ActiveSelectionRange
    Set s = ActiveShape
    Set nr = New ShapeRange

    If Documents.Count = 0 Then 'En caso de no haber documento abrieto termina el proceso
        MsgBox "No hay documentos abiertos.", vbOKOnly + vbCritical
        Exit Sub
    Else
        If sr.Count < 1 Then
            MsgBox "No hay objetos seleccionados.", vbOKOnly + vbCritical
            Exit Sub
        End If
    End If
     
    ActiveDocument.Unit = cdrMillimeter 'Indicamos que la unidad de medida son milimetros
    ActiveDocument.ReferencePoint = cdrCenter 'Indicamos que el punto de referencia es el centro de los objetos
    
    If sr.Count >= 1 Then 'Si hay objetos seleccionados
            
        size = InputBox("Ingrese el tamaño")
        
        If size = "" Then
            Exit Sub
        Else

         If IsNumeric(size) Then
             sr.UngroupAll 'Desagrupamos los objetos
             For Each s In sr
                 nr.AddRange s.BreakApartEx
                 Next s
                 nr.CreateSelection
         
             For Each s In nr
                 s.SetSize size, size
                 Next s
         
             If nr.Count > 1 Then 'Si los objetos seleccionados son mas de uno
                 Pregunta = MsgBox("Quieres combinar los objetos?", vbYesNo + vbQuestion)
        
                 If Pregunta = vbYes Then
                     Set s = nr.Combine
                     s.CreateSelection
                 End If
             End If
         Else
             MsgBox "El dato introducido no tiene un formato correcto.", vbOKOnly + vbCritical
             Exit Sub
         End If
        End If
    Else
        MsgBox "No hay objetos seleccionados.", vbOKOnly + vbExclamation
    End If
End Sub