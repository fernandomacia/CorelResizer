Attribute VB_Name = "modCorelResizer"
Option Explicit

Sub CorelResizer()

    'Macro que redimensiona círculos a una medida dada por teclado
    '@autor: Fernando Maciá
    '@version: 1.1.0 29/10/2026
    '@repositorio: https://github.com/fernandomacia/CorelResizer
    '@licencia: GPL-3.0 - https://www.gnu.org/licenses/gpl-3.0.html
    'Copyright (C) 2023 Fernando Maciá

    Dim sr As ShapeRange, nr As ShapeRange, s As Shape, entrada As String, sep As String, size As Double, Pregunta
    Dim unidadOriginal As cdrUnit, refOriginal As cdrReferencePoint

    If Documents.Count = 0 Then 'En caso de no haber documento abierto termina el proceso
        MsgBox "No hay documentos abiertos.", vbOKOnly + vbCritical
        Exit Sub
    End If

    Set sr = ActiveSelectionRange
    Set nr = New ShapeRange

    If sr.Count < 1 Then 'En caso de no haber elementos seleccionados termina el proceso
        MsgBox "No hay objetos seleccionados.", vbOKOnly + vbCritical
        Exit Sub
    End If

    entrada = InputBox("Ingrese el tamaño en mm")

    If entrada = "" Then
        MsgBox "No se ha seleccionado el tamaño.", vbOKOnly + vbCritical
        Exit Sub
    End If

    sep = Mid$(CStr(0.5), 2, 1) 'Separador decimal de la configuración regional
    entrada = Replace(Replace(Trim$(entrada), ".", sep), ",", sep) 'Aceptamos punto o coma como separador decimal

    If Not IsNumeric(entrada) Then
        MsgBox "El dato introducido no tiene un formato correcto.", vbOKOnly + vbCritical
        Exit Sub
    End If

    size = CDbl(entrada)
    If size <= 0 Then
        MsgBox "El tamaño debe ser mayor que 0.", vbOKOnly + vbCritical
        Exit Sub
    End If

    unidadOriginal = ActiveDocument.Unit 'Guardamos la configuración para restaurarla al final
    refOriginal = ActiveDocument.ReferencePoint
    ActiveDocument.Unit = cdrMillimeter 'Indicamos que la unidad de medida son milimetros
    ActiveDocument.ReferencePoint = cdrCenter 'Indicamos que el punto de referencia es el centro de los objetos

    On Error GoTo Fin
    ActiveDocument.BeginCommandGroup "CorelResizer" 'Todo se deshace con un solo Ctrl+Z
    Optimization = True 'Desactivamos el redibujado mientras procesamos
    EventsEnabled = False

    sr.UngroupAll 'Desagrupamos los objetos
    Set sr = ActiveSelectionRange 'Releemos la selección con los objetos ya desagrupados
    For Each s In sr
        nr.AddRange s.BreakApartEx
    Next s
    nr.CreateSelection

    For Each s In nr
        s.SetSize size, size
    Next s

    Optimization = False 'Mostramos el resultado antes de preguntar si combinar
    EventsEnabled = True
    ActiveWindow.Refresh

    If nr.Count > 1 Then 'Si los objetos seleccionados son más de uno
        Pregunta = MsgBox("¿Quieres combinar los objetos?", vbYesNo + vbQuestion)

        If Pregunta = vbYes Then
            Set s = nr.Combine
            s.CreateSelection
        End If
    End If

Fin:
    Optimization = False
    EventsEnabled = True
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    ActiveDocument.Unit = unidadOriginal
    ActiveDocument.ReferencePoint = refOriginal
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbOKOnly + vbCritical

End Sub
