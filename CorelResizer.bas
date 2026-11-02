Attribute VB_Name = "modCorelResizer"
Option Explicit

Public Enum TipoForma
    tfCirculo = 1
    tfElipse
    tfCuadrado
    tfRectangulo
    tfRombo 'Incluye las cometas: ambos se miden por sus diagonales
End Enum

Public Enum AccionFinal
    afNada = 0
    afAgrupar
    afCombinar
End Enum

Private Const TOLERANCIA As Double = 0.01 'Margen relativo en las comparaciones geométricas (1 %)
Private Const PI As Double = 3.14159265358979

Sub CorelResizer()

    'Macro que redimensiona círculos, elipses, cuadrados, rectángulos, rombos y cometas
    'a las medidas indicadas, manteniendo su posición
    '@autor: Fernando Maciá
    '@version: 1.2.0 02/11/2026
    '@repositorio: https://github.com/fernandomacia/CorelResizer
    '@licencia: GPL-3.0 - https://www.gnu.org/licenses/gpl-3.0.html
    'Copyright (C) 2023 Fernando Maciá

    Dim sr As ShapeRange, nr As ShapeRange, s As Shape
    Dim tipo As TipoForma, largo As Double, corto As Double, accion As AccionFinal
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

    frmCorelResizer.Preparar sr 'El formulario pide las medidas y comprueba las formas antes de cerrarse
    frmCorelResizer.Show
    If Not frmCorelResizer.Aceptado Then
        Unload frmCorelResizer
        Exit Sub
    End If
    tipo = frmCorelResizer.Tipo
    largo = frmCorelResizer.Medida1
    corto = frmCorelResizer.Medida2
    accion = frmCorelResizer.AlTerminar
    Unload frmCorelResizer

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
        If s.Type <> cdrCurveShape Then s.ConvertToCurves 'Rectángulos y elipses nativos
        nr.AddRange s.BreakApartEx 'Cada forma de una curva combinada por separado
    Next s

    For Each s In nr
        TransformarForma s, tipo, largo, corto
    Next s

    If nr.Count > 1 And accion = afAgrupar Then
        nr.Group.CreateSelection
    ElseIf nr.Count > 1 And accion = afCombinar Then
        nr.Combine.CreateSelection
    Else
        nr.CreateSelection
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

'Convierte el texto en una medida mayor que 0, aceptando punto o coma como separador decimal
Public Function LeerMedida(ByVal texto As String, ByRef valor As Double) As Boolean
    Dim sep As String
    sep = Mid$(CStr(0.5), 2, 1) 'Separador decimal de la configuración regional
    texto = Replace(Replace(Trim$(texto), ".", sep), ",", sep)
    If texto = "" Or Not IsNumeric(texto) Then Exit Function
    valor = CDbl(texto)
    LeerMedida = valor > 0
End Function

'Cuenta las formas de la selección: cada subtrayecto de una curva (combinada o no) es una forma
Public Function ContarFormas(sr As ShapeRange) As Long
    Dim s As Shape, n As Long
    For Each s In sr
        Select Case s.Type
            Case cdrGroupShape
                n = n + ContarFormas(s.Shapes.All)
            Case cdrCurveShape, cdrRectangleShape, cdrEllipseShape, cdrPolygonShape
                n = n + s.DisplayCurve.SubPaths.Count
        End Select
    Next s
    ContarFormas = n
End Function

'Comprueba sin modificar nada que todas las formas pueden ser del tipo elegido.
'Devuelve "" si todo es correcto, o el motivo por el que no se puede redimensionar
Public Function ComprobarFormas(sr As ShapeRange, tipo As TipoForma) As String
    Dim noFormas As Long, noEncajan As Long
    RevisarFormas sr, tipo, noFormas, noEncajan
    If noFormas > 0 Then
        ComprobarFormas = "La selección contiene objetos que no son formas (textos, imágenes...)."
    ElseIf noEncajan > 0 Then
        ComprobarFormas = noEncajan & IIf(noEncajan = 1, " forma no encaja", " formas no encajan") & _
            " con el tipo elegido. Revisa que no haya tipos mezclados ni esquinas redondeadas."
    End If
End Function

Private Sub RevisarFormas(sr As ShapeRange, tipo As TipoForma, ByRef noFormas As Long, ByRef noEncajan As Long)
    Dim s As Shape, sp As SubPath
    For Each s In sr
        Select Case s.Type
            Case cdrGroupShape
                RevisarFormas s.Shapes.All, tipo, noFormas, noEncajan
            Case cdrCurveShape, cdrRectangleShape, cdrEllipseShape, cdrPolygonShape
                For Each sp In s.DisplayCurve.SubPaths 'DisplayCurve: la forma como curva, sin convertirla
                    If Not EncajaConTipo(sp, tipo) Then noEncajan = noEncajan + 1
                Next sp
            Case Else
                noFormas = noFormas + 1
        End Select
    Next s
End Sub

'Comprueba si un subtrayecto puede ser del tipo elegido aunque esté estirado, girado o inclinado,
'usando solo propiedades que no cambian al deformar
Private Function EncajaConTipo(sp As SubPath, tipo As TipoForma) As Boolean
    Dim x(1 To 4) As Double, y(1 To 4) As Double, i As Long, curvos As Long, margen As Double

    If Not sp.Closed Or sp.Nodes.Count <> 4 Then Exit Function
    For i = 1 To 4
        x(i) = sp.Nodes(i).PositionX
        y(i) = sp.Nodes(i).PositionY
    Next i
    For i = 1 To sp.Segments.Count
        If sp.Segments(i).Type = cdrCurveSegment Then curvos = curvos + 1
    Next i
    margen = TOLERANCIA * (Distancia(x(1), y(1), x(3), y(3)) + Distancia(x(2), y(2), x(4), y(4)))

    Select Case tipo
        Case tfCirculo, tfElipse 'Una elipse deformada sigue siendo una elipse de 4 tramos curvos
            EncajaConTipo = (curvos = 4)
        Case tfCuadrado, tfRectangulo 'Siguen siendo paralelogramos: las diagonales se cortan en su punto medio
            EncajaConTipo = (curvos = 0) And _
                Distancia((x(1) + x(3)) / 2, (y(1) + y(3)) / 2, (x(2) + x(4)) / 2, (y(2) + y(4)) / 2) <= margen
        Case tfRombo 'Rombo o cometa: una diagonal pasa por el punto medio de la otra
            EncajaConTipo = (curvos = 0) And _
                (DistanciaARecta((x(2) + x(4)) / 2, (y(2) + y(4)) / 2, x(1), y(1), x(3), y(3)) <= margen Or _
                 DistanciaARecta((x(1) + x(3)) / 2, (y(1) + y(3)) / 2, x(2), y(2), x(4), y(4)) <= margen)
    End Select
End Function

Private Function Distancia(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    Distancia = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

'Distancia del punto P a la recta que pasa por A y B
Private Function DistanciaARecta(ByVal px As Double, ByVal py As Double, ByVal ax As Double, ByVal ay As Double, _
                                 ByVal bx As Double, ByVal by As Double) As Double
    DistanciaARecta = Abs((bx - ax) * (py - ay) - (by - ay) * (px - ax)) / Distancia(ax, ay, bx, by)
End Function

'Ajusta una forma (curva de un solo subtrayecto, ya comprobada) a las medidas indicadas,
'manteniendo su centro y la orientación que mejor encaja con su forma actual
Public Sub TransformarForma(s As Shape, tipo As TipoForma, largo As Double, corto As Double)
    Select Case tipo
        Case tfCirculo: s.SetSize largo, largo
        Case tfElipse: AjustarElipse s, largo, corto
        Case tfCuadrado, tfRectangulo: AjustarParalelogramo s.Curve.SubPaths(1), largo, corto
        Case tfRombo: AjustarPorDiagonales s.Curve.SubPaths(1), largo, corto
    End Select
End Sub

Private Sub AjustarElipse(s As Shape, largo As Double, corto As Double)
    Dim sp As SubPath, cx As Double, cy As Double, ux As Double, uy As Double, vx As Double, vy As Double
    Dim a As Double, b As Double, c As Double, angulo As Double

    Set sp = s.Curve.SubPaths(1)
    cx = (sp.Nodes(1).PositionX + sp.Nodes(3).PositionX) / 2
    cy = (sp.Nodes(1).PositionY + sp.Nodes(3).PositionY) / 2
    ux = (sp.Nodes(1).PositionX - sp.Nodes(3).PositionX) / 2 'Semidiámetros conjugados: los nodos opuestos
    uy = (sp.Nodes(1).PositionY - sp.Nodes(3).PositionY) / 2
    vx = (sp.Nodes(2).PositionX - sp.Nodes(4).PositionX) / 2
    vy = (sp.Nodes(2).PositionY - sp.Nodes(4).PositionY) / 2

    'Dirección del eje largo: autovector principal de M·M^T, con M = [u v]
    a = ux * ux + vx * vx
    b = ux * uy + vx * vy
    c = uy * uy + vy * vy
    angulo = Atan2(2 * b, a - c) / 2 * 180 / PI

    s.RotateEx -angulo, cx, cy 'Eje largo en horizontal
    s.SetSize largo, corto
    s.RotateEx angulo, cx, cy
End Sub

Private Sub AjustarParalelogramo(sp As SubPath, largo As Double, corto As Double)
    Dim x(1 To 4) As Double, y(1 To 4) As Double, cx As Double, cy As Double
    Dim e1x As Double, e1y As Double, e2x As Double, e2y As Double
    Dim f1x As Double, f1y As Double, f2x As Double, f2y As Double, l1 As Double, l2 As Double

    LeerNodos sp, x, y
    cx = (x(1) + x(2) + x(3) + x(4)) / 4
    cy = (y(1) + y(2) + y(3) + y(4)) / 4
    e1x = (x(2) - x(1) + x(3) - x(4)) / 2 'Lados: media de cada pareja de lados opuestos
    e1y = (y(2) - y(1) + y(3) - y(4)) / 2
    e2x = (x(4) - x(1) + x(3) - x(2)) / 2
    e2y = (y(4) - y(1) + y(3) - y(2)) / 2
    Orientacion e1x, e1y, e2x, e2y, f1x, f1y, f2x, f2y

    If Distancia(0, 0, e1x, e1y) >= Distancia(0, 0, e2x, e2y) Then l1 = largo: l2 = corto Else l1 = corto: l2 = largo

    x(1) = cx - l1 / 2 * f1x - l2 / 2 * f2x
    y(1) = cy - l1 / 2 * f1y - l2 / 2 * f2y
    MoverNodos sp, x(1), y(1), x(1) + l1 * f1x, y(1) + l1 * f1y, _
        x(1) + l1 * f1x + l2 * f2x, y(1) + l1 * f1y + l2 * f2y, x(1) + l2 * f2x, y(1) + l2 * f2y
End Sub

Private Sub AjustarPorDiagonales(sp As SubPath, largo As Double, corto As Double)
    Dim x(1 To 4) As Double, y(1 To 4) As Double, d1x As Double, d1y As Double, d2x As Double, d2y As Double
    Dim cruz As Double, t As Double, u As Double, qx As Double, qy As Double
    Dim f1x As Double, f1y As Double, f2x As Double, f2y As Double, l1 As Double, l2 As Double

    LeerNodos sp, x, y
    d1x = x(3) - x(1): d1y = y(3) - y(1)
    d2x = x(4) - x(2): d2y = y(4) - y(2)

    'Punto Q donde se cortan las diagonales: Q = P1 + t·d1 = P2 + u·d2.
    't y u son las proporciones en que Q divide cada diagonal; se conservan para que la cometa siga siéndolo
    cruz = d1x * d2y - d1y * d2x
    t = ((x(2) - x(1)) * d2y - (y(2) - y(1)) * d2x) / cruz
    u = ((x(2) - x(1)) * d1y - (y(2) - y(1)) * d1x) / cruz
    qx = x(1) + t * d1x
    qy = y(1) + t * d1y
    Orientacion d1x, d1y, d2x, d2y, f1x, f1y, f2x, f2y

    If Distancia(0, 0, d1x, d1y) >= Distancia(0, 0, d2x, d2y) Then l1 = largo: l2 = corto Else l1 = corto: l2 = largo

    MoverNodos sp, qx - t * l1 * f1x, qy - t * l1 * f1y, qx - u * l2 * f2x, qy - u * l2 * f2y, _
        qx + (1 - t) * l1 * f1x, qy + (1 - t) * l1 * f1y, qx + (1 - u) * l2 * f2x, qy + (1 - u) * l2 * f2y
End Sub

'Ejes perpendiculares (f1, f2) que mejor encajan con las direcciones d1 y d2: el giro más cercano
'a la matriz [d1 d2] (descomposición polar). Conserva el sentido de los nodos para no voltear la forma
Private Sub Orientacion(ByVal d1x As Double, ByVal d1y As Double, ByVal d2x As Double, ByVal d2y As Double, _
                        ByRef f1x As Double, ByRef f1y As Double, ByRef f2x As Double, ByRef f2y As Double)
    Dim sentido As Double, angulo As Double
    sentido = IIf(d1x * d2y - d1y * d2x < 0, -1, 1)
    angulo = Atan2(d1y - sentido * d2x, d1x + sentido * d2y)
    f1x = Cos(angulo): f1y = Sin(angulo)
    f2x = -sentido * f1y: f2y = sentido * f1x
End Sub

Private Sub LeerNodos(sp As SubPath, x() As Double, y() As Double)
    Dim i As Long
    For i = 1 To 4
        x(i) = sp.Nodes(i).PositionX
        y(i) = sp.Nodes(i).PositionY
    Next i
End Sub

Private Sub MoverNodos(sp As SubPath, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                       ByVal x3 As Double, ByVal y3 As Double, ByVal x4 As Double, ByVal y4 As Double)
    sp.Nodes(1).PositionX = x1: sp.Nodes(1).PositionY = y1
    sp.Nodes(2).PositionX = x2: sp.Nodes(2).PositionY = y2
    sp.Nodes(3).PositionX = x3: sp.Nodes(3).PositionY = y3
    sp.Nodes(4).PositionX = x4: sp.Nodes(4).PositionY = y4
End Sub

'Ángulo en radianes del punto (x, y), en todo el rango de -PI a PI (VBA no trae Atan2)
Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        Atan2 = Atn(y / x) + IIf(y >= 0, PI, -PI)
    Else
        Atan2 = IIf(y >= 0, PI / 2, -PI / 2)
    End If
End Function
