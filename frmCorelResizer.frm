VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorelResizer 
   Caption         =   "CorelResizer"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmCorelResizer.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCorelResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Resultado que lee modCorelResizer después de cerrar el formulario
Public Aceptado As Boolean
Public Tipo As TipoForma
Public Medida1 As Double 'Diámetro, lado o medida larga
Public Medida2 As Double 'Medida corta (solo en los tipos que la necesitan)
Public AlTerminar As AccionFinal

Private WithEvents optCirculo As MSForms.OptionButton
Private WithEvents optElipse As MSForms.OptionButton
Private WithEvents optCuadrado As MSForms.OptionButton
Private WithEvents optRectangulo As MSForms.OptionButton
Private WithEvents optRombo As MSForms.OptionButton
Private WithEvents txtMedida1 As MSForms.TextBox
Private WithEvents txtMedida2 As MSForms.TextBox
Private WithEvents btnAplicar As MSForms.CommandButton
Private WithEvents btnCancelar As MSForms.CommandButton
Private lblObjetos As MSForms.Label
Private lblMedida1 As MSForms.Label
Private lblMedida2 As MSForms.Label
Private lblAlTerminar As MSForms.Label
Private optNada As MSForms.OptionButton
Private optAgrupar As MSForms.OptionButton
Private optCombinar As MSForms.OptionButton
Private lblAviso As MSForms.Label
Private mSeleccion As ShapeRange 'Selección que se comprueba al pulsar Aplicar

Private Const FUENTE As String = "Segoe UI"
Private Const MARGEN As Single = 12
Private Const ANCHO As Single = 270
Private Const APP As String = "CorelResizer" 'Clave del registro donde se recuerda la última configuración

Private Sub UserForm_Initialize()
    Me.Width = ANCHO
    Me.BackColor = vbWhite

    Set lblObjetos = NuevoControl("Forms.Label.1", MARGEN, MARGEN, ANCHO - 2 * MARGEN)
    NuevoControl("Forms.Label.1", MARGEN, 36, 200).Caption = "Tipo de forma"
    Set optCirculo = NuevaOpcion("Círculo", 24, 54, "Tipo")
    Set optElipse = NuevaOpcion("Elipse", 130, 54, "Tipo")
    Set optCuadrado = NuevaOpcion("Cuadrado", 24, 72, "Tipo")
    Set optRectangulo = NuevaOpcion("Rectángulo", 130, 72, "Tipo")
    Set optRombo = NuevaOpcion("Rombo / cometa", 24, 90, "Tipo")
    NuevoControl("Forms.Label.1", MARGEN, 118, 200).Caption = "Medidas (mm)"
    Set lblMedida1 = NuevoControl("Forms.Label.1", 24, 0, 100)
    Set txtMedida1 = NuevoControl("Forms.TextBox.1", 130, 0, 70)
    Set lblMedida2 = NuevoControl("Forms.Label.1", 24, 0, 100)
    Set txtMedida2 = NuevoControl("Forms.TextBox.1", 130, 0, 70)
    Set lblAlTerminar = NuevoControl("Forms.Label.1", MARGEN, 0, 200)
    lblAlTerminar.Caption = "Al terminar"
    Set optNada = NuevaOpcion("Nada", 24, 0, "Final", 70)
    Set optAgrupar = NuevaOpcion("Agrupar", 100, 0, "Final", 70)
    Set optCombinar = NuevaOpcion("Combinar", 176, 0, "Final", 70)
    Set lblAviso = NuevoControl("Forms.Label.1", MARGEN, 0, ANCHO - 2 * MARGEN)
    lblAviso.ForeColor = RGB(192, 0, 0)
    lblAviso.WordWrap = True
    Set btnCancelar = NuevoControl("Forms.CommandButton.1", 0, 0, 75)
    btnCancelar.Caption = "Cancelar"
    btnCancelar.Height = 24 'Los botones necesitan más alto que el resto de controles
    btnCancelar.Cancel = True 'Esc
    Set btnAplicar = NuevoControl("Forms.CommandButton.1", 0, 0, 75)
    btnAplicar.Caption = "Aplicar"
    btnAplicar.Height = 24
    btnAplicar.Default = True 'Enter

    CargarConfiguracion
End Sub

'Debe llamarse antes de Show con la selección que se va a redimensionar
Public Sub Preparar(sr As ShapeRange)
    Dim n As Long
    Set mSeleccion = sr
    n = ContarFormas(sr)
    lblObjetos.Caption = n & IIf(n = 1, " forma seleccionada", " formas seleccionadas")
End Sub

Private Function NuevoControl(progId As String, x As Single, y As Single, ancho As Single) As Object
    Dim c As Object
    Set c = Me.Controls.Add(progId)
    c.Left = x
    c.Top = y
    c.Width = ancho
    c.Height = 18
    c.Font.Name = FUENTE
    c.Font.Size = 9
    Set NuevoControl = c
End Function

Private Function NuevaOpcion(texto As String, x As Single, y As Single, grupo As String, _
                             Optional ancho As Single = 105) As MSForms.OptionButton
    Set NuevaOpcion = NuevoControl("Forms.OptionButton.1", x, y, ancho)
    NuevaOpcion.Caption = texto
    NuevaOpcion.GroupName = grupo 'Cada grupo de opciones se marca por separado
    NuevaOpcion.BackStyle = fmBackStyleTransparent
End Function

'Cambia los textos de las medidas según el tipo y recoloca los controles de debajo
Private Sub ActualizarCampos()
    Dim y As Single, dos As Boolean

    Select Case TipoElegido
        Case tfCirculo: lblMedida1.Caption = "Diámetro"
        Case tfElipse: lblMedida1.Caption = "Eje largo": lblMedida2.Caption = "Eje corto"
        Case tfCuadrado: lblMedida1.Caption = "Lado"
        Case tfRectangulo: lblMedida1.Caption = "Lado largo": lblMedida2.Caption = "Lado corto"
        Case tfRombo: lblMedida1.Caption = "Diagonal larga": lblMedida2.Caption = "Diagonal corta"
    End Select
    dos = TipoElegido = tfElipse Or TipoElegido = tfRectangulo Or TipoElegido = tfRombo

    y = 138
    lblMedida1.Top = y + 2
    txtMedida1.Top = y
    y = y + 24
    lblMedida2.Visible = dos
    txtMedida2.Visible = dos
    If dos Then
        lblMedida2.Top = y + 2
        txtMedida2.Top = y
        y = y + 24
    End If
    lblAlTerminar.Top = y + 6
    optNada.Top = y + 24
    optAgrupar.Top = y + 24
    optCombinar.Top = y + 24
    lblAviso.Top = y + 48
    lblAviso.Height = 40
    btnAplicar.Top = y + 94
    btnCancelar.Top = y + 94
    btnAplicar.Left = ANCHO - MARGEN - 75 - 6
    btnCancelar.Left = btnAplicar.Left - 75 - 6
    Me.Height = btnAplicar.Top + 24 + MARGEN + (Me.Height - Me.InsideHeight)

    Validar
End Sub

Private Function TipoElegido() As TipoForma
    If optElipse.Value Then
        TipoElegido = tfElipse
    ElseIf optCuadrado.Value Then
        TipoElegido = tfCuadrado
    ElseIf optRectangulo.Value Then
        TipoElegido = tfRectangulo
    ElseIf optRombo.Value Then
        TipoElegido = tfRombo
    Else
        TipoElegido = tfCirculo
    End If
End Function

Private Function AccionElegida() As AccionFinal
    If optAgrupar.Value Then
        AccionElegida = afAgrupar
    ElseIf optCombinar.Value Then
        AccionElegida = afCombinar
    Else
        AccionElegida = afNada
    End If
End Function

'Comprueba las medidas, muestra el aviso y activa Aplicar solo si todo es correcto
Private Sub Validar()
    Dim m1 As Double, m2 As Double, aviso As String

    If Not LeerMedida(txtMedida1.Text, m1) Then
        aviso = "Introduce en """ & lblMedida1.Caption & """ un número mayor que 0."
    ElseIf txtMedida2.Visible Then
        If Not LeerMedida(txtMedida2.Text, m2) Then
            aviso = "Introduce en """ & lblMedida2.Caption & """ un número mayor que 0."
        ElseIf m2 > m1 Then
            aviso = """" & lblMedida1.Caption & """ no puede ser menor que """ & lblMedida2.Caption & """."
        End If
    End If

    lblAviso.Caption = aviso
    btnAplicar.Enabled = (aviso = "")
End Sub

Private Sub optCirculo_Click(): ActualizarCampos: End Sub
Private Sub optElipse_Click(): ActualizarCampos: End Sub
Private Sub optCuadrado_Click(): ActualizarCampos: End Sub
Private Sub optRectangulo_Click(): ActualizarCampos: End Sub
Private Sub optRombo_Click(): ActualizarCampos: End Sub
Private Sub txtMedida1_Change(): Validar: End Sub
Private Sub txtMedida2_Change(): Validar: End Sub

Private Sub btnAplicar_Click()
    Dim motivo As String
    motivo = ComprobarFormas(mSeleccion, TipoElegido)
    If motivo <> "" Then 'El formulario sigue abierto para cambiar el tipo o cancelar
        lblAviso.Caption = motivo
        Exit Sub
    End If
    Tipo = TipoElegido
    LeerMedida txtMedida1.Text, Medida1
    If txtMedida2.Visible Then LeerMedida txtMedida2.Text, Medida2 Else Medida2 = Medida1
    AlTerminar = AccionElegida
    Aceptado = True
    GuardarConfiguracion
    Me.Hide
End Sub

Private Sub btnCancelar_Click()
    Me.Hide
End Sub

'La X de la ventana equivale a Cancelar
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub CargarConfiguracion()
    Select Case CLng(GetSetting(APP, "Ultimo", "Tipo", CStr(tfCirculo)))
        Case tfElipse: optElipse.Value = True
        Case tfCuadrado: optCuadrado.Value = True
        Case tfRectangulo: optRectangulo.Value = True
        Case tfRombo: optRombo.Value = True
        Case Else: optCirculo.Value = True
    End Select
    txtMedida1.Text = GetSetting(APP, "Ultimo", "Medida1", "")
    txtMedida2.Text = GetSetting(APP, "Ultimo", "Medida2", "")
    Select Case CLng(GetSetting(APP, "Ultimo", "AlTerminar", CStr(afNada)))
        Case afAgrupar: optAgrupar.Value = True
        Case afCombinar: optCombinar.Value = True
        Case Else: optNada.Value = True
    End Select
    ActualizarCampos
End Sub

Private Sub GuardarConfiguracion()
    SaveSetting APP, "Ultimo", "Tipo", CStr(Tipo)
    SaveSetting APP, "Ultimo", "Medida1", txtMedida1.Text
    SaveSetting APP, "Ultimo", "Medida2", txtMedida2.Text
    SaveSetting APP, "Ultimo", "AlTerminar", CStr(AlTerminar)
End Sub
