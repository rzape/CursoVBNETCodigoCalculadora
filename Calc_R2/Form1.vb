'######################################################################################################
' NOTAS
' MODIF_PROG: 2023-07-03  La visualización por pantalla de los valores calculados se debe
'                         trabajar en una subrutina a fin de truncar o limitar los valores
'                         a visualizar.
' MODIF_PROG: 2023-07-06  Se corrige codigo para sumar operandos de tipo entero y fraccion
' 
' MODIF_PROG: 2023-07-09  Se separa el código en módulos.
'                         Se rediseña para trabajar con máquina de estados.
'                         Se agregan y/o modifican operaciones "+","-","x" y "%"
'                         Queda para agregar operaciones "/","sqrt","pot" y corregir operacion "CE"
'######################################################################################################
Public Class Form1


    '=========================================================================================================
    ' CODIGO ASIGNADO A BOTONES DE OPERACIONES
    '=========================================================================================================
    Private Sub btnSumar_Click(sender As Object, e As EventArgs) Handles btnSumar.Click
        ' Evento por operacion, no por ingreso de numero        
        bNuevaOperacion = True
        sBoton = "boton_+"
        SM_Calculadora()
    End Sub
    Private Sub btnRestar_Click(sender As Object, e As EventArgs) Handles btnRestar.Click
        ' Evento por operacion, no por ingreso de numero        
        bNuevaOperacion = True
        sBoton = "boton_-"
        SM_Calculadora()
    End Sub
    Private Sub btnMultiplicar_Click(sender As Object, e As EventArgs) Handles btnMultiplicar.Click
        ' Evento por operacion, no por ingreso de numero        
        bNuevaOperacion = True
        sBoton = "boton_x"
        SM_Calculadora()
    End Sub
    Private Sub btnDividir_Click(sender As Object, e As EventArgs) Handles btnDividir.Click
    End Sub
    Private Sub btnPorcentaje_Click(sender As Object, e As EventArgs) Handles btnPorcentaje.Click
        'Se pulsó botón para nueva operación
        bNuevaOperacion = True
        sBoton = "boton_%"
        SM_Calculadora()
    End Sub
    Private Sub btnPotenciar_Click(sender As Object, e As EventArgs) Handles btnPotenciar.Click
    End Sub
    Private Sub btnRaizCuadrada_Click(sender As Object, e As EventArgs) Handles btnRaizCuadrada.Click
    End Sub
    Private Sub btnFraccion_Click(sender As Object, e As EventArgs) Handles btnFraccion.Click
        '--------------------------------------------------------------
        'NO se permite el uso del separador decimal en modo fracción.
        '--------------------------------------------------------------------
        ' Si ya se encontraba en modo fracción
        If bModoFraccion = True Then
            ' luego de un resultado, habrá una fracción en el display...
            If bResultadoEnDisplay = True Then
                ' Se sale del modo fracción y se pasa a modo decimal (solo despues de un resultado). 
                bModoFraccion = False
                'btnSeparadorDecimal.Enabled = True
                ' En esta versión, se debe reiniciar el display, no se convierte a decimal
                ReinicioVariables()
                ' pasa el foco al elemento tb_Display (barra parpadeando)
                tbDisplay.Select()
                Exit Sub
            End If
            ' ...O se estaba ingresando una fracción
        Else
            ' Si se encontraba en decimal (no fracción) se pasa a fracción
            bModoFraccion = True
            'btnSeparadorDecimal.Enabled = False
        End If
        ' Ingreso la fracción, el número en display al array que contiene los componentes de la fracción
        ' Debe comenzar en cero!!
        '-------------------------------------------------------------------------------------------
        ' Para convertir, se debe leer el string, tomar el último caracter y convertirlo en número
        ' ##### 1.- ERROR HAY QUE TOMAR LOS CARACTERES ENTRE LOS SIMBOLOS DE FRACCION
        ' ##### 2.- TAMPOCO GUARDAR OPERANDOS SI SE SUPERA EL CONTADOR.
        ' ##### 3.- LOS OPERANDOS HAY QUE CONVERTIRLOS A NUMERO.
        ' O se debió trabajar todo en números y luego convertir para volcar en display
        '-------------------------------------------------------------------------------------------
        'iFraccionOperandos(iContadorBarraFraccional) = Mid(tbDisplay.Text, Len(tbDisplay.Text), 1)
        ' Aumento contador de barras de fracción (valor inicial cero, máximo aceptable como fracción dos)
        iContadorBarraFraccional += 1
        ' Verifico si anteriormente ya había ingresado componentes de fracción; hasta 3 componentes
        ' permitido, se toma como fracción impropia. (rango contador = 0..2)
        ' NOTA Son tres componentes máximo y dos barras máximo: a∟b∟c
        ' Si es mayor o igual a 2, hay un error matemático, se avisa a traves de la variable booleana
        ' Esto se maneja al solicitar resultado, mostrando error
        If iContadorBarraFraccional >= 3 Then
            bFraccionErrorIngreso = True
            'Exit Sub
        End If
        ' Agrego el caracter de fraccion
        tbDisplay.Text = tbDisplay.Text & "∟"
        ' Indica que hay un punto decimal
        bBarraFraccionPresente = True
        ' Indica que no es cero el dígito
        bDigitoActualContieneCero = False
        ' Indica que no es el primero
        bPosicionDigitoActualEsElPrimero = False
        ' Se Ingresó un número o simbolo
        bResultadoEnDisplay = False
        ' pasa el foco al elemento tb_Display (barra parpadeando)
        tbDisplay.Select()
    End Sub
    Private Sub btnClearEntry_Click(sender As Object, e As EventArgs) Handles btnClearEntry.Click
        '######
        ' REVISAR genera error por barra de fraccion
        '######
        '' borra el display y flag punto decimal presente
        'Reinicia_tb_Display()
        'MuestraCeroEnDisplay()
        '' pasa el foco al elemento tb_Display (barra parpadeando)
        'tbDisplay.Select()
    End Sub
    Private Sub btnClearOp_Click(sender As Object, e As EventArgs) Handles btnClearOp.Click
        ' borra operandos en memoria
        BorraMemoriaOperandos()
        ' borra operación
        bOperacionEnCurso = False
        sOperacion = ""
        ' borra el display y flag punto decimal presente
        Reinicia_tb_Display()
        MuestraCeroEnDisplay()
        ' Esta subrutina es más completa, verificar.
        ReinicioVariables()
        ' pasa el foco al elemento tb_Display (barra parpadeando)
        tbDisplay.Select()
    End Sub
    Private Sub btnResultado_Click(sender As Object, e As EventArgs) Handles btnResultado.Click
        'Se pulsó botón para nueva operación
        bNuevaOperacion = True
        sBoton = "boton_="
        SM_Calculadora()
    End Sub


    '//////////////////////////////////////
    ' SUBRUTINAS
    '//////////////////////////////////////

    Public Sub IngresoOperandoA()
        'El operando ingresado es fracción?

        If bModoFraccion = True Then
            '######################################################
            ' LA VERIFICACION DEBE HACERSE EN SM_Calculadora()
            '######################################################
            'verifica que el denominador no sea cero (cond. error)
            'Si es válido, entonces lo ingresa a memoria, sino modo error
            VerificacionCondicionErrorFraccion()
            'modo fracción, guardar en memoria formato decimal
            AlmacenaOperandoDisplayFraccionEnMemoria(iFraccionA(0), iFraccionA(1), iFraccionA(2))
            ' Indica que es primer operando es fracción
            bPriOperTipoFraccion = True
            'reinicia modo ingreso, es decimal por defecto.
            bModoFraccion = False
            ' Separador decimal se encuentra desafectado.
            bSeparadorDecimalDeshabilitado = True
            '###################################################
            'Borra pantalla
            Reinicia_tb_Display()
            MuestraCeroEnDisplay()
            '#######################################################
        Else
            'modo decimal, guardar en memoria formato decimal
            AlmacenaOperandoDisplayEnteroEnMemoria("A")
            'Indica que es primer operando es entero (en principio lo trunca)
            bPriOperTipoFraccion = False
        End If
        'OPA leido de pantalla, no copiado de resultado operacion anterior.
        bOperando_aCopiadoDeOperando_r = False

        'Resetea el contador de fracciones
        'Para poder ingresar una nueva fracción.
        iContadorBarraFraccional = 0

    End Sub
    Public Sub FinalizaOperacionEjecucion()

        ' Almacena formato operacion previa
        bOperacionFraccionalPrevia = bOperacionFraccional

        If bOperacionFraccional = True Then

            ' Inicializa variables de operaciones con fracciones

            ' Si queda el resultado en display, se debe recordar que es
            ' una fracciòn

            bFlagAlmacenarOperandoFraccionDos = False
            'bBarraFraccionPresente = False
            'bModoFraccion = False
            bIngresoComponenteFraccion = False
            bIngresoOperandoFraccionUno = False
            bIngresoOperandoFraccionDos = False
            'iContadorBarraFraccional = 0
            bFraccionErrorIngreso = False
            ' almacena el resultado como el operando a
            'iFraccionA(0) = 0
            'iFraccionA(1) = F2b
            'iFraccionA(2) = F2c

            ' ####ya realizada la operacion, no se guarda en memoria
            'bOperacionEnCurso = False
            'Reinicia formato salida
            'bOperacionFraccional = False
            '### Se ejecutó operación, esperar a nueva o repetir anterior
            'bNuevaOperacion = False
            '### reinicializa estados de primer dígito y punto decimal presente
            'Reinicia_tb_Display()
            '### pasa el foco al elemento tb_Display (barra parpadeando)
            'tbDisplay.Select()
        Else
            ' almacena el resultado como el operando a
            'AlmacenaOperandoDisplayEnMemoria("A")
            ' ya realizada la operacion, no se guarda en memoria
            bOperacionEnCurso = False
            'Reinicia formato salida
            'bOperacionFraccional = False
            ' Se ejecutó operación, esperar a nueva o repetir anterior
            bNuevaOperacion = False
            ' reinicializa estados de primer dígito y punto decimal presente
            Reinicia_tb_Display()
            ' pasa el foco al elemento tb_Display (barra parpadeando)
            tbDisplay.Select()

        End If
    End Sub
    Sub VerificacionCondicionErrorFraccion()
        If bFraccionErrorIngreso = True Then
            'Si error, entonces
            tbDisplay.Text = msgErrMat
            'Sale del modo fraccion
            bModoFraccion = False
            ' Borra display en proximo ingreso de un numero
            bFlagBorraDisplayEnProximoIngreso = True
            ' pasa el foco al elemento tb_Display (barra parpadeando)
            tbDisplay.Select()
            Exit Sub
        End If
    End Sub
    Public Sub NivelaTipoDeOperandos()
        'Ingreso el valor del display
        'Si es fraccion
        If bModoFraccion = True Then
            'y no hay error 
            VerificacionCondicionErrorFraccion()

            'Lo guarda en memoria (fraccion)
            AlmacenaOperandoDisplayFraccionEnMemoria(iFraccionB(0), iFraccionB(1), iFraccionB(2))
            ' Indica que es segundo operando es fracción
            bSegOperTipoFraccion = True
            'reinicia modo ingreso, es decimal por defecto.
            bModoFraccion = False
            ' Separador decimal se encuentra desafectado.
            bSeparadorDecimalDeshabilitado = True

            '-------------------------------------------------------------
            'modo fracción en el segundo operando, primer operario entero
            '-------------------------------------------------------------
            If bPriOperTipoFraccion = False Then
                ' Si el primer operando era un decimal:
                ' 1.- El formato del primer operando se debe cambiar a fracción
                '##################
                'ver de truncar
                '##################
                iFraccionA(0) = 0
                iFraccionA(1) = dOperando_a
                iFraccionA(2) = 1
                ' 2.- Se debe cargar el segundo operando fraccional
                ' AlmacenaOperandoFraccionDosEnMemoria()
            End If


            'Las operaciones a ejecutar son con ambos operarios expresados en fracciones
            bOperacionFraccional = True
        Else
            ' Almacena operando b en memoria (el número del display, una sola vez,
            ' excepto que ingrese nueva operacion
            AlmacenaOperandoDisplayEnteroEnMemoria("B")
            'Las operaciones a ejecutar son con ambos operarios expresados en decimales
            bOperacionFraccional = False

            ' End If
            '#################################################
            If bPriOperTipoFraccion = True Then
                ' Si el primer operando era un decimal:
                ' 1.- El formato del primer operando se debe cambiar a fracción
                '##################
                'ver de truncar
                '##################
                iFraccionB(0) = 0
                iFraccionB(1) = dOperando_b
                iFraccionB(2) = 1
                ' 2.- Se debe cargar el segundo operando fraccional
                ' AlmacenaOperandoFraccionDosEnMemoria()
                ' Indica que es segundo operando es fracción
                bSegOperTipoFraccion = True
                'reinicia modo ingreso, es decimal por defecto.
                bModoFraccion = False
                ' Separador decimal se encuentra desafectado.
                bSeparadorDecimalDeshabilitado = True
                'Las operaciones a ejecutar son con ambos operarios expresados en fracciones
                bOperacionFraccional = True
            End If
        End If

    End Sub
    Public Sub CopiaResultadoEnOpA()
        'sino, copia resultado en operando a
        If bOperacionFraccional = True Then
            iFraccionA(0) = iFraccionR(0)
            iFraccionA(1) = iFraccionR(1)
            iFraccionA(2) = iFraccionR(2)
            'indica primer operando fraccion
            bPriOperTipoFraccion = True
            '## agregado, se producia error al multiplicar
            'Resetea el contador de fracciones
            'Para poder ingresar una nueva fracción.
            iContadorBarraFraccional = 0
        Else
            'indica primer operando entero
            bPriOperTipoFraccion = False
            'copia el valor
            dOperando_a = dOperando_r
        End If
    End Sub
    Public Sub CopiaResultadoEnOpA_REV_2()
        If bResultadoTipoFraccion = True Then
            'iFraccionA(0) = 0
            iFraccionA(0) = iFraccionR(0)
            iFraccionA(1) = iFraccionR(1)
            iFraccionA(2) = iFraccionR(2)
            'indica primer operando fraccion
            bPriOperTipoFraccion = True
            'verifica que no haya cero en el denominador
            If iFraccionA(2) = 0 Then
                bFraccionErrorIngreso = True
            End If
            '## agregado, se producia error al multiplicar
            'Resetea el contador de fracciones
            'Para poder ingresar una nueva fracción.
            iContadorBarraFraccional = 0
        Else
            'indica primer operando entero
            bPriOperTipoFraccion = False
            'copia el valor
            dOperando_a = dOperando_r
            ' operando en memoria
            bOperando_a_EnMemoria = True
        End If
        bOperando_aCopiadoDeOperando_r = True
    End Sub
    Sub BorraMemoriaOperandos()
        ' almacena "0" en memoria a
        dOperando_a = Convierte(0)
        ' idem memoria b
        dOperando_b = dOperando_a
        ' y borra flags de memoria ocupada
        bOperando_a_EnMemoria = False
        bOperando_b_EnMemoria = False
    End Sub
    Sub ReinicioVariables()
        ' No hay operaciones en curso
        bOperacionEnCurso = False
        ' Inicializa sin operando_a en memoria
        bOperando_a_EnMemoria = False
        ' Posición digito = primero, sin coma
        Reinicia_tb_Display()
        ' Se inicia mostrando un cero.
        MuestraCeroEnDisplay()
        ' Se ingresó un número en display (cero)
        bNumeroEn_tb_Display = True
        ' No borra pantalla en proximo ingreso
        bFlagBorraDisplayEnProximoIngreso = False
        ' borra almacenar operando b
        'bFlagAlmacenarOperando_b = False
        ' borra operacion
        sOperacion = ""
        ' Inicializa calcular porcentajer
        bCalcularPorcentaje = False
        ' Inicializa variables de operaciones con fracciones
        bFlagAlmacenarOperandoFraccionDos = False
        bBarraFraccionPresente = False
        bModoFraccion = False
        bIngresoComponenteFraccion = False
        bIngresoOperandoFraccionUno = False
        bIngresoOperandoFraccionDos = False
        iContadorBarraFraccional = 0
        bFraccionErrorIngreso = False
        ' Flag para verificar que se solicitó resultado
        bResultadoEnDisplay = False
        ' Flag de ingreso numero
        bNumeroIngresado = False
        bSeparadorDecimalIngresado = False
        ' Indico que los operandos inicialmente son del tipo entero, no fraccion
        bPriOperTipoFraccion = False
        bSegOperTipoFraccion = False
        bSeparadorDecimalDeshabilitado = False
        bOperacionFraccional = False
        bResultadoTipoFraccion = False
        bOperando_aCopiadoDeOperando_r = False
        bCalcPorcentajeExtension = False
        bNuevaOperacion = False
        bPorcentajeEnDisplay = False
        bTasaDeCambioEnDisplay = False
        iSMCalculadora = S000
        sBoton = ""
    End Sub

    '=========================================================================================================
    ' CODIGO ASIGNADO A BOTONES NUMERICOS Y SEPARADOR DECIMAL
    '=========================================================================================================
    Private Sub btnSeparadorDecimal_Click(sender As Object, e As EventArgs) Handles btnSeparadorDecimal.Click
        CodigoSeparadorDecimal()
        SM_Calculadora()
    End Sub
    Private Sub btnCero_Click(sender As Object, e As EventArgs) Handles btnCero.Click
        CodigoNumeroCero()
        SM_Calculadora()
    End Sub
    Private Sub btnUno_Click(sender As Object, e As EventArgs) Handles btnUno.Click
        CodigoNumeroUno()
        SM_Calculadora()
    End Sub
    Private Sub btnDos_Click(sender As Object, e As EventArgs) Handles btnDos.Click
        CodigoNumeroDos()
        SM_Calculadora()
    End Sub
    Private Sub btnTres_Click(sender As Object, e As EventArgs) Handles btnTres.Click
        CodigoNumeroTres()
        SM_Calculadora()
    End Sub
    Private Sub btnCuatro_Click(sender As Object, e As EventArgs) Handles btnCuatro.Click
        CodigoNumeroCuatro()
        SM_Calculadora()
    End Sub
    Private Sub btnCinco_Click(sender As Object, e As EventArgs) Handles btnCinco.Click
        CodigoNumeroCinco()
        SM_Calculadora()
    End Sub
    Private Sub btnSeis_Click(sender As Object, e As EventArgs) Handles btnSeis.Click
        CodigoNumeroSeis()
        SM_Calculadora()
    End Sub
    Private Sub btnSiete_Click(sender As Object, e As EventArgs) Handles btnSiete.Click
        CodigoNumeroSiete()
        SM_Calculadora()
    End Sub
    Private Sub btnOcho_Click(sender As Object, e As EventArgs) Handles btnOcho.Click
        CodigoNumeroOcho()
        SM_Calculadora()
    End Sub
    Private Sub btnNueve_Click(sender As Object, e As EventArgs) Handles btnNueve.Click
        CodigoNumeroNueve()
        SM_Calculadora()
    End Sub

    '=========================================================================================================
    ' CODIGO ASIGNADO A INGRESO POR TECLADO NUMERICO
    '=========================================================================================================
    ' El código se toma desde el evento KeyDown del textbox del display.
    '
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles tbDisplay.KeyDown
        ' Si se ingresa por teclado numérico el "0"
        'If e.KeyCode = Keys.NumPad0 Then
        'Muestro un cero
        'TextBox1.Text = 0
        'o un codigo en dos lineas, forma 1
        'TextBox1.Text= "2+2=" & VbcrLf & "4"
        'o un codigo en dos lineas, forma 2
        'Dim linea(2) As String
        'linea(0) = "this is line1"
        ' linea(1) = "this is line2"
        ' TextBox1.Lines = linea
        'CodigoNumeroCero()
        'End If
        Select Case e.KeyCode
            Case Keys.Decimal
                CodigoSeparadorDecimal()
            Case Keys.NumPad0
                CodigoNumeroCero()
            Case Keys.NumPad1
                CodigoNumeroUno()
            Case Keys.NumPad2
                CodigoNumeroDos()
            Case Keys.NumPad3
                CodigoNumeroTres()
            Case Keys.NumPad4
                CodigoNumeroCuatro()
            Case Keys.NumPad5
                CodigoNumeroCinco()
            Case Keys.NumPad6
                CodigoNumeroSeis()
            Case Keys.NumPad7
                CodigoNumeroSiete()
            Case Keys.NumPad8
                CodigoNumeroOcho()
            Case Keys.NumPad9
                CodigoNumeroNueve()

        End Select
    End Sub

    '=========================================================================================================
    ' INICIALIZACION
    '=========================================================================================================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' El texto del formulario se alinea a la derecha
        ' El tipo y tamaño de texto es "console", 16pt
        ' Ver form designer

        ' Deshabilita la escritura en el cuadro de texto display
        tbDisplay.ReadOnly = True
        ' Reinicia todos los componentes
        ReinicioVariables()
    End Sub
    '=========================================================================================================
    ' FUNCIONES
    '=========================================================================================================
    Public Function Convierte(ByVal Texto As String) As Double
        ' Se busca en la cadena recibida como parámetro el caracter coma ","
        ' y se lo reemplaza por el caracter punto "." para luego convertir
        ' el número recibido en formato texto a formato valor (numérico)
        Convierte = Val(Replace(Texto, ",", "."))
    End Function

End Class