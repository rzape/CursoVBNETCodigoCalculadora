﻿' NOTAS
' MP1: 2023-07-03  La visualización por pantalla de los valores calculados se debe
'                  trabajar en una subrutina a fin de truncar o limitar los valores
'                  a visualizar.
Public Class Form1
    Private Const msgErrMat As String = "ERROR matemático"
    Private Const msgErrStx As String = "ERROR de sintaxis"
    '//////////////////////////////////////
    ' Variables
    '//////////////////////////////////////
    Dim bPuntoDecimalPresente As Boolean
    Dim bPosicionDigitoActualEsElPrimero As Boolean
    Dim bDigitoActualContieneCero As Boolean
    Dim bNumeroEn_tb_Display As Boolean
    Dim bOperando_a_EnMemoria As Boolean
    Dim bOperando_b_EnMemoria As Boolean
    Dim bFlagBorraDisplayEnProximoIngreso As Boolean
    Dim bOperacionEnCurso As Boolean
    Dim bFlagAlmacenarOperando_b As Boolean
    Dim bCalcularPorcentaje As Boolean
    Dim bModoFraccion As Boolean
    Dim bIngresoComponenteFraccion As Boolean
    Dim bIngresoOperandoFraccionUno As Boolean
    Dim bIngresoOperandoFraccionDos As Boolean
    Dim bFraccionErrorIngreso As Boolean
    Dim bResultadoEnDisplay As Boolean
    Dim bNumeroIngresado As Boolean
    Dim bBarraFraccionPresente As Boolean
    Dim bFlagAlmacenarOperandoFraccionDos As Boolean




    Dim iContadorBarraFraccional As Integer
    Dim iFraccionA(3) As Integer
    Dim iFraccionB(3) As Integer


    Dim dOperando_a As Double
    Dim dOperando_b As Double
    Dim sOperacion As String


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
    ' CODIGO ASIGNADO A BOTONES DE OPERACIONES
    '=========================================================================================================
    Private Sub btnSumar_Click(sender As Object, e As EventArgs) Handles btnSumar.Click
        ' Si no hay una operacion previa en curso, se guarda en memoria
        ' el valor que se encuentre en display, y se asigna la operacion
        ' sino solamente se modificará la operación a realizar con el operando
        ' en memoria.
        ' Si no se ingreso un numero nuevo (entonces en display esta el resultado
        ' de una operación anterior)
        If bOperacionEnCurso = False Then
            If bModoFraccion = True Then
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
                'modo fracción, guardar en memoria formato decimal
                AlmacenaOperandoFraccionUnoEnMemoria()

                'Borra pantalla
                Reinicia_tb_Display()
                MuestraCeroEnDisplay()
                'Resetea el contador de fracciones
                iContadorBarraFraccional = 0
            Else
                'modo decimal, guardar en memoria formato decimal
                AlmacenaOperando_aEnMemoria()
            End If
            ' Avisa que hay una operacion en curso
            bOperacionEnCurso = True
        End If
        ' define la operación a realizar como suma
        sOperacion = "adicion"
        If bModoFraccion = True Then
            'modo fracción

            'habilita almacenar el proximo numero Fraccion Dos
            bFlagAlmacenarOperandoFraccionDos = True
        Else
            ' modo decimal...
            ' habilita almacenar el proximo numero como operando b 
            bFlagAlmacenarOperando_b = True
        End If

        ' Borra display en proximo ingreso de un numero
        bFlagBorraDisplayEnProximoIngreso = True
        ' pasa el foco al elemento tb_Display (barra parpadeando)
        tbDisplay.Select()
    End Sub

    Private Sub btnRestar_Click(sender As Object, e As EventArgs) Handles btnRestar.Click
        ' Si no hay una operacion previa en curso, se guarda en memoria
        ' el valor que se encuentre en display, y se asigna la operacion
        ' sino solamente se modificará la operación a realizar con el operando
        ' en memoria.
        If bOperacionEnCurso = False Then
            If bModoFraccion = True Then
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
                'modo fracción, guardar en memoria formato decimal
                AlmacenaOperandoFraccionUnoEnMemoria()

                'Borra pantalla
                Reinicia_tb_Display()
                MuestraCeroEnDisplay()
                'Resetea el contador de fracciones
                iContadorBarraFraccional = 0
            Else
                'modo decimal, guardar en memoria formato decimal
                AlmacenaOperando_aEnMemoria()
            End If
            ' Avisa que hay una operacion en curso
            bOperacionEnCurso = True
        End If
        ' define la operación a realizar como resta
        sOperacion = "sustraccion"

        If bModoFraccion = True Then
            'modo fracción

            'habilita almacenar el proximo numero Fraccion Dos
            bFlagAlmacenarOperandoFraccionDos = True
        Else
            ' modo decimal...
            ' habilita almacenar el proximo numero como operando b 
            bFlagAlmacenarOperando_b = True
        End If

        ' borra el display
        bFlagBorraDisplayEnProximoIngreso = True
        ' pasa el foco al elemento tb_Display (barra parpadeando)
        tbDisplay.Select()
    End Sub

    Private Sub btnMultiplicar_Click(sender As Object, e As EventArgs) Handles btnMultiplicar.Click
        ' Si no hay una operacion previa en curso, se guarda en memoria
        ' el valor que se encuentre en display, y se asigna la operacion
        ' sino solamente se modificará la operación a realizar con el operando
        ' en memoria.
        If bOperacionEnCurso = False Then
            ' Si no se ingreso un numero nuevo (entonces en display esta el resultado
            ' de una operación anterior)
            AlmacenaOperando_aEnMemoria()
            ' Avisa que hay una operacion en curso
            bOperacionEnCurso = True
        End If
        ' define la operación a realizar como multiplicacion
        sOperacion = "multiplicacion"
        ' borra el display, en cuanto ingrese un número (código en subrutinas de número)
        bFlagBorraDisplayEnProximoIngreso = True
        ' habilita almacenar el proximo numero como operando b  
        bFlagAlmacenarOperando_b = True
        ' pasa el foco al elemento tb_Display (barra parpadeando)
        tbDisplay.Select()
    End Sub

    Private Sub btnDividir_Click(sender As Object, e As EventArgs) Handles btnDividir.Click
        ' Si no hay una operacion previa en curso, se guarda en memoria
        ' el valor que se encuentre en display, y se asigna la operacion
        ' sino solamente se modificará la operación a realizar con el operando
        ' en memoria.
        If bOperacionEnCurso = False Then
            ' Si no se ingreso un numero nuevo (entonces en display esta el resultado
            ' de una operación anterior)
            AlmacenaOperando_aEnMemoria()
            ' Avisa que hay una operacion en curso
            bOperacionEnCurso = True
        End If
        ' define la operación a realizar como division
        sOperacion = "division"
        ' borra el display
        bFlagBorraDisplayEnProximoIngreso = True
        ' habilita almacenar el proximo numero como operando b 
        bFlagAlmacenarOperando_b = True
        ' pasa el foco al elemento tb_Display (barra parpadeando)
        tbDisplay.Select()
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

    Private Sub btnPorcentaje_Click(sender As Object, e As EventArgs) Handles btnPorcentaje.Click
        bCalcularPorcentaje = True
        EjecutaOperaciones()
    End Sub

    Private Sub btnPotenciar_Click(sender As Object, e As EventArgs) Handles btnPotenciar.Click
        ' Si no hay una operacion previa en curso, se guarda en memoria
        ' el valor que se encuentre en display, y se asigna la operacion
        ' sino solamente se modificará la operación a realizar con el operando
        ' en memoria.
        If bOperacionEnCurso = False Then
            ' Si no se ingreso un numero nuevo (entonces en display esta el resultado
            ' de una operación anterior)
            AlmacenaOperando_aEnMemoria()
            ' Avisa que hay una operacion en curso
            bOperacionEnCurso = True
        End If
        ' define la operación a realizar como suma
        sOperacion = "potenciacion"
        ' habilita almacenar el proximo numero como operando b 
        bFlagAlmacenarOperando_b = True
        ' Borra display en proximo ingreso de un numero
        bFlagBorraDisplayEnProximoIngreso = True
        ' pasa el foco al elemento tb_Display (barra parpadeando)
        tbDisplay.Select()
    End Sub

    Private Sub btnRaizCuadrada_Click(sender As Object, e As EventArgs) Handles btnRaizCuadrada.Click
        ' Si no hay una operacion previa en curso, se guarda en memoria
        ' el valor que se encuentre en display, y se asigna la operacion
        ' sino solamente se modificará la operación a realizar con el operando
        ' en memoria.
        If bOperacionEnCurso = False Then
            ' Si no se ingreso un numero nuevo (entonces en display esta el resultado
            ' de una operación anterior)
            AlmacenaOperando_aEnMemoria()
            ' Avisa que hay una operacion en curso
            bOperacionEnCurso = True
        End If
        ' define la operación a realizar como suma
        sOperacion = "RaizCuadrada"
        ' y ejecuta la operación
        EjecutaOperaciones()
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
        ' borra el display y flag punto decimal presente
        Reinicia_tb_Display()
        MuestraCeroEnDisplay()
        ' pasa el foco al elemento tb_Display (barra parpadeando)
        tbDisplay.Select()
    End Sub
    Private Sub btnResultado_Click(sender As Object, e As EventArgs) Handles btnResultado.Click
        EjecutaOperaciones()
        ' Se generó un resultado
        bResultadoEnDisplay = True
        ' bOperando_a_EnMemoria = False
    End Sub


    '=========================================================================================================
    ' CODIGO ASIGNADO A BOTONES NUMERICOS Y SEPARADOR DECIMAL
    '=========================================================================================================
    Private Sub btnSeparadorDecimal_Click(sender As Object, e As EventArgs) Handles btnSeparadorDecimal.Click
        CodigoSeparadorDecimal()
    End Sub
    Private Sub btnCero_Click(sender As Object, e As EventArgs) Handles btnCero.Click
        CodigoNumeroCero()
    End Sub
    Private Sub btnUno_Click(sender As Object, e As EventArgs) Handles btnUno.Click
        CodigoNumeroUno()
    End Sub
    Private Sub btnDos_Click(sender As Object, e As EventArgs) Handles btnDos.Click
        CodigoNumeroDos()
    End Sub
    Private Sub btnTres_Click(sender As Object, e As EventArgs) Handles btnTres.Click
        CodigoNumeroTres()
    End Sub
    Private Sub btnCuatro_Click(sender As Object, e As EventArgs) Handles btnCuatro.Click
        CodigoNumeroCuatro()
    End Sub
    Private Sub btnCinco_Click(sender As Object, e As EventArgs) Handles btnCinco.Click
        CodigoNumeroCinco()
    End Sub
    Private Sub btnSeis_Click(sender As Object, e As EventArgs) Handles btnSeis.Click
        CodigoNumeroSeis()
    End Sub
    Private Sub btnSiete_Click(sender As Object, e As EventArgs) Handles btnSiete.Click
        CodigoNumeroSiete()
    End Sub
    Private Sub btnOcho_Click(sender As Object, e As EventArgs) Handles btnOcho.Click
        CodigoNumeroOcho()
    End Sub
    Private Sub btnNueve_Click(sender As Object, e As EventArgs) Handles btnNueve.Click
        CodigoNumeroNueve()
    End Sub
    '=========================================================================================================
    ' CODIGO SUBRUTINAS COMUNES TECLADO Y BOTONES
    '=========================================================================================================
    Sub CodigoSeparadorDecimal()
        ' Este fragmento se ejecuta si no se encuentra en modo fracción
        If bModoFraccion = False Then
            VerificaBorradoDisplay()
            ' Si no se encuentra un punto decimal presente en el numero actual
            If bPuntoDecimalPresente = False Then
                ' Concatena muestra el símbolo ",", a lo que ya se encontraba en el textbox del display
                tbDisplay.Text = tbDisplay.Text & ","
                ' Indica que hay un punto decimal
                bPuntoDecimalPresente = True
                ' Indica que no es cero el dígito
                bDigitoActualContieneCero = False
                ' Indica que no es el primero
                bPosicionDigitoActualEsElPrimero = False
                ' Se Ingresó un número o simbolo
                bResultadoEnDisplay = False
            End If
        End If
        tbDisplay.Select()
    End Sub

    Sub CodigoNumeroCero()
        ' Si el modo no es fracción, se verifica posición del dígito y se borra el display
        If bModoFraccion = False Then
            VerificaBorradoDisplay()
        End If
        ' Si el modo es fracción, no se borra el display cada vez que se ingresa
        ' un número, porque ahora el operando es una fracción que contiene varios
        ' números y barras de fracción.
        If bPosicionDigitoActualEsElPrimero = True Then
            If bDigitoActualContieneCero = True Then
                tbDisplay.Text = 0
                bPosicionDigitoActualEsElPrimero = True
            Else
                ' Concatena muestra el símbolo "Numero", a lo que ya se encontraba en el textbox del display
                tbDisplay.Text = tbDisplay.Text & 0
                bPosicionDigitoActualEsElPrimero = False
            End If
        Else
            tbDisplay.Text = tbDisplay.Text & 0
            bPosicionDigitoActualEsElPrimero = False
        End If

        bDigitoActualContieneCero = True
        ' Ya hay un numero en display.
        bNumeroEn_tb_Display = True
        ' NO es el resultado de una operacion
        bResultadoEnDisplay = False
        ' Numero ingresado
        bNumeroIngresado = True
        ' Devuelve el foco al cuadro de texto, caso contrario no se puede ingresar por teclado alfanumerico
        ' dado que está asociado al evento de este objeto
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroUno()
        MuestraNumeroEnDisplay(1)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroDos()
        MuestraNumeroEnDisplay(2)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroTres()
        MuestraNumeroEnDisplay(3)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroCuatro()
        MuestraNumeroEnDisplay(4)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroCinco()
        MuestraNumeroEnDisplay(5)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroSeis()
        MuestraNumeroEnDisplay(6)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroSiete()
        MuestraNumeroEnDisplay(7)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroOcho()
        MuestraNumeroEnDisplay(8)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroNueve()
        MuestraNumeroEnDisplay(9)
        tbDisplay.Select()
    End Sub


    '//////////////////////////////////////
    ' SUBRUTINAS NIVELES INFERIORES
    '//////////////////////////////////////
    Sub EjecutaOperaciones()
        Dim F1a, F1b, F1c As Integer
        Dim F2a, F2b, F2c As Integer


        If bModoFraccion = True Then
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

            'modo fracción
            'habilita almacenar el proximo numero Fraccion Dos
            If bFlagAlmacenarOperandoFraccionDos = True Then
                ' deshabilita almacenar el proximo numero como operando b 
                bFlagAlmacenarOperandoFraccionDos = False
                'modo fracción, guardar en memoria formato decimal
                AlmacenaOperandoFraccionDosEnMemoria()
            End If
            F1a = iFraccionA(0)
            F1b = iFraccionA(1)
            F1c = iFraccionA(2)

            F2a = iFraccionB(0)
            F2b = iFraccionB(1)
            F2c = iFraccionB(2)
        Else
            'modo decimal, guardar en memoria formato decimal
            If bFlagAlmacenarOperando_b = True Then
                ' deshabilita almacenar el proximo numero como operando b 
                bFlagAlmacenarOperando_b = False

                ' Almacena operando b en memoria (el número del display, una sola vez,
                ' excepto que ingrese nueva operacion
                AlmacenaOperando_bEnMemoria()
            End If
        End If


        ' de acuerdo a la variable "operación", seleccionamos que haremos con los operandos
        Select Case sOperacion

            Case "adicion"
                If bModoFraccion = True Then

                    Dim res As Double

                    F1b = F1a * F1c + F1b
                    F2b = F2a * F2c + F2b

                    F2b = (F2c * F1b) + (F1c * F2b)
                    F2c = F1c * F2c

                    ' simplificación básica

                    ' obtiene el módulo de los componentes (enteros) de la fraccion convertidos a
                    ' doble, si son divisibles, luego se toma como factor comun
                    res = CDbl(F2b) Mod CDbl(F2c)

                    ' y se calculan los componentes de la fracción simplificados
                    ' si el divisor es exacto para ambos
                    If res <> 0D Then
                        F2b = CInt(CDbl(F2b) / res)
                        F2c = CInt(CDbl(F2c) / res)
                    End If

                    'Se muestra en display
                    tbDisplay.Text = F2b & "∟" & F2c

                Else
                    If bCalcularPorcentaje = True Then
                        ' realiza la operacion de TASA DE CAMBIO (cociente entre (suma aritmetica de operandos)
                        ' y operando b y luego multiplica por cien)
                        tbDisplay.Text = (((dOperando_a + dOperando_b) / dOperando_b) * 100) & " %"
                        ' fin calculo de porcentaje
                        bCalcularPorcentaje = False
                        sOperacion = ""
                    Else
                        ' realiza la operacion de suma aritmetica
                        tbDisplay.Text = dOperando_a + dOperando_b
                    End If
                End If

            Case "sustraccion"
                If bModoFraccion = True Then

                    Dim res As Double

                    F1b = F1a * F1c + F1b
                    F2b = F2a * F2c + F2b

                    F2b = (F2c * F1b) - (F1c * F2b)
                    F2c = F1c * F2c

                    ' simplificación básica

                    ' obtiene el módulo de los componentes (enteros) de la fraccion convertidos a
                    ' doble, si son divisibles, luego se toma como factor comun
                    res = CDbl(F2b) Mod CDbl(F2c)

                    ' y se calculan los componentes de la fracción simplificados
                    ' si el divisor es exacto para ambos
                    If res <> 0D Then
                        F2b = CInt(CDbl(F2b) / res)
                        F2c = CInt(CDbl(F2c) / res)
                    End If

                    'Se muestra en display
                    tbDisplay.Text = F2b & "∟" & F2c

                Else
                    If bCalcularPorcentaje = True Then
                        ' realiza la operacion de TASA DE CAMBIO (cociente entre (suma aritmetica de operandos)
                        ' y operando b y luego multiplica por cien)
                        tbDisplay.Text = (((dOperando_a - dOperando_b) / dOperando_b) * 100) & " %"
                        ' fin calculo de porcentaje
                        bCalcularPorcentaje = False
                        sOperacion = ""
                    Else
                        ' realiza la operacion de resta aritmetica
                        tbDisplay.Text = dOperando_a - dOperando_b
                    End If
                End If

            Case "division"
                ' Si el segundo operando es cero, muestra "Error" en display
                If dOperando_b = 0 Then
                    tbDisplay.Text = msgErrMat
                Else
                    If bCalcularPorcentaje = True Then
                        ' realiza la operacion de RELACION (division aritmetica y luego multiplica por cien)
                        tbDisplay.Text = ((dOperando_a / dOperando_b) * 100) & " %"
                        ' fin calculo de porcentaje
                        bCalcularPorcentaje = False
                        sOperacion = ""
                    Else
                        ' realiza la operacion de division aritmetica
                        tbDisplay.Text = dOperando_a / dOperando_b
                    End If
                End If

            Case "multiplicacion"
                If bCalcularPorcentaje = True Then
                    ' realiza la operacion de PORCENTAJE (producto aritmetico y luego divide por cien)
                    tbDisplay.Text = ((dOperando_a * dOperando_b) / 100) & " %"
                    ' fin calculo de porcentaje
                    bCalcularPorcentaje = False
                    sOperacion = ""
                Else
                    ' realiza la operacion de producto aritmetico
                    tbDisplay.Text = dOperando_a * dOperando_b
                End If

            Case "potenciacion"
                If bCalcularPorcentaje = True Then
                    tbDisplay.Text = msgErrStx
                    bCalcularPorcentaje = False
                    sOperacion = ""
                Else
                    ' realiza la operacion de potenciación
                    tbDisplay.Text = Math.Pow(dOperando_a, dOperando_b)
                End If

            Case "RaizCuadrada"
                If bCalcularPorcentaje = True Then
                    tbDisplay.Text = "ERROR de sintaxis"
                    bCalcularPorcentaje = False
                    sOperacion = ""
                Else
                    ' realiza la operacion de potenciación
                    tbDisplay.Text = Math.Sqrt(dOperando_a)
                    sOperacion = ""
                End If
        End Select

        If bModoFraccion = True Then
            ' ya realizada la operacion, no se guarda en memoria
            bOperacionEnCurso = False

            ' Inicializa variables de operaciones con fracciones
            bFlagAlmacenarOperandoFraccionDos = False
            bBarraFraccionPresente = False
            bModoFraccion = False
            bIngresoComponenteFraccion = False
            bIngresoOperandoFraccionUno = False
            bIngresoOperandoFraccionDos = False
            iContadorBarraFraccional = 0
            bFraccionErrorIngreso = False

            ' reinicializa estados de primer dígito y punto decimal presente
            Reinicia_tb_Display()
            ' pasa el foco al elemento tb_Display (barra parpadeando)
            tbDisplay.Select()
        Else
            ' almacena el resultado como el operando a
            AlmacenaOperando_aEnMemoria()
            ' ya realizada la operacion, no se guarda en memoria
            bOperacionEnCurso = False
            ' reinicializa estados de primer dígito y punto decimal presente
            Reinicia_tb_Display()
            ' pasa el foco al elemento tb_Display (barra parpadeando)
            tbDisplay.Select()

        End If
    End Sub

    Sub VerificaBorradoDisplay()
        ' Si el primer digito a ingresar es cero, no se concatena 
        ' si no es el primero, se concatena (el primero puede ser un punto)
        If bFlagBorraDisplayEnProximoIngreso = True Then
            ' reinicializa estados de primer dígito y punto decimal presente
            Reinicia_tb_Display()
            MuestraCeroEnDisplay()
            bFlagBorraDisplayEnProximoIngreso = False
        End If
    End Sub
    Sub MuestraNumeroEnDisplay(Numero As Integer)
        ' Si el modo no es fracción, se verifica posición del dígito y se borra el display
        If bModoFraccion = False Then
            VerificaBorradoDisplay()
        End If
        ' Si el modo es fracción, no se borra el display cada vez que se ingresa
        ' un número, porque ahora el operando es una fracción que contiene varios
        ' números y barras de fracción.

        If bPosicionDigitoActualEsElPrimero = True Then
            If bDigitoActualContieneCero = True Then
                tbDisplay.Text = Numero
                bDigitoActualContieneCero = False
                bPosicionDigitoActualEsElPrimero = True
            Else
                ' Concatena muestra el símbolo "Numero", a lo que ya se encontraba en el textbox del display
                tbDisplay.Text = tbDisplay.Text & Numero
                bDigitoActualContieneCero = False
                bPosicionDigitoActualEsElPrimero = False
            End If
        Else
            tbDisplay.Text = tbDisplay.Text & Numero
            bDigitoActualContieneCero = False
            bPosicionDigitoActualEsElPrimero = False
        End If

        ' Ya hay un numero en display.
        bNumeroEn_tb_Display = True
        ' NO es el resultado de una operacion
        bResultadoEnDisplay = False
        ' Numero ingresado
        bNumeroIngresado = True
    End Sub
    Sub Reinicia_tb_Display()
        ' Sin punto decimal
        bPuntoDecimalPresente = False
        ' Posición digito ingresado = primero
        bPosicionDigitoActualEsElPrimero = True
    End Sub
    Sub MuestraCeroEnDisplay()
        tbDisplay.Text = 0
        bDigitoActualContieneCero = True
    End Sub
    Sub AlmacenaOperando_aEnMemoria()
        ' almacena como primer operando al número contenido en el cuadro de texto 
        ' convertido previamente de texto a número.
        dOperando_a = Convierte(tbDisplay.Text)
        ' operando en memoria
        bOperando_a_EnMemoria = True
    End Sub
    Sub AlmacenaOperando_bEnMemoria()
        ' almacena como primer operando al número contenido en el cuadro de texto 
        ' convertido previamente de texto a número.
        dOperando_b = Convierte(tbDisplay.Text)
        ' operando en memoria
        bOperando_b_EnMemoria = True
    End Sub

    Sub AlmacenaOperandoFraccionUnoEnMemoria()

        Dim iLargoCadena As Integer
        Dim sNumeroFormatoTexto As String

        Dim iPosBarraUno As Integer
        Dim iPosBarraDos As Integer



        ' Para saber cuantas barras de fraccion, leemos iContadorBarraFraccional
        ' y luego podemos conocer cuantos números y tipo de fraccion tenemos

        iLargoCadena = Len(tbDisplay.Text)

        ' En este segmento obtengo las posiciones de los separadores de fraccion
        ' primera posición
        iPosBarraUno = InStr(tbDisplay.Text, "∟", CompareMethod.Text)
        ' Segunda posición (si no hay, indicará cero)
        iPosBarraDos = InStr((iPosBarraUno + 1), tbDisplay.Text, "∟", CompareMethod.Text)
        ' En este segmento debo leer los dos o tres números de la fraccion de acuerdo 
        ' a iContadorBarraFraccional

        ' Primer número
        ' Obtiene la cadena
        sNumeroFormatoTexto = Mid(tbDisplay.Text, 1, (iPosBarraUno - 1))
        ' Y la convierte a número
        iFraccionA(0) = Val(sNumeroFormatoTexto)

        Select Case iContadorBarraFraccional
            Case 1
                ' Acomoda la fracción para que ocupe b∟c, y a = 0
                iFraccionA(1) = iFraccionA(0)
                iFraccionA(0) = 0
                ' Si hay una sola barra, es de dos números
                sNumeroFormatoTexto = Mid(tbDisplay.Text, iPosBarraUno + 1, iLargoCadena)
                ' convierte a número
                iFraccionA(2) = Val(sNumeroFormatoTexto)

            Case 2
                ' Si hay dos barras, es de tres números
                sNumeroFormatoTexto = Mid(tbDisplay.Text, iPosBarraUno + 1, (iPosBarraDos - iPosBarraUno) - 1)
                ' convierte a número
                iFraccionA(1) = Val(sNumeroFormatoTexto)
                sNumeroFormatoTexto = Mid(tbDisplay.Text, iPosBarraDos + 1, iLargoCadena)
                ' convierte a número
                iFraccionA(2) = Val(sNumeroFormatoTexto)
            Case Else
                'mensaje error, no debería haber incongruencias.
        End Select

        'finalmente, chequea que el denominador sea distinto de cero
        If iFraccionB(2) = 0 Then
            bFraccionErrorIngreso = True
        End If

    End Sub

    Sub AlmacenaOperandoFraccionDosEnMemoria()

        Dim iLargoCadena As Integer
        Dim sNumeroFormatoTexto As String

        Dim iPosBarraUno As Integer
        Dim iPosBarraDos As Integer



        ' Para saber cuantas barras de fraccion, leemos iContadorBarraFraccional
        ' y luego podemos conocer cuantos números y tipo de fraccion tenemos

        iLargoCadena = Len(tbDisplay.Text)

        ' En este segmento obtengo las posiciones de los separadores de fraccion
        ' primera posición
        iPosBarraUno = InStr(tbDisplay.Text, "∟", CompareMethod.Text)
        ' Segunda posición (si no hay, indicará cero)
        iPosBarraDos = InStr((iPosBarraUno + 1), tbDisplay.Text, "∟", CompareMethod.Text)
        ' En este segmento debo leer los dos o tres números de la fraccion de acuerdo 
        ' a iContadorBarraFraccional

        ' Primer número
        ' Obtiene la cadena
        sNumeroFormatoTexto = Mid(tbDisplay.Text, 1, (iPosBarraUno - 1))
        ' Y la convierte a número
        iFraccionB(0) = Val(sNumeroFormatoTexto)

        Select Case iContadorBarraFraccional
            Case 1
                ' Si hay una sola barra, es de dos números
                ' Acomoda la fracción para que ocupe b∟c, y a = 0
                iFraccionB(1) = iFraccionB(0)
                iFraccionB(0) = 0
                sNumeroFormatoTexto = Mid(tbDisplay.Text, iPosBarraUno + 1, iLargoCadena)
                ' convierte a número
                iFraccionB(2) = Val(sNumeroFormatoTexto)
            Case 2
                ' Si hay dos barras, es de tres números
                sNumeroFormatoTexto = Mid(tbDisplay.Text, iPosBarraUno + 1, (iPosBarraDos - iPosBarraUno) - 1)
                ' convierte a número
                iFraccionB(1) = Val(sNumeroFormatoTexto)
                sNumeroFormatoTexto = Mid(tbDisplay.Text, iPosBarraDos + 1, iLargoCadena)
                ' convierte a número
                iFraccionB(2) = Val(sNumeroFormatoTexto)
            Case Else
                'mensaje error, no debería haber incongruencias.
        End Select
        'finalmente, chequea que el denominador sea distinto de cero
        If iFraccionB(2) = 0 Then
            bFraccionErrorIngreso = True
        End If

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
        bFlagAlmacenarOperando_b = False
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
        '
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
    Private Function Convierte(ByVal Texto As String) As Double
        ' Se busca en la cadena recibida como parámetro el caracter coma ","
        ' y se lo reemplaza por el caracter punto "." para luego convertir
        ' el número recibido en formato texto a formato valor (numérico)
        Convierte = Val(Replace(Texto, ",", "."))
    End Function

End Class