' NOTAS
' MP1: 2023-07-03  La visualización por pantalla de los valores calculados se debe
'                  trabajar en una subrutina a fin de truncar o limitar los valores
'                  a visualizar.
Public Class Form1
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
    Dim bCalcularFraccion As Boolean



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
        If bOperacionEnCurso = False Then
            ' Si no se ingreso un numero nuevo (entonces en display esta el resultado
            ' de una operación anterior)
            AlmacenaOperando_aEnMemoria()
            ' Avisa que hay una operacion en curso
            bOperacionEnCurso = True
        End If
        ' define la operación a realizar como suma
        sOperacion = "adicion"
        ' habilita almacenar el proximo numero como operando b 
        bFlagAlmacenarOperando_b = True
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
            ' Si no se ingreso un numero nuevo (entonces en display esta el resultado
            ' de una operación anterior)
            AlmacenaOperando_aEnMemoria()
            ' Avisa que hay una operacion en curso
            bOperacionEnCurso = True
        End If
        ' define la operación a realizar como resta
        sOperacion = "sustraccion"
        ' borra el display
        bFlagBorraDisplayEnProximoIngreso = True
        ' habilita almacenar el proximo numero como operando b 
        bFlagAlmacenarOperando_b = True
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
        ' define el modo de calculo y presentación
        bCalcularFraccion = True
        ' habilita almacenar el proximo numero como operando b 
        bFlagAlmacenarOperando_b = True
        ' Borra display en proximo ingreso de un numero
        bFlagBorraDisplayEnProximoIngreso = True
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
        '
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
        End If
        tbDisplay.Select()
    End Sub

    Sub CodigoNumeroCero()
        '
        VerificaBorradoDisplay()
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
        'Ya hay un numero en display.
        bNumeroEn_tb_Display = True
        'Ya hay un numero en display.
        bNumeroEn_tb_Display = True
        ' Devuelve el foco al cuadro de texto, caso contrario no se puede ingresar por teclado alfanumerico
        ' dado que está asociado al evento de este objeto
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroUno()
        VerificaBorradoDisplay()
        MuestraNumeroEnDisplay(1)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroDos()
        VerificaBorradoDisplay()
        MuestraNumeroEnDisplay(2)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroTres()
        VerificaBorradoDisplay()
        MuestraNumeroEnDisplay(3)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroCuatro()
        VerificaBorradoDisplay()
        MuestraNumeroEnDisplay(4)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroCinco()
        VerificaBorradoDisplay()
        MuestraNumeroEnDisplay(5)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroSeis()
        VerificaBorradoDisplay()
        MuestraNumeroEnDisplay(6)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroSiete()
        VerificaBorradoDisplay()
        MuestraNumeroEnDisplay(7)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroOcho()
        VerificaBorradoDisplay()
        MuestraNumeroEnDisplay(8)
        tbDisplay.Select()
    End Sub
    Sub CodigoNumeroNueve()
        VerificaBorradoDisplay()
        MuestraNumeroEnDisplay(9)
        tbDisplay.Select()
    End Sub


    '//////////////////////////////////////
    ' SUBRUTINAS NIVELES INFERIORES
    '//////////////////////////////////////
    Sub EjecutaOperaciones()
        If bFlagAlmacenarOperando_b = True Then
            ' deshabilita almacenar el proximo numero como operando b 
            bFlagAlmacenarOperando_b = False
            ' Almacena operando b en memoria (el número del display, una sola vez,
            ' excepto que ingrese nueva operacion
            AlmacenaOperando_bEnMemoria()
        End If

        ' de acuerdo a la variable "operación", seleccionamos que haremos con los operandos
        Select Case sOperacion

            Case "adicion"
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

            Case "sustraccion"
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

            Case "division"
                ' Si el segundo operando es cero, muestra "Error" en display
                If dOperando_b = 0 Then
                    tbDisplay.Text = "Error de division por cero"
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
                    tbDisplay.Text = "Función no válida"
                    bCalcularPorcentaje = False
                    sOperacion = ""
                Else
                    ' realiza la operacion de potenciación
                    tbDisplay.Text = Math.Pow(dOperando_a, dOperando_b)
                End If

            Case "RaizCuadrada"
                If bCalcularPorcentaje = True Then
                    tbDisplay.Text = "Función no válida"
                    bCalcularPorcentaje = False
                    sOperacion = ""
                Else
                    ' realiza la operacion de potenciación
                    tbDisplay.Text = Math.Sqrt(dOperando_a)
                    sOperacion = ""
                End If
        End Select

        ' almacena el resultado como el operando a
        AlmacenaOperando_aEnMemoria()
        ' ya realizada la operacion, no se guarda en memoria
        bOperacionEnCurso = False
        ' reinicializa estados de primer dígito y punto decimal presente
        Reinicia_tb_Display()
        ' pasa el foco al elemento tb_Display (barra parpadeando)
        tbDisplay.Select()
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

    Sub BorraMemoriaOperandos()
        ' almacena "0" en memoria a
        dOperando_a = Convierte(0)
        ' idem memoria b
        dOperando_b = dOperando_a
        ' y borra flags de memoria ocupada
        bOperando_a_EnMemoria = False
        bOperando_b_EnMemoria = False
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
