Module SubDisplay

    Sub AlmacenaOperandoDisplayEnteroEnMemoria(Operando)
        ' almacena como primer operando al número contenido en el cuadro de texto 
        ' convertido previamente de texto a número.
        Dim dOpr As Double

        dOpr = Form1.Convierte(Form1.tbDisplay.Text)

        Select Case Operando
            Case "A"
                ' guarda operando en memoria
                dOperando_a = dOpr
                ' operando en memoria
                bOperando_a_EnMemoria = True
            Case "B"
                ' guarda operando en memoria
                dOperando_b = dOpr
                ' operando en memoria
                bOperando_b_EnMemoria = True
            Case Else
                ' No efectua operacion
        End Select

    End Sub
    Sub AlmacenaOperandoDisplayFraccionEnMemoria(ByRef iOperando0 As Integer, ByRef iOperando1 As Integer, ByRef iOperando2 As Integer)

        Dim iLargoCadena As Integer
        Dim sNumeroFormatoTexto As String

        Dim iPosBarraUno As Integer
        Dim iPosBarraDos As Integer



        ' Para saber cuantas barras de fraccion, leemos iContadorBarraFraccional
        ' y luego podemos conocer cuantos números y tipo de fraccion tenemos
        ' Si es distinto de cero es una fracción, sino es un entero
        If iContadorBarraFraccional <> 0 Then

            iLargoCadena = Len(Form1.tbDisplay.Text)

            ' En este segmento obtengo las posiciones de los separadores de fraccion
            ' primera posición
            iPosBarraUno = InStr(Form1.tbDisplay.Text, "∟", CompareMethod.Text)
            ' Segunda posición (si no hay, indicará cero)
            iPosBarraDos = InStr((iPosBarraUno + 1), Form1.tbDisplay.Text, "∟", CompareMethod.Text)
            ' En este segmento debo leer los dos o tres números de la fraccion de acuerdo 
            ' a iContadorBarraFraccional

            ' Primer número
            ' Obtiene la cadena
            sNumeroFormatoTexto = Mid(Form1.tbDisplay.Text, 1, (iPosBarraUno - 1))

            ' Y la convierte a número
            iOperando0 = Val(sNumeroFormatoTexto)
        Else
            ' Es decimal, lo convierte a número, luego a formato entero y luego lo trunca
            iOperando1 = Math.Truncate(CInt(Form1.Convierte(Form1.tbDisplay.Text)))
        End If


        Select Case iContadorBarraFraccional
            Case 0
                ' Es un número entero
                ' lo acomoda al formato b∟c = b∟1
                iOperando0 = 0
                iOperando2 = 1

            Case 1
                ' Si hay una sola barra, es de dos números
                ' Acomoda la fracción para que ocupe b∟c, y a = 0
                iOperando1 = iOperando0
                iOperando0 = 0
                sNumeroFormatoTexto = Mid(Form1.tbDisplay.Text, iPosBarraUno + 1, iLargoCadena)
                ' convierte a número
                iOperando2 = Val(sNumeroFormatoTexto)
            Case 2
                ' Si hay dos barras, es de tres números
                sNumeroFormatoTexto = Mid(Form1.tbDisplay.Text, iPosBarraUno + 1, (iPosBarraDos - iPosBarraUno) - 1)
                ' convierte a número
                iOperando1 = Val(sNumeroFormatoTexto)
                sNumeroFormatoTexto = Mid(Form1.tbDisplay.Text, iPosBarraDos + 1, iLargoCadena)
                ' convierte a número
                iOperando2 = Val(sNumeroFormatoTexto)
            Case Else
                'mensaje error, no debería haber incongruencias.
        End Select
        'finalmente, chequea que el denominador sea distinto de cero
        If iOperando2 = 0 Then
            bFraccionErrorIngreso = True
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
        If bResultadoEnDisplay = True Then
            Reinicia_tb_Display()
            MuestraCeroEnDisplay()
        End If
        ' Si el modo no es fracción, se verifica posición del dígito y se borra el display
        If bModoFraccion = False Then
            VerificaBorradoDisplay()
        End If
        ' Si el modo es fracción, no se borra el display cada vez que se ingresa
        ' un número, porque ahora el operando es una fracción que contiene varios
        ' números y barras de fracción.

        If bPosicionDigitoActualEsElPrimero = True Then
            If bDigitoActualContieneCero = True Then
                Form1.tbDisplay.Text = Numero
                bDigitoActualContieneCero = False
                bPosicionDigitoActualEsElPrimero = True
            Else
                ' Concatena muestra el símbolo "Numero", a lo que ya se encontraba en el textbox del display
                Form1.tbDisplay.Text = Form1.tbDisplay.Text & Numero
                bDigitoActualContieneCero = False
                bPosicionDigitoActualEsElPrimero = False
            End If
        Else
            Form1.tbDisplay.Text = Form1.tbDisplay.Text & Numero
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
        Form1.tbDisplay.Text = 0
        bDigitoActualContieneCero = True
    End Sub

End Module
