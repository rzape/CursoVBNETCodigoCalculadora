Module OperacionesMatematicas
    Sub OperacionSuma()
        If bOperacionFraccional = True Then



            iFraccionA(1) = iFraccionA(0) * iFraccionA(2) + iFraccionA(1)
            iFraccionB(1) = iFraccionB(0) * iFraccionB(2) + iFraccionB(1)

            iFraccionR(1) = (iFraccionB(2) * iFraccionA(1)) + (iFraccionA(2) * iFraccionB(1))
            iFraccionR(2) = iFraccionA(2) * iFraccionB(2)

            'VERIFICAR SI EL DENOMINADOR NO ES CERO
            ' simplificación básica

            'Se muestra en display
            Form1.tbDisplay.Text = iFraccionR(1) & "∟" & iFraccionR(2)
            bResultadoTipoFraccion = True
        Else
            ' realiza la operacion de suma aritmetica
            dOperando_r = dOperando_a + dOperando_b
            'visualiza en display
            Form1.tbDisplay.Text = dOperando_r
            bResultadoTipoFraccion = False
        End If
    End Sub
    Sub OperacionResta()
        If bOperacionFraccional = True Then
            'convierte la fracción impropia en fracción propia
            iFraccionA(1) = iFraccionA(0) * iFraccionA(2) + iFraccionA(1)
            iFraccionB(1) = iFraccionB(0) * iFraccionB(2) + iFraccionB(1)
            'Calcula la resta de fracciones
            iFraccionR(1) = (iFraccionB(2) * iFraccionA(1)) - (iFraccionA(2) * iFraccionB(1))
            iFraccionR(2) = iFraccionA(2) * iFraccionB(2)

            'VERIFICAR SI EL DENOMINADOR NO ES CERO
            ' simplificación básica

            'Se muestra en display
            Form1.tbDisplay.Text = iFraccionR(1) & "∟" & iFraccionR(2)
            bResultadoTipoFraccion = True
        Else
            ' realiza la operacion de resta aritmetica
            dOperando_r = dOperando_a - dOperando_b
            'visualiza en display
            Form1.tbDisplay.Text = dOperando_r
            bResultadoTipoFraccion = False
        End If


    End Sub
    Sub OperacionMultiplicacion()
        If bOperacionFraccional = True Then

            'convierte la fracción impropia en fracción propia
            iFraccionA(1) = iFraccionA(0) * iFraccionA(2) + iFraccionA(1)
            iFraccionB(1) = iFraccionB(0) * iFraccionB(2) + iFraccionB(1)

            'Calcula el producto de fracciones
            iFraccionR(1) = iFraccionA(1) * iFraccionB(1)
            iFraccionR(2) = iFraccionA(2) * iFraccionB(2)

            'VERIFICAR SI EL DENOMINADOR NO ES CERO
            ' simplificación básica

            'Se muestra en display
            Form1.tbDisplay.Text = iFraccionR(1) & "∟" & iFraccionR(2)
            bResultadoTipoFraccion = True
        Else
            ' realiza la operacion de producto aritmetico
            dOperando_r = dOperando_a * dOperando_b
            Form1.tbDisplay.Text = dOperando_r
        End If
    End Sub
    Sub OPeracionPorcentaje()
        ' realiza la operacion de PORCENTAJE (producto aritmetico y luego divide por cien)
        dOperando_r = ((dOperando_a * dOperando_b) / 100)
        Form1.tbDisplay.Text = dOperando_r & " %"
    End Sub
    Sub OperacionPremium()
        ' Que es A incrementado en B %   ??
        dOperando_r = dOperando_a + ((dOperando_a * dOperando_b) / 100)
        Form1.tbDisplay.Text = dOperando_r & "%"
    End Sub
    Sub OperacionDescuento()
        ' Que es A reducido en B %   ??
        dOperando_r = dOperando_a - ((dOperando_a * dOperando_b) / 100)
        Form1.tbDisplay.Text = dOperando_r & "%"
    End Sub
    Sub OperacionTasaDeCambio1()
        'Esta operación sólo se puede realizar
        'si antes se calculó el porcentaje
        dOperando_r = ((dOperando_a + dOperando_b) / dOperando_b) * 100
        Form1.tbDisplay.Text = dOperando_r & "%"
    End Sub
    Sub OperacionTasaDeCambio2()
        'Esta operación sólo se puede realizar
        'si antes se calculó el porcentaje
        dOperando_r = ((dOperando_a - dOperando_b) / dOperando_b) * 100
        Form1.tbDisplay.Text = dOperando_r & "%"
    End Sub

End Module
