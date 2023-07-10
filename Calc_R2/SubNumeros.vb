Module SubNumeros
    '=========================================================================================================
    ' CODIGO SUBRUTINAS COMUNES TECLADO Y BOTONES
    '=========================================================================================================
    Sub CodigoSeparadorDecimal()
        ' Este fragmento se ejecuta si no se encuentra en modo fracción
        If bSeparadorDecimalDeshabilitado = False Then
            VerificaBorradoDisplay()
            ' Si no se encuentra un punto decimal presente en el numero actual
            If bPuntoDecimalPresente = False Then
                ' Concatena muestra el símbolo ",", a lo que ya se encontraba en el textbox del display
                Form1.tbDisplay.Text = Form1.tbDisplay.Text & ","
                ' Indica que hay un punto decimal
                bPuntoDecimalPresente = True
                ' Indica que no es cero el dígito
                bDigitoActualContieneCero = False
                ' Indica que no es el primero
                bPosicionDigitoActualEsElPrimero = False
                ' Se Ingresó un número o simbolo
                bResultadoEnDisplay = False
                ' Se ingresó separador decimal
                bSeparadorDecimalIngresado = True
            End If
        End If
        Form1.tbDisplay.Select()

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
                Form1.tbDisplay.Text = 0
                bPosicionDigitoActualEsElPrimero = True
            Else
                ' Concatena muestra el símbolo "Numero", a lo que ya se encontraba en el textbox del display
                Form1.tbDisplay.Text = Form1.tbDisplay.Text & 0
                bPosicionDigitoActualEsElPrimero = False
            End If
        Else
            Form1.tbDisplay.Text = Form1.tbDisplay.Text & 0
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
        Form1.tbDisplay.Select()
    End Sub
    Sub CodigoNumeroUno()
        MuestraNumeroEnDisplay(1)
        Form1.tbDisplay.Select()
    End Sub
    Sub CodigoNumeroDos()
        MuestraNumeroEnDisplay(2)
        Form1.tbDisplay.Select()
    End Sub
    Sub CodigoNumeroTres()
        MuestraNumeroEnDisplay(3)
        Form1.tbDisplay.Select()
    End Sub
    Sub CodigoNumeroCuatro()
        MuestraNumeroEnDisplay(4)
        Form1.tbDisplay.Select()
    End Sub
    Sub CodigoNumeroCinco()
        MuestraNumeroEnDisplay(5)
        Form1.tbDisplay.Select()
    End Sub
    Sub CodigoNumeroSeis()
        MuestraNumeroEnDisplay(6)
        Form1.tbDisplay.Select()
    End Sub
    Sub CodigoNumeroSiete()
        MuestraNumeroEnDisplay(7)
        Form1.tbDisplay.Select()
    End Sub
    Sub CodigoNumeroOcho()
        MuestraNumeroEnDisplay(8)
        Form1.tbDisplay.Select()
    End Sub
    Sub CodigoNumeroNueve()
        MuestraNumeroEnDisplay(9)
        Form1.tbDisplay.Select()
    End Sub

End Module
