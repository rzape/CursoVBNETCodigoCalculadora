Module SubSM_Calculadora
    Sub SM_Calculadora()



        Select Case iSMCalculadora

            Case S000
                ' En este estado se espera hasta que se ingresa un boton de operacion
                ' luego se carga el numero en display al operando A, se carga la operacion
                ' y se salta al estado 100
                If bNuevaOperacion = True Then
                    'ingresa el operando A nuevo
                    Form1.IngresoOperandoA()
                    Select Case sBoton
                        Case "boton_+"
                            ' define la operación a realizar como suma
                            sOperacion = "adicion"
                        Case "boton_-"
                            ' define la operación a realizar como suma
                            sOperacion = "sustraccion"
                        Case "boton_x"
                            ' define la operación a realizar como multiplicacion
                            sOperacion = "multiplicacion"
                        Case Else
                            'sOperacion = ""
                    End Select
                    'salta al proximo estado
                    iSMCalculadora = S100
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' Borra display en proximo ingreso de un numero
                    bFlagBorraDisplayEnProximoIngreso = True
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If

            Case S100
                ' En este estado se puede cambiar de operacion.
                ' Se espera hasta que se ingresa un número, luego ya no se
                ' puede cambiar la operacion
                If bNuevaOperacion = True Then
                    'si se cambia el boton, cambia operación, sino quedará la misma
                    Select Case sBoton
                        Case "boton_+"
                            ' define la operación a realizar como suma
                            sOperacion = "adicion"
                        Case "boton_-"
                            ' define la operación a realizar como suma
                            sOperacion = "sustraccion"
                        Case "boton_x"
                            ' define la operación a realizar como multiplicacion
                            sOperacion = "multiplicacion"
                        Case Else
                            'sOperacion = ""
                    End Select
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                ElseIf bSeparadorDecimalIngresado = True Or bNumeroIngresado = True Then
                    ' Espera a que se presione el botón de "=" una vez finalizado
                    ' el ingreso del operando B
                    Select Case sOperacion
                        Case "adicion"
                            'salta al proximo estado
                            iSMCalculadora = S200
                        Case "sustraccion"
                            'salta al proximo estado
                            iSMCalculadora = S202
                        Case "multiplicacion"
                            'salta al proximo estado
                            iSMCalculadora = S201
                        Case Else
                            'sOperacion = ""
                    End Select
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If

            Case S200
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_="
                            ' Ingresa operando
                            Form1.NivelaTipoDeOperandos()
                            ' Realiza la operación suma
                            OperacionSuma()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S300
                        Case "boton_%"
                            ' Ingresa operando
                            Form1.NivelaTipoDeOperandos()
                            ' Realiza la operación suma
                            OperacionTasaDeCambio1()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S400
                        Case Else
                            'sOperacion = ""
                    End Select
                Else
                    ' No ocurre nada
                End If
                ' borra nueva operacion
                bNuevaOperacion = False
                ' borra el boton llamador
                sBoton = ""
                ' pasa el foco al elemento tb_Display (barra parpadeando)
                Form1.tbDisplay.Select()


            Case S201
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_="
                            ' Ingresa operando
                            Form1.NivelaTipoDeOperandos()
                            ' Realiza la operación suma
                            OperacionMultiplicacion()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S301
                        Case "boton_%"
                            ' Ingresa operando
                            Form1.NivelaTipoDeOperandos()
                            ' Realiza la operación suma
                            OPeracionPorcentaje()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S302
                        Case Else
                            'sOperacion = ""
                    End Select
                Else
                    ' No ocurre nada
                End If
                ' borra nueva operacion
                bNuevaOperacion = False
                ' borra el boton llamador
                sBoton = ""
                ' pasa el foco al elemento tb_Display (barra parpadeando)
                Form1.tbDisplay.Select()

            Case S202
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_="
                            ' Ingresa operando
                            Form1.NivelaTipoDeOperandos()
                            ' Realiza la operación suma
                            OperacionResta()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S303
                        Case "boton_%"
                            ' Ingresa operando
                            Form1.NivelaTipoDeOperandos()
                            ' Realiza la operación tasa de cambio2
                            OperacionTasaDeCambio2()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S402
                        Case Else
                            'sOperacion = ""
                    End Select
                Else
                    ' No ocurre nada
                End If
                ' borra nueva operacion
                bNuevaOperacion = False
                ' borra el boton llamador
                sBoton = ""
                ' pasa el foco al elemento tb_Display (barra parpadeando)
                Form1.tbDisplay.Select()


            Case S300
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_+"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "adicion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_%"
                            'Los operandos son A y B ya ingresados
                            ' Realiza la operación Tasa de Cambio 1
                            ' ### BUG
                            ' #En calculadora repite una vez más la operación
                            ' #antes de efectuar la tasa de cambio
                            ' ejemplo
                            ' codigo: 1+2=3 +2 (5) +2 (7) % (350) 
                            ' casio : 1+2=3 ans+2 (5) ans+2 (7) (ans+2)% (450) 
                            OperacionTasaDeCambio1()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S400
                        Case "boton_="
                            ' ###  BUG
                            ' #Repite la operación, pero se debe verificar cual es la operacion
                            ' #anterior y además si es cálculo de porcentaje NO se debe
                            ' #repetir.
                            'Copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' Operando B es el mismo
                            ' Realiza la operación suma
                            OperacionSuma()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S300
                        Case Else
                            'nada
                    End Select
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' Borra display en proximo ingreso de un numero
                    bFlagBorraDisplayEnProximoIngreso = True
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                ElseIf bSeparadorDecimalIngresado = True Or bNumeroIngresado = True Then
                    ' Comienza nuevamente desde el principio
                    iSMCalculadora = S000
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If

            Case S301
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_x"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "multiplicacion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_%"
                            'Los operandos son A y B ya ingresados
                            ' Realiza la operación Tasa de Cambio 1
                            ' ### BUG
                            ' #En calculadora repite una vez más la operación
                            ' #antes de efectuar la tasa de cambio
                            ' ejemplo
                            ' codigo: 1+2=3 +2 (5) +2 (7) % (350) 
                            ' casio : 1+2=3 ans+2 (5) ans+2 (7) (ans+2)% (450) 
                            OPeracionPorcentaje()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S400
                        Case "boton_="
                            ' ###  BUG
                            ' #Repite la operación, pero se debe verificar cual es la operacion
                            ' #anterior y además si es cálculo de porcentaje NO se debe
                            ' #repetir.
                            'Copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' Operando B es el mismo
                            ' Realiza la operación suma
                            OperacionMultiplicacion()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S301
                        Case Else
                            'nada
                    End Select
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' Borra display en proximo ingreso de un numero
                    bFlagBorraDisplayEnProximoIngreso = True
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                ElseIf bSeparadorDecimalIngresado = True Or bNumeroIngresado = True Then
                    ' Comienza nuevamente desde el principio
                    iSMCalculadora = S000
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If

            Case S302
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_+"
                            'NOTA, en esta operación se debe tomar el resultado como operador A
                            'cuando se repite la operacion (en Casio fx82, 1*2=2 ANS*3 % 0.06 + (genera OPA=0.06, OPB=3)
                            'caso contrario es el operador ingresado. Modificar al utilizar las dos filas en el txtbox
                            ' define la operación a realizar como Premium (porcentaje)
                            OperacionPremium()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S401
                        Case "boton_-"
                            'NOTA, en esta operación se debe tomar el resultado como operador A
                            'cuando se repite la operacion (en Casio fx82, 1*2=2 ANS*3 % 0.06 + (genera OPA=0.06, OPB=3)
                            'caso contrario es el operador ingresado. Modificar al utilizar las dos filas en el txtbox
                            ' define la operación a realizar como Premium (porcentaje)
                            OperacionDescuento()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S403
                        Case "boton_x"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "multiplicacion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_="
                            'nada
                            iSMCalculadora = S302
                        Case Else
                            'nada
                    End Select
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' Borra display en proximo ingreso de un numero
                    bFlagBorraDisplayEnProximoIngreso = True
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                ElseIf bSeparadorDecimalIngresado = True Or bNumeroIngresado = True Then
                    ' Comienza nuevamente desde el principio
                    iSMCalculadora = S000
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If

            Case S303
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_-"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "sustraccion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_%"
                            'Los operandos son A y B ya ingresados
                            ' Realiza la operación Tasa de Cambio 1
                            ' ### BUG
                            ' #En calculadora repite una vez más la operación
                            ' #antes de efectuar la tasa de cambio
                            ' ejemplo
                            ' codigo: 1+2=3 +2 (5) +2 (7) % (350) 
                            ' casio : 1+2=3 ans+2 (5) ans+2 (7) (ans+2)% (450) 
                            OperacionTasaDeCambio2()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S402
                        Case "boton_="
                            ' ###  BUG
                            ' #Repite la operación, pero se debe verificar cual es la operacion
                            ' #anterior y además si es cálculo de porcentaje NO se debe
                            ' #repetir.
                            'Copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' Operando B es el mismo
                            ' Realiza la operación suma
                            OperacionSuma()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S303
                        Case Else
                            'nada
                    End Select
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' Borra display en proximo ingreso de un numero
                    bFlagBorraDisplayEnProximoIngreso = True
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                ElseIf bSeparadorDecimalIngresado = True Or bNumeroIngresado = True Then
                    ' Comienza nuevamente desde el principio
                    iSMCalculadora = S000
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If
            Case S400
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_+"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "adicion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_="
                            ' Copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' Operando B es el mismo
                            ' Realiza la operación suma
                            OperacionSuma()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S300
                        Case Else
                            'nada
                    End Select
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' Borra display en proximo ingreso de un numero
                    bFlagBorraDisplayEnProximoIngreso = True
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                ElseIf bSeparadorDecimalIngresado = True Or bNumeroIngresado = True Then
                    ' Comienza nuevamente desde el principio
                    iSMCalculadora = S000
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If

            Case S401
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_+"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "adicion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_x"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "multiplicacion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_="
                            'No se efectua otra operacion, queda esperando
                            'salta al proximo estado
                            iSMCalculadora = S401
                        Case Else
                            'nada
                    End Select
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' Borra display en proximo ingreso de un numero
                    bFlagBorraDisplayEnProximoIngreso = True
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                ElseIf bSeparadorDecimalIngresado = True Or bNumeroIngresado = True Then
                    ' Comienza nuevamente desde el principio
                    iSMCalculadora = S000
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If

            Case S402
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_-"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "sustraccion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_="
                            ' Copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' Operando B es el mismo
                            ' Realiza la operación suma
                            OperacionResta()
                            ' Se generó un resultado
                            bResultadoEnDisplay = True
                            'salta al proximo estado
                            iSMCalculadora = S303
                        Case Else
                            'nada
                    End Select
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' Borra display en proximo ingreso de un numero
                    bFlagBorraDisplayEnProximoIngreso = True
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                ElseIf bSeparadorDecimalIngresado = True Or bNumeroIngresado = True Then
                    ' Comienza nuevamente desde el principio
                    iSMCalculadora = S000
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If

            Case S403
                If bNuevaOperacion = True Then
                    Select Case sBoton
                        Case "boton_-"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "sustraccion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_x"
                            'copia el valor del operando resultado en el operando A
                            Form1.CopiaResultadoEnOpA()
                            ' define la operación a realizar como suma
                            sOperacion = "multiplicacion"
                            ' Salta al proximo estado  
                            ' a ingresar operando B
                            iSMCalculadora = S100
                        Case "boton_/"
                            'Falta agregarlo
                        Case "boton_="
                            'No se efectua otra operacion, queda esperando
                            'salta al proximo estado
                            iSMCalculadora = S403
                        Case Else
                            'nada
                    End Select
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' Borra display en proximo ingreso de un numero
                    bFlagBorraDisplayEnProximoIngreso = True
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                    ' borra nueva operacion
                    bNuevaOperacion = False
                    ' borra el boton llamador
                    sBoton = ""
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                ElseIf bSeparadorDecimalIngresado = True Or bNumeroIngresado = True Then
                    ' Comienza nuevamente desde el principio
                    iSMCalculadora = S000
                    ' pasa el foco al elemento tb_Display (barra parpadeando)
                    Form1.tbDisplay.Select()
                End If

        End Select


        Form1.FinalizaOperacionEjecucion()
        'ESTE CODIGO SE EJECUTA CON EL BOTON =
        'If bNuevaOperacion = True Then
        '    'operación nueva, ingresar operando b
        '    'y verificar si coinciden los tipos entero o fraccion;
        '    'En caso de ser alguno fraccional, convierte el otro al 
        '    'mismo tipo
        '    NivelaTipoDeOperandos()
        'Else
        '    CopiaResultadoEnOpA()
        'End If

        ' ESTE CODIGO LEIA O COPIABA EL OPERANDO A DEPENDIENDO SI HABIA RESULTADO
        'Si hay un resultado en display, se copia el mismo en operando A
        'If bResultadoEnDisplay = True Then
        '    CopiaResultadoEnOpA_REV_2()
        'Else
        '    'caso contrario ingresa el operando A nuevo
        '    IngresoOperandoA()
        'End If
    End Sub
End Module
