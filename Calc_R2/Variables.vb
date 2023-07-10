Module Variables
    '//////////////////////////////////////
    ' Variables
    '//////////////////////////////////////
    Public bPuntoDecimalPresente As Boolean
    Public bPosicionDigitoActualEsElPrimero As Boolean
    Public bDigitoActualContieneCero As Boolean
    Public bNumeroEn_tb_Display As Boolean
    Public bOperando_a_EnMemoria As Boolean
    Public bOperando_b_EnMemoria As Boolean
    Public bFlagBorraDisplayEnProximoIngreso As Boolean
    Public bOperacionEnCurso As Boolean
    Public bFlagAlmacenarOperando_b As Boolean
    Public bCalcularPorcentaje As Boolean
    Public bCalcPorcentajeExtension As Boolean
    Public bModoFraccion As Boolean
    Public bIngresoComponenteFraccion As Boolean
    Public bIngresoOperandoFraccionUno As Boolean
    Public bIngresoOperandoFraccionDos As Boolean
    Public bFraccionErrorIngreso As Boolean
    Public bResultadoEnDisplay As Boolean
    Public bNumeroIngresado As Boolean
    Public bBarraFraccionPresente As Boolean
    Public bFlagAlmacenarOperandoFraccionDos As Boolean
    ' Variable para definir si el prim er operando es del tipo fracción o entero
    Public bPriOperTipoFraccion As Boolean
    ' Variable para definir si el segundo operando es del tipo fracción o entero
    Public bSegOperTipoFraccion As Boolean
    ' Variable para definir si el resultado es del tipo fracción o entero
    Public bResultadoTipoFraccion As Boolean
    Public bOperando_aCopiadoDeOperando_r As Boolean
    'tipo de modo de operación a realizar
    Public bOperacionFraccional As Boolean
    Public bPorcentajeEnDisplay As Boolean
    Public bTasaDeCambioEnDisplay As Boolean
    Public bSeparadorDecimalDeshabilitado As Boolean
    Public bOperacionFraccionalPrevia As Boolean
    Public bNuevaOperacion As Boolean
    Public bSeparadorDecimalIngresado As Boolean


    Public iContadorBarraFraccional As Integer
    Public iFraccionA(3) As Integer
    Public iFraccionB(3) As Integer
    Public iFraccionR(3) As Integer
    Public F1a, F1b, F1c As Integer
    Public F2a, F2b, F2c As Integer
    Public FRa, FRb, FRc As Integer
    Public iSMCalculadora As Integer


    Public dOperando_a As Double
    Public dOperando_b As Double
    Public dOperando_r As Double

    Public sBoton As String
    Public sOperacion As String
    Public sOperacionPrevia As String

    '//////////////////////////////////////
    ' Constantes
    '//////////////////////////////////////
    Public Const S000 = 0
    Public Const S100 = 100
    Public Const S101 = 101
    Public Const S200 = 200
    Public Const S201 = 201
    Public Const S202 = 202
    Public Const S300 = 300
    Public Const S301 = 301
    Public Const S302 = 302
    Public Const S303 = 303
    Public Const S400 = 400
    Public Const S401 = 401
    Public Const S402 = 402
    Public Const S403 = 403



    Public Const msgErrMat As String = "ERROR matemático"
    Public Const msgErrStx As String = "ERROR de sintaxis"


End Module
