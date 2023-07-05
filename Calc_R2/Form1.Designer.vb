<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tbDisplay = New System.Windows.Forms.TextBox()
        Me.btnPorcentaje = New System.Windows.Forms.Button()
        Me.btnSiete = New System.Windows.Forms.Button()
        Me.btnOcho = New System.Windows.Forms.Button()
        Me.btnSumar = New System.Windows.Forms.Button()
        Me.btnMultiplicar = New System.Windows.Forms.Button()
        Me.btnNueve = New System.Windows.Forms.Button()
        Me.btnResultado = New System.Windows.Forms.Button()
        Me.btnDividir = New System.Windows.Forms.Button()
        Me.btnSeis = New System.Windows.Forms.Button()
        Me.btnCinco = New System.Windows.Forms.Button()
        Me.btnCuatro = New System.Windows.Forms.Button()
        Me.btnFraccion = New System.Windows.Forms.Button()
        Me.btnClearOp = New System.Windows.Forms.Button()
        Me.btnClearEntry = New System.Windows.Forms.Button()
        Me.btnSeparadorDecimal = New System.Windows.Forms.Button()
        Me.btnCero = New System.Windows.Forms.Button()
        Me.btnRaizCuadrada = New System.Windows.Forms.Button()
        Me.btnRestar = New System.Windows.Forms.Button()
        Me.btnTres = New System.Windows.Forms.Button()
        Me.btnDos = New System.Windows.Forms.Button()
        Me.btnUno = New System.Windows.Forms.Button()
        Me.btnPotenciar = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'tbDisplay
        '
        Me.tbDisplay.Font = New System.Drawing.Font("Consolas", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbDisplay.Location = New System.Drawing.Point(37, 29)
        Me.tbDisplay.Multiline = True
        Me.tbDisplay.Name = "tbDisplay"
        Me.tbDisplay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbDisplay.Size = New System.Drawing.Size(330, 71)
        Me.tbDisplay.TabIndex = 0
        Me.tbDisplay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnPorcentaje
        '
        Me.btnPorcentaje.Location = New System.Drawing.Point(37, 131)
        Me.btnPorcentaje.Name = "btnPorcentaje"
        Me.btnPorcentaje.Size = New System.Drawing.Size(50, 46)
        Me.btnPorcentaje.TabIndex = 1
        Me.btnPorcentaje.Text = "%"
        Me.btnPorcentaje.UseVisualStyleBackColor = True
        '
        'btnSiete
        '
        Me.btnSiete.Location = New System.Drawing.Point(93, 131)
        Me.btnSiete.Name = "btnSiete"
        Me.btnSiete.Size = New System.Drawing.Size(50, 46)
        Me.btnSiete.TabIndex = 2
        Me.btnSiete.Text = "7"
        Me.btnSiete.UseVisualStyleBackColor = True
        '
        'btnOcho
        '
        Me.btnOcho.Location = New System.Drawing.Point(149, 131)
        Me.btnOcho.Name = "btnOcho"
        Me.btnOcho.Size = New System.Drawing.Size(50, 46)
        Me.btnOcho.TabIndex = 3
        Me.btnOcho.Text = "8"
        Me.btnOcho.UseVisualStyleBackColor = True
        '
        'btnSumar
        '
        Me.btnSumar.Font = New System.Drawing.Font("Consolas", 14.25!)
        Me.btnSumar.Location = New System.Drawing.Point(317, 131)
        Me.btnSumar.Name = "btnSumar"
        Me.btnSumar.Size = New System.Drawing.Size(50, 98)
        Me.btnSumar.TabIndex = 6
        Me.btnSumar.Text = "+"
        Me.btnSumar.UseVisualStyleBackColor = True
        '
        'btnMultiplicar
        '
        Me.btnMultiplicar.Font = New System.Drawing.Font("Consolas", 14.25!)
        Me.btnMultiplicar.Location = New System.Drawing.Point(261, 131)
        Me.btnMultiplicar.Name = "btnMultiplicar"
        Me.btnMultiplicar.Size = New System.Drawing.Size(50, 46)
        Me.btnMultiplicar.TabIndex = 5
        Me.btnMultiplicar.Text = "*"
        Me.btnMultiplicar.UseVisualStyleBackColor = True
        '
        'btnNueve
        '
        Me.btnNueve.Location = New System.Drawing.Point(205, 131)
        Me.btnNueve.Name = "btnNueve"
        Me.btnNueve.Size = New System.Drawing.Size(50, 46)
        Me.btnNueve.TabIndex = 4
        Me.btnNueve.Text = "9"
        Me.btnNueve.UseVisualStyleBackColor = True
        '
        'btnResultado
        '
        Me.btnResultado.Font = New System.Drawing.Font("Consolas", 14.25!)
        Me.btnResultado.Location = New System.Drawing.Point(317, 235)
        Me.btnResultado.Name = "btnResultado"
        Me.btnResultado.Size = New System.Drawing.Size(50, 98)
        Me.btnResultado.TabIndex = 17
        Me.btnResultado.Text = "="
        Me.btnResultado.UseVisualStyleBackColor = True
        '
        'btnDividir
        '
        Me.btnDividir.Font = New System.Drawing.Font("Consolas", 14.25!)
        Me.btnDividir.Location = New System.Drawing.Point(261, 183)
        Me.btnDividir.Name = "btnDividir"
        Me.btnDividir.Size = New System.Drawing.Size(50, 46)
        Me.btnDividir.TabIndex = 11
        Me.btnDividir.Text = "/"
        Me.btnDividir.UseVisualStyleBackColor = True
        '
        'btnSeis
        '
        Me.btnSeis.Location = New System.Drawing.Point(205, 183)
        Me.btnSeis.Name = "btnSeis"
        Me.btnSeis.Size = New System.Drawing.Size(50, 46)
        Me.btnSeis.TabIndex = 10
        Me.btnSeis.Text = "6"
        Me.btnSeis.UseVisualStyleBackColor = True
        '
        'btnCinco
        '
        Me.btnCinco.Location = New System.Drawing.Point(149, 183)
        Me.btnCinco.Name = "btnCinco"
        Me.btnCinco.Size = New System.Drawing.Size(50, 46)
        Me.btnCinco.TabIndex = 9
        Me.btnCinco.Text = "5"
        Me.btnCinco.UseVisualStyleBackColor = True
        '
        'btnCuatro
        '
        Me.btnCuatro.Location = New System.Drawing.Point(93, 183)
        Me.btnCuatro.Name = "btnCuatro"
        Me.btnCuatro.Size = New System.Drawing.Size(50, 46)
        Me.btnCuatro.TabIndex = 8
        Me.btnCuatro.Text = "4"
        Me.btnCuatro.UseVisualStyleBackColor = True
        '
        'btnFraccion
        '
        Me.btnFraccion.Location = New System.Drawing.Point(37, 183)
        Me.btnFraccion.Name = "btnFraccion"
        Me.btnFraccion.Size = New System.Drawing.Size(50, 46)
        Me.btnFraccion.TabIndex = 7
        Me.btnFraccion.Text = "a/b"
        Me.btnFraccion.UseVisualStyleBackColor = True
        '
        'btnClearOp
        '
        Me.btnClearOp.Location = New System.Drawing.Point(261, 287)
        Me.btnClearOp.Name = "btnClearOp"
        Me.btnClearOp.Size = New System.Drawing.Size(50, 46)
        Me.btnClearOp.TabIndex = 22
        Me.btnClearOp.Text = "C"
        Me.btnClearOp.UseVisualStyleBackColor = True
        '
        'btnClearEntry
        '
        Me.btnClearEntry.Location = New System.Drawing.Point(205, 287)
        Me.btnClearEntry.Name = "btnClearEntry"
        Me.btnClearEntry.Size = New System.Drawing.Size(50, 46)
        Me.btnClearEntry.TabIndex = 21
        Me.btnClearEntry.Text = "CE"
        Me.btnClearEntry.UseVisualStyleBackColor = True
        '
        'btnSeparadorDecimal
        '
        Me.btnSeparadorDecimal.Font = New System.Drawing.Font("Consolas", 14.25!)
        Me.btnSeparadorDecimal.Location = New System.Drawing.Point(149, 287)
        Me.btnSeparadorDecimal.Name = "btnSeparadorDecimal"
        Me.btnSeparadorDecimal.Size = New System.Drawing.Size(50, 46)
        Me.btnSeparadorDecimal.TabIndex = 20
        Me.btnSeparadorDecimal.Text = "."
        Me.btnSeparadorDecimal.UseVisualStyleBackColor = True
        '
        'btnCero
        '
        Me.btnCero.Location = New System.Drawing.Point(93, 287)
        Me.btnCero.Name = "btnCero"
        Me.btnCero.Size = New System.Drawing.Size(50, 46)
        Me.btnCero.TabIndex = 19
        Me.btnCero.Text = "0"
        Me.btnCero.UseVisualStyleBackColor = True
        '
        'btnRaizCuadrada
        '
        Me.btnRaizCuadrada.Location = New System.Drawing.Point(37, 287)
        Me.btnRaizCuadrada.Name = "btnRaizCuadrada"
        Me.btnRaizCuadrada.Size = New System.Drawing.Size(50, 46)
        Me.btnRaizCuadrada.TabIndex = 18
        Me.btnRaizCuadrada.Text = "sqrt"
        Me.btnRaizCuadrada.UseVisualStyleBackColor = True
        '
        'btnRestar
        '
        Me.btnRestar.Font = New System.Drawing.Font("Consolas", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRestar.Location = New System.Drawing.Point(261, 235)
        Me.btnRestar.Name = "btnRestar"
        Me.btnRestar.Size = New System.Drawing.Size(50, 46)
        Me.btnRestar.TabIndex = 16
        Me.btnRestar.Text = "-"
        Me.btnRestar.UseVisualStyleBackColor = True
        '
        'btnTres
        '
        Me.btnTres.Location = New System.Drawing.Point(205, 235)
        Me.btnTres.Name = "btnTres"
        Me.btnTres.Size = New System.Drawing.Size(50, 46)
        Me.btnTres.TabIndex = 15
        Me.btnTres.Text = "3"
        Me.btnTres.UseVisualStyleBackColor = True
        '
        'btnDos
        '
        Me.btnDos.Location = New System.Drawing.Point(149, 235)
        Me.btnDos.Name = "btnDos"
        Me.btnDos.Size = New System.Drawing.Size(50, 46)
        Me.btnDos.TabIndex = 14
        Me.btnDos.Text = "2"
        Me.btnDos.UseVisualStyleBackColor = True
        '
        'btnUno
        '
        Me.btnUno.Location = New System.Drawing.Point(93, 235)
        Me.btnUno.Name = "btnUno"
        Me.btnUno.Size = New System.Drawing.Size(50, 46)
        Me.btnUno.TabIndex = 13
        Me.btnUno.Text = "1"
        Me.btnUno.UseVisualStyleBackColor = True
        '
        'btnPotenciar
        '
        Me.btnPotenciar.Location = New System.Drawing.Point(37, 235)
        Me.btnPotenciar.Name = "btnPotenciar"
        Me.btnPotenciar.Size = New System.Drawing.Size(50, 46)
        Me.btnPotenciar.TabIndex = 12
        Me.btnPotenciar.Text = "a^b"
        Me.btnPotenciar.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(406, 373)
        Me.Controls.Add(Me.btnClearOp)
        Me.Controls.Add(Me.btnClearEntry)
        Me.Controls.Add(Me.btnSeparadorDecimal)
        Me.Controls.Add(Me.btnCero)
        Me.Controls.Add(Me.btnRaizCuadrada)
        Me.Controls.Add(Me.btnRestar)
        Me.Controls.Add(Me.btnTres)
        Me.Controls.Add(Me.btnDos)
        Me.Controls.Add(Me.btnUno)
        Me.Controls.Add(Me.btnPotenciar)
        Me.Controls.Add(Me.btnResultado)
        Me.Controls.Add(Me.btnDividir)
        Me.Controls.Add(Me.btnSeis)
        Me.Controls.Add(Me.btnCinco)
        Me.Controls.Add(Me.btnCuatro)
        Me.Controls.Add(Me.btnFraccion)
        Me.Controls.Add(Me.btnSumar)
        Me.Controls.Add(Me.btnMultiplicar)
        Me.Controls.Add(Me.btnNueve)
        Me.Controls.Add(Me.btnOcho)
        Me.Controls.Add(Me.btnSiete)
        Me.Controls.Add(Me.btnPorcentaje)
        Me.Controls.Add(Me.tbDisplay)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tbDisplay As TextBox
    Friend WithEvents btnPorcentaje As Button
    Friend WithEvents btnSiete As Button
    Friend WithEvents btnOcho As Button
    Friend WithEvents btnSumar As Button
    Friend WithEvents btnMultiplicar As Button
    Friend WithEvents btnNueve As Button
    Friend WithEvents btnResultado As Button
    Friend WithEvents btnDividir As Button
    Friend WithEvents btnSeis As Button
    Friend WithEvents btnCinco As Button
    Friend WithEvents btnCuatro As Button
    Friend WithEvents btnFraccion As Button
    Friend WithEvents btnClearOp As Button
    Friend WithEvents btnClearEntry As Button
    Friend WithEvents btnSeparadorDecimal As Button
    Friend WithEvents btnCero As Button
    Friend WithEvents btnRaizCuadrada As Button
    Friend WithEvents btnRestar As Button
    Friend WithEvents btnTres As Button
    Friend WithEvents btnDos As Button
    Friend WithEvents btnUno As Button
    Friend WithEvents btnPotenciar As Button
End Class
