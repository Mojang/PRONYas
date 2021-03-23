Public Class uctArgumentMultiplication
    Inherits uctArgumentNumeric


    Public Overrides Sub subInitialize(ByVal bvArgument As clsArgument)

        Call MyBase.subInitialize(bvArgument:=bvArgument)

    End Sub


    Protected Overrides Function fncGetValueInputedByUser() As String

        Return Me.gymArgument.fncCalculateMultiplicatedValue

    End Function

    Public Overrides Sub subDisplayValue(ByVal bvValue As String)

        MyBase.NumericUpDown2.Text = bvValue

    End Sub

    Public Overrides Function fncGetValueToPrint() As String

        Dim sValue As String
        sValue = Me.gymArgument.fncCalculateMultiplicatedValue


        If Me.gymArgument.ThousandSeparated Then
            'Text, not value, as text also includes thousand separators



            'Dim msNumerircUpDown As New NumericUpDown
            'msNumerircUpDown.ThousandsSeparator = True
            'msNumerircUpDown.Value = sValue
            'sValue = msNumerircUpDown.Text

            sValue = MyBase.NumericUpDown2.Text

        End If

        If Me.gymArgument.NoOutput Then
            sValue = ""
        End If

        Return sValue

    End Function

End Class
