Public Class uctArgumentNumeric

    Inherits uctArgumentBase


    Private Sub Control_TextChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged

        Call Me.subSetValueInputedByUser()

    End Sub


    Public Overrides Sub subInitialize(ByVal bvArgument As clsArgument)

        Call MyBase.subInitialize(bvArgument:=bvArgument)

        Me.NumericUpDown2.Maximum = Me.gymArgument.NumerMax

        'DefaultValue
        If Me.gymArgument.HasDefaultValue Then
            Me.NumericUpDown2.Value = MyBase.gymArgument.DefaultValue
        End If

        'ThousandSeparated
        If Me.gymArgument.ThousandSeparated Then
            Me.NumericUpDown2.ThousandsSeparator = True
        End If

    End Sub

    Protected Overrides Function fncGetValueInputedByUser() As String

        Return Me.NumericUpDown2.Value

    End Function

    Public Overrides Sub subDisplayValue(ByVal bvValue As String)

        Try



            Me.NumericUpDown2.Value = bvValue

        Catch ex As Exception
            Me.NumericUpDown2.Value = 0

        End Try
    End Sub

    Public Overrides Function fncGetValueToPrint() As String

        Dim sValue As String

        If Me.gymArgument.IsNumericAsText Then
            sValue = clsNumeriCon.fncConvertNum(Input:=Me.NumericUpDown2.Value)
        Else
            'Text, not value, as text also includes thousand separators
            sValue = Me.NumericUpDown2.Text
        End If

        If Me.gymArgument.NoOutput Then
            sValue = ""
        End If

        Return sValue

    End Function


End Class
