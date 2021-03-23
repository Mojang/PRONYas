Public Class uctArgumentText

    Inherits uctArgumentBase


    Private Sub Control_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        Call Me.subSetValueInputedByUser()

    End Sub


    Public Overrides Sub subInitialize(ByVal bvArgument As clsArgument)

        Call MyBase.subInitialize(bvArgument:=bvArgument)

        'Multiline
        If Me.gymArgument.IsMultiline Then
            Me.TextBox2.Multiline = True

            If Me.gymArgument.NumberOfLines = 0 Then
                Me.gymArgument.NumberOfLines = 2
            End If

            Dim iHeightIncrease As Integer
            iHeightIncrease = Me.TextBox2.Height * (Me.gymArgument.NumberOfLines - 1)

            Me.TextBox2.Height += iHeightIncrease
            Me.Height += iHeightIncrease

        End If

        'DefaultValue
        If Me.gymArgument.HasDefaultValue Then
            Me.TextBox2.Text = Me.gymArgument.DefaultValue
        End If

    End Sub

    Protected Overrides Function fncGetValueInputedByUser() As String
        Return Me.TextBox2.Text
    End Function

    Public Overrides Sub subDisplayValue(ByVal bvValue As String)
        Me.TextBox2.Text = bvValue
    End Sub

    Public Overrides Function fncGetValueToPrint() As String

        Dim sValue As String

        sValue = Me.TextBox2.Text

        If Not Me.gymArgument.NoValueOnEmpty Then
            If sValue = "" Then
                sValue = "[Insert details]"
            End If
        End If

        If Me.gymArgument.NoOutput Then
            sValue = ""
        End If

        Return sValue

    End Function


End Class
