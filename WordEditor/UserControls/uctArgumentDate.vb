Public Class uctArgumentDate

    Inherits uctArgumentBase


    Private Sub Control_TextChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

        Call Me.subSetValueInputedByUser()

    End Sub


    Public Overrides Sub subInitialize(ByVal bvArgument As clsArgument)

        Call MyBase.subInitialize(bvArgument:=bvArgument)

        'DateFormat
        If Me.gymArgument.DateFormat IsNot Nothing Then
            Me.DateTimePicker2.Format = DateTimePickerFormat.Custom
            Me.DateTimePicker2.CustomFormat = Me.gymArgument.DateFormat
        End If

        'DefaultValue
        If Me.gymArgument.HasDefaultValue Then
            If Me.gymArgument.DefaultValue.Trim.ToUpper = clsWordDocument.gcsCharNOW Then
                Me.DateTimePicker2.Value = New Date(year:=Now.Year, month:=Now.Month, day:=Now.Day)
                'ElseIf Me.gymArgument.DefaultValue.Trim.ToUpper = clsWordDocument.gcsCharToday Then
                '    Me.DateTimePicker2.Value = New Date(year:=Now.Year, month:=Now.Month, day:=Now.Day)
            Else

                Me.DateTimePicker2.Value = Me.gymArgument.DefaultValue
            End If

        End If

        If Me.gymArgument.InitalizeAtEndOfMonth Or Me.gymArgument.EndOfMonth Then
            Me.DateTimePicker2.Value = Functions.fncLastDay(bvDate:=Me.DateTimePicker2.Value)
        End If

    End Sub

    Protected Overrides Function fncGetValueInputedByUser() As String

        Return Me.DateTimePicker2.Value
        'Return Format(Expression:=Me.DateTimePicker2.Value, Style:=Me.gymArgument.DateFormat)

    End Function

    Public Overrides Sub subDisplayValue(ByVal bvValue As String)

        If Not IsDate(bvValue) Then

        End If


        Try
            'This will result in an error in cases where it's a date control that displays a date as in the format "mmmm". That makes 2020-01-01 into the string "Januari". 
            'The different loops in the program might result in trying to feed that string value back into this control as a date, and "Januari" can't be simply converted into a date
            'so in that case, just keep the current value
            Me.DateTimePicker2.Value = bvValue

        Catch ex As Exception
            'keep current value
            Stop

        End Try

    End Sub

    Public Overrides Function fncGetValueToPrint() As String

        Dim sValue As String

        If Me.gymArgument.HasDateFormat Then
            'sValue = Format(Me.DateTimePicker2.Value, Me.gymArgument.DateFormat)
            sValue = Me.gymArgument.CurrentValueAsStringWithDateFormat
        Else
            sValue = Me.gymArgument.CurrentValueAsString
        End If

        If Me.gymArgument.NoOutput Then
            sValue = ""
        End If

        If sValue Is Nothing Then
            Return Nothing
        End If

        sValue = sValue.Replace("januari", "January")
        sValue = sValue.Replace("februari", "February")
        sValue = sValue.Replace("mars", "March")
        sValue = sValue.Replace("april", "April")
        sValue = sValue.Replace("maj", "May")
        sValue = sValue.Replace("juni", "June")
        sValue = sValue.Replace("juli", "July")
        sValue = sValue.Replace("augusti", "August")
        sValue = sValue.Replace("september", "September")
        sValue = sValue.Replace("oktober", "October")
        sValue = sValue.Replace("november", "November")
        sValue = sValue.Replace("december", "December")

        Return sValue


    End Function

End Class
