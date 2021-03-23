Public Class uctArgumentBase

    Protected gymArgument As clsArgument
    'Private gbProgramaticallyChangingTextboxValue As Boolean

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Overridable Sub subInitialize(ByVal bvArgument As clsArgument)


        Me.gymArgument = bvArgument
        Me.gymArgument.MyUserControl = Me

        If Me.gymArgument.NoSequenceWasProvided Then
            Me.Label1.Text = "#?: "
        Else
            Me.Label1.Text = "#" & Me.gymArgument.Sequence & ": "
        End If

        If Me.gymArgument.TextToShow = "" Then
            Me.Label1.Text &= "[not specified]:"
        Else
            Me.Label1.Text &= Me.gymArgument.TextToShow & ":"
        End If

        'INVISIBLE
        If Me.gymArgument.INVISIBLE Then
            Me.Visible = False
        End If

        'DISABLED
        If Me.gymArgument.DISABLED Then
            Me.Enabled = False
        End If

        'HasCopyFromArgument
        If Me.gymArgument.IsCopyFromArgument Then

            If Me.gymArgument.TextToShow = "" Then
                Me.Label1.Text = Me.gymArgument.Sequence & ": " & "Copy:" & Me.gymArgument.CopyFromArgumentTexToShow & ":"
            End If

        End If

        'Reference
        If Me.gymArgument.ImportsFromReference Then
            'do nothing?
        End If

    End Sub

    Protected Overridable Function fncGetValueInputedByUser() As String
        'MustOverride
        Throw New NotImplementedException
    End Function

    Public Sub subSetValueInputedByUser()

        Dim sValue As String
        sValue = Me.fncGetValueInputedByUser

        Me.gymArgument.subSetValue(bvValue:=sValue)

    End Sub

    Public Sub subSetCurrentValueAsDefaultParameter()
        Call Me.gymArgument.subSetCurrentValueAsDefaultParameter()
    End Sub

    Public Sub CommitValueToParagraph()
        Me.gymArgument.CommitValueToParagraph(bvValueToPrint:=Me.fncGetValueToPrint)
    End Sub

    Public Sub RevertToOriginalParagraphText()
        Me.gymArgument.RevertToOriginalParagraphText()
    End Sub

    Public Overridable Sub subDisplayValue(ByVal bvValue As String)
        'MustOverride
        Throw New NotImplementedException
    End Sub

    Public Overridable Function fncGetValueToPrint() As String
        'MustOverride
        Throw New NotImplementedException
    End Function

End Class
