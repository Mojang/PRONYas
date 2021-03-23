Imports Spire.Doc.Documents

Public Class clsArgument

    Public Sequence As Integer
    Public NoSequenceWasProvided As Boolean
    Public ArgumentAsIsInDocument As String
    Public TextToShow As String
    Public DefaultValue As String
    Public HasDefaultValue As String

    Public ImportsFromReference As Boolean
    Public ReferenceTextToShow As String

    Public IsMultiline As Boolean
    Public NumberOfLines As Integer
    Public KeyMultiline As String
    Public IsNumeric As Boolean
    Public ThousandSeparated As Boolean
    Public IsNumericAsText As Boolean
    Public NumerMax As Integer

    Public IsMultiplication As Boolean
    Public IsAddition As Boolean
    Public MathValueFromArgument1 As String
    Public _MathValueFromArgument1 As clsArgument
    Public MathValueFromArgument2 As String
    Public _MathValueFromArgument2 As clsArgument

    Public MathReceivers As ArrayList

    Public IsDate_ As Boolean
    Public HasDateFormat As Boolean
    Public DateFormat As String
    Public AddDays As Integer
    Public EndOfMonth As Boolean
    Public InitalizeAtEndOfMonth As Boolean

    Public IsCopyFromArgument As Boolean
    Public CopyFromArgumentTexToShow As String

    Public NoValueOnEmpty As Boolean
    Public NoOutput As Boolean
    Public INVISIBLE As Boolean
    Public DISABLED As Boolean

    Public ActivateFunction As String
    Public WorkDaysStart As String
    Public WorkDaysEnd As String
    Public RedoWorkDayCalculationsOnValueChange As Boolean

    Public LoadReference As String
    Public HasActionSequence As Boolean
    Public ActionSequence As Integer

    Public CurrentValueAsString As String
    Public ReadOnly Property CurrentValueAsStringWithDateFormat As String
        Get

            If Me.IsDate_ AndAlso Me.HasDateFormat AndAlso Information.IsDate(CurrentValueAsString) Then
                Return Format(CDate(CurrentValueAsString), Me.DateFormat)
            Else
                Return CurrentValueAsString
            End If

        End Get
    End Property


    Public MyUserControl As uctArgumentBase

    Private gsOriginalParagraphText As String
    Private _Paragraph As Spire.Doc.Documents.Paragraph
    Public ReadOnly Property Paragraph As Spire.Doc.Documents.Paragraph
        Get
            Return Me._Paragraph
        End Get
    End Property

    Public ListOfArgumentsToSendValueTo As ArrayList

    Public Sub subSetValue(ByVal bvValue As String)

        If Me.IsMultiplication Then
            Me.CurrentValueAsString = Me.fncCalculateMultiplicatedValue
        ElseIf Me.IsAddition Then
            Me.CurrentValueAsString = Me.fncCalculateAdditionValue
        Else
            Me.CurrentValueAsString = bvValue
        End If

        If MyUserControl IsNot Nothing Then
            MyUserControl.subDisplayValue(bvValue:=Me.CurrentValueAsString)
        End If

        Call Me.subSendValueToReceivers()

        Call Me.UpdateMathReceivers()

        If Me.RedoWorkDayCalculationsOnValueChange Then
            Call clsMain.gfrmMain.subDocumentFunction_CalculateWorkdays()
        End If

    End Sub

    Sub subSetValueFromOtherArgument(ByVal bvValue As String)

        If Me.IsDate_ Then

            If bvValue Is Nothing Then Exit Sub

            If Me.AddDays > 0 Then
                Dim dDate As Date
                dDate = bvValue
                dDate = dDate.AddDays(Me.AddDays)
                bvValue = dDate
            End If

            If Me.EndOfMonth Then

                Dim dDate As Date
                dDate = Functions.fncLastDay(bvDate:=bvValue)

                'Dim sFormatedValue As String
                'sFormatedValue = Format(Expression:=dDate, Style:=Me.DateFormat)
                'bvValue = sFormatedValue
                bvValue = dDate
            End If

            If clsMain.gfrmMain.gbIsLoadingTemplate Then
                If Me.InitalizeAtEndOfMonth Then

                    Dim dDate As Date
                    dDate = Functions.fncLastDay(bvDate:=bvValue)

                    'Dim sFormatedValue As String
                    'sFormatedValue = Format(Expression:=dDate, Style:=Me.DateFormat)
                    'bvValue = sFormatedValue
                    bvValue = dDate
                End If
            End If

        End If

        Call Me.subSetValue(bvValue:=bvValue)

    End Sub

    Public Sub CommitValueToParagraph(ByVal bvValueToPrint As String) ''

        'The Paragraph should only be changed as the last action, as that will result in the key strings in the document to be replaced and unable to be found again.
        Me.Paragraph.Text = Me.Paragraph.Text.Replace(oldValue:=Me.ArgumentAsIsInDocument, newValue:=bvValueToPrint)

    End Sub

    Public Sub RevertToOriginalParagraphText()

        Me.Paragraph.Text = Me.gsOriginalParagraphText

    End Sub


    Public Sub subSendValueToReceivers()

        If Me.ListOfArgumentsToSendValueTo Is Nothing Then
            Exit Sub
        End If

        For Each feArgument As clsArgument In Me.ListOfArgumentsToSendValueTo
            feArgument.subSetValueFromOtherArgument(bvValue:=CurrentValueAsString)
        Next

    End Sub

    Function fncCalculateMultiplicatedValue() As String
        Return Me._MathValueFromArgument1.CurrentValueAsString * Me._MathValueFromArgument2.CurrentValueAsString
    End Function

    Function fncCalculateAdditionValue() As String
        Return CDec(Me._MathValueFromArgument1.CurrentValueAsString) + CDec(Me._MathValueFromArgument2.CurrentValueAsString)
    End Function

    Private Sub UpdateMathReceivers()

        If Me.MathReceivers Is Nothing Then
            Exit Sub
        End If

        For Each feArgument As clsArgument In Me.MathReceivers
            feArgument.subSetValue(bvValue:=Nothing)
        Next

    End Sub

    Sub subSetCurrentValueAsDefaultParameter()

        Me.HasDefaultValue = True
        Me.DefaultValue = Me.CurrentValueAsString

    End Sub

    Public Sub subSetParagraphFromSource(ByVal bvSourceParagraph As Paragraph)
        Me._Paragraph = bvSourceParagraph
        Me.gsOriginalParagraphText = Me._Paragraph.Text
    End Sub
End Class
