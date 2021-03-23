Public Class Form1

    Private gymMainWordDocument As clsWordDocument
    Private gymReferenceWordDocument As clsWordDocument

    Private gsFolderNameOfTemplates As String
    Private gsFolderNameOfReferences As String
    Private gsFolderNameOfOutputs As String

    Public Const gcFileExtension As String = ".docx"
    Public Const gcTemplatesFolderName As String = "\Templates"
    Public Const gcReferencesFolderName As String = "\References"
    Public Const gcOutputsFolderName As String = "\Outputs"
    Public Const gcTempFolderName As String = "\temp"

    Public gbIsLoadingTemplate As Boolean


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.gsFolderNameOfTemplates = Application.StartupPath() & Form1.gcTemplatesFolderName
        Me.gsFolderNameOfReferences = Application.StartupPath() & Form1.gcReferencesFolderName
        Me.gsFolderNameOfOutputs = Application.StartupPath() & Form1.gcOutputsFolderName

    End Sub

    Private Sub btnLoadDocument_Click(sender As Object, e As EventArgs) Handles btnLoadDocument.Click

        Me.gbIsLoadingTemplate = True

        Call Me.subbtnLoadDocument_Click(bvImportReferenceDataBeforeExecutingFunctions:=False)

        Me.gbIsLoadingTemplate = False

    End Sub

    Private Sub subbtnLoadDocument_Click(ByVal bvImportReferenceDataBeforeExecutingFunctions As Boolean)

        'Select a file
        Dim sFullPath As String
        sFullPath = Me.fncHaveUserSelectFile(bvInitialDirectory:=Me.gsFolderNameOfTemplates)

        If sFullPath = "" Then
            Exit Sub
        End If

        'Load Document
        Me.gymMainWordDocument = New clsWordDocument
        Call Me.gymMainWordDocument.subLoadDocument(bvFullPath:=sFullPath)

        'Create GUI components
        Call subFillFlowLayoutPanel1()

        If bvImportReferenceDataBeforeExecutingFunctions Then
            'Import the reference values
            Call subDocumentFunction_ImportReferenceData()
        End If

        'Activate functions
        Call subListArgumentsThatActivateFunctionsAndActivateThoseFunctionsPerSequence(bvWordDocument:=Me.gymMainWordDocument)

        Me.Text = "Yasminator: " & IO.Path.GetFileName(sFullPath)


    End Sub

    Private Sub btnLoadReferenceDocument_Click(sender As Object, e As EventArgs) Handles btnLoadReferenceDocument.Click

        Call Me.subHaveUserSelectFileAndThenLoadReferenceDocument()

    End Sub

    Private Sub subHaveUserSelectFileAndThenLoadReferenceDocument()

        'Select a file
        Dim sFullPath As String
        sFullPath = Me.fncHaveUserSelectFile(bvInitialDirectory:=Me.gsFolderNameOfReferences)

        Call Me.subDocumentFunction_LoadReferenceDocument(bvFullPath:=sFullPath)

    End Sub

    Private Sub btnKeepInMemoryAsRefenceAndOpenNew_Click(sender As Object, e As EventArgs) Handles btnKeepInMemoryAsRefenceAndOpenNew.Click

        'Me.gymMainWordDocument.PrepareToBeAssignedAsReferenceDocumentInRam()
        For Each feUCT As uctArgumentBase In Me.FlowLayoutPanel1.Controls
            feUCT.subSetCurrentValueAsDefaultParameter()
        Next


        Me.FlowLayoutPanel1.Controls.Clear()

        Me.gymReferenceWordDocument = Me.gymMainWordDocument

        Call Me.subbtnLoadDocument_Click(bvImportReferenceDataBeforeExecutingFunctions:=True)


    End Sub

    Private Sub btnSaveAndLaunch_Click(sender As Object, e As EventArgs) Handles btnSaveAndLaunch.Click

        For Each uctArgumentGUI As uctArgumentBase In Me.FlowLayoutPanel1.Controls
            uctArgumentGUI.CommitValueToParagraph()
        Next

        'Save and Launch
        My.Computer.FileSystem.CreateDirectory(Me.gsFolderNameOfOutputs)

        Dim sNewFullPath As String = Me.gsFolderNameOfOutputs & "\" & gymMainWordDocument.gsFileNameWithoutExtension & " (edited " & Me.fncGetNowAsStringForSaveFile & ")" & Form1.gcFileExtension

        Call gymMainWordDocument.subSaveToFile(bvNewFullPath:=sNewFullPath)


        System.Diagnostics.Process.Start(sNewFullPath)

        For Each uctArgumentGUI As uctArgumentBase In Me.FlowLayoutPanel1.Controls
            uctArgumentGUI.RevertToOriginalParagraphText()
        Next

    End Sub

    Private Function fncGetNowAsStringForSaveFile()
        Return Format(Now.Year, "0000") & "-" &
               Format(Now.Month, "00") & "-" &
               Format(Now.Day, "00") & " " &
               Format(Now.Hour, "00") & ";" &
               Format(Now.Minute, "00") & ";" &
               Format(Now.Second, "00")
    End Function

    Private Function fncHaveUserSelectFile(ByVal bvInitialDirectory As String) As String

        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Title = "Select a docx file to edit"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Docx Files|*" & Form1.gcFileExtension
        OpenFileDialog1.InitialDirectory = bvInitialDirectory

        OpenFileDialog1.ShowDialog()

        If OpenFileDialog1.FileName = "" Then
            MsgBox("No file selected")
            Return ""
        End If

        Return OpenFileDialog1.FileName

    End Function

    Private Sub subFillFlowLayoutPanel1()

        Me.FlowLayoutPanel1.Controls.Clear()

        Dim ymArgumentGUI As uctArgumentBase

        'New uctArgumentGUI
        For Each feArgument As clsArgument In Me.gymMainWordDocument.gmsListOfArguments

            If feArgument.IsNumeric Or feArgument.IsNumericAsText Then

                ymArgumentGUI = New uctArgumentNumeric

            ElseIf feArgument.IsMultiplication Then
                ymArgumentGUI = New uctArgumentMultiplication

            ElseIf feArgument.IsDate_ Then
                ymArgumentGUI = New uctArgumentDate

            Else
                ymArgumentGUI = New uctArgumentText

            End If



            ymArgumentGUI.subInitialize(bvArgument:=feArgument)

            Me.FlowLayoutPanel1.Controls.Add(ymArgumentGUI)

        Next

        'IsCopyFromArgument
        For Each feArgumentThatIsPontentiallyCopying As clsArgument In Me.gymMainWordDocument.gmsListOfArguments
            If feArgumentThatIsPontentiallyCopying.IsCopyFromArgument Then
                For Each feArgumentToCopyFrom As clsArgument In Me.gymMainWordDocument.gmsListOfArguments
                    If Form1.CompareText(bvText1:=feArgumentToCopyFrom.TextToShow,
                                         vText2:=feArgumentThatIsPontentiallyCopying.CopyFromArgumentTexToShow) Then

                        If feArgumentToCopyFrom.ListOfArgumentsToSendValueTo Is Nothing Then
                            feArgumentToCopyFrom.ListOfArgumentsToSendValueTo = New ArrayList
                        End If

                        feArgumentToCopyFrom.ListOfArgumentsToSendValueTo.Add(feArgumentThatIsPontentiallyCopying)

                        feArgumentThatIsPontentiallyCopying.subSetValueFromOtherArgument(bvValue:=feArgumentToCopyFrom.CurrentValueAsString)

                    End If
                Next

            End If
        Next

        'IsMultiplication
        For Each feArgumentThatPontentiallyHasMath As clsArgument In Me.gymMainWordDocument.gmsListOfArguments
            If feArgumentThatPontentiallyHasMath.IsMultiplication OrElse feArgumentThatPontentiallyHasMath.IsAddition Then
                For Each feArgumentToCopyFrom As clsArgument In Me.gymMainWordDocument.gmsListOfArguments

                    If feArgumentToCopyFrom.TextToShow = feArgumentThatPontentiallyHasMath.MathValueFromArgument1 Then

                        feArgumentThatPontentiallyHasMath._MathValueFromArgument1 = feArgumentToCopyFrom

                        If feArgumentToCopyFrom.MathReceivers Is Nothing Then
                            feArgumentToCopyFrom.MathReceivers = New ArrayList
                        End If
                        feArgumentToCopyFrom.MathReceivers.Add(feArgumentThatPontentiallyHasMath)

                    End If

                    If feArgumentToCopyFrom.TextToShow = feArgumentThatPontentiallyHasMath.MathValueFromArgument2 Then

                        feArgumentThatPontentiallyHasMath._MathValueFromArgument2 = feArgumentToCopyFrom

                        If feArgumentToCopyFrom.MathReceivers Is Nothing Then
                            feArgumentToCopyFrom.MathReceivers = New ArrayList
                        End If
                        feArgumentToCopyFrom.MathReceivers.Add(feArgumentThatPontentiallyHasMath)

                    End If

                Next

                'No need to provide a value, it's calculated
                feArgumentThatPontentiallyHasMath.subSetValue(bvValue:=Nothing)


            End If

        Next

    End Sub

    Private Sub subDocumentFunction_ImportReferenceData()

        If Me.gymMainWordDocument Is Nothing Then
            Exit Sub
        End If

        For Each feArgument As clsArgument In Me.gymMainWordDocument.gmsListOfArguments
            If feArgument.ImportsFromReference Then
                For Each feArgumentThatCanBeImported As clsArgument In Me.gymReferenceWordDocument.gmsListOfArguments
                    If Form1.CompareText(bvText1:=feArgumentThatCanBeImported.TextToShow,
                                         vText2:=feArgument.ReferenceTextToShow.ToUpper) Then

                        feArgument.subSetValue(bvValue:=feArgumentThatCanBeImported.DefaultValue)

                    End If
                Next
            End If
        Next

    End Sub

    Private Sub subListArgumentsThatActivateFunctionsAndActivateThoseFunctionsPerSequence(ByVal bvWordDocument As clsWordDocument)

        'Capture all arguments that are function calls
        'Calls to functions such as automatically loading a references are found in documents, and there may be more than one in each document.
        'They can have a ActionSequence to specify the order they are executed. 
        'It is possible the user repeats the same ActionSequence for more than one Argument.
        Dim msDictionaryOfListOfArgumetns As Dictionary(Of Integer, ArrayList)
        msDictionaryOfListOfArgumetns = New Dictionary(Of Integer, ArrayList)

        '-1 will result in no loops if nothing is added when it loops from 0
        Dim iHighestActionSequence As Integer = -1

        'Loop though all argumetns
        For Each feArgument As clsArgument In bvWordDocument.gmsListOfArguments

            'If this argument is a function call
            If feArgument.ActivateFunction <> "" Then

                'ActionSequence has to be given a positive Parameter Value. Any negative values will be processed as 0.
                Dim iActionSequence As Integer
                iActionSequence = feArgument.ActionSequence
                If iActionSequence < 0 Then
                    iActionSequence = 0
                End If

                If iActionSequence > iHighestActionSequence Then
                    iHighestActionSequence = iActionSequence
                End If

                'See if there is an array list for it's sequence number in the dictionary
                Dim aListOfArguments As ArrayList
                aListOfArguments = Nothing
                msDictionaryOfListOfArgumetns.TryGetValue(key:=feArgument.ActionSequence, value:=aListOfArguments)

                'If this is the first Argument with this ActionSequence, create an Arraylist for this ActionSequence number.
                If aListOfArguments Is Nothing Then
                    aListOfArguments = New ArrayList
                    msDictionaryOfListOfArgumetns.Add(key:=feArgument.ActionSequence, value:=aListOfArguments)
                End If

                aListOfArguments.Add(feArgument)
            End If

        Next

        'Execute them in accordance with their ActionSequence
        For iActionSequence = 0 To iHighestActionSequence

            Dim aListOfArguments As ArrayList
            aListOfArguments = Nothing

            msDictionaryOfListOfArgumetns.TryGetValue(key:=iActionSequence, value:=aListOfArguments)

            'Are there any arguments at this ActionSequence?
            If aListOfArguments IsNot Nothing Then

                'Now that we have the relevant arguments, we excecute them one at a time.
                For Each feArguments As clsArgument In aListOfArguments

                    Call Me.subActivateFunction(bvArgument:=feArguments)

                Next

            End If
        Next

    End Sub

    Private Sub subActivateFunction(bvArgument As clsArgument)

        If bvArgument.ActivateFunction <> "" Then

            Dim sFunctionToActivate As String
            sFunctionToActivate = bvArgument.ActivateFunction.Trim.ToUpper

            If sFunctionToActivate.StartsWith("LoadReference".ToUpper) Then

                Dim sPath As String
                sPath = sFunctionToActivate.Substring(startIndex:="LoadReference:".Length)

                Dim sFullPath As String
                sFullPath = Application.StartupPath() & "\" & sPath.Trim

                Call Me.subDocumentFunction_LoadReferenceDocument(bvFullPath:=sFullPath)

            ElseIf sFunctionToActivate.StartsWith("CalculateWorkdays".ToUpper) Then
                Call subDocumentFunction_CalculateWorkdays()

            ElseIf sFunctionToActivate.StartsWith("ImportReferenceData".ToUpper) Then
                'Import the reference values
                Call subDocumentFunction_ImportReferenceData()

            End If

        End If

    End Sub

    Public Sub subDocumentFunction_CalculateWorkdays()

        If Me.gymMainWordDocument Is Nothing Then
            Exit Sub
        End If

        For Each feArgument As clsArgument In Me.gymMainWordDocument.gmsListOfArguments
            'example false: "" AndAlso ""
            'example true: Milestone2StartDate" AndAlso " Milestone 2 due date"
            If feArgument.WorkDaysStart <> Nothing AndAlso feArgument.WorkDaysEnd <> Nothing Then

                Dim iAmountDays As Integer
                iAmountDays = Me.fncGetAmountDaydsWorked(bvArgument:=feArgument)

                feArgument.subSetValue(bvValue:=iAmountDays)

            End If
        Next

    End Sub

    Private Function fncGetAmountDaydsWorked(ByVal bvArgument As clsArgument) As Integer

        Dim ymArgument As clsArgument
        Dim sWorkDaysStart As String = Nothing
        Dim sWorkDaysEnd As String = Nothing

        Dim dWorkDaysStart As Date
        Dim dWorkDaysEnd As Date

        Dim iErrorValue As Integer = -1

        If bvArgument.WorkDaysStart <> Nothing Then
            ymArgument = Me.fncGetArgumentThroughTextToShow(vbTextToShow:=bvArgument.WorkDaysStart)
            ymArgument.RedoWorkDayCalculationsOnValueChange = True
            sWorkDaysStart = ymArgument.CurrentValueAsString

            ''Take the string input of sWorkDaysStart and using the DateFormat, create the date variable dWorkDaysStart
            'If sWorkDaysStart <> Nothing Then
            '    dWorkDaysStart = Date.ParseExact(s:=sWorkDaysStart, format:=ymArgument.DateFormat, provider:=Nothing)
            'Else
            '    'dWorkDaysStart = Nothing
            '    Return iErrorValue
            'End If
            If sWorkDaysStart <> Nothing Then
                dWorkDaysStart = CDate(sWorkDaysStart)
            End If
        End If

        If bvArgument.WorkDaysEnd <> Nothing Then
            ymArgument = Me.fncGetArgumentThroughTextToShow(vbTextToShow:=bvArgument.WorkDaysEnd)
            ymArgument.RedoWorkDayCalculationsOnValueChange = True
            sWorkDaysEnd = ymArgument.CurrentValueAsString
            'dWorkDaysEnd = Date.ParseExact(s:=sWorkDaysEnd, format:=ymArgument.DateFormat, provider:=Nothing)
            If sWorkDaysEnd <> Nothing Then
                dWorkDaysEnd = CDate(sWorkDaysEnd)
            End If
        End If

        If dWorkDaysStart > dWorkDaysEnd Then
            Return iErrorValue
        End If

        Dim iAmountDays As Integer

        iAmountDays = (dWorkDaysEnd - dWorkDaysStart).Days


        Dim dDateBuffer As Date

        dDateBuffer = dWorkDaysStart
        iAmountDays = 0


        'Start with the first day, evaluate it, and keep as long as you havent passed the last day
        Do Until dDateBuffer > dWorkDaysEnd

            'Is this a weekend day?
            If dDateBuffer.DayOfWeek <> DayOfWeek.Saturday AndAlso
                dDateBuffer.DayOfWeek <> DayOfWeek.Sunday Then

                'Add a workday
                iAmountDays += 1

            End If

            dDateBuffer = dDateBuffer.AddDays(1)

        Loop


        'Check again for weekdays that are aren't workdays according to the reference file
        dDateBuffer = dWorkDaysStart

        Dim iWorkDaysOff As Integer = 0

        'Start with the first day, evaluate it, and keep as long as you havent passed the last day
        Do Until dDateBuffer > dWorkDaysEnd

            For Each feArgument As clsArgument In Me.gymReferenceWordDocument.gmsListOfArguments

                Try

                    'Is this day a day off?
                    If feArgument.DefaultValue <> "" AndAlso feArgument.DefaultValue = dDateBuffer Then

                        'remove a workday
                        iWorkDaysOff += 1

                    End If

                Catch ex As Exception
                    If Not My.Computer.Keyboard.CtrlKeyDown Then
                        MsgBox("Reference file for days off isn't loaded or is corrupt. " & Environment.NewLine &
                           "The calculation for days off does not consider the reference file for days off. " & Environment.NewLine &
                           "" & Environment.NewLine &
                           "This could be due to another reference file not automatically opening the days off reference file after it's done." & Environment.NewLine &
                           "If that is the case, the WordEditor tool is working as intented, the reference file isn't." & Environment.NewLine &
                           "A short term fix is to open some other reference file that is configured to automatically open the reference files for days off." & Environment.NewLine &
                           "Contact the administrator for further help." & Environment.NewLine &
                           "" & Environment.NewLine &
                           "Hold the ctr button to prevent this message from appearing." & Environment.NewLine)
                    End If

                    iWorkDaysOff = 0
                    Exit Do
                End Try
            Next


            dDateBuffer = dDateBuffer.AddDays(1)

        Loop

        iAmountDays = iAmountDays - iWorkDaysOff

        Return iAmountDays

    End Function

    Private Function fncGetArgumentThroughTextToShow(ByVal vbTextToShow As String) As clsArgument

        For Each feArgument As clsArgument In Me.gymMainWordDocument.gmsListOfArguments

            If Form1.CompareText(bvText1:=feArgument.TextToShow,
                                 vText2:=vbTextToShow) Then

                Return feArgument

            End If

        Next

        Return Nothing

    End Function

    Private Sub subDocumentFunction_LoadReferenceDocument(ByVal bvFullPath As String)

        If bvFullPath = "" Then
            Exit Sub
        End If

        'Load Document
        Me.gymReferenceWordDocument = New clsWordDocument
        Call Me.gymReferenceWordDocument.subLoadDocument(bvFullPath:=bvFullPath)


        'Activate functions
        Call Me.subListArgumentsThatActivateFunctionsAndActivateThoseFunctionsPerSequence(bvWordDocument:=Me.gymReferenceWordDocument)

    End Sub

    Private Shared Function CompareText(bvText1 As String, vText2 As String) As Boolean

        If bvText1.Trim.ToUpper = vText2.Trim.ToUpper Then
            Return True
        Else
            Return False
        End If

    End Function



End Class