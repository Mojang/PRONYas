Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Public Class clsWordDocument


    Private gspDocument As New Spire.Doc.Document()
    Private gsFullPath As String

    Public gsFileNameWithoutExtension As String

    Public gmsListOfArguments As New ArrayList

    Private Const gcsCharArgumentStart As String = "{["
    Private Const gcsCharArgumentEnd As String = "]}"
    Private Const gcsCharArgumentParameterDelimiter As String = "|"
    Private Const gcsCharArgumentParameterValueDelimiter As String = ":"
    Private Const gcsCharPrefixSequence As String = "#" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharPrefixTexToShow As String = "$" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharPrefixTexToShowExplicit As String = "TexToShow" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharPrefixTextToShowDefault As String = "D$" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharPrefixImportFromReference As String = "R$" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharKeyMultiline As String = "MULTILINE" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharIsDate As String = "DATE" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Public Const gcsCharNOW As String = "NOW"
    Private Const gcsAddDays As String = "AddDays" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsEndOfMonth As String = "EndOfMonth"
    Private Const gcsInitalizeAtEndOfMonth As String = "InitalizeAtEndOfMonth"
    Private Const gcsCharIsNumeric As String = "NUM" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharThousandSeparated As String = "ThousandSeparated"
    Private Const gcsCharIsNumAsText As String = "NumAsText" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharIsCOMMENT As String = "COMMENT" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharNoValueOnEmpty As String = "NoValueOnEmpty"
    Private Const gcsCharNoOutput As String = "NoOutput"
    Private Const gcsCharIsINVISIBLE As String = "INVISIBLE"
    Private Const gcsCharIsDISABLED As String = "NoUserInput"
    Private Const gcsDoNotProcess As String = "DoNotProcess"
    Private Const gcsNewline As String = "*#newline#*"

    'Math
    Private Const gcsMULTIPLICATION As String = "MULTIPLICATION"
    Private Const gcsAddition As String = "Addition"
    Private Const gcsMathValueFromArgument1 As String = "MathValueFromArgument1" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsMathValueFromArgument2 As String = "MathValueFromArgument2" & clsWordDocument.gcsCharArgumentParameterValueDelimiter

    'Functions
    Private Const gcsActionSequence As String = "ActionSequence" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsCharIsActivateFunction As String = "ActivateFunction" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsWorkDaysStart As String = "WorkDaysStart" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsWorkDaysEnd As String = "WorkDaysEnd" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gcsLoadReference As String = "LoadReference" & clsWordDocument.gcsCharArgumentParameterValueDelimiter

    'Copy
    Private Const gcsCharIsCopyFromArgument As String = "CopyFromArgument" & clsWordDocument.gcsCharArgumentParameterValueDelimiter
    Private Const gciNothing As Integer = -1

    Public Sub subLoadDocument(ByVal bvFullPath As String)

        Dim sFileName As String
        Dim sFolderNameOfFile As String

        Me.gsFullPath = bvFullPath

        sFileName = My.Computer.FileSystem.GetFileInfo(file:=Me.gsFullPath).Name
        gsFileNameWithoutExtension = sFileName.Remove(startIndex:=sFileName.Length - Form1.gcFileExtension.Length)
        sFolderNameOfFile = My.Computer.FileSystem.GetParentPath(path:=Me.gsFullPath)

        Try

            Try

                Me.gspDocument.LoadFromFile(fileName:=Me.gsFullPath)

            Catch ex As System.IO.IOException
                Dim sCopyFullPath As String
                sCopyFullPath = sFolderNameOfFile & Form1.gcTempFolderName & "\" & gsFileNameWithoutExtension & " (deletable temp)" & Form1.gcFileExtension
                My.Computer.FileSystem.CopyFile(sourceFileName:=Me.gsFullPath, destinationFileName:=sCopyFullPath, overwrite:=True)
                Me.gspDocument.LoadFromFile(fileName:=sCopyFullPath)
            Catch ex As Exception
                MsgBox("Error. Check if the selecfted file is open in another application.")
                Exit Sub
            End Try


            Me.gmsListOfArguments = fncFindArgumentsInParagraphs(msListOfArguments:=gmsListOfArguments)
            Try
                Me.gmsListOfArguments = fncFindArgumentsInTables(msListOfArguments:=gmsListOfArguments)
            Catch ex As Exception
                Stop
            End Try

        Catch ex As Exception
            MsgBox("Error. This is barebones, so the error isn't specified.")
            Exit Sub
        End Try

    End Sub


    Private Function fncFindArgumentsInParagraphs(ByRef msListOfArguments As ArrayList
                                                  ) As ArrayList

        For Each esSection As Spire.Doc.Section In gspDocument.Sections
            For Each esParagraph As Spire.Doc.Documents.Paragraph In esSection.Paragraphs

                msListOfArguments = fncParseParagraphs(bvSourceParagraph:=esParagraph,
                                                       msListOfArguments:=msListOfArguments,
                                                       bvStartLookingFrom:=0)

            Next
        Next

        Return msListOfArguments

    End Function

    Private Function fncFindArgumentsInTables(ByRef msListOfArguments As ArrayList
                                              ) As ArrayList


        For Each esSection As Spire.Doc.Section In gspDocument.Sections
            For Each esTable As Spire.Doc.Table In esSection.Tables
                For Each feRow As Spire.Doc.TableRow In esTable.Rows
                    For Each feCell As Spire.Doc.TableCell In feRow.Cells
                        For Each esParagraph As Spire.Doc.Documents.Paragraph In feCell.Paragraphs

                            Try


                                msListOfArguments = fncParseParagraphs(bvSourceParagraph:=esParagraph,
                                                                   msListOfArguments:=msListOfArguments,
                                                                   bvStartLookingFrom:=0)
                            Catch ex As Exception
                                MsgBox("error:" & esParagraph.Text)
                            End Try
                        Next
                    Next
                Next
            Next
        Next

        Return msListOfArguments

    End Function

    Private Function fncParseParagraphs(ByVal bvSourceParagraph As Spire.Doc.Documents.Paragraph,
                                        ByRef msListOfArguments As ArrayList,
                                        ByVal bvStartLookingFrom As Integer) As ArrayList

        Dim sSourceText As String = bvSourceParagraph.Text

        If Not sSourceText.Contains(value:=clsWordDocument.gcsCharArgumentStart) Then
            Return msListOfArguments
        End If


        Dim iSequenceAndKeyStart As Integer
        Dim iSequenceAndKeyEnd As Integer
        Dim iSequenceAndKeyLength As Integer
        Dim sEntireSequenceAndKeyString As String = ""
        Dim sArguments As String = ""

        iSequenceAndKeyStart = sSourceText.IndexOf(value:=clsWordDocument.gcsCharArgumentStart, startIndex:=bvStartLookingFrom)

        If iSequenceAndKeyStart = -1 Then
            Return msListOfArguments
        End If

        Try
            iSequenceAndKeyEnd = sSourceText.IndexOf(startIndex:=iSequenceAndKeyStart, value:=clsWordDocument.gcsCharArgumentEnd)
            iSequenceAndKeyLength = iSequenceAndKeyEnd - iSequenceAndKeyStart
            sEntireSequenceAndKeyString = sSourceText.Substring(startIndex:=iSequenceAndKeyStart, length:=iSequenceAndKeyLength + clsWordDocument.gcsCharArgumentEnd.Length)
            sArguments = sEntireSequenceAndKeyString.Substring(startIndex:=clsWordDocument.gcsCharArgumentStart.Length)
            sArguments = sArguments.Substring(startIndex:=0, length:=sArguments.Length - clsWordDocument.gcsCharArgumentEnd.Length)
        Catch ex As Exception
            MsgBox("Could not process an Argument in the following text:" & Environment.NewLine &
                   sSourceText & Environment.NewLine &
                   Environment.NewLine &
                   "The program in an unsteady state, please restart it.")

            Err.Raise("The program in an unsteady state, please restart it")
        End Try

        Dim iSequence As Integer
        Dim sTextToShow As String = ""
        Dim sDefaultValue As String = ""
        Dim bHasDefaultValue As Boolean = False
        Dim iIsMultiline As Boolean = False
        Dim iNumberOfLines As Integer = 2
        Dim bImportsFromReference As Boolean = False
        Dim sReferenceTextToShow As String = ""

        'Date
        Dim bIsDate As Boolean = False
        Dim bHasDateFormat As Boolean = False
        Dim sDateFormat As String = ""
        Dim iAddDays As Integer = 0
        Dim bEndOfMonth As Boolean = False
        Dim bInitalizeAtEndOfMonth As Boolean = False

        Dim bIsNumeric As Boolean = False
        Dim bThousandSeparated As Boolean = False
        Dim bIsNumericAsText As Boolean = False
        Dim iNumberMax As Integer = Integer.MaxValue
        Dim bNoValueOnEmpty As Boolean = False
        Dim bNoOutput As Boolean = False
        Dim bINVISIBLE As Boolean = False
        Dim bDISABLED As Boolean = False
        Dim bDoNotProcess As Boolean = False

        'Math
        Dim bIsMultiplication As Boolean = False
        Dim bIsAddition As Boolean = False
        Dim sMathValueFromArgument1 As String = ""
        Dim sMathValueFromArgument2 As String = ""

        'Functions
        Dim sActivateFunction As String = ""
        Dim sWorkDaysStart As String = ""
        Dim sWorkDaysEnd As String = ""

        'Automatation
        Dim sLoadReference As String = ""
        Dim bHasActionSequence As Boolean = False
        Dim iActionSequence As String = 0

        'Copy
        Dim IsCopyFromArgument As Boolean = False
        Dim CopyFromArgumentSequenceTextToShow As String = ""

        Dim msParameters As ArrayList
        msParameters = New ArrayList

        msParameters.AddRange(sArguments.Split(clsWordDocument.gcsCharArgumentParameterDelimiter))

        For Each feParameter As String In msParameters

            feParameter = feParameter.Trim

            'DoNotProcess
            If feParameter.StartsWith(clsWordDocument.gcsDoNotProcess) Then
                bDoNotProcess = True
                Exit For
            End If

            'COMMENT
            If feParameter.StartsWith(clsWordDocument.gcsCharIsCOMMENT) Then
                Continue For
            End If

            'Sequence
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharPrefixSequence,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=iSequence)

            'sTextToShow
            '“$” is the short form of “TexToShow”. If an Argument has both as Parameters, only “TexToShow” will be used. 
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharPrefixTexToShow,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=sTextToShow)

            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharPrefixTexToShowExplicit,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=sTextToShow)

            'Default
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharPrefixTextToShowDefault,
                                     brHasVariable:=bHasDefaultValue,
                                     brValueVariable:=sDefaultValue)

            If bHasDefaultValue Then
                sDefaultValue = sDefaultValue.Replace(oldValue:="*#newline#*", newValue:=Environment.NewLine)
            End If

            'Multiline
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharKeyMultiline,
                                     brHasVariable:=iIsMultiline,
                                     brValueVariable:=iNumberOfLines)

            'Reference
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharPrefixImportFromReference,
                                     brHasVariable:=bImportsFromReference,
                                     brValueVariable:=sReferenceTextToShow)

            If Not bImportsFromReference Then

                Call Me.subLoadParameter(bvParameter:=feParameter,
                                         bvKey:=clsWordDocument.gcsCharPrefixImportFromReference.Substring(startIndex:=0, length:=clsWordDocument.gcsCharPrefixImportFromReference.Length - clsWordDocument.gcsCharArgumentParameterValueDelimiter.Length),
                                         brHasVariable:=bImportsFromReference)

                'sReferenceTextToShow = sTextToShow will be added after the for loop, since it's not guaranteeded that sTextToShow was loaded at this point

            End If

            'Date
            If feParameter.ToUpper.StartsWith(clsWordDocument.gcsCharIsDate.ToUpper) Then
                bIsDate = True

                If feParameter.Length > clsWordDocument.gcsCharIsDate.Length Then
                    bHasDateFormat = True
                    sDateFormat = feParameter.Substring(startIndex:=clsWordDocument.gcsCharIsDate.Length)
                End If

            End If

            'AddDays
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsAddDays,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=iAddDays)

            'EndOfMonth
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsEndOfMonth,
                                     brHasVariable:=bEndOfMonth)

            'InitalizeAtEndOfMonth
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsInitalizeAtEndOfMonth,
                                     brHasVariable:=bInitalizeAtEndOfMonth)

            'Math
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsMULTIPLICATION,
                                     brHasVariable:=bIsMultiplication)

            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsAddition,
                                     brHasVariable:=bIsAddition)

            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsMathValueFromArgument1,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=sMathValueFromArgument1)

            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsMathValueFromArgument2,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=sMathValueFromArgument2)

            'Numeric
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharIsNumeric,
                                     brHasVariable:=bIsNumeric,
                                     brValueVariable:=iNumberMax)


            If bIsNumeric = False Then


                Call Me.subLoadParameter(bvParameter:=feParameter & clsWordDocument.gcsCharArgumentParameterValueDelimiter,
                                         bvKey:=clsWordDocument.gcsCharIsNumeric,
                                         brHasVariable:=bIsNumeric)
            End If

            'ThousandSeparated
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharThousandSeparated,
                                     brHasVariable:=bThousandSeparated)
            'IsNumericAsText
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharIsNumAsText,
                                     brHasVariable:=bIsNumericAsText,
                                     brValueVariable:=iNumberMax)

            If bIsNumeric = False Then
                Call Me.subLoadParameter(bvParameter:=feParameter & clsWordDocument.gcsCharArgumentParameterValueDelimiter,
                                         bvKey:=clsWordDocument.gcsCharIsNumAsText,
                                         brHasVariable:=bIsNumericAsText)
            End If


            'NoValueOnEmpty
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharNoValueOnEmpty,
                                     brHasVariable:=bNoValueOnEmpty)

            'NoOutput
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharNoOutput,
                                     brHasVariable:=bNoOutput)

            'INVISIBLE
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                             bvKey:=clsWordDocument.gcsCharIsINVISIBLE,
                                     brHasVariable:=bINVISIBLE)

            'DISABLED
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                            bvKey:=clsWordDocument.gcsCharIsDISABLED,
                                     brHasVariable:=bDISABLED)

            'ActivateFunction
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharIsActivateFunction,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=sActivateFunction)

            'WorkDaysStart
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsWorkDaysStart,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=sWorkDaysStart)

            'WorkDaysEnd
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsWorkDaysEnd,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=sWorkDaysEnd)

            'sLoadReference
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsLoadReference,
                                     brHasVariable:=Nothing,
                                     brValueVariable:=sLoadReference)

            'ActionSequence
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsActionSequence,
                                     brHasVariable:=bHasActionSequence,
                                     brValueVariable:=iActionSequence)

            'CopyFromArgument
            Call Me.subLoadParameter(bvParameter:=feParameter,
                                     bvKey:=clsWordDocument.gcsCharIsCopyFromArgument,
                                     brHasVariable:=IsCopyFromArgument,
                                     brValueVariable:=CopyFromArgumentSequenceTextToShow)

        Next

        'if "R$" is used instead of "R$: Mojang Employee E-mail"
        If bImportsFromReference And sReferenceTextToShow = "" Then
            sReferenceTextToShow = sTextToShow
        End If

        If Not bDoNotProcess Then
            msListOfArguments = Me.fncAddToDictionary(msListOfArguments:=msListOfArguments,
                                                      bvSequence:=iSequence,
                                                      bvKey:=sTextToShow,
                                                      bvDefaultValue:=sDefaultValue,
                                                      bvHasDefaultValue:=bHasDefaultValue,
                                                      bvIsMultiline:=iIsMultiline,
                                                      bvNumberOfLines:=iNumberOfLines,
                                                      bvImportsFromReference:=bImportsFromReference,
                                                      bvReferenceTextToShow:=sReferenceTextToShow,
                                                      bvEntireSequenceAndKeyString:=sEntireSequenceAndKeyString,
                                                      bvIsDate:=bIsDate,
                                                      bvHasDateFormat:=bHasDateFormat,
                                                      bvDateFormat:=sDateFormat,
                                                      bvAddDays:=iAddDays,
                                                      bvEndOfMonth:=bEndOfMonth,
                                                      bvInitalizeAtEndOfMonth:=bInitalizeAtEndOfMonth,
                                                      bvIsNumeric:=bIsNumeric,
                                                      bvThousandSeparated:=bThousandSeparated,
                                                      bvIsNumericAsText:=bIsNumericAsText,
                                                      bvNumerMax:=iNumberMax,
                                                      bvNoValueOnEmpty:=bNoValueOnEmpty,
                                                      bvNoOutput:=bNoOutput,
                                                      bvINVISIBLE:=bINVISIBLE,
                                                      bvDISABLED:=bDISABLED,
                                                      bvActivateFunction:=sActivateFunction,
                                                      bvWorkDaysStart:=sWorkDaysStart,
                                                      bvWorkDaysEnd:=sWorkDaysEnd,
                                                      bvIsMultiplication:=bIsMultiplication,
                                                      bvIsAddition:=bIsAddition,
                                                      bvMathValueFromArgument1:=sMathValueFromArgument1,
                                                      bvMathValueFromArgument2:=sMathValueFromArgument2,
                                                      bvLoadReference:=sLoadReference,
                                                      bvHasActionSequence:=bHasActionSequence,
                                                      bvActionSequence:=iActionSequence,
                                                      bvIsCopyFromArgument:=IsCopyFromArgument,
                                                      bvCopyFromArgumentTextToShow:=CopyFromArgumentSequenceTextToShow,
                                                      bvSourceParagraph:=bvSourceParagraph)
        End If

        'Get the rest of the keys
        msListOfArguments = Me.fncParseParagraphs(bvSourceParagraph:=bvSourceParagraph,
                                                  msListOfArguments:=msListOfArguments,
                                                  bvStartLookingFrom:=iSequenceAndKeyEnd)

        Return msListOfArguments

    End Function

    Private Function fncAddToDictionary(ByRef msListOfArguments As ArrayList,
                                        ByVal bvSequence As Integer,
                                        ByVal bvKey As String,
                                        ByVal bvDefaultValue As String,
                                        ByVal bvHasDefaultValue As Boolean,
                                        ByVal bvIsMultiline As Boolean,
                                        ByVal bvNumberOfLines As Integer,
                                        ByVal bvImportsFromReference As Boolean,
                                        ByVal bvReferenceTextToShow As String,
                                        ByVal bvEntireSequenceAndKeyString As String,
                                        ByVal bvIsDate As Boolean,
                                        ByVal bvHasDateFormat As Boolean,
                                        ByVal bvDateFormat As String,
                                        ByVal bvAddDays As Integer,
                                        ByVal bvEndOfMonth As Boolean,
                                        ByVal bvInitalizeAtEndOfMonth As Boolean,
                                        ByVal bvIsNumeric As Boolean,
                                        ByVal bvThousandSeparated As Boolean,
                                        ByVal bvIsNumericAsText As Boolean,
                                        ByVal bvNumerMax As Integer,
                                        ByVal bvNoValueOnEmpty As Boolean,
                                        ByVal bvNoOutput As Boolean,
                                        ByVal bvINVISIBLE As Boolean,
                                        ByVal bvDISABLED As Boolean,
                                        ByVal bvActivateFunction As String,
                                        ByVal bvWorkDaysStart As String,
                                        ByVal bvWorkDaysEnd As String,
                                        ByVal bvIsMultiplication As Boolean,
                                        ByVal bvIsAddition As Boolean,
                                        ByVal bvMathValueFromArgument1 As String,
                                        ByVal bvMathValueFromArgument2 As String,
                                        ByVal bvLoadReference As String,
                                        ByVal bvHasActionSequence As Boolean,
                                        ByVal bvActionSequence As Integer,
                                        ByVal bvIsCopyFromArgument As Boolean,
                                        ByVal bvCopyFromArgumentTextToShow As String,
                                        ByVal bvSourceParagraph As Spire.Doc.Documents.Paragraph
                                        ) As ArrayList

        Dim bNoSequenceWasProvided As Boolean

        If bvSequence = clsWordDocument.gciNothing Then
            bNoSequenceWasProvided = True
        Else
            bNoSequenceWasProvided = False
        End If

        ''If the provided sequense is already taken, put it into the next one.
        'Do While Me.fncIsSequenceTaken(brDictionaryOfKeys:=msListOfArguments, _
        '                               bvSequence:=bvSequence)
        '    bvSequence += 1
        'Loop


        'Create the holder for the key
        Dim ymArgument As clsArgument
        ymArgument = New clsArgument

        ymArgument.subSetParagraphFromSource(bvSourceParagraph:=bvSourceParagraph)

        ymArgument.Sequence = bvSequence
        ymArgument.TextToShow = bvKey
        ymArgument.DefaultValue = bvDefaultValue
        ymArgument.HasDefaultValue = bvHasDefaultValue
        ymArgument.IsMultiline = bvIsMultiline
        ymArgument.NumberOfLines = bvNumberOfLines
        ymArgument.ImportsFromReference = bvImportsFromReference
        ymArgument.ReferenceTextToShow = bvReferenceTextToShow
        ymArgument.ArgumentAsIsInDocument = bvEntireSequenceAndKeyString
        ymArgument.NoSequenceWasProvided = bNoSequenceWasProvided
        ymArgument.IsDate_ = bvIsDate
        ymArgument.HasDateFormat = bvHasDateFormat
        ymArgument.DateFormat = bvDateFormat
        ymArgument.AddDays = bvAddDays
        ymArgument.EndOfMonth = bvEndOfMonth
        ymArgument.InitalizeAtEndOfMonth = bvInitalizeAtEndOfMonth
        ymArgument.IsNumeric = bvIsNumeric
        ymArgument.ThousandSeparated = bvThousandSeparated
        ymArgument.IsNumericAsText = bvIsNumericAsText
        ymArgument.NumerMax = bvNumerMax
        ymArgument.NoValueOnEmpty = bvNoValueOnEmpty
        ymArgument.NoOutput = bvNoOutput
        ymArgument.INVISIBLE = bvINVISIBLE
        ymArgument.DISABLED = bvDISABLED
        ymArgument.ActivateFunction = bvActivateFunction
        ymArgument.WorkDaysStart = bvWorkDaysStart
        ymArgument.WorkDaysEnd = bvWorkDaysEnd

        ymArgument.IsMultiplication = bvIsMultiplication
        ymArgument.IsAddition = bvIsAddition
        ymArgument.MathValueFromArgument1 = bvMathValueFromArgument1
        ymArgument.MathValueFromArgument2 = bvMathValueFromArgument2

        ymArgument.LoadReference = bvLoadReference
        ymArgument.HasActionSequence = bvHasActionSequence
        ymArgument.ActionSequence = bvActionSequence

        ymArgument.IsCopyFromArgument = bvIsCopyFromArgument
        ymArgument.CopyFromArgumentTexToShow = bvCopyFromArgumentTextToShow

        'Put the value in sequential order
        If msListOfArguments.Count = 0 Then
            msListOfArguments.Add(value:=ymArgument)
        Else
            Dim iIndex As Integer = 0

            For Each feArgument As clsArgument In msListOfArguments
                If feArgument.Sequence > bvSequence Then
                    Exit For
                Else
                    iIndex += 1
                End If
            Next

            msListOfArguments.Insert(value:=ymArgument, index:=iIndex)

        End If

        Return msListOfArguments

    End Function

    Private Function fncIsSequenceTaken(ByRef brDictionaryOfKeys As ArrayList,
                                            ByVal bvSequence As Integer
                                            ) As Boolean

        For Each feSequenceKeyHolder As clsArgument In brDictionaryOfKeys
            If feSequenceKeyHolder.Sequence = bvSequence Then
                Return True
            End If
        Next

        Return False

    End Function

    Public Sub subSaveToFile(ByVal bvNewFullPath As String)

        Me.gspDocument.SaveToFile(bvNewFullPath, FileFormat.Docx)

    End Sub

    Private Sub subLoadParameter(ByVal bvParameter As String, _
                                 ByVal bvKey As String, _
                                 ByRef brHasVariable As Boolean, _
                                 ByRef brValueVariable As String)

        'brVariable has a default value, so it is received by ref in order to maintain its default value and change it only if requried.
        'The alternative is to return "" when the IF results in false, and that would overwrite the default value, and correcting for that would generate complex code.

        If bvParameter.Trim.ToUpper.StartsWith(bvKey.ToUpper) Then
            brHasVariable = True
            brValueVariable = bvParameter.Substring(startIndex:=bvKey.Length)
        End If

    End Sub

    Private Sub subLoadParameter(ByVal bvParameter As String, _
                                 ByVal bvKey As String, _
                                 ByRef brHasVariable As Boolean)


        'brVariable has a default value, so it is received by ref in order to maintain its default value and change it only if requried.
        'The alternative is to return "" when the IF results in false, and that would overwrite the default value, and correcting for that would generate complex code.

        If bvParameter.Trim.ToUpper.StartsWith(bvKey.ToUpper) Then
            brHasVariable = True
        End If

    End Sub



End Class
