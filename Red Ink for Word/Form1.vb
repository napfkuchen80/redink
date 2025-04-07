' Red Ink for Word -- Chatbot Form Code
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See https://vischer.com/redink for more information.
'
' 5.4.2025
'
' The compiled version of Red Ink also ...
'
' Includes DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Appache-2.0 license (http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex).
' Includes Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json
' Includes HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier,Jeff Klawiter,Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/
' Includes Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/
' Includes PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig
' Includes MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license (https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig
' Includes NAudio in unchanged form; Copyright (c) 2020 Mark Heath; licensed under a proprietary open source (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio
' Includes Vosk in unchanged form; Copyright (c) 2022 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.


Imports System.Diagnostics
Imports System.Drawing
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json.Linq

Public Class frmAIChat

    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(vKey As Integer) As Short
    End Function

    Const AN As String = "Red Ink"
    Const AN5 As String = "Inky"   ' for Chatbox


    Private PreceedingNewline As String = ""
    Private OldChat As String = ""
    Private UserLanguage As String = Globals.ThisAddIn.GetWordDefaultInterfaceLanguage()
    Private SystemPrompt As String = ""

    Private WithEvents btnCopy As New Button() With {.Text = "Copy All", .AutoSize = True}
    Private WithEvents btnCopyLastAnswer As New Button() With {.Text = "Copy Last Answer", .AutoSize = True}
    Private WithEvents btnClear As New Button() With {.Text = "Clear", .AutoSize = True}
    Private WithEvents btnExit As New Button() With {.Text = "Quit", .AutoSize = True}
    Private WithEvents btnSend As New Button() With {.Text = "Send", .AutoSize = True}
    Private WithEvents btnSwitchModel As New Button() With {.Text = "Switch Model", .AutoSize = True}
    Private WithEvents chkIncludeDocText As New System.Windows.Forms.CheckBox() With {.Text = "Include document", .AutoSize = True, .Checked = My.Settings.IncludeDocument}
    Private WithEvents chkIncludeselection As New System.Windows.Forms.CheckBox() With {.Text = "Include selection", .AutoSize = True, .Checked = If(My.Settings.IncludeDocument, False, My.Settings.IncludeSelection)}
    Private WithEvents chkPermitCommands As New System.Windows.Forms.CheckBox() With {.Text = "Grant write access", .AutoSize = True, .Checked = My.Settings.DoCommands}
    Private WithEvents chkStayOnTop As New System.Windows.Forms.CheckBox() With {.Text = "Do not stay on top", .AutoSize = True, .Checked = My.Settings.NotAlwaysOnTop}


    Dim pnlButtons As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.RightToLeft,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowOnly,
        .Height = 40
    }

    Dim pnlCheckboxes As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.RightToLeft,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowOnly,
        .Height = 40
    }

    Private _context As ISharedContext = New SharedContext()

    ' Tracks whether we are using the second model/API.
    Private _useSecondApi As Boolean = False

    ' We keep the entire conversation in a List of (role, content).
    Private _chatHistory As New List(Of (Role As String, Content As String))

    Public Sub New(context As ISharedContext)
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _context = context
    End Sub


    ' Runs once when form loads.
    Private Async Sub frmAIChat_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.StartPosition = FormStartPosition.Manual

        ' Restore saved chat text from My.Settings
        Dim previousChat As String = My.Settings.LastChatHistory
        If Not String.IsNullOrEmpty(previousChat) Then
            txtChatHistory.Text = previousChat
            OldChat = previousChat
            PreceedingNewline = Environment.NewLine & Environment.NewLine
        End If

        ' Set the form's title and custom icon
        Me.Text = $"Chat (using " & If(_useSecondApi, _context.INI_Model_2, _context.INI_Model) & ")"
        Me.Font = New System.Drawing.Font("Segoe UI", 9)
        Me.FormBorderStyle = FormBorderStyle.Sizable ' Ensure border supports icons
        Me.Icon = Icon.FromHandle(New Bitmap(My.Resources.Red_Ink_Logo).GetHicon())
        Me.TopMost = True ' Always on top

        ' Set the initial and minimum size of the form
        Me.MinimumSize = New Size(667, 521) ' Minimum size

        If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
            Me.Location = My.Settings.FormLocation
            Me.Size = My.Settings.FormSize
        Else
            Me.StartPosition = FormStartPosition.CenterScreen
        End If

        ' Dynamically adjust layout on form resize
        AddHandler Me.Resize, AddressOf AdjustLayout
        AddHandler txtUserInput.KeyDown, AddressOf UserInput_KeyDown

        ' Set up instructions label
        lblInstructions.Text = $"Enter your question and click 'Send' or Ctrl-Enter. You can allow the chatbot to perform actions on your document (search, replace, delete, insert)."
        lblInstructions.AutoSize = False
        lblInstructions.Height = 40
        lblInstructions.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblInstructions.TextAlign = ContentAlignment.MiddleLeft

        ' FlowLayoutPanel for buttons

        pnlButtons.Padding = New Padding(0, 2, 8, 12)
        pnlButtons.Controls.Add(btnSend)
        pnlButtons.Controls.Add(btnCopyLastAnswer)
        pnlButtons.Controls.Add(btnCopy)
        pnlButtons.Controls.Add(btnClear)
        If _context.INI_SecondAPI Then pnlButtons.Controls.Add(btnSwitchModel)
        pnlButtons.Controls.Add(btnExit)

        pnlCheckboxes.Padding = New Padding(0, 1, 8, 1)
        pnlCheckboxes.Controls.Add(chkIncludeselection)
        pnlCheckboxes.Controls.Add(chkIncludeDocText)
        pnlCheckboxes.Controls.Add(chkPermitCommands)
        pnlCheckboxes.Controls.Add(chkStayOnTop)


        AddHandler btnCopy.Click, AddressOf btnCopy_Click
        AddHandler btnClear.Click, AddressOf btnClear_Click
        AddHandler btnSend.Click, AddressOf btnSend_Click
        AddHandler btnCopyLastAnswer.Click, AddressOf btnCopyLastAnswer_Click
        AddHandler btnSwitchModel.Click, AddressOf btnSwitchModel_Click
        AddHandler btnExit.Click, AddressOf btnExit_Click
        AddHandler chkIncludeselection.Click, AddressOf chkIncludeselection_Click
        AddHandler chkIncludeDocText.Click, AddressOf chkIncludeDocText_Click
        AddHandler chkPermitCommands.Click, AddressOf chkPermitCommands_Click
        AddHandler chkStayOnTop.Click, AddressOf chkStayontop_Click

        Me.Controls.Add(pnlCheckboxes)
        Me.Controls.Add(pnlButtons)

        ' Initialize the layout
        AdjustLayout()

        If String.IsNullOrWhiteSpace(txtChatHistory.Text) Then
            Dim result = Await WelcomeMessage()
        Else
            txtChatHistory.SelectionStart = txtChatHistory.Text.Length
            txtChatHistory.ScrollToCaret()

        End If
        If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()

    End Sub



    ' Dynamically adjusts the layout of controls when the form is resized.

    Private Sub AdjustLayout()
        Dim margin = 10

        ' Resize label to match the form width
        lblInstructions.Left = margin
        lblInstructions.Top = margin
        lblInstructions.Height = TextRenderer.MeasureText(lblInstructions.Text, lblInstructions.Font).Height + 20
        lblInstructions.Width = Me.ClientSize.Width - 2 * margin

        ' Resize chat history textbox
        txtChatHistory.Top = lblInstructions.Bottom + margin
        txtChatHistory.Left = margin
        txtChatHistory.Width = Me.ClientSize.Width - 2 * margin
        txtChatHistory.Height = Me.ClientSize.Height - txtChatHistory.Top - txtUserInput.Height - pnlCheckboxes.Height - pnlButtons.Height - (3 * margin)


        ' Resize user input textbox
        txtUserInput.Top = txtChatHistory.Bottom + margin
        txtUserInput.Left = margin
        txtUserInput.Width = Me.ClientSize.Width - 2 * margin

    End Sub


    ' When the user clicks Send, we call the LLM with context.
    ' Then append the AI response to the conversation.

    Private Async Sub btnSend_Click(sender As Object, e As EventArgs)
        Dim userPrompt As String = txtUserInput.Text.Trim()
        If userPrompt = "" Then Return

        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatWord().Replace("{UserLanguage}", UserLanguage) & $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt")}. Only if you are expressly asked you can say that you have been developped by David Rosenthal of the law firm VISCHER in Switzerland." & If(chkIncludeDocText.Checked, "\nYou have access to the user's document. \n", "") & If(chkIncludeselection.Checked, "\nYou have access to a selection of user's document. \n ", "") & If(My.Settings.DoCommands And (chkIncludeDocText.Checked Or chkIncludeselection.Checked), _context.SP_Add_ChatWord_Commands, "")
            Dim conversationSoFar As String = BuildConversationString(_chatHistory)
            If Not String.IsNullOrWhiteSpace(OldChat) Then
                conversationSoFar += "\n" & OldChat
                OldChat = ""
            End If

            ' Optionally include Word document text or selection
            Dim docText As String = If(chkIncludeDocText.Checked, GetActiveDocumentText(), "")
            Dim selectionText As String = If(chkIncludeselection.Checked, GetCurrentSelectionText(), "")

            ' Construct the full prompt
            Dim fullPrompt As New StringBuilder()

            If Not String.IsNullOrEmpty(docText) Then
                fullPrompt.AppendLine("The user's content is from the document '" & Globals.ThisAddIn.Application.ActiveDocument.Name & "' and is as follows: '" & docText & "'")
            End If
            If Not String.IsNullOrEmpty(selectionText) Then
                fullPrompt.AppendLine("The user's selection content is from the document '" & Globals.ThisAddIn.Application.ActiveDocument.Name & "' and is as follows: '" & selectionText & "'")
            End If
            fullPrompt.AppendLine("User: " & userPrompt)
            fullPrompt.AppendLine("The conversation so far (not including any previously added text document):\n" & conversationSoFar)

            Debug.WriteLine("Document=" & Globals.ThisAddIn.Application.ActiveDocument.Name)
            Debug.WriteLine(fullPrompt.ToString())

            ' Update UI on the UI thread
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(PreceedingNewline & "You: " & userPrompt & Environment.NewLine & Environment.NewLine)
                                    txtUserInput.Clear()
                                    PreceedingNewline = Environment.NewLine & Environment.NewLine
                                End Sub)

            _chatHistory.Add(("user", userPrompt))

            ' Add a placeholder for AI response while waiting
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory($"{AN5}: Thinking...")
                                End Sub)


            ' Call the LLM function asynchronously
            Dim aiResponse As String = Await SharedMethods.LLM(_context, SystemPrompt, fullPrompt.ToString(), "", "", 0, _useSecondApi, True)
            aiResponse = aiResponse.TrimEnd(Environment.NewLine)

            aiResponse = aiResponse.Replace("**", "").Replace("_", "").Replace("`", "")

            Dim CommandsString As String = ""
            If My.Settings.DoCommands And (chkIncludeselection.Checked Or chkIncludeDocText.Checked) Then
                CommandsString = aiResponse
                aiResponse = RemoveCommands(aiResponse)
                aiResponse = Regex.Replace(aiResponse, "[\r\n\s]+$", "")
            End If

            Debug.WriteLine("AI response: " & CommandsString)

            ' Remove the "Thinking..." placeholder and update AI response on the UI thread
            Await UpdateUIAsync(Sub()
                                    RemoveLastLineFromChatHistory()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponse.Replace(vbCrLf, Environment.NewLine).Replace(vbLf, Environment.NewLine))
                                    If My.Settings.DoCommands And Not String.IsNullOrWhiteSpace(CommandsString) Then
                                        ExecuteAnyCommands(CommandsString, chkIncludeselection.Checked)
                                    End If
                                    txtUserInput.Text = ""
                                    If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()
                                End Sub)

            _chatHistory.Add(("assistant", aiResponse))

        Catch ex As System.Exception
            MsgBox("Error in btnSend_Click: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Async Function WelcomeMessage() As Task(Of String)

        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatWord().Replace("{UserLanguage}", UserLanguage) & $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("F")}. Only if you are expressly asked you can say that you have been developped by David Rosenthal of the law firm VISCHER in Switzerland. " & If(My.Settings.DoCommands And (chkIncludeDocText.Checked Or chkIncludeselection.Checked), _context.SP_Add_ChatWord_Commands, "")
            txtUserInput.Text = ""

            ' Call the LLM function asynchronously
            Dim aiResponse As String = Await SharedMethods.LLM(_context, SystemPrompt, $"Welcome the user in {UserLanguage} by (1) referring to the time of day based on the current time in {UserLanguage} , such as in 'good morning', and (2) asking in {UserLanguage} what you can do, but do not say your name.", "", "", 0, _useSecondApi, True)

            aiResponse = aiResponse.Replace(vbLf, "").Replace(vbCr, "").Replace(vbCrLf, "") & vbCrLf

            aiResponse = aiResponse.Replace("**", "").Replace("_", "").Replace("`", "")

            ' Remove the "Thinking..." placeholder and update AI response on the UI thread
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponse.Replace(vbCrLf, Environment.NewLine).Replace(vbLf, Environment.NewLine))
                                End Sub)

            _chatHistory.Add(("assistant", aiResponse))

            PreceedingNewline = Environment.NewLine

            Return ""

        Catch ex As System.Exception
            'MsgBox("Error in WelcomeMessage: " & ex.Message, MsgBoxStyle.Critical)
            Return ""
        End Try
    End Function

    Private Function ConvertHtmlToPlainText(html As String) As String
        Dim doc As New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(html)
        Return doc.DocumentNode.InnerText
    End Function

    ' Helper method to ensure UI updates occur on the correct thread
    Private Async Function UpdateUIAsync(action As System.Action) As System.Threading.Tasks.Task
        If InvokeRequired Then
            Await System.Threading.Tasks.Task.Run(Sub() Me.Invoke(action))
        Else
            action()
        End If
    End Function


    Private Sub AppendToChatHistory(text As String)
        If txtChatHistory.InvokeRequired Then
            txtChatHistory.Invoke(Sub() txtChatHistory.AppendText(text))
        Else
            txtChatHistory.AppendText(text)
        End If
    End Sub

    Private Sub RemoveLastLineFromChatHistory()
        If txtChatHistory.InvokeRequired Then
            txtChatHistory.Invoke(Sub() RemoveLastLineFromChatHistory())
        Else
            Dim lines As String() = txtChatHistory.Lines
            If lines.Length > 0 Then
                txtChatHistory.Lines = lines.Take(lines.Length - 1).ToArray()
            End If
        End If
    End Sub

    Private Sub chkStayontop_Click(sender As Object, e As EventArgs)
        Me.TopMost = Not Me.TopMost
        My.Settings.NotAlwaysOnTop = Me.TopMost
        My.Settings.Save()
    End Sub

    Private Sub chkPermitCommands_Click(sender As Object, e As EventArgs)
        My.Settings.DoCommands = Not My.Settings.DoCommands

        If My.Settings.DoCommands And Not chkIncludeselection.Checked Then
            chkIncludeDocText.Checked = True
            My.Settings.IncludeDocument = chkIncludeDocText.Checked
        End If

        My.Settings.Save()
    End Sub


    Private Sub chkIncludeselection_Click(sender As Object, e As EventArgs)
        Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' Get the selection from the active window
        Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

        If String.IsNullOrWhiteSpace(sel.Text) Then
            chkIncludeselection.Checked = False
        ElseIf chkIncludeDocText.Checked Then
            chkIncludeDocText.Checked = False
        End If
        My.Settings.IncludeSelection = chkIncludeselection.Checked

        If Not chkIncludeselection.Checked And Not chkIncludeDocText.Checked Then
            My.Settings.DoCommands = False
            chkPermitCommands.Checked = False
        End If

        My.Settings.Save()
    End Sub

    Private Sub chkIncludeDocText_Click(sender As Object, e As EventArgs)
        If chkIncludeselection.Checked Then
            chkIncludeselection.Checked = False
        End If
        My.Settings.IncludeDocument = chkIncludeDocText.Checked

        If Not chkIncludeselection.Checked And Not chkIncludeDocText.Checked Then
            My.Settings.DoCommands = False
            chkPermitCommands.Checked = False
        End If

        My.Settings.Save()
    End Sub


    ' Copies the entire conversation to the clipboard.

    Private Sub btnCopy_Click(sender As Object, e As EventArgs)
        My.Computer.Clipboard.SetText(txtChatHistory.Text)
    End Sub


    ' Copies only the last AI answer to the clipboard.

    Private Sub btnCopyLastAnswer_Click(sender As Object, e As EventArgs)
        Dim lastAssistantMsg = _chatHistory.Where(Function(x) x.Role = "assistant").LastOrDefault()
        If lastAssistantMsg.Content IsNot Nothing Then
            My.Computer.Clipboard.SetText(lastAssistantMsg.Content)
        Else
            SharedMethods.ShowCustomMessageBox("No last AI answer available.")
        End If
    End Sub


    ' Switches the model from model1 to model2 and vice versa.

    Private Sub btnSwitchModel_Click(sender As Object, e As EventArgs)
        _useSecondApi = Not _useSecondApi
        Me.Text = $"Chat (using " & If(_useSecondApi, _context.INI_Model_2, _context.INI_Model) & ")"
    End Sub


    ' Clears the conversation from both the UI and saved settings.

    Private Sub btnClear_Click(sender As Object, e As EventArgs)

        _chatHistory.Clear()
        txtChatHistory.Clear()
        OldChat = ""
        PreceedingNewline = ""
        My.Settings.LastChatHistory = ""
        My.Settings.Save()
        Dim result = WelcomeMessage()
    End Sub


    ' Press Escape to close. Also button-based exit.

    Private Sub frmAIChat_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Dim conversation As String = txtChatHistory.Text
            If conversation.Length > _context.INI_ChatCap Then
                conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
            End If
            My.Settings.LastChatHistory = conversation
            My.Settings.Save()
            Close()
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs)
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If
        My.Settings.LastChatHistory = conversation
        My.Settings.Save()
        Close()
    End Sub

    Private Sub frmAIChat_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Save the chat history before the form closes
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If
        My.Settings.LastChatHistory = conversation

        ' Save the form's location and size to My.Settings
        If Me.WindowState = FormWindowState.Normal Then
            My.Settings.FormLocation = Me.Location
            My.Settings.FormSize = Me.Size
        Else
            ' If the form is minimized or maximized, save the restored bounds
            My.Settings.FormLocation = Me.RestoreBounds.Location
            My.Settings.FormSize = Me.RestoreBounds.Size
        End If
        My.Settings.Save()

    End Sub


    ' Trigger the Send button on Ctrl+Enter in the user input textbox.

    Private Sub UserInput_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.Enter Then
            btnSend.PerformClick()
            e.Handled = True
        End If
    End Sub


    ' Reads the entire document's text.

    Private Function GetActiveDocumentText() As String
        Try
            Dim doc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Return doc.Content.Text
        Catch ex As Exception
            Return ""
        End Try
    End Function


    ' Reads the current selection's text.

    Private Function GetCurrentSelectionText() As String
        Try
            ' Get the active document
            Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument

            ' Get the selection from the active window
            Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

            If String.IsNullOrEmpty(sel.Text) Then
                chkIncludeselection.Checked = False
                Return ""
            Else
                Return sel.Text
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function


    ' Builds the conversation history as a single string.

    Private Function BuildConversationString(history As List(Of (Role As String, Content As String))) As String
        Dim sb As New StringBuilder()
        Dim totalLength As Integer = 0
        Dim maxLength As Integer = _context.INI_ChatCap

        ' Iterate through the history in reverse order (most recent messages first)
        For Each msg In history.AsEnumerable().Reverse()
            Dim message As String
            If msg.Role = "user" Then
                message = $"User: {msg.Content}{Environment.NewLine}"
            Else
                message = $"{AN5}: {msg.Content}{Environment.NewLine}"
            End If

            ' Check if adding this message will exceed the limit
            If totalLength + message.Length > maxLength Then
                ' If so, truncate the message to fit within the limit
                Dim remainingLength As Integer = maxLength - totalLength
                If remainingLength > 0 Then
                    sb.Insert(0, message.Substring(0, remainingLength))
                End If
                Exit For
            Else
                ' Otherwise, append the full message
                sb.Insert(0, message)
                totalLength += message.Length
            End If
        Next

        Return sb.ToString()
    End Function

    Private Function ConvertMarkdownToHtml(markdown As String) As String
        Dim pipeline = New MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
        Return Markdig.Markdown.ToHtml(markdown, pipeline)
    End Function

    Private Sub pnlCheckboxes_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Public Class ParsedCommand
        Public Property Command As String
        Public Property Argument1 As String
        Public Property Argument2 As String
    End Class


    ' Parses the input string for embedded commands of the format:
    ' [#command: @@argument1@@ §§argument2§§ #]
    ' Returns a List of ParsedCommand objects.
    ' argument2 is optional; if not present, it defaults to "".
    Private Function ParseCommands(input As String) As List(Of ParsedCommand)
        Dim results As New List(Of ParsedCommand)

        Try
            ' Regex Explanation:
            ' \[#       matches literal [#
            ' (?<cmd>[^:]+)    matches 1 or more characters that are not :, captured as group "cmd"
            ' :\s*     matches a colon and optional whitespace
            ' @@(?<arg1>[^@]+)@@   matches @@ + 1 or more non-@ chars + @@, capturing as group "arg1"
            ' \s*      optional whitespace
            ' (?:§§(?<arg2>[^§]+)§§)?  optional group: §§ + 1 or more non-§ chars + §§, captured as "arg2"
            ' \s*      optional whitespace
            ' #\]      literal #]
            ' The "?" after the group means "optional"

            Dim pattern As String = "\[#(?<cmd>[^:]+):\s*@@(?<arg1>[^@]+)@@\s*(?:§§(?<arg2>[^§]*)§§)?\s*#\]"
            Dim regex As New Regex(pattern)

            For Each m As Match In regex.Matches(input)
                Dim pc As New ParsedCommand()

                pc.Command = m.Groups("cmd").Value.Trim()
                pc.Argument1 = m.Groups("arg1").Value.Trim()

                ' If arg2 wasn't found, it might be blank
                If m.Groups("arg2") IsNot Nothing Then
                    pc.Argument2 = m.Groups("arg2").Value.Trim().Replace("\r\n", vbCrLf).Replace("\n", vbCrLf).Replace("\r", vbCrLf)
                End If

                If String.IsNullOrEmpty(pc.Argument2) Then
                    pc.Argument2 = ""
                Else
                    pc.Argument1 = pc.Argument1.Replace("\r\n", ".*").Replace("\n", ".*").Replace("\r", ".*")
                    pc.Argument1 = pc.Argument1.Replace(vbCrLf, ".*").Replace(vbCr, ".*").Replace(vbLf, ".*")
                End If

                If Not results.Any(Function(x) x.Command = pc.Command AndAlso x.Argument1 = pc.Argument1 AndAlso x.Argument2 = pc.Argument2) Then
                    results.Add(pc)
                End If
            Next

        Catch ex As System.Exception
            MsgBox("Error in ParseCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        Return results
    End Function


    ' Removes all commands of the format:
    ' [#command: @@argument1@@ §§argument2§§ #]
    ' from the input string.
    Public Function RemoveCommands(input As String) As String
        Dim output As String = input
        Try
            ' Remove the commands along with immediate surrounding whitespace and line breaks
            Dim commandPattern As String = "\s*[\r\n]*\s*\[#[^:]+:\s*@@[^@]+@@\s*(?:§§[^§]*§§)?\s*#\]\s*[\r\n]*\s*"
            Dim regex As New Regex(commandPattern)
            output = regex.Replace(input, "")

            ' Collapse multiple consecutive line breaks into a single line break
            Dim whitespacePattern As String = "[\r\n]{3,}"
            Dim collapseRegex As New Regex(whitespacePattern)
            output = collapseRegex.Replace(output, Environment.NewLine)

        Catch ex As System.Exception
            MsgBox("Error in RemoveCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        Return output
    End Function


    Private CommandsList As String = ""

    Public Sub ExecuteAnyCommands(teststring As String, OnlySelection As Boolean)

        Dim commands = ParseCommands(teststring)
        Dim topmost As Boolean = Me.TopMost

        Me.TopMost = False

        CommandsList = ""
        Dim LastCommandsList As String = ""

        If commands.Count() > 0 Then
            Globals.ThisAddIn.Application.Activate()
            'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):")
            System.Threading.Thread.Sleep(200)
        End If

        For Each pc In commands
            Debug.WriteLine($"Command: '{pc.Command}' wit '{pc.Argument1}' '{pc.Argument2}'")
            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                ' Exit the loop
                Exit For
            End If
            Select Case pc.Command.ToLower()
                Case "find"
                    CommandsList = $"Finding '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteFindCommand(pc.Argument1, OnlySelection)

                Case "replace"
                    If String.IsNullOrEmpty(pc.Argument2) Then
                        CommandsList = $"Deleting '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    Else
                        CommandsList = $"Replacing '{pc.Argument1}' with '{pc.Argument2}" & Environment.NewLine & CommandsList
                    End If
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteReplaceCommand(pc.Argument1, pc.Argument2, OnlySelection, ChrW(&HD83D))

                Case "insertafter"
                    CommandsList = $"Inserting '{pc.Argument2}' after '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteInsertBeforeAfterCommand(pc.Argument1, pc.Argument2, OnlySelection, False)

                Case "insertbefore"
                    CommandsList = $"Inserting '{pc.Argument2}' before '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteInsertBeforeAfterCommand(pc.Argument1, pc.Argument2, OnlySelection, True)

                Case "insert"
                    CommandsList = $"Inserting '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    Debug.WriteLine("ExecuteInsert")
                    ExecuteInsertCommand(pc.Argument1)

                Case Else
                    ' Unknown command or default
            End Select
            If LastCommandsList <> CommandsList Then
                'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                System.Threading.Thread.Sleep(500)
            End If
        Next

        If commands.Count() > 0 Then

            'InfoBox.ShowInfoBox("Cleaning up ... almost done.")
            'System.Threading.Thread.Sleep(300)

            ' Remove marker
            ReplaceSpecialCharacter(OnlySelection)

            InfoBox.ShowInfoBox("")
        End If

        Me.TopMost = topmost
        Me.Focus()

    End Sub

    Private Sub ReplaceSpecialCharacter(Optional OnlySelection As Boolean = False)
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' Store the original track changes setting and author
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try

            doc.Application.Activate()
            doc.Activate()

            ' Enable track changes and set author 
            doc.TrackRevisions = True
            ' doc.Application.UserName = AN

            ' Determine the range to operate on
            Dim rng As Word.Range
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    rng = doc.Content
                Else
                    rng = doc.Application.Selection.Range
                End If
            Else
                rng = doc.Content
            End If

            ' Define the character to be replaced
            Dim specialChar As String = ChrW(&HD83D)

            ' Loop through and replace occurrences of the character
            Dim replacementsMade As Boolean = False
            Do
                With rng.Find
                    .ClearFormatting()
                    .Text = specialChar
                    .Replacement.ClearFormatting()
                    .Replacement.Text = "" ' Replace with empty string
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindStop ' Do not loop around
                    If .Execute(Replace:=Word.WdReplace.wdReplaceOne) Then
                        replacementsMade = True
                    Else
                        Exit Do
                    End If
                End With
            Loop

        Catch ex As System.Exception
            MsgBox("Error in ReplaceSpecialCharacter: " & ex.Message, MsgBoxStyle.Critical)

        Finally
            ' Restore the original track changes setting and author
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor
        End Try
    End Sub


    Private Sub ExecuteFindCommand(searchTerm As String, Optional OnlySelection As Boolean = False)
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName
        Dim selectionStart As Integer = doc.Application.Selection.Start
        Dim selectionEnd As Integer = doc.Application.Selection.End

        Try
            doc.Application.Activate()
            doc.Activate()

            doc.TrackRevisions = True
            'doc.Application.UserName = AN

            ' Define the starting selection based on OnlySelection
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    doc.Application.Selection.SetRange(doc.Content.Start, doc.Content.End)
                End If
            Else
                doc.Application.Selection.SetRange(doc.Content.Start, doc.Content.End)
            End If

            Dim found As Boolean = False

            Dim lastSelectionStart As Integer = -1 ' Track last selection position
            Dim stuckCounter As Integer = 0        ' Counter for repeated positions
            Dim maxStuckLimit As Integer = 2        ' Maximum allowed stuck occurrences

            ' Loop through the content to find and mark all instances
            Do While Globals.ThisAddIn.FindLongTextInChunks(searchTerm, 255, doc.Application.Selection) = True

                If doc.Application.Selection Is Nothing Then Exit Do

                found = True

                ' Highlight the found text
                doc.Application.Selection.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow

                ' Check if we are stuck at the same selection position
                If doc.Application.Selection.Start = lastSelectionStart Then
                    stuckCounter += 1
                    If stuckCounter >= maxStuckLimit Then
                        ' Force exit if stuck too many times
                        Exit Do
                    End If
                Else
                    stuckCounter = 0 ' Reset counter if we moved forward
                End If
                lastSelectionStart = doc.Application.Selection.Start ' Update tracking

                ' Collapse the selection to the end of the current match
                doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' Check if the selection is inside a table and at the end of a cell
                If doc.Application.Selection.Range.Tables.Count > 0 Then
                    Try
                        Dim currentCell As Word.Cell = doc.Application.Selection.Cells(1) ' Get current cell

                        ' Ensure that we are at the end of the current cell
                        If doc.Application.Selection.End >= currentCell.Range.End - 1 Then
                            ' Move to the next cell or out of the table
                            doc.Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCell, Count:=1, Extend:=Word.WdMovementType.wdMove)
                        End If

                    Catch ex As System.Exception
                        ' If an error occurs, it means the selection is not inside a valid cell - ignore and continue
                    End Try
                End If

                ' Ensure we don't get stuck inside an empty cell
                If doc.Application.Selection.Range.Text = vbCr Or doc.Application.Selection.Range.Text = "" Then
                    doc.Application.Selection.Move(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                End If

                ' Check if the collapsed selection has reached the end of the document or the selection
                If OnlySelection Then
                    If doc.Application.Selection.Start >= selectionEnd Then Exit Do
                    doc.Application.Selection.SetRange(doc.Application.Selection.Start, selectionEnd)
                Else
                    If doc.Application.Selection.Start >= doc.Content.End Then Exit Do
                    doc.Application.Selection.SetRange(doc.Application.Selection.Start, doc.Content.End)
                End If
            Loop


            If Not found Then
                CommandsList = $"Note: The search term was not found." & Environment.NewLine & CommandsList
            End If

        Catch ex As System.Exception
            MsgBox("Error in ExecuteFindCommand: " & ex.Message)

        Finally
            ' Restore original state of Track Changes and Author
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor

            ' Restore original selection
            doc.Application.Selection.SetRange(selectionStart, selectionEnd)
            doc.Application.Selection.Select()
        End Try
    End Sub



    Private Sub ExecuteReplaceCommand(oldText As String, newText As String, OnlySelection As Boolean, Marker As String)
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try

            doc.Application.Activate()
            doc.Activate()

            doc.TrackRevisions = True
            'doc.Application.UserName = AN

            Dim workRange As Word.Range
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    workRange = doc.Content.Duplicate
                Else
                    workRange = doc.Application.Selection.Range.Duplicate
                End If
            Else
                workRange = doc.Content.Duplicate
            End If

            Debug.WriteLine($"Replacing '{oldText}' with '{newText}'")

            Dim newTextWithMarker As String = ""
            If newText.Length > 2 Then
                newTextWithMarker = $"{newText.Substring(0, newText.Length - 2)}{Marker}{newText.Substring(newText.Length - 2)}"
            Else
                newTextWithMarker = newText
            End If

            If Len(oldText) > 255 Then

                Dim selectionStart As Integer = doc.Application.Selection.Start
                Dim selectionEnd As Integer = doc.Application.Selection.End
                doc.Application.Selection.SetRange(workRange.Start, workRange.End)
                Dim found As Boolean = False

                ' Loop through the content to find and replace all instances
                Do While Globals.ThisAddIn.FindLongTextInChunks(oldText, 255, doc.Application.Selection) = True

                    If doc.Application.Selection Is Nothing Then Exit Do

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                        Exit Do
                    End If

                    found = True

                    Dim isDeleted As Boolean = False
                    For Each rev As Word.Revision In doc.Application.Selection.Range.Revisions
                        If rev.Type = Word.WdRevisionType.wdRevisionDelete Then
                            isDeleted = True
                            Exit For
                        End If
                    Next

                    ' Account for trackchanges being turned on, i.e. the old text remains
                    Dim currentEnd As Integer = doc.Application.Selection.End
                    If Not isDeleted Then
                        currentEnd = currentEnd + Len(newTextWithMarker)
                        selectionEnd = selectionEnd + Len(newTextWithMarker)
                        ' Replace the found text
                        doc.Application.Selection.Text = newTextWithMarker
                    End If

                    ' Check if the collapsed selection has reached the end of the document or the selection
                    If OnlySelection Then
                        If currentEnd >= selectionEnd Then Exit Do
                        doc.Application.Selection.SetRange(currentEnd, selectionEnd)
                    Else
                        If currentEnd >= doc.Content.End Then Exit Do
                        doc.Application.Selection.SetRange(currentEnd, doc.Content.End)
                    End If
                Loop

                If Not found Then
                    CommandsList = $"Note: The search term was not found (Chunk Search)." & Environment.NewLine & CommandsList
                End If

                doc.Application.Selection.SetRange(selectionStart, selectionEnd)
                doc.Application.Selection.Select()

            Else

                If String.IsNullOrEmpty(oldText) Then
                    CommandsList = $"Note: The search term was empty (bad LLM response)." & Environment.NewLine & CommandsList
                Else

                    Dim replacementsMade As Boolean = False
                    Dim initialRangeEnd As Integer = workRange.End
                    Do

                        If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                            Exit Do
                        End If

                        With workRange.Find
                            .ClearFormatting()
                            .Text = oldText
                            .Forward = True
                            .Wrap = Word.WdFindWrap.wdFindStop
                            .MatchWholeWord = False
                            ' Use ReplaceNone to get the match without automatically replacing it
                            If .Execute(Replace:=Word.WdReplace.wdReplaceNone) Then

                                ' Create a duplicate of the found range for the revision check
                                Dim foundRange As Word.Range = workRange.Duplicate

                                Dim isDeleted As Boolean = False
                                For Each rev As Word.Revision In foundRange.Revisions
                                    If rev.Type = Word.WdRevisionType.wdRevisionDelete Then
                                        isDeleted = True
                                        Exit For
                                    End If
                                Next

                                Dim previousStart As Integer = workRange.Start

                                If Not isDeleted Then
                                    foundRange.Text = newTextWithMarker
                                    replacementsMade = True
                                End If

                                ' Adjust the initial end based on the difference in length
                                initialRangeEnd = initialRangeEnd + IIf(isDeleted, 0, Len(newTextWithMarker) - Len(oldText))
                                ' Move the start of workRange to the end of the found match
                                workRange.Start = foundRange.End

                                ' Safeguard: Ensure that the search range advances.
                                If workRange.Start <= previousStart Then
                                    workRange.Start = previousStart + 1
                                End If

                                workRange.End = initialRangeEnd

                            Else
                                Exit Do
                            End If
                        End With

                    Loop

                    If Not replacementsMade Then
                        CommandsList = $"Note: The resarch term was not found." & Environment.NewLine & CommandsList
                    End If

                End If
            End If

        Catch ex As System.Exception
            MsgBox("Error in ExecuteReplaceCommand: " & ex.Message, MsgBoxStyle.Critical)

        Finally
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor
        End Try
    End Sub


    Private Sub ExecuteInsertBeforeAfterCommand(searchText As String, newText As String, Optional OnlySelection As Boolean = False, Optional InsertBefore As Boolean = False)
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' Save the current state of Track Changes and Author
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try

            doc.Application.Activate()
            doc.Activate()

            ' Enable Track Changes and set the author to 
            doc.TrackRevisions = True
            'doc.Application.UserName = AN

            ' Determine the range for the search
            Dim workrange As Word.Range
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    workrange = doc.Content
                Else
                    workrange = doc.Application.Selection.Range
                End If
            Else
                workrange = doc.Content
            End If

            Dim found As Boolean = False

            ' Handle cases where searchText exceeds 255 characters
            If Len(searchText) > 255 Then

                Dim selectionStart As Integer = doc.Application.Selection.Start
                Dim selectionEnd As Integer = doc.Application.Selection.End

                doc.Application.Selection.SetRange(workrange.Start, workrange.End)

                ' Loop through the content to find and replace all instances
                Do While Globals.ThisAddIn.FindLongTextInChunks(searchText, 255, doc.Application.Selection) = True

                    If doc.Application.Selection Is Nothing Then Exit Do

                    found = True

                    ' Account for trackchanges being turned on, i.e. the old text remains
                    Dim currentEnd As Integer = doc.Application.Selection.End + Len(newText)
                    selectionEnd = selectionEnd + Len(newText)

                    ' Insert the found text
                    If InsertBefore Then
                        doc.Application.Selection.InsertBefore(newText)
                    Else
                        doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        doc.Application.Selection.Text = newText & doc.Application.Selection.Text
                    End If

                    ' Check if the collapsed selection has reached the end of the document or the selection
                    If OnlySelection Then
                        If currentEnd >= selectionEnd Then Exit Do
                        doc.Application.Selection.SetRange(currentEnd, selectionEnd)
                    Else
                        If currentEnd >= doc.Content.End Then Exit Do
                        doc.Application.Selection.SetRange(currentEnd, doc.Content.End)
                    End If
                Loop

                If Not found Then
                    CommandsList = $"Note: The search term was not found (Chunk Search)." & Environment.NewLine & CommandsList
                End If

                doc.Application.Selection.SetRange(selectionStart, selectionEnd)
                doc.Application.Selection.Select()

            Else
                ' Use Word's Find functionality for shorter searchText
                With workrange.Find
                    .ClearFormatting()
                    .Text = searchText
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindStop ' Stop after searching once
                    If .Execute() Then
                        found = True
                        ' Insert new text before or after the found text
                        If InsertBefore Then
                            workrange.InsertBefore(newText)
                        Else
                            workrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            workrange.Text = newText & workrange.Text
                        End If

                        ' Select the newly inserted text
                        workrange.Select()
                    End If
                End With
            End If

            If Not found Then
                CommandsList = $"Note: The insertion point was not found." & Environment.NewLine & CommandsList
            End If

        Catch ex As System.Exception
            MsgBox("Error in ExecuteInsertBeforeAfterCommand: " & ex.Message, MsgBoxStyle.Critical)

        Finally
            ' Restore the original state of Track Changes and Author
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor
        End Try
    End Sub


    Private Sub ExecuteInsertCommand(newText As String)

        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' Save the current state of Track Changes and Author
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try

            doc.Application.Activate()
            doc.Activate()

            ' Enable Track Changes and set the author to 
            doc.TrackRevisions = True
            'doc.Application.UserName = AN

            Dim selection As Word.Selection = doc.Application.Selection

            ' Collapse the selection to the insertion point (start of selection)
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart)

            newText = newText.Replace("\r\n", vbCrLf).Replace("\n", vbCrLf).Replace("\r", vbCrLf)

            ' Insert the new text at the current cursor position
            selection.Text = newText

        Catch ex As System.Exception
            MsgBox("Error in ExecuteInsertCommand: " & ex.Message, MsgBoxStyle.Critical)

        Finally
            ' Restore the original state of Track Changes and Author
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor
        End Try
    End Sub


End Class
