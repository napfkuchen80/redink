' Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See License.txt or https://vischer.com/redink for more information.
'
' 9.6.2025
'
' The compiled version of Red Ink also ...
'
' Includes DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Appache-2.0 license (http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex).
' Includes Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json
' Includes HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier,Jeff Klawiter,Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/
' Includes Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/
' Includes PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig
' Includes MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license (https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig
' Includes NAudio in unchanged form; Copyright (c) 2024 Mark Heath; licensed under a proprietary open source license (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio
' Includes Vosk in unchanged form; Copyright (c) 2024 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes Grpc.core in unchanged form; Copyright (c) 2023 The gRPC Authors; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/grpc/grpc
' Includes Google Speech V1 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes Google Protobuf in unchanged form; Copyright (c) 2025 Google Inc.; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/protocolbuffers/protobuf
' Includes MarkdownToRTF in modified form; Copyright (c) 2025 Gustavo Hennig; original licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/GustavoHennig/MarkdownToRtf
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.

Option Explicit On

Imports System.Collections.Concurrent
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.WebSockets
Imports System.Runtime.InteropServices
Imports System.Security.Policy
Imports System.Speech.Synthesis
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports Google.Api.Gax.Grpc
Imports Google.Cloud.Speech.V1
Imports Google.Cloud.Speech.V1.LanguageCodes
Imports Google.Protobuf
Imports Google.Rpc.Context.AttributeContext.Types
Imports Grpc.Core
Imports HtmlAgilityPack
Imports Markdig
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Imports NAudio
Imports NAudio.CoreAudioApi
Imports NAudio.Wave
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Vosk
Imports Whisper.net
Imports Whisper.net.LibraryLoader
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods



Public Class StopForm
    Inherits Form

    Public Property StopRequested As Boolean = False

    Public Sub New()
        Me.Text = "Transkription stoppen"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Width = 200
        Me.Height = 100

        Dim btnStop As New System.Windows.Forms.Button() With {
            .Text = "Stop",
            .Dock = DockStyle.Fill
        }
        AddHandler btnStop.Click, Sub(s, e)
                                      Me.StopRequested = True
                                      Me.Close()
                                  End Sub

        Me.Controls.Add(btnStop)
    End Sub
End Class

Module Module1
    ' Correct attribute declaration for DllImport
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function

End Module

#Region "BridgeSubs"

<ComVisible(True)>
Public Class BridgeSubs
    Public Sub DoInLanguage1()
        Globals.ThisAddIn.InLanguage1()
    End Sub

    Public Sub DoInLanguage2()
        Globals.ThisAddIn.InLanguage2()
    End Sub

    Public Sub DoInOther()
        Globals.ThisAddIn.InOther()
    End Sub

    Public Sub DoCorrect()
        Globals.ThisAddIn.Correct()
    End Sub

    Public Sub DoImprove()
        Globals.ThisAddIn.Improve()
    End Sub

    Public Sub DoNoFillers()
        Globals.ThisAddIn.NoFillers()
    End Sub

    Public Sub DoConvincing()
        Globals.ThisAddIn.Convincing()
    End Sub

    Public Sub DoFriendly()
        Globals.ThisAddIn.Friendly()
    End Sub

    Public Sub DoShorten()
        Globals.ThisAddIn.Shorten()
    End Sub

    Public Sub DoAnonymize()
        Globals.ThisAddIn.Anonymize()
    End Sub

    Public Sub DoSwitchParty()
        Globals.ThisAddIn.SwitchParty()
    End Sub

    Public Sub DoSummarize()
        Globals.ThisAddIn.Summarize()
    End Sub

    Public Sub DoFreestyleNM()
        Globals.ThisAddIn.FreeStyleNM()
    End Sub

    Public Sub DoFreestyleAM()
        Globals.ThisAddIn.FreeStyleAM()
    End Sub

    Public Sub DoContextSearch()
        Globals.ThisAddIn.ContextSearch()
    End Sub

    Public Sub DoCompareSelectionHalves()
        Globals.ThisAddIn.CompareSelectionHalves()
    End Sub

    Public Sub DoAcceptFormatting()
        Globals.ThisAddIn.AcceptFormatting()
    End Sub

    Public Sub DoCalculateUserMarkupTimeSpan()
        Globals.ThisAddIn.CalculateUserMarkupTimeSpan()
    End Sub
    Public Sub DoRegexSearchReplace()
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Public Sub DoImportTextFile()
        Globals.ThisAddIn.ImportTextFile()
    End Sub
    Public Sub DoAddContextMenu()
        Globals.ThisAddIn.AddContextMenu()
    End Sub

End Class

#End Region


Public Class ThisAddIn



    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function SetThreadExecutionState(esFlags As UInteger) As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function FindWindow(
                                ByVal lpClassName As String,
                                ByVal lpWindowName As String
                            ) As IntPtr
    End Function

    Private Function GetWordMainWindowHandle() As IntPtr
        ' Word’s top-level windows all have the class name "OpusApp" (Office 2013+)
        Dim hwnd = FindWindow("OpusApp", Nothing)
        Return hwnd
    End Function

    Private mainThreadControl As New System.Windows.Forms.Control()
    Public StartupInitialized As Boolean = False
    Private WithEvents wordApp As Word.Application

    Private ReadOnly _uiContext As SynchronizationContext = WindowsFormsSynchronizationContext.Current

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ' Necessary for Update Handler to work correctly

        ' 1) Force the creation of the Control's handle on the Office UI thread
        Dim dummy = mainThreadControl.Handle

        ' 2) Give that Control to the UpdateHandler so it can Invoke on it
        UpdateHandler.MainControl = mainThreadControl

        ' 3) Capture the host window’s HWND (Word / Excel / Outlook)
        Dim hwnd As IntPtr
        Dim progId = Me.Application.GetType().Name.ToLowerInvariant()
        If progId.Contains("word") OrElse progId.Contains("excel") Then
            hwnd = New IntPtr(CInt(Me.Application.Hwnd))
        Else
            hwnd = FindWindow("rctrl_renwnd32", Nothing)
        End If
        UpdateHandler.HostHandle = hwnd

        ' Other tasks that need to be done at startup

        SharedMethods.Initialize(Me.CustomTaskPanes)
        wordApp = Application
        Try
            If wordApp IsNot Nothing Then
                AddHandler wordApp.WindowActivate, AddressOf WordApp_WindowActivate
                AddHandler wordApp.DocumentOpen, AddressOf WordApp_DocumentOpen
                AddHandler wordApp.NewDocument, AddressOf WordApp_NewDocument
                AddHandler wordApp.ProtectedViewWindowOpen, AddressOf WordApp_ProtectedViewWindowOpen
                AddHandler wordApp.ProtectedViewWindowBeforeEdit, AddressOf WordApp_ProtectedViewWindowBeforeEdit
                AddHandler wordApp.ProtectedViewWindowActivate, AddressOf WordApp_ProtectedViewWindowActivate
            Else
                mainThreadControl.BeginInvoke(CType(AddressOf DelayedStartupTasks, MethodInvoker))
                StartupInitialized = True
            End If
        Catch ex As System.Exception
            ' Handle exceptions gracefully.
        End Try
    End Sub

    Private Sub RemoveStartupHandlers()
        StartupInitialized = True
        Try
            RemoveHandler wordApp.WindowActivate, AddressOf WordApp_WindowActivate
            RemoveHandler wordApp.DocumentOpen, AddressOf WordApp_DocumentOpen
            RemoveHandler wordApp.NewDocument, AddressOf WordApp_NewDocument
            RemoveHandler wordApp.ProtectedViewWindowOpen, AddressOf WordApp_ProtectedViewWindowOpen
            RemoveHandler wordApp.ProtectedViewWindowBeforeEdit, AddressOf WordApp_ProtectedViewWindowBeforeEdit
            RemoveHandler wordApp.ProtectedViewWindowActivate, AddressOf WordApp_ProtectedViewWindowActivate
        Catch ex As System.Exception
            ' Handle exceptions gracefully.
        End Try
    End Sub

    Private Sub WordApp_WindowActivate(ByVal Doc As Word.Document, ByVal Wn As Word.Window)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    Private Sub WordApp_DocumentOpen(doc As Word.Document)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    Private Sub WordApp_NewDocument(doc As Word.Document)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub


    ' Fires when a file opens in Protected View.
    Private Sub WordApp_ProtectedViewWindowOpen(
            pvWin As ProtectedViewWindow)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    ' Fires just before the user clicks “Edit” in Protected View.
    Private Sub WordApp_ProtectedViewWindowBeforeEdit(
            pvWin As ProtectedViewWindow,
            ByRef Cancel As Boolean)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    ' Fires when the Protected View window is activated.
    Private Sub WordApp_ProtectedViewWindowActivate(
            pvWin As ProtectedViewWindow)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    Private Sub DelayedStartupTasks()
        Try
            InitializeAddInFeatures()
            StartupHttpListener()
        Catch ex As System.Exception
            ' Handle exceptions gracefully.
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        RemoveOldContextMenu()
        ShutdownHttpListener()
    End Sub

    ' Hardcoded config values

    Public Const Version As String = "V.090625 Gen2 Beta Test"

    Public Const AN As String = "Red Ink"
    Public Const AN2 As String = "redink"
    Public Const AN5 As String = "RI" ' for bubble comments 
    Public Const AN6 As String = "Inky" ' for chat

    Private Const ISearch_MinChars = 500         ' minimum characters for a search hit to be relevant
    Private Const ISearch_MaxChars = 4000        ' characters that will be used per search result (rest will be cut off)
    Private Const ISearch_MaxCrawlErrors = 3     ' maximum number of errors before search is aborted
    Private Const ShortenPercent As Integer = 20
    Private Const SummaryPercent As Integer = 20
    Private Const NetTrigger As String = "(net)"
    Private Const LibTrigger As String = "(lib)"
    Private Const AllTrigger As String = "(all)"
    Private Const TPMarkupTrigger As String = "(rev)"
    Private Const TPMarkupTriggerL As String = "(rev:"
    Private Const TPMarkupTriggerR As String = ")"
    Private Const TPMarkupTriggerInstruct As String = "(rev[:user])"
    Private Const ExtTrigger As String = "{doc}"
    Private Const NoFormatTrigger As String = "(noformat)"
    Private Const NoFormatTrigger2 As String = "(nf)"
    Private Const KFTrigger As String = "(keepformat)"
    Private Const KFTrigger2 As String = "(kf)"
    Private Const KPFTrigger As String = "(keepparaformat)"
    Private Const KPFTrigger2 As String = "(kpf)"
    Private Const ObjectTrigger As String = "(file)"
    Private Const ObjectTrigger2 As String = "(clip)"
    Private Const InPlacePrefix As String = "Replace:"
    Private Const MarkupPrefix As String = "Markup:"
    Private Const MarkupPrefixDiff As String = "MarkupDiff:"
    Private Const MarkupPrefixDiffW As String = "MarkupDiffW:"
    Private Const MarkupPrefixWord As String = "MarkupWord:"
    Private Const MarkupPrefixRegex As String = "MarkupRegex:"
    Private Const MarkupPrefixAll As String = "Markup[Diff|DiffW|Word|Regex]:"
    Private Const PurePrefix As String = "Pure:"
    Private Const ClipboardPrefix As String = "Clipboard:"
    Private Const ClipboardPrefix2 As String = "Clip:"
    Private Const PanePrefix As String = "Pane:"
    Private Const BubblesPrefix As String = "Bubbles:"
    Private Const BubbleCutText As String = " (" & ChrW(&H2702) & ")"
    Private Const SearchNextTrigger As String = "Next:"
    Private Const BoWTrigger As String = "(bow)"
    Private Const EmbedTrigger As String = "(embed)"
    Private Const RefreshTrigger As String = "(refresh)"

    Private Const RegexSeparator1 As String = "|||"  ' Set also in SharedLibrary
    Private Const RegexSeparator2 As String = "§§§"  ' Set also in SharedLibrary 
    Private Const RIMenu = AN
    Private Const OldRIMenu = AN & " " & ChrW(&HD83D) & ChrW(&HDC09)
    Private Const MinHelperVersion = 1 ' Minimum version of the helper file that is required
    Private Const VoskSource = "https://alphacephei.com/vosk/models"
    Private Const WhisperSource = "https://huggingface.co/ggerganov/whisper.cpp/tree/main"

    Public Shared WhisperSupportedLanguages As New HashSet(Of String) From {
                            "af", "sq", "am", "ar", "hy", "as", "az", "ba", "eu", "be", "bn", "bs", "br", "bg",
                            "ca", "zh", "hr", "cs", "da", "nl", "en", "et", "fo", "fi", "fr", "gl", "ka", "de",
                            "el", "gu", "ht", "ha", "he", "hi", "hu", "is", "id", "it", "ja", "jv", "kn", "kk",
                            "km", "rw", "ky", "ko", "lv", "lt", "lb", "mk", "mg", "ms", "ml", "mt", "mi", "mr",
                            "mn", "my", "ne", "no", "oc", "ps", "fa", "pl", "pt", "pa", "ro", "ru", "sa", "sr",
                            "sd", "si", "sk", "sl", "so", "es", "su", "sw", "sv", "tl", "tg", "ta", "tt", "te",
                            "th", "tr", "uk", "ur", "uz", "vi", "cy", "yi", "yo", "zu", "auto"
                        }

    Public Shared GoogleTTSsupportedLanguages As String() = {
            "en-US", "en-GB", "de-DE", "fr-FR", "es-ES", "it-IT",
            "af-ZA", "sq-AL", "am-ET", "ar-SA", "eu-ES", "bn-BD",
            "bs-BA", "bg-BG", "yue-HK", "ca-ES", "zh-CN", "zh-TW",
            "hr-HR", "cs-CZ", "da-DK", "nl-NL", "en-AU", "en-IN",
            "en-NG", "et-EE", "fil-PH", "fi-FI", "fr-CA", "gl-ES",
            "el-GR", "gu-IN", "ha-NG", "he-IL", "hi-IN", "hu-HU",
            "is-IS", "id-ID", "ja-JP", "jv-ID", "kn-IN", "km-KH",
            "ko-KR", "la-LA", "lv-LV", "lt-LT", "ms-MY", "ml-IN",
            "mr-IN", "my-MM", "ne-NP", "nb-NO", "pl-PL", "pt-BR",
            "pt-PT", "pa-IN", "ro-RO", "ru-RU", "sr-RS", "si-LK",
            "sk-SK", "es-US", "su-ID", "sw-KE", "sv-SE", "ta-IN",
            "te-IN", "th-TH", "tr-TR", "uk-UA", "ur-PK", "vi-VN", "cy-GB"
        }

    ' Human-readable descriptions for each OpenAI voice.
    Private Shared ReadOnly OpenAIDescriptions As New Dictionary(Of String, String) From {
    {"alloy", "Female: Versatile and balanced"},
    {"ash", "Male: Clear and precise"},
    {"ballad", "Male: Melodic and smooth"},
    {"coral", "Female: Warm and friendly"},
    {"echo", "Male: Warm and natural"},
    {"fable", "Male: Engaging storyteller"},
    {"nova", "Female: Bright and energetic"},
    {"onyx", "Male: Deep and authoritative"},
    {"sage", "Male: Calm and thoughtful"},
    {"shimmer", "Female: Clear and expressive"},
    {"verse", "Male: Versatile and expressive"}
}

    Private Const TTS_OpenAI_Model = "gpt-4o-mini-tts"

    Private Shared ReadOnly OpenAIVoices As String() = OpenAIDescriptions.Keys.ToArray()
    Private Shared ReadOnly OpenAILanguages As String() = {
    "de", "en", "es", "fr", "it", "ja", "ko", "pt", "ru", "zh",
    "ar", "bg", "ca", "cs", "da", "el", "et", "fi", "hi", "hu",
    "id", "nl", "no", "pl", "ro", "sv", "th", "tr", "uk", "vi"
}

    Private Const ES_CONTINUOUS As UInteger = &H80000000UI
    Private Const ES_SYSTEM_REQUIRED As UInteger = &H1UI
    Private Const ES_DISPLAY_REQUIRED As UInteger = &H2UI
    Private Const ES_AWAYMODE_REQUIRED As UInteger = &H40UI

    Private Const ES_KEEP_SYSTEM_ONLY As UInteger = ES_CONTINUOUS Or ES_SYSTEM_REQUIRED
    Private Const ES_KEEP_SYSTEM_AND_DISPLAY As UInteger = ES_CONTINUOUS Or ES_SYSTEM_REQUIRED Or ES_DISPLAY_REQUIRED
    Private Const ES_KEEP_CURRENT_SETTING As UInteger = ES_KEEP_SYSTEM_ONLY

    Private Shared prevExecState As UInteger = Nothing

    Private Const Code_JsonTemplateFormatter As String = "Public Module JsonTemplateFormatter" & vbCrLf & "''' <summary>''' Hauptfunktion für JSON-String + Template''' </summary>" & vbCrLf & "Public Function FormatJsonWithTemplate(json As String, ByVal template As String) As String" & vbCrLf & "Dim jObj As JObject" & vbCrLf & "Try" & vbCrLf & "jObj = JObject.Parse(json)" & vbCrLf & "Catch ex As Newtonsoft.Json.JsonReaderException" & vbCrLf & "Return $""[Fehler beim Parsen des JSON: {ex.Message}]""" & vbCrLf & "End Try" & vbCrLf & "Return FormatJsonWithTemplate(jObj, template)" & vbCrLf & "End Function" & vbCrLf & "''' <summary>''' Hauptfunktion für direkten JObject + Template''' </summary>" & vbCrLf & "Public Function FormatJsonWithTemplate(jObj As JObject, ByVal template As String) As String" & vbCrLf & "If String.IsNullOrWhiteSpace(template) Then Return """"" & vbCrLf & "template = template.Replace(""\\N"", vbCrLf).Replace(""\\n"", vbCrLf).Replace(""\\R"", vbCrLf).Replace(""\\r"", vbCrLf)" & vbCrLf & "template = Regex.Replace(template, ""<cr>"", vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "Dim hasLoop = Regex.IsMatch(template, ""\\{\\%\\s*for\\s+([^\\s\\%]+)\\s*\\%\\}"", RegexOptions.Singleline)" & vbCrLf & "Dim hasPh = Regex.IsMatch(template, ""\\{([^}]+)\\}"")" & vbCrLf & "If Not hasLoop AndAlso Not hasPh Then Return FindJsonProperty(jObj, template)" & vbCrLf & "Dim loopRegex = New Regex(""\\{\\%\\s*for\\s+([^%\\s]+)\\s*\\%\\}(.*?)\\{\\%\\s*endfor\\s*\\%\\}"", RegexOptions.Singleline Or RegexOptions.IgnoreCase)" & vbCrLf & "Dim mLoop = loopRegex.Match(template)" & vbCrLf & "While mLoop.Success" & vbCrLf & "Dim fullBlock = mLoop.Value" & vbCrLf & "Dim rawPath = mLoop.Groups(1).Value.Trim()" & vbCrLf & "Dim innerTpl = mLoop.Groups(2).Value" & vbCrLf & "Dim path = If(rawPath.StartsWith(""$""), rawPath, ""$."" & rawPath)" & vbCrLf & "Dim tokens = jObj.SelectTokens(path)" & vbCrLf & "Dim items = tokens.SelectMany(Function(t) If t.Type = JTokenType.Array Then Return CType(t, JArray).OfType(Of JObject)() ElseIf t.Type = JTokenType.Object Then Return {CType(t, JObject)} Else Return Enumerable.Empty(Of JObject)() End If)" & vbCrLf & "Dim rendered = items.Select(Function(o) FormatJsonWithTemplate(o, innerTpl)).ToArray()" & vbCrLf & "template = template.Replace(fullBlock, If(rendered.Any, String.Join(vbCrLf & vbCrLf, rendered), """"""))" & vbCrLf & "mLoop = loopRegex.Match(template)" & vbCrLf & "End While" & vbCrLf & "Dim phRegex = New Regex(""\\{(.+?)\\}"", RegexOptions.Singleline)" & vbCrLf & "Dim result = template" & vbCrLf & "For Each mPh As Match In phRegex.Matches(template)" & vbCrLf & "Dim fullPh = mPh.Value" & vbCrLf & "Dim content = mPh.Groups(1).Value" & vbCrLf & "Dim isHtml As Boolean = False" & vbCrLf & "Dim isNoCr As Boolean = False" & vbCrLf & "If content.StartsWith(""htmlnocr:"", StringComparison.OrdinalIgnoreCase) Then isHtml = True : isNoCr = True : content = content.Substring(""htmlnocr:"".Length) ElseIf content.StartsWith(""html:"", StringComparison.OrdinalIgnoreCase) Then isHtml = True : content = content.Substring(""html:"".Length) ElseIf content.StartsWith(""nocr:"", StringComparison.OrdinalIgnoreCase) Then isNoCr = True : content = content.Substring(""nocr:"".Length)" & vbCrLf & "Dim parts = content.Split(New Char() {""|""c}, 2)" & vbCrLf & "Dim pathPh = parts(0).Trim()" & vbCrLf & "Dim remainder = If(parts.Length > 1, parts(1), String.Empty)" & vbCrLf & "Dim sep As String = vbCrLf" & vbCrLf & "Dim mappings As Dictionary(Of String, String) = Nothing" & vbCrLf & "If Not String.IsNullOrEmpty(remainder) Then If remainder.Contains(""=""c) Then mappings = ParseMappings(remainder) Else sep = remainder.Replace(""\\n"", vbCrLf)" & vbCrLf & "Dim replacement = RenderTokens(jObj, pathPh, sep, isHtml, isNoCr, mappings)" & vbCrLf & "result = result.Replace(fullPh, replacement)" & vbCrLf & "Next" & vbCrLf & "Return result" & vbCrLf & "End Function" & vbCrLf & "Private Function RenderTokens(jObj As JObject, path As String, sep As String, isHtml As Boolean, isNoCr As Boolean, mappings As Dictionary(Of String, String)) As String" & vbCrLf & "Try" & vbCrLf & "If Not path.StartsWith(""$"") AndAlso Not path.StartsWith(""@"") Then path = ""$."" & path" & vbCrLf & "Dim tokens = jObj.SelectTokens(path)" & vbCrLf & "Dim list As New List(Of String)" & vbCrLf & "For Each t In tokens" & vbCrLf & "Dim raw = t.ToString()" & vbCrLf & "If mappings IsNot Nothing AndAlso mappings.ContainsKey(raw) Then raw = mappings(raw)" & vbCrLf & "If isHtml Then raw = HtmlToMarkdownSimple(raw)" & vbCrLf & "If isNoCr Then raw = Regex.Replace(raw, ""[\r\n]+"", "" "") : raw = Regex.Replace(raw, ""\s{2,}"", "" "") : raw = Regex.Replace(raw, ""[\u2022\u2023\u25E6]"", String.Empty) : raw = raw.Trim()" & vbCrLf & "list.Add(raw)" & vbCrLf & "Next" & vbCrLf & "Return If(list.Count = 0, """", String.Join(sep, list))" & vbCrLf & "Catch ex As System.Exception" & vbCrLf & "Return """"" & vbCrLf & "End Try" & vbCrLf & "End Function" & vbCrLf & "Private Function ParseMappings(defs As String) As Dictionary(Of String, String)" & vbCrLf & "Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)" & vbCrLf & "For Each pair In defs.Split("";""c)" & vbCrLf & "Dim kv = pair.Split(New Char() {""=""c}, 2)" & vbCrLf & "If kv.Length = 2 Then dict(kv(0).Trim()) = kv(1).Trim()" & vbCrLf & "Next" & vbCrLf & "Return dict" & vbCrLf & "End Function" & vbCrLf & "Public Function HtmlToMarkdownSimple(html As String) As String" & vbCrLf & "Dim s = WebUtility.HtmlDecode(html)" & vbCrLf & "s = Regex.Replace(s, ""</?p\s*/?>"", vbCrLf & vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "s = Regex.Replace(s, ""<br\s*/?>"", vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "s = Regex.Replace(s, ""<strong>(.*?)</strong>"", ""**$1**"", RegexOptions.IgnoreCase)" & vbCrLf & "s = Regex.Replace(s, ""<em>(.*?)</em>"", ""*$1*"", RegexOptions.IgnoreCase)" & vbCrLf & "s = Regex.Replace(s, ""<span\b[^>]*>(.*?)</span>"", ""*$1*"", RegexOptions.IgnoreCase)" & vbCrLf & "s = Regex.Replace(s, ""<li>(.*?)</li>"", ""- $1"" & vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "s = Regex.Replace(s, ""<[^>]+>"", String.Empty)" & vbCrLf & "s = Regex.Replace(s, ""("" & vbCrLf & ""){3,}"", vbCrLf & vbCrLf)" & vbCrLf & "Return s.Trim()" & vbCrLf & "End Function" & vbCrLf & "End Module"
    Private Const SP_GenerateResponseKey As String = "I have code that will generate from a JSON string an Markdown output using a Template, which the code will parse together with the JSON file. I want you to create me a working template taking into account (i) the code, (ii) the structure of the JSON file and (iii) my instructions. If the JSON has arrays, make sure you correctly handle them. To produce your output, first provide the barebones template one one single line (do not use placeholders, provide the text how template should look like; for linebreaks, use only <cr>), then provide a brief explanation without any formatting. I will provide you in the following first the code, and then you will get the (sample) JSON file and my instructions. Follow them carefully."

    Private Const NER_Model = "anon\model.onnx"
    Private Const NER_Token = "anon\bpe.model"
    Private Const NER_Label = "anon\label_map.txt"
    Private Const Embed_Model = "embed\model.onnx"
    Private Const Embed_Vocab = "embed\vocab.txt"
    Private Default_Embed_Min_Score = 0.2
    Private Default_Embed_Top_K = 5
    Private Default_Embed_Chunks = 2
    Private Default_Embed_Overlap = 1
    Private Default_Embed_Chunks_bow = 1
    Private Default_Embed_Overlap_bow = 0

    Public Shared DragDropFormLabel As String = ""
    Public Shared DragDropFormFilter As String = ""

    Public Shared TTSDefaultFile As String = $"{AN2}-output.mp3"
    Public Const TTSLargeText As Integer = 2500
    Public Shared hostTags As String() = {"H:", "Host:", "A:", "1:"}
    Public Shared guestTags As String() = {"G:", "Guest:", "Gast:", "B:", "2:"}
    Public Shared GoogleIdentifier As String = "googleapis.com"
    Public Shared OpenAIIdentifier As String = "openai.com"

    Public Shared TTS_googleAvailable As Boolean = False
    Private Shared TTS_googleSecondary As Boolean = False
    Public Shared TTS_openAIAvailable As Boolean = False
    Private Shared TTS_openAISecondary As Boolean = False
    Private Shared TTS_GoogleEndpoint As String = ""
    Private Shared TTS_OpenAIEndpoint As String = ""

    Public Shared GoogleSTT_Desc As String = "Google STT V1 (run in EU)"
    Public Shared STTEndpoint As String = "eu-speech.googleapis.com"


    Public Shared OpenAISTTModel As String = "gpt-4o-realtime-preview"
    Public Shared OpenAISTT_Desc As String = $"OpenAI Streaming"
    Public Shared STTEndpoint_OpenAI As String = $"wss://api.openai.com/v1/realtime?model=gpt-4o-realtime-preview"

    Public Shared GoogleSTTsupportedLanguages As String() = {
    "en-US", "de-DE",
    "de-AT", "de-CH", "es-AR", "es-BO", "es-CL", "es-CO", "es-CR", "es-DO", "es-EC", "es-ES", "es-GT",
    "es-HN", "es-MX", "es-NI", "es-PA", "es-PE", "es-PR", "es-PY", "es-SV", "es-UY", "es-VE",
    "fr-BE", "fr-CA", "fr-CH", "fr-FR", "it-CH", "it-IT", "nl-BE", "nl-NL",
    "af-ZA", "am-ET", "ar-BH", "ar-DZ", "ar-EG", "ar-IQ", "ar-IL", "ar-JO", "ar-KW", "ar-LB", "ar-MA",
    "ar-MR", "ar-OM", "ar-PS", "ar-QA", "ar-SA", "ar-SY", "ar-TN", "ar-AE", "ar-YE", "az-AZ", "bg-BG",
    "bn-BD", "bn-IN", "bs-BA", "ca-ES", "cmn-Hans-CN", "cmn-Hans-HK", "cmn-Hant-TW", "cs-CZ",
    "da-DK", "el-GR", "en-AU", "en-CA", "en-GH", "en-HK", "en-IE", "en-IN", "en-KE", "en-NG",
    "en-NZ", "en-PH", "en-PK", "en-SG", "en-TZ", "en-ZA", "et-EE", "eu-ES", "fa-IR", "fi-FI",
    "fil-PH", "gl-ES", "gu-IN", "hi-IN", "hr-HR", "hu-HU", "hy-AM", "id-ID", "is-IS", "iw-IL",
    "ja-JP", "jv-ID", "ka-GE", "kk-KZ", "km-KH", "kn-IN", "ko-KR", "lo-LA", "lt-LT", "lv-LV",
    "ml-IN", "mn-MN", "mr-IN", "ms-MY", "my-MM", "ne-NP", "no-NO", "pa-Guru-IN", "pl-PL",
    "pt-BR", "pt-PT", "ro-RO", "rw-RW", "si-LK", "sk-SK", "sl-SI", "sr-RS", "ss-Latn-ZA", "st-ZA",
    "su-ID", "sv-SE", "sw-KE", "sw-TZ", "ta-IN", "ta-LK", "ta-MY", "ta-SG", "te-IN", "th-TH",
    "tn-Latn-ZA", "tr-TR", "uk-UA", "ur-IN", "ur-PK", "uz-UZ", "ve-ZA", "vi-VN", "xh-ZA",
    "yue-Hant-HK", "zu-ZA"
        }




    ' Definition of the SharedProperties for context for exchanging values with the SharedLibrary

#Region "SharedProperties"

    Private Shared _context As ISharedContext = New SharedContext()

    Public Shared Property INI_APIKey As String
        Get
            Return _context.INI_APIKey
        End Get
        Set(value As String)
            _context.INI_APIKey = value
        End Set
    End Property

    Public Shared Property INI_APIKeyBack As String
        Get
            Return _context.INI_APIKeyBack
        End Get
        Set(value As String)
            _context.INI_APIKeyBack = value
        End Set
    End Property

    Public Shared Property INI_Temperature As String
        Get
            Return _context.INI_Temperature
        End Get
        Set(value As String)
            _context.INI_Temperature = value
        End Set
    End Property

    Public Shared Property INI_Timeout As Long
        Get
            Return _context.INI_Timeout
        End Get
        Set(value As Long)
            _context.INI_Timeout = value
        End Set
    End Property

    Public Shared Property INI_MaxOutputToken As Integer
        Get
            Return _context.INI_MaxOutputToken
        End Get
        Set(value As Integer)
            _context.INI_MaxOutputToken = value
        End Set
    End Property

    Public Shared Property INI_Model As String
        Get
            Return _context.INI_Model
        End Get
        Set(value As String)
            _context.INI_Model = value
        End Set
    End Property

    Public Shared Property INI_Endpoint As String
        Get
            Return _context.INI_Endpoint
        End Get
        Set(value As String)
            _context.INI_Endpoint = value
        End Set
    End Property

    Public Shared Property INI_HeaderA As String
        Get
            Return _context.INI_HeaderA
        End Get
        Set(value As String)
            _context.INI_HeaderA = value
        End Set
    End Property

    Public Shared Property INI_HeaderB As String
        Get
            Return _context.INI_HeaderB
        End Get
        Set(value As String)
            _context.INI_HeaderB = value
        End Set
    End Property

    Public Shared Property INI_APICall As String
        Get
            Return _context.INI_APICall
        End Get
        Set(value As String)
            _context.INI_APICall = value
        End Set
    End Property

    Public Shared Property INI_APICall_Object As String
        Get
            Return _context.INI_APICall_Object
        End Get
        Set(value As String)
            _context.INI_APICall_Object = value
        End Set
    End Property


    Public Shared Property INI_Response As String
        Get
            Return _context.INI_Response
        End Get
        Set(value As String)
            _context.INI_Response = value
        End Set
    End Property

    Public Shared Property INI_Anon As String
        Get
            Return _context.INI_Anon
        End Get
        Set(value As String)
            _context.INI_Anon = value
        End Set
    End Property

    Public Shared Property INI_DoubleS As Boolean
        Get
            Return _context.INI_DoubleS
        End Get
        Set(value As Boolean)
            _context.INI_DoubleS = value
        End Set
    End Property

    Public Shared Property INI_Clean As Boolean
        Get
            Return _context.INI_Clean
        End Get
        Set(value As Boolean)
            _context.INI_Clean = value
        End Set
    End Property


    Public Shared Property INI_PreCorrection As String
        Get
            Return _context.INI_PreCorrection
        End Get
        Set(value As String)
            _context.INI_PreCorrection = value
        End Set
    End Property

    Public Shared Property INI_PostCorrection As String
        Get
            Return _context.INI_PostCorrection
        End Get
        Set(value As String)
            _context.INI_PostCorrection = value
        End Set
    End Property

    Public Shared Property INI_APIEncrypted As Boolean
        Get
            Return _context.INI_APIEncrypted
        End Get
        Set(value As Boolean)
            _context.INI_APIEncrypted = value
        End Set
    End Property

    Public Shared Property INI_APIKeyPrefix As String
        Get
            Return _context.INI_APIKeyPrefix
        End Get
        Set(value As String)
            _context.INI_APIKeyPrefix = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodOutlook As Integer
        Get
            Return _context.INI_MarkupMethodOutlook
        End Get
        Set(value As Integer)
            _context.INI_MarkupMethodOutlook = value
        End Set
    End Property

    Public Shared Property INI_MarkupDiffCap As Integer
        Get
            Return _context.INI_MarkupDiffCap
        End Get
        Set(value As Integer)
            _context.INI_MarkupDiffCap = value
        End Set
    End Property

    Public Shared Property INI_MarkupRegexCap As Integer
        Get
            Return _context.INI_MarkupRegexCap
        End Get
        Set(value As Integer)
            _context.INI_MarkupRegexCap = value
        End Set
    End Property

    Public Shared Property INI_OpenSSLPath As String
        Get
            Return _context.INI_OpenSSLPath
        End Get
        Set(value As String)
            _context.INI_OpenSSLPath = value
        End Set
    End Property


    Public Shared Property INI_OAuth2 As Boolean
        Get
            Return _context.INI_OAuth2
        End Get
        Set(value As Boolean)
            _context.INI_OAuth2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ClientMail As String
        Get
            Return _context.INI_OAuth2ClientMail
        End Get
        Set(value As String)
            _context.INI_OAuth2ClientMail = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Scopes As String
        Get
            Return _context.INI_OAuth2Scopes
        End Get
        Set(value As String)
            _context.INI_OAuth2Scopes = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Endpoint As String
        Get
            Return _context.INI_OAuth2Endpoint
        End Get
        Set(value As String)
            _context.INI_OAuth2Endpoint = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ATExpiry As Long
        Get
            Return _context.INI_OAuth2ATExpiry
        End Get
        Set(value As Long)
            _context.INI_OAuth2ATExpiry = value
        End Set
    End Property

    Public Shared Property INI_SecondAPI As Boolean
        Get
            Return _context.INI_SecondAPI
        End Get
        Set(value As Boolean)
            _context.INI_SecondAPI = value
        End Set
    End Property

    Public Shared Property INI_APIKey_2 As String
        Get
            Return _context.INI_APIKey_2
        End Get
        Set(value As String)
            _context.INI_APIKey_2 = value
        End Set
    End Property

    Public Shared Property INI_APIKeyBack_2 As String
        Get
            Return _context.INI_APIKeyBack_2
        End Get
        Set(value As String)
            _context.INI_APIKeyBack_2 = value
        End Set
    End Property

    Public Shared Property INI_Temperature_2 As String
        Get
            Return _context.INI_Temperature_2
        End Get
        Set(value As String)
            _context.INI_Temperature_2 = value
        End Set
    End Property

    Public Shared Property INI_Timeout_2 As Long
        Get
            Return _context.INI_Timeout_2
        End Get
        Set(value As Long)
            _context.INI_Timeout_2 = value
        End Set
    End Property
    Public Shared Property INI_MaxOutputToken_2 As Integer
        Get
            Return _context.INI_MaxOutputToken_2
        End Get
        Set(value As Integer)
            _context.INI_MaxOutputToken_2 = value
        End Set
    End Property

    Public Shared Property INI_Model_2 As String
        Get
            Return _context.INI_Model_2
        End Get
        Set(value As String)
            _context.INI_Model_2 = value
        End Set
    End Property

    Public Shared Property INI_Endpoint_2 As String
        Get
            Return _context.INI_Endpoint_2
        End Get
        Set(value As String)
            _context.INI_Endpoint_2 = value
        End Set
    End Property

    Public Shared Property INI_HeaderA_2 As String
        Get
            Return _context.INI_HeaderA_2
        End Get
        Set(value As String)
            _context.INI_HeaderA_2 = value
        End Set
    End Property

    Public Shared Property INI_HeaderB_2 As String
        Get
            Return _context.INI_HeaderB_2
        End Get
        Set(value As String)
            _context.INI_HeaderB_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_2 As String
        Get
            Return _context.INI_APICall_2
        End Get
        Set(value As String)
            _context.INI_APICall_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_Object_2 As String
        Get
            Return _context.INI_APICall_Object_2
        End Get
        Set(value As String)
            _context.INI_APICall_Object_2 = value
        End Set
    End Property


    Public Shared Property INI_Response_2 As String
        Get
            Return _context.INI_Response_2
        End Get
        Set(value As String)
            _context.INI_Response_2 = value
        End Set
    End Property

    Public Shared Property INI_Anon_2 As String
        Get
            Return _context.INI_Anon_2
        End Get
        Set(value As String)
            _context.INI_Anon_2 = value
        End Set
    End Property

    Public Shared Property INI_APIEncrypted_2 As Boolean
        Get
            Return _context.INI_APIEncrypted_2
        End Get
        Set(value As Boolean)
            _context.INI_APIEncrypted_2 = value
        End Set
    End Property

    Public Shared Property INI_APIKeyPrefix_2 As String
        Get
            Return _context.INI_APIKeyPrefix_2
        End Get
        Set(value As String)
            _context.INI_APIKeyPrefix_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2_2 As Boolean
        Get
            Return _context.INI_OAuth2_2
        End Get
        Set(value As Boolean)
            _context.INI_OAuth2_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ClientMail_2 As String
        Get
            Return _context.INI_OAuth2ClientMail_2
        End Get
        Set(value As String)
            _context.INI_OAuth2ClientMail_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Scopes_2 As String
        Get
            Return _context.INI_OAuth2Scopes_2
        End Get
        Set(value As String)
            _context.INI_OAuth2Scopes_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Endpoint_2 As String
        Get
            Return _context.INI_OAuth2Endpoint_2
        End Get
        Set(value As String)
            _context.INI_OAuth2Endpoint_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ATExpiry_2 As Long
        Get
            Return _context.INI_OAuth2ATExpiry_2
        End Get
        Set(value As Long)
            _context.INI_OAuth2ATExpiry_2 = value
        End Set
    End Property

    Public Shared Property INI_APIDebug As Boolean
        Get
            Return _context.INI_APIDebug
        End Get
        Set(value As Boolean)
            _context.INI_APIDebug = value
        End Set
    End Property

    Public Shared Property INI_UsageRestrictions As String
        Get
            Return _context.INI_UsageRestrictions
        End Get
        Set(value As String)
            _context.INI_UsageRestrictions = value
        End Set
    End Property

    Public Shared Property INI_Language1 As String
        Get
            Return _context.INI_Language1
        End Get
        Set(value As String)
            _context.INI_Language1 = value
        End Set
    End Property

    Public Shared Property INI_Language2 As String
        Get
            Return _context.INI_Language2
        End Get
        Set(value As String)
            _context.INI_Language2 = value
        End Set
    End Property

    Public Shared Property INI_KeepFormat1 As Boolean
        Get
            Return _context.INI_KeepFormat1
        End Get
        Set(value As Boolean)
            _context.INI_KeepFormat1 = value
        End Set
    End Property

    Public Shared Property INI_MarkdownConvert As Boolean
        Get
            Return _context.INI_MarkdownConvert
        End Get
        Set(value As Boolean)
            _context.INI_MarkdownConvert = value
        End Set
    End Property


    Public Shared Property INI_KeepFormat2 As Boolean
        Get
            Return _context.INI_KeepFormat2
        End Get
        Set(value As Boolean)
            _context.INI_KeepFormat2 = value
        End Set
    End Property

    Public Shared Property INI_KeepParaFormatInline As Boolean
        Get
            Return _context.INI_KeepParaFormatInline
        End Get
        Set(value As Boolean)
            _context.INI_KeepParaFormatInline = value
        End Set
    End Property

    Public Shared Property INI_KeepFormatCap As Integer
        Get
            Return _context.INI_KeepFormatCap
        End Get
        Set(value As Integer)
            _context.INI_KeepFormatCap = value
        End Set
    End Property


    Public Shared Property INI_ReplaceText1 As Boolean
        Get
            Return _context.INI_ReplaceText1
        End Get
        Set(value As Boolean)
            _context.INI_ReplaceText1 = value
        End Set
    End Property

    Public Shared Property INI_ReplaceText2 As Boolean
        Get
            Return _context.INI_ReplaceText2
        End Get
        Set(value As Boolean)
            _context.INI_ReplaceText2 = value
        End Set
    End Property

    Public Shared Property INI_DoMarkupOutlook As Boolean
        Get
            Return _context.INI_DoMarkupOutlook
        End Get
        Set(value As Boolean)
            _context.INI_DoMarkupOutlook = value
        End Set
    End Property

    Public Shared Property INI_DoMarkupWord As Boolean
        Get
            Return _context.INI_DoMarkupWord
        End Get
        Set(value As Boolean)
            _context.INI_DoMarkupWord = value
        End Set
    End Property

    Public Shared Property INI_RoastMe As Boolean
        Get
            Return _context.INI_RoastMe
        End Get
        Set(value As Boolean)
            _context.INI_RoastMe = value
        End Set
    End Property


    Public Shared Property SP_Translate As String
        Get
            Return _context.SP_Translate
        End Get
        Set(value As String)
            _context.SP_Translate = value
        End Set
    End Property

    Public Shared Property SP_Correct As String
        Get
            Return _context.SP_Correct
        End Get
        Set(value As String)
            _context.SP_Correct = value
        End Set
    End Property

    Public Shared Property SP_Improve As String
        Get
            Return _context.SP_Improve
        End Get
        Set(value As String)
            _context.SP_Improve = value
        End Set
    End Property

    Public Shared Property SP_Explain As String
        Get
            Return _context.SP_Explain
        End Get
        Set(value As String)
            _context.SP_Explain = value
        End Set
    End Property

    Public Shared Property SP_SuggestTitles As String
        Get
            Return _context.SP_SuggestTitles
        End Get
        Set(value As String)
            _context.SP_SuggestTitles = value
        End Set
    End Property

    Public Shared Property SP_Friendly As String
        Get
            Return _context.SP_Friendly
        End Get
        Set(value As String)
            _context.SP_Friendly = value
        End Set
    End Property

    Public Shared Property SP_Convincing As String
        Get
            Return _context.SP_Convincing
        End Get
        Set(value As String)
            _context.SP_Convincing = value
        End Set
    End Property

    Public Shared Property SP_NoFillers As String
        Get
            Return _context.SP_NoFillers
        End Get
        Set(value As String)
            _context.SP_NoFillers = value
        End Set
    End Property

    Public Shared Property SP_Podcast As String
        Get
            Return _context.SP_Podcast
        End Get
        Set(value As String)
            _context.SP_Podcast = value
        End Set
    End Property

    Public Shared Property SP_Shorten As String
        Get
            Return _context.SP_Shorten
        End Get
        Set(value As String)
            _context.SP_Shorten = value
        End Set
    End Property

    Public Shared Property SP_InsertClipboard As String
        Get
            Return _context.SP_InsertClipboard
        End Get
        Set(value As String)
            _context.SP_InsertClipboard = value
        End Set
    End Property

    Public Shared Property SP_Summarize As String
        Get
            Return _context.SP_Summarize
        End Get
        Set(value As String)
            _context.SP_Summarize = value
        End Set
    End Property

    Public Shared Property SP_MailReply As String
        Get
            Return _context.SP_MailReply
        End Get
        Set(value As String)
            _context.SP_MailReply = value
        End Set
    End Property

    Public Shared Property SP_MailSumup As String
        Get
            Return _context.SP_MailSumup
        End Get
        Set(value As String)
            _context.SP_MailSumup = value
        End Set
    End Property

    Public Shared Property SP_MailSumup2 As String
        Get
            Return _context.SP_MailSumup2
        End Get
        Set(value As String)
            _context.SP_MailSumup2 = value
        End Set
    End Property

    Public Shared Property SP_FreestyleText As String
        Get
            Return _context.SP_FreestyleText
        End Get
        Set(value As String)
            _context.SP_FreestyleText = value
        End Set
    End Property

    Public Shared Property SP_FreestyleNoText As String
        Get
            Return _context.SP_FreestyleNoText
        End Get
        Set(value As String)
            _context.SP_FreestyleNoText = value
        End Set
    End Property

    Public Shared Property SP_SwitchParty As String
        Get
            Return _context.SP_SwitchParty
        End Get
        Set(value As String)
            _context.SP_SwitchParty = value
        End Set
    End Property

    Public Shared Property SP_Anonymize As String
        Get
            Return _context.SP_Anonymize
        End Get
        Set(value As String)
            _context.SP_Anonymize = value
        End Set
    End Property

    Public Shared Property SP_ContextSearch As String
        Get
            Return _context.SP_ContextSearch
        End Get
        Set(value As String)
            _context.SP_ContextSearch = value
        End Set
    End Property

    Public Shared Property SP_ContextSearchMulti As String
        Get
            Return _context.SP_ContextSearchMulti
        End Get
        Set(value As String)
            _context.SP_ContextSearchMulti = value
        End Set
    End Property


    Public Shared Property SP_RangeOfCells As String
        Get
            Return _context.SP_RangeOfCells
        End Get
        Set(value As String)
            _context.SP_RangeOfCells = value
        End Set
    End Property

    Public Shared Property SP_WriteNeatly As String
        Get
            Return _context.SP_WriteNeatly
        End Get
        Set(value As String)
            _context.SP_WriteNeatly = value
        End Set
    End Property

    Public Shared Property SP_Add_KeepFormulasIntact As String
        Get
            Return _context.SP_Add_KeepFormulasIntact
        End Get
        Set(value As String)
            _context.SP_Add_KeepFormulasIntact = value
        End Set
    End Property

    Public Shared Property SP_Add_KeepHTMLIntact As String
        Get
            Return _context.SP_Add_KeepHTMLIntact
        End Get
        Set(value As String)
            _context.SP_Add_KeepHTMLIntact = value
        End Set
    End Property

    Public Shared Property SP_Add_KeepInlineIntact As String
        Get
            Return _context.SP_Add_KeepInlineIntact
        End Get
        Set(value As String)
            _context.SP_Add_KeepInlineIntact = value
        End Set
    End Property

    Public Shared Property SP_Add_Bubbles As String
        Get
            Return _context.SP_Add_Bubbles
        End Get
        Set(value As String)
            _context.SP_Add_Bubbles = value
        End Set
    End Property


    Public Shared Property SP_BubblesExcel As String
        Get
            Return _context.SP_BubblesExcel
        End Get
        Set(value As String)
            _context.SP_BubblesExcel = value
        End Set
    End Property

    Public Shared Property SP_Add_Revisions As String
        Get
            Return _context.SP_Add_Revisions
        End Get
        Set(value As String)
            _context.SP_Add_Revisions = value
        End Set
    End Property
    Public Shared Property SP_MarkupRegex As String
        Get
            Return _context.SP_MarkupRegex
        End Get
        Set(value As String)
            _context.SP_MarkupRegex = value
        End Set
    End Property

    Public Shared Property SP_ChatWord As String
        Get
            Return _context.SP_ChatWord
        End Get
        Set(value As String)
            _context.SP_ChatWord = value
        End Set
    End Property

    Public Shared Property SP_Add_ChatWord_Commands As String
        Get
            Return _context.SP_Add_ChatWord_Commands
        End Get
        Set(value As String)
            _context.SP_Add_ChatWord_Commands = value
        End Set
    End Property

    Public Shared Property SP_ChatExcel As String
        Get
            Return _context.SP_ChatExcel
        End Get
        Set(value As String)
            _context.SP_ChatExcel = value
        End Set
    End Property

    Public Shared Property SP_Add_ChatExcel_Commands As String
        Get
            Return _context.SP_Add_ChatExcel_Commands
        End Get
        Set(value As String)
            _context.SP_Add_ChatExcel_Commands = value
        End Set
    End Property
    Public Shared Property INI_ChatCap As Integer
        Get
            Return _context.INI_ChatCap
        End Get
        Set(value As Integer)
            _context.INI_ChatCap = value
        End Set
    End Property



    Public Shared ReadOnly Property RDV As String = "Word (" & Version & ")"
    Public Shared Property DecodedAPI As String
        Get
            Return _context.DecodedAPI
        End Get
        Set(value As String)
            _context.DecodedAPI = value
        End Set
    End Property

    Public Shared Property DecodedAPI_2 As String
        Get
            Return _context.DecodedAPI_2
        End Get
        Set(value As String)
            _context.DecodedAPI_2 = value
        End Set
    End Property

    Public Shared Property TokenExpiry As DateTime
        Get
            Return _context.TokenExpiry
        End Get
        Set(value As DateTime)
            _context.TokenExpiry = value
        End Set
    End Property

    Public Shared Property TokenExpiry_2 As DateTime
        Get
            Return _context.TokenExpiry_2
        End Get
        Set(value As DateTime)
            _context.TokenExpiry_2 = value
        End Set
    End Property

    Public Shared Property Codebasis As String
        Get
            Return _context.Codebasis
        End Get
        Set(value As String)
            _context.Codebasis = value
        End Set
    End Property

    Public Shared Property GPTSetupError As Boolean
        Get
            Return _context.GPTSetupError
        End Get
        Set(value As Boolean)
            _context.GPTSetupError = value
        End Set
    End Property

    Public Shared Property INIloaded As Boolean
        Get
            Return _context.INIloaded
        End Get
        Set(value As Boolean)
            _context.INIloaded = value
        End Set
    End Property



    Public Shared Property INI_ISearch As Boolean
        Get
            Return _context.INI_ISearch
        End Get
        Set(value As Boolean)
            _context.INI_ISearch = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Approve As Boolean
        Get
            Return _context.INI_ISearch_Approve
        End Get
        Set(value As Boolean)
            _context.INI_ISearch_Approve = value
        End Set
    End Property

    Public Shared Property INI_ISearch_URL As String
        Get
            Return _context.INI_ISearch_URL
        End Get
        Set(value As String)
            _context.INI_ISearch_URL = value
        End Set
    End Property

    Public Shared Property INI_ISearch_ResponseURLStart As String
        Get
            Return _context.INI_ISearch_ResponseURLStart
        End Get
        Set(value As String)
            _context.INI_ISearch_ResponseURLStart = value
        End Set
    End Property

    Public Shared Property INI_ISearch_ResponseMask1 As String
        Get
            Return _context.INI_ISearch_ResponseMask1
        End Get
        Set(value As String)
            _context.INI_ISearch_ResponseMask1 = value
        End Set
    End Property

    Public Shared Property INI_ISearch_ResponseMask2 As String
        Get
            Return _context.INI_ISearch_ResponseMask2
        End Get
        Set(value As String)
            _context.INI_ISearch_ResponseMask2 = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Name As String
        Get
            Return _context.INI_ISearch_Name
        End Get
        Set(value As String)
            _context.INI_ISearch_Name = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Tries As Integer
        Get
            Return _context.INI_ISearch_Tries
        End Get
        Set(value As Integer)
            _context.INI_ISearch_Tries = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Results As Integer
        Get
            Return _context.INI_ISearch_Results
        End Get
        Set(value As Integer)
            _context.INI_ISearch_Results = value
        End Set
    End Property

    Public Shared Property INI_ISearch_MaxDepth As Integer
        Get
            Return _context.INI_ISearch_MaxDepth
        End Get
        Set(value As Integer)
            _context.INI_ISearch_MaxDepth = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Timeout As Long
        Get
            Return _context.INI_ISearch_Timeout
        End Get
        Set(value As Long)
            _context.INI_ISearch_Timeout = value
        End Set
    End Property

    Public Shared Property INI_ISearch_SearchTerm_SP As String
        Get
            Return _context.INI_ISearch_SearchTerm_SP
        End Get
        Set(value As String)
            _context.INI_ISearch_SearchTerm_SP = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Apply_SP_Markup As String
        Get
            Return _context.INI_ISearch_Apply_SP_Markup
        End Get
        Set(value As String)
            _context.INI_ISearch_Apply_SP_Markup = value
        End Set
    End Property
    Public Shared Property INI_ISearch_Apply_SP As String
        Get
            Return _context.INI_ISearch_Apply_SP
        End Get
        Set(value As String)
            _context.INI_ISearch_Apply_SP = value
        End Set
    End Property

    Public Shared Property INI_Lib As Boolean
        Get
            Return _context.INI_Lib
        End Get
        Set(value As Boolean)
            _context.INI_Lib = value
        End Set
    End Property

    Public Shared Property INI_Lib_File As String
        Get
            Return _context.INI_Lib_File
        End Get
        Set(value As String)
            _context.INI_Lib_File = value
        End Set
    End Property

    Public Shared Property INI_Lib_Timeout As Long
        Get
            Return _context.INI_Lib_Timeout
        End Get
        Set(value As Long)
            _context.INI_Lib_Timeout = value
        End Set
    End Property

    Public Shared Property INI_Lib_Find_SP As String
        Get
            Return _context.INI_Lib_Find_SP
        End Get
        Set(value As String)
            _context.INI_Lib_Find_SP = value
        End Set
    End Property

    Public Shared Property INI_Lib_Apply_SP_Markup As String
        Get
            Return _context.INI_Lib_Apply_SP_Markup
        End Get
        Set(value As String)
            _context.INI_Lib_Apply_SP_Markup = value
        End Set
    End Property

    Public Shared Property INI_Lib_Apply_SP As String
        Get
            Return _context.INI_Lib_Apply_SP
        End Get
        Set(value As String)
            _context.INI_Lib_Apply_SP = value
        End Set
    End Property

    Public Shared Property INI_Placeholder_01 As String
        Get
            Return _context.INI_Placeholder_01
        End Get
        Set(value As String)
            _context.INI_Placeholder_01 = value
        End Set
    End Property

    Public Shared Property INI_Placeholder_02 As String
        Get
            Return _context.INI_Placeholder_02
        End Get
        Set(value As String)
            _context.INI_Placeholder_02 = value
        End Set
    End Property

    Public Shared Property INI_Placeholder_03 As String
        Get
            Return _context.INI_Placeholder_03
        End Get
        Set(value As String)
            _context.INI_Placeholder_03 = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodHelper As Integer
        Get
            Return _context.INI_MarkupMethodHelper
        End Get
        Set(value As Integer)
            _context.INI_MarkupMethodHelper = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodWord As Integer
        Get
            Return _context.INI_MarkupMethodWord
        End Get
        Set(value As Integer)
            _context.INI_MarkupMethodWord = value
        End Set
    End Property

    Public Shared Property INI_ContextMenu As Boolean
        Get
            Return _context.INI_ContextMenu
        End Get
        Set(value As Boolean)
            _context.INI_ContextMenu = value
        End Set
    End Property

    Public Shared Property INI_UpdateCheckInterval As Integer
        Get
            Return _context.INI_UpdateCheckInterval
        End Get
        Set(value As Integer)
            _context.INI_UpdateCheckInterval = value
        End Set
    End Property

    Public Shared Property INI_UpdatePath As String
        Get
            Return _context.INI_UpdatePath
        End Get
        Set(value As String)
            _context.INI_UpdatePath = value
        End Set
    End Property

    Public Shared Property INI_SpeechModelPath As String
        Get
            Return _context.INI_SpeechModelPath
        End Get
        Set(value As String)
            _context.INI_SpeechModelPath = value
        End Set
    End Property

    Public Shared Property INI_LocalModelPath As String
        Get
            Return _context.INI_LocalModelPath
        End Get
        Set(value As String)
            _context.INI_LocalModelPath = value
        End Set
    End Property



    Public Shared Property INI_TTSEndpoint As String
        Get
            Return _context.INI_TTSEndpoint
        End Get
        Set(value As String)
            _context.INI_TTSEndpoint = value
        End Set
    End Property

    Public Shared Property INI_ShortcutsWordExcel As String
        Get
            Return _context.INI_ShortcutsWordExcel
        End Get
        Set(value As String)
            _context.INI_ShortcutsWordExcel = value
        End Set
    End Property

    Public Shared Property INI_PromptLib As Boolean
        Get
            Return _context.INI_PromptLib
        End Get
        Set(value As Boolean)
            _context.INI_PromptLib = value
        End Set
    End Property

    Public Shared Property INI_PromptLibPath As String
        Get
            Return _context.INI_PromptLibPath
        End Get
        Set(value As String)
            _context.INI_PromptLibPath = value
        End Set
    End Property

    Public Shared Property INI_PromptLibPath_Transcript As String
        Get
            Return _context.INI_PromptLibPath_Transcript
        End Get
        Set(value As String)
            _context.INI_PromptLibPath_Transcript = value
        End Set
    End Property

    Public Shared Property INI_AlternateModelPath As String
        Get
            Return _context.INI_AlternateModelPath
        End Get
        Set(value As String)
            _context.INI_AlternateModelPath = value
        End Set
    End Property

    Public Shared Property INI_SpecialServicePath As String
        Get
            Return _context.INI_SpecialServicePath
        End Get
        Set(value As String)
            _context.INI_SpecialServicePath = value
        End Set
    End Property


    Public Shared Property PromptLibrary() As List(Of String)
        Get
            Return _context.PromptLibrary
        End Get
        Set(value As List(Of String))
            _context.PromptLibrary = value
        End Set
    End Property

    Public Shared Property PromptTitles() As List(Of String)
        Get
            Return _context.PromptTitles
        End Get
        Set(value As List(Of String))
            _context.PromptTitles = value
        End Set
    End Property

    Public Shared Property MenusAdded As Boolean
        Get
            Return _context.MenusAdded
        End Get
        Set(value As Boolean)
            _context.MenusAdded = value
        End Set
    End Property

    Public Shared Property InitialConfigFailed As Boolean
        Get
            Return _context.InitialConfigFailed
        End Get
        Set(value As Boolean)
            _context.InitialConfigFailed = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter1 As String
        Get
            Return _context.INI_Model_Parameter1
        End Get
        Set(value As String)
            _context.INI_Model_Parameter1 = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter2 As String
        Get
            Return _context.INI_Model_Parameter2
        End Get
        Set(value As String)
            _context.INI_Model_Parameter2 = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter3 As String
        Get
            Return _context.INI_Model_Parameter3
        End Get
        Set(value As String)
            _context.INI_Model_Parameter3 = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter4 As String
        Get
            Return _context.INI_Model_Parameter4
        End Get
        Set(value As String)
            _context.INI_Model_Parameter4 = value
        End Set
    End Property

    Public Shared Property SP_MergePrompt As String
        Get
            Return _context.SP_MergePrompt
        End Get
        Set(value As String)
            _context.SP_MergePrompt = value
        End Set
    End Property
    Public Shared Property SP_Add_MergePrompt As String
        Get
            Return _context.SP_Add_MergePrompt
        End Get
        Set(value As String)
            _context.SP_Add_MergePrompt = value
        End Set
    End Property


#End Region

    ' Functions of SharedLibrary

#Region "SharedLibrary"
    Public Sub InitializeConfig(FirstTime As Boolean, Reload As Boolean)
        _context.InitialConfigFailed = False
        _context.RDV = "Word (" & Version & ")"
        SharedMethods.InitializeConfig(_context, FirstTime, Reload)
    End Sub
    Private Function INIValuesMissing()
        Return SharedMethods.INIValuesMissing(_context)
    End Function
    Public Shared Async Function PostCorrection(inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
        Return Await SharedMethods.PostCorrection(_context, inputText, UseSecondAPI)
    End Function
    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal Hidesplash As Boolean = False, Optional ByVal AddUserPrompt As String = "", Optional ByVal FileObject As String = "") As Task(Of String)
        Return Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, Hidesplash, AddUserPrompt, FileObject)
    End Function
    Private Function ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String))
        SharedMethods.ShowSettingsWindow(Settings, SettingsTips, _context)
    End Function
    Private Function ShowPromptSelector(filePath As String, enableMarkup As Boolean, enableBubbles As Boolean) As (String, Boolean, Boolean, Boolean)
        Return SharedMethods.ShowPromptSelector(filePath, enableMarkup, enableBubbles, _context)
    End Function

#End Region

    Public Enum CustomWdKey
        wdKeyUp = 38
        wdKeyDown = 40
        wdKeyLeft = 37
        wdKeyRight = 39
        wdKeySpace = 32
    End Enum

    Private automationObject As BridgeSubs

    Protected Overrides Function RequestComAddInAutomationService() As Object
        If automationObject Is Nothing Then
            automationObject = New BridgeSubs()
        End If
        Return automationObject
    End Function

    Public Sub InitializeAddInFeatures()
        InitializeConfig(True, True)
        AddContextMenu()
        UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, RDV, INI_UpdatePath)
    End Sub

#Region "Build Menu"
    Public Sub AddContextMenu()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

        If MenusAdded Then Return

        ' Remove existing context menus from relevant context menus
        If RemoveMenu Then
            RemoveOldContextMenu()
            RemoveMenu = False
        End If

        If Not INI_ContextMenu Then Return

        If Not VBAModuleWorking() Then Return

        If INIloaded = False Then Return

        MenusAdded = True

        ' List of relevant context menus
        Dim contextMenus As String() = {
        "Text", "Spelling", "Grammar", "Grammar (2)", "Linked Text", "Lists", "Headings", "Rotate Text", "Table Text",
"Footnotes", "Endnotes", "Frames", "Fields", "Form Fields", "Display Fields", "Field Display List Numbers", "Field AutoText",
"Comment", "Track Changes", "Track Changes Indicator", "Hyperlink Context Menu",
"Table Cells", "Whole Table", "Linked Table", "Table Lists", "Table Pictures",
"Inline Picture", "Floating Picture", "OLE Object", "ActiveX Control", "Inline ActiveX Control",
"Business Card", "Equation Popup", "WordArt Context Menu",
"Drop Caps", "Font Popup", "Font Paragraph", "Format consistency",
"Format Inspector Popup in Normal Mode", "Format Inspector Popup in Compare Mode", "AutoSignature Popup"
}
        Dim application As Word.Application = Globals.ThisAddIn.Application

        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Check if the menu already exists
                    If Not ContextMenuExists(cb, RIMenu) Then
                        Dim myControl As CommandBarPopup = Nothing
                        Try
                            myControl = CType(cb.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
                        Catch ex As System.Exception
                            ' Handle potential errors
                        End Try

                        If myControl IsNot Nothing Then
                            myControl.Caption = RIMenu
                            myControl.Visible = True
                            myControl.Enabled = True

                            ' Add submenu items
                            AddSubMenuItems(myControl)
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Function ContextMenuExists(cb As CommandBar, menuName As String) As Boolean
        For Each ctrl As CommandBarControl In cb.Controls
            If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = menuName Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Sub AddSubMenuItems(myControl As CommandBarPopup)

        Try
            Dim subControl As CommandBarButton
            Dim wordHelpersMenu As CommandBarPopup
            Dim improveMenu As CommandBarPopup
            Dim subSubControl As CommandBarButton
            Dim shortcutsArray() As String
            Dim shortcutPair() As String
            Dim shortcutDict As New Dictionary(Of String, String) ' Use native .NET Dictionary
            Dim i As Integer

            ' Parse the shortcuts from INI_ShortcutsWordExcel
            shortcutsArray = INI_ShortcutsWordExcel.Split(";")

            ' Populate the dictionary
            For i = 0 To shortcutsArray.Length - 1
                If shortcutsArray(i).Contains("=") Then
                    shortcutPair = shortcutsArray(i).Split("=")
                    shortcutDict(shortcutPair(0).Trim()) = shortcutPair(1).Trim()
                End If
            Next
            myControl.Visible = True

            ' Add menu items and assign shortcuts
            ' The OnAction refers to a Word Macro that has to be loaded as a helper for this to work; it will call up the BridgeSubs methods

            If Not String.IsNullOrWhiteSpace(INI_Language1) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language1
                subControl.FaceId = 6112
                subControl.Visible = True
                subControl.OnAction = "CallInLanguage1"
                If shortcutDict.ContainsKey(subControl.Caption) Then ' Check for key existence
                    subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption) ' Access the value
                End If
            End If

            If Not String.IsNullOrWhiteSpace(INI_Language2) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language2
                subControl.OnAction = "CallInLanguage2"
                subControl.FaceId = 6112

                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
                End If
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "To Other"
            subControl.OnAction = "CallInOther"
            subControl.FaceId = 6112
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Correct" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallCorrect"
            subControl.FaceId = 329

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            ' Create new submenu "Improve"
            improveMenu = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
            improveMenu.Caption = "Improve"

            ' Add submenu items to "Improve"
            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Improve" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallImprove"
            subSubControl.FaceId = 329
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "No Filler Words" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallNoFillers"
            subSubControl.FaceId = 4242
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "More Friendly" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallFriendly"
            subSubControl.FaceId = 59
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "More Convincing" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallConvincing"
            subSubControl.FaceId = 343
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If


            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Shorten" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallShorten"
            subControl.FaceId = 292
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Anonymize" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallAnonymize"
            subControl.FaceId = 7502
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Switch Party" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallSwitchParty"
            subControl.FaceId = 327
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Summarize"
            subControl.OnAction = "CallSummarize"
            subControl.FaceId = 602

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If
            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Freestyle"
            subControl.OnAction = "CallFreestyleNM"
            subControl.FaceId = 346
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            If INI_SecondAPI Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "Freestyle (" & INI_Model_2 & ")"
                subControl.OnAction = "CallFreestyleAM"
                subControl.FaceId = 346
                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
                End If
            End If


            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Context Search"
            subControl.OnAction = "CallContextSearch"
            subControl.FaceId = 46
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            ' Create new submenu "Word helpers"
            wordHelpersMenu = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
            wordHelpersMenu.Caption = "Word helpers"

            ' Add submenu items to "Word helpers"
            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Self-Compare Selection"
            subSubControl.OnAction = "CallCompareSelectionHalves"
            subSubControl.FaceId = 304
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Accept Format Changes"
            subSubControl.OnAction = "CallAcceptFormatting"
            subSubControl.FaceId = 161
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Markup Time Span"
            subSubControl.OnAction = "CallCalculateUserMarkupTimeSpan"
            subSubControl.FaceId = 33
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Regex Search && Replace"
            subSubControl.OnAction = "CallRegexSearchReplace"
            subSubControl.FaceId = 288
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Import Text File"
            subSubControl.OnAction = "CallImportTextFile"
            subSubControl.FaceId = 2311
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            If Not String.IsNullOrWhiteSpace(INI_ShortcutsWordExcel) Then

                ' Assign shortcuts using the dictionary
                If Not String.IsNullOrWhiteSpace(INI_Language1) Then AssignShortcut("To " & INI_Language1, "CallInLanguage1", shortcutDict)
                If Not String.IsNullOrWhiteSpace(INI_Language2) Then AssignShortcut("To " & INI_Language2, "CallInLanguage2", shortcutDict)
                AssignShortcut("To Other", "CallInOther", shortcutDict)
                AssignShortcut("Correct (Markup)", "CallCorrect", shortcutDict)
                AssignShortcut("Correct", "CallCorrect", shortcutDict)
                AssignShortcut("Improve (Markup)", "CallImprove", shortcutDict)
                AssignShortcut("Improve", "CallImprove", shortcutDict)
                AssignShortcut("No Filler Words (Markup)", "CallNoFillers", shortcutDict)
                AssignShortcut("No Filler Words", "CallNoFillers", shortcutDict)
                AssignShortcut("More Friendly (Markup)", "CallFriendly", shortcutDict)
                AssignShortcut("More Friendly", "CallFriendly", shortcutDict)
                AssignShortcut("More Convincing (Markup)", "CallConvincing", shortcutDict)
                AssignShortcut("More Convincing", "CallConvincing", shortcutDict)
                AssignShortcut("Shorten (Markup)", "CallShorten", shortcutDict)
                AssignShortcut("Shorten", "CallShorten", shortcutDict)
                AssignShortcut("Anonymize (Markup)", "CallAnonymize", shortcutDict)
                AssignShortcut("Anonymize", "CallAnonymize", shortcutDict)
                AssignShortcut("Switch Party (Markup)", "CallSwitchParty", shortcutDict)
                AssignShortcut("Switch Party", "CallSwitchParty", shortcutDict)
                AssignShortcut("Summarize", "CallSummarize", shortcutDict)
                AssignShortcut("Freestyle", "CallFreestyleNM", shortcutDict)
                AssignShortcut("Context Search", "CallContextSearch", shortcutDict)

                ' Assign shortcuts for second API if applicable
                If INI_SecondAPI Then
                    AssignShortcut("Freestyle (" & INI_Model_2 & ")", "CallFreestyleAM", shortcutDict)
                End If

                ' Assign shortcuts for submenu "Word helpers"
                AssignShortcut("Self-Compare Selection", "CallCompareSelectionHalves", shortcutDict)
                AssignShortcut("Accept Format Changes", "CallAcceptFormatting", shortcutDict)
                AssignShortcut("Markup Time Span", "CallCalculateUserMarkupTimeSpan", shortcutDict)
                AssignShortcut("Regex Search & Replace", "CallRegexSearchReplace", shortcutDict)
                AssignShortcut("Regex Search && Replace", "CallRegexSearchReplace", shortcutDict)
                AssignShortcut("Import Text File", "CallImportTextFile", shortcutDict)

            End If
        Catch ex As System.Exception

        End Try
    End Sub
    Public Sub AssignShortcut(ByVal controlName As String, ByVal macro As String, ByRef shortcutDict As Dictionary(Of String, String))
        Dim shortcutKey As String
        Dim keyCode As Long
        Try
            ' Check if there is a shortcut assigned for this menu item
            If shortcutDict.ContainsKey(controlName) Then
                shortcutKey = shortcutDict(controlName)
            Else
                Return ' No shortcut assigned
            End If

            ' Build KeyCode from shortcutKey text
            keyCode = BuildKeyCodeFromText(shortcutKey)

            If keyCode > 0 Then
                Globals.ThisAddIn.Application.CustomizationContext = Globals.ThisAddIn.Application.NormalTemplate
                Globals.ThisAddIn.Application.KeyBindings.Add(KeyCode:=keyCode, KeyCategory:=WdKeyCategory.wdKeyCategoryMacro, Command:=macro)
            End If
        Catch ex As System.Exception
            ' Handle exceptions gracefully
            ' Debug.WriteLine("Error in AssignShortcut " & ex.Message)
        End Try
    End Sub

    Public Function BuildKeyCodeFromText(ByVal shortcutKey As String) As Long
        Dim parts() As String
        Dim keysCollection As New List(Of Integer)()
        Dim keyCode As Long = 0

        Try
            parts = shortcutKey.Split("-"c)

            For Each part As String In parts
                Select Case part.Trim().ToUpper()
                    Case "CTRL"
                        keysCollection.Add(WdKey.wdKeyControl)
                    Case "SHIFT"
                        keysCollection.Add(WdKey.wdKeyShift)
                    Case "ALT"
                        keysCollection.Add(WdKey.wdKeyAlt)

                ' Map digits directly
                    Case "0"
                        keysCollection.Add(WdKey.wdKey0)
                    Case "1"
                        keysCollection.Add(WdKey.wdKey1)
                    Case "2"
                        keysCollection.Add(WdKey.wdKey2)
                    Case "3"
                        keysCollection.Add(WdKey.wdKey3)
                    Case "4"
                        keysCollection.Add(WdKey.wdKey4)
                    Case "5"
                        keysCollection.Add(WdKey.wdKey5)
                    Case "6"
                        keysCollection.Add(WdKey.wdKey6)
                    Case "7"
                        keysCollection.Add(WdKey.wdKey7)
                    Case "8"
                        keysCollection.Add(WdKey.wdKey8)
                    Case "9"
                        keysCollection.Add(WdKey.wdKey9)

                ' Map function keys directly
                    Case "F1"
                        keysCollection.Add(WdKey.wdKeyF1)
                    Case "F2"
                        keysCollection.Add(WdKey.wdKeyF2)
                    Case "F3"
                        keysCollection.Add(WdKey.wdKeyF3)
                    Case "F4"
                        keysCollection.Add(WdKey.wdKeyF4)
                    Case "F5"
                        keysCollection.Add(WdKey.wdKeyF5)
                    Case "F6"
                        keysCollection.Add(WdKey.wdKeyF6)
                    Case "F7"
                        keysCollection.Add(WdKey.wdKeyF7)
                    Case "F8"
                        keysCollection.Add(WdKey.wdKeyF8)
                    Case "F9"
                        keysCollection.Add(WdKey.wdKeyF9)
                    Case "F10"
                        keysCollection.Add(WdKey.wdKeyF10)
                    Case "F11"
                        keysCollection.Add(WdKey.wdKeyF11)
                    Case "F12"
                        keysCollection.Add(WdKey.wdKeyF12)

                ' Navigation and special keys
                    Case "LEFT"
                        keysCollection.Add(CustomWdKey.wdKeyLeft)
                    Case "RIGHT"
                        keysCollection.Add(CustomWdKey.wdKeyRight)
                    Case "UP"
                        keysCollection.Add(CustomWdKey.wdKeyUp)
                    Case "DOWN"
                        keysCollection.Add(CustomWdKey.wdKeyDown)
                    Case "HOME"
                        keysCollection.Add(WdKey.wdKeyHome)
                    Case "END"
                        keysCollection.Add(WdKey.wdKeyEnd)
                    Case "PAGEUP"
                        keysCollection.Add(WdKey.wdKeyPageUp)
                    Case "PAGEDOWN"
                        keysCollection.Add(WdKey.wdKeyPageDown)
                    Case "ESC"
                        keysCollection.Add(WdKey.wdKeyEsc)
                    Case "TAB"
                        keysCollection.Add(WdKey.wdKeyTab)
                    Case "BACKSPACE"
                        keysCollection.Add(WdKey.wdKeyBackspace)
                    Case "DELETE"
                        keysCollection.Add(WdKey.wdKeyDelete)
                    Case "INSERT"
                        keysCollection.Add(WdKey.wdKeyInsert)
                    Case "SPACE"
                        keysCollection.Add(CustomWdKey.wdKeySpace)

                ' Letters mapped directly
                    Case "A" To "Z"
                        keysCollection.Add([Enum].Parse(GetType(WdKey), "wdKey" & part.Trim().ToUpper()))
                    Case Else
                        ' Unknown key
                        Return 0
                End Select
            Next

            ' Build the KeyCode using Application.BuildKeyCode
            Select Case keysCollection.Count
                Case 1
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0))
                Case 2
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0), keysCollection(1))
                Case 3
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0), keysCollection(1), keysCollection(2))
                Case 4
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0), keysCollection(1), keysCollection(2), keysCollection(3))
                Case Else
                    ' Unknown key
                    Return 0
            End Select

            'Debug.WriteLine("Shortcutkey " & shortcutKey & "  Keycode: " & keyCode)

            Return keyCode

        Catch ex As System.Exception
            ' Handle errors gracefully
            Return 0
        End Try
    End Function

    Public Sub RemoveOldContextMenu()
        Dim application As Word.Application = Globals.ThisAddIn.Application

        ' Array of relevant context menus
        Dim contextMenus As String() = {
"Text", "Spelling", "Grammar", "Grammar (2)", "Linked Text", "Lists", "Headings", "Rotate Text", "Table Text",
"Footnotes", "Endnotes", "Frames", "Fields", "Form Fields", "Display Fields", "Field Display List Numbers", "Field AutoText",
"Comment", "Track Changes", "Track Changes Indicator", "Hyperlink Context Menu",
"Table Cells", "Whole Table", "Linked Table", "Table Lists", "Table Pictures",
"Inline Picture", "Floating Picture", "OLE Object", "ActiveX Control", "Inline ActiveX Control",
"Business Card", "Equation Popup", "WordArt Context Menu",
"Drop Caps", "Font Popup", "Font Paragraph", "Format consistency",
"Format Inspector Popup in Normal Mode", "Format Inspector Popup in Compare Mode", "AutoSignature Popup"
}

        ' Iterate through all CommandBars
        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Remove the context menu if it exists
                    For Each ctrl As CommandBarControl In cb.Controls
                        If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = RIMenu Then
                            Try
                                ctrl.Delete()
                            Catch ex As System.Exception
                                ' Handle errors if needed, e.g., logging
                                'Debug.WriteLine($"Error removing control {ex.Message}")
                            End Try
                        End If
                    Next
                End If
            End If
        Next
    End Sub

    Public Sub RemoveVeryOldContextMenu()
        Dim application As Word.Application = Globals.ThisAddIn.Application

        ' Array of relevant context menus
        Dim contextMenus As String() = {
"Text", "Spelling", "Grammar", "Grammar (2)", "Linked Text", "Lists", "Headings", "Rotate Text", "Table Text",
"Footnotes", "Endnotes", "Frames", "Fields", "Form Fields", "Display Fields", "Field Display List Numbers", "Field AutoText",
"Comment", "Track Changes", "Track Changes Indicator", "Hyperlink Context Menu",
"Table Cells", "Whole Table", "Linked Table", "Table Lists", "Table Pictures",
"Inline Picture", "Floating Picture", "OLE Object", "ActiveX Control", "Inline ActiveX Control",
"Business Card", "Equation Popup", "WordArt Context Menu",
"Drop Caps", "Font Popup", "Font Paragraph", "Format consistency",
"Format Inspector Popup in Normal Mode", "Format Inspector Popup in Compare Mode", "AutoSignature Popup"
}

        ' Iterate through all CommandBars
        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Remove the context menu if it exists
                    For Each ctrl As CommandBarControl In cb.Controls
                        If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = OldRIMenu Then
                            Try
                                ctrl.Delete()
                            Catch ex As System.Exception
                                ' Handle errors if needed, e.g., logging
                                'Debug.WriteLine($"Error removing control {ex.Message}")
                            End Try
                        End If
                    Next
                End If
            End If
        Next
    End Sub

#End Region


    ' Declare them publicly so that InterpolateAtRuntime can access them; case-sensitive

    Public TranslateLanguage As String
    Public ShortenLength As Double
    Public SummaryLength As Integer
    Public OtherPrompt As String = ""
    Public SearchTerms As String
    Public SearchContext As String
    Public CurrentDate As String
    Public SysPrompt As String
    Public OldParty, NewParty As String
    Public SelectedText As String
    Public LibraryText As String
    Public LibResult As String
    Public SearchResult As String
    Public doc As String
    Public HostName As String
    Public GuestName As String
    Public Language As String
    Public Duration As String
    Public TargetAudience As String
    Public DialogueContext As String
    Public ExtraInstructions As String

    Private chatForm As frmAIChat

    Public Sub Transcriptor()
        If INILoadFail() Then Return
        If Not String.IsNullOrEmpty(INI_SpeechModelPath) Then
            Dim SpeechPath As String = ExpandEnvironmentVariables(INI_SpeechModelPath)
            If Not String.IsNullOrEmpty(SpeechPath) AndAlso Not SpeechPath.EndsWith("\") Then
                SpeechPath = SpeechPath & "\"
            End If
            Dim currentPath As String = Environment.GetEnvironmentVariable("PATH")

            If Not currentPath.Contains(SpeechPath) Then
                Environment.SetEnvironmentVariable("PATH", currentPath & ";" & SpeechPath)
            End If
            RuntimeOptions.LibraryPath = SpeechPath
            'RuntimeOptions.RuntimeLibraryOrder = New List(Of RuntimeLibrary) From {RuntimeLibrary.Cuda, RuntimeLibrary.Cpu}

        End If

        Dim TranscriptionForm = New TranscriptionForm()
        TranscriptionForm.Show()
    End Sub

    Public Sub ShowChatForm()
        If INILoadFail() Then Return
        If chatForm Is Nothing OrElse chatForm.IsDisposed Then
            chatForm = New frmAIChat(_context)

            ' Set the location and size before showing the form
            If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
                chatForm.StartPosition = FormStartPosition.Manual
                chatForm.Location = My.Settings.FormLocation
                chatForm.Size = My.Settings.FormSize
            Else
                ' Default to center screen if no settings are available
                chatForm.StartPosition = FormStartPosition.Manual
                Dim screenBounds As System.Drawing.Rectangle = Screen.PrimaryScreen.WorkingArea
                chatForm.Location = New System.Drawing.Point((screenBounds.Width - chatForm.Width) \ 2, (screenBounds.Height - chatForm.Height) \ 2)
                chatForm.Size = New Size(650, 500) ' Set default size if needed
            End If
        End If

        ' Show and bring the form to the front
        chatForm.Show()
        chatForm.BringToFront()
    End Sub

    Public Function INILoadFail() As Boolean
        If Not INIloaded Then
            If Not StartupInitialized Then
                DelayedStartupTasks()
                RemoveStartupHandlers()
                If Not INIloaded Then Return True
                Return False
            Else
                InitializeConfig(False, False)
                If Not INIloaded Then
                    Return True
                End If
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Async Sub ImageGenerator()

        OtherPrompt = SLib.ShowCustomInputBox("Describe the image to generate (the image will be saved to the desktop):", $"{AN} Image Generator", True)

        If String.IsNullOrWhiteSpace(OtherPrompt) Then
            Return
        End If

        Dim result2 As String = Await LLM("You are an image generator. You will complete the following command.", "Create the following image: " & OtherPrompt, "", "", 0, True)

        ShowCustomMessageBox(result2)

    End Sub

    Public Async Sub InLanguage1()

        If INILoadFail() Then Return
        TranslateLanguage = INI_Language1
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, False, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub InLanguage2()

        If INILoadFail() Then Return
        TranslateLanguage = INI_Language2
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, False, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub InOther()
        If INILoadFail() Then Return
        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, False, False, False, True, False, INI_KeepFormatCap)
        End If
    End Sub

    Public Async Sub Correct()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Correct), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Improve()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Improve), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Friendly()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Friendly), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Convincing()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Convincing), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub NoFillers()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_NoFillers), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Anonymize()
        If INILoadFail() Then Return

        Dim DoMarkup As Boolean = INI_DoMarkupWord
        Dim DoReplace As Boolean = INI_ReplaceText2
        If Not DoMarkup Or Not DoReplace Then
            Dim result2 As Integer = ShowCustomYesNoBox($"As per your current settings no markup will be applied. For anonymizing a larger text, doing a markup may be a better choice. How Do you want To Continue?", "Continue As Is", "Continue With a markup")
            If result2 = 2 Then
                DoMarkup = True
            End If
        End If

        Dim MarkupMethod As Integer = INI_MarkupMethodWord
        If INI_DoMarkupWord And MarkupMethod <> 4 Then
            Dim MarkupNow As String = ""
            Select Case INI_MarkupMethodWord
                Case 1
                    MarkupNow = "Word markup method"
                Case 2
                    MarkupNow = "Diff markup method"
                Case 3
                    MarkupNow = "Diff markup method (with the output in a separate window)"
            End Select

            Dim result2 As Integer = ShowCustomYesNoBox($"You have chosen the {MarkupNow}. If you are anonymizing a larger text, the 'Regex' markup method may be a better choice. How do you want to continue?", "Continue as is", "Use Regex")
            If result2 = 2 Then
                MarkupMethod = 4
                DoReplace = True
            End If
        End If


        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Anonymize), True, INI_KeepFormat2, INI_KeepParaFormatInline, DoReplace, DoMarkup, MarkupMethod, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Explain()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Explain), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, True, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub SuggestTitles()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_SuggestTitles), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, True, False, True, False, INI_KeepFormatCap)
    End Sub

    Public Async Sub InsertClipboard()

        If String.IsNullOrWhiteSpace(INI_APICall_Object) Then
            ShowCustomMessageBox($"Your model ({INI_Model}) is not configured to process clipboard data (i.e. binary objects).")
            Return
        End If

        With Globals.ThisAddIn.Application.Selection
            If .Start <> .End Then
                .Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            End If
        End With

        If INILoadFail() Then Return

        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_InsertClipboard), False, False, False, False, False, 0, False, False, False, False, 0, False, "", False, "clipboard")

        If result <> "" Then
            Globals.ThisAddIn.Application.Selection.TypeParagraph()
            Globals.ThisAddIn.Application.Selection.TypeParagraph()
            InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & result & vbCrLf, False)
        End If

    End Sub


    Public Async Sub Shorten()

        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection

        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        Dim Textlength As Integer = GetSelectedTextLength()
        Dim UserInput As String
        Dim ShortenPercentValue As Integer = 0
        Do
            UserInput = SLib.ShowCustomInputBox("Enter the percentage by which your text should be shortened (it has " & Textlength & " words; " & ShortenPercent & "% will cut approx. " & (Textlength * ShortenPercent / 100) & " words)", $"{AN} Shortener", True, CStr(ShortenPercent) & "%").Trim()
            If String.IsNullOrEmpty(UserInput) Then
                Return
            End If
            UserInput = UserInput.Replace("%", "").Trim()
            If Integer.TryParse(UserInput, ShortenPercentValue) AndAlso ShortenPercentValue >= 1 AndAlso ShortenPercentValue <= 99 Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid percentage between 1 And 99.")
            End If
        Loop
        If ShortenPercentValue = 0 Then Return
        ShortenLength = (Textlength - (Textlength * (100 - ShortenPercentValue) / 100))
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Improve), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub SwitchParty()
        If INILoadFail() Then Return
        Dim UserInput As String
        Do
            UserInput = SLib.ShowCustomInputBox("Please provide the original party name And the New party name, separated by a comma (example: Elvis Presley, Taylor Swift):", $"{AN} Switch Party", True).Trim()

            If String.IsNullOrEmpty(UserInput) Then
                Return
            End If

            Dim parts() As String = UserInput.Split(","c)
            If parts.Length = 2 Then
                OldParty = parts(0).Trim()
                NewParty = parts(1).Trim()
                Exit Do
            Else
                ShowCustomMessageBox("Please enter two names separated by a comma.")
            End If
        Loop

        Dim DoMarkup As Boolean = INI_DoMarkupWord
        Dim DoReplace As Boolean = INI_ReplaceText2
        If Not DoMarkup Or Not DoReplace Then
            Dim result2 As Integer = ShowCustomYesNoBox($"As per your current settings no markup will be applied. For using 'Switch Party' on a larger texts, markup may be a better choice. How do you want to continue?", "Continue as is", "Continue with a markup")
            If result2 = 2 Then
                DoMarkup = True
                DoReplace = True
            End If
        End If

        Dim MarkupMethod As Integer = INI_MarkupMethodWord
        If INI_DoMarkupWord And MarkupMethod <> 4 Then
            Dim MarkupNow As String = ""
            Select Case INI_MarkupMethodWord
                Case 1
                    MarkupNow = "Word markup method"
                Case 2
                    MarkupNow = "Diff markup method"
                Case 3
                    MarkupNow = "Diff markup method (with the output in a separate window)"
            End Select

            Dim result2 As Integer = ShowCustomYesNoBox($"You have chosen the {MarkupNow}. If you are using 'Switch Party' with a larger text, the 'Regex' markup method may be a better choice. How do you want to continue?", "Continue as is", "Use Regex")
            If result2 = 2 Then
                MarkupMethod = 4
                DoReplace = True
            End If
        End If

        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_SwitchParty), True, INI_KeepFormat2, INI_KeepParaFormatInline, DoReplace, DoMarkup, MarkupMethod, False, False, True, False, INI_KeepFormatCap)

    End Sub
    Public Async Sub Summarize()
        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection

        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        Dim Textlength As Integer = GetSelectedTextLength()

        Dim UserInput As String
        SummaryLength = 0

        Do
            UserInput = SLib.ShowCustomInputBox("Enter the number of words your summary shall have (the selected text has " & Textlength & " words; the proposal " & SummaryPercent & "%):", $"{AN} Summarizer", True, CStr(CInt(SummaryPercent * Textlength / 100))).Trim()

            If String.IsNullOrEmpty(UserInput) Then
                Return
            End If

            If Integer.TryParse(UserInput, SummaryLength) AndAlso SummaryLength >= 1 AndAlso SummaryLength <= Textlength Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid word count between 1 and " & Textlength & ".")
            End If
        Loop
        If SummaryLength = 0 Then Return

        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Summarize), False, False, False, False, False, False, True, False, True, False, 0)
    End Sub

    Public Async Sub CreatePodcast()
        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection

        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        HostName = My.Settings.Hostname
        GuestName = My.Settings.Guestname
        TargetAudience = My.Settings.TargetAudience
        Duration = My.Settings.Duration
        Language = My.Settings.Language
        DialogueContext = My.Settings.DialogueContext
        ExtraInstructions = My.Settings.ExtraInstructions

        Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Host name", HostName),
                    New SLib.InputParameter("Guest name", GuestName),
                    New SLib.InputParameter("Target audience", TargetAudience),
                    New SLib.InputParameter("Context, background info", DialogueContext),
                    New SLib.InputParameter("Target length", Duration),
                    New SLib.InputParameter("Language of dialogue", Language),
                    New SLib.InputParameter("Extra instructions", ExtraInstructions)
                    }

        If ShowCustomVariableInputForm("Please enter the following parameters to take into account when creating your podcast script:", $"Create Podcast Script", params) Then

            HostName = params(0).Value.ToString()
            GuestName = params(1).Value.ToString()
            TargetAudience = params(2).Value.ToString()
            DialogueContext = params(3).Value.ToString()
            Duration = params(4).Value.ToString()
            Language = params(5).Value.ToString()
            ExtraInstructions = params(6).Value.ToString()

            My.Settings.Hostname = HostName
            My.Settings.Guestname = GuestName
            My.Settings.TargetAudience = TargetAudience
            My.Settings.DialogueContext = DialogueContext
            My.Settings.Duration = Duration
            My.Settings.Language = Language
            My.Settings.ExtraInstructions = ExtraInstructions
            My.Settings.Save()

            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Podcast), True, False, False, False, False, 3, True, False, True, False, 0, False, "", True)

        End If

    End Sub

    Public Async Sub CreateAudio()
        If INILoadFail() Then Return

        DetectTTSEngines()

        ' — Nothing at all? bail out —
        If Not TTS_googleAvailable AndAlso Not TTS_openAIAvailable Then
            Return   ' no TTS provider configured
        End If


        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection
        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If
        SelectedText = selection.Text.Trim()
        If SelectedText.Contains("H: ") And SelectedText.Contains("G: ") Then
            ReadPodcast(SelectedText)
        Else
            If selection.Text.Trim().StartsWith("{") Then
                Dim selectedoutputpath As String = (If(String.IsNullOrEmpty(My.Settings.TTSOutputPath), Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile), My.Settings.TTSOutputPath))
                selectedoutputpath = ShowCustomInputBox("Where should the audio generated from your JSON TTS file be saved to?", $"{AN} Create Audiobook", True, selectedoutputpath)
                If String.IsNullOrWhiteSpace(selectedoutputpath) Then
                    ' Use default path (Desktop) with default filename
                    selectedoutputpath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
                ElseIf selectedoutputpath.EndsWith("\") OrElse selectedoutputpath.EndsWith("/") Then
                    ' If only a folder is given, append default filename
                    selectedoutputpath = Path.Combine(selectedoutputpath, TTSDefaultFile)
                Else
                    Dim dir As String = Path.GetDirectoryName(selectedoutputpath)
                    Dim fileName As String = Path.GetFileName(selectedoutputpath)

                    ' If no directory is found, assume Desktop as the base
                    If String.IsNullOrWhiteSpace(dir) Then
                        selectedoutputpath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName)
                        dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    End If

                    ' If no filename is given, use the default filename
                    If String.IsNullOrWhiteSpace(fileName) Then
                        selectedoutputpath = Path.Combine(dir, TTSDefaultFile)
                    End If

                    ' Ensure the filename has ".mp3" extension
                    If Not fileName.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) Then
                        selectedoutputpath = Path.Combine(dir, fileName & ".mp3")
                    End If
                End If
                GenerateAndPlayAudio(selection.Text, selectedoutputpath, "", "")
                Return
            Else
                Dim Voices As Integer = ShowCustomYesNoBox("Do you want to use alternate voices to read the text?", "No, one voice", "Yes, alternate", "Create Audio")
                If Voices = 0 Then Return
                Using frm As New TTSSelectionForm("Select the voice you wish To use For creating your audio file And configure where To save it.", $"{AN} Text-To-Speech - Select Voices", Voices = 2) ' TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voice you wish To use For creating your audio file And configure where To save it.", $"{AN} Text-To-Speech - Select Voices", Voices = 2)
                    If frm.ShowDialog() = DialogResult.OK Then
                        Dim selectedVoices As List(Of String) = frm.SelectedVoices
                        Dim selectedLanguage As String = frm.SelectedLanguage
                        Dim outputPath As String = frm.SelectedOutputPath
                        GenerateAndPlayAudioFromSelectionParagraphs(outputPath, selectedLanguage, selectedVoices(0).Replace(" (male)", "").Replace(" (female)", ""), If(Voices = 2, selectedVoices(1).Replace(" (male)", "").Replace(" (female)", ""), ""))
                    End If
                End Using
            End If
        End If
    End Sub
    Public Async Sub FreeStyleNM()
        If INILoadFail() Then Return
        FreeStyle(False)
    End Sub
    Public Async Sub FreeStyleAM()
        If INILoadFail() Then Return

        If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then

            If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                originalConfigLoaded = False
                Return
            End If

        End If

        FreeStyle(True)

    End Sub


    Public Async Sub SpecialModel()
        Try
            If INILoadFail() Then Return

            Dim DoPane As Boolean = True

            If String.IsNullOrWhiteSpace(INI_SpecialServicePath) Then
                ShowCustomMessageBox("No special service path is configured.")
                Return
            End If

            If INILoadFail() Then Return
            Dim application As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Word.Selection = application.Selection

            If selection.Type = Word.WdSelectionType.wdSelectionIP Then
                ShowCustomMessageBox("Please select the text to be processed.")
                Return
            End If

            OptionChecked = False

            If Not ShowModelSelection(_context, INI_SpecialServicePath, "Special Service", "Select the special service you want to query:", "Output in a pane (not directly in the document)", 2) Then
                originalConfigLoaded = False
                Return
            End If

            Dim iniValues() As String = {INI_Model_Parameter1, INI_Model_Parameter2, INI_Model_Parameter3, INI_Model_Parameter4}
            Dim parameterDefs As New List(Of SharedLibrary.SharedLibrary.SharedMethods.InputParameter)()
            Dim typesList As New List(Of String)()
            Dim rangesList As New List(Of Tuple(Of Integer, Integer))()
            Dim optsDisplayList As New List(Of List(Of String))()
            Dim optsCodeList As New List(Of List(Of String))()

            For Each raw As String In iniValues
                If String.IsNullOrWhiteSpace(raw) Then Continue For
                Dim segments = raw.Split(";"c).Select(Function(s) s.Trim()).ToArray()
                Dim desc = segments(0)
                Dim t As String = segments(1).ToLowerInvariant()
                Dim defaultStr = segments(2)

                ' Range-Parsing (nur bei String-Typ): "min-max"
                Dim rangeTuple As Tuple(Of Integer, Integer) = Nothing
                Dim optsRaw As List(Of String) = Nothing

                ' Neu: nur numerische Typen
                If (t = "integer" OrElse t = "long" OrElse t = "double") _
                           AndAlso segments.Length > 3 _
                           AndAlso System.Text.RegularExpressions.Regex.IsMatch(segments(3), "^\d+-\d+$") Then

                    Dim parts = segments(3).Split("-"c)
                    Dim minVal = Integer.Parse(parts(0))
                    Dim maxVal = Integer.Parse(parts(1))
                    rangeTuple = Tuple.Create(minVal, maxVal)

                    ' Falls noch Auswahl-Optionen im Feld 4 stehen …
                    If segments.Length > 4 Then
                        optsRaw = segments(4).Split(","c).Select(Function(o) o.Trim()).ToList()
                    End If
                End If

                If t = "string" AndAlso segments.Length > 3 Then
                    optsRaw = segments(3).
                                  Split(","c).
                                  Select(Function(o) o.Trim()).
                                  ToList()
                End If

                ' Aufteilen der rohen Optionen in Display-Text und internen Code
                Dim displayList As List(Of String) = Nothing
                Dim codeList As List(Of String) = Nothing
                If optsRaw IsNot Nothing Then
                    displayList = New List(Of String)()
                    codeList = New List(Of String)()
                    For Each o As String In optsRaw
                        Dim lbl = o
                        Dim code = o
                        Dim idx1 = o.IndexOf("<"c)
                        Dim idx2 = o.IndexOf(">"c)
                        If idx1 >= 0 AndAlso idx2 > idx1 Then
                            lbl = o.Substring(0, idx1).Trim()
                            code = o.Substring(idx1 + 1, idx2 - idx1 - 1).Trim()
                        End If
                        displayList.Add(lbl)
                        codeList.Add(code)
                    Next
                End If

                ' Default-Wert für das Formular: Display-Text anstelle des Codes
                Dim defaultDisplay As Object = defaultStr
                If codeList IsNot Nothing Then
                    Dim idxDef = codeList.IndexOf(defaultStr)
                    If idxDef >= 0 Then defaultDisplay = displayList(idxDef)
                End If

                ' Typumwandlung für den Default-Wert
                Dim val As Object
                Select Case t
                    Case "boolean"
                        Dim b As Boolean
                        Boolean.TryParse(defaultStr, b)
                        val = b
                    Case "integer"
                        Dim i As Integer
                        Integer.TryParse(defaultStr, i)
                        val = i
                    Case "long"
                        Dim l As Long
                        Long.TryParse(defaultStr, l)
                        val = l
                    Case "double"
                        Dim d As Double
                        Double.TryParse(defaultStr, d)
                        val = d
                    Case Else
                        val = defaultDisplay
                End Select

                ' ParameterDefs für ShowCustomVariableInputForm aufbauen
                If displayList IsNot Nothing Then
                    parameterDefs.Add(New SharedLibrary.SharedLibrary.SharedMethods.InputParameter(desc, val, displayList))
                Else
                    parameterDefs.Add(New SharedLibrary.SharedLibrary.SharedMethods.InputParameter(desc, val))
                End If

                ' Metadaten merken
                typesList.Add(t)
                rangesList.Add(rangeTuple)
                optsDisplayList.Add(displayList)
                optsCodeList.Add(codeList)
            Next

            OtherPrompt = ""

            If parameterDefs.Count > 0 Then
                Dim parameters() As SharedLibrary.SharedLibrary.SharedMethods.InputParameter = parameterDefs.ToArray()
                If ShowCustomVariableInputForm("Please configure your parameters:", "Use '" & INI_Model_2 & "'", parameters) Then

                    ' === NEU: Werte auslesen mit Range-Clamping und Mapping ===
                    For i As Integer = 0 To parameters.Length - 1
                        Dim p As SharedLibrary.SharedLibrary.SharedMethods.InputParameter = parameters(i)
                        Dim rawValue As String = p.Value.ToString().Trim()
                        Dim t As String = typesList(i)
                        Dim range As Tuple(Of Integer, Integer) = rangesList(i)
                        Dim dispList = optsDisplayList(i)
                        Dim codeList = optsCodeList(i)

                        ' paramValue muss vor allen Verzweigungen deklariert werden
                        Dim paramValue As String

                        ' 1) Boolean → "true"/"false"
                        If TypeOf p.Value Is System.Boolean Then
                            paramValue = CType(p.Value, System.Boolean).ToString().ToLowerInvariant()

                        Else
                            ' 2) Range-Clamping bei numerischen Typen
                            If (t = "integer" OrElse t = "long" OrElse t = "double") AndAlso range IsNot Nothing Then

                                Dim num As Double
                                If Double.TryParse(rawValue, num) Then
                                    ' Clamp auf [Min;Max]
                                    num = Math.Max(range.Item1, Math.Min(range.Item2, num))

                                    ' Für Integer/Long als Ganzzahl zurückgeben
                                    If t = "integer" OrElse t = "long" Then
                                        rawValue = CInt(Math.Round(num)).ToString()
                                    Else
                                        rawValue = num.ToString()
                                    End If
                                End If
                            End If


                            ' 3) Mapping von Display-Text → interner Code
                            If dispList IsNot Nothing Then
                                Dim idx As Integer = dispList.IndexOf(rawValue)
                                If idx >= 0 Then
                                    paramValue = codeList(idx)
                                Else
                                    ' Fallback: unverändert
                                    paramValue = rawValue
                                End If

                            Else
                                ' 4) Normaler String-Fall: (all)/(alle)/--- filtern
                                Dim rvLower As String = rawValue.ToLowerInvariant()
                                If rvLower.Contains("(all)") OrElse rvLower.Contains("(alle)") OrElse rawValue.Contains("---") Then
                                    rawValue = ""
                                End If
                                paramValue = rawValue
                            End If
                        End If

                        ' 5) Sonderfall Prompt
                        If p.Name.ToLowerInvariant().Contains("prompt") Then
                            OtherPrompt = paramValue
                        End If

                        ' 6) Platzhalter ersetzen
                        INI_Endpoint_2 = INI_Endpoint_2.Replace("{" & "parameter" & (i + 1) & "}", paramValue)
                        INI_APICall_2 = INI_APICall_2.Replace("{" & "parameter" & (i + 1) & "}", paramValue)
                        INI_APICall_Object_2 = INI_APICall_Object_2.Replace("{" & "parameter" & (i + 1) & "}", paramValue)
                    Next

                Else
                    Return
                End If
            End If

            SelectedText = selection.Text.Trim()
            Dim llmresult As String = Await LLM(OtherPrompt, SelectedText, "", "", 0, True)

            SP_MergePrompt_Cached = SP_MergePrompt

            If Not String.IsNullOrWhiteSpace(llmresult) Then
                If OptionChecked Then

                    Dim ClipPaneText1 As String = "Your service has provided the following result (you can edit it):"
                    Dim ClipText2 As String = "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made (without formatting), or you can directly insert the original text in your document. If you select Cancel, nothing will be put into the clipboard."

                    If DoPane Then

                        If _uiContext IsNot Nothing Then  ' Make sure we run in the UI Thread
                            _uiContext.Post(Sub(s)

                                                ShowPaneAsync(
                                        ClipPaneText1,
                                        llmresult,
                                        "",
                                        AN,
                                        noRTF:=False,
                                        insertMarkdown:=True
                                        )
                                            End Sub, Nothing)
                        Else

                            ShowPaneAsync(ClipPaneText1, llmresult, "", AN, noRTF:=False, insertMarkdown:=True)

                        End If

                    Else

                        Dim dialogResult As String = ""

                        If _uiContext IsNot Nothing Then
                            Dim doneEvent As New ManualResetEventSlim(False)            ' Make sure we run in the UI Thread

                            _uiContext.Post(Sub(state)
                                                Try

                                                    Dim wordHwnd As IntPtr = GetWordMainWindowHandle()

                                                    dialogResult = ShowCustomWindow(ClipPaneText1,
                                                                            llmresult,
                                                                            ClipText2,
                                                                            AN,
                                                                            NoRTF:=False,
                                                                            Getfocus:=False,
                                                                            InsertMarkdown:=True,
                                                                            TransferToPane:=True,
                                                                            parentWindowHwnd:=wordHwnd)

                                                    If dialogResult <> "" And dialogResult <> "Pane" Then
                                                        If dialogResult = "Markdown" Then
                                                            Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                            InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & llmresult, False)
                                                        Else
                                                            SLib.PutInClipboard(dialogResult)
                                                        End If
                                                    ElseIf dialogResult = "Pane" Then

                                                        Debug.WriteLine($"SP_Mergeprompt = {SP_MergePrompt}")
                                                        Debug.WriteLine($"SP_Mergeprompt_Cached = {SP_MergePrompt_Cached}")


                                                        ShowPaneAsync(
                                                                            ClipPaneText1,
                                                                            llmresult,
                                                                            "",
                                                                            AN,
                                                                            noRTF:=False,
                                                                            insertMarkdown:=True
                                                                            )
                                                    End If

                                                Finally
                                                    doneEvent.Set()
                                                End Try
                                            End Sub, Nothing)
                            ' doneEvent.Wait()

                        Else
                            dialogResult = ShowCustomWindow(
                                            ClipPaneText1,
                                            llmresult,
                                            ClipText2,
                                            AN,
                                            NoRTF:=False,
                                            Getfocus:=False,
                                            InsertMarkdown:=True,
                                            TransferToPane:=True)

                            If dialogResult <> "" And dialogResult <> "Pane" Then
                                If dialogResult = "Markdown" Then
                                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                    InsertTextWithMarkdown(selection, vbCrLf & llmresult, False)
                                Else
                                    SLib.PutInClipboard(dialogResult)
                                End If
                            ElseIf dialogResult = "Pane" Then

                                ShowPaneAsync(
                                                    ClipPaneText1,
                                                    llmresult,
                                                    "",
                                                    AN,
                                                    noRTF:=False,
                                                    insertMarkdown:=True
                                                    )
                            End If

                        End If

                    End If
                Else
                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                    InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, llmresult, False)
                End If
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in SpecialModel: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If originalConfig IsNot Nothing Then
                RestoreDefaults(_context, originalConfig)
            End If
            originalConfigLoaded = False
        End Try
    End Sub




    Public Async Sub FreeStyle(UseSecondAPI)
        If INILoadFail() Then Return
        Try
            OtherPrompt = ""
            SysPrompt = ""

            Dim NoText As Boolean = False
            Dim DoMarkup As Boolean = False
            Dim DoClipboard As Boolean = False
            Dim DoBubbles As Boolean = False
            Dim DoInplace As Boolean = INI_ReplaceText2
            Dim MarkupMethod As Integer = INI_MarkupMethodWord
            Dim DoLib As Boolean = False
            Dim DoNet As Boolean = False
            Dim DoTPMarkup As Boolean = False
            Dim TPMarkupName As String = ""
            Dim KeepFormatCap = INI_KeepFormatCap
            Dim DoKeepFormat As Boolean = INI_KeepFormat2
            Dim DoKeepParaFormat As Boolean = INI_KeepParaFormatInline
            Dim DoFileObject As Boolean = False
            Dim DoFileObjectClip As Boolean = False
            Dim DoPane As Boolean = False

            Dim MarkupInstruct As String = $"start With '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"with '{InPlacePrefix}' for replacing the selection"
            Dim BubblesInstruct As String = $"with '{BubblesPrefix}' for having your text commented"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}' or '{PanePrefix}' for separate output"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim ExtInstruct As String = $"; inlcude '{ExtTrigger}' for text of a file (txt, docx, pdf)"
            Dim TPMarkupInstruct As String = $"; add '{TPMarkupTriggerInstruct}' if revisions [of user] should be pointed out to the LLM"
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}' for overriding formatting defaults"
            Dim AllInstruct As String = $"; add '{AllTrigger}' to select all"
            Dim LibInstruct As String = $"; add '{LibTrigger}' for library search"
            Dim NetInstruct As String = $"; add '{NetTrigger}' for internet search"
            Dim PureInstruct As String = $"; use '{PurePrefix}' for direct prompting"
            Dim ObjectInstruct As String = $"; add '{ObjectTrigger}'/'{ObjectTrigger2}' for adding a file object"
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
            Dim FileObject As String = ""

            Dim application As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Selection = application.Selection

            If selection.Type = WdSelectionType.wdSelectionIP Then NoText = True

            Dim AddOnInstruct As String = AllInstruct

            If Not NoText Then
                AddOnInstruct += NoFormatInstruct.Replace("; add", ", ")
                AddOnInstruct += TPMarkupInstruct.Replace("; add", ", ")
            End If
            If INI_Lib Then
                AddOnInstruct += LibInstruct.Replace("; add", ",")
            End If
            If INI_ISearch Then
                AddOnInstruct += NetInstruct.Replace("; add", ", ")
            End If
            If UseSecondAPI Then
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object_2) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
            Else
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
            End If

            Dim lastCommaIndex As Integer = AddOnInstruct.LastIndexOf(","c)
            If lastCommaIndex <> -1 Then
                AddOnInstruct = AddOnInstruct.Substring(0, lastCommaIndex) & ", and" & AddOnInstruct.Substring(lastCommaIndex + 1)
            End If


            If Not NoText Then
                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {ClipboardInstruct}, {InplaceInstruct} or {BubblesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{LastPromptInstruct}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt).Trim()
            Else
                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct} or {BubblesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{LastPromptInstruct}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt).Trim()
            End If

            SelectedText = ""

            If Not NoText Then

                SelectedText = selection.Text

                If OtherPrompt.StartsWith("codebasis", StringComparison.OrdinalIgnoreCase) Then
                    SLib.WriteToRegistry(RemoveCR(RegPath_CodeBasis), RemoveCR(selection.Text))
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Return
                End If
                If OtherPrompt.StartsWith("inipath", StringComparison.OrdinalIgnoreCase) Then
                    SLib.WriteToRegistry(RemoveCR(RegPath_IniPath), RemoveCR(selection.Text))
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Return
                End If
                If OtherPrompt.StartsWith("encode", StringComparison.OrdinalIgnoreCase) Then
                    Dim Key As String = CodeAPIKey(RemoveCR(selection.Text))
                    SLib.PutInClipboard(Key)
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    selection.TypeText(vbCrLf & "Encoded key (also in clipboard):" & vbCrLf & Key)
                    selection.ParagraphFormat.Hyphenation = False ' Turn off hyphenation
                    SLib.PutInClipboard(Key)
                    Return
                End If

                If OtherPrompt.StartsWith("decode", StringComparison.OrdinalIgnoreCase) Then
                    Dim Key As String = DeCodeAPIKey(RemoveCR(selection.Text))
                    SLib.PutInClipboard(Key)
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    selection.TypeText(vbCrLf & "Decoded key (also in clipboard):" & vbCrLf & Key)
                    selection.ParagraphFormat.Hyphenation = False ' Turn off hyphenation
                    Return
                End If

            End If
            If OtherPrompt.StartsWith("domain", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"{AN} is running in the domain '{GetDomain()}' and configured to run in {If(String.IsNullOrEmpty(SLib.alloweddomains), "any domain ('alloweddomains' has not been set).", "'" & SLib.alloweddomains & "'.")}", "")
                Return
            End If
            If OtherPrompt.StartsWith("model", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox("I am using the " & INI_Model & " model as my primary model with a default timeout of " & (INI_Timeout / 1000) & " seconds (" & Format(INI_Timeout / 60000, "0.00") & " minutes)." & If(INI_MaxOutputToken > 0, "The maximum output token length is " & INI_MaxOutputToken & ".", ""))
                Return
            End If
            If OtherPrompt.StartsWith("terms", StringComparison.OrdinalIgnoreCase) Then
                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                selection.TypeText(vbCrLf & If(INI_UsageRestrictions = "", "No usage restrictions or permissions have been defined in the configuration file.", "The defined usage restrictions or permissions defined in the configuration file are: " & INI_UsageRestrictions) & vbCrLf)
                Return
            End If
            If OtherPrompt.StartsWith("anonymize", StringComparison.OrdinalIgnoreCase) Then
                Call AnonymizeSelection()
                Return
            End If
            If OtherPrompt.StartsWith("insertclipboard", StringComparison.OrdinalIgnoreCase) OrElse OtherPrompt.StartsWith("insertclip", StringComparison.OrdinalIgnoreCase) OrElse OtherPrompt.StartsWith("clipboard", StringComparison.OrdinalIgnoreCase) OrElse OtherPrompt.StartsWith("iclip", StringComparison.OrdinalIgnoreCase) Then
                Call InsertClipboard()
                Return
            End If

            If OtherPrompt.StartsWith("generateresponsekey", StringComparison.OrdinalIgnoreCase) Or OtherPrompt.StartsWith("generateresponsetemplate", StringComparison.OrdinalIgnoreCase) Then

                If NoText Then
                    ShowCustomMessageBox("No text has been selected. Select the text containing both the JSON payload to interpret and what you want the output to look like (by referencing to the JSON fields and structure in natural text).")
                    Return
                End If

                Dim response As String = Await LLM(SP_GenerateResponseKey & vbCrLf & Code_JsonTemplateFormatter, vbCrLf & SelectedText, "", "", 0, UseSecondAPI)

                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                selection.InsertAfter(vbCrLf & vbCrLf & response)

                Return
            End If


            If OtherPrompt.StartsWith("switch", StringComparison.OrdinalIgnoreCase) Then
                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                If INI_SecondAPI Then
                    SwitchModels(_context)
                    ShowCustomMessageBox("You have temporarily switched the two configured models. Primary is now '" & INI_Model & "', and secondary is '" & INI_Model_2 & "'.")
                Else
                    ShowCustomMessageBox("You have defined only one model ('" & INI_Model & "').")
                End If
                Return
            End If
            If OtherPrompt.StartsWith("version", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox("You are using " & Version & $" of {AN}. (c) by David Rosenthal, VISCHER. Go to https://vischer.com/{AN2} for more information. This copy of {AN} is set to expire on {LicensedTill.ToString("dd-MMM-yyyy")}", AN)
                Return
            End If
            If OtherPrompt.StartsWith("reset", StringComparison.OrdinalIgnoreCase) Then
                If ShowCustomYesNoBox($"Do you really want to reset your local configuration file and settings (if any) by removing non-mandatory entries? The current configuration file '{AN2}.ini' will NOT be saved to a '.bak' file. If you only want to reload the configuration settings for giving up any temporary changes, use 'reload' instead.", "Yes", "No") = 1 Then
                    INIloaded = False
                    ResetLocalAppConfig(_context)
                    MenusAdded = False
                    AddContextMenu()
                    ShowCustomMessageBox($"Following the reset, the configuration file '{AN2}.ini' has been be reloaded.")
                End If
                Return
            End If

            If OtherPrompt.StartsWith("speech", StringComparison.OrdinalIgnoreCase) Then
                Transcriptor()
                Return

            End If

            If OtherPrompt.StartsWith("readlocal", StringComparison.OrdinalIgnoreCase) Then
                SpeakSelectedText()
                Return

            End If

            If OtherPrompt.StartsWith("voiceslocal", StringComparison.OrdinalIgnoreCase) Then
                SelectVoiceByNumber()
                Return
            End If

            If OtherPrompt.StartsWith("voices2", StringComparison.OrdinalIgnoreCase) Then
                Using frm As New TTSSelectionForm("Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", True) ' TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", True)
                    If frm.ShowDialog() = DialogResult.OK Then
                        ' Retrieve selected voices
                        Dim selectedVoices As List(Of String) = frm.SelectedVoices
                        Dim outputPath As String = frm.SelectedOutputPath
                        ' Use the selected values
                        If selectedVoices.Count > 0 Then
                            MessageBox.Show("Selected Voice(s): " & String.Join(", ", selectedVoices))
                        Else
                            MessageBox.Show("No voices selected.")
                        End If

                        If outputPath = "" Then
                            MessageBox.Show("Temporary output selected.")
                        Else
                            MessageBox.Show("Output path: " & outputPath)
                        End If
                    Else
                        MessageBox.Show("Voice selection was cancelled.")
                    End If
                End Using

                Return
            End If

            If OtherPrompt.StartsWith("voices", StringComparison.OrdinalIgnoreCase) Then
                Using frm As New TTSSelectionForm("Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", False) ' TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", False)
                    If frm.ShowDialog() = DialogResult.OK Then
                        ' Retrieve selected voices
                        Dim selectedVoices As List(Of String) = frm.SelectedVoices
                        Dim outputPath As String = frm.SelectedOutputPath
                        ' Use the selected values
                        If selectedVoices.Count > 0 Then
                            MessageBox.Show("Selected Voice(s): " & String.Join(", ", selectedVoices))
                        Else
                            MessageBox.Show("No voices selected.")
                        End If

                        If outputPath = "" Then
                            MessageBox.Show("Temporary output selected.")
                        Else
                            MessageBox.Show("Output path: " & outputPath)
                        End If
                    Else
                        MessageBox.Show("Voice selection was cancelled.")
                    End If
                End Using

                Return
            End If

            If OtherPrompt.StartsWith("createpodcast", StringComparison.OrdinalIgnoreCase) Then
                CreatePodcast()
                Return
            End If

            If OtherPrompt.StartsWith("readpodcast", StringComparison.OrdinalIgnoreCase) Then
                ReadPodcast(selection.Text)
                Return
            End If

            If OtherPrompt.StartsWith("read", StringComparison.OrdinalIgnoreCase) Then
                CreateAudio()
                Return
            End If

            If OtherPrompt.StartsWith("cleanmenu", StringComparison.OrdinalIgnoreCase) Then
                RemoveOldContextMenu()
                RemoveVeryOldContextMenu()
                MenusAdded = False
                AddContextMenu()
                Return
            End If

            If OtherPrompt.StartsWith("reload", StringComparison.OrdinalIgnoreCase) Then
                INIloaded = False
                InitializeConfig(False, True)
                MenusAdded = False
                AddContextMenu()
                ShowCustomMessageBox($"The configuration file '{AN2}.ini' has been be reloaded.")
                Return
            End If
            If OtherPrompt.StartsWith("settings", StringComparison.OrdinalIgnoreCase) Then
                ShowSettings()
                Return
            End If

            If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then

                Dim promptlibresult As (String, Boolean, Boolean, Boolean)

                promptlibresult = ShowPromptSelector(INI_PromptLibPath, Not NoText, Not NoText)

                OtherPrompt = promptlibresult.Item1
                DoMarkup = promptlibresult.Item2
                DoBubbles = promptlibresult.Item3
                DoClipboard = promptlibresult.Item4

                If OtherPrompt = "" Then
                    Return
                End If
            Else
                If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return
            End If

            My.Settings.LastPrompt = OtherPrompt
            My.Settings.Save()

            If OtherPrompt.IndexOf(AllTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(AllTrigger, "").Trim()
                Dim document As Word.Document = application.ActiveDocument
                document.Content.Select()
                NoText = False
            End If

            If OtherPrompt.IndexOf(LibTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(LibTrigger, "").Trim()
                DoLib = True
            End If

            If OtherPrompt.IndexOf(TPMarkupTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(TPMarkupTrigger, "").Trim()
                DoTPMarkup = True
            End If

            ' Formatting Trigger

            If OtherPrompt.IndexOf(NoFormatTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger, "").Trim()
                KeepFormatCap = 1
            End If
            If OtherPrompt.IndexOf(NoFormatTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger2, "").Trim()
                KeepFormatCap = 1
            End If
            If OtherPrompt.IndexOf(KFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger, "").Trim()
                DoKeepFormat = True
            End If
            If OtherPrompt.IndexOf(KFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger2, "").Trim()
                DoKeepFormat = True
            End If
            If OtherPrompt.IndexOf(KPFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger, "").Trim()
                DoKeepParaFormat = True
            End If
            If OtherPrompt.IndexOf(KPFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger2, "").Trim()
                DoKeepParaFormat = True
            End If
            If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger, "(a file object follows)").Trim()
            ElseIf DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a file object follows)").Trim()
                DoFileObjectClip = True
            Else
                DoFileObject = False
            End If


            ' Regular expression to find text in the format "(markup:..." and extract until ")"
            Dim pattern As String = $"\{TPMarkupTriggerL}}}(.*?)\{TPMarkupTriggerR}"
            ' Match the pattern in the input string
            Dim match As Match = Regex.Match(OtherPrompt, pattern, RegexOptions.IgnoreCase)
            If match.Success Then
                ' Extract the captured group (the text between "(markup:" and ")")
                TPMarkupName = match.Groups(1).Value
                DoTPMarkup = True
                OtherPrompt = Regex.Replace(OtherPrompt, pattern, String.Empty, RegexOptions.IgnoreCase)
            End If

            If OtherPrompt.StartsWith(ClipboardPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix.Length).Trim()
                DoClipboard = True
            ElseIf OtherPrompt.StartsWith(ClipboardPrefix2, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix2.Length).Trim()
                DoClipboard = True
            ElseIf OtherPrompt.StartsWith(BubblesPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(BubblesPrefix.Length).Trim()
                DoBubbles = True
            ElseIf OtherPrompt.StartsWith(InPlacePrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(InPlacePrefix.Length).Trim()
                DoInplace = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                DoMarkup = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefixRegex, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixRegex.Length).Trim()
                DoMarkup = True
                MarkupMethod = 4
            ElseIf OtherPrompt.StartsWith(MarkupPrefixWord, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixWord.Length).Trim()
                DoMarkup = True
                MarkupMethod = 1
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiffW, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiffW.Length).Trim()
                DoMarkup = True
                MarkupMethod = 3
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiff, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiff.Length).Trim()
                DoMarkup = True
                MarkupMethod = 2
            ElseIf OtherPrompt.StartsWith(PanePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PanePrefix.Length).Trim()
                DoPane = True
                DoClipboard = True
            End If


            If OtherPrompt.IndexOf(NetTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NetTrigger, "").Trim()
                DoNet = True
            End If


            If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                DragDropFormLabel = ""
                DragDropFormFilter = ""
                doc = GetFileContent()
                If String.IsNullOrWhiteSpace(doc) Then
                    ShowCustomMessageBox("The file you have selected is empty or not supported - exiting.")
                    Return
                End If
                OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ExtTrigger), doc, RegexOptions.IgnoreCase)
                ShowCustomMessageBox($"This file will be included in your prompt where you have referred to {ExtTrigger}: " & vbCrLf & vbCrLf & doc)
            End If

            If DoFileObject Then
                If DoFileObjectClip Then
                    FileObject = "clipboard"
                Else
                    DragDropFormLabel = "All file types that are supported by your LLM."
                    DragDropFormFilter = "Supported Files|*.*"
                    FileObject = GetFileName()
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    If String.IsNullOrWhiteSpace(FileObject) Then
                        ShowCustomMessageBox("No file object has been selected - will abort. You can try again (use Ctrl-P to re-insert your prompt).")
                        Return
                    End If
                End If
            End If


            If NoText And DoBubbles Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Ask the LLM to comment on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Return
                End If
            End If

            If NoText And DoMarkup Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Do the markup on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Return
                End If
            End If

            If Not DoInplace And DoMarkup Then
                Dim AppendMarkup As Integer = ShowCustomYesNoBox("You have asked for a markup to be created, but according to the configuration, it will not replace your current selection but added to it at the end. Is this really what you want?", "Yes, add markup ", "No, replace text with markup")
                If AppendMarkup = 0 Then
                    Return
                ElseIf AppendMarkup = 2 Then
                    DoInplace = True
                End If
            End If

            If OtherPrompt.StartsWith(PurePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PurePrefix.Length).Trim()
                SysPrompt = OtherPrompt
            Else
                If DoLib Then
                    Dim isSuccess As Boolean = Await ConsultLibrary(DoMarkup) ' updates SysPrompt
                    If Not isSuccess Then Return
                ElseIf DoNet Then
                    Dim isSuccess As Boolean = Await ConsultInternet(DoMarkup) ' updates SysPrompt
                    If Not isSuccess Then Return
                ElseIf NoText Then
                    SysPrompt = SP_FreestyleNoText
                Else
                    SysPrompt = SP_FreestyleText
                    If DoBubbles Then SysPrompt = SysPrompt & " " & SP_Add_Bubbles
                End If
            End If

            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SysPrompt), True, DoKeepFormat, DoKeepParaFormat, DoInplace, DoMarkup, MarkupMethod, DoClipboard, DoBubbles, False, UseSecondAPI, KeepFormatCap, DoTPMarkup, TPMarkupName, False, FileObject, DoPane)

            If UseSecondAPI And originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Async Function ConsultLibrary(DoMarkup As Boolean) As Task(Of Boolean)

        Try

            Dim SysPromptTemp As String

            ' Load the library text

            Dim LibFilePath As String = ExpandEnvironmentVariables(INI_Lib_File)

            InfoBox.ShowInfoBox("Loading the library from " & LibFilePath & " ...")

            LibraryText = ReadTextFile(LibFilePath)

            If String.IsNullOrWhiteSpace(LibraryText) Then
                InfoBox.ShowInfoBox("")
                ShowCustomMessageBox("The library file '" & LibFilePath & "' is empty or could not be read.")
                Return False
            End If

            InfoBox.ShowInfoBox("Asking the LLM to search the library based on the intruction ....")

            SysPromptTemp = InterpolateAtRuntime(INI_Lib_Find_SP)

            LibResult = Await LLM(SysPromptTemp, SelectedText, "", "", INI_Lib_Timeout)

            If String.IsNullOrWhiteSpace(LibResult) Then
                InfoBox.ShowInfoBox("")
                ShowCustomMessageBox("The LLM failed to retrieve relevant content from the library. Will abort.")
                Return False
            End If

            InfoBox.ShowInfoBox("Having the LLM apply the result from the library search: " & LibResult, 5)

            If DoMarkup And Not String.IsNullOrWhiteSpace(SelectedText) Then
                SysPrompt = InterpolateAtRuntime(INI_Lib_Apply_SP_Markup)
            Else
                SysPrompt = InterpolateAtRuntime(INI_Lib_Apply_SP)
            End If

            Return True

        Catch ex As System.Exception
            MessageBox.Show("Error in ConsultLibrary: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function
    Public Async Function ConsultInternet(DoMarkup As Boolean) As Task(Of Boolean)

        Try

            InfoBox.ShowInfoBox("Asking the LLM to determine the necessary searchterms for your instruction ...")

            Dim SysPromptTemp As String
            Dim SearchResults As List(Of String)

            CurrentDate = DateAndTime.Now.ToString("MMMM d, yyyy")

            SysPromptTemp = InterpolateAtRuntime(INI_ISearch_SearchTerm_SP)

            SearchTerms = Await LLM(SysPromptTemp, If(SelectedText = "", "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0)

            If String.IsNullOrWhiteSpace(SearchTerms) Then
                InfoBox.ShowInfoBox("")
                ShowCustomMessageBox("The LLM failed to establish searchterms. Will abort.")
                Return False
            End If

            If INI_ISearch_Approve Then
                InfoBox.ShowInfoBox("")
                Dim approveresult As Integer = ShowCustomYesNoBox("These are the searchterms that the LLM wants to issue to " & INI_ISearch_Name & ": {SearchTerms}", "Approve", "Abort", $"{AN} Internet Search", 5, " = 'Approve'")
                If approveresult = 0 Or 2 Then Return False
            End If

            InfoBox.ShowInfoBox($"Now using {INI_ISearch_Name} to search for '{SearchTerms}' ...")

            SearchResults = Await PerformSearchGrounding(SearchTerms, INI_ISearch_URL, INI_ISearch_ResponseMask1, INI_ISearch_ResponseMask2, INI_ISearch_Tries, INI_ISearch_MaxDepth)

            SearchResult = String.Join(Environment.NewLine, SearchResults.Select(Function(result, index) $"<SEARCHRESULT{index + 1}>{result}</SEARCHRESULT{index + 1}>"))

            InfoBox.ShowInfoBox($"Having the LLM execute your instruction using also the {SearchResults.Count} result(s) from the Internet search ...", 3)
            If DoMarkup And Not String.IsNullOrWhiteSpace(SelectedText) Then
                SysPrompt = InterpolateAtRuntime(INI_ISearch_Apply_SP_Markup)
            Else
                SysPrompt = InterpolateAtRuntime(INI_ISearch_Apply_SP)
            End If

            Return True

        Catch ex As System.Exception
            MessageBox.Show("Error in ConsultInternet: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function



    Public Async Sub EasterEgg()

        Dim splash As New SplashScreen($"{AN6} is preparing to tickle{If(INI_RoastMe, " (inofficial version)", "")}...")
        splash.Show()
        splash.Refresh()

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection
        Dim inputtext As String = Nothing
        Dim result As String
        Dim UserLanguage As String = Globals.ThisAddIn.GetWordDefaultInterfaceLanguage()
        Dim basicinstruction As String = $"Write me in {UserLanguage} a piece of wisdom based on hasidic humour related To the following input. The output shall be politically correct, contain nothing religious and not be offensive, but within those bounds it can be provocative. Do tickle my brain!"
        If INI_RoastMe Then basicinstruction = $"Write me in {UserLanguage} a piece of wisdom based on hasidic humour related To the following input. You do not have to be politically correct. You can use black humour, if this fits the purpose. You don't have to take any sensitivities into consideration. But don't say anything that could be considered discrimination based on race, ethnicity, religion or sexual orientation. Be direct! Be provocative! Roast me! Tickle my brain!"

        If selection IsNot Nothing Then
            ' Case 1: A selection exists
            If Not String.IsNullOrWhiteSpace(selection.Text) Then
                inputtext = selection.Text
            End If
        End If

        If String.IsNullOrWhiteSpace(inputtext) Then
            ' Case 2: No selection or empty selection, use paragraph text at the cursor
            Dim currentParagraph As Word.Paragraph = Nothing

            If selection IsNot Nothing AndAlso selection.Range IsNot Nothing Then
                currentParagraph = selection.Range.Paragraphs.First
            End If

            If currentParagraph Is Nothing OrElse String.IsNullOrWhiteSpace(currentParagraph.Range.Text) Then
                ' Case 3: No cursor paragraph, fallback to the first paragraph with text
                For Each paragraph As Word.Paragraph In application.ActiveDocument.Paragraphs
                    If Not String.IsNullOrWhiteSpace(paragraph.Range.Text) Then
                        currentParagraph = paragraph
                        Exit For
                    End If
                Next
            End If

            If currentParagraph IsNot Nothing Then
                inputtext = currentParagraph.Range.Text.Trim()
            End If
        End If

        If String.IsNullOrWhiteSpace(inputtext) Then
            ' Case 4: No text in document, use fallback logic
            Dim userName As String = Globals.ThisAddIn.Application.UserName
            Dim currentMonth As String = DateTime.Now.ToString("MMMM") ' Full month name
            Dim currentDay As String = DateTime.Now.ToString("dddd")  ' Full day name
            result = Await LLM($"{basicinstruction} The only input you get is the name of the current user ({userName}) of this Word application (you may create friendly variations), the current month is {currentMonth} and the current day of week is {currentDay}.", "", "", "", 0, False, True)
        Else
            result = Await LLM($"{basicinstruction} This Is the text: {inputtext}", "", "", "", 0, False, True)
        End If
        splash.Close()

        ShowCustomMessageBox(result, $"{AN6} tickles your brain ...")
    End Sub


    ' ProcessSelectedText Parameters:
    ' - SysCommand: A string command to be processed.
    ' - CheckMaxToken: Boolean flag to check the maximum token limit.
    ' - KeepFormat: Boolean flag to maintain the formatting of the text.
    ' - ParaFormatInline: Boolean flag to format paragraphs inline.
    ' - InPlace: Boolean flag to indicate that the output should replace the selected text.
    ' - DoMarkup: Boolean flag to indicate that the output should be provided as a markup of the selected text.
    ' - MarkupMethod: Integer to indicate the markup method to be used: 1 = Word, 2 = Diff, 3 = Regex
    ' - PutInClipboard: Boolean flag to output the processed text in the clipboard.
    ' - PutInBubbles: Boolean flag to output the processed text in bubbles
    ' - SelectionMandatory: Boolean flag to enforce text selection before processing.
    ' - UseSecondAPI: Boolean flag to decide if a secondary API should be utilized.
    ' - FormattingCap: Number indicating the maximum number of characters for preserving format
    ' - DoTPMarkup: Boolean flag to indicate that markups in the output should marked.
    ' - TPMarkupname: String containing the user of whom the tags will be marked, if any.
    ' - CreatePodcast: Boolean flag to indicate that the output should be used to create a podcast.
    ' - FileObject: String containing the file path to the object to be added to the LLM request if supported by the API.
    ' - DoPane: Boolean flag to indicate that the output should be shown in a pane.

    ' Global array to store paragraph formatting information
    Structure ParagraphFormatStructure
        Dim Style As Word.Style
        Dim FontName As String
        Dim FontSize As Single
        Dim FontBold As Integer
        Dim FontItalic As Integer
        Dim FontUnderline As Word.WdUnderline
        Dim FontColor As Word.WdColor
        Dim ListType As Word.WdListType
        Dim ListTemplate As Word.ListTemplate
        Dim ListLevel As Integer
        Dim ListNumber As Integer
        Dim HasListFormat As Boolean
        Dim Alignment As Word.WdParagraphAlignment
        Dim LineSpacing As Single
        Dim SpaceBefore As Single
        Dim SpaceAfter As Single
    End Structure

    Dim paragraphFormat() As ParagraphFormatStructure
    Dim paraCount As Integer

    Private Async Function ProcessSelectedText(SysCommand As String, CheckMaxToken As Boolean, KeepFormat As Boolean, ParaFormatInline As Boolean, InPlace As Boolean, DoMarkup As Boolean, MarkupMethod As Integer, PutInClipboard As Boolean, PutInBubbles As Boolean, SelectionMandatory As Boolean, UseSecondAPI As Boolean, FormattingCap As Integer, Optional DoTPMarkup As Boolean = False, Optional TPMarkupname As String = "", Optional CreatePodcast As Boolean = False, Optional FileObject As String = "", Optional DoPane As Boolean = False) As Task(Of String)

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection

        If SysCommand = "" Then
            ShowCustomMessageBox("The (system-)prompt For the LLM Is missing.")
            Return ""
        End If

        If selection.Type = WdSelectionType.wdSelectionIP And SelectionMandatory Then
            ShowCustomMessageBox("Please Select the text To be processed.")
            Return ""
        End If

        If selection.Type = WdSelectionType.wdSelectionIP Or selection.Tables.Count = 0 Or PutInClipboard Or PutInBubbles Then

            Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast, FileObject, DoPane)

        Else

            Dim userdialog As Integer = ShowCustomYesNoBox("Your text contains tables. Shall each text section and each cell content be processed separately to maintain the table? This may take more time" & If(DoMarkup And MarkupMethod <> 2, " (will also switch to Diff [shown in text] markup). If you do not want to change the text, select 'No'.", "."), "No", "Yes, process each cell individually", $"{AN} Table Processing")

            If userdialog = 2 Then

                MarkupMethod = 2

                Dim selRange As Range = selection.Range
                Dim docTables As Tables = selRange.Tables

                Dim isEntirelyWithinTable As Boolean = False
                Dim isWholeTable As Boolean = False
                Dim isPartialTableSelection As Boolean = False

                If selection.Tables.Count = 1 Then
                    Dim tbl As Table = selRange.Tables(1)
                    Dim tblRange As Range = tbl.Range

                    ' Check if the selection is entirely within the table boundaries.
                    isEntirelyWithinTable = (selRange.Start >= tblRange.Start AndAlso selRange.End <= tblRange.End)

                    ' Get trimmed texts. Adjust the characters to trim as needed.
                    Dim selText As String = selRange.Text.Trim(vbCr, vbLf, " "c)
                    Dim tblText As String = tblRange.Text.Trim(vbCr, vbLf, " "c)

                    ' Compare the texts. If they differ, then the selection is not the whole table.
                    isWholeTable = (selText = tblText)

                    ' If the selection is fully contained in the table but does not equal the entire table's text,
                    ' then it is entirely within the table but is only a part of it.
                    If isEntirelyWithinTable AndAlso Not isWholeTable Then
                        isPartialTableSelection = True
                    End If
                End If

                If isEntirelyWithinTable Or isWholeTable Then

                    Dim tbl As Table = selRange.Tables(1)

                    For Each row As Word.Row In tbl.Rows
                        ' Cycle through each cell in the current row.
                        For Each cell As Word.Cell In row.Cells
                            ' Work with a duplicate of the cell's range.
                            Dim cellRange As Word.Range = cell.Range.Duplicate
                            ' Exclude the cell marker.
                            cellRange.End -= 1

                            ' Check if this cell's range intersects with the selection.
                            If cellRange.End >= selRange.Start AndAlso cellRange.Start <= selRange.End Then
                                ' Calculate the intersection between the cell and selection ranges.
                                Dim intersectionRange As Word.Range = selRange.Duplicate
                                intersectionRange.Start = Math.Max(cellRange.Start, selRange.Start)
                                intersectionRange.End = Math.Min(cellRange.End, selRange.End)

                                ' Only process if there is a valid range.
                                If intersectionRange.Start < intersectionRange.End Then
                                    ' (Optional) Use DoEvents if you need to keep the UI responsive.
                                    System.Windows.Forms.Application.DoEvents()

                                    ' Select the intersection (for visual feedback or further processing).
                                    intersectionRange.Select()

                                    ' Process the selected text.
                                    ' (Replace the parameters with the ones required by your method.)
                                    Dim result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken,
                                                                                KeepFormat, ParaFormatInline, InPlace,
                                                                                DoMarkup, MarkupMethod, PutInClipboard,
                                                                                PutInBubbles, SelectionMandatory, UseSecondAPI,
                                                                                FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane)
                                    ' Optionally delay between processing cells.
                                    Await System.Threading.Tasks.Task.Delay(500)
                                End If
                            End If
                        Next
                    Next

                Else

                    ' Sort tables by their start positions in the selection
                    Dim tableList As New List(Of Table)
                    For i As Integer = 1 To docTables.Count
                        tableList.Add(docTables(i))
                    Next
                    tableList.Sort(Function(t1, t2) t1.Range.Start.CompareTo(t2.Range.Start))

                    Dim lastPos As Integer = selRange.Start

                    Dim splash As New SplashScreen("Processing table(s)... press 'Esc' to abort")
                    splash.Show()
                    splash.Refresh()

                    Dim IsExit As Boolean = False

                    For Each tbl As Table In tableList

                        System.Windows.Forms.Application.DoEvents()

                        If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                            Exit For
                        End If

                        If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Or IsExit Then
                            ' Exit the loop
                            Exit For
                        End If

                        Dim tblStart As Integer = tbl.Range.Start
                        Dim tblEnd As Integer = tbl.Range.End

                        ' Text chunk BEFORE the table
                        If tblStart > lastPos Then
                            Dim textChunk As Range = selRange.Duplicate
                            textChunk.Start = lastPos
                            textChunk.End = tblStart - 1

                            ' Double-check you haven't snagged any table content
                            If textChunk.Tables.Count = 0 Then
                                ' Also verify it's not empty
                                If textChunk.Start < textChunk.End Then
                                    textChunk.Select()
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane)
                                    Await System.Threading.Tasks.Task.Delay(500)
                                End If
                            Else

                                Do
                                    textChunk.Start += 1
                                Loop While textChunk.Tables.Count <> 0 And Not textChunk.Start = textChunk.End

                                If textChunk.Tables.Count = 0 AndAlso textChunk.Start < textChunk.End Then
                                    textChunk.Select()
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane)
                                    Await System.Threading.Tasks.Task.Delay(500)
                                End If

                            End If
                        End If

                        ' Process the table itself (cells)
                        For Each row As Row In tbl.Rows
                            System.Windows.Forms.Application.DoEvents()

                            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                                Exit For
                            End If

                            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Or IsExit Then
                                ' Exit the loop
                                Exit For
                            End If
                            For Each cell As Cell In row.Cells
                                System.Windows.Forms.Application.DoEvents()

                                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                                    Exit For
                                End If

                                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Or IsExit Then
                                    ' Exit the loop
                                    Exit For
                                End If
                                Dim cellRange As Range = cell.Range
                                cellRange.End -= 1  ' Exclude cell marker
                                If cellRange.Start < cellRange.End Then
                                    cellRange.Select()
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane)
                                    Await System.Threading.Tasks.Task.Delay(500)
                                End If
                            Next
                        Next

                        ' Move lastPos to end of this table
                        lastPos = tblEnd + 1
                    Next

                    ' Text chunk AFTER the last table
                    If lastPos <= selRange.End And Not IsExit Then
                        Dim finalChunk As Range = selRange.Duplicate
                        finalChunk.Start = lastPos
                        finalChunk.End = selRange.End

                        If finalChunk.Tables.Count = 0 AndAlso finalChunk.Start < finalChunk.End Then

                            finalChunk.Select()
                            Dim text = selection.Text
                            Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane)
                        Else
                            Do
                                finalChunk.Start += 1
                            Loop While finalChunk.Tables.Count <> 0 And Not finalChunk.Start = finalChunk.End

                            finalChunk.End = selRange.End

                            If finalChunk.Tables.Count = 0 AndAlso finalChunk.Start < finalChunk.End Then
                                finalChunk.Select()
                                Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane)
                            End If
                        End If
                    End If

                    splash.Close()
                End If

            ElseIf userdialog = 1 Then

                Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast, FileObject, DoPane)

            End If

        End If

        If Not PutInClipboard Then
            selection.Collapse(WdCollapseDirection.wdCollapseEnd)
            selection.MoveStart(WdUnits.wdCharacter, 0)
            selection.MoveEnd(WdUnits.wdCharacter, 0)
        End If

        Return ""

    End Function
    Private Async Function TrueProcessSelectedText(SysCommand As String, CheckMaxToken As Boolean, KeepFormat As Boolean, ParaFormatInline As Boolean, InPlace As Boolean, DoMarkup As Boolean, MarkupMethod As Integer, PutInClipboard As Boolean, PutInBubbles As Boolean, SelectionMandatory As Boolean, UseSecondAPI As Boolean, FormattingCap As Integer, Optional DoTPMarkup As Boolean = False, Optional TPMarkupname As String = "", Optional CreatePodcast As Boolean = False, Optional FileObject As String = "", Optional DoPane As Boolean = False) As Task(Of String)

        Try
            Dim SelectedText As String = ""
            Dim rng As Range
            Dim i As Integer
            Dim NoFormatting As Boolean = False
            Dim NoSelectedText As Boolean = False
            Dim trailingCR As Boolean

            Dim application As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Selection = application.Selection

            If selection.Type = WdSelectionType.wdSelectionIP And SelectionMandatory Then
                Return ""
            End If

            If selection.Type = WdSelectionType.wdSelectionIP Then NoSelectedText = True

            rng = selection.Range

            If Not NoSelectedText Then

                If rng.Text.Length = 0 Then NoSelectedText = True
                If Not NoSelectedText And FormattingCap > 0 And rng.Text.Length > FormattingCap Then NoFormatting = True

            End If

            If PutInBubbles Or PutInClipboard Or NoSelectedText Then NoFormatting = True

            If PutInBubbles Then
                DoMarkup = False
                PutInClipboard = False
            End If

            If PutInClipboard Then DoMarkup = False

            If DoTPMarkup Then NoFormatting = True

            If MarkupMethod = 4 Then NoFormatting = True

            paraCount = 0

            If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoSelectedText Then

                paraCount = rng.Paragraphs.Count

                ReDim paragraphFormat(paraCount - 1)
                Array.Clear(paragraphFormat, 0, paragraphFormat.Length)

                For i = 1 To paraCount
                    Dim para As Word.Paragraph = rng.Paragraphs(i)
                    Dim paraRange As Word.Range = para.Range

                    Try
                        paragraphFormat(i - 1) = New ParagraphFormatStructure With {
                            .Style = If(para.Style IsNot Nothing, para.Style, Nothing),
                            .FontName = If(Not String.IsNullOrEmpty(paraRange.Font.Name), paraRange.Font.Name, ""),
                            .FontSize = If(paraRange.Font.Size > 0, paraRange.Font.Size, 0),
                            .FontBold = paraRange.Font.Bold,
                            .FontItalic = paraRange.Font.Italic,
                            .FontUnderline = paraRange.Font.Underline,
                            .FontColor = paraRange.Font.Color,
                            .ListType = paraRange.ListFormat.ListType,
                            .ListTemplate = If(paraRange.ListFormat.ListType <> Word.WdListType.wdListNoNumbering, paraRange.ListFormat.ListTemplate, Nothing),
                            .ListLevel = If(paraRange.ListFormat.ListType <> Word.WdListType.wdListNoNumbering, paraRange.ListFormat.ListLevelNumber, 0),
                            .ListNumber = If(paraRange.ListFormat.ListType <> Word.WdListType.wdListNoNumbering, paraRange.ListFormat.ListValue, 0),
                            .HasListFormat = paraRange.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                            .Alignment = para.Alignment,
                            .LineSpacing = para.LineSpacing,
                            .SpaceBefore = para.SpaceBefore,
                            .SpaceAfter = para.SpaceAfter
                    }
                    Catch ex As System.Exception
                        'Debug.Print($"Error extracting paragraph formatting for paragraph {i}: {ex.Message}")
                    End Try

                Next

            End If

            If Not NoSelectedText AndAlso INI_MarkdownConvert AndAlso Not KeepFormat AndAlso (Not DoMarkup OrElse MarkupMethod = 3) AndAlso rng.Text.Length < INI_MarkupDiffCap Then
                ConvertRangeFormattingToMarkdown()
                rng = selection.Range
            End If

            If Not NoSelectedText Then

                If KeepFormat And Not NoFormatting Then
                    SelectedText = SLib.GetRangeHtml(rng)
                Else
                    'SelectedText = rng.Text
                    If NoFormatting Then
                        If DoTPMarkup Then
                            SelectedText = AddMarkupTags(rng, TPMarkupname)
                        Else
                            SelectedText = rng.Text
                        End If
                    Else
                        SelectedText = GetTextWithSpecialElementsInline(rng, If(NoFormatting, False, ParaFormatInline))
                    End If
                    trailingCR = (SelectedText.EndsWith(vbCrLf) Or SelectedText.EndsWith(vbLf) Or SelectedText.EndsWith(vbCr))

                End If

                Dim MaxToken As Integer = If(UseSecondAPI, INI_MaxOutputToken_2, INI_MaxOutputToken)
                Dim EstimatedTokens As Integer = EstimateTokenCount(SelectedText)

                If CheckMaxToken And MaxToken > 0 And EstimatedTokens > MaxToken And (InPlace Or DoMarkup) Then
                    ShowCustomMessageBox("Your selected text is larger than the maximum output your LLM can supposedly generate. Therefore, the output may be shorter than expected based on maximum tokens supported, which is " & MaxToken & " tokens. Your input (with formatting information, as the case may be) has an estimated to be " & EstimatedTokens & " tokens). Therefore, check whether the output is complete.", AN, 15)
                End If

                If DoMarkup And MarkupMethod = 2 And Len(SelectedText) > INI_MarkupDiffCap Then
                    Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Diff markup method at {INI_MarkupDiffCap} chars (your selection has {Len(SelectedText)} chars). {If(KeepFormat, "This may be because HTML codes have been inserted to keep the formatting (you can turn this off in the settings). ", "")}How do you want to continue?", "Use Diff in Window compare instead", "Use Diff")
                    Select Case MarkupChange
                        Case 1
                            MarkupMethod = 3
                        Case 2
                            MarkupMethod = 2
                        Case Else
                            Return ""
                    End Select
                End If

                If DoMarkup And MarkupMethod = 4 And Len(SelectedText) > INI_MarkupRegexCap Then
                    Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Regex markup method at {INI_MarkupRegexCap} chars (your selection has {Len(SelectedText)} chars). {If(KeepFormat, "This may be because HTML codes have been inserted to keep the formatting (you can turn this off in the settings). ", "")}How do you want to continue?", "Use Word compare instead", "Use Regex")
                    Select Case MarkupChange
                        Case 1
                            MarkupMethod = 1
                        Case 2
                            MarkupMethod = 4
                        Case Else
                            Return ""
                    End Select
                End If

            Else

                SelectedText = ""

            End If

            Dim LLMResult = Await LLM(SysCommand & " " & If(NoFormatting, "", If(KeepFormat, " " & SP_Add_KeepHTMLIntact, " " & SP_Add_KeepInlineIntact)), If(NoSelectedText, "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

            OtherPrompt = ""

            LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

            If Not String.IsNullOrEmpty(LLMResult) Then
                LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
            End If

            If ParaFormatInline Then LLMResult = CorrectPFORMarkers(LLMResult)

            If DoTPMarkup Then LLMResult = RemoveMarkupTags(LLMResult)

            If (MarkupMethod <> 4 Or Not DoMarkup) And InPlace And Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
            If (MarkupMethod <> 4 Or Not DoMarkup) And InPlace And Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)

            If (MarkupMethod <> 4 Or Not DoMarkup) And trailingCR And (LLMResult.EndsWith(ControlChars.Cr) Or LLMResult.EndsWith(ControlChars.Lf)) Then LLMResult = LLMResult.Replace(ControlChars.Cr, ControlChars.CrLf).Replace(ControlChars.Lf, ControlChars.CrLf)

            If Not String.IsNullOrEmpty(LLMResult) Then

                Dim ClipPaneText1 As String = "The LLM has provided the following result (you can edit it):"
                Dim ClipText2 As String = "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made (without formatting), or you can directly insert the original text in your document. If you select Cancel, nothing will be put into the clipboard."
                Dim PaneText2 As String = "Choose to put your edited or original text in the clipboard, or inserted the original with formatting; the pane will close. You can also copy & paste from the pane."

                If CreatePodcast Then
                    Dim TTSAvailable As Boolean = False

                    DetectTTSEngines()

                    If Not TTS_googleAvailable AndAlso Not TTS_openAIAvailable Then
                        TTSAvailable = False
                    Else
                        TTSAvailable = True
                    End If


                    If TTSAvailable Then
                        Dim FinalText = ShowCustomWindow("The LLM has created the following podcast script for you (you can edit it; you do not have to manually remove the SSML codes, if you do not like them):", LLMResult, "The next step is the production of an audio file. You can choose whether you want to use the original text or your text with any changes you have made. The text will also be put in the clipboard. If you select Cancel, the original text will only be put into the clipboard.", AN, True)

                        If FinalText = "" Then
                            SLib.PutInClipboard(LLMResult)
                        Else
                            FinalText = FinalText.Trim()
                            SLib.PutInClipboard(FinalText)
                            If FinalText.Contains("H: ") AndAlso FinalText.Contains("G: ") Then ReadPodcast(FinalText)
                        End If
                    Else
                        Dim FinalText = ShowCustomWindow("The LLM has created the following podcast script for you (you can edit it; you do not have to manually remove the SSML codes, if you do not like them):", LLMResult, $"The next step is the production of an audio file. Since you have not configured {AN} for Google, you unfortunately cannot do that here. However, you can choose whether you want the original text or the text with your changes to put in the clipboard for further use. If you select Cancel, no text will be put in the clipboard.", AN, True)

                        If FinalText <> "" Then
                            SLib.PutInClipboard(LLMResult)
                        Else
                            FinalText = FinalText.Trim()
                            SLib.PutInClipboard(FinalText)
                        End If
                    End If

                ElseIf DoPane Then

                    If _uiContext IsNot Nothing Then  ' Make sure we run in the UI Thread
                        _uiContext.Post(Sub(s)
                                            SP_MergePrompt_Cached = SP_MergePrompt
                                            ShowPaneAsync(
                                        ClipPaneText1,
                                        LLMResult,
                                        PaneText2,
                                        AN,
                                        noRTF:=False,
                                        insertMarkdown:=True
                                        )
                                        End Sub, Nothing)
                    Else

                        SP_MergePrompt_Cached = SP_MergePrompt
                        ShowPaneAsync(ClipPaneText1, LLMResult, PaneText2, AN, noRTF:=False, insertMarkdown:=True)
                    End If

                ElseIf PutInClipboard Then

                    Dim dialogResult As String = ""

                    If _uiContext IsNot Nothing Then
                        Dim doneEvent As New ManualResetEventSlim(False)            ' Make sure we run in the UI Thread

                        _uiContext.Post(Sub(state)
                                            Try

                                                Dim wordHwnd As IntPtr = GetWordMainWindowHandle()

                                                dialogResult = ShowCustomWindow(ClipPaneText1,
                                                                            LLMResult,
                                                                            ClipText2,
                                                                            AN,
                                                                            NoRTF:=False,
                                                                            Getfocus:=False,
                                                                            InsertMarkdown:=True,
                                                                            TransferToPane:=True,
                                                                            parentWindowHwnd:=wordHwnd)

                                                If dialogResult <> "" And dialogResult <> "Pane" Then
                                                    If dialogResult = "Markdown" Then
                                                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                        InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, LLMResult, False)
                                                    Else
                                                        SLib.PutInClipboard(dialogResult)
                                                    End If
                                                ElseIf dialogResult = "Pane" Then
                                                    SP_MergePrompt_Cached = SP_MergePrompt
                                                    ShowPaneAsync(
                                                                            ClipPaneText1,
                                                                            LLMResult,
                                                                            PaneText2,
                                                                            AN,
                                                                            noRTF:=False,
                                                                            insertMarkdown:=True
                                                                            )
                                                End If

                                            Finally
                                                doneEvent.Set()
                                            End Try
                                        End Sub, Nothing)
                        ' doneEvent.Wait()

                    Else
                        dialogResult = ShowCustomWindow(
                                            ClipPaneText1,
                                            LLMResult,
                                            ClipText2,
                                            AN,
                                            NoRTF:=False,
                                            Getfocus:=False,
                                            InsertMarkdown:=True,
                                            TransferToPane:=True)

                        If dialogResult <> "" And dialogResult <> "Pane" Then
                            If dialogResult = "Markdown" Then
                                Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, LLMResult, False)
                            Else
                                SLib.PutInClipboard(dialogResult)
                            End If
                        ElseIf dialogResult = "Pane" Then
                            SP_MergePrompt_Cached = SP_MergePrompt
                            ShowPaneAsync(
                                                    ClipPaneText1,
                                                    LLMResult,
                                                    PaneText2,
                                                    AN,
                                                    noRTF:=False,
                                                    insertMarkdown:=True
                                                    )
                        End If

                    End If

                ElseIf PutInBubbles Then

                    Dim responseItems() As String = LLMResult.Split({"§§§"}, StringSplitOptions.RemoveEmptyEntries)
                    Dim wrongformatresponse As New List(Of String)
                    Dim notfoundresponse As New List(Of String)
                    Dim originalRange As Word.Range = selection.Range.Duplicate ' Save the original selection range
                    Dim BubblecutHappened As Boolean = False

                    Dim splash As New SplashScreen("Adding bubbles To your text... press 'Esc' to abort")
                    splash.Show()
                    splash.Refresh()

                    For Each item In responseItems

                        System.Windows.Forms.Application.DoEvents()

                        If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                        If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                            Exit For
                        End If

                        Dim parts() As String = item.Split({"@@"}, StringSplitOptions.None)
                        If parts.Length = 2 Then

                            Dim findText As String = parts(0).Trim().Trim("'"c).Trim(""""c)
                            Dim commentText As String = parts(1).Trim()

                            Try
                                If findText.Length <= 255 Then
                                    ' Use the built-in Find directly if <= 255 characters
                                    If selection.Find.Execute(FindText:=findText) Then
                                        Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, $"{AN5}: " & commentText)
                                    Else
                                        notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}: " & commentText & vbCrLf & vbCrLf)
                                    End If
                                Else
                                    ' Use chunk-by-chunk search for > 255 characters
                                    If FindLongTextInChunks(findText, 255, selection) Then
                                        ' If found, selection now covers the entire matched text
                                        Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, $"{AN5}: " & commentText)
                                    Else
                                        notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}: " & commentText & vbCrLf & vbCrLf)
                                    End If
                                End If

                            Catch ex As Exception
                                notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}: " & commentText & " [Error: " & ex.Message & "]" & vbCrLf & vbCrLf)
                            End Try

                        Else
                            wrongformatresponse.Add(item)
                        End If

                        selection.SetRange(originalRange.Start, originalRange.End) ' Restore the original selection
                    Next

                    splash.Close()

                    Dim ErrorList As String = ""
                    If notfoundresponse.Count > 0 Then
                        ErrorList += "The following comments could not be assigned to your text (they were not found):" & vbCrLf & vbCrLf
                        For Each item In notfoundresponse
                            If item.Trim() <> "" Then ErrorList += Trim("- " & item & vbCrLf)
                        Next
                        ErrorList += vbCrLf
                    End If

                    If wrongformatresponse.Count > 0 Then
                        ErrorList += "The following responses could not be identified as bubble comments:" & vbCrLf & vbCrLf
                        For Each item In wrongformatresponse
                            If item.Trim() <> "" Then ErrorList += Trim("- " & item & vbCrLf)
                        Next
                        ErrorList += vbCrLf
                    End If
                    If Not String.IsNullOrWhiteSpace(ErrorList) Then
                        If BubblecutHappened Then
                            ErrorList = $"Some of the sections to which the bubble comments relate were too long for selecting. Only the initial part has been selected. This is indicated by '{BubbleCutText}' in the bubble comments, as applicable." & vbCrLf & vbCrLf & ErrorList
                        End If

                        ErrorList = ShowCustomWindow("Errors when implementing the 'bubbles' feedback of the LLM:", ErrorList, "The above error list will be included in a final comment at the end of your selection (it will also be included in the clipboard). You can have the original list included, or you can now make changes and have this version used. If you select Cancel, nothing will be put added to the document.", AN, True)

                        If ErrorList <> "" And ErrorList.ToLower() <> "esc" Then
                            SLib.PutInClipboard(ErrorList)
                            Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, $"{AN5}: " & ErrorList)
                        End If

                    Else

                        ShowCustomMessageBox("The bubble comments provided by the LLM have been added to to your text." & If(BubblecutHappened, $"Some of the sections to which the bubble comments relate were too long for selecting. Only the initial part has been selected. This is indicated by '{BubbleCutText}' in the bubble comments, as applicable.", ""))
                    End If

                ElseIf MarkupMethod = 4 Then

                    Dim RegexResult = Await LLM(SP_MarkupRegex, "<ORIGINALTEXT>" & SelectedText & "</ORIGINALTEXT> /n <NEWTEXT>" & LLMResult & "</NEWTEXT>", "", "", 0, UseSecondAPI)

                    MarkupSelectedTextWithRegex(RegexResult)

                    ' End Extended Selection Mode
                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ElseIf NoSelectedText Then
                    selection.TypeText(vbCrLf & vbCrLf)
                    InsertTextWithMarkdown(selection, LLMResult, trailingCR)

                    ' End Extended Selection Mode
                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ElseIf KeepFormat AndAlso Not NoFormatting Then
                    SelectedText = selection.Text
                    SLib.InsertTextWithFormat(LLMResult, rng, InPlace)
                    If DoMarkup Then
                        LLMResult = SLib.RemoveHTML(LLMResult)
                        If MarkupMethod = 2 Or MarkupMethod = 3 Then
                            Dim SaveRng As Range = rng.Duplicate
                            CompareAndInsert(SelectedText, LLMResult, rng, MarkupMethod = 3, "This is the markup of the text inserted:")
                            If Not ParaFormatInline And Not NoFormatting Then
                                ApplyParagraphFormat(rng)
                            End If
                            RestoreSpecialTextElements(SaveRng)
                        Else
                            CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                            RestoreSpecialTextElements(rng)
                        End If
                    End If

                    ' End Extended Selection Mode
                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                Else
                    SelectedText = selection.Text

                    If InPlace Then
                        If DoMarkup Then
                            If MarkupMethod = 2 Or MarkupMethod = 3 Then
                                If MarkupMethod = 3 Then
                                    InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                                    'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                    rng = selection.Range
                                Else
                                    If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                End If
                                Dim SaveRng As Range = rng.Duplicate
                                CompareAndInsert(SelectedText, LLMResult, rng, MarkupMethod = 3, "This is the markup of the text inserted:")
                                If Not ParaFormatInline AndAlso Not NoFormatting Then
                                    ApplyParagraphFormat(rng)
                                End If
                                RestoreSpecialTextElements(SaveRng)
                            Else
                                If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                                RestoreSpecialTextElements(rng)
                            End If
                        Else
                            InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                            rng = selection.Range
                            Dim SaveRng As Range = rng.Duplicate
                            If Not ParaFormatInline And Not NoFormatting Then
                                ApplyParagraphFormat(rng)
                            End If
                            RestoreSpecialTextElements(SaveRng)
                        End If

                    Else
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        selection.TypeText(vbCrLf & vbCrLf)
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        rng = selection.Range
                        If DoMarkup Then
                            If MarkupMethod = 2 Or MarkupMethod = 3 Then
                                If MarkupMethod = 3 Then
                                    Dim pattern As String = "\{\{.*?\}\}"
                                    If System.Text.RegularExpressions.Regex.IsMatch(LLMResult, pattern) Then
                                        SLib.InsertTextWithBoldMarkers(selection, LLMResult)
                                        'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                    Else
                                        InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                                        'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                    End If
                                    rng = selection.Range
                                End If
                                Dim SaveRng As Range = rng.Duplicate
                                CompareAndInsert(SelectedText, LLMResult, rng.Duplicate, MarkupMethod = 3, "This is the markup of the text inserted:")
                                If Not ParaFormatInline And Not NoFormatting Then
                                    ApplyParagraphFormat(rng)
                                End If
                                RestoreSpecialTextElements(SaveRng)
                            Else
                                If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                                RestoreSpecialTextElements(rng)
                            End If
                        Else
                            Dim pattern As String = "\{\{.*?\}\}"
                            If System.Text.RegularExpressions.Regex.IsMatch(LLMResult, pattern) Then
                                SLib.InsertTextWithBoldMarkers(selection, LLMResult)
                                'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                            Else
                                InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                                'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                            End If
                            rng = selection.Range
                            Dim SaveRng As Range = rng.Duplicate
                            If Not ParaFormatInline And Not NoFormatting Then
                                ApplyParagraphFormat(rng)
                            End If
                            RestoreSpecialTextElements(SaveRng)
                        End If

                    End If

                    ' End Extended Selection Mode
                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                End If

            Else
                ShowCustomMessageBox("The LLM did not return any content to process.")
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in ProcessSelectedText: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return ""
    End Function


    Public Sub ConvertRangeFormattingToMarkdown()
        Dim app As Word.Application = Globals.ThisAddIn.Application
        Dim sel As Word.Selection = app.Selection

        ' --- 1. PRELIMINARY CHECKS ---
        If sel Is Nothing OrElse sel.Type = Word.WdSelectionType.wdSelectionIP Then
            Return
        End If

        Dim rng As Word.Range = sel.Range
        Dim originalStart As Integer = rng.Start
        If rng.Characters.Count = 0 Then Return

        Dim sb As New StringBuilder()

        ' --- 2. STATE TRACKING ---
        Dim isBold As Boolean = False
        Dim isItalic As Boolean = False
        Dim isUnderlined As Boolean = False

        ' --- 3. ITERATE THROUGH CHARACTERS ---
        For i As Integer = 1 To rng.Characters.Count
            Dim charRng As Word.Range = rng.Characters(i)

            Dim currentBold As Boolean = (charRng.Font.Bold = -1)
            Dim currentItalic As Boolean = (charRng.Font.Italic = -1)
            Dim currentUnderline As Boolean = (charRng.Font.Underline <> Word.WdUnderline.wdUnderlineNone)
            Dim charText As String = charRng.Text
            Dim isEndOfLine As Boolean = (charText = vbCr)

            ' --- A. HANDLE CLOSING TAGS ---
            If isUnderlined AndAlso (Not currentUnderline OrElse isEndOfLine) Then sb.Append("</u>")
            If isItalic AndAlso (Not currentItalic OrElse isEndOfLine) Then sb.Append("*")
            If isBold AndAlso (Not currentBold OrElse isEndOfLine) Then sb.Append("**")

            ' --- B. HANDLE OPENING TAGS ---
            If Not isBold AndAlso currentBold AndAlso Not isEndOfLine Then sb.Append("**")
            If Not isItalic AndAlso currentItalic AndAlso Not isEndOfLine Then sb.Append("*")
            If Not isUnderlined AndAlso currentUnderline AndAlso Not isEndOfLine Then sb.Append("<u>")

            ' Update state for the next character
            If isEndOfLine Then
                isBold = False
                isItalic = False
                isUnderlined = False
            Else
                isBold = currentBold
                isItalic = currentItalic
                isUnderlined = currentUnderline
            End If

            ' --- C. APPEND THE CHARACTER TEXT ---
            sb.Append(charText)
        Next

        ' --- 4. CLOSE ANY REMAINING DANGLING TAGS ---
        If isUnderlined Then sb.Append("</u>")
        If isItalic Then sb.Append("*")
        If isBold Then sb.Append("**")

        ' For debugging: Check final string in Visual Studio's "Output" window.
        Debug.WriteLine("Generated Markdown: " & sb.ToString())

        ' --- 5. REPLACE TEXT, FIX FORMATTING, AND UPDATE SELECTION ---

        ' First, replace the original range's text. At this moment, Word might
        ' incorrectly apply formatting from the first original character.
        rng.Text = sb.ToString()

        ' *** THE CORRECTED FIX ***
        ' The 'rng' object now refers to the new content we just inserted.
        ' We now explicitly remove the character formatting from this new range
        ' to ensure it's plain text, leaving only our markdown tags.
        ' Word Interop uses 0 for False.
        rng.Font.Bold = 0
        rng.Font.Italic = 0
        rng.Font.Underline = Word.WdUnderline.wdUnderlineNone

        ' Reselect the range to include the newly added tags.
        Dim newEnd As Integer = originalStart + sb.Length
        app.ActiveDocument.Range(originalStart, newEnd).Select()

    End Sub



    Public Function RemoveMarkdownFormatting(ByVal input As String) As String
        Try
            Dim output As String = input

            ' 1) Fett+Kursiv: ***Text*** → Text
            output = Regex.Replace(output, "\*\*\*(.+?)\*\*\*", "$1", RegexOptions.Singleline)

            ' 2) Fett: **Text** → Text
            output = Regex.Replace(output, "\*\*(.+?)\*\*", "$1", RegexOptions.Singleline)

            ' 3) Kursiv: *Text* → Text
            output = Regex.Replace(output, "(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", "$1", RegexOptions.Singleline)

            ' 4) Durchgestrichen: ~~Text~~ → Text
            output = Regex.Replace(output, "~~(.+?)~~", "$1", RegexOptions.Singleline)

            ' 5) Superscript: <sup>Text</sup> → Text
            output = Regex.Replace(output, "<sup>(.+?)</sup>", "$1", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

            ' 6) Subscript: <sub>Text</sub> → Text
            output = Regex.Replace(output, "<sub>(.+?)</sub>", "$1", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

            Return output

        Catch ex As System.Exception
            ' Hier könntest Du Logging oder eine Meldung einfügen
            Throw New System.Exception("Error in RemoveMarkdownFormatting: " & ex.Message, ex)
        End Try
    End Function

    Public Sub ApplyParagraphFormat(ByRef rng As Range)

        Dim maxParaStylesCount As Integer = paragraphFormat.Length
        paraCount = rng.Paragraphs.Count

        If paraCount > 0 Then
            For i = 1 To paraCount
                If i - 1 < maxParaStylesCount Then
                    With rng.Paragraphs(i).Range
                        Dim format = paragraphFormat(i - 1)

                        ' Apply the stored style
                        If format.Style IsNot Nothing Then
                            Try
                                .Style = format.Style
                            Catch ex As System.Exception
                                'Debug.Print($"Error applying style: {ex.Message}")
                            End Try
                        End If

                        ' Apply the stored font formatting
                        With .Font
                            Try
                                If Not String.IsNullOrEmpty(format.FontName) Then .Name = format.FontName
                                If format.FontSize > 0 Then .Size = format.FontSize
                                If format.FontBold = 0 Or format.FontBold = -1 Then .Bold = format.FontBold
                                If format.FontItalic = 0 Or format.FontItalic = -1 Then .Italic = format.FontItalic
                                .Underline = format.FontUnderline
                                .Color = format.FontColor
                            Catch ex As System.Exception
                                'Debug.Writeline($"Error applying font properties: {ex.Message}")
                            End Try
                        End With

                        ' Apply list formatting if applicable
                        If format.HasListFormat AndAlso format.ListTemplate IsNot Nothing Then
                            Try
                                If .ListFormat.ListType <> Word.WdListType.wdListNoNumbering Then
                                    .ListFormat.RemoveNumbers()
                                End If

                                .ListFormat.ApplyListTemplateWithLevel(
                                        ListTemplate:=format.ListTemplate,
                                        ContinuePreviousList:=If(format.ListNumber > 0, True, False),
                                        ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList,
                                        DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior
                                    )
                                .ListFormat.ListLevelNumber = format.ListLevel
                            Catch ex As System.Exception
                                'Debug.Writeline($"Error applying list format: {ex.Message}")
                            End Try
                        End If

                        ' Apply paragraph alignment
                        Try
                            .ParagraphFormat.Alignment = format.Alignment
                        Catch ex As System.Exception
                            'Debug.Writeline($"Error applying alignment: {ex.Message}")
                        End Try

                        ' Apply line spacing
                        Try
                            .ParagraphFormat.LineSpacing = format.LineSpacing
                        Catch ex As System.Exception
                            'Debug.Writeline($"Error applying line spacing: {ex.Message}")
                        End Try

                        ' Apply spacing before and after
                        Try
                            .ParagraphFormat.SpaceBefore = format.SpaceBefore
                            .ParagraphFormat.SpaceAfter = format.SpaceAfter
                        Catch ex As System.Exception
                            'Debug.Writeline($"Error applying spacing: {ex.Message}")
                        End Try

                    End With
                End If
            Next
        End If

    End Sub

    Public Function CorrectPFORMarkers(ByVal input As String) As String
        Try
            Dim output As New StringBuilder()
            Dim i As Integer = 0
            Dim length As Integer = input.Length

            While i < length
                ' Detect PFOR markers
                If i <= length - 9 AndAlso input.Substring(i, 7) = "{{PFOR:" Then
                    ' Check if it's "PFOR:0"
                    Dim endIndex As Integer = input.IndexOf("}}", i)
                    If endIndex <> -1 Then
                        Dim markerContent As String = input.Substring(i + 7, endIndex - (i + 7)) ' Extract "nnn"
                        If markerContent = "0" Then
                            ' If it's PFOR:0, copy as-is and move the pointer
                            output.Append(input.Substring(i, endIndex - i + 2))
                            i = endIndex + 2
                            Continue While
                        End If
                    End If

                    ' Check preceding character
                    If output.Length > 0 Then
                        Dim prevChar As Char = output(output.Length - 1)
                        If prevChar <> vbCr AndAlso prevChar <> vbLf Then
                            output.Append(vbCrLf) ' Add newline before the marker
                        End If
                    End If

                    ' Append the marker
                    Dim markerEnd As Integer = input.IndexOf("}}", i) + 2
                    output.Append(input.Substring(i, markerEnd - i))
                    i = markerEnd
                Else
                    ' Copy character-by-character
                    output.Append(input(i))
                    i += 1
                End If
            End While

            Return output.ToString()
        Catch ex As System.Exception
            Debug.WriteLine("An error occurred while correcting PFOR markers: " & ex.Message, ex)
        End Try
    End Function

    Public Function FindLongTextInChunks(ByVal findText As String, ByVal chunkSize As Integer, ByRef selection As Word.Selection) As Boolean
        ' Store original selection to restore if needed
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim originalStart As Integer = selection.Start
        Dim originalEnd As Integer = selection.End

        ' Break the long text into chunks of up to chunkSize characters
        Dim chunks As New List(Of String)
        Dim startIndex As Integer = 0
        While startIndex < findText.Length
            Dim length As Integer = Math.Min(chunkSize, findText.Length - startIndex)
            chunks.Add(findText.Substring(startIndex, length))
            startIndex += length
        End While

        ' We'll need to track the final Start/End of the matched text
        Dim overallMatchStart As Integer = -1
        Dim overallMatchEnd As Integer = -1

        ' Move the selection to the beginning of the document (or keep at original if you prefer)

        For i As Integer = 0 To chunks.Count - 1
            Dim currentChunk As String = chunks(i)

            Dim chunk As String = chunks(i)

            With selection.Find
                .Forward = True
                .Wrap = Word.WdFindWrap.wdFindStop
                .Format = False
                If INI_Clean Then
                    .MatchWildcards = True
                    ' replace every literal space with [ ]@ (one-or-more spaces)
                    .Text = chunk.Replace(" ", "[ ]@")
                Else
                    .MatchWildcards = False
                    .Text = chunk
                End If
            End With

            ' If this is the first chunk, search from the current selection
            ' If it's not found, restore and return False
            If Not selection.Find.Execute() Then
                ' Not found
                selection.SetRange(originalStart, originalEnd) ' restore original selection
                Return False
            End If

            ' If we found the chunk:
            If i = 0 Then
                ' This is the first chunk found; record the overallMatchStart
                overallMatchStart = selection.Start
            Else
                ' For subsequent chunks, ensure continuity:
                ' The new chunk's 'Start' should be exactly the previous chunk's 'End'
                ' If there's a gap, it means it's not contiguous
                If selection.Start <> overallMatchEnd Then
                    ' Not contiguous; fail
                    selection.SetRange(originalStart, originalEnd)
                    Return False
                End If
            End If

            ' Update the overallMatchEnd to this chunk's end
            overallMatchEnd = selection.End

            ' Move the selection just after this chunk so that next chunk search begins from that point
            ' (to enforce sequential matching).
            selection.SetRange(overallMatchEnd, overallMatchEnd)
        Next

        ' If we reach here, all chunks were found contiguously
        ' We now have overallMatchStart and overallMatchEnd
        selection.SetRange(overallMatchStart, overallMatchEnd)
        Return True
    End Function


    Public Sub MarkupSelectedTextWithRegex(regexResult As String)
        Dim regexList As List(Of (Pattern As String, Replacement As String)) = ParseRegexString(regexResult)
        Dim errorCount As Integer = 0

        If regexList.Count = 0 Then
            ShowCustomMessageBox("The Regex markup method did not work, as the the LLM delivered no valid regex patterns. You may want to retry.")
            Return
        End If

        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Selection = app.Selection

            If selection Is Nothing OrElse selection.Range Is Nothing Then
                MessageBox.Show("Error in MarkupSelectedTextWithRegex: No text selected (anymore). Can't proceed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            'Dim splash As New SplashScreen("Applying changes... press 'Esc' to abort") 
            'splash.Show()
            'splash.Refresh()

            ShowProgressBarInSeparateThread($"{AN} Regex Markup", "Applying changes...")
            ProgressBarModule.CancelOperation = False

            ' Ensure Track Changes is enabled
            Dim originalTrackChangesSetting As Boolean = app.ActiveDocument.TrackRevisions
            Dim originalUserName As String = app.UserName
            app.ActiveDocument.TrackRevisions = True
            ' app.UserName = AN

            ' Define the character to be replaced
            Dim specialChar As String = ChrW(&HD83D)

            Dim selectedRange As Range = selection.Range
            Dim Exited As Boolean = False

            Dim regexIndex As Integer = 0

            For Each regexPair In regexList
                Try

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                        Exited = True
                        Exit For
                    End If

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                        Exited = True
                        Exit For
                    End If

                    selectedRange.Select()
                    SearchAndReplace(regexPair.Pattern, regexPair.Replacement, True, specialChar)

                    GlobalProgressMax = regexList.Count + 1

                    ' Update the current progress value and status label.
                    GlobalProgressValue = regexIndex + 1
                    GlobalProgressLabel = $"Search & Replace command {regexIndex + 1} of {regexList.Count}"

                    regexIndex += 1

                Catch ex As Exception
                    errorCount += 1
                End Try
            Next

            selectedRange.Select()

            If Not Exited Then

                GlobalProgressValue = regexIndex + 1
                GlobalProgressLabel = $"Cleaning up..."

                ' Loop through and replace occurrences of the character
                Dim replacementsMade As Boolean = False
                Do
                    With selectedRange.Find
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
            End If

            ProgressBarModule.CancelOperation = True

            ' Restore original Track Changes setting
            app.ActiveDocument.TrackRevisions = originalTrackChangesSetting
            ' app.UserName = originalUserName

            'splash.Close()

            If errorCount > 0 Then
                ShowCustomMessageBox($"Some markups were applied. However, in {errorCount} cases this did not work because the LLM did not return the correct results. You may want to retry.")
            End If

        Catch ex As Exception
            MessageBox.Show("Error in MarkupSelectedTextWithRegex: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Parses the input string into a list of regex patterns and replacements
    Private Function ParseRegexString(input As String) As List(Of (String, String))
        Dim result As New List(Of (String, String))
        Dim entries() As String = input.Split(New String() {RegexSeparator2}, StringSplitOptions.RemoveEmptyEntries)

        For Each entry In entries
            Dim parts() As String = entry.Split(New String() {RegexSeparator1}, StringSplitOptions.None)

            If parts.Length = 2 Then
                Dim key As String = parts(0).Trim()
                Dim value As String = parts(1).Trim()

                ' Only add if the tuple does not yet exist in result
                If Not result.Any(Function(item) item.Item1 = key AndAlso item.Item2 = value) Then
                    result.Add((key, value))
                End If
            End If
        Next

        Return result
    End Function

    Private Sub SearchAndReplace(oldText As String, newText As String, OnlySelection As Boolean, Marker As String)

        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try

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
            If newText.Length > 2 And Marker <> "" Then
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

                    ' Replace the found text
                    If Not isDeleted Then
                        currentEnd = currentEnd + Len(newTextWithMarker)
                        selectionEnd = selectionEnd + Len(newTextWithMarker)
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
                    Debug.WriteLine($"Note: The search term was not found (Chunk Search)." & Environment.NewLine)
                End If

                doc.Application.Selection.SetRange(selectionStart, selectionEnd)
                doc.Application.Selection.Select()

            Else

                If String.IsNullOrEmpty(oldText) Then
                    Debug.WriteLine($"Note: The search term was empty (bad LLM response)." & Environment.NewLine)
                Else
                    Dim replacementsMade As Boolean = False
                    ' Capture the initial end of the workRange
                    Dim initialRangeEnd As Integer = workRange.End

                    Do

                        If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                            Exit Do
                        End If

                        With workRange.Find
                            .ClearFormatting()
                            '.Text = oldText
                            If ThisAddIn.INI_Clean Then
                                .MatchWildcards = True
                                ' turn each " " into "[ ]@" so Word will match 1+ spaces
                                .Text = oldText.Replace(" ", "[ ]@")
                            Else
                                .MatchWildcards = False
                                .Text = oldText
                            End If
                            .Forward = True
                            .Wrap = Word.WdFindWrap.wdFindStop
                            .MatchWholeWord = True ' Ensures only whole words are matched
                            .MatchWildcards = False ' Ensure wildcard mode is on
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
                        Debug.WriteLine($"Note: The sarch term was not found." & Environment.NewLine)
                    End If
                End If
            End If

        Catch ex As System.Exception
            MsgBox("Error in SearchReplace: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Public Shared Sub InsertTextWithMarkdown(selection As Microsoft.Office.Interop.Word.Selection, Result As String, Optional TrailingCR As Boolean = False)

        If selection Is Nothing Then
            MessageBox.Show("Error in InsertTextWithMarkdown: The selection object is null", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim insertionStart As Integer = selection.Range.Start

        ' Extract the range from the selection
        Dim range As Microsoft.Office.Interop.Word.Range = selection.Range

        Dim LeadingTrailingSpace As Boolean = False

        If range.Start < range.End AndAlso Not TrailingCR Then

            ' Prüfen, ob vor und hinter range Platz im Dokument ist; erforderlich, weil beim Löschen eines solchen Texts Word automatisch einen Space entfernt
            Dim docStart As Integer = range.Document.Content.Start
            Dim docEnd As Integer = range.Document.Content.End

            If range.Start > docStart AndAlso range.End < docEnd Then
                ' Ein 1‐Zeichen‐Range vor range
                Dim beforerange As Range = range.Document.Range(range.Start - 1, range.Start)
                ' Ein 1‐Zeichen‐Range nach range
                Dim afterrange As Range = range.Document.Range(range.End, range.End + 1)

                If beforerange.Text = " " AndAlso afterrange.Text = " " Then
                    LeadingTrailingSpace = True
                Else
                    LeadingTrailingSpace = False
                End If
            Else
                LeadingTrailingSpace = False
            End If
        End If

        If range.Start < range.End Then
            If TrailingCR Then
                range.End = range.End - 1
            End If
        End If

        range.Delete()

        Dim markdownpipeline As MarkdownPipeline = New MarkdownPipelineBuilder() _
                        .UseAdvancedExtensions() _
                        .UseEmojiAndSmiley() _
                        .UseTaskLists() _
                        .UseMathematics() _
                        .UseGenericAttributes() _
                        .Build()

        Dim htmlResult As String = Markdown.ToHtml(Result, markdownpipeline).Trim

        ' Load the HTML into HtmlDocument
        Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
        htmlDoc.LoadHtml(htmlResult)

        Dim nodesToRemove = htmlDoc.DocumentNode.SelectNodes("//*[not(node()) and not(normalize-space())]")
        If nodesToRemove IsNot Nothing Then
            For Each emptyNode In nodesToRemove
                emptyNode.Remove()
            Next
        End If

        'Dim fullHtml As String = htmlDoc.DocumentNode.OuterHtml
        'Debug.WriteLine(fullHtml)

        ' Parse and insert HTML content into the Word range
        ParseHtmlNode(htmlDoc.DocumentNode, range)

        selection.Document.Fields.Update()

        If LeadingTrailingSpace Then
            range.Collapse(WdCollapseDirection.wdCollapseEnd)
            range.InsertAfter(" ")
        End If

        'If Trailing CR Then
        'range.Collapse(WdCollapseDirection.wdCollapseEnd)
        'range.InsertParagraphAfter()
        'range.Collapse(WdCollapseDirection.wdCollapseEnd)
        'End If

        Dim InsertionEnd As Integer = range.End

        Dim doc As Microsoft.Office.Interop.Word.Document = selection.Document
        selection.SetRange(insertionStart, insertionEnd)
        selection.Select()

    End Sub


    Private Shared Sub ParseHtmlNode(node As HtmlNode, range As Range)
        For Each childNode As HtmlNode In node.ChildNodes

            ' ——— Pre-Check auf erstes verschachteltes <a> ———
            Dim nestedLinkNode As HtmlNode = Nothing
            If Not childNode.Name.Equals("a", StringComparison.OrdinalIgnoreCase) Then
                nestedLinkNode = childNode.SelectSingleNode(".//a")
            End If
            Dim nestedHref As String = If(nestedLinkNode IsNot Nothing,
                                     nestedLinkNode.GetAttributeValue("href", String.Empty),
                                     String.Empty)
            ' ——————————————————————————————————————————————

            Dim footnotesDict As New Dictionary(Of String, String)
            Dim fnDefs = node.OwnerDocument.DocumentNode _
                  .SelectNodes("//div[@class='footnotes']//li")
            If fnDefs IsNot Nothing Then
                For Each liDef As HtmlNode In fnDefs
                    Dim id = liDef.GetAttributeValue("id", "")  ' z.B. "fn:1"
                    Dim pNode = liDef.SelectSingleNode("p")        ' das <p>…
                    ' entferne das Back-Ref-Link-Element, falls vorhanden
                    Dim backRef = pNode.SelectSingleNode("a[@class='footnote-back-ref']")
                    If backRef IsNot Nothing Then backRef.Remove()

                    ' nimm jetzt nur noch reinen Text
                    Dim text = HtmlEntity.DeEntitize(pNode.InnerText).Trim()
                    footnotesDict(id) = text
                Next
            End If


            range.Style = WdBuiltinStyle.wdStyleNormal
            Select Case childNode.Name.ToLower()

                Case "div"
                    ' Fußnoten-Container einfach überspringen
                    If childNode.GetAttributeValue("class", "") = "footnotes" Then
                        Exit Select
                    End If

                Case "#text"
                    Dim txt As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    InsertInline(range, txt, Sub(r) r.Font.Reset(), nestedHref)
                Case "strong", "b"
                    ' Fett (+ evtl. verschachtelt kursiv/underline)
                    Dim txtB As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    Dim hasItalic As Boolean = (childNode.SelectSingleNode(".//em|.//i") IsNot Nothing)
                    Dim hasUnderline As Boolean = (childNode.SelectSingleNode(".//u") IsNot Nothing)

                    InsertInline(range, txtB,
                        Sub(r)
                            r.Font.Bold = True
                            If hasItalic Then r.Font.Italic = True
                            If hasUnderline Then r.Font.Underline = WdUnderline.wdUnderlineSingle
                        End Sub,
                        nestedHref)
                Case "em", "i"
                    ' Kursiv (+ evtl. verschachtelt fett/underline)
                    Dim txtI As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    Dim hasBold As Boolean = (childNode.SelectSingleNode(".//strong|.//b") IsNot Nothing)
                    Dim hasUnderline As Boolean = (childNode.SelectSingleNode(".//u") IsNot Nothing)

                    InsertInline(range, txtI,
                        Sub(r)
                            r.Font.Italic = True
                            If hasBold Then r.Font.Bold = True
                            If hasUnderline Then r.Font.Underline = WdUnderline.wdUnderlineSingle
                        End Sub,
                        nestedHref)

                Case "u"
                    ' Underline (+ evtl. verschachtelt fett/kursiv)
                    Dim txtU As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    Dim hasBold As Boolean = (childNode.SelectSingleNode(".//strong|.//b") IsNot Nothing)
                    Dim hasItalic As Boolean = (childNode.SelectSingleNode(".//em|.//i") IsNot Nothing)

                    InsertInline(range, txtU,
                        Sub(r)
                            r.Font.Underline = WdUnderline.wdUnderlineSingle
                            If hasBold Then r.Font.Bold = True
                            If hasItalic Then r.Font.Italic = True
                        End Sub,
                        nestedHref)

                Case "br"
                    range.Font.Reset()
                    range.Text = vbCr
                    range.Collapse(WdCollapseDirection.wdCollapseEnd)

                Case "h1", "h2", "h3"
                    ' 1) Welcher Built-In Heading-Style?
                    Dim style As WdBuiltinStyle =
                                If(childNode.Name.Equals("h1", StringComparison.OrdinalIgnoreCase),
                                   WdBuiltinStyle.wdStyleHeading1,
                                If(childNode.Name.Equals("h2", StringComparison.OrdinalIgnoreCase),
                                   WdBuiltinStyle.wdStyleHeading2,
                                   WdBuiltinStyle.wdStyleHeading3))

                    Dim txt As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    Dim href As String = nestedHref

                    ' 2) Neuen Absatz einfügen und Range dorthin setzen
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    range.InsertParagraphAfter()
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ' 3) Text einfügen
                    Dim paraStart As Integer = range.Start
                    range.InsertAfter(txt)
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ' 4) Absatz-Range ermitteln
                    Dim paraRg As Word.Range = range.Document.Range(paraStart, range.End)

                    ' 5) Absatz-Stil anwenden
                    paraRg.Style = style

                    ' 6) Hyperlink (falls nötig)
                    If href <> String.Empty Then
                        Dim hl As Word.Hyperlink =
                                            range.Document.Hyperlinks.Add(
                                                Anchor:=paraRg,
                                                Address:=href
                                            )
                        hl.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        range.SetRange(hl.Range.End, hl.Range.End)
                    End If

                    ' 7) Absatz-Umbruch ans Ende
                    range.InsertParagraphAfter()
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                Case "a"
                    Dim cls = childNode.GetAttributeValue("class", "")
                    Dim href = childNode.GetAttributeValue("href", "")

                    ' → echte Fußnoten-Referenz?
                    If cls.Contains("footnote-ref") AndAlso href.StartsWith("#fn:", StringComparison.OrdinalIgnoreCase) Then
                        ' baue Dictionary 'footnotesDict' einmal vor der Schleife auf:
                        '   Dim footnotesDict As New Dictionary(Of String,String)
                        '   ...node.OwnerDocument.SelectNodes("//div[@class='footnotes']//li")...
                        '   footnotesDict("fn:1") = "Dies ist die Fussnote."

                        ' 1) halte die ID ohne '#'
                        Dim fnId = href.TrimStart("#"c)

                        If footnotesDict.ContainsKey(fnId) Then
                            ' 2) echten Word-Footnote-Marker anlegen
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Dim fn As Word.Footnote =
                              range.Document.Footnotes.Add(Range:=range)

                            ' 3) Fussnoten-Text unten einsetzen
                            fn.Range.Text = footnotesDict(fnId)

                            ' 4) Cursor hinter Marker setzen
                            fn.Reference.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            range.SetRange(fn.Reference.End, fn.Reference.End)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        End If

                    ElseIf cls.Contains("footnote-back-ref") Then
                        ' Rückverweis nicht rendern
                        Exit Select

                    Else
                        ' ganz normaler Link-Text
                        Dim txt = HtmlEntity.DeEntitize(childNode.InnerText)
                        InsertInline(range, txt,
                            Sub(r) 'keine extra Formatierung 

                            End Sub,
                                href)
                    End If



                Case "blockquote"
                    Dim txt As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    InsertInline(range, txt,
                                            Sub(r)
                                                r.ParagraphFormat.LeftIndent += 18
                                                r.Font.Italic = True
                                            End Sub,
                                            nestedHref)
                Case "ul"
                    If childNode.GetAttributeValue("class", "").Contains("contains-task-list") Then

                        For Each li As HtmlNode In childNode.SelectNodes("li")
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                            Dim chkNode = li.SelectSingleNode(".//input[@type='checkbox']")
                            Dim isChecked As Boolean = False
                            If chkNode IsNot Nothing Then
                                isChecked = chkNode.GetAttributeValue("checked", False)
                            End If

                            Dim symbol As String = If(isChecked, "☑", "☐")

                            Dim labelText = HtmlEntity.DeEntitize(li.InnerText.Trim())
                            range.InsertAfter(symbol & " " & labelText & vbCr)

                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Next

                        Exit Select
                    Else

                        ' Unordered list
                        Dim listStart As Integer = range.Start
                        For Each li As HtmlNode In childNode.SelectNodes("li")
                            ParseHtmlNode(li, range)
                            range.Text = vbCr
                            range.Collapse(False)
                        Next
                        Dim ulRange As Microsoft.Office.Interop.Word.Range = range.Document.Range(listStart, range.End)
                        ulRange.ListFormat.ApplyBulletDefault()
                        ulRange.ListFormat.ListIndent()
                        With ulRange.ParagraphFormat
                            .LeftIndent = .Application.CentimetersToPoints(0.75)
                            .FirstLineIndent = - .Application.CentimetersToPoints(0.75)
                        End With
                        range.SetRange(ulRange.End, ulRange.End)
                    End If

                Case "ol"
                    ' Ordered list
                    Dim numStart As Integer = range.Start
                    For Each li As HtmlNode In childNode.SelectNodes("li")
                        ParseHtmlNode(li, range)
                        range.Text = vbCr
                        range.Collapse(False)
                    Next
                    Dim olRange As Microsoft.Office.Interop.Word.Range = range.Document.Range(numStart, range.End)
                    olRange.ListFormat.ApplyNumberDefault()
                    olRange.ListFormat.ListIndent()
                    With olRange.ParagraphFormat
                        .LeftIndent = .Application.CentimetersToPoints(0.75)
                        .FirstLineIndent = - .Application.CentimetersToPoints(0.75)
                    End With

                    range.SetRange(olRange.End, olRange.End)

                Case "dl"
                    ' Definition list
                    For Each dt As HtmlNode In childNode.SelectNodes("dt")
                        ' Term
                        Dim term As Microsoft.Office.Interop.Word.Range = range.Duplicate
                        term.Text = HtmlEntity.DeEntitize(dt.InnerText) & vbTab
                        term.Font.Bold = True
                        term.Collapse(False)
                        range.SetRange(term.End, term.End)
                        ' Definition
                        Dim dd As HtmlNode = dt.NextSibling
                        If dd IsNot Nothing AndAlso dd.Name.ToLower() = "dd" Then
                            Dim defn As Microsoft.Office.Interop.Word.Range = range.Duplicate
                            defn.Text = HtmlEntity.DeEntitize(dd.InnerText) & vbCr
                            defn.ParagraphFormat.LeftIndent += 18
                            defn.Collapse(False)
                            range.SetRange(defn.End, defn.End)
                        End If
                    Next

                Case "sub"
                    ' Subscript
                    Dim txtSub As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    InsertInline(range, txtSub,
                        Sub(r) r.Font.Subscript = True,
                        nestedHref)

                Case "sup"
                    Dim txtSup As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    InsertInline(range, txtSup,
                         Sub(r) r.Font.Superscript = True,
                         nestedHref)

                Case "hr"

                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    Dim hrPara As Word.Paragraph = range.Document.Paragraphs.Add(range)
                    hrPara.Range.Text = ""  ' leer lassen, wir brauchen nur den Rahmen

                    With hrPara.Range.ParagraphFormat.Borders(Word.WdBorderType.wdBorderBottom)
                        .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                        .Color = Word.WdColor.wdColorAutomatic
                    End With

                    Dim afterHr As Word.Range = hrPara.Range.Duplicate
                    afterHr.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    range.SetRange(afterHr.Start, afterHr.Start)
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)



                Case "input"
                    ' Checkbox (ContentControl)
                    If childNode.GetAttributeValue("type", String.Empty).ToLower() = "checkbox" Then
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Dim cc As Word.ContentControl =
                                        range.Document.ContentControls.Add(
                                            Word.WdContentControlType.wdContentControlCheckBox,
                                            range
                                        )
                        cc.Checked = (childNode.GetAttributeValue("checked", String.Empty).ToLower() = "checked")

                        range.SetRange(cc.Range.End, cc.Range.End)
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    End If

                Case "img"
                    ' Image
                    Dim src As String = childNode.GetAttributeValue("src", String.Empty)
                    If Not String.IsNullOrEmpty(src) Then
                        Dim pic As Microsoft.Office.Interop.Word.InlineShape =
                    range.InlineShapes.AddPicture(src, LinkToFile:=False, SaveWithDocument:=True)
                        range.SetRange(pic.Range.End, pic.Range.End)
                    End If

                Case "pre"
                    ' Code block
                    Dim codeBlock As Microsoft.Office.Interop.Word.Range = range.Duplicate
                    codeBlock.Text = HtmlEntity.DeEntitize(childNode.InnerText) & vbCr
                    codeBlock.Font.Name = "Courier New"
                    codeBlock.Font.Size = 10
                    codeBlock.ParagraphFormat.LeftIndent += 14.18
                    codeBlock.Collapse(False)
                    range.SetRange(codeBlock.End, codeBlock.End)


                Case "code"
                    Dim txt As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    InsertInline(range, txt,
                                        Sub(r)
                                            r.Font.Name = "Courier New"
                                            r.Font.Size = 10
                                            r.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25
                                        End Sub,
                                        nestedHref)

                Case "span"
                    Dim cls = childNode.GetAttributeValue("class", String.Empty)
                    Dim txt As String = HtmlEntity.DeEntitize(childNode.InnerText)

                    If cls.Contains("emoji") Then
                        InsertInline(range, txt,
                            Sub(r)
                                r.Font.Name = "Segoe UI Emoji"
                                r.Font.Color = Word.WdColor.wdColorWhite
                                r.Shading.BackgroundPatternColor =
                                    System.Drawing.ColorTranslator.ToOle(
                                        System.Drawing.Color.FromArgb(0, 112, 192))
                            End Sub,
                            nestedHref)

                    ElseIf cls.Contains("math") Then
                        InsertInline(range, txt,
                            Sub(r)
                                ' Math-Inline hier nur Text einsetzen,
                                ' OMath füge danach manuell hinzu:
                            End Sub,
                            nestedHref)
                        ' Anschließend:
                        Dim mathRg As Word.Range = range.Duplicate
                        mathRg.OMaths.Add(mathRg)
                        mathRg.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    Else
                        ParseHtmlNode(childNode, range)
                    End If

                Case "table"
                    ' Table
                    Dim rowList = childNode.SelectNodes(".//tr")
                    If rowList Is Nothing OrElse rowList.Count = 0 Then Exit Select
                    Dim firstCells = rowList(0).SelectNodes("th|td")
                    If firstCells Is Nothing OrElse firstCells.Count = 0 Then Exit Select
                    Dim tbl As Microsoft.Office.Interop.Word.Table =
                    range.Document.Tables.Add(range, rowList.Count, firstCells.Count)
                    Dim r As Integer = 1
                    For Each tr As HtmlNode In rowList
                        Dim cells = tr.SelectNodes("th|td")
                        If cells IsNot Nothing AndAlso cells.Count > 0 Then
                            Dim c As Integer = 1
                            For Each cell As HtmlNode In cells
                                Dim text = HtmlEntity.DeEntitize(cell.InnerText)
                                tbl.Cell(r, c).Range.Text = text
                                If cell.Name.Equals("th", StringComparison.OrdinalIgnoreCase) Then
                                    tbl.Cell(r, c).Range.Font.Bold = True
                                End If
                                c += 1
                            Next
                        End If
                        r += 1
                    Next
                    range.SetRange(tbl.Range.End, tbl.Range.End)
                Case Else
                    ParseHtmlNode(childNode, range)
            End Select

        Next
    End Sub

    Private Shared Sub InsertInline(
    ByRef mainRg As Word.Range,
    txt As String,
    styleAction As Action(Of Word.Range),
    Optional href As String = ""
)
        ' 1) Stelle sicher, dass mainRg an der Einfügestelle eine 0-Len-Range ist
        mainRg.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

        ' 2) Baue eine Duplicate-Range, die Du gleich füllst und formatierst
        Dim wrk As Word.Range = mainRg.Duplicate
        wrk.Text = txt
        styleAction(wrk)  ' z.B. wrk.Font.Bold = True

        If href <> "" Then
            ' 3a) Hyperlink ANLEGEN, während wrk noch auf den Text zeigt
            Dim hl As Word.Hyperlink =
            mainRg.Document.Hyperlinks.Add(
                Anchor:=wrk,
                Address:=href
            )
            ' 4a) Den Hyperlink-Feld-Range kollabieren
            hl.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            ' 5a) mainRg direkt hinter das Link-Feld setzen
            mainRg.SetRange(hl.Range.End, hl.Range.End)
        Else
            ' 3b) kein Link: jetzt wrk erst kollabieren
            wrk.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            ' 4b) mainRg hinter den Plain-Text setzen
            mainRg.SetRange(wrk.End, wrk.End)
        End If
    End Sub





    Public Function AddMarkupTags(rng As Range, Optional TPMarkupName As String = Nothing) As String
        Dim resultBuilder As New StringBuilder()

        ' Check revisions in the range to apply markup for deletions and insertions
        For Each rev As Revision In rng.Revisions
            Dim includeMarkup As Boolean = True

            ' Check if we need to filter by TPMarkupName
            If Not String.IsNullOrEmpty(TPMarkupName) Then
                If Not String.Equals(rev.Author, TPMarkupName, StringComparison.OrdinalIgnoreCase) Then
                    includeMarkup = False
                End If
            End If

            ' Apply markup based on the type of revision
            If includeMarkup Then
                If rev.Type = WdRevisionType.wdRevisionDelete Then
                    resultBuilder.Append("<del>").Append(rev.Range.Text).Append("</del>")
                ElseIf rev.Type = WdRevisionType.wdRevisionInsert Then
                    resultBuilder.Append("<ins>").Append(rev.Range.Text).Append("</ins>")
                Else
                    resultBuilder.Append(rev.Range.Text)
                End If
            Else
                resultBuilder.Append(rev.Range.Text)
            End If
        Next

        ' Return the result
        Return resultBuilder.ToString()
    End Function

    Public Function RemoveMarkupTags(text As String) As String
        ' Remove <del>, </del>, <ins>, and </ins> tags using regular expressions
        Dim result As String = System.Text.RegularExpressions.Regex.Replace(text, "<del>|</del>|<ins>|</ins>", String.Empty)
        Return result
    End Function

    Private Sub CompareAndInsertComparedoc(originalText As String, newText As String, targetrange As Range, Optional paraformatinline As Boolean = False, Optional noformatting As Boolean = True)

        Dim splash As New SLib.SplashScreen("Creating markup using the Word compare functionality (ignore any flickering and press 'No' if prompted) ...")
        splash.Show()
        splash.Refresh()

        Dim wordApp As Word.Application = Globals.ThisAddIn.Application
        Dim tempOriginalDoc As Word.Document = Nothing
        Dim tempNewDoc As Word.Document = Nothing
        Dim comparisonDoc As Word.Document = Nothing
        Dim originalAuthor As String = wordApp.UserName
        Dim originalScreenUpdating As Boolean = wordApp.ScreenUpdating
        Dim rng As Word.Range

        Try
            ' Disable screen updating to reduce flickers
            wordApp.ScreenUpdating = False

            ' Set the temporary author name to app
            ' wordApp.UserName = AN

            ' Create temporary documents for original and new text
            tempOriginalDoc = wordApp.Documents.Add
            tempNewDoc = wordApp.Documents.Add

            ' Minimize the windows of the temporary documents
            tempOriginalDoc.Windows(1).WindowState = Word.WdWindowState.wdWindowStateMinimize
            tempNewDoc.Windows(1).WindowState = Word.WdWindowState.wdWindowStateMinimize

            ' Insert original text into the first temporary document
            tempOriginalDoc.Content.Text = originalText

            ' Insert new text into the second temporary document
            tempNewDoc.Content.Text = newText

            ' Define the entire newly added text to be the range rng
            rng = tempNewDoc.Content
            If Not paraformatinline And Not noformatting Then
                ApplyParagraphFormat(rng)
            End If

            ' Perform the comparison
            comparisonDoc = wordApp.CompareDocuments(
                OriginalDocument:=tempOriginalDoc,
                RevisedDocument:=tempNewDoc,
                Destination:=WdCompareDestination.wdCompareDestinationNew,
                Granularity:=WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=False,
                CompareCaseChanges:=False,
                CompareWhitespace:=False,
                CompareTables:=False,
                CompareHeaders:=False,
                CompareFootnotes:=False,
                CompareTextboxes:=False,
                CompareFields:=False,
                CompareComments:=False,
                CompareMoves:=False,
                RevisedAuthor:=Application.UserName
            )

            ' Copy the comparison document's content while keeping the original format
            comparisonDoc.Content.Copy()

            ' Insert the compared content at the specified range
            targetrange.Paste()

        Catch ex As System.Exception
            MessageBox.Show("Error in CompareAndInsertComparedoc: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Restore screen updating
            wordApp.ScreenUpdating = originalScreenUpdating

            ' Restore the original author name
            'wordApp.UserName = originalAuthor

            ' Clean up temporary documents
            If tempOriginalDoc IsNot Nothing Then tempOriginalDoc.Close(SaveChanges:=False)
            If tempNewDoc IsNot Nothing Then tempNewDoc.Close(SaveChanges:=False)
            If comparisonDoc IsNot Nothing Then comparisonDoc.Close(SaveChanges:=False)

            splash.Close()

        End Try
    End Sub
    Private Sub CompareAndInsert(text1 As String, text2 As String, targetRange As Range, Optional ShowInWindow As Boolean = False, Optional TextforWindow As String = "A text with these changes will be inserted ('Esc' to abort):", Optional paraformatinline As Boolean = False, Optional noformatting As Boolean = True)
        Try
            Dim diffBuilder As New InlineDiffBuilder(New Differ())
            Dim sText As String = String.Empty

            ' Pre-process the texts to replace line breaks with a unique marker
            text1 = text1.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")
            text2 = text2.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")

            ' Normalize the texts by removing extra spaces
            text1 = text1.Replace("  ", " ").Trim()
            text2 = text2.Replace("  ", " ").Trim()

            ' Split the texts into words and convert them into a line-by-line format
            Dim words1 As String = String.Join(Environment.NewLine, text1.Split(" "c))
            Dim words2 As String = String.Join(Environment.NewLine, text2.Split(" "c))

            ' Generate word-based diff using DiffPlex
            Dim diffResult As DiffPaneModel = diffBuilder.BuildDiffModel(words1, words2)

            ' Build the formatted output based on the diff results
            For Each line In diffResult.Lines
                Select Case line.Type
                    Case ChangeType.Inserted
                        sText &= "[INS_START]" & line.Text.Trim() & "[INS_END] "
                    Case ChangeType.Deleted
                        sText &= "[DEL_START]" & line.Text.Trim() & "[DEL_END] "
                    Case ChangeType.Unchanged
                        sText &= line.Text.Trim() & " "
                End Select
            Next

            ' Remove preceding and trailing spaces around placeholders
            sText = sText.Replace("{vbCr}", "{vbCrLf}")
            sText = sText.Replace("{vbLf}", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf} ", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf}", "{vbCrLf}")
            sText = sText.Replace("{vbCrLf} ", "{vbCrLf}")

            ' Remove instances of line breaks surrounded by [DEL_START] and [DEL_END]
            sText = sText.Replace("[DEL_START]{vbCrLf}[DEL_END] ", "")

            ' Include instances of line breaks surrounded by [INS_START] and [INS_END] without the [INS...] text
            sText = sText.Replace("[INS_START]{vbCrLf}[INS_END] ", "{vbCrLf}")

            ' Replace placeholders with actual line breaks
            sText = sText.Replace("{vbCrLf}", vbCrLf)

            ' Adjust overlapping tags
            sText = sText.Replace("[DEL_END] [INS_START]", "[DEL_END][INS_START]")
            sText = sText.Replace("[INS_START][INS_END] ", "")

            ' Insert formatted text into the specified range
            If Not ShowInWindow Then
                InsertMarkupText(sText & vbCrLf, targetRange)
            Else
                sText = Regex.Replace(sText, "\{\{.*?\}\}", String.Empty)

                Dim htmlContent As String = ConvertMarkupToRTF(TextforWindow & "\r\r" & sText)

                System.Threading.Tasks.Task.Run(Sub()
                                                    ShowRTFCustomMessageBox(htmlContent)
                                                End Sub)
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in CompareAndInsertText: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub InsertMarkupText(inputText As String, targetRange As Range)
        Try
            Dim wordApp As Word.Application = Globals.ThisAddIn.Application
            Dim TextArray() As String = {}
            Dim FormatArray() As Integer = {}

            Dim splash As New SplashScreen("Creating your markup ... press 'Esc' to abort")
            splash.Show()
            splash.Refresh()

            ' Parse the input text into chunks with formatting information
            ParseText(inputText, TextArray, FormatArray)

            ' Store the original TrackRevisions and TrackFormatting states
            Dim originalTrackRevisions As Boolean = wordApp.ActiveDocument.TrackRevisions
            Dim originalAuthor As String = wordApp.ActiveDocument.BuiltInDocumentProperties("Author").Value

            ' Enable TrackRevisions and set the author to app
            wordApp.ActiveDocument.TrackRevisions = True
            wordApp.ActiveDocument.BuiltInDocumentProperties("Author").Value = AN
            For i = 0 To TextArray.Length - 1

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    ' Exit the loop
                    Exit For
                End If

                Select Case FormatArray(i)
                    Case 1 ' [INS_START]...[INS_END]: Insert text
                        ' Insert the text as a tracked insertion
                        Dim insertRange As Range = targetRange.Duplicate
                        insertRange.Text = TextArray(i)
                        targetRange.Start = insertRange.End

                    Case 2 ' [DEL_START]...[DEL_END]: Delete text
                        ' Delete the text as a tracked deletion
                        Dim deleteRange As Range = targetRange.Duplicate
                        deleteRange.Text = TextArray(i)
                        deleteRange.Select()
                        wordApp.Selection.Delete()

                    Case Else ' Normal text
                        ' Insert normal text without tracking
                        wordApp.ActiveDocument.TrackRevisions = False
                        targetRange.Text = TextArray(i)
                        targetRange.Collapse(WdCollapseDirection.wdCollapseEnd)
                        wordApp.ActiveDocument.TrackRevisions = True
                End Select
            Next

            splash.Close()

            ' Restore the original author and TrackRevisions state
            wordApp.ActiveDocument.BuiltInDocumentProperties("Author").Value = originalAuthor
            wordApp.ActiveDocument.TrackRevisions = originalTrackRevisions

        Catch ex As System.Exception
            MessageBox.Show("Error in InsertMarkupText: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub InsertMarkupTextColor(inputText As String, targetRange As Range)
        Try
            Dim wordApp As Word.Application = Globals.ThisAddIn.Application
            Dim TextArray() As String = {}
            Dim FormatArray() As Integer = {}
            Dim originalFontColor As WdColor = WdColor.wdColorBlack
            Dim originalUnderline As WdUnderline = WdUnderline.wdUnderlineNone
            Dim originalStrikeThrough As Boolean = False
            Dim originalBold As Integer = 0
            Dim originalItalic As Integer = 0

            ' Parse the input text into chunks with formatting information
            ParseText(inputText, TextArray, FormatArray)

            ' Store original font properties from the range
            With targetRange.Font
                originalFontColor = .Color
                originalUnderline = .Underline
                originalStrikeThrough = .StrikeThrough
                originalBold = .Bold
                originalItalic = .Italic
            End With

            ' Insert each text chunk with the appropriate formatting
            For i = 0 To TextArray.Length - 1
                ' Reset formatting to original before each insertion
                With targetRange.Font
                    .Color = originalFontColor
                    .Underline = originalUnderline
                    .StrikeThrough = originalStrikeThrough
                    .Bold = originalBold
                    .Italic = originalItalic
                End With

                ' Insert the text at the target range
                targetRange.Text = TextArray(i)

                ' Define the range for the inserted text
                Dim insertedRange As Range = targetRange.Duplicate
                insertedRange.Start = targetRange.Start
                insertedRange.End = targetRange.Start + TextArray(i).Length

                ' Apply formatting based on the tag
                Select Case FormatArray(i)
                    Case 1 ' [INS_START]...[INS_END]: Blue underline
                        With insertedRange.Font
                            .Color = RGB(0, 0, 255)
                            .Underline = WdUnderline.wdUnderlineSingle
                            .StrikeThrough = False
                        End With
                    Case 2 ' [DEL_START]...[DEL_END]: Red strikethrough
                        With insertedRange.Font
                            .Color = RGB(255, 0, 0)
                            .StrikeThrough = True
                            .Underline = WdUnderline.wdUnderlineNone
                        End With
                    Case Else ' Normal text
                        ' Already reset to original formatting
                End Select

                ' Collapse the range to the end for the next insertion
                targetRange.Collapse(WdCollapseDirection.wdCollapseEnd)
            Next

            ' Ensure formatting is reset after all insertions
            With targetRange.Font
                .Color = originalFontColor
                .Underline = originalUnderline
                .StrikeThrough = originalStrikeThrough
                .Bold = originalBold
                .Italic = originalItalic
            End With
        Catch ex As System.Exception
            MessageBox.Show("Error in InsertMarkupTextColor: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub ParseText(inputText As String, ByRef TextArray() As String, ByRef FormatArray() As Integer)
        Dim pos As Integer = 1
        Dim lenText As Integer = inputText.Length
        Dim nextTagPos As Integer
        Dim tagEndPos As Integer
        Dim tagText As String
        Dim chunkIndex As Integer = 0
        Dim tagType As Integer
        Dim nextInsPos As Integer
        Dim nextDelPos As Integer

        While pos <= lenText
            If inputText.Substring(pos - 1, System.Math.Min(11, lenText - pos + 1)) = "[INS_START]" Then
                pos += 11
                tagType = 1 ' Insert formatting
                tagEndPos = inputText.IndexOf("[INS_END]", pos - 1) + 1
                If tagEndPos = -1 Then
                    MessageBox.Show("Error in ParseText: Missing [INS_END] tag.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                tagText = inputText.Substring(pos - 1, tagEndPos - pos)
                pos = tagEndPos + 9
            ElseIf inputText.Substring(pos - 1, System.Math.Min(11, lenText - pos + 1)) = "[DEL_START]" Then
                pos += 11
                tagType = 2 ' Delete formatting
                tagEndPos = inputText.IndexOf("[DEL_END]", pos - 1) + 1
                If tagEndPos = -1 Then
                    MessageBox.Show("Error in ParseText: Missing [DEL_END] tag.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                tagText = inputText.Substring(pos - 1, tagEndPos - pos)
                pos = tagEndPos + 9
            Else
                tagType = 0
                nextInsPos = inputText.IndexOf("[INS_START]", pos - 1) + 1
                If nextInsPos = 0 Then nextInsPos = lenText + 1
                nextDelPos = inputText.IndexOf("[DEL_START]", pos - 1) + 1
                If nextDelPos = 0 Then nextDelPos = lenText + 1
                nextTagPos = System.Math.Min(nextInsPos, nextDelPos)
                tagText = inputText.Substring(pos - 1, nextTagPos - pos)
                pos = nextTagPos
            End If

            chunkIndex += 1
            ReDim Preserve TextArray(chunkIndex - 1)
            ReDim Preserve FormatArray(chunkIndex - 1)
            TextArray(chunkIndex - 1) = tagText
            FormatArray(chunkIndex - 1) = tagType
        End While
    End Sub
    Private Function RGB(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer) As Integer
        Return red Or (green << 8) Or (blue << 16)
    End Function


    Private Function GetTextWithSpecialElementsInline(ByRef workingrange As Word.Range, PreserveParagraphFormatInline As Boolean) As String

        Try
            Dim resultText As String = workingrange.Text
            Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Dim currentOffset As Integer = 0
            Dim paraCount As Integer = workingrange.Paragraphs.Count

            Dim StartOfRange As Integer = workingrange.Start
            Dim EndOfRange As Integer = workingrange.End
            Dim EndOfDocument As Boolean = If(EndOfRange < doc.Content.End, False, True)
            If doc.Bookmarks.Exists("RTEX1") Then
                doc.Bookmarks("RTEX1").Delete()
            End If
            If Not EndOfDocument Then
                doc.Bookmarks.Add("RTEX1", doc.Range(EndOfRange, EndOfRange))
            End If

            ' Process footnotes in the selected range
            For Each footnote As Word.Footnote In workingrange.Footnotes
                Dim footnoteText As String = $"{{{{WFNT:{footnote.Range.Text}}}}}"

                ' Find the exact position of the footnote reference
                Dim referenceStart As Integer = footnote.Reference.Start - workingrange.Start
                Dim referenceLength As Integer = footnote.Reference.End - footnote.Reference.Start

                ' Replace the footnote reference directly in the document
                Dim referenceRange As Word.Range = doc.Range(footnote.Reference.Start, footnote.Reference.End)
                referenceRange.Text = footnoteText
            Next

            ' Process endnotes in the selected range
            For Each endnote As Word.Endnote In workingrange.Endnotes
                Dim endnoteText As String = $"{{{{WENT:{endnote.Range.Text}}}}}"

                ' Find the exact position of the endnote reference
                Dim referenceStart As Integer = endnote.Reference.Start - workingrange.Start
                Dim referenceLength As Integer = endnote.Reference.End - endnote.Reference.Start

                ' Replace the endnote reference directly in the document
                Dim referenceRange As Word.Range = doc.Range(endnote.Reference.Start, endnote.Reference.End)
                referenceRange.Text = endnoteText
            Next

            ' Process field in the selected range
            For Each field As Word.Field In workingrange.Fields
                Dim fieldCode As String = field.Code.Text.Trim() ' Field code (e.g., "HYPERLINK \"http://example.com\"")
                Dim fieldText As String = $"{{{{WFLD:{fieldCode}}}}}" ' Store only the field code

                ' Find the exact position of the field reference
                Dim fieldRange As Word.Range = field.Result

                ' Replace the field reference directly in the document
                field.Delete()
                fieldRange.Text = fieldText

            Next

            workingrange.Start = StartOfRange
            If doc.Bookmarks.Exists("RTEX1") Then
                workingrange.End = doc.Bookmarks("RTEX1").Range.Start
                doc.Bookmarks("RTEX1").Delete()
            End If
            If EndOfDocument Then workingrange.End = doc.Content.End

            Dim updatedRange As Word.Range = doc.Range(workingrange.Start, workingrange.End)

            updatedRange.TextRetrievalMode.IncludeHiddenText = True
            updatedRange.TextRetrievalMode.IncludeFieldCodes = True

            Dim preservedText As New StringBuilder(updatedRange.Text)

            If PreserveParagraphFormatInline Then

                paraCount = updatedRange.Paragraphs.Count

                ReDim paragraphFormat(paraCount - 1)
                Array.Clear(paragraphFormat, 0, paragraphFormat.Length)

                ' Process each paragraph
                For i As Integer = 1 To paraCount
                    Dim para As Word.Paragraph = updatedRange.Paragraphs(i)

                    ' Check if the paragraph range is fully contained in the working range
                    If para.Range.Start >= updatedRange.Start AndAlso para.Range.End <= updatedRange.End Then
                        ' Store all relevant paragraph formatting settings
                        paragraphFormat(i - 1) = New ParagraphFormatStructure With {
                        .Style = para.Style,
                        .FontName = para.Range.Font.Name,
                        .FontSize = para.Range.Font.Size,
                        .FontBold = para.Range.Font.Bold,
                        .FontItalic = para.Range.Font.Italic,
                        .FontUnderline = para.Range.Font.Underline,
                        .FontColor = para.Range.Font.Color,
                        .ListType = para.Range.ListFormat.ListType,
                        .ListTemplate = If(para.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering, para.Range.ListFormat.ListTemplate, Nothing),
                        .ListLevel = If(para.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering, para.Range.ListFormat.ListLevelNumber, 0),
                        .ListNumber = If(para.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering, para.Range.ListFormat.ListValue, 0),
                        .HasListFormat = para.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                        .Alignment = para.Alignment,
                        .LineSpacing = para.LineSpacing,
                        .SpaceBefore = para.SpaceBefore,
                        .SpaceAfter = para.SpaceAfter
                        }

                        ' Insert the placeholder PFOR:nnn into the string builder
                        Dim placeholder As String = $"{{{{PFOR:{i - 1}}}}}"
                        preservedText.Insert(para.Range.Start - updatedRange.Start + currentOffset, placeholder)

                        ' Adjust the offset to account for the newly inserted placeholder
                        currentOffset += placeholder.Length
                    End If
                Next
            End If
            Return preservedText.ToString()

        Catch ex As System.Exception
            'MsgBox("An error occurred: " & ex.Message & " " & ex.Source, MsgBoxStyle.Critical)
            Return workingrange.Text
        End Try

    End Function


    Private Sub RestoreSpecialTextElements(workingrange As Word.Range)

        Try
            Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

            Dim StartOfRange As Integer = workingrange.Start
            Dim EndOfRange As Integer = workingrange.End
            Dim EndOfDocument As Boolean = If(EndOfRange < doc.Content.End, False, True)
            If doc.Bookmarks.Exists("RTEX1") Then
                doc.Bookmarks("RTEX1").Delete()
            End If
            If Not EndOfDocument Then
                doc.Bookmarks.Add("RTEX1", doc.Range(EndOfRange, EndOfRange))
            End If

            ' Process Footnotes
            ProcessInTextPlaceholders(workingrange, doc, "WFNT:", AddressOf AddFootnote)
            workingrange.Start = StartOfRange
            If doc.Bookmarks.Exists("RTEX1") Then
                workingrange.End = doc.Bookmarks("RTEX1").Range.Start
            End If
            If EndOfDocument Then workingrange.End = doc.Content.End

            ' Process Endnotes
            ProcessInTextPlaceholders(workingrange, doc, "WENT:", AddressOf AddEndnote)
            workingrange.Start = StartOfRange
            If doc.Bookmarks.Exists("RTEX1") Then
                workingrange.End = doc.Bookmarks("RTEX1").Range.Start
            End If
            If EndOfDocument Then workingrange.End = doc.Content.End

            ' Process Fields
            ProcessInTextPlaceholders(workingrange, doc, "WFLD:", AddressOf AddField)
            workingrange.Start = StartOfRange
            If doc.Bookmarks.Exists("RTEX1") Then
                workingrange.End = doc.Bookmarks("RTEX1").Range.Start
            End If
            If EndOfDocument Then workingrange.End = doc.Content.End

            ' Process Formatting
            ProcessInTextPlaceholders(workingrange, doc, "PFOR:", AddressOf AddFormat)
            workingrange.Start = StartOfRange
            If doc.Bookmarks.Exists("RTEX1") Then
                workingrange.End = doc.Bookmarks("RTEX1").Range.Start
                doc.Bookmarks("RTEX1").Delete()
            End If
            If EndOfDocument Then workingrange.End = doc.Content.End

        Catch ex As System.Exception
            'MsgBox("An error occurred: " & ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub
    Private Sub ProcessInTextPlaceholders(ByRef workingrange As Word.Range, doc As Word.Document, placeholderPrefix As String, addNoteAction As Action(Of Word.Document, Word.Range, String))
        Dim PreserveRange As Range = workingrange
        With workingrange.Find
            .Text = "\{\{" & placeholderPrefix & "*\}\}"
            .MatchWildcards = True
            Do While .Execute()
                Dim startPos As Integer = workingrange.Start
                Dim endPos As Integer = workingrange.End

                ' Extract note or field text by trimming prefix and suffix
                Dim placeholderText As String = workingrange.Text
                Dim noteText As String = placeholderText.Substring(placeholderPrefix.Length + 2, placeholderText.Length - (placeholderPrefix.Length + 4))

                ' Remove placeholder
                workingrange.Text = ""

                ' Add footnote, endnote, or field
                Dim insertionRange As Word.Range = doc.Range(startPos, startPos)
                addNoteAction.Invoke(doc, insertionRange, noteText)

                ' Adjust range position for the next match
                ' workingrange.Start = endPos + 1
                workingrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                If placeholderPrefix = "PFOR:" Then
                    workingrange.MoveStart(Unit:=Word.WdUnits.wdParagraph, Count:=1)
                    If workingrange.Start < doc.Content.End Then
                        Dim nextChar As String = doc.Range(workingrange.Start, workingrange.Start + 1).Text
                        ' In Word the paragraph mark is typically Chr(13) (which is vbCr)
                        If nextChar = vbCr Then
                            workingrange.MoveStart(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                        End If
                        If nextChar = vbLf Then
                            workingrange.MoveStart(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                        End If
                    End If
                Else
                    workingrange.MoveStart(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                End If
            Loop
        End With
    End Sub
    Private Sub AddFootnote(doc As Word.Document, insertionRange As Word.Range, noteText As String)
        doc.Footnotes.Add(Range:=insertionRange, Text:=noteText)
    End Sub
    Private Sub AddEndnote(doc As Word.Document, insertionRange As Word.Range, noteText As String)
        doc.Endnotes.Add(Range:=insertionRange, Text:=noteText)
    End Sub
    Private Sub AddField(doc As Word.Document, insertionRange As Word.Range, fieldText As String)
        ' Use the fieldText directly as the field code
        Dim fieldCode As String = fieldText.Trim()

        ' Insert the field
        Dim fieldRange As Word.Range = insertionRange.Duplicate
        Dim field As Word.Field = doc.Fields.Add(fieldRange)
        field.Code.Text = fieldCode

        ' Update the field to display its calculated result
        field.Update()
    End Sub
    Private Sub AddFormat(doc As Word.Document, insertionRange As Word.Range, formatIndexText As String)
        Try
            ' Parse the format index from the input text
            Dim formatIndex As Integer = Integer.Parse(formatIndexText.Trim())

            ' Ensure the format index is within bounds
            If formatIndex >= 0 AndAlso formatIndex < paragraphFormat.Length Then
                ' Retrieve the specific paragraph format
                Dim format = paragraphFormat(formatIndex)

                ' Expand the range to the entire paragraph
                Dim targetRange As Word.Range = insertionRange.Paragraphs(1).Range
                If targetRange.End > targetRange.Start Then
                    targetRange.End = targetRange.End - 1  ' Exclude the paragraph mark
                End If

                With targetRange
                    ' Apply the stored style
                    If format.Style IsNot Nothing Then .Style = format.Style

                    ' Apply the stored font formatting
                    With .Font
                        If format.FontName IsNot Nothing Then .Name = format.FontName
                        If format.FontSize > 0 Then .Size = format.FontSize
                        .Bold = format.FontBold
                        .Italic = format.FontItalic
                        .Underline = format.FontUnderline
                        .Color = format.FontColor
                    End With

                    ' Apply list formatting if applicable
                    If format.HasListFormat AndAlso format.ListTemplate IsNot Nothing Then
                        Try
                            .ListFormat.ApplyListTemplateWithLevel(
                            ListTemplate:=format.ListTemplate,
                            ContinuePreviousList:=If(format.ListNumber > 0, True, False),
                            ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList,
                            DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior
                        )
                            .ListFormat.ListLevelNumber = format.ListLevel
                        Catch ex As System.Exception
                            'MsgBox("Error applying list format: " & ex.Message, MsgBoxStyle.Exclamation)
                        End Try
                    End If

                    ' Apply paragraph alignment
                    .ParagraphFormat.Alignment = format.Alignment

                    ' Apply line spacing
                    .ParagraphFormat.LineSpacing = format.LineSpacing

                    ' Apply spacing before and after
                    .ParagraphFormat.SpaceBefore = format.SpaceBefore
                    .ParagraphFormat.SpaceAfter = format.SpaceAfter
                End With
            Else
                ' MsgBox("Invalid format index: " & formatIndex, MsgBoxStyle.Exclamation)
            End If
        Catch ex As System.Exception
            ' MsgBox("Error applying format: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' Word Helper Functions

    Public Sub ImportTextFile()
        Dim sel As Word.Range = Globals.ThisAddIn.Application.Selection.Range
        Dim Doc = GetFileContent()
        sel.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
        sel.Text = Doc
        'sel.End = sel.Start + Doc.Length
        sel.Select()
    End Sub

    Public Sub AcceptFormatting()

        Dim sel As Word.Range = Globals.ThisAddIn.Application.Selection.Range
        Dim formatChangeCount As Integer = 0
        Dim DocRef As String = "in the selected text"

        ' Ensure a selection is made
        If sel Is Nothing OrElse String.IsNullOrWhiteSpace(sel.Text) Then
            sel = Globals.ThisAddIn.Application.ActiveDocument.Content
            DocRef = "in the document"
        End If

        ' Check if there are any markups
        If sel.Revisions.Count = 0 Then
            ShowCustomMessageBox($"No revisions found {DocRef}.")
            Return
        End If

        Dim splash As New SplashScreen("Accepting revisions related to formatting... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        ' Loop through all markups in the range
        For Each rev As Word.Revision In sel.Revisions

            System.Windows.Forms.Application.DoEvents()

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                ' Exit the loop
                Exit For
            End If

            ' Check if the revision is a formatting change (exclude text insertions/deletions)
            Select Case rev.Type
                Case Word.WdRevisionType.wdRevisionProperty,
                 Word.WdRevisionType.wdRevisionParagraphNumber,
                 Word.WdRevisionType.wdRevisionParagraphProperty,
                 Word.WdRevisionType.wdRevisionSectionProperty,
                 Word.WdRevisionType.wdRevisionStyle,
                 Word.WdRevisionType.wdRevisionStyleDefinition,
                 Word.WdRevisionType.wdRevisionTableProperty

                    ' Accept the revision
                    rev.Accept()
                    formatChangeCount += 1
            End Select
        Next

        splash.Close()

        ' Provide final feedback
        If formatChangeCount > 0 Then
            ShowCustomMessageBox(formatChangeCount.ToString() & $" formatting revision(s) {DocRef} (including paragraph numbering) have been found and accepted.")
        Else
            ShowCustomMessageBox($"No revisions related (only) to formatting were found {DocRef}.")
        End If
    End Sub



    Private Async Sub ShowPaneAsync(
                              introLine As String,
                              bodyText As String,
                              finalRemark As String,
                              header As String,
                              Optional noRTF As Boolean = False,
                              Optional insertMarkdown As Boolean = False
                            )
        Try

            Dim OriginalText As String = bodyText

            Dim result As String = Await PaneManager.ShowMyPane(introLine, bodyText, finalRemark, header, noRTF, insertMarkdown, New IntelligentMergeCallback(AddressOf HandleIntelligentMerge))

            If result <> "" Then
                If result = "Markdown" Then
                    Dim currentSelection As Word.Selection = Globals.ThisAddIn.Application.Selection
                    currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                    InsertTextWithMarkdown(currentSelection, OriginalText, False)
                End If
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in ShowPaneAsync: " & ex.Message)
        End Try
    End Sub


    Private Sub HandleIntelligentMerge(selectedText As String)
        ' Hier Deine bestehende Merge-Logik aufrufen:
        IntelligentMerge(selectedText)
    End Sub

    Public Async Sub IntelligentMerge(newtext As String)
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection
        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text in your document with which your selection in the pane shall be merged.")
            Return
        End If
        OtherPrompt = SLib.ShowCustomInputBox("If you want, you can amend the prompt that will be used to intelligently merge your selection into your document:", $"{AN} Intelligent Merge", False, SP_MergePrompt_Cached).Trim()
        If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return
        Dim result As String = Await ProcessSelectedText(OtherPrompt & " " & SP_Add_MergePrompt & " <INSERT>" & newtext & "</INSERT> ", True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub

    Public ONNX_initialized As Boolean = False

    Private Function EnsureInitialized() As Boolean

        If Not ONNX_initialized AndAlso Not String.IsNullOrEmpty(Globals.ThisAddIn.INI_LocalModelPath) Then
            ' Pfade an dein Add-In anpassen oder aus der Config laden
            Try
                Dim modelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_LocalModelPath), NER_Model)
                Dim vocabpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_LocalModelPath), NER_Token)
                Dim labelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_LocalModelPath), NER_Label)

                OnnxAnonymizer.Initialize(modelpath, vocabpath, labelpath, 128)
                ONNX_initialized = True
                Return True
            Catch ex As Exception
                SLib.ShowCustomMessageBox($"Error loading and initializing the NER model ({ex.Message}).")
                ONNX_initialized = False
                Return False
            End Try
        Else
            Return ONNX_initialized
        End If
    End Function


    Public Sub AnonymizeSelection()

        Dim sel As String = Globals.ThisAddIn.Application.Selection.Text

        If String.IsNullOrWhiteSpace(sel) Then
            SLib.ShowCustomMessageBox("Please select a text to anonymize.")
            Return
        End If

        Dim AnonSetting As String = INI_Anon
        Dim OverrideAnonSetting As String = LoadAnonSettingsForModel(INI_Model)

        If Not String.IsNullOrWhiteSpace(OverrideAnonSetting) Then AnonSetting = OverrideAnonSetting
        If Not String.IsNullOrWhiteSpace(AnonSetting) Then
            Dim AnonType As Integer = ShowCustomYesNoBox($"Which anonymization type do you want (using keys for '{INI_Model}')?", "3 - file based", "4 - prompt", $"{AN} Anonymization") + 2
            If AnonType > 2 Then
                Dim AnonMode As String = "silent"
                Dim AnonText As String = AnonymizeText(sel, INI_Model, AnonMode, AnonType)
                AnonText = AnonText & vbCrLf & vbCrLf & "**Entities:**  " & vbCrLf & vbCrLf & ExportEntitiesMappings()
                Dim result As String = ShowCustomWindow("The anonymization returned the following text:", AnonText, $"Beware that this anonymization depends entirely on the keys you provided in your file '{AnonFile}' (for your model '{INI_Model}') or your prompt. Check the result. Choose what to put into the clipboard.", $"{AN} Anonymization", True)

                If result <> "" Then
                    SLib.PutInClipboard(result)
                End If
            End If
        End If

        Return

        If String.IsNullOrEmpty(Globals.ThisAddIn.INI_LocalModelPath) Then
            SLib.ShowCustomMessageBox("No path set for the NER model ('LocalModelPath').")
            Return
        End If

        If Not EnsureInitialized() Then Return

        'Dim sel As String = Globals.ThisAddIn.Application.Selection.Text
        If String.IsNullOrWhiteSpace(sel) Then
            SLib.ShowCustomMessageBox("Please select a text to anonymize.")
            Return
        End If

        Dim anon As String = OnnxAnonymizer.Anonymize(sel)

        Dim sb As New StringBuilder()
        sb.AppendLine(anon)
        sb.AppendLine()
        sb.AppendLine("Entity-Mapping:")
        sb.AppendLine()
        For Each kvp In OnnxAnonymizer.mapping
            sb.AppendLine($"{kvp.Key} -> {kvp.Value}")
        Next

        Dim FinalText As String = ShowCustomWindow("The NER anonymization returned the following text:", sb.ToString(), "Beware that this anonymization method is fast, but not of very high precision. Check the result.", AN, True)

        If FinalText <> "" Then
            SLib.PutInClipboard(FinalText)
        End If

    End Sub


    Private Shared LastRegexPattern As String = String.Empty
    Private Shared LastRegexOptions As String = String.Empty
    Private Shared LastRegexReplace As String = String.Empty
    Public Sub RegexSearchReplace()
        Dim sel As Word.Range = Globals.ThisAddIn.Application.Selection.Range
        Dim DocRef As String = "in the selected text"

        ' Ensure a selection is made
        If sel Is Nothing OrElse String.IsNullOrWhiteSpace(sel.Text) Then
            Globals.ThisAddIn.Application.ActiveDocument.Content.Select()
            sel = Globals.ThisAddIn.Application.Selection.Range
            DocRef = "in the document"
        End If

        ' Step 1: Get regex patterns
        Dim RegexPattern As String = ShowCustomInputBox("Step 1: Enter your Regex pattern(s), one per line (more info about Regex: vischerlnk.com/regexinfo):", "Regex Search & Replace", False, LastRegexPattern)?.Trim()
        If String.IsNullOrEmpty(RegexPattern) Then Return

        ' Step 2: Get regex options
        Dim optionsInput As String = ShowCustomInputBox("Enter regex option(s) (i for IgnoreCase, m for Multiline, s for Singleline, c for Compiled, r for RightToLeft, e for ExplicitCapture):", "Regex Search & Replace", True, LastRegexOptions)

        Dim regexOptions As RegexOptions = RegexOptions.None

        If Not String.IsNullOrEmpty(optionsInput) Then
            ' Add specific options based on user input
            If optionsInput.Contains("i") Then regexOptions = regexOptions Or RegexOptions.IgnoreCase
            If optionsInput.Contains("m") Then regexOptions = regexOptions Or RegexOptions.Multiline
            If optionsInput.Contains("s") Then regexOptions = regexOptions Or RegexOptions.Singleline
            If optionsInput.Contains("c") Then regexOptions = regexOptions Or RegexOptions.Compiled
            If optionsInput.Contains("r") Then regexOptions = regexOptions Or RegexOptions.RightToLeft
            If optionsInput.Contains("e") Then regexOptions = regexOptions Or RegexOptions.ExplicitCapture
        End If

        ' Step 3: Get replacement text
        Dim Replacementtext As String = ShowCustomInputBox("Step 2: Enter your replacement text(s), one on each line, matching to your pattern(s) (leave empty or cancel to only search for the first hit):", "Regex Search & Replace", False, LastRegexReplace)

        ' Update the last-used regex pattern and options
        LastRegexPattern = RegexPattern
        LastRegexOptions = optionsInput
        LastRegexReplace = Replacementtext

        ' Split patterns and replacements into lines
        Dim patterns() As String = RegexPattern.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
        Dim replacements() As String = If(Not String.IsNullOrEmpty(Replacementtext), Replacementtext.Split(New String() {Environment.NewLine}, StringSplitOptions.None), Nothing)

        ' Check if patterns and replacements match
        If replacements IsNot Nothing AndAlso patterns.Length <> replacements.Length Then
            ShowCustomMessageBox("The number of regex patterns does not match the number of replacement lines. Aborting without any replacements done.")
            Return
        End If

        ' Validate all regex patterns first
        For Each pattern As String In patterns
            Try
                Dim regexTest As New Regex(pattern, regexOptions)
            Catch ex As ArgumentException
                ShowCustomMessageBox($"Your regex pattern '{pattern}' is invalid ({ex.Message}). Aborting without any replacements done.")
                Return
            End Try
        Next

        ' Perform replacements after validation
        Dim totalReplacements As Integer = 0

        For i As Integer = 0 To patterns.Length - 1
            Dim pattern As String = patterns(i)
            Dim replacement As String = If(replacements IsNot Nothing, replacements(i), Nothing)

            Dim regex As New Regex(pattern, regexOptions)

            If Not String.IsNullOrEmpty(replacement) Then
                ' Perform replacement
                Dim replacementCount As Integer = 0
                sel.Text = regex.Replace(sel.Text, Function(match)
                                                       replacementCount += 1
                                                       Return replacement
                                                   End Function)
                totalReplacements += replacementCount
            Else
                ' Perform search only
                Dim match As Match = regex.Match(sel.Text)
                If match.Success Then
                    ' Highlight the first match
                    sel.Start = sel.Start + match.Index
                    sel.End = sel.Start + match.Length
                    Globals.ThisAddIn.Application.Selection.Select()
                    Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(sel, True)
                    Return
                Else
                    ShowCustomMessageBox($"No matches found for '{pattern}' {DocRef}.")
                    Return
                End If
            End If
        Next

        If replacements IsNot Nothing Then
            ShowCustomMessageBox($"{totalReplacements} replacement(s) made {DocRef}.")
        Else
            ShowCustomMessageBox("Search complete. No replacements were made.")
        End If
    End Sub

    Public Sub CalculateUserMarkupTimeSpan()

        Try
            Dim userName As String
            Dim docRevisions As Word.Revisions
            Dim rev As Word.Revision
            Dim comment As Word.Comment
            Dim firstTimestamp As Date
            Dim lastTimestamp As Date
            Dim found As Boolean
            Dim userInput As String
            Dim userNames As New Collection
            Dim selRange As Word.Range
            Dim outputUserNames As String
            Dim DocRef As String = "in the selected text"

            ' Initialize
            found = False
            firstTimestamp = #1/1/1900# ' Default initialization
            lastTimestamp = #1/1/1900# ' Default initialization

            ' Prompt for user input
            userName = Globals.ThisAddIn.Application.UserName

            ' Prompt for user input
            userInput = ShowCustomInputBox("Please enter the name of the user (leave empty for all users):", "Markup Time Span", True, userName)
            userInput = userInput.Trim()

            ' Check selection
            If Globals.ThisAddIn.Application.Selection Is Nothing OrElse String.IsNullOrWhiteSpace(Globals.ThisAddIn.Application.Selection.Range.Text) Then
                ' If no selection, select the entire document
                Globals.ThisAddIn.Application.ActiveDocument.Content.Select()
                DocRef = "in the document"
            End If
            selRange = Globals.ThisAddIn.Application.Selection.Range
            docRevisions = selRange.Revisions ' Only consider changes in the selected range

            ' Process revisions
            For Each rev In docRevisions
                If String.IsNullOrEmpty(userInput) OrElse rev.Author.Equals(userInput, StringComparison.OrdinalIgnoreCase) Then
                    ' Update timestamps
                    If Not found Then
                        firstTimestamp = rev.Date
                        lastTimestamp = rev.Date
                        found = True
                    Else
                        If rev.Date < firstTimestamp Then firstTimestamp = rev.Date
                        If rev.Date > lastTimestamp Then lastTimestamp = rev.Date
                    End If
                    ' Collect user names if processing all
                    Try
                        userNames.Add(rev.Author, rev.Author.ToLower())
                    Catch ex As Exception
                        ' Ignore duplicates
                    End Try
                End If
            Next

            ' Process comments
            For Each comment In selRange.Comments
                If String.IsNullOrEmpty(userInput) OrElse comment.Author.Equals(userInput, StringComparison.OrdinalIgnoreCase) Then
                    ' Update timestamps
                    If Not found Then
                        firstTimestamp = comment.Date
                        lastTimestamp = comment.Date
                        found = True
                    Else
                        If comment.Date < firstTimestamp Then firstTimestamp = comment.Date
                        If comment.Date > lastTimestamp Then lastTimestamp = comment.Date
                    End If
                    ' Collect user names if processing all
                    Try
                        userNames.Add(comment.Author, comment.Author.ToLower())
                    Catch ex As Exception
                        ' Ignore duplicates
                    End Try
                End If
            Next

            ' Display results
            If found Then
                Dim timeSpan As String
                Dim timeDiff As Double
                timeDiff = DateDiff(DateInterval.Minute, firstTimestamp, lastTimestamp) ' Time difference in minutes
                timeSpan = Math.Floor(timeDiff / 1440).ToString() & " days, " &
                       ((timeDiff Mod 1440) \ 60).ToString("00") & " hours, " &
                       (timeDiff Mod 60).ToString("00") & " minutes"

                ' Format timestamps without seconds
                Dim formattedFirstTimestamp As String
                Dim formattedLastTimestamp As String
                formattedFirstTimestamp = firstTimestamp.ToString("dd/MM/yyyy HH:mm")
                formattedLastTimestamp = lastTimestamp.ToString("dd/MM/yyyy HH:mm")
                If String.IsNullOrEmpty(userInput) Then
                    ' Display all users
                    Dim user As Object
                    outputUserNames = "Users involved:" & vbCrLf
                    For Each user In userNames
                        outputUserNames &= "- " & user.ToString() & vbCrLf
                    Next
                Else
                    outputUserNames = "User: " & userInput
                End If
                ShowCustomMessageBox(outputUserNames & vbCrLf & "First markup/comment: " & formattedFirstTimestamp & vbCrLf &
    "Last markup/comment: " & formattedLastTimestamp & vbCrLf &
    "Time span: " & timeSpan)
            Else
                If String.IsNullOrEmpty(userInput) Then
                    ShowCustomMessageBox($"No markups or comments found {DocRef}.")
                Else
                    ShowCustomMessageBox("No markups or comments found for user '" & userInput & $"' {DocRef}.")
                End If
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in CalculateUserMarkupTimeSpan: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub CompareSelectionHalves()

        Dim sel As Word.Range
        Dim nonEmptyParaCount As Long
        Dim halfParaCount As Long
        Dim firstRange As Word.Range
        Dim secondRange As Word.Range
        Dim paraIndices() As Long
        Dim i As Long, index As Long

        ' Get the selected text
        sel = Globals.ThisAddIn.Application.Selection.Range

        ' Count non-empty paragraphs and store their indices
        ReDim paraIndices(0 To sel.Paragraphs.Count - 1)
        index = 0
        For i = 1 To sel.Paragraphs.Count
            If Len(sel.Paragraphs(i).Range.Text.Trim()) > 1 Then ' Greater than 1 to account for paragraph mark
                index += 1
                paraIndices(index - 1) = i
            End If
        Next

        ' Update nonEmptyParaCount
        nonEmptyParaCount = index

        ' If number of non-empty paragraphs is uneven or zero, abort
        If nonEmptyParaCount Mod 2 <> 0 Or nonEmptyParaCount = 0 Then
            ShowCustomMessageBox("The number of non-empty paragraphs in the selection is uneven or zero. Please select an even number of non-empty paragraphs.")
            Return
        End If

        ' Determine the halfway point
        halfParaCount = nonEmptyParaCount \ 2

        ' Get the first half and second half ranges
        firstRange = sel.Paragraphs(paraIndices(0)).Range
        firstRange.End = sel.Paragraphs(paraIndices(halfParaCount - 1)).Range.End

        secondRange = sel.Paragraphs(paraIndices(halfParaCount)).Range
        secondRange.End = sel.Paragraphs(paraIndices(nonEmptyParaCount - 1)).Range.End


        ' Get text from the first and second range without the final paragraph marks
        Dim text1 As String = Left(firstRange.Text, Len(firstRange.Text) - 1)
        Dim text2 As String = Left(secondRange.Text, Len(secondRange.Text) - 1)

        If INI_MarkupMethodHelper <> 1 Then
            CompareAndInsert(text1, text2, secondRange, INI_MarkupMethodHelper = 3, "These are the differences of the second (set of) paragraph(s) of the text selected:")
        Else
            CompareAndInsertComparedoc(text1, text2, secondRange)
        End If
    End Sub

    Private embed_store As EmbeddingStore
    Private embed_indexedDocs As HashSet(Of String) = New HashSet(Of String)()

    Public Async Sub ContextSearch()

        Dim EmbedModel As String = ""
        Dim EmbedVocab As String = ""

        If Not String.IsNullOrEmpty(INI_LocalModelPath) Then

            EmbedModel = System.IO.Path.Combine(ExpandEnvironmentVariables(INI_LocalModelPath), Embed_Model)
            EmbedVocab = System.IO.Path.Combine(ExpandEnvironmentVariables(INI_LocalModelPath), Embed_Vocab)

            If File.Exists(EmbedModel) And File.Exists(EmbedVocab) Then
                If embed_store Is Nothing Then embed_store = New EmbeddingStore(EmbedModel, EmbedVocab)
            End If
        End If

        Dim EmbeddingAvailable As Boolean = Not embed_store Is Nothing

        Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
        Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.ActiveDocument
        Dim DoSearchNext As Boolean = False
        Dim EmbedInstruct As String = If(EmbeddingAvailable, $"add '{EmbedTrigger} to use embeddings, ", "")
        Dim DoBoW As Boolean = False
        Dim DoRefresh As Boolean = False
        Dim DoEmbed As Boolean = False

        Dim lastcontextsearch As String = If(String.IsNullOrWhiteSpace(My.Settings.LastContextSearch), "", My.Settings.LastContextSearch)

        SearchContext = ShowCustomInputBox($"Enter the search term (use '{SearchNextTrigger}' if you only want to find the next term; {EmbedInstruct}'{BoWTrigger}' to use Bag of Words and '{RefreshTrigger}' to refresh the index first):", "Context Search", True, lastcontextsearch).Trim()
        If String.IsNullOrWhiteSpace(SearchContext) Or SearchContext = "ESC" Then Return

        My.Settings.LastContextSearch = SearchContext
        My.Settings.Save()

        If SearchContext.StartsWith(SearchNextTrigger, StringComparison.OrdinalIgnoreCase) Then
            SearchContext = SearchContext.Substring(SearchNextTrigger.Length).Trim()
            DoSearchNext = True
        End If

        If SearchContext.IndexOf(EmbedTrigger, StringComparison.OrdinalIgnoreCase) >= 0 And EmbeddingAvailable Then
            SearchContext = SearchContext.Replace(EmbedTrigger, "").Trim()
            DoEmbed = True
        ElseIf SearchContext.IndexOf(BoWTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            SearchContext = SearchContext.Replace(BoWTrigger, "").Trim()
            DoBow = True
        End If
        If SearchContext.IndexOf(RefreshTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            SearchContext = SearchContext.Replace(RefreshTrigger, "").Trim()
            DoRefresh = True
        End If

        SearchContext = SearchContext.Replace("  ", "")

        If DoEmbed Then
            RunSearch_Embed(EmbedModel, EmbedVocab, SearchContext, DoSearchNext, DoRefresh)
            Return
        ElseIf DoBoW Then
            RunSearch_bow(SearchContext, DoSearchNext, DoRefresh)
            Return
        End If

        Dim SearchText As String = ""

        If Not String.IsNullOrWhiteSpace(selection.Text) And Len(selection.Text) > 3 And DoSearchNext Then
            SearchText = selection.Text
        ElseIf selection.Start < selection.Document.Content.End And DoSearchNext Then
            SearchText = selection.Document.Range(selection.Start, selection.Document.Content.End).Text
            selection.SetRange(selection.Start, selection.Document.Content.End)
        Else
            SearchText = selection.Document.Content.Text
            selection.SetRange(0, selection.Document.Content.End)
            DoSearchNext = False
        End If

        Dim LLMResult As String = Await LLM(InterpolateAtRuntime(If(DoSearchNext, SP_ContextSearch, SP_ContextSearchMulti)), "<TEXTTOSEARCH>" & SearchText & "</TEXTTOSEARCH>", "", "", 0)

        LLMResult = LLMResult.Replace("<TEXTTOSEARCH>", "").Replace("</TEXTTOSEARCH>", "")

        If Not DoSearchNext Then

            Dim parts() As String = LLMResult.Split(New String() {"@@@"}, StringSplitOptions.RemoveEmptyEntries)
            Dim notFoundParts As New List(Of String)
            Dim originalStart As Integer = selection.Start
            Dim originalEnd As Integer = selection.End

            If parts.Count > 0 Then

                Dim splash As New SplashScreen($"Highlighting hits... Press 'Esc' to abort")
                splash.Show()
                splash.Refresh()

                Dim Aborted As Boolean = False

                Dim trackChangesEnabled As Boolean = doc.TrackRevisions
                Dim originalAuthor As String = doc.Application.UserName

                doc.TrackRevisions = True

                Dim SuccessHits As Integer = 0

                For Each part As String In parts

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                        Aborted = True
                        Exit For
                    End If

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                        ' Exit the loop
                        Aborted = True
                        Exit For
                    End If

                    Dim findText As String = part.Trim()
                    If FindLongTextInChunks(findText, 255, selection) And selection IsNot Nothing Then
                        'selection.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow
                        doc.Comments.Add(selection.Range, $"{AN5}: '{SearchContext}'")
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        SuccessHits += 1
                    Else
                        notFoundParts.Add(findText)
                    End If
                Next

                splash.Close()

                If Aborted Then
                    ShowCustomMessageBox($"Search aborted. {SuccessHits} hit(s) have been highlighted so far.", "Context Search")
                ElseIf notFoundParts.Count > 0 Then
                    ShowCustomMessageBox($"{SuccessHits} hit(s) have been highlighted. The following hit(s) could not be found:" & vbCrLf & vbCrLf & String.Join(vbCrLf, notFoundParts), "Context Search")
                Else
                    ShowCustomMessageBox($"{SuccessHits} hit(s) have been highlighted.", "Context Search")
                End If

                ' Restore the original selection
                selection.SetRange(originalStart, originalEnd)
                doc.TrackRevisions = trackChangesEnabled

            Else
                ShowCustomMessageBox($"The LLM has found no hits for the context '{SearchContext}'.", "Context Search")
            End If

        Else
            If Not String.IsNullOrWhiteSpace(LLMResult) Then
                Dim FindText As String = LLMResult.Trim()

                If FindLongTextInChunks(FindText, 255, selection) And selection IsNot Nothing Then
                    wordApp.ActiveWindow.ScrollIntoView(selection.Range, True)
                Else
                    ShowCustomMessageBox($"The LLM found this section:" & vbCrLf & vbCrLf & FindText & vbCrLf & vbCrLf & $"However, {AN} could not locate it in the document for technical reasons (may be due to special characters, line breaks of the LLM not quoting the text properly).", "Context Search")
                End If
            Else
                ShowCustomMessageBox($"The LLM did not find any (further) hits for the context '{SearchContext}'.", "Context Search")
            End If
        End If
    End Sub


    Public Sub RunIndexing_Embed(refresh As Boolean, EmbedModel As String, EmbedVocab As String, ChunkLength As Integer, ChunkOverlap As Integer)

        If embed_store Is Nothing Then embed_store = New EmbeddingStore(EmbedModel, EmbedVocab)

        Dim doc = Application.ActiveDocument
        Dim docId = doc.FullName

        ' 0) Early return, wenn schon indexiert und kein Refresh gewünscht
        If embed_indexedDocs.Contains(docId) AndAlso Not refresh Then
            Return
        End If

        ' 1) Parameter validieren
        Dim nn As Integer = ChunkLength     ' Sätze pro Chunk
        Dim mm As Integer = ChunkOverlap     ' Überlappung
        Dim stepSize = nn - mm
        If nn <= 0 OrElse mm < 0 OrElse stepSize <= 0 Then
            Throw New ArgumentException("Bitte nn>0, mm≥0 und mm<nn wählen.")
        End If

        ' 2) Sätze holen und leere filtern
        Dim sentences = doc.Sentences.Cast(Of Range)() _
                        .Where(Function(r) Not String.IsNullOrWhiteSpace(r.Text)) _
                        .ToList()
        Dim total = sentences.Count

        If total < nn Then
            Return
        End If

        ' 3) Chunks bauen (nur volle nn-Satz-Chunks)
        Dim chunks As New List(Of TextChunk)()
        For idx As Integer = 0 To total - nn Step stepSize
            Dim startIdx = idx
            Dim endIdx = idx + nn - 1  ' garantiert ≤ total-1

            ' Text zusammenbauen
            Dim parts = sentences.Skip(startIdx).Take(nn).Select(Function(r) r.Text.Trim())
            Dim chunkText = String.Join(" ", parts)

            ' Sehr kurze Chunks überspringen
            If chunkText.Length < 10 Then
                Continue For
            End If

            ' 4) Offset berechnen – direkt aus Range.Start, kein doc.Range mehr
            Dim rangeStart = sentences(startIdx).Start
            Dim startOffset = If(rangeStart < 0, 0, rangeStart)
            Dim rangeEnd = sentences(endIdx).End

            ' Chunk hinzufügen
            chunks.Add(New TextChunk With {
            .Text = chunkText,
            .StartOffset = startOffset,
            .EndOffset = rangeEnd
        })
        Next
        ' 4) Indexieren
        embed_store.IndexDocument(docId, chunks)
        If Not embed_indexedDocs.Contains(docId) Then embed_indexedDocs.Add(docId)
    End Sub

    Public Sub RunSearch_Embed(EmbedModel As String, EmbedVocab As String, SearchContext As String, DoNext As Boolean, DoRefresh As Boolean)
        Try

            ' 1) Parameter

            Dim ChunkLength As Integer = Default_Embed_Chunks
            Dim ChunkOverlap As Integer = Default_Embed_Overlap
            Dim Min_Score As Double = Default_Embed_Min_Score
            Dim Top_K As Integer = Default_Embed_Top_K
            Dim allDocs As Boolean = False
            Dim Fallback As Boolean = False


            If Not embed_indexedDocs.Contains(Application.ActiveDocument.FullName) Or DoRefresh Then

                Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Sentences per chunk:", ChunkLength),
                    New SLib.InputParameter("Overlap per chunk", ChunkOverlap),
                    New SLib.InputParameter("Minimum relevance", Min_Score),
                    New SLib.InputParameter("Maximum hits", Top_K),
                    New SLib.InputParameter("Always hits", Fallback)
                    }

                If ShowCustomVariableInputForm("Please set your embedding and search values:", $"Context Search (Embedding)", params) Then

                    ChunkLength = CInt(params(0).Value)
                    ChunkOverlap = CInt(params(1).Value)
                    Min_Score = CDbl(params(2).Value)
                    Top_K = CInt(params(3).Value)
                    Fallback = CBool(params(4).Value)

                Else
                    Return
                End If

            Else
                Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Minimum relevance", Min_Score),
                    New SLib.InputParameter("Maximum hits:", Top_K),
                    New SLib.InputParameter("Always hits", Fallback)
                    }
                If ShowCustomVariableInputForm("Please set your search values:", $"Context Search (Embedding)", params) Then

                    Min_Score = CDbl(params(0).Value)
                    Top_K = CBool(params(1).Value)
                    Fallback = CBool(params(2).Value)

                Else
                    Return
                End If

            End If

            ' 2) Für Next-Suche: Selektion zurücksetzen & Cursor ans Ende
            Dim selRange As Word.Range = Application.Selection.Range
            If DoNext Then selRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

            ' 3) Index ggf. neu aufbauen
            RunIndexing_Embed(DoRefresh, EmbedModel, EmbedVocab, ChunkLength, ChunkOverlap)

            Dim currentDocId = Application.ActiveDocument.FullName
            Dim cursorPos = selRange.Start

            If Not DoNext Then
                ' --- COMPLETE: Suche im (Rest)Dokument oder in allen Docs ---
                Dim rawHits = embed_store.Search(SearchContext, allDocs, True, currentDocId, cursorPos) _
                            .Where(Function(r) r.Score > 0 _
                                             AndAlso (allDocs OrElse r.StartOffset > cursorPos)) _
                            .OrderByDescending(Function(r) r.Score) _
                            .ToList()

                ' Treffer über Schwellwert
                Dim scoredHits = rawHits.Where(Function(r) r.Score >= Min_Score).ToList()
                Dim hits As List(Of SearchResult)
                If scoredHits.Count > 0 Then
                    hits = scoredHits.Take(Top_K).ToList()
                ElseIf Fallback Then
                    ' Fallback: die besten TOP_K unabhängig vom Score
                    hits = rawHits.Take(Top_K).ToList()
                End If

                If hits.Count = 0 Then
                    ShowCustomMessageBox($"No hits found for '{SearchContext}'" & If(Fallback, ".", " and minimum relevance of {Min_Score}."))
                    Return
                End If

                Dim trackChangesEnabled As Boolean = Application.ActiveDocument.TrackRevisions
                Application.ActiveDocument.TrackRevisions = True

                For Each r In hits
                    Dim doc = If(r.DocId = currentDocId,
                             Application.ActiveDocument,
                             Application.Documents.Open(r.DocId))
                    Dim rng = doc.Range(r.StartOffset, r.EndOffset)
                    doc.Comments.Add(rng, $"{AN5}: '{SearchContext}' (Score {r.Score:F3})")
                Next

                Application.ActiveDocument.TrackRevisions = trackChangesEnabled

                ShowCustomMessageBox($"{hits.Count} hits found for '{SearchContext}', a minimum relevance of {Min_Score} and a maximum of {Top_K} hits. Comments have been added to them.")
            Else
                ' --- NEXT: Suche nur ab Cursor im aktuellen Dokument ---
                Dim rawHits = embed_store.Search(SearchContext, False, True, currentDocId, cursorPos) _
                            .Where(Function(r) r.Score > 0 AndAlso r.StartOffset > cursorPos) _
                            .OrderByDescending(Function(r) r.Score) _
                            .ToList()

                ' Treffer über Schwellwert
                Dim scoredHits = rawHits.Where(Function(r) r.Score >= Min_Score).ToList()
                Dim hits As List(Of SearchResult)
                If scoredHits.Count > 0 Then
                    hits = scoredHits.Take(Top_K).ToList()
                ElseIf Fallback Then
                    ' Fallback: die besten TOP_K unabhängig vom Score
                    hits = rawHits.Take(Top_K).ToList()
                End If

                If hits.Count = 0 Then
                    ShowCustomMessageBox($"No (further) hits found for '{SearchContext}'" & If(Fallback, ".", " and minimum relevance of {Min_Score}."))
                    Return
                End If

                Dim trackChangesEnabled As Boolean = Application.ActiveDocument.TrackRevisions
                Application.ActiveDocument.TrackRevisions = True

                For Each r In hits
                    Dim doc = If(r.DocId = currentDocId,
                             Application.ActiveDocument,
                             Application.Documents.Open(r.DocId))
                    Dim rng = doc.Range(r.StartOffset, r.EndOffset)
                    doc.Comments.Add(rng, $"{AN5}: '{SearchContext}' (Score {r.Score:F3})")
                Next

                Application.ActiveDocument.TrackRevisions = trackChangesEnabled

                ShowCustomMessageBox($"The (next) {hits.Count} have been found for '{SearchContext}', with a maximum of {Top_K} hits. Comments have been added to them.")
            End If

        Catch ex As Exception
            MessageBox.Show("Error in RunSearch_Embed: " & ex.Message)
        End Try
    End Sub


    ' Bag-of-Words–Store und Index-Tracking
    Private store_bow As EmbeddingStore_BagofWords = New EmbeddingStore_BagofWords()
    Private indexedDocs_bow As HashSet(Of String) = New HashSet(Of String)()

    Public Sub RunIndexing_bow(refresh As Boolean, ChunkLength As Integer, ChunkOverlap As Integer)
        Dim doc = Application.ActiveDocument
        Dim docId = doc.FullName

        ' 0) Early return, wenn schon indexiert und kein Refresh gewünscht
        If indexedDocs_bow.Contains(docId) AndAlso Not refresh Then
            Return
        End If

        ' 1) Parameter validieren
        Dim nn As Integer = ChunkLength     ' Sätze pro Chunk
        Dim mm As Integer = ChunkOverlap     ' Überlappung
        Dim stepSize = nn - mm
        If nn <= 0 OrElse mm < 0 OrElse stepSize <= 0 Then
            Throw New System.ArgumentException("Bitte nn>0, mm≥0 und mm<nn wählen.")
        End If

        ' 2) Sätze holen und leere filtern
        Dim sentences = doc.Sentences.Cast(Of Word.Range)() _
                    .Where(Function(r) Not String.IsNullOrWhiteSpace(r.Text)) _
                    .ToList()
        Dim total = sentences.Count
        If total < nn Then
            Return
        End If

        ' 3) Chunks bauen (nur volle nn-Satz-Chunks)
        Dim chunks As New List(Of TextChunk)()
        For idx As Integer = 0 To total - nn Step stepSize
            Dim startIdx = idx
            Dim endIdx = idx + nn - 1  ' garantiert ≤ total-1

            ' Text zusammenbauen
            Dim parts = sentences.Skip(startIdx).Take(nn).Select(Function(r) r.Text.Trim())
            Dim chunkText = String.Join(" ", parts)
            If chunkText.Length < 10 Then Continue For

            ' Offsets aus den Ranges
            Dim rangeStart = sentences(startIdx).Start
            Dim startOffset = If(rangeStart < 0, 0, rangeStart)
            Dim rangeEnd = sentences(endIdx).End

            chunks.Add(New TextChunk With {
            .Text = chunkText,
            .StartOffset = startOffset,
            .EndOffset = rangeEnd
        })
        Next

        ' 4) Indexieren
        store_bow.IndexDocument(docId, chunks)
        If Not indexedDocs_bow.Contains(docId) Then indexedDocs_bow.Add(docId)
    End Sub

    Public Sub RunSearch_bow(SearchContext As String, DoNext As Boolean, DoRefresh As Boolean)
        Try

            ' 1) Parameter

            Dim ChunkLength As Integer = Default_Embed_Chunks_bow
            Dim ChunkOverlap As Integer = Default_Embed_Overlap_bow
            Dim Min_Score As Double = Default_Embed_Min_Score
            Dim Top_K As Integer = Default_Embed_Top_K
            Dim allDocs As Boolean = False
            Dim Fallback As Boolean = False

            If embed_store Is Nothing Or DoRefresh Then

                Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Sentences per chunk:", ChunkLength),
                    New SLib.InputParameter("Overlap per chunk", ChunkOverlap),
                    New SLib.InputParameter("Minimum relevance", Min_Score),
                    New SLib.InputParameter("Maximum hits", Top_K),
                    New SLib.InputParameter("Always hits", Fallback)
                    }

                If ShowCustomVariableInputForm("Please set your 'Bag of Words' and search values:", $"Context Search (Bag of Words)", params) Then

                    ChunkLength = CInt(params(0).Value)
                    ChunkOverlap = CInt(params(1).Value)
                    Min_Score = CDbl(params(2).Value)
                    Top_K = CInt(params(3).Value)
                    Fallback = CBool(params(4).Value)

                Else
                    Return
                End If

            Else
                Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Minimum relevance", Min_Score),
                    New SLib.InputParameter("Maximum hits:", Top_K),
                    New SLib.InputParameter("Always hits", Fallback),
                    New SLib.InputParameter("Search all indexed docs:", allDocs)
                    }
                If ShowCustomVariableInputForm("Please set your search values:", $"Context Search (Bag of Words)", params) Then

                    Min_Score = CDbl(params(0).Value)
                    Top_K = CBool(params(1).Value)
                    Fallback = CBool(params(2).Value)
                    allDocs = CBool(params(3).Value)

                Else
                    Return
                End If

            End If

            ' 2) Für Next-Suche: Selektion zurücksetzen & Cursor ans Ende
            Dim selRange As Word.Range = Application.Selection.Range
            If DoNext Then selRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

            ' 3) Index ggf. neu aufbauen
            RunIndexing_bow(DoRefresh, ChunkLength, ChunkOverlap)

            Dim currentDocId = Application.ActiveDocument.FullName
            Dim cursorPos = selRange.Start

            If Not DoNext Then
                ' --- COMPLETE: Suche im (Rest-)Dokument oder in allen Docs ---
                Dim rawHits = store_bow.Search(SearchContext, allDocs, True, currentDocId, cursorPos) _
                        .Where(Function(r) r.Score > 0 _
                                         AndAlso (allDocs OrElse r.StartOffset > cursorPos)) _
                        .OrderByDescending(Function(r) r.Score) _
                        .ToList()

                ' Treffer über Schwellwert
                Dim scoredHits = rawHits.Where(Function(r) r.Score >= Min_Score).ToList()
                Dim hits As List(Of SearchResult)
                If scoredHits.Count > 0 Then
                    hits = scoredHits.Take(Top_K).ToList()
                ElseIf Fallback Then
                    ' Fallback: die besten TOP_K unabhängig vom Score
                    hits = rawHits.Take(Top_K).ToList()
                End If

                If hits.Count = 0 Then
                    ShowCustomMessageBox($"No hits found for '{SearchContext}'" & If(Fallback, ".", " And minimum relevance of {Min_Score}."))
                    Return
                End If

                Dim trackChangesEnabled As Boolean = Application.ActiveDocument.TrackRevisions
                Application.ActiveDocument.TrackRevisions = True

                For Each r In hits
                    Dim docTarget = If(r.DocId = currentDocId,
                                   Application.ActiveDocument,
                                   Application.Documents.Open(r.DocId))
                    Dim rng = docTarget.Range(r.StartOffset, r.EndOffset)
                    docTarget.Comments.Add(rng, $"{AN5}: '{SearchContext}' (BoW score {r.Score:F3})")
                Next

                Application.ActiveDocument.TrackRevisions = trackChangesEnabled

                ShowCustomMessageBox($"{hits.Count} hits found for '{SearchContext}', a minimum relevance of {Min_Score} and a maximum of {Top_K} hits. Comments have been added to them.")
            Else
                ' --- NEXT: Suche nur ab Cursor im aktuellen Dokument ---
                Dim rawHits = store_bow.Search(SearchContext, False, True, currentDocId, cursorPos) _
                        .Where(Function(r) r.Score > 0 AndAlso r.StartOffset > cursorPos) _
                        .OrderByDescending(Function(r) r.Score) _
                        .ToList()

                Dim scoredHits = rawHits.Where(Function(r) r.Score >= Min_Score).ToList()
                Dim hits As List(Of SearchResult)
                If scoredHits.Count > 0 Then
                    hits = scoredHits.Take(Top_K).ToList()
                ElseIf Fallback Then
                    ' Fallback: die besten TOP_K unabhängig vom Score
                    hits = rawHits.Take(Top_K).ToList()
                End If

                If hits.Count = 0 Then
                    ShowCustomMessageBox($"No (further) hits found for '{SearchContext}'" & If(Fallback, ".", " And minimum relevance of {Min_Score}."))
                    Return
                End If

                Dim trackChangesEnabled As Boolean = Application.ActiveDocument.TrackRevisions
                Application.ActiveDocument.TrackRevisions = True

                For Each r In hits
                    Dim docTarget = If(r.DocId = currentDocId,
                                   Application.ActiveDocument,
                                   Application.Documents.Open(r.DocId))
                    Dim rng = docTarget.Range(r.StartOffset, r.EndOffset)
                    docTarget.Comments.Add(rng, $"{AN5}: '{SearchContext}' (BoW score {r.Score:F3})")
                Next

                Application.ActiveDocument.TrackRevisions = trackChangesEnabled

                ShowCustomMessageBox($"The (next) {hits.Count} have been found for '{SearchContext}', with a maximum of {Top_K} hits. Comments have been added to them.")
            End If
        Catch ex As System.Exception
            MessageBox.Show("Error in RunSearch_BoW: " & ex.Message)
        End Try
    End Sub





    ' Other helper functions
    Private Function GetSelectedTextLength() As Integer
        Try
            ' Get the active Word application
            Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application

            ' Get the current selection in the active document
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection

            ' Check if there is any selected text
            Dim selectedText As String = selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                Return 0
            End If

            ' Split the text on whitespace to count words,
            ' ignoring empty entries from multiple spaces/newlines
            Dim words = selectedText.Split(New Char() {" "c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf},
                                       StringSplitOptions.RemoveEmptyEntries)
            Return words.Length

        Catch ex As System.Exception ' Explicitly referencing System.Exception
            ' Handle any exceptions and return 0 if an error occurs
            Return 0
        End Try
    End Function
    Public Function InterpolateAtRuntime(ByVal template As String) As String
        If template Is Nothing Then
            MessageBox.Show("Error InterpolateAtRuntime: Template is Nothing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End If

        template = Regex.Replace(template, "{Codebasis}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack_2}", "", RegexOptions.IgnoreCase)

        Dim result As String = template

        Dim placeholderPattern As String = "\{([^}]+)\}"
        Dim matches As MatchCollection = Regex.Matches(template, placeholderPattern)

        For Each m As Match In matches
            Dim placeholder As String = m.Value          ' e.g. "{Name}"
            Dim varName As String = m.Groups(1).Value    ' e.g. "Name"

            ' Debug.WriteLine($"placeholder = {placeholder}  Varname = {varName}")
            ' Search for Field
            Dim fieldInfo = Me.GetType().GetField(varName)
            If fieldInfo IsNot Nothing Then
                Dim fieldValue = fieldInfo.GetValue(Me)
                If fieldValue IsNot Nothing Then
                    result = result.Replace(placeholder, fieldValue.ToString())
                End If
                Continue For
            End If

            ' Search for Property
            Dim propInfo = Me.GetType().GetProperty(varName)
            If propInfo IsNot Nothing Then
                Dim propValue = propInfo.GetValue(Me)
                If propValue IsNot Nothing Then
                    result = result.Replace(placeholder, propValue.ToString())
                End If
            End If
        Next

        Return result
    End Function


    Public Function VBAModuleWorking() As Boolean

        Dim xlApp As Microsoft.Office.Interop.Word.Application = Me.Application

        Try
            ' Call the VBA function
            Dim HelperVersion As Integer = CType(xlApp.Run("CheckAppHelper"), Integer)

            If HelperVersion >= MinHelperVersion Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
        End Try

    End Function


    Public Sub ShowSettings()

        If INILoadFail() Then Return
        Dim Settings As New Dictionary(Of String, String) From {
                {"Temperature", "Temperature of {model}"},
                {"Timeout", "Timeout of {model}"},
                {"Temperature_2", "Temperature of {model2}"},
                {"Timeout_2", "Timeout of {model2}"},
                {"DoubleS", "Convert '" & ChrW(223) & "' to 'ss'"},
                {"MarkdownConvert", "Keep character formatting"},
                {"KeepFormat1", "Keep format (translations)"},
                {"ReplaceText1", "Replace text (translations)"},
                {"KeepFormat2", "Keep format (other commands)"},
                {"ReplaceText2", "Replace text (other commands)"},
                {"KeepParaFormatInline", "Keep paragraph format"},
                {"KeepFormatCap", "Maximum text for keeping format (chars)"},
                {"DoMarkupWord", "Output as a markup (some functions)"},
                {"MarkupMethodHelper", "Markup method helpers (1 = Word, 2 = Diff, 3 = DiffW)"},
                {"MarkupMethodWord", "Markup method (1 = Word, 2 = Diff, 3 = DiffW, 4 = Regex)"},
                {"MarkupDiffCap", "Maximum characters for Diff Markup"},
                {"MarkupRegexCap", "Maximum characters for Regex Markup"},
                {"PreCorrection", "Additional instruction for prompts"},
                {"PostCorrection", "Prompt to apply after queries"},
                {"Language1", "Default translation language 1"},
                {"Language2", "Default translation language 2"},
                {"PromptLibPath", "Prompt library file"},
                {"PromptLibPath_Transcript", "Transcript prompt library file"},
                {"ShortcutsWordExcel", "Key shortcuts (for direct access)"},
                {"ChatCap", "Chat conversation memory (chars)"},
                {"SpeechModelPath", "Path to the speech recognition models"}
            }
        Dim SettingsTips As New Dictionary(Of String, String) From {
                {"Temperature", "The higher, the more creative the LLM will be (0.0-2.0)"},
                {"Timeout", "In milliseconds"},
                {"Temperature_2", "The higher, the more creative the LLM will be (0.0-2.0)"},
                {"Timeout_2", "In milliseconds"},
                {"DoubleS", "For Switzerland"},
                {"MarkdownConvert", "If selected, bold, italic, underline and some more formatting will be preserved converting it to Markdown coding before passing it to the LLM (most LLM support it)"},
                {"KeepFormat1", "If selected, the original's text basic character and paragraph formatting of a translated text will be retained (by HTML encoding, takes time!)"},
                {"ReplaceText1", "If selected, the response of the LLM for translations will replace the original text"},
                {"KeepFormat2", "If selected, the original's text basic character formatting will be retained for commands other than translations (by HTML encoding, takes time!)"},
                {"ReplaceText2", "If selected, the response of the LLM for other commands (than translate) will replace the original text"},
                {"KeepParaFormatInline", "If selected, the basic formatting of each paragraph will be retained by encoding it into the text (takes time, but less time encoding HTML), unless 'Keep Format' is selected"},
                {"KeepFormatCap", "If a text has more characters, then the format will not be retained (to prevent having to wait too long)"},
                {"DoMarkupWord", "Whether a markup should be done for functions that change only parts of a text"},
                {"MarkupMethodHelper", "Which markup method to use: 1 = Word compare, 2 = Simple Differ, 3 = Diff shown in a window"},
                {"MarkupMethodWord", "Which markup method to use: 1 = Word compare, 2 = Simple Differ, 3 = Diff shown in a window, 4 = LLM-based Regex Markup"},
                {"MarkupDiffCap", "The maximum size of the text that should be processed using the Diff method (to avoid you having to wait too long)"},
                {"MarkupRegexCap", "The maximum size of the text that should be processed using the Regex method (to avoid you having to wait too long)"},
                {"PreCorrection", "Add prompting text that will be added to all basic requests (e.g., for special language tasks)"},
                {"PostCorrection", "Add a prompt that will be applied to each result before it is further processed (slow!)"},
                {"Language1", "The language (in English) that will be used for the first quick access button in the ribbon"},
                {"Language2", "The language (in English) that will be used for the second quick access button in the ribbon"},
                {"PromptLibPath", "The filename (including path, support environmental variables) for your prompt library (if any)"},
                {"PromptLibPath_Transcript", "The filename (including path, support environmental variables) for your transcript prompt library (if any)"},
                {"ShortcutsWordExcel", "You can add key shortcuts by giving the name of the context menu, e.g., 'Correct=Ctrl-Shift-C', separated by ';' (only works if context menus are enabled and the Word helper is installed)"},
                {"ChatCap", "Use this to limit how many characters of your past chat discussion the chatbot will memorize (for saving costs and time)"},
                {"SpeechModelPath", "This is the path where you have to store the Vosk and Whisper models (and the Whisper.net runtime) for running the Transcriptor."}
            }

        ShowSettingsWindow(Settings, SettingsTips)

        Dim splash As New SplashScreen("Updating menu following your changes ...")
        splash.Show()
        splash.Refresh()

        AddContextMenu()

        splash.Close()

    End Sub

    Public Function GetWordDefaultInterfaceLanguage() As String
        Try
            ' Get the language ID of the Word user interface
            Dim uiLanguageID As Integer = Globals.ThisAddIn.Application.LanguageSettings.LanguageID(MsoAppLanguageID.msoLanguageIDUI)

            ' Convert the language ID to a human-readable name
            Dim cultureInfo As Globalization.CultureInfo = New Globalization.CultureInfo(uiLanguageID)

            ' Return the language display name
            Return cultureInfo.DisplayName
        Catch ex As System.Exception
            Return "English"
        End Try
    End Function

    Private Function CodeAPIKey(ByVal apiKey As String) As String
        Dim modifiedKey As String
        Dim resultKey As String
        Dim xcodebasis As String
        Dim HadPrefix As Boolean = False

        Dim PrefixValue As String = INI_APIKeyPrefix

        ' Check if an API key is provided
        apiKey = apiKey.Trim()
        If String.IsNullOrEmpty(apiKey) Then
            ShowCustomMessageBox("No text selected to encode. Select the API Key you wish to encode.")
            Return "Error"
        End If

        PrefixValue = SLib.ShowCustomInputBox("Please enter the API key prefix (as used in the configuration file, if any):", "API Key Encryptor", True, PrefixValue)

        xcodebasis = SLib.ShowCustomInputBox("Please enter the secret key:", "API Key Encryptor", True)
        If String.IsNullOrEmpty(xcodebasis) Then
            ShowCustomMessageBox("No secret key entered.")
            Return "Error"
        End If

        ' Check if the API key has the prefix
        If Not String.IsNullOrEmpty(PrefixValue) AndAlso apiKey.StartsWith(PrefixValue) Then
            HadPrefix = True
            ' Encrypt only the part after the prefix
            modifiedKey = apiKey.Substring(PrefixValue.Length)
        Else
            ' Encrypt the entire key if no prefix is present
            modifiedKey = apiKey
        End If

        ' Encrypt the modified key (without the prefix)
        resultKey = CodeString(modifiedKey, xcodebasis)

        ' Add the prefix back if it was present
        If HadPrefix Then
            resultKey = PrefixValue & resultKey
        End If

        Return resultKey
    End Function
    Private Function DeCodeAPIKey(ByVal apiKey As String) As String
        Dim modifiedKey As String
        Dim resultKey As String
        Dim xcodebasis As String

        Dim PrefixValue As String = INI_APIKeyPrefix

        ' Check if an API key is provided
        apiKey = apiKey.Trim()
        If String.IsNullOrEmpty(apiKey) Then
            ShowCustomMessageBox("No text selected to decode. Select the API Key you wish to decode.")
            Return "Error"
        End If

        PrefixValue = SLib.ShowCustomInputBox("Please enter the API key prefix (as used in the configuration file, if any):", "API Key Decryptor", True, PrefixValue)

        xcodebasis = SLib.ShowCustomInputBox("Please enter the secret key:", "API Key Decryptor", True)
        If String.IsNullOrEmpty(xcodebasis) Then
            ShowCustomMessageBox("No secret key entered.")
            Return "Error"
        End If

        ' Check if the key starts with the prefix
        If Not String.IsNullOrEmpty(PrefixValue) AndAlso apiKey.StartsWith(PrefixValue) Then
            ' Decrypt only the part after the prefix
            modifiedKey = apiKey.Substring(PrefixValue.Length)
        Else
            ' Decrypt the entire key if no prefix is present
            modifiedKey = apiKey
        End If

        ' Decrypt the modified key (without the prefix)
        resultKey = DecodeString(modifiedKey, xcodebasis)

        ' Add the prefix back only if it was in the original key
        If Not String.IsNullOrEmpty(PrefixValue) AndAlso apiKey.StartsWith(PrefixValue) Then
            resultKey = PrefixValue & resultKey
        End If

        Return resultKey
    End Function
    Public Function GetFileContent(Optional ByVal optionalFilePath As String = Nothing, Optional Silent As Boolean = False) As String
        Dim filePath As String = ""
        Try

            If optionalFilePath IsNot Nothing Then
                filePath = ExpandEnvironmentVariables(optionalFilePath)
            End If

            If String.IsNullOrWhiteSpace(filePath) Then
                Using form As New DragDropForm()
                    If form.ShowDialog() = DialogResult.OK Then
                        filePath = form.SelectedFilePath
                    Else
                        ' User cancelled or closed form
                        Return String.Empty
                    End If
                End Using
            End If

            filePath = RemoveCR(filePath.Trim())
            filePath = Path.GetFullPath(filePath)
            If Not File.Exists(filePath) Then
                If Not Silent Then ShowCustomMessageBox($"The file '{filePath}' was not found.")
                Return ""
            End If

            If Not String.IsNullOrWhiteSpace(filePath) AndAlso IO.File.Exists(filePath) Then
                Dim ext As String = IO.Path.GetExtension(filePath).ToLowerInvariant()
                Dim FromFile As String
                Select Case ext
                    Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm"
                        FromFile = ReadTextFile(filePath)
                    Case ".rtf"
                        FromFile = ReadRtfAsText(filePath)
                    Case ".doc", ".docx"
                        FromFile = ReadWordDocument(filePath)
                    Case ".pdf"
                        FromFile = ReadPdfAsText(filePath)
                    Case Else
                        FromFile = "Error: File type not supported."
                End Select
                If FromFile.StartsWith("Error") And Len(FromFile) < 100 And Not Silent Then
                    ShowCustomMessageBox(FromFile)
                    Return ""
                Else
                    Return FromFile
                End If
            End If
        Catch ex As System.Exception
            If Not Silent Then ShowCustomMessageBox($"An error occurred reading the file '{filePath}': {ex.Message}")
            Return ""
        End Try
    End Function

    Public Function GetFileName() As String
        Dim filePath As String = ""
        Try
            If String.IsNullOrWhiteSpace(filePath) Then
                Using form As New DragDropForm()
                    If form.ShowDialog() = DialogResult.OK Then
                        filePath = form.SelectedFilePath
                    Else
                        ' User cancelled or closed form
                        Return String.Empty
                    End If
                End Using
            End If

            filePath = RemoveCR(filePath.Trim())
            filePath = Path.GetFullPath(filePath)
            If Not File.Exists(filePath) Then
                ShowCustomMessageBox($"The file '{filePath}' was not found.")
                Return ""
            End If
            Return filePath

        Catch ex As System.Exception
            ShowCustomMessageBox($"An error occurred reading the file '{filePath}': {ex.Message}")
            Return ""
        End Try
    End Function

    ' Internet Helper Functions

    Public Async Function PerformSearchGrounding(SGTerms As String, ISearch_URL As String, ISearch_ResponseMask1 As String, ISearch_ResponseMask2 As String, ISearch_Tries As Integer, ISearch_MaxDepth As Integer) As Task(Of List(Of String))
        Dim results As New List(Of String)
        Using httpClient As New HttpClient()
            Try
                ' Construct the search URL
                Dim searchUrl As String = ISearch_URL & Uri.EscapeDataString(SGTerms)

                InfoBox.ShowInfoBox($"Searching {searchUrl} ...")

                ' Get search results HTML
                httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
                httpClient.Timeout = TimeSpan.FromSeconds(30) ' Set to an appropriate value
                Dim searchResponse As String = Await httpClient.GetStringAsync(searchUrl)
                'Debug.WriteLine("Search response: " & Left(searchResponse, 10000))

                InfoBox.ShowInfoBox($"Extracting URLs ...")

                ' Extract URLs using the defined start and mask
                Dim urlPattern As String = Regex.Escape(ISearch_ResponseMask1) & "(.*?)" & Regex.Escape(ISearch_ResponseMask2)
                Dim matches As MatchCollection = Regex.Matches(searchResponse, urlPattern)

                Dim extractedUrls As New List(Of String)
                Dim URLList As String = "URLS found so far:" & vbCrLf & vbCrLf
                For Each match As Match In matches
                    Dim rawUrl As String = match.Groups(1).Value
                    Dim decodedUrl As String = WebUtility.UrlDecode(rawUrl.Replace(ISearch_ResponseMask1, ""))

                    ' Check if the decoded URL already exists in the list
                    If Not extractedUrls.Contains(decodedUrl) Then
                        extractedUrls.Add(decodedUrl)
                        URLList += decodedUrl & vbCrLf
                        InfoBox.ShowInfoBox(URLList)
                        'Debug.WriteLine("URL added: " & decodedUrl)
                    Else
                        'Debug.WriteLine("Duplicate URL skipped: " & decodedUrl)
                    End If

                    If extractedUrls.Count >= ISearch_Tries Then Exit For
                Next

                ' Visit each extracted URL and retrieve content
                For Each url In extractedUrls
                    Try
                        Dim content As String = Await RetrieveWebsiteContent(url, ISearch_MaxDepth, httpClient)
                        'Debug.WriteLine("URL {url} provides:" & content)
                        If Not String.IsNullOrWhiteSpace(content) Then
                            If Len(content) > ISearch_MinChars Then
                                results.Add(content)
                                InfoBox.ShowInfoBox($"{url} resulted in: " & Left(content.Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, ""), 1000))
                                'Debug.WriteLine("Content=" & content)
                            Else
                                'Debug.WriteLine("Content (not considered)=" & content)
                            End If
                        End If
                    Catch ex As Exception
                        'Debug.WriteLine($"Error retrieving content from URL: {url} - {ex.Message}")
                    End Try
                Next

            Catch ex As HttpRequestException
                'Debug.WriteLine($"HTTP Request Error: {ex.Message}")
                ShowCustomMessageBox("An error occurred when searching and analyzing the Internet (HTTP request error: " & ex.Message & ")")
            Catch ex As TaskCanceledException
                'Debug.WriteLine("Request timed out or was canceled.")
                ShowCustomMessageBox("An error occurred when searching and analyzing the Internet (request timed-out or was canceled: " & ex.Message & ")")
            Catch ex As Exception
                'Debug.WriteLine($"An error occurred: {ex.Message}")
                ShowCustomMessageBox("An error occurred when searching and analyzing the Internet (" & ex.Message & ")")
            Finally
                httpClient.Dispose()
                InfoBox.ShowInfoBox("")
            End Try
        End Using
        Return results
    End Function

    Private Async Function RetrieveWebsiteContent(
                        baseUrl As String,
                        subTries As Integer,
                        httpClient As HttpClient
                    ) As Task(Of String)

        ' Create a single HttpClient for the entire crawl (optional if you already have one)
        Dim client As New HttpClient()

        ' Create the shared context object
        Dim context As New CrawlContext With {
                    .VisitedUrls = New HashSet(Of String)(),
                    .ContentBuilder = New StringBuilder(),
                    .ErrorCount = 0,
                    .MaxErrors = ISearch_MaxCrawlErrors  ' e.g. the user-defined max # of errors
                }

        ' Create one CancellationTokenSource for the entire crawl (30s in your example)
        Dim cts As New CancellationTokenSource(TimeSpan.FromSeconds(INI_ISearch_Timeout))

        ' Call the CrawlWebsite function with the context
        '   - 'subTries' is your maxDepth
        '   - '0' is your currentDepth
        '   - pass the same 'cts.Token' so the entire crawl times out in 30s

        Await CrawlWebsite(
                    currentUrl:=baseUrl,
                    maxDepth:=subTries,
                    currentDepth:=0,
                    httpClient:=client,
                    context:=context,
                    cancellationToken:=cts.Token,
                    timeOutSeconds:=INI_ISearch_Timeout
                     )

        ' Return plain text with HTML tags removed (up to ISearch_MaxChars)
        Return Left(
                    Regex.Replace(context.ContentBuilder.ToString(), "<.*?>", String.Empty).Trim(),
                    ISearch_MaxChars
                    )
    End Function

    Public Class CrawlContext
        Public Property VisitedUrls As HashSet(Of String)
        Public Property ContentBuilder As StringBuilder
        Public Property ErrorCount As Integer
        Public Property MaxErrors As Integer
    End Class


    Private Async Function CrawlWebsite(
    currentUrl As String,
    maxDepth As Integer,
    currentDepth As Integer,
    httpClient As HttpClient,
    context As CrawlContext,
    Optional cancellationToken As CancellationToken = Nothing,
    Optional timeOutSeconds As Integer = 10
) As Task(Of String)

        ' If the function has no valid CancellationToken, create one that cancels after 30 seconds
        Dim localCts As CancellationTokenSource = Nothing
        If cancellationToken = CancellationToken.None Then
            localCts = New CancellationTokenSource(TimeSpan.FromSeconds(timeOutSeconds))
            cancellationToken = localCts.Token
        End If

        Dim results As String = ""

        ' If we've already exceeded the max errors, abort quickly
        If context.ErrorCount >= context.MaxErrors Then
            Return results
        End If

        ' Early exit if depth is exceeded or already visited
        If currentDepth > maxDepth OrElse context.VisitedUrls.Contains(currentUrl) Then
            Return results
        End If

        Try
            context.VisitedUrls.Add(currentUrl)

            ' Use the cancellation token to abort if it exceeds the specified time
            Dim response As HttpResponseMessage = Await httpClient.GetAsync(currentUrl, cancellationToken)
            Dim pageHtml As String = Await response.Content.ReadAsStringAsync()

            Dim doc As New HtmlAgilityPack.HtmlDocument()
            doc.LoadHtml(pageHtml)

            ' Safely extract paragraph text
            Dim pNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//p")
            If pNodes IsNot Nothing Then
                For Each node In pNodes
                    context.ContentBuilder.AppendLine(node.InnerText.Trim())
                Next
            End If

            ' Follow links if depth permits
            If currentDepth < maxDepth Then
                Dim links As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//a[@href]")
                If links IsNot Nothing Then
                    For Each link In links
                        Dim hrefValue As String = link.GetAttributeValue("href", "").Trim()
                        Dim absoluteUrl As String = GetAbsoluteUrl(currentUrl, hrefValue)
                        ' You should already have a GetAbsoluteUrl function that resolves relative paths

                        If Not String.IsNullOrEmpty(absoluteUrl) Then
                            Await CrawlWebsite(
                            absoluteUrl,
                            maxDepth,
                            currentDepth + 1,
                            httpClient,
                            context,
                            cancellationToken,
                            timeOutSeconds
                        )

                            ' If error count has now exceeded the limit, stop immediately
                            If context.ErrorCount >= context.MaxErrors Then
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

        Catch ex As System.Threading.Tasks.TaskCanceledException
            ' Decide if a cancellation/timeout should increment errorCount
            context.ErrorCount += 1
            Debug.WriteLine($"Task canceled while crawling URL: {currentUrl} - {ex.Message}")

        Catch ex As System.Exception
            context.ErrorCount += 1
            Debug.WriteLine($"Error crawling URL: {currentUrl} - {ex.Message}")
        Finally
            If localCts IsNot Nothing Then
                localCts.Dispose()
            End If
        End Try

        Return results
    End Function


    Private Function GetAbsoluteUrl(baseUrl As String, relativeUrl As String) As String
        Try
            Dim baseUri As New Uri(baseUrl)
            Dim absoluteUri As New Uri(baseUri, relativeUrl)
            Return absoluteUri.ToString()
        Catch ex As Exception
            ' Invalid relative URL handling
            Return String.Empty
        End Try
    End Function

    ' WebExtension integration

    Private httpListener As HttpListener
    Private listenerThread As Thread
    Private isShuttingDown As Boolean = False


    Private Sub StartupHttpListener()
        ' Start the HTTP listener on a background thread.
        listenerThread = New Thread(AddressOf StartHttpListener)
        listenerThread.IsBackground = True
        listenerThread.Start()
    End Sub


    Private Sub ShutdownHttpListener()
        ' Cleanly stop the listener if it's running.
        isShuttingDown = True
        If httpListener IsNot Nothing AndAlso httpListener.IsListening Then
            httpListener.Stop()
            httpListener.Close()
        End If
    End Sub

    Private Async Function StartHttpListener() As Task(Of String)
        Dim prefix As String = "http://127.0.0.1:12334/"
        Dim consecutiveFailures As Integer = 0

        Try
            ' Initialize the listener once.
            If httpListener Is Nothing Then
                httpListener = New HttpListener()
                httpListener.Prefixes.Add(prefix)
                httpListener.Start()
                Debug.WriteLine("HttpListener started.")
            End If

            While Not isShuttingDown
                Dim delayNeeded As Boolean = False

                ' If for some reason the listener is not active, restart it.
                If httpListener Is Nothing OrElse Not httpListener.IsListening Then
                    Try
                        If httpListener IsNot Nothing Then
                            httpListener.Close()
                        End If
                    Catch ex As System.Exception
                        Debug.WriteLine("Error closing HttpListener: " & ex.Message)
                    End Try

                    httpListener = New HttpListener()
                    httpListener.Prefixes.Add(prefix)
                    httpListener.Start()
                    Debug.WriteLine("HttpListener restarted.")
                End If

                Try
                    ' Asynchronously wait for an incoming request.
                    Dim context As HttpListenerContext = Await httpListener.GetContextAsync()
                    Dim result As String = Await HandleHttpRequest(context)
                    Debug.WriteLine("Request handled successfully.")
                    ' Reset the failure counter on success.
                    consecutiveFailures = 0
                Catch ex As System.ObjectDisposedException
                    Debug.WriteLine("HttpListener was disposed. Restarting listener...")
                    consecutiveFailures += 1
                    delayNeeded = True
                Catch ex As System.Exception
                    Debug.WriteLine("Error handling HTTP request: " & ex.Message)
                    consecutiveFailures += 1
                    delayNeeded = True
                End Try

                ' Check if we have reached the maximum number of consecutive failures.
                If consecutiveFailures >= 10 Then
                    Debug.WriteLine("Too many consecutive failures. Shutting down.")
                    isShuttingDown = True
                    Exit While
                End If

                ' If an error occurred, delay before restarting.
                If delayNeeded Then
                    Await System.Threading.Tasks.Task.Delay(5000)
                End If
            End While
        Catch ex As System.Exception
            Debug.WriteLine("Error in StartHttpListener: " & ex.Message)
        End Try

        Return ""
    End Function


    'Private Sub HandleHttpRequest(ByVal context As HttpListenerContext)
    Private Async Function HandleHttpRequest(ByVal context As HttpListenerContext) As Task(Of String)
        Try
            ' 1) Retrieve the request
            Dim request As HttpListenerRequest = context.Request
            Dim response As HttpListenerResponse = context.Response

            ' Debug logs for incoming request details
            Debug.Print("Raw URL: " & request.RawUrl)
            Debug.Print("HTTP Method: " & request.HttpMethod)
            Debug.Print("Content-Length: " & request.ContentLength64)
            Debug.Print("Content-Type: " & request.ContentType)
            Debug.Print("Has Entity Body: " & request.HasEntityBody.ToString())
            ' --- full requested URI ---
            Debug.Print("Full URL: " & request.Url.AbsoluteUri)
            ' --- referring page (if the browser sent a Referer header) ---
            If request.UrlReferrer IsNot Nothing Then
                Debug.Print("Referrer: " & request.UrlReferrer.AbsoluteUri)
            Else
                Debug.Print("Referrer: (none)")
            End If

            ' Handle preflight (OPTIONS) request
            If request.HttpMethod = "OPTIONS" Then
                Debug.Print("Handling preflight (OPTIONS) request...")
                response.AddHeader("Access-Control-Allow-Origin", "*")
                response.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization")
                response.StatusCode = 204 ' No Content
                'response.OutputStream.Close()
                response.Close()
                Return ""
            End If

            ' Initialize request body variable
            Dim requestBody As String = String.Empty

            ' Handle entity body (for POST/PUT, etc.)
            If request.HasEntityBody Then
                Debug.Print("Processing request body...")
                Using reader As New StreamReader(request.InputStream, Encoding.UTF8)
                    requestBody = reader.ReadToEnd()
                End Using
                Debug.Print("Request Body: " & requestBody)
            End If

            ' 2) Process the request
            '    - Parse JSON or handle requestBody
            Dim responseText As String = ProcessRequestInAddIn(requestBody, request.RawUrl)

            ' 3) Write a response with CORS headers
            Dim buffer As Byte() = Encoding.UTF8.GetBytes(responseText)
            response.ContentLength64 = buffer.Length
            response.ContentType = "text/plain; charset=utf-8"
            response.AddHeader("Access-Control-Allow-Origin", "*") ' Allow cross-origin requests

            Using output As Stream = response.OutputStream
                output.Write(buffer, 0, buffer.Length)
            End Using
            context.Response.Close()
            Debug.WriteLine("HTTP Request completed without errors.")
            Return ""
        Catch ex As System.Exception
            ' If there's an error, return an error response to the caller
            Try
                Dim errorStr = "Error: " & ex.Message
                Dim errorBytes = Encoding.UTF8.GetBytes(errorStr)
                context.Response.ContentLength64 = errorBytes.Length
                context.Response.StatusCode = 500  ' Internal server error
                context.Response.OutputStream.Write(errorBytes, 0, errorBytes.Length)
                'context.Response.OutputStream.Close()
                context.Response.Close()
            Catch
                ' If we can’t even write an error, just ignore
            End Try
            Debug.WriteLine("HTTP Request completed with errors.")
            Return ""
        End Try
    End Function


    Private Function ProcessRequestInAddIn(requestBody As String, rawUrl As String) As String

        Dim result As String = ""

        Try
            ' Parse the JSON string
            Dim jsonObject As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(requestBody)

            Debug.WriteLine("Requestbody = " & requestBody)

            ' Check if the "command" segment contains "redink_sendtoword"
            Dim command As String = jsonObject("Command")?.ToString()
            If command IsNot Nothing AndAlso command.Equals("redink_sendtoword", StringComparison.OrdinalIgnoreCase) Then
                ' Extract the "text" segment
                Dim textToInsert As String = jsonObject("Text")?.ToString()
                Dim SourceURL As String = jsonObject("URL")?.ToString()
                If textToInsert IsNot Nothing Then
                    ' Get the active Word document and the selection
                    Dim app As Word.Application = Globals.ThisAddIn.Application
                    Dim selection As Word.Selection = app.Selection

                    ' Insert the text at the current cursor position
                    selection.TypeText(textToInsert & " (" & SourceURL & ")")

                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"Error in ProcessRequestInAddIn: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return result

    End Function

    Public Class TranscriptionForm

        Inherits Form

        Private STT_PreExistingSleepBlocker As Boolean = False

        Private RichTextBox1 As Forms.RichTextBox
        Private StartButton As Forms.Button
        Private StopButton As Forms.Button
        Private ClearButton As Forms.Button
        Private LoadButton As Forms.Button
        Private AudioButton As Forms.Button
        Private QuitButton As Forms.Button
        Private ProcessButton As Forms.Button
        Private cultureComboBox As Forms.ComboBox
        Private deviceComboBox As Forms.ComboBox
        Private processCombobox As Forms.ComboBox
        Private SpeakerIdent As System.Windows.Forms.CheckBox
        Private SpeakerDistance As Forms.TextBox
        Private Label1 As Label
        Private Label2 As Label
        Private StatusLabel As Label
        Private PartialTextLabel As Label
        Private ButtonPanel As Panel

        Private TranscriptPromptsTitles As New List(Of String)
        Private TranscriptPromptsLibrary As New List(Of String)

        Private recognizer As VoskRecognizer
        Private waveIn As WaveInEvent
        Private capturing As Boolean = False
        Private partialText As String = ""
        Private finalText As New StringBuilder()
        Private Const VoskTooltip = "Only for Vosk: Set similarity threshold for speaker identification (0.5-0.7 for real-time speaker tracking, 1.0-1.5 for meetings/interviews)"
        Private Const VoskToggle = "Iden"

        Private WhisperRecognizer As WhisperProcessor
        Private audioBuffer As New List(Of Single)
        Private STTCanceled As Boolean = False
        Private cts As CancellationTokenSource = New CancellationTokenSource()
        Private Const WhisperTooltip = "Only for Whisper: Select if text shall be translated to English and the threshold for detecting voice (default = 0.6, increase for noisy environments)"
        Private Const WhisperToggle = "Trans"
        Private STTModel As String = "whisper"

        Private GoogleSpeech As Boolean = False
        Private STTSecondAPI As Boolean = False
        Private IsGoogle As Boolean = False
        Private Const GoogleTooltip = "Only for Google: Set the maximum number of speakers expected for diarization (speaker tracking)"
        Private Const GoogleToggle = "Iden"
        Private googleReaderTask As System.Threading.Tasks.Task
        Private readerCts = New CancellationTokenSource()
        Private _stream As SpeechClient.StreamingRecognizeStream
        Private googleTranscriptStart As Integer = 0
        Private client As SpeechClient
        Private GoogleLanguageCode As String = ""
        Private audioQueue As New System.Collections.Concurrent.BlockingCollection(Of ByteString)()
        Private _googleStreamCompleted As Boolean = False
        Private Const STREAMING_LIMIT_MS As Integer = 290000  ' 4 Minuten 50 Sekunden
        Private streamingStartTime As DateTime

        Private ReadOnly ringBuffer As New Queue(Of Google.Protobuf.ByteString)()
        Private Const RING_BUFFER_SIZE As Integer = 50

        Private ReadOnly recoverySemaphore As New System.Threading.SemaphoreSlim(1, 1)
        Private writerTask As System.Threading.Tasks.Task

        ' The watchdog timer that will check for API responsiveness.
        Private _apiWatchdogTimer As System.Threading.Timer
        ' Tracks the last time we received ANY response (partial or final) from Google.
        ' We use a long representing Ticks for thread-safe updates.
        Private _lastApiResponseTicks As Long
        ' Configurable: The number of seconds of API silence before triggering a restart.
        Private Const API_RESPONSE_TIMEOUT_SECONDS As Integer = 3
        Private _lastKnownPartialResult As String = ""
        Private _justCommittedPartialText As String = ""

        ' Maps a temporary SpeakerTag (e.g., 1, 2) from the API to a consistent,
        ' human-readable label (e.g., "Speaker 1", "Speaker 2").
        Private _speakerTagToLabelMap As New Dictionary(Of Integer, String)

        ' Counter to ensure we always assign a new, unique speaker number.
        Private _nextSpeakerNumber As Integer = 1

        Private loopback As WasapiLoopbackCapture
        Private loopbackBuffer As BufferedWaveProvider
        Private loopbackCapture As WasapiLoopbackCapture
        Private loopbackRawProvider As BufferedWaveProvider
        Private loopbackResampler As MediaFoundationResampler
        Private _multiSourceSelected As Boolean = False

        Private ReadOnly Property MultiSourceEnabled As Boolean
            Get
                Return _multiSourceSelected
            End Get
        End Property

        Private sttAccessToken1 As String = String.Empty
        Private sttTokenExpiry1 As DateTime = DateTime.MinValue
        Private sttAccessToken2 As String = String.Empty
        Private sttTokenExpiry2 As DateTime = DateTime.MinValue



        ' Hilfs‐Methode: PrivateKey in 64-Zeichen-Zeilen brechen
        Public Shared Function FormatPrivateKey(rawKey As String) As String
            Dim noEscapes = rawKey.Replace("\n", "")
            Dim sb As New System.Text.StringBuilder()
            For i As Integer = 0 To noEscapes.Length - 1 Step 64
                Dim chunk = If(i + 64 <= noEscapes.Length,
                      noEscapes.Substring(i, 64),
                      noEscapes.Substring(i))
                sb.AppendLine(chunk)
            Next
            Return "-----BEGIN PRIVATE KEY-----" & vbLf &
           sb.ToString() &
           "-----END PRIVATE KEY-----" & vbLf
        End Function

        ' Neu: Holt lokal einen frischen STT-Token für die gewählte API
        Private Async Function GetFreshSTTToken(useSecond As Boolean) As System.Threading.Tasks.Task(Of String)

            Try
                Dim token As String
                Dim expiry As DateTime

                If useSecond Then
                    token = sttAccessToken2
                    expiry = sttTokenExpiry2
                Else
                    token = sttAccessToken1
                    expiry = sttTokenExpiry1
                End If

                If String.IsNullOrEmpty(token) OrElse DateTime.UtcNow >= expiry Then
                    ' Parameter je nach API auswählen
                    Dim clientEmail = If(useSecond, INI_OAuth2ClientMail_2, INI_OAuth2ClientMail)
                    Dim scopes = If(useSecond, INI_OAuth2Scopes_2, INI_OAuth2Scopes)
                    Dim rawKey = If(useSecond, INI_APIKey_2, INI_APIKey)
                    Dim authServer = If(useSecond, INI_OAuth2Endpoint_2, INI_OAuth2Endpoint)
                    Dim life = If(useSecond, INI_OAuth2ATExpiry_2, INI_OAuth2ATExpiry)

                    ' GoogleOAuthHelper konfigurieren
                    GoogleOAuthHelper.client_email = clientEmail
                    GoogleOAuthHelper.private_key = FormatPrivateKey(rawKey)
                    GoogleOAuthHelper.scopes = scopes
                    GoogleOAuthHelper.token_uri = authServer
                    GoogleOAuthHelper.token_life = life

                    ' neuen Token holen
                    Dim newToken As String = Await GoogleOAuthHelper.GetAccessToken()
                    Dim newExpiry As DateTime = DateTime.UtcNow.AddSeconds(life - 300)

                    If useSecond Then
                        sttAccessToken2 = newToken
                        sttTokenExpiry2 = newExpiry
                    Else
                        sttAccessToken1 = newToken
                        sttTokenExpiry1 = newExpiry
                    End If

                    token = newToken
                End If

                Return token

            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(
            $"Error fetching STT token: {ex.Message}",
            "Transcription Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
                Return String.Empty
            End Try
        End Function

        Public Sub New()
            ' Initialize UI Components
            InitializeComponents()

            Me.AutoScaleMode = AutoScaleMode.Dpi
            ''Me.AutoScaleMode = AutoScaleMode.Font

            ' Load available Vosk models
            Dim modelPath As String = Globals.ThisAddIn.INI_SpeechModelPath
            Dim modelsexist As Boolean = False

            Dim Endpoint As String = INI_Endpoint
            Dim Endpoint_2 As String = INI_Endpoint_2

            If Endpoint.Contains(GoogleIdentifier) And INI_OAuth2 Then
                STTSecondAPI = False
                IsGoogle = True
            ElseIf Endpoint_2.Contains(GoogleIdentifier) And INI_OAuth2_2 Then
                STTSecondAPI = True
                IsGoogle = True
            End If
            If IsGoogle And Not String.IsNullOrWhiteSpace(STTEndpoint) Then
                GoogleSpeech = True
                cultureComboBox.Items.Add(GoogleSTT_Desc)
                modelsexist = True
            End If

            If Directory.Exists(modelPath) Then
                For Each dir As String In Directory.GetDirectories(modelPath)
                    Dim dirName As String = Path.GetFileName(dir)
                    If dirName.StartsWith("vosk-model") Then
                        cultureComboBox.Items.Add(dirName)
                        modelsexist = True
                    End If
                Next

                For Each file As String In Directory.GetFiles(modelPath)
                    Dim fileName As String = Path.GetFileName(file)
                    If fileName.StartsWith("ggml") Then
                        cultureComboBox.Items.Add(fileName)
                        modelsexist = True
                    End If
                Next

            End If

            ' Pre-select the last used model if it exists in the list
            Dim lastModel As String = My.Settings.LastSpeechModel
            If Not String.IsNullOrEmpty(lastModel) AndAlso cultureComboBox.Items.Contains(lastModel) Then
                cultureComboBox.SelectedItem = lastModel
            End If

            AddHandler Me.cultureComboBox.MouseMove, AddressOf cultureComboBox_MouseMove

            LoadAudioDevices()

            AddHandler Me.deviceComboBox.MouseMove, AddressOf deviceComboBox_MouseMove

            AddHandler Me.deviceComboBox.SelectedIndexChanged, AddressOf Me.deviceComboBox_SelectedIndexChanged

            LoadAndPopulateProcessComboBox(Globals.ThisAddIn.INI_PromptLibPath_Transcript, processCombobox)

            Dim index As Integer = Me.cultureComboBox.SelectedIndex
            If index >= 0 Then
                If Me.cultureComboBox.Items(index).startswith(GoogleSTT_Desc) Then
                    Me.SpeakerIdent.Text = GoogleToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, GoogleTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, GoogleTooltip)

                ElseIf Me.cultureComboBox.Items(index).startswith("ggml") Then
                    Me.SpeakerIdent.Text = WhisperToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, WhisperTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, WhisperTooltip)
                Else
                    Me.SpeakerIdent.Text = VoskToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, VoskTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, VoskTooltip)
                End If
            End If

            ' Wire up event handlers
            AddHandler StartButton.Click, AddressOf StartButton_Click
            AddHandler StopButton.Click, AddressOf StopButton_Click
            AddHandler ClearButton.Click, AddressOf ClearButton_Click
            AddHandler LoadButton.Click, AddressOf LoadButton_Click
            AddHandler AudioButton.Click, AddressOf AudioButton_Click
            AddHandler QuitButton.Click, AddressOf QuitButton_Click
            AddHandler ProcessButton.Click, AddressOf ProcessButton_Click

            ' Make window resizable
            Me.MinimumSize = New Size(800, 440)

            If Not modelsexist Then
                ShowCustomMessageBox($"No Vosk or Whisper models have been found at the configured path ('{modelPath}'). A model is necessary for transcribing. You can download models for free at {VoskSource} and {WhisperSource}.", $"{AN} Transcriptor")
                Me.Close()
            End If
        End Sub

        Private ToolTip As New Forms.ToolTip()

        Private Sub deviceComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
            ' runs on UI thread—safe to read SelectedItem here
            Dim s As String = TryCast(Me.deviceComboBox.SelectedItem, String)
            _multiSourceSelected = Not String.IsNullOrEmpty(s) _
                            AndAlso s.EndsWith("(plus audio output)")
        End Sub

        Private Sub cultureComboBox_MouseMove(sender As Object, e As MouseEventArgs)
            Dim index As Integer = Me.cultureComboBox.SelectedIndex
            If index >= 0 Then
                ToolTip.SetToolTip(Me.cultureComboBox, Me.cultureComboBox.Items(index).ToString())
                If Me.cultureComboBox.Items(index).startswith(GoogleSTT_Desc) Then
                    Me.SpeakerIdent.Text = GoogleToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, GoogleTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, GoogleTooltip)
                ElseIf Me.cultureComboBox.Items(index).startswith("ggml") Then
                    Me.SpeakerIdent.Text = WhisperToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, WhisperTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, WhisperTooltip)
                Else
                    Me.SpeakerIdent.Text = VoskToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, VoskTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, VoskTooltip)
                End If
            End If
        End Sub

        Private Sub deviceComboBox_MouseMove(sender As Object, e As MouseEventArgs)
            Dim index As Integer = Me.deviceComboBox.SelectedIndex
            If index >= 0 Then
                ToolTip.SetToolTip(Me.deviceComboBox, Me.deviceComboBox.Items(index).ToString())
            End If
        End Sub



        Public Sub ConfigureAudioOutputDevice()
            ' 1) Alle aktiven Render-Endpoints ermitteln
            Dim enumerator As New MMDeviceEnumerator()
            Dim devices As MMDeviceCollection =
        enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)

            ' 2) FriendlyNames und zugehörige IDs in parallele Arrays packen, inkl. Default als Index 0
            Dim totalCount As Integer = devices.Count + 1
            Dim deviceNames(totalCount - 1) As String
            Dim deviceIds(totalCount - 1) As String

            ' 2a) Default Audio Output Device (wie von WasapiLoopbackCapture)
            deviceNames(0) = "Default Audio Output Device"
            deviceIds(0) = String.Empty

            ' 2b) Alle anderen Geräte ab Index 1
            For i As Integer = 0 To devices.Count - 1
                deviceNames(i + 1) = devices(i).FriendlyName
                deviceIds(i + 1) = devices(i).ID
            Next

            ' 3) Aktuell in den Settings gespeichertes Device ermitteln (leere ID → Default)
            Dim currentDeviceId As String = My.Settings.AudioOutputDevice
            Dim currentDeviceName As String = String.Empty
            Dim idxSaved As Integer = Array.IndexOf(deviceIds, currentDeviceId)
            If idxSaved >= 0 Then
                currentDeviceName = deviceNames(idxSaved)
            End If

            ' 4) Prompt für den Auswahl-Dialog zusammenbauen
            Dim prompt As String = "Choose the audio output device for capturing"
            If Not String.IsNullOrEmpty(currentDeviceName) Then
                prompt &= $" (currently: {currentDeviceName})"
            End If
            prompt &= ":"

            ' 5) Auswahl-Dialog anzeigen
            Dim selection As String = ShowSelectionForm(
        prompt,
        $"{AN} Transcriptor",
        deviceNames)

            ' 6) Wenn Auswahl gültig, Index ermitteln und Settings setzen/clearen
            If Not String.IsNullOrEmpty(selection) AndAlso selection <> "esc" Then
                Dim chosenIndex As Integer = Array.IndexOf(deviceNames, selection)
                If chosenIndex >= 0 Then
                    If chosenIndex = 0 Then
                        ' Default gewählt → Setting leeren
                        My.Settings.AudioOutputDevice = String.Empty
                    Else
                        ' Konkrete Device-ID speichern
                        My.Settings.AudioOutputDevice = deviceIds(chosenIndex)
                    End If

                    Try
                        My.Settings.Save()
                    Catch ex As System.Exception
                        ' Volle Referenz auf Exception
                        ShowCustomMessageBox($"Error saving audio output device setting: {ex.Message}")
                    End Try
                End If
            End If
        End Sub





        Private Sub InitializeComponents()
            ' --- DPI‐aware form setup ---
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.AutoScaleMode = AutoScaleMode.Font
            Me.Text = $"{AN} Transcriptor (editable text, audio will not be stored)"
            Me.FormBorderStyle = FormBorderStyle.Sizable

            ' --- Create controls ---

            ' Transcript area
            Me.RichTextBox1 = New RichTextBox() With {
        .Font = New System.Drawing.Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point),
        .Multiline = True,
        .ScrollBars = RichTextBoxScrollBars.Vertical,
        .Dock = DockStyle.Fill
    }

            ' Selector labels
            Me.Label1 = New Label() With {.Text = "Model:", .AutoSize = True}
            Me.Label2 = New Label() With {.Text = "Source:", .AutoSize = True}

            ' Model / source dropdowns (start 50px wider)
            Me.cultureComboBox = New System.Windows.Forms.ComboBox() With {
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Width = 250
    }
            Me.deviceComboBox = New System.Windows.Forms.ComboBox() With {
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Width = 450
    }

            ' Speaker toggle + threshold
            Me.SpeakerIdent = New System.Windows.Forms.CheckBox() With {.Text = VoskToggle, .AutoSize = True}
            Me.SpeakerDistance = New System.Windows.Forms.TextBox() With {
        .Text = If(My.Settings.LastSpeakerDistance <= 0, "1.0", My.Settings.LastSpeakerDistance.ToString()),
        .Width = 50,
        .AutoSize = False
    }

            ' Status + partial text
            Me.StatusLabel = New Label() With {
        .Text = "Transcribing:",
        .AutoSize = True,
        .Dock = DockStyle.Top
    }
            Me.PartialTextLabel = New Label() With {
        .Text = "...",
        .AutoSize = True,
        .MinimumSize = New System.Drawing.Size(0, 70),
        .Dock = DockStyle.Top
    }

            ' Action buttons + bottom combobox
            Me.StartButton = New System.Windows.Forms.Button() With {.Text = "Start", .AutoSize = True}
            Me.StopButton = New System.Windows.Forms.Button() With {.Text = "Stop", .AutoSize = True, .Enabled = False}
            Me.ClearButton = New System.Windows.Forms.Button() With {.Text = "Clear", .AutoSize = True}
            Me.LoadButton = New System.Windows.Forms.Button() With {.Text = "Load", .AutoSize = True}
            Me.AudioButton = New System.Windows.Forms.Button() With {.Text = "Dev", .AutoSize = True}
            Me.QuitButton = New System.Windows.Forms.Button() With {.Text = "Quit", .AutoSize = True}
            Me.ProcessButton = New System.Windows.Forms.Button() With {.Text = "Process:", .AutoSize = True}
            Me.processCombobox = New System.Windows.Forms.ComboBox() With {
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Width = 250
    }

            ' Add a little right‐margin so controls aren’t jammed
            Dim pad As New Padding(0, 0, 10, 0)
            For Each ctl In {Label1, cultureComboBox, Label2, deviceComboBox, SpeakerIdent, SpeakerDistance,
                     StartButton, StopButton, ClearButton, LoadButton, AudioButton, QuitButton, ProcessButton}
                ctl.Margin = pad
            Next
            processCombobox.Margin = pad

            ' --- Build layout ---

            ' Root: 3 rows—top selectors, middle transcript, bottom actions
            Dim root As New TableLayoutPanel() With {
        .Dock = DockStyle.Fill,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .ColumnCount = 1,
        .RowCount = 3,
        .Padding = New Padding(10)
    }
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))    ' row0: selectors
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100)) ' row1: transcript
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))    ' row2: actions

            ' Row 0: selectors laid out in a TableLayoutPanel so combos stretch
            Dim topRow As New TableLayoutPanel() With {
        .Dock = DockStyle.Top,
        .AutoSize = False,
        .Height = cultureComboBox.PreferredHeight + 10,
        .ColumnCount = 6,
        .RowCount = 1,
        .Padding = New Padding(0, 0, 0, 10)
    }
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            cultureComboBox.Dock = DockStyle.Fill
            deviceComboBox.Dock = DockStyle.Fill

            topRow.Controls.Add(Label1, 0, 0)
            topRow.Controls.Add(cultureComboBox, 1, 0)
            topRow.Controls.Add(Label2, 2, 0)
            topRow.Controls.Add(deviceComboBox, 3, 0)
            topRow.Controls.Add(SpeakerIdent, 4, 0)
            topRow.Controls.Add(SpeakerDistance, 5, 0)

            root.Controls.Add(topRow, 0, 0)

            ' Row 1: status, partial, then main RichTextBox
            Dim mid As New TableLayoutPanel() With {
        .Dock = DockStyle.Fill,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .ColumnCount = 1,
        .RowCount = 3
    }
            mid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            mid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mid.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            mid.Controls.Add(StatusLabel, 0, 0)
            mid.Controls.Add(PartialTextLabel, 0, 1)
            mid.Controls.Add(RichTextBox1, 0, 2)

            root.Controls.Add(mid, 0, 1)

            ' Row 2: bottom actions in a stretchy TableLayoutPanel
            Dim bottomRow As New TableLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .AutoSize = False,
        .Height = StartButton.PreferredSize.Height + 20,
        .ColumnCount = 8,
        .RowCount = 1,
        .Padding = New Padding(0, 10, 0, 0)
    }
            ' first six columns auto‐size, last column (processCombobox) fills
            For i = 1 To 7
                bottomRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            Next
            bottomRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            processCombobox.Dock = DockStyle.Fill

            bottomRow.Controls.Add(StartButton, 0, 0)
            bottomRow.Controls.Add(StopButton, 1, 0)
            bottomRow.Controls.Add(ClearButton, 2, 0)
            bottomRow.Controls.Add(LoadButton, 3, 0)
            bottomRow.Controls.Add(AudioButton, 4, 0)
            bottomRow.Controls.Add(QuitButton, 5, 0)
            bottomRow.Controls.Add(ProcessButton, 6, 0)
            bottomRow.Controls.Add(processCombobox, 7, 0)

            root.Controls.Add(bottomRow, 0, 2)

            ' Swap in our root layout
            Me.Controls.Clear()
            Me.Controls.Add(root)

            ' Freeze minimum size once first shown
            AddHandler Me.Shown, Sub() Me.MinimumSize = Me.Size

            ' Set icon
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        End Sub


        Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
            Dim minWidth As Integer = SpeakerDistance.Left + SpeakerDistance.Width + 40
            If Me.Width < minWidth Then
                Me.Width = minWidth ' Force minimum width dynamically
            End If
        End Sub

        Private Async Function StopRecording() As System.Threading.Tasks.Task

            If loopbackCapture IsNot Nothing Then
                RemoveHandler loopbackCapture.DataAvailable, AddressOf OnLoopbackDataAvailable
                loopbackCapture.StopRecording()
                loopbackCapture.Dispose()
                loopbackCapture = Nothing
            End If

            If loopbackResampler IsNot Nothing Then
                loopbackResampler.Dispose()
                loopbackResampler = Nothing
                loopbackRawProvider = Nothing
            End If


            If waveIn IsNot Nothing Then
                RemoveHandler waveIn.DataAvailable, AddressOf OnGoogleDataAvailable
                RemoveHandler waveIn.DataAvailable, AddressOf OnAudioDataAvailable
                waveIn.StopRecording()
                waveIn.Dispose()
                waveIn = Nothing
            End If

            CancelTranscription()

            If STTModel = "google" AndAlso _stream IsNot Nothing Then
                Await SafeCompleteAndDisposeGoogleStreamAsync(readerCts.Token)
            End If

            If WhisperRecognizer IsNot Nothing Then
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper stopped...")
                Await WhisperRecognizer.DisposeAsync()
                WhisperRecognizer = Nothing
            End If

            If IsNothing(prevExecState) Then
                SetThreadExecutionState(ES_CONTINUOUS)
            Else
                If Not STT_PreExistingSleepBlocker Then
                    SetThreadExecutionState(prevExecState)
                    prevExecState = Nothing
                End If
            End If

        End Function


        Private Sub StopButton_Click(sender As Object, e As EventArgs)

            If Not capturing Then Return

            STTCanceled = True

            ' Verhindere Mehrfachklicks
            Me.StopButton.Enabled = False
            If STTModel <> "vosk" Then
                PartialTextLabel.Text = "Stopping…"
            End If

            System.Threading.Tasks.Task.Run(Async Function()
                                                Try
                                                    Await StopRecording()
                                                    If STTModel = "google" Then StopApiWatchdogTimer()
                                                Catch ex As System.Exception

                                                End Try

                                                Me.Invoke(Sub()
                                                              Me.StartButton.Enabled = True
                                                              Me.LoadButton.Enabled = True
                                                              Me.AudioButton.Enabled = True
                                                              Me.cultureComboBox.Enabled = True
                                                              Me.deviceComboBox.Enabled = True
                                                              Me.SpeakerIdent.Enabled = True
                                                              Me.SpeakerDistance.Enabled = True

                                                              If STTModel = "vosk" Then
                                                                  Addline(PartialTextLabel.Text)
                                                              End If
                                                              PartialTextLabel.Text = String.Empty
                                                          End Sub)
                                            End Function)

            capturing = False

        End Sub


        Private Sub ClearButton_Click(sender As Object, e As EventArgs)
            RichTextBox1.Invoke(Sub()
                                    RichTextBox1.Text = ""
                                    RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                    RichTextBox1.ScrollToCaret()
                                End Sub)
        End Sub

        Private Sub FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.Closing

            If e.CloseReason = CloseReason.UserClosing Then
                If capturing Then

                    STTCanceled = True

                    Me.StopButton.Enabled = False
                    Me.AudioButton.Enabled = False
                    Me.QuitButton.Enabled = False
                    If STTModel <> "vosk" Then
                        PartialTextLabel.Text = "Stopping…"
                    End If

                    System.Threading.Tasks.Task.Run(Async Function()
                                                        Try
                                                            Await StopRecording()
                                                            If STTModel = "google" Then StopApiWatchdogTimer()
                                                        Catch ex As System.Exception

                                                        End Try

                                                        Me.Invoke(Sub()
                                                                      Me.StartButton.Enabled = False
                                                                      Me.LoadButton.Enabled = False

                                                                      If STTModel = "vosk" Then
                                                                          Addline(PartialTextLabel.Text)
                                                                      End If
                                                                      PartialTextLabel.Text = String.Empty
                                                                  End Sub)
                                                    End Function)

                    capturing = False

                End If
            End If
        End Sub

        Private Sub AudioButton_Click(sender As Object, e As EventArgs)
            ConfigureAudioOutputDevice()
        End Sub

        Private Sub QuitButton_Click(sender As Object, e As EventArgs)

            If capturing Then

                STTCanceled = True

                Me.StopButton.Enabled = False
                Me.AudioButton.Enabled = False
                Me.QuitButton.Enabled = False
                If STTModel <> "vosk" Then
                    PartialTextLabel.Text = "Stopping…"
                End If

                System.Threading.Tasks.Task.Run(Async Function()
                                                    Try
                                                        Await StopRecording()
                                                    Catch ex As System.Exception

                                                    End Try

                                                    Me.Invoke(Sub()
                                                                  Me.StartButton.Enabled = False
                                                                  Me.LoadButton.Enabled = False

                                                                  If STTModel = "vosk" Then
                                                                      Addline(PartialTextLabel.Text)
                                                                  End If
                                                                  PartialTextLabel.Text = String.Empty
                                                              End Sub)
                                                End Function)

                capturing = False

            End If
            Me.Close()
        End Sub

        Private Async Sub LoadButton_Click(sender As Object, e As EventArgs)
            If capturing Then Return

            Dim filepath As String = ""

            DragDropFormLabel = "Supported are audio files (*.wav, *.mp3, *.aac, *.m4a, *.mp4 and *.wma)"
            DragDropFormFilter = "Supported Files|*.wav;*.mp3;*.aac;*.m4a;*.mp4;*.wma|" &
                             "Wave files (*.wav)|*.wav|" &
                             "MP3 files (*.mp3)|*.mp3|" &
                             "AAC files (*.aac, *.m4a, *.mp4)|*.aac;*.m4a;*.mp4|" &
                             "WMA files (*.wma)|*.wma|" &
                             "All files|*.*"

            Using form As New DragDropForm()

                If form.ShowDialog() = DialogResult.OK Then
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    filepath = form.SelectedFilePath
                    If Not File.Exists(filepath) Then
                        ShowCustomMessageBox("The selected file was not found.")
                        Return
                    End If
                Else
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    Return
                End If
            End Using
            DragDropFormLabel = ""
            DragDropFormFilter = ""

            Dim splash As New SplashScreen($"Loading model...")
            splash.Show()
            splash.Refresh()

            cts = New CancellationTokenSource()
            STTCanceled = False
            audioBuffer.Clear()

            Try
                If Me.cultureComboBox.SelectedItem.ToString().StartsWith(GoogleSTT_Desc) Then
                    STTModel = "google"
                ElseIf Me.cultureComboBox.SelectedItem.ToString().StartsWith("ggml") Then
                    STTModel = "whisper"
                Else
                    STTModel = "vosk"
                End If

                Select Case STTModel

                    Case "google"

                        readerCts = New CancellationTokenSource()

                        ' Ask user for language code
                        Dim language As String = ShowSelectionForm("Select the language code you want to transcribe in:", $"{GoogleSTT_Desc}", GoogleSTTsupportedLanguages)

                        language = language.Trim()

                        If String.IsNullOrWhiteSpace(language) OrElse String.Equals(language, "ESC", StringComparison.OrdinalIgnoreCase) Then
                            splash.Close()
                            Return
                        End If

                        If Not GoogleSTTsupportedLanguages.Any(Function(code) code.Trim().Normalize().IndexOf(language, StringComparison.OrdinalIgnoreCase) = 0) Then
                            splash.Close()
                            ShowCustomMessageBox("This language code is not supported. Supported are: " & String.Join(", ", GoogleSTTsupportedLanguages))
                            Return
                        End If

                        ' Configure the streaming recognizer

                        GoogleLanguageCode = language

                    Case "vosk"
                        StartVosk()

                    Case "whisper"

                        Dim language As String = ShowCustomInputBox("Enter the language ISO code you want Whisper to transcribe (e.g. en, de, fr, etc.) or go with 'auto':", "Whisper Language Code", True, "auto")

                        language = language.ToLower()

                        If String.IsNullOrWhiteSpace(language) Or language = "esc" Or Not WhisperSupportedLanguages.Contains(language.ToLower()) Then
                            splash.Close()
                            If Not WhisperSupportedLanguages.Contains(language.ToLower()) And language <> "esc" Then
                                ShowCustomMessageBox("This language code is not supported. Supported are: Afrikaans (af), Albanian (sq), Amharic (am), Arabic (ar), Armenian (hy), Assamese (as), Azerbaijani (az), Bashkir (ba), Basque (eu), Belarusian (be), Bengali (bn), Bosnian (bs), Breton (br), Bulgarian (bg), Catalan (ca), Chinese (zh), Croatian (hr), Czech (cs), Danish (da), Dutch (nl), English (en), Estonian (et), Faroese (fo), Finnish (fi), French (fr), Galician (gl), Georgian (ka), German (de), Greek (el), Gujarati (gu), Haitian Creole (ht), Hausa (ha), Hebrew (he), Hindi (hi), Hungarian (hu), Icelandic (is), Indonesian (id), Italian (it), Japanese (ja), Javanese (jv), Kannada (kn), Kazakh (kk), Khmer (km), Kinyarwanda (rw), Kirghiz (ky), Korean (ko), Latvian (lv), Lithuanian (lt), Luxembourgish (lb), Macedonian (mk), Malagasy (mg), Malay (ms), Malayalam (ml), Maltese (mt), Maori (mi), Marathi (mr), Mongolian (mn), Myanmar (my), Nepali (ne), Norwegian (no), Occitan (oc), Pashto (ps), Persian (fa), Polish (pl), Portuguese (pt), Punjabi (pa), Romanian (ro), Russian (ru), Sanskrit (sa), Serbian (sr), Sindhi (sd), Sinhala (si), Slovak (sk), Slovenian (sl), Somali (so), Spanish (es), Sundanese (su), Swahili (sw), Swedish (sv), Tagalog (tl), Tajik (tg), Tamil (ta), Tatar (tt), Telugu (te), Thai (th), Turkish (tr), Ukrainian (uk), Urdu (ur), Uzbek (uz), Vietnamese (vi), Welsh (cy), Yiddish (yi), Yoruba (yo), Zulu (zu)")
                            End If
                            STTCanceled = True
                            Return
                        End If

                        StartWhisper(language)
                        STTCanceled = False
                        PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper is listening and working... (no partial results shown, please wait)")

                    Case Else
                        splash.Close()
                        ShowCustomMessageBox($"No valid model selected. Please select a model.")
                        Return

                End Select

                My.Settings.LastAudioSource = Me.deviceComboBox.SelectedItem.ToString()
                My.Settings.LastSpeechModel = Me.cultureComboBox.SelectedItem.ToString()
                My.Settings.LastSpeakerEnabled = Me.SpeakerIdent.Checked
                similarityThreshold = Double.Parse(Me.SpeakerDistance.Text)
                If STTModel = "google" Then
                    If similarityThreshold < 1 Then similarityThreshold = 1.0
                Else
                    If similarityThreshold = 0 Then similarityThreshold = 1.0
                    If similarityThreshold < 0.2 Then similarityThreshold = 0.2
                    If similarityThreshold > 2.5 Then similarityThreshold = 2.5
                End If
                My.Settings.LastSpeakerDistance = similarityThreshold

                My.Settings.Save()

                capturing = True
                Me.StartButton.Enabled = False
                Me.cultureComboBox.Enabled = False
                Me.deviceComboBox.Enabled = False
                Me.SpeakerIdent.Enabled = False
                Me.SpeakerDistance.Enabled = False
                Me.StopButton.Enabled = True
                Me.LoadButton.Enabled = False
                Me.AudioButton.Enabled = False
                splash.Close()

                Select Case STTModel
                    Case "google"
                        googleTranscriptStart = RichTextBox1.TextLength
                        Dim methodChoice As Integer = ShowCustomYesNoBox("Select your Google transcription method (you may have to try which one works better):", "Send chunks (faster)", "Stream (less gaps)")

                        Debug.WriteLine("Choice = " & methodChoice)

                        If methodChoice = 0 Then
                            splash.Close()
                            Return
                        End If

                        ' Splash schließen, UI ist bereits deaktiviert
                        splash.Close()

                        splash = New SplashScreen($"Transcribing file ...")
                        splash.Show()
                        splash.Refresh()

                        Try

                            ' Chunking vs. Streaming aufrufen
                            If methodChoice = 1 Then
                                Await GoogleChunkedTranscribeAudioFile(filepath)
                            Else
                                Await GoogleFileStreamTranscription(filepath)
                            End If

                        Catch ex As Exception
                            splash.Close()
                            ShowCustomMessageBox($"Error in Transcribing File using Google: {ex.Message}")
                        Finally
                            splash.Close()
                            Me.Invoke(Sub()
                                          capturing = False
                                          StartButton.Enabled = True
                                          StopButton.Enabled = False
                                          LoadButton.Enabled = True
                                          AudioButton.Enabled = True
                                          cultureComboBox.Enabled = True
                                          deviceComboBox.Enabled = True
                                          SpeakerIdent.Enabled = True
                                          SpeakerDistance.Enabled = True
                                      End Sub)
                        End Try

                    Case "vosk"
                        VoskTranscribeAudioFile(filepath)
                    Case "whisper"
                        WhisperTranscribeAudioFile(filepath)
                        ShowCustomMessageBox($"Transcription using Whisper has started In the background. You can continue working. Do not quit Word. Press 'Stop' to stop transcription.")
                End Select

            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"There has been an Error starting the transcription engine (Error: {ex.Message}).")

            End Try

        End Sub


        Private Async Sub ProcessButton_Click(sender As Object, e As EventArgs)
            If processCombobox.SelectedIndex >= 0 Then
                Dim selectedIndex As Integer = processCombobox.SelectedIndex
                If selectedIndex < TranscriptPromptsLibrary.Count Then
                    Dim OtherPrompt As String = TranscriptPromptsLibrary(selectedIndex)
                    Dim SelectedText As String = ""
                    If String.IsNullOrWhiteSpace(RichTextBox1.SelectedText) Then
                        SelectedText = RichTextBox1.Text
                    Else
                        SelectedText = RichTextBox1.SelectedText
                    End If
                    Dim LLMResult As String = Await LLM(OtherPrompt, SelectedText, "", "", False)

                    Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
                    Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection

                    If wordApp.Documents.Count > 0 Then
                        ' Collapse any existing selection towards the end
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        ' Insert the markdown text
                        InsertTextWithMarkdown(selection, LLMResult, True)
                    End If
                End If
            End If
        End Sub

        Private Sub LoadAudioDevices()
            deviceComboBox.Items.Clear()
            Dim i As Integer = 0
            For i = 0 To WaveInEvent.DeviceCount - 1
                Dim capabilities = WaveInEvent.GetCapabilities(i)
                Dim micName As String = $"{i}: {capabilities.ProductName}"
                '  a) plain mic
                deviceComboBox.Items.Add(micName)
                '  b) mic + system audio
                deviceComboBox.Items.Add($"{micName} (plus audio output)")
            Next

            ' Select default device (if available)
            Dim lastAudioSource As String = My.Settings.LastAudioSource
            If Not String.IsNullOrEmpty(lastAudioSource) AndAlso deviceComboBox.Items.Contains(lastAudioSource) Then
                deviceComboBox.SelectedItem = lastAudioSource
            ElseIf deviceComboBox.Items.Count > 0 Then
                deviceComboBox.SelectedIndex = 0
            End If
            Dim sel = TryCast(deviceComboBox.SelectedItem, String)
            _multiSourceSelected = (sel IsNot Nothing AndAlso sel.EndsWith("(plus audio output)"))
        End Sub

        Private Sub StartVosk()
            Dim modelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), Me.cultureComboBox.SelectedItem.ToString())
            Dim model As New Model(modelpath)
            recognizer = New VoskRecognizer(model, 16000.0F)
            If Me.SpeakerIdent.Checked Then
                ' Get the first available speaker model in the directory
                Dim speakerModelPath As String = System.IO.Directory.GetDirectories(System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), "Speaker\"), "vosk-model*").FirstOrDefault()
                If String.IsNullOrEmpty(speakerModelPath) Then
                    ShowCustomMessageBox($"No speaker model found (at {System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), "Speaker\")}. Speaker recognition will be disabled.")
                    Me.SpeakerIdent.Checked = False
                Else
                    Dim speakerModel As SpkModel = New SpkModel(speakerModelPath)
                    recognizer.SetSpkModel(speakerModel)

                End If

                Debug.WriteLine("Vosk recognizer initialized")

            End If

            recognizer.SetMaxAlternatives(0) ' Forces earlier finalization
            recognizer.SetWords(True) ' Enable word timestamps
            recognizer.SetPartialWords(True) ' Partial words emitted faster
        End Sub

        Private Sub StartWhisper(Optional language As String = "auto")
            Dim modelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), Me.cultureComboBox.SelectedItem.ToString())

            ' Load the model using WhisperFactory with the specified runtime options
            Dim factory As WhisperFactory = WhisperFactory.FromPath(modelpath)

            ' Configure the builder with language, threads, etc.
            If Me.SpeakerIdent.Checked Then
                Dim builder = factory.CreateBuilder() _
                    .WithLanguage(language) _
                    .WithThreads(Environment.ProcessorCount) _
                    .WithNoSpeechThreshold(Double.Parse(Me.SpeakerDistance.Text)) _
                    .WithTemperature(0.3) _
                    .WithTranslate()

                ' Build the recognizer
                WhisperRecognizer = builder.Build()
            Else
                Dim builder = factory.CreateBuilder() _
                    .WithLanguage(language) _
                    .WithThreads(Environment.ProcessorCount) _
                    .WithNoSpeechThreshold(Double.Parse(Me.SpeakerDistance.Text)) _
                    .WithTemperature(0.3)

                ' Build the recognizer
                WhisperRecognizer = builder.Build()
            End If
        End Sub

        Private Async Function StartGoogleSTT() As System.Threading.Tasks.Task
            ' ─── 1) Interceptor definieren, der bei jedem neuen Streaming-Aufruf einen frischen Token holt ───

            Dim callCreds As Grpc.Core.CallCredentials = Grpc.Core.CallCredentials.FromInterceptor(
                    Async Function(contextCall, metadata)
                        ' Nicht mehr context.GetFresh…, sondern unser lokaler Helper
                        Dim tokenToSend As String = Await GetFreshSTTToken(STTSecondAPI)
                        metadata.Add("Authorization", $"Bearer {tokenToSend}")
                        Await System.Threading.Tasks.Task.CompletedTask
                    End Function
                )

            ' ─── 2) Baue die ChannelCredentials mit Secure SSL + unserem Interceptor ───
            Dim channelCreds As Grpc.Core.ChannelCredentials = Grpc.Core.ChannelCredentials.Create(
                            Grpc.Core.ChannelCredentials.SecureSsl,
                            callCreds
                        )

            ' ─── 3) Erzeuge einen brandneuen SpeechClient, der das obige channelCreds verwendet ───
            Dim builder As New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
                            .Endpoint = STTEndpoint,
                            .ChannelCredentials = channelCreds
                        }
            client = builder.Build()

            ' ─── 4) Öffne die Streaming-Verbindung mit InitializeGoogleStream() ───
            '      Das ruft im Hintergrund “_stream = client.StreamingRecognize() …” und sendet die 
            '      StreamingConfig per WriteAsync. Beim ersten WriteAsync wird der Interceptor aktiv.
            Await InitializeGoogleStream()

            SyncLock ringBuffer
                ringBuffer.Clear()
            End SyncLock

            StartAudioQueueWriter()

        End Function

        Private Sub ResetGoogleStreamFlag()
            _googleStreamCompleted = False
        End Sub

        Private Async Function InitializeGoogleStream() As System.Threading.Tasks.Task

            streamingStartTime = DateTime.UtcNow
            ResetGoogleStreamFlag()

            Try
                ' Bidirektionales Streaming öffnen

                If Me.SpeakerIdent.Checked Then

                    'Dim maxSpk As Integer = CInt(Math.Ceiling(Double.Parse(Me.SpeakerDistance.Text)))

                    Dim minSpeakers As Integer = 2
                    Dim maxSpeakers As Integer = 6 ' Standard-Maximum, anpassen falls nötig

                    ' Versuchen Sie, die Werte aus der UI zu lesen, mit sicheren Standardwerten
                    Try
                        ' Annahme: SpeakerDistance ist jetzt MaxCount und ein neues TextFeld ist MinCount
                        maxSpeakers = CInt(Double.Parse(Me.SpeakerDistance.Text))
                    Catch
                        ' Bei Fehler Standardwerte verwenden
                    End Try

                    ' Die Werte auf den von Google unterstützten Bereich begrenzen
                    minSpeakers = Math.Max(2, minSpeakers)
                    maxSpeakers = Math.Max(minSpeakers, maxSpeakers)

                    _stream = client.StreamingRecognize()
                    Dim streamingConfig As New StreamingRecognitionConfig With {
                        .Config = New RecognitionConfig With {
                            .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                            .SampleRateHertz = 16000,
                            .LanguageCode = GoogleLanguageCode,
                            .EnableAutomaticPunctuation = True,
                            .EnableSpokenPunctuation = True,
                            .EnableWordTimeOffsets = False,
                            .EnableWordConfidence = False,
                            .Model = "latest_long",
                            .UseEnhanced = True,
                            .DiarizationConfig = New SpeakerDiarizationConfig With {
                                .EnableSpeakerDiarization = Me.SpeakerIdent.Checked,
                                        .MinSpeakerCount = minSpeakers,
                                    .MaxSpeakerCount = maxSpeakers
                                                            }
                        },
                        .InterimResults = True,
                        .SingleUtterance = False
                    }
                    Await _stream.WriteAsync(New StreamingRecognizeRequest With {.StreamingConfig = streamingConfig})


                Else
                    _stream = client.StreamingRecognize()
                    Dim streamingConfig As New StreamingRecognitionConfig With {
                    .Config = New RecognitionConfig With {
                        .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                        .SampleRateHertz = 16000,
                        .LanguageCode = GoogleLanguageCode,
                        .EnableAutomaticPunctuation = True,
                            .EnableSpokenPunctuation = True,
                            .EnableWordTimeOffsets = False,
                            .EnableWordConfidence = False,
                            .Model = "latest_long",
                            .UseEnhanced = True
                                },
                                .InterimResults = True,
                                .SingleUtterance = False
                            }
                    Await _stream.WriteAsync(New StreamingRecognizeRequest With {.StreamingConfig = streamingConfig})
                End If

                'StartAudioQueueWriter()

            Catch ex As System.Exception

                ShowCustomMessageBox("No speaker diarization available for this language (or other error).", $"{GoogleSTT_Desc} Language Code")
                _stream = client.StreamingRecognize()
                Dim streamingConfig As New StreamingRecognitionConfig With {
                .Config = New RecognitionConfig With {
                    .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                    .SampleRateHertz = 16000,
                    .LanguageCode = GoogleLanguageCode
                            },
                            .InterimResults = True
                        }
                _stream.WriteAsync(New StreamingRecognizeRequest With {.StreamingConfig = streamingConfig}).Wait()

            End Try

        End Function

        Private Async Sub StartButton_Click(sender As Object, e As EventArgs)

            If capturing Then
                Return
            End If

            Dim splash As New SplashScreen($"Loading model...")
            splash.Show()
            splash.Refresh()

            cts = New CancellationTokenSource()
            STTCanceled = False
            audioBuffer.Clear()

            Try
                If Me.cultureComboBox.SelectedItem.ToString().StartsWith(GoogleSTT_Desc) Then
                    STTModel = "google"
                ElseIf Me.cultureComboBox.SelectedItem.ToString().StartsWith("ggml") Then
                    STTModel = "whisper"
                Else
                    STTModel = "vosk"
                End If

                Select Case STTModel

                    Case "google"

                        readerCts = New CancellationTokenSource()

                        Dim language As String = ShowSelectionForm("Select the language code you want to transcribe in:", $"{GoogleSTT_Desc}", GoogleSTTsupportedLanguages)

                        language = language.Trim()

                        ' first handle empty or escape
                        If String.IsNullOrWhiteSpace(language) OrElse String.Equals(language, "esc", StringComparison.OrdinalIgnoreCase) Then
                            splash.Close()
                            STTCanceled = True
                            Return
                        End If

                        ' now do a true case‑insensitive lookup
                        If Not GoogleSTTsupportedLanguages.Any(
                                Function(code)
                                    Return String.Equals(code, language, StringComparison.OrdinalIgnoreCase)
                                End Function) Then
                            splash.Close()
                            ShowCustomMessageBox("This language code is not supported. Supported are: " & String.Join(", ", GoogleSTTsupportedLanguages), $"{GoogleSTT_Desc} Language Code")
                            STTCanceled = True
                            Return
                        End If

                        Try
                            GoogleLanguageCode = language
                            Await StartGoogleSTT()
                        Catch ex As System.Exception
                            ShowCustomMessageBox("Error starting transcription service: {ex.Message}", $"{GoogleSTT_Desc}")
                            STTCanceled = True
                            Return
                        End Try

                        If Not StartRecording() Then
                            splash.Close()
                            Return
                        End If

                        googleTranscriptStart = RichTextBox1.TextLength

                        _speakerTagToLabelMap.Clear()
                        _nextSpeakerNumber = 1

                        Me.googleReaderTask = StartGoogleReaderTask()

                    Case "vosk"

                        StartVosk()

                        If Not StartRecording() Then
                            splash.Close()
                            Return
                        End If

                    Case "whisper"

                        ' Define supported ISO 639-1 language codes

                        Dim language As String = ShowCustomInputBox("Enter the language ISO code you want Whisper to transcribe (e.g. en, de, fr, etc.) or go with 'auto':", "Whisper Language Code", True, "auto")

                        language = language.ToLower()

                        If String.IsNullOrWhiteSpace(language) Or language = "esc" Or Not WhisperSupportedLanguages.Contains(language.ToLower()) Then
                            splash.Close()
                            If Not WhisperSupportedLanguages.Contains(language.ToLower()) And language <> "esc" Then
                                ShowCustomMessageBox("This language code is not supported. Supported are: Afrikaans (af), Albanian (sq), Amharic (am), Arabic (ar), Armenian (hy), Assamese (as), Azerbaijani (az), Bashkir (ba), Basque (eu), Belarusian (be), Bengali (bn), Bosnian (bs), Breton (br), Bulgarian (bg), Catalan (ca), Chinese (zh), Croatian (hr), Czech (cs), Danish (da), Dutch (nl), English (en), Estonian (et), Faroese (fo), Finnish (fi), French (fr), Galician (gl), Georgian (ka), German (de), Greek (el), Gujarati (gu), Haitian Creole (ht), Hausa (ha), Hebrew (he), Hindi (hi), Hungarian (hu), Icelandic (is), Indonesian (id), Italian (it), Japanese (ja), Javanese (jv), Kannada (kn), Kazakh (kk), Khmer (km), Kinyarwanda (rw), Kirghiz (ky), Korean (ko), Latvian (lv), Lithuanian (lt), Luxembourgish (lb), Macedonian (mk), Malagasy (mg), Malay (ms), Malayalam (ml), Maltese (mt), Maori (mi), Marathi (mr), Mongolian (mn), Myanmar (my), Nepali (ne), Norwegian (no), Occitan (oc), Pashto (ps), Persian (fa), Polish (pl), Portuguese (pt), Punjabi (pa), Romanian (ro), Russian (ru), Sanskrit (sa), Serbian (sr), Sindhi (sd), Sinhala (si), Slovak (sk), Slovenian (sl), Somali (so), Spanish (es), Sundanese (su), Swahili (sw), Swedish (sv), Tagalog (tl), Tajik (tg), Tamil (ta), Tatar (tt), Telugu (te), Thai (th), Turkish (tr), Ukrainian (uk), Urdu (ur), Uzbek (uz), Vietnamese (vi), Welsh (cy), Yiddish (yi), Yoruba (yo), Zulu (zu)")
                            End If
                            STTCanceled = True
                            Return
                        End If
                        StartWhisper(language)

                        If Not StartRecording() Then
                            splash.Close()
                            STTCanceled = True
                            Return
                        End If
                        STTCanceled = False

                        PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper is listening and working... (no partial results shown, please wait)")
                    Case Else
                        splash.Close()
                        ShowCustomMessageBox($"No valid model selected. Please select a model.")
                        Return

                End Select
                My.Settings.LastAudioSource = Me.deviceComboBox.SelectedItem.ToString()
                My.Settings.LastSpeechModel = Me.cultureComboBox.SelectedItem.ToString()
                My.Settings.LastSpeakerEnabled = Me.SpeakerIdent.Checked
                similarityThreshold = Double.Parse(Me.SpeakerDistance.Text)
                If STTModel = "google" Then
                    If similarityThreshold < 1 Then similarityThreshold = 1.0
                Else
                    If similarityThreshold = 0 Then similarityThreshold = 1.0
                    If similarityThreshold < 0.2 Then similarityThreshold = 0.2
                    If similarityThreshold > 2.5 Then similarityThreshold = 2.5
                End If
                My.Settings.LastSpeakerDistance = similarityThreshold

                My.Settings.Save()

                If STTModel = "google" Then StartApiWatchdogTimer()

                capturing = True
                Me.StartButton.Enabled = False
                Me.cultureComboBox.Enabled = False
                Me.deviceComboBox.Enabled = False
                Me.SpeakerIdent.Enabled = False
                Me.SpeakerDistance.Enabled = False
                Me.StopButton.Enabled = True
                Me.LoadButton.Enabled = False
                Me.AudioButton.Enabled = False
                splash.Close()

            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"There has been an error starting the transcription engine (Error: {ex.Message}).")

            End Try
        End Sub


        Private Function StartRecording() As Boolean

            Dim ss As String = TryCast(Me.deviceComboBox.SelectedItem, String)
            Dim deviceIndex As Integer

            Dim pos As Integer = If(ss?.IndexOf(":"c), -1)
            If pos < 0 OrElse Not Integer.TryParse(ss.Substring(0, pos), deviceIndex) Then
                ShowCustomMessageBox($"Invalid device selection: '{ss}'")
                Return False
            End If

            waveIn = New WaveInEvent() With {
                    .DeviceNumber = deviceIndex,
                    .WaveFormat = New WaveFormat(16000, 1)
                }


            If MultiSourceEnabled Then

                ' Versuche, das in den Einstellungen gesetzte Ausgabegerät zu verwenden
                Dim audioOutputDeviceId As String = My.Settings.AudioOutputDevice
                Dim chosenDevice As MMDevice = Nothing

                'Debug.WriteLine("audioOutputDeviceId=" & audioOutputDeviceId)

                If Not String.IsNullOrEmpty(audioOutputDeviceId) Then
                    Try
                        Dim enumerator As New MMDeviceEnumerator()
                        chosenDevice = enumerator.GetDevice(audioOutputDeviceId)
                    Catch ex As System.Exception
                        ' Ungültige ID oder Gerät nicht gefunden → Fallback auf Default
                        chosenDevice = Nothing
                    End Try
                End If

                ' 1) LoopbackCapture mit spezifischem Gerät oder Default erstellen
                If chosenDevice IsNot Nothing Then
                    loopbackCapture = New WasapiLoopbackCapture(chosenDevice)
                Else
                    loopbackCapture = New WasapiLoopbackCapture()
                End If

                ' 2) Raw-Provider in native Format
                loopbackRawProvider = New BufferedWaveProvider(loopbackCapture.WaveFormat) With {
                                .DiscardOnBufferOverflow = True
                            }
                AddHandler loopbackCapture.DataAvailable, Sub(s, ev)
                                                              loopbackRawProvider.AddSamples(ev.Buffer, 0, ev.BytesRecorded)
                                                          End Sub

                ' 3) Resample von native → Mic-Format (16 kHz mono 16-bit)
                loopbackResampler = New MediaFoundationResampler(loopbackRawProvider, waveIn.WaveFormat) With {
                            .ResamplerQuality = 60
                        }

                ' 4) Aufnahme starten
                Try
                    loopbackCapture.StartRecording()
                Catch ex As System.Exception
                    ' Gerät evtl. exklusiv belegt → Fallback auf Mic-only
                    ShowCustomMessageBox("Cannot capture system audio: Device is in exclusive use or invalid. Continuing with mic only.")
                    loopbackCapture.Dispose()
                    loopbackCapture = Nothing
                    loopbackResampler?.Dispose()
                    loopbackResampler = Nothing
                    loopbackRawProvider = Nothing
                End Try
            End If

            If STTModel = "google" Then
                AddHandler waveIn.DataAvailable, AddressOf OnGoogleDataAvailable
            Else
                AddHandler waveIn.DataAvailable, AddressOf OnAudioDataAvailable
            End If
            waveIn.StartRecording()

            If IsNothing(prevExecState) Then
                prevExecState = SetThreadExecutionState(ES_KEEP_CURRENT_SETTING)
                STT_PreExistingSleepBlocker = False
            Else
                STT_PreExistingSleepBlocker = True
            End If

            Return True

        End Function

        Private Sub StartApiWatchdogTimer()
            ' Initialize the last response time to now.
            System.Threading.Interlocked.Exchange(_lastApiResponseTicks, DateTime.UtcNow.Ticks)

            ' Dispose of any existing timer to prevent orphans.
            _apiWatchdogTimer?.Dispose()

            ' Create a new timer that will call the CheckApiResponse method every 1000ms (1 sec).
            _apiWatchdogTimer = New System.Threading.Timer(
                                        AddressOf CheckApiResponse,
                                        Nothing,
                                        TimeSpan.FromSeconds(1),
                                        TimeSpan.FromSeconds(1)
                                    )
        End Sub

        Private Sub StopApiWatchdogTimer()
            _apiWatchdogTimer?.Dispose()
            _apiWatchdogTimer = Nothing
        End Sub

        Private Sub CheckApiResponse(state As Object)
            ' If we are not capturing, or a recovery is already in progress, do nothing.
            If Not capturing OrElse recoverySemaphore.CurrentCount = 0 Then
                Return
            End If

            ' Atomically read the last response time.
            Dim lastResponseTime As New DateTime(System.Threading.Interlocked.Read(_lastApiResponseTicks))

            ' Check if the elapsed time has exceeded our timeout.
            If (DateTime.UtcNow - lastResponseTime).TotalSeconds > API_RESPONSE_TIMEOUT_SECONDS Then
                ' The API has not responded in time. The stream is likely hung.
                Debug.WriteLine($"[ApiWatchdog] No API response for >{API_RESPONSE_TIMEOUT_SECONDS}s. Forcing stream recovery.")

                ' Stop the timer to prevent it from re-triggering while we recover.
                StopApiWatchdogTimer()

                ' Use our existing thread-safe recovery method to restart the stream.
                ' Then, after recovery, the watchdog will be restarted by the recovery logic itself.
                System.Threading.Tasks.Task.Run(Async Sub() Await TryRecoverGoogleStreamAsync())
            End If
        End Sub

        Private Function StartGoogleReaderTask() As System.Threading.Tasks.Task
            If readerCts IsNot Nothing Then
                Try
                    readerCts.Cancel()
                Catch
                End Try
            End If

            readerCts = New CancellationTokenSource()

            Dim newTask = System.Threading.Tasks.Task.Run(
                                    Async Sub()
                                        Dim token = readerCts.Token
                                        Try
                                            Dim enumerator = _stream.GetResponseStream().GetAsyncEnumerator(token)

                                            While Await enumerator.MoveNextAsync()
                                                System.Threading.Interlocked.Exchange(_lastApiResponseTicks, DateTime.UtcNow.Ticks)

                                                For Each result In enumerator.Current.Results
                                                    If result.IsFinal Then
                                                        ' --- NEW, CORRECTED FINAL RESULT LOGIC ---
                                                        If result.Alternatives.Count > 0 Then
                                                            Dim bestAlternative = result.Alternatives(0)
                                                            Dim finalTranscript As String = bestAlternative.Transcript.Trim()
                                                            _lastKnownPartialResult = ""

                                                            ' Check if we should ignore this result because it's a duplicate of a
                                                            ' partial result that was just committed during recovery.
                                                            If Not String.IsNullOrEmpty(_justCommittedPartialText) AndAlso
                                                               String.Equals(finalTranscript, _justCommittedPartialText.Trim(), StringComparison.OrdinalIgnoreCase) Then

                                                                ' This is a duplicate. Log it, clear the flag, and do nothing more.
                                                                Debug.WriteLine($"[ReaderTask] Ignoring duplicate final result: '{finalTranscript}'")
                                                                _justCommittedPartialText = ""

                                                            Else
                                                                ' This is a new, valid final result. Clear the "ignore" flag and proceed.
                                                                _justCommittedPartialText = ""

                                                                ' Now, apply formatting based on diarization settings.
                                                                If Me.SpeakerIdent.Checked AndAlso bestAlternative.Words.Count > 0 Then

                                                                    Dim currentSegment As New System.Text.StringBuilder()
                                                                    ' Get the label for the very first word's speaker.
                                                                    Dim currentSpeakerLabel As String = GetSpeakerLabel(bestAlternative.Words(0).SpeakerTag)

                                                                    For Each wordInfo In bestAlternative.Words
                                                                        Dim wordSpeakerLabel As String = GetSpeakerLabel(wordInfo.SpeakerTag)

                                                                        If wordSpeakerLabel <> currentSpeakerLabel Then
                                                                            ' The speaker has changed. Commit the previous speaker's segment.
                                                                            Dim segmentToCommit As String = $"{currentSpeakerLabel}: {currentSegment.ToString().Trim()}"
                                                                            Addline(segmentToCommit)

                                                                            ' Start a new segment for the new speaker.
                                                                            currentSegment.Clear()
                                                                            currentSpeakerLabel = wordSpeakerLabel
                                                                        End If

                                                                        ' Append the current word to the segment.
                                                                        currentSegment.Append(wordInfo.Word & " ")
                                                                    Next

                                                                    ' After the loop, commit the final segment.
                                                                    If currentSegment.Length > 0 Then
                                                                        Dim finalSegmentToCommit As String = $"{currentSpeakerLabel}: {currentSegment.ToString().Trim()}"
                                                                        Addline(finalSegmentToCommit)
                                                                    End If
                                                                Else
                                                                    ' --- Standard non-diarization logic ---
                                                                    ' Just use the simple Addline method with the raw transcript.
                                                                    Addline(finalTranscript)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        ' --- Interim result logic (remains the same) ---
                                                        If result.Alternatives.Count > 0 Then
                                                            Dim partialTranscript = result.Alternatives(0).Transcript
                                                            PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = partialTranscript)
                                                            _lastKnownPartialResult = partialTranscript
                                                        End If
                                                    End If
                                                Next
                                            End While

                                            ' --- The rest of the Catch blocks remain the same ---
                                        Catch ex As OperationCanceledException
                                            Debug.WriteLine($"[ReaderTask] Gracefully cancelled via OperationCanceledException.")
                                        Catch rex As RpcException
                                            If token.IsCancellationRequested OrElse rex.StatusCode = StatusCode.Cancelled Then
                                                Debug.WriteLine($"[ReaderTask] Gracefully cancelled via RpcException (Status: {rex.StatusCode}).")
                                            Else
                                                Debug.WriteLine($"[ReaderTask] Unexpected RpcException (Status: {rex.StatusCode}). Requesting recovery...")
                                                System.Threading.Tasks.Task.Run(Async Sub() Await TryRecoverGoogleStreamAsync())
                                            End If
                                        Catch ex As Exception
                                            Debug.WriteLine($"[ReaderTask] UNEXPECTED FATAL ERROR: {ex.ToString()}")
                                        End Try
                                    End Sub)
            Return newTask
        End Function

        Private Function GetSpeakerLabel(speakerTag As Integer) As String
            ' Check if we've already seen this tag in this session.
            If _speakerTagToLabelMap.ContainsKey(speakerTag) Then
                ' Yes, return the consistent label we already assigned.
                Return _speakerTagToLabelMap(speakerTag)
            Else
                ' No, this is a new speaker tag. Assign it a new label.
                Dim newLabel As String = $"Speaker {_nextSpeakerNumber}"
                _nextSpeakerNumber += 1

                ' Store the mapping for future use.
                _speakerTagToLabelMap.Add(speakerTag, newLabel)

                Return newLabel
            End If
        End Function



        Private Async Function TryRecoverGoogleStreamAsync() As System.Threading.Tasks.Task

            ' Check if there is a pending partial result that we need to commit.
            If Not String.IsNullOrWhiteSpace(_lastKnownPartialResult) Then
                ' Create a copy to avoid any potential race conditions.
                Dim partialToCommit As String = _lastKnownPartialResult

                ' Reset the class-level variable immediately.
                _lastKnownPartialResult = ""
                _justCommittedPartialText = partialToCommit

                ' Use the existing Addline method to append it to the RichTextBox.
                ' Addline is already thread-safe as it uses Invoke.
                Debug.WriteLine($"[TryRecover] Committing lost partial result: '{partialToCommit}'")
                Addline(partialToCommit)
            End If

            ' Asynchronously wait to acquire the semaphore. If another thread already has it,
            ' this thread will wait here without blocking a thread-pool thread.
            Await recoverySemaphore.WaitAsync()
            Try
                ' Now that we have the lock, perform the actual recovery.
                ' Any other threads calling this method will be waiting on the line above.
                Debug.WriteLine($"[TryRecover] Acquired semaphore. Starting recovery... ts={DateTime.UtcNow:HH:mm:ss.fff}")
                Await RecoverGoogleStream()
                streamingStartTime = DateTime.UtcNow ' Reset the timer *after* successful recovery
                Me.Invoke(Sub() StartApiWatchdogTimer())
            Finally
                ' CRITICAL: Always release the semaphore in a Finally block to prevent deadlocks.
                recoverySemaphore.Release()
                Debug.WriteLine($"[TryRecover] Released semaphore. ts={DateTime.UtcNow:HH:mm:ss.fff}")
            End Try
        End Function

        Private Sub StartAudioQueueWriter()
            writerTask = System.Threading.Tasks.Task.Run(
                                        Async Sub()
                                            Try
                                                ' This loop will automatically exit when the queue is completed by
                                                ' SafeCompleteAndDisposeGoogleStreamAsync.
                                                For Each chunk As Google.Protobuf.ByteString In audioQueue.GetConsumingEnumerable()
                                                    Try
                                                        ' If the stream was disposed during recovery, exit the writer immediately.
                                                        If _stream Is Nothing Then
                                                            Debug.WriteLine("[Writer] Stream is null. Exiting task.")
                                                            Return
                                                        End If

                                                        ' Send the audio chunk.
                                                        Await _stream.WriteAsync(New StreamingRecognizeRequest With {.AudioContent = chunk})

                                                    Catch ex As RpcException
                                                        ' A gRPC error occurred (e.g., the stream was cancelled).
                                                        ' This is an expected part of the shutdown/recovery cycle.
                                                        ' We just log it and exit the writer task gracefully.
                                                        Debug.WriteLine($"[Writer] RpcException (Status: {ex.StatusCode}). Exiting writer task.")
                                                        Return ' Exit the task.

                                                    Catch ex As NullReferenceException
                                                        ' This can happen if _stream is set to Nothing by another thread.
                                                        Debug.WriteLine("[Writer] Stream became null. Exiting writer task.")
                                                        Return ' Exit the task.

                                                    Catch ex As InvalidOperationException
                                                        ' This can happen if the stream is used after being closed.
                                                        Debug.WriteLine($"[Writer] InvalidOperationException (likely closed stream). Exiting writer task.")
                                                        Return ' Exit the task.

                                                    Catch ex As Exception
                                                        ' Catch any other unexpected error.
                                                        Debug.WriteLine($"[Writer] Unhandled exception in write loop: {ex.GetType().Name}. Exiting writer task.")
                                                        Return ' Exit the task gracefully.
                                                    End Try
                                                Next

                                            Catch ex As InvalidOperationException
                                                ' This exception occurs if GetConsumingEnumerable is called on a collection
                                                ' that has already been marked as complete and then disposed. This is an
                                                ' expected and normal part of the recovery cycle.
                                                Debug.WriteLine("[Writer] Task ending gracefully due to completed or disposed audio queue.")

                                            Catch ex As Exception
                                                ' A truly unexpected error occurred at the task level.
                                                Debug.WriteLine($"[Writer] UNEXPECTED FATAL ERROR in writer task: {ex.ToString()}")
                                            End Try
                                        End Sub)
        End Sub

        Private Async Sub OnGoogleDataAvailable(sender As Object, e As WaveInEventArgs)
            If _googleStreamCompleted Then Return

            'Debug.WriteLine($"[OnGoogleDataAvailable] start  ts={DateTime.UtcNow:HH:mm:ss.fff} ring={ringBuffer.Count} queue={audioQueue.Count}")

            Dim now = DateTime.UtcNow
            Dim elapsed = (now - streamingStartTime).TotalMilliseconds

            ' ——————— 1) MIX IN LOOPBACK, falls aktiviert ———————
            If MultiSourceEnabled AndAlso loopbackCapture IsNot Nothing AndAlso loopbackResampler IsNot Nothing Then
                Dim mixBuf(e.BytesRecorded - 1) As Byte
                Dim bytesRead = loopbackResampler.Read(mixBuf, 0, e.BytesRecorded)
                If bytesRead > 0 Then
                    For i As Integer = 0 To bytesRead - 1 Step 2
                        Dim micSample As Integer = BitConverter.ToInt16(e.Buffer, i)
                        Dim outSample As Integer = BitConverter.ToInt16(mixBuf, i)
                        Dim summedSample As Integer = micSample + outSample

                        If summedSample > Short.MaxValue Then summedSample = Short.MaxValue
                        If summedSample < Short.MinValue Then summedSample = Short.MinValue

                        Dim ba() As Byte = BitConverter.GetBytes(CShort(summedSample))
                        e.Buffer(i) = ba(0)
                        e.Buffer(i + 1) = ba(1)
                    Next
                End If
            End If

            Dim chunk As Google.Protobuf.ByteString =
                Google.Protobuf.ByteString.CopyFrom(e.Buffer, 0, e.BytesRecorded)

            ' 1) Ins Ring-Buffer schreiben (maximal 50 Chunks)
            SyncLock ringBuffer
                ringBuffer.Enqueue(chunk)
                If ringBuffer.Count > RING_BUFFER_SIZE Then ringBuffer.Dequeue()
            End SyncLock

            'Debug.WriteLine($"[OnGoogleDataAvailable] afterRing   ts={DateTime.UtcNow:HH:mm:ss.fff} ring={ringBuffer.Count}")

            ' 2) In die Queue schreiben (sofern offen)
            If Not audioQueue.IsAddingCompleted Then
                audioQueue.Add(chunk)
            End If

            'Debug.WriteLine($"[OnGoogleDataAvailable] afterQueue  ts={DateTime.UtcNow:HH:mm:ss.fff} queue={audioQueue.Count}")

            ' 3) Timeout prüfen und global recovern
            If elapsed > STREAMING_LIMIT_MS Then
                Debug.WriteLine($"[OnGoogleDataAvailable] Timeout detected. Requesting recovery... ts={DateTime.UtcNow:HH:mm:ss.fff}")
                ' Fire-and-forget the recovery task so we don't block the audio processing event.
                System.Threading.Tasks.Task.Run(Async Sub() Await TryRecoverGoogleStreamAsync())
                ' We now reset the timer *inside* the safe recovery method, not here.
                streamingStartTime = DateTime.UtcNow ' Resetting here is fine to prevent this from firing again immediately
                Return
            End If


        End Sub


        Private Async Function RecoverGoogleStream() As System.Threading.Tasks.Task
            Debug.WriteLine($"[RecoverGoogleStream] Starting...")

            ' --- 1. SHUTDOWN OLD COMPONENTS ---
            ' Store the old task before we overwrite the class-level variable.
            Dim oldReaderTask As System.Threading.Tasks.Task = Me.googleReaderTask

            ' Cancel the old reader's token source.
            If readerCts IsNot Nothing Then
                Try
                    readerCts.Cancel()
                    Debug.WriteLine($"[RecoverGoogleStream] Old CancellationTokenSource cancelled.")
                Catch ex As Exception
                    ' Ignore
                End Try
            End If

            ' Gracefully complete and dispose of the old stream object.
            ' This will help the old reader task exit cleanly.
            Await SafeCompleteAndDisposeGoogleStreamAsync(readerCts.Token)
            Debug.WriteLine($"[RecoverGoogleStream] Old stream disposed.")

            ' Now, explicitly wait for the old reader task to finish.
            ' This is the KEY to preventing the race condition.
            If oldReaderTask IsNot Nothing Then
                Try
                    Await oldReaderTask
                    Debug.WriteLine($"[RecoverGoogleStream] Old reader task has completed.")
                Catch ex As Exception
                    ' We expect exceptions here (like TaskCanceled), so we just log and continue.
                    Debug.WriteLine($"[RecoverGoogleStream] Awaiting old reader task threw: {ex.GetType().Name}")
                End Try
            End If

            ' --- 2. INITIALIZE NEW COMPONENTS ---

            ' This block is mostly the same, creating the new client and stream.
            Dim newToken As String = Await GetFreshSTTToken(STTSecondAPI)
            Dim callCreds = Grpc.Core.CallCredentials.FromInterceptor(
            Async Function(contextCall, metadata)
                metadata.Add("Authorization", $"Bearer {newToken}")
                Await System.Threading.Tasks.Task.CompletedTask
            End Function)
            Dim channelCreds = Grpc.Core.ChannelCredentials.Create(
            Grpc.Core.ChannelCredentials.SecureSsl, callCreds)
            Dim builder = New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
                                .Endpoint = STTEndpoint,
                                .ChannelCredentials = channelCreds
                            }
            client = builder.Build()

            ' Initialize the new stream.
            ' It's important that the call to InitializeGoogleStream happens *after*
            ' the old stream and reader are fully dead.
            ResetGoogleStreamFlag()
            Await InitializeGoogleStream()
            Debug.WriteLine($"[RecoverGoogleStream] New stream initialized.")

            ' --- 3. START NEW TASKS ---

            ' First, start the writer task. It needs a fresh audioQueue.
            ' You have a bug in SafeCompleteAndDisposeGoogleStreamAsync where you permanently close the queue.
            ' Let's fix that too. First, we need a NEW audio queue.
            audioQueue = New System.Collections.Concurrent.BlockingCollection(Of ByteString)()
            StartAudioQueueWriter()
            Debug.WriteLine($"[RecoverGoogleStream] New writer task started.")

            ' Now that everything old is gone, start the new reader task
            ' and assign it to our class-level variable.
            Me.googleReaderTask = StartGoogleReaderTask()
            Debug.WriteLine($"[RecoverGoogleStream] New reader task started.")
        End Function


        Private Async Function xRecoverGoogleStream() As System.Threading.Tasks.Task

            'Debug.WriteLine($"[RecoverGoogleStream] start      ts={DateTime.UtcNow:HH:mm:ss.fff} ring={ringBuffer.Count} queue={audioQueue.Count}")

            Try
                ' ─── 1) Neuer Token & Client wie gehabt ───
                Dim newToken As String = Await GetFreshSTTToken(STTSecondAPI)
                Dim callCreds = Grpc.Core.CallCredentials.FromInterceptor(
            Async Function(contextCall, metadata)
                metadata.Add("Authorization", $"Bearer {newToken}")
                Await System.Threading.Tasks.Task.CompletedTask
            End Function)
                Dim channelCreds = Grpc.Core.ChannelCredentials.Create(
            Grpc.Core.ChannelCredentials.SecureSsl, callCreds)
                Dim builder = New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
            .Endpoint = STTEndpoint,
            .ChannelCredentials = channelCreds
        }
                client = builder.Build()

                ' ─── 2) Stream neu initialisieren ───
                streamingStartTime = DateTime.UtcNow
                ResetGoogleStreamFlag()
                Await InitializeGoogleStream()

                'Debug.WriteLine($"[RecoverGoogleStream] inited     ts={DateTime.UtcNow:HH:mm:ss.fff}")

                ' ─── 3) Offset zurücksetzen ───
                Dim offset As Integer = 0
                Me.Invoke(Sub() offset = RichTextBox1.TextLength)
                googleTranscriptStart = offset

                ' ─── 4) Ring-Buffer wieder in die Queue spielen ───
                SyncLock ringBuffer
                    For Each oldChunk In ringBuffer
                        audioQueue.Add(oldChunk)
                    Next
                End SyncLock

                'Debug.WriteLine($"[RecoverGoogleStream] requeued   ts={DateTime.UtcNow:HH:mm:ss.fff} ring={ringBuffer.Count} queue={audioQueue.Count}")

                ' ─── 5) Reader neu starten ───
                StartGoogleReaderTask()

                SyncLock ringBuffer
                    ringBuffer.Clear()
                End SyncLock


                'Debug.WriteLine($"[RecoverGoogleStream] completed  ts={DateTime.UtcNow:HH:mm:ss.fff}")


            Catch ex As System.Exception
                'Debug.WriteLine($"[RecoverGoogleStream] ERROR      ts={DateTime.UtcNow:HH:mm:ss.fff} ex={ex.Message}")

            End Try
        End Function



        Private Async Function xxSafeCompleteAndDisposeGoogleStreamAsync(token As CancellationToken) As System.Threading.Tasks.Task
            ' 1) Beende den Stream sauber
            Try
                If _stream IsNot Nothing AndAlso Not _googleStreamCompleted Then
                    Await _stream.WriteCompleteAsync()
                    _googleStreamCompleted = True
                    ' ► KEIN CompleteAdding() hier!
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"Error in SafeComplete…: {ex.Message}")
            End Try

            ' 2) Jetzt erst: Queue schließen und auf writerTask warten
            audioQueue.CompleteAdding()
            Await writerTask

            ' 3) Finally: Stream-Objekt freigeben
            If _stream IsNot Nothing Then
                _stream.Dispose()
                _stream = Nothing
            End If
        End Function

        Private Async Function SafeCompleteAndDisposeGoogleStreamAsync(token As CancellationToken) As System.Threading.Tasks.Task
            ' 1) Beende den Stream sauber
            Try
                ' ONLY try to complete the stream if it's still valid AND hasn't been forcibly cancelled.
                If _stream IsNot Nothing AndAlso Not _googleStreamCompleted AndAlso Not token.IsCancellationRequested Then
                    Await _stream.WriteCompleteAsync()
                End If
            Catch ex As RpcException When ex.StatusCode = StatusCode.Cancelled
                ' This is an expected exception if the stream was cancelled while we tried to complete it.
                ' We can safely ignore it and proceed with cleanup.
                Debug.WriteLine($"[SafeComplete] Ignored expected RpcException (Cancelled).")
            Catch ex As Exception
                ' Catch other potential errors but don't let them stop the cleanup process.
                Debug.WriteLine($"[SafeComplete] Error during WriteCompleteAsync: {ex.Message}")
            End Try

            _googleStreamCompleted = True

            ' 2) Wait for the writerTask to finish.
            ' It will finish either because the queue was completed or it hit an exception.
            If writerTask IsNot Nothing AndAlso Not writerTask.IsCompleted Then
                Try
                    ' Don't try to complete the queue if it's already done.
                    If Not audioQueue.IsAddingCompleted Then
                        audioQueue.CompleteAdding()
                    End If
                    Await writerTask
                Catch ex As Exception
                    Debug.WriteLine($"[SafeComplete] Error while awaiting writerTask: {ex.Message}")
                End Try
            End If

            ' 3) Finally: Stream-Objekt freigeben
            ' This is a local method call, not the gRPC object. This is safe.
            _stream?.Dispose()
            _stream = Nothing
        End Function



        Private Function ConvertAudioToFloat(buffer As Byte()) As Single()
            ' Each sample = 2 bytes (16-bit), so half as many float samples
            Dim floatArray As Single() = New Single((buffer.Length \ 2) - 1) {}

            ' Convert raw 16-bit PCM -> -1.0f..+1.0f
            For i As Integer = 0 To buffer.Length - 2 Step 2
                Dim sample As Short = BitConverter.ToInt16(buffer, i)
                floatArray(i \ 2) = sample / 32768.0F
            Next

            Return floatArray
        End Function

        Private Sub OnLoopbackDataAvailable(sender As Object, e As WaveInEventArgs)
            ' Buffer the system audio for later mixing
            loopbackBuffer.AddSamples(e.Buffer, 0, e.BytesRecorded)
        End Sub


        Private Async Sub OnAudioDataAvailable(sender As Object, e As WaveInEventArgs)

            If MultiSourceEnabled AndAlso loopbackCapture IsNot Nothing AndAlso loopbackResampler IsNot Nothing Then
                Dim mixBuf(e.BytesRecorded - 1) As Byte
                ' read the same # of bytes from our resampler (16kHz mono 16-bit)
                Dim bytesRead = loopbackResampler.Read(mixBuf, 0, e.BytesRecorded)
                If bytesRead > 0 Then
                    For i As Integer = 0 To bytesRead - 1 Step 2
                        Dim micSample As Integer = BitConverter.ToInt16(e.Buffer, i)
                        Dim outSample As Integer = BitConverter.ToInt16(mixBuf, i)
                        Dim summedSample As Integer = micSample + outSample

                        ' clamp to Int16
                        If summedSample > Short.MaxValue Then summedSample = Short.MaxValue
                        If summedSample < Short.MinValue Then summedSample = Short.MinValue

                        Dim ba() As Byte = BitConverter.GetBytes(CShort(summedSample))
                        e.Buffer(i) = ba(0)
                        e.Buffer(i + 1) = ba(1)
                    Next
                End If
            End If

            Dim buffer As Byte() = e.Buffer
            Dim bytesRecorded As Integer = e.BytesRecorded

            ' Convert to 16-bit PCM samples
            Dim sampleCount As Integer = bytesRecorded / 2
            Dim samples(sampleCount - 1) As Single ' Float array for normalized audio

            For i As Integer = 0 To sampleCount - 1
                ' Convert 16-bit PCM to float (-1.0 to 1.0)
                Dim sample As Short = BitConverter.ToInt16(buffer, i * 2)
                Dim floatSample As Single = sample / 32768.0F
                samples(i) = floatSample
            Next

            ' **Normalize Samples**
            Dim maxSample As Single = samples.Max(Function(x) Math.Abs(x))
            If maxSample > 0 Then
                Dim gain As Single = 1.0F / maxSample ' Compute normalization factor
                For i As Integer = 0 To sampleCount - 1
                    samples(i) *= gain ' Apply normalization
                Next
            End If

            ' Convert back to 16-bit PCM
            For i As Integer = 0 To sampleCount - 1
                Dim normalizedSample As Short = CShort(samples(i) * 32767)
                Dim bytes As Byte() = BitConverter.GetBytes(normalizedSample)
                buffer(i * 2) = bytes(0)
                buffer(i * 2 + 1) = bytes(1)
            Next

            Select Case STTModel
                Case "vosk"
                    If recognizer IsNot Nothing AndAlso capturing Then
                        Dim jsonResult As String = ""
                        jsonResult = If(recognizer.AcceptWaveform(e.Buffer, e.BytesRecorded),
                                                recognizer.Result, recognizer.PartialResult)
                        ProcessTranscriptionJson(jsonResult)
                    End If

                Case "whisper"

                    If WhisperRecognizer Is Nothing Then Return

                    Try
                        ' Convert audio buffer to float array
                        Dim whispersamples As Single() = ConvertAudioToFloat(e.Buffer)

                        ' Append to buffer
                        audioBuffer.AddRange(whispersamples)
                        ' Only process when buffer has enough data 
                        If audioBuffer.Count < 32000 Then Return ' Adjust threshold based on sample rate
                        ' Copy buffered audio and clear buffer
                        Dim processSamples = audioBuffer.ToArray()
                        audioBuffer.Clear()
                        e.Buffer.Initialize() ' Clear the buffer    

                        ' Process transcription asynchronously
                        Await ProcessWhisper(processSamples)
                    Catch ex As Exception
                        Debug.WriteLine($"Error in OnAudioDataAvailable: {ex.Message}")
                    End Try
            End Select

        End Sub

        Private Async Function ProcessWhisper(samples As Single()) As Threading.Tasks.Task
            Try
                If STTCanceled Then Return

                Dim segments As IAsyncEnumerable(Of SegmentData) = WhisperRecognizer.ProcessAsync(samples)

                ' Iterate over the transcription results (only once)
                Dim enumerator = segments.GetAsyncEnumerator()

                If Await enumerator.MoveNextAsync() Then ' Only process the first result batch
                    Dim result As SegmentData = enumerator.Current
                    Dim text As String = result.Text

                    'Debug.WriteLine(text)
                    text = Regex.Replace(text, "\[.*?\]", String.Empty)
                    text = Regex.Replace(text, "\*.*?\*", String.Empty)

                    If Not String.IsNullOrWhiteSpace(text) And Not STTCanceled Then
                        Me.Invoke(Sub()
                                      RichTextBox1.AppendText(text & vbCrLf)
                                      RichTextBox1.ScrollToCaret()
                                  End Sub)
                    End If
                End If

                Await enumerator.DisposeAsync()

            Catch ex As Exception
                Debug.WriteLine($"Error in ProcessWhisper: {ex.Message}")
            End Try
        End Function

        Public Async Function WhisperTranscribeAudioFile(filepath As String) As Threading.Tasks.Task

            Try
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper is reading and transcribing your file...")

                Dim samples As Single() = LoadAudioToFloatArray(filepath) ' Use LoadAudioToFloatArray for MP3/FLAC

                Dim segments As IAsyncEnumerable(Of SegmentData) = WhisperRecognizer.ProcessAsync(samples)

                Dim enumerator = segments.GetAsyncEnumerator()

                Dim Exited As Boolean = False

                While Await enumerator.MoveNextAsync()

                    If cts.Token.IsCancellationRequested Then
                        Exited = True
                        Exit While
                    End If

                    Dim result As SegmentData = enumerator.Current
                    Dim Text = result.Text
                    If Not String.IsNullOrWhiteSpace(Text) And Not STTCanceled Then
                        Me.Invoke(Sub()
                                      RichTextBox1.AppendText(Text & vbCrLf)
                                      RichTextBox1.ScrollToCaret()
                                  End Sub)
                    End If

                End While
                Await enumerator.DisposeAsync()

                STTCanceled = True
                Await StopRecording()
                capturing = False
                Me.StartButton.Enabled = True
                Me.StopButton.Enabled = False
                Me.AudioButton.Enabled = True
                Me.LoadButton.Enabled = True
                Me.cultureComboBox.Enabled = True
                Me.deviceComboBox.Enabled = True
                Me.SpeakerIdent.Enabled = True
                Me.SpeakerDistance.Enabled = True
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")

                If Exited Then
                    ShowCustomMessageBox("Transcription aborted.")
                Else
                    ShowCustomMessageBox("The transcription of your file is complete.")
                End If

            Catch ex As Exception
                Debug.WriteLine($"Error in WhisperTranscribeAudioFile: {ex.Message}")
            End Try
        End Function

        Public Sub CancelTranscription()
            If cts IsNot Nothing Then
                cts.Cancel()
            End If
        End Sub


        Public Function LoadAudioToFloatArray(filepath As String) As Single()
            Using reader As New MediaFoundationReader(filepath) ' Supports MP3, WAV, FLAC, etc.
                ' Convert audio to 16kHz Mono (Whisper requires this format)
                Dim waveFormat = New WaveFormat(16000, 1) ' 16kHz, Mono
                Using resampler As New MediaFoundationResampler(reader, waveFormat)
                    resampler.ResamplerQuality = 60

                    ' Convert to floating point explicitly
                    Dim floatProvider As ISampleProvider = resampler.ToSampleProvider()

                    ' Read audio data into a floating-point array
                    Dim samples As New List(Of Single)()
                    Dim buffer As Single() = New Single(1024 - 1) {} ' Buffer for PCM float samples
                    Dim samplesRead As Integer

                    Do
                        samplesRead = floatProvider.Read(buffer, 0, buffer.Length)
                        If samplesRead > 0 Then
                            samples.AddRange(buffer.Take(samplesRead))
                        End If
                    Loop While samplesRead > 0

                    Return samples.ToArray()
                End Using
            End Using
        End Function


        ''' <summary>
        ''' Teilt eine Audiodatei in <60 s-Slices, ruft RecognizeAsync auf und beendet sich danach selbst.
        ''' </summary>
        Public Async Function GoogleChunkedTranscribeAudioFile(filepath As String) _
        As System.Threading.Tasks.Task

            ' ─── 0) Stelle sicher, dass client initialisiert ist ───
            If client Is Nothing Then
                Dim tokenToSend As String = Await GetFreshSTTToken(STTSecondAPI)
                Dim callCreds As Grpc.Core.CallCredentials = Grpc.Core.CallCredentials.FromInterceptor(
            Async Function(contextCall, metadata)
                metadata.Add("Authorization", $"Bearer {tokenToSend}")
                Await System.Threading.Tasks.Task.CompletedTask
            End Function
        )
                Dim channelCreds As Grpc.Core.ChannelCredentials = Grpc.Core.ChannelCredentials.Create(
            Grpc.Core.ChannelCredentials.SecureSsl,
            callCreds
        )
                Dim builder As New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
            .Endpoint = STTEndpoint,
            .ChannelCredentials = channelCreds
        }
                client = builder.Build()
            End If

            ' ─── 1) Lade PCM-Daten (16 kHz, mono, 16 Bit) ───
            Dim pcmData As Byte() = LoadAudioToPCM(filepath)

            ' ─── 2) Chunk-Parameter ───
            Dim bytesPerSec As Integer = 16000 * 2      ' 32 000 B/s
            Dim sliceLenSec As Integer = 50
            Dim overlapSec As Integer = 2
            Dim sliceSize As Integer = sliceLenSec * bytesPerSec
            Dim overlapSize As Integer = overlapSec * bytesPerSec
            Dim offset As Integer = 0

            ' ─── 3) Schleife über alle Slices ───
            While offset < pcmData.Length AndAlso Not STTCanceled
                Dim endPos = Math.Min(offset + sliceSize, pcmData.Length)
                Dim slice(endPos - offset - 1) As Byte
                Array.Copy(pcmData, offset, slice, 0, endPos - offset)

                ' ─── 4) RecognitionConfig bauen ───
                Dim config As New Google.Cloud.Speech.V1.RecognitionConfig With {
            .Encoding = Google.Cloud.Speech.V1.RecognitionConfig.Types.AudioEncoding.Linear16,
            .SampleRateHertz = 16000,
            .LanguageCode = GoogleLanguageCode,
            .EnableAutomaticPunctuation = True,
            .Model = "latest_long",
            .UseEnhanced = True
        }
                Dim audio As Google.Cloud.Speech.V1.RecognitionAudio =
            Google.Cloud.Speech.V1.RecognitionAudio.FromBytes(slice)

                ' ─── 5) Sync-API-Call ───
                Dim response As Google.Cloud.Speech.V1.RecognizeResponse =
            Await client.RecognizeAsync(config, audio)

                ' ─── 6) Ergebnisse anhängen ───
                For Each result As Google.Cloud.Speech.V1.SpeechRecognitionResult In response.Results
                    If result.Alternatives.Count > 0 Then
                        Addline(result.Alternatives(0).Transcript)
                    End If
                Next

                ' ─── 7) Wenn das letzte Slice war, Abbruch ───
                If endPos >= pcmData.Length Then
                    Exit While
                End If

                ' ansonsten Offset mit Überlappung weiter
                offset = endPos - overlapSize
                If offset < 0 Then offset = 0
            End While

            ' ─── 8) Abschlussmeldung ───
            ShowCustomMessageBox("Chunked transcription complete.", $"{AN} Transcriptor")
        End Function



        ''' <summary>
        ''' Transcribe a local file via StreamingRecognize by feeding the existing
        ''' audioQueue → writerTask → readerTask pipeline, then cleanly shutting down.
        ''' </summary>
        Public Async Function GoogleFileStreamTranscription(filepath As String) As System.Threading.Tasks.Task
            ' 1) UI-Status
            PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = $"{GoogleSTT_Desc} streaming file…")

            ' 2) Initialisiere den Stream und Reader
            readerCts = New CancellationTokenSource()
            Await StartGoogleSTT()                             ' öffnet _stream & schreibt Config
            googleTranscriptStart = RichTextBox1.TextLength
            googleReaderTask = StartGoogleReaderTask()         ' startet das Lesen der Antworten

            ' 3) Queue & Writer zurücksetzen
            audioQueue = New BlockingCollection(Of Google.Protobuf.ByteString)()
            StartAudioQueueWriter()                            ' schreibt später aus audioQueue in _stream

            ' 4) PCM-Daten laden
            Dim pcmFull As Byte() = LoadAudioToPCM(filepath)

            ' 5) RIFF-Header entfernen, falls WAV
            Dim pcmData = If(
        pcmFull.Length > 44 AndAlso
        System.Text.Encoding.ASCII.GetString(pcmFull, 0, 4) = "RIFF",
        pcmFull.Skip(44).ToArray(),
        pcmFull
    )

            ' 6) In file-Tempo (16 kHz) in die Queue legen
            Const chunkSize As Integer = 4096
            Dim bytesPerSec As Integer = 16000 * 2  ' 16 kHz × 16 Bit Mono = 32 000 B/s
            Dim pos As Integer = 0

            While pos < pcmData.Length AndAlso Not STTCanceled
                Dim len = Math.Min(chunkSize, pcmData.Length - pos)
                Dim chunk = Google.Protobuf.ByteString.CopyFrom(pcmData, pos, len)
                audioQueue.Add(chunk)

                ' → hier wird das Tempo gedrosselt:
                Dim delayMs = CInt(1000.0 * len / bytesPerSec)
                Await System.Threading.Tasks.Task.Delay(delayMs)

                pos += len
            End While

            ' 7) Queue schließen → Writer weiß, dass kein Nachschub mehr kommt
            audioQueue.CompleteAdding()

            ' 8) Stream sauber beenden & auf readerTask warten
            Await SafeCompleteAndDisposeGoogleStreamAsync(readerCts.Token)
            Await googleReaderTask

            ' 9) Cleanup & UI wieder freigeben
            StopApiWatchdogTimer()
            PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")
            ShowCustomMessageBox("Streaming transcription complete.", $"{AN} Transcriptor")
            Me.Invoke(Sub()
                          capturing = False
                          StartButton.Enabled = True
                          StopButton.Enabled = False
                          LoadButton.Enabled = True
                          AudioButton.Enabled = True
                          cultureComboBox.Enabled = True
                          deviceComboBox.Enabled = True
                          SpeakerIdent.Enabled = True
                          SpeakerDistance.Enabled = True
                      End Sub)
        End Function


        Public Async Function VoskTranscribeAudioFile(filepath As String) As Threading.Tasks.Task
            Try
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Vosk is reading and transcribing your file... press 'Esc' to abort")

                Dim Exited As Boolean = False

                ' Load PCM audio directly (no float conversion needed)
                Dim pcmData As Byte() = LoadAudioToPCM(filepath)

                ' Initialize Vosk recognizer 
                recognizer.Reset()

                ' Stream PCM data to Vosk recognizer
                Dim chunkSize As Integer = 4096 ' Process in small chunks
                Dim offset As Integer = 0

                While offset < pcmData.Length

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                        Exited = True
                        Exit While
                    End If

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                        ' Exit the loop
                        Exited = True
                        Exit While
                    End If

                    Dim chunkLength As Integer = Math.Min(chunkSize, pcmData.Length - offset)
                    Dim chunk As Byte() = pcmData.Skip(offset).Take(chunkLength).ToArray()

                    ' Feed the chunk into the recognizer
                    Dim resultAvailable As Boolean = recognizer.AcceptWaveform(chunk, chunk.Length)

                    ' Retrieve transcription
                    Dim resultText As String
                    If resultAvailable Then
                        Dim resultJson As String = recognizer.Result()
                        resultText = ExtractTextFromJson(resultJson)
                    Else
                        Dim partialJson As String = recognizer.PartialResult()
                        resultText = ExtractTextFromJson(partialJson)
                    End If

                    ' Update UI with transcribed text
                    If Not String.IsNullOrWhiteSpace(resultText) And Not STTCanceled Then
                        Me.Invoke(Sub()
                                      RichTextBox1.AppendText(resultText & vbCrLf)
                                      RichTextBox1.ScrollToCaret()
                                  End Sub)
                    End If

                    offset += chunkLength
                End While

                ' Get final result
                Dim finalResultJson As String = recognizer.FinalResult()
                Dim finalText As String = ExtractTextFromJson(finalResultJson)

                If Not String.IsNullOrWhiteSpace(finalText) Then
                    Me.Invoke(Sub()
                                  RichTextBox1.AppendText(finalText & vbCrLf)
                                  RichTextBox1.ScrollToCaret()
                              End Sub)
                End If

                ' Reset flags and UI
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")
                STTCanceled = True
                Await StopRecording()
                capturing = False
                Me.StartButton.Enabled = True
                Me.StopButton.Enabled = False
                Me.LoadButton.Enabled = True
                Me.AudioButton.Enabled = True
                Me.cultureComboBox.Enabled = True
                Me.deviceComboBox.Enabled = True
                Me.SpeakerIdent.Enabled = True
                Me.SpeakerDistance.Enabled = True
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")

                If Exited Then
                    ShowCustomMessageBox("Transcription aborted.")
                Else
                    ShowCustomMessageBox("The transcription of your file is complete.")
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error in VoskTranscribeAudioFile: {ex.Message}")
            End Try
        End Function

        Private Function ExtractTextFromJson(jsonString As String) As String
            Try
                Dim json As JObject = JObject.Parse(jsonString)
                If json.ContainsKey("text") Then
                    Return json("text").ToString()
                Else
                    Return String.Empty
                End If
            Catch ex As Exception
                Debug.WriteLine($"JSON Parsing Error: {ex.Message}")
                Return String.Empty
            End Try
        End Function


        Public Function LoadAudioToPCM(filepath As String) As Byte()
            Using reader As New MediaFoundationReader(filepath) ' Supports MP3, WAV, FLAC, etc.
                ' Convert audio to 16kHz Mono PCM (Vosk requires this format)
                Dim waveFormat = New WaveFormat(16000, 16, 1) ' 16kHz, 16-bit, Mono

                Using resampler As New MediaFoundationResampler(reader, waveFormat)
                    resampler.ResamplerQuality = 60

                    ' Use MemoryStream to store PCM data
                    Using memoryStream As New MemoryStream()
                        Using pcmWriter As New WaveFileWriter(memoryStream, waveFormat)
                            Dim buffer(4096 - 1) As Byte
                            Dim bytesRead As Integer

                            Do
                                bytesRead = resampler.Read(buffer, 0, buffer.Length)
                                If bytesRead > 0 Then
                                    pcmWriter.Write(buffer, 0, bytesRead)
                                End If
                            Loop While bytesRead > 0

                            pcmWriter.Flush()
                        End Using

                        ' Return raw PCM byte array
                        Return memoryStream.ToArray()
                    End Using
                End Using
            End Using
        End Function

        Private Sub ProcessTranscriptionJson(jsonString As String)
            Try
                Dim jsonObject As JObject = JObject.Parse(jsonString)



                If jsonObject.ContainsKey("text") AndAlso jsonObject("text") IsNot Nothing Then
                    Dim completedLine As String = jsonObject("text").ToString()
                    If Not String.IsNullOrWhiteSpace(completedLine) Then

                        ' Check if speaker embeddings are available
                        If jsonObject.ContainsKey("spk") AndAlso jsonObject("spk").Type = JTokenType.Array Then
                            Dim speakerArray As JArray = jsonObject("spk")
                            Dim speakerEmbedding As List(Of Double) = speakerArray.Select(Function(x) CDbl(x)).ToList()

                            ' Identify the speaker using cosine similarity
                            Dim speakerID As String = IdentifySpeaker(speakerEmbedding)
                            completedLine = $"{speakerID}: " & completedLine
                        End If

                        ' Add line to UI or output
                        Addline(completedLine)
                    End If
                ElseIf jsonObject.ContainsKey("partial") AndAlso jsonObject("partial") IsNot Nothing Then
                    partialText = jsonObject("partial").ToString()
                    PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = partialText)
                End If

            Catch ex As Exception
                MessageBox.Show("Error in ProcessTranscriptionJson: " & ex.Message, "Error")
            End Try
        End Sub

        ' Dictionary to store multiple embeddings per speaker for better matching
        Dim knownSpeakers As New Dictionary(Of String, List(Of List(Of Double)))
        Dim similarityThreshold As Double = 1.0 ' Adjusted for Euclidean Distance

        Private Function IdentifySpeaker(newEmbedding As List(Of Double)) As String
            ' Normalize new embedding
            newEmbedding = NormalizeEmbedding(newEmbedding)

            Dim bestMatch As String = "Unknown"
            Dim bestDistance As Double = Double.MaxValue

            For Each kvp In knownSpeakers
                Dim existingEmbeddings As List(Of List(Of Double)) = kvp.Value

                ' Compute similarity with the average embedding of the stored speaker
                Dim avgEmbedding As List(Of Double) = GetAverageEmbedding(existingEmbeddings)
                Dim distance As Double = EuclideanDistance(avgEmbedding, newEmbedding)

                ' Consider as the same speaker if distance is below threshold
                If distance < bestDistance AndAlso distance < similarityThreshold Then
                    bestMatch = kvp.Key
                    bestDistance = distance
                End If
            Next

            ' If no match, assign a new speaker ID
            If bestMatch = "Unknown" Then
                Dim newSpeakerID As String = "Speaker " & (knownSpeakers.Count + 1).ToString()
                knownSpeakers(newSpeakerID) = New List(Of List(Of Double)) From {newEmbedding}
                Return newSpeakerID
            Else
                ' Store the new embedding for future matches (stabilizes detection)
                knownSpeakers(bestMatch).Add(newEmbedding)

                ' Limit stored embeddings to the last 5 to prevent memory overuse
                If knownSpeakers(bestMatch).Count > 5 Then
                    knownSpeakers(bestMatch).RemoveAt(0)
                End If

                Return bestMatch
            End If
        End Function

        ' Normalize the embedding (ensures embeddings are comparable)
        Private Function NormalizeEmbedding(embedding As List(Of Double)) As List(Of Double)
            Dim norm As Double = Math.Sqrt(embedding.Sum(Function(x) x * x))
            If norm = 0 Then Return embedding
            Return embedding.Select(Function(x) x / norm).ToList()
        End Function

        ' Compute the average embedding
        Private Function GetAverageEmbedding(embeddings As List(Of List(Of Double))) As List(Of Double)
            Dim embeddingSize As Integer = embeddings(0).Count
            Dim avgEmbedding As New List(Of Double)(New Double(embeddingSize - 1) {})

            ' Sum up all embeddings
            For Each emb In embeddings
                For i As Integer = 0 To embeddingSize - 1
                    avgEmbedding(i) += emb(i)
                Next
            Next

            ' Divide by the number of stored embeddings
            For i As Integer = 0 To embeddingSize - 1
                avgEmbedding(i) /= embeddings.Count
            Next

            Return avgEmbedding
        End Function

        ' Compute Euclidean Distance between two speaker embeddings
        Private Function EuclideanDistance(vec1 As List(Of Double), vec2 As List(Of Double)) As Double
            Dim sum As Double = 0
            For i As Integer = 0 To vec1.Count - 1
                sum += (vec1(i) - vec2(i)) ^ 2
            Next
            Return Math.Sqrt(sum)
        End Function


        ' Function to compute cosine similarity between two speaker embeddings
        Private Function CosineSimilarity(vec1 As List(Of Double), vec2 As List(Of Double)) As Double
            Dim dotProduct As Double = vec1.Zip(vec2, Function(a, b) a * b).Sum()
            Dim magnitude1 As Double = Math.Sqrt(vec1.Sum(Function(a) a * a))
            Dim magnitude2 As Double = Math.Sqrt(vec2.Sum(Function(b) b * b))

            If magnitude1 = 0 OrElse magnitude2 = 0 Then
                Return 0
            End If

            Return dotProduct / (magnitude1 * magnitude2)
        End Function



        Private Sub Addline(completedline As String)
            completedline = completedline.Trim()

            SyncLock finalText
                finalText.AppendLine(completedline)
            End SyncLock

            ' This block is now deadlock-safe because it only writes to the UI.
            RichTextBox1.Invoke(Sub()
                                    ' Clear the partial text label
                                    PartialTextLabel.Text = ""

                                    ' Append the new completed line. AppendText is generally a safe "write" operation.
                                    RichTextBox1.AppendText(completedline & vbCrLf)

                                    RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                    RichTextBox1.ScrollToCaret()
                                    If STTModel = "google" Then googleTranscriptStart = RichTextBox1.TextLength
                                End Sub)
        End Sub


        Private Sub ReplaceAndAddLine(fullTranscript As String)
            RichTextBox1.Invoke(Sub()
                                    ' 1) select everything from the start index to the end…
                                    RichTextBox1.Select(googleTranscriptStart, RichTextBox1.TextLength - googleTranscriptStart)
                                    ' 2) replace it with the entire new transcript
                                    RichTextBox1.SelectedText = fullTranscript & Environment.NewLine
                                    ' 3) reset the caret to the end
                                    RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                    RichTextBox1.ScrollToCaret()
                                    If STTModel = "google" Then googleTranscriptStart = RichTextBox1.TextLength
                                End Sub)
        End Sub



        Public Sub LoadAndPopulateProcessComboBox(filePath As String, processComboBox As Forms.ComboBox)
            ' Execute LoadPrompts function
            Dim resultCode As Integer = LoadTranscriptPrompts(ExpandEnvironmentVariables(filePath))

            ' Clear the combo box before populating
            processComboBox.Items.Clear()

            ' Check if prompts were successfully loaded
            If resultCode = 0 AndAlso TranscriptPromptsTitles.Count > 0 Then
                ' Add the titles to the combo box
                For Each title As String In TranscriptPromptsTitles
                    processComboBox.Items.Add(title)
                Next
            End If
        End Sub

        Private Function LoadTranscriptPrompts(filePath As String) As Integer

            ' Initialize the return code to 0 (no error)
            Dim returnCode As Integer = 0

            filePath = ExpandEnvironmentVariables(filePath)

            'Debug.WriteLine($"Filepath = {filePath}")

            Try
                ' Verify the file exists
                If Not System.IO.File.Exists(filePath) Then
                    ShowCustomMessageBox("The transcript prompt library file was not found.")
                    Return 1
                End If

                TranscriptPromptsTitles.Clear()
                TranscriptPromptsLibrary.Clear()

                ' Read all lines from the file
                Dim lines = System.IO.File.ReadAllLines(filePath)

                For Each line As String In lines
                    ' Trim leading and trailing spaces
                    Dim trimmedLine = line.Trim()

                    ' Ignore empty lines and lines starting with ';'
                    If Not String.IsNullOrEmpty(trimmedLine) AndAlso Not trimmedLine.StartsWith(";") Then
                        ' Split the line by the delimiter '|'
                        Dim promptData = trimmedLine.Split("|"c)

                        ' Ensure there are at least two parts (title and prompt)
                        If promptData.Length >= 2 Then
                            Dim title = promptData(0).Trim()
                            Dim prompt = String.Join("|", promptData.Skip(1)).Trim()

                            ' Add title and prompt to the respective lists
                            TranscriptPromptsTitles.Add(title)
                            TranscriptPromptsLibrary.Add(prompt)
                        End If
                    End If
                Next

                ' Check if no prompts were found
                If TranscriptPromptsLibrary.Count = 0 Then
                    returnCode = 3
                    ShowCustomMessageBox("No prompts have been found in the configured transcript prompt library file.")
                End If

            Catch ex As System.IO.FileNotFoundException
                returnCode = 1
                ShowCustomMessageBox("The transcript prompt library file was not found: " & ex.Message)

            Catch ex As IndexOutOfRangeException
                returnCode = 2
                ShowCustomMessageBox("The format of the transcript prompt library file is not correct (is a '|' or text thereafter missing?): " & ex.Message)

            Catch ex As Exception
                returnCode = 99
                ShowCustomMessageBox("An unexpected error occurred while loading transcript prompts: " & ex.Message)
            End Try

            Return returnCode
        End Function

    End Class

    ' Text-to-Speech

    Private synth As New SpeechSynthesizer()

    Public Shared Sub SelectVoiceByNumber()
        ' Ensure the SpeechSynthesizer is available
        Dim synth As New SpeechSynthesizer()

        ' (1) Retrieve all available voices
        Dim installedVoices As List(Of InstalledVoice) = synth.GetInstalledVoices().ToList()
        Dim voiceNames As New List(Of String)()

        ' (2) Populate voice list
        Dim sb As New StringBuilder()
        sb.AppendLine("Available voices for Text-to-Speech:" & vbCrLf)

        For i As Integer = 0 To installedVoices.Count - 1
            Dim voiceInfo As VoiceInfo = installedVoices(i).VoiceInfo
            voiceNames.Add(voiceInfo.Name)
            sb.AppendLine($"{i}: {voiceInfo.Name}")
        Next

        If voiceNames.Count = 0 Then
            ShowCustomMessageBox("No voices available on this system.", "Text-to-Speech")
            Return
        End If

        Dim UserInput As String = ShowCustomInputBox(sb.ToString(), "Select Voice for Text Reader", True)

        If String.IsNullOrWhiteSpace(UserInput) Then Return

        Dim selectedIndex As Integer
        If Integer.TryParse(UserInput, selectedIndex) AndAlso selectedIndex >= 0 AndAlso selectedIndex < voiceNames.Count Then
            ' Get the selected voice name
            Dim chosenVoice As String = voiceNames(selectedIndex)
            Try
                synth.SelectVoice(chosenVoice)
                My.Settings.LastVoice = chosenVoice
                My.Settings.Save()

                synth.Speak($"Hello! I am now using the voice: {chosenVoice}")
            Catch ex As Exception
                MsgBox("Error selecting voice: " & ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        Else
            ShowCustomMessageBox("Invalid voice number entered.", "Text-to-Speech")
        End If
    End Sub

    Public Sub SpeakSelectedText()

        Debug.WriteLine("Status: " & synth.State.ToString())

        If synth.State = SynthesizerState.Speaking Then
            synth.SpeakAsyncCancelAll()
            ShowCustomMessageBox("Reading out aborted.", "Text-to-Speech")
            Return
        End If

        Try
            ' Get the active Word application
            Dim wordApp As Word.Application = Globals.ThisAddIn.Application

            ' Get the selected text
            Dim selectedText As String = wordApp.Selection.Text.Trim()

            If String.IsNullOrEmpty(selectedText) Then
                ShowCustomMessageBox("No text selected in Word.", "Text-to-Speech")
                Return
            End If

            ' Speak the selected text

            synth.SelectVoice(My.Settings.LastVoice)

            synth.SpeakAsync(selectedText)

            ShowCustomMessageBox($"Reading out the selected text (using {My.Settings.LastVoice}). You can stop this by again calling this function.", "Text-to-Speech")

        Catch ex As Exception
            MessageBox.Show("Error in SpeakSelectedText: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' Text-to-Speech Engine 

    Private Shared TTS_PreExistingSleepBlocker As Boolean = False

    Public Enum TTSEngine
        Google = 0
        OpenAI = 1
    End Enum

    Public Shared TTS_SelectedEngine As TTSEngine = TTSEngine.Google

    Public Sub DetectTTSEngines()
        ' — split auth endpoints —

        Dim auth1 As String = ThisAddIn.INI_Endpoint
        Dim auth2 As String = ThisAddIn.INI_Endpoint_2

        ' — split TTS endpoints —
        Dim ttsEps = If(String.IsNullOrEmpty(ThisAddIn.INI_TTSEndpoint),
                     Array.Empty(Of String)(),
                     INI_TTSEndpoint.Split("¦"c))
        Dim tts1 As String = If(ttsEps.Length > 0, ttsEps(0), "")
        Dim tts2 As String = If(ttsEps.Length > 1, ttsEps(1), "")

        ' reset
        TTS_googleAvailable = False : TTS_googleSecondary = False
        TTS_openAIAvailable = False : TTS_openAISecondary = False
        TTS_GoogleEndpoint = "" : TTS_OpenAIEndpoint = ""

        ' — Google (needs OAuth2 flags) —
        If auth1.Contains(GoogleIdentifier) AndAlso ThisAddIn.INI_OAuth2 Then
            TTS_googleAvailable = True
            TTS_googleSecondary = False
        End If
        If auth2.Contains(GoogleIdentifier) AndAlso ThisAddIn.INI_OAuth2_2 Then
            TTS_googleAvailable = True
            TTS_googleSecondary = True
        End If

        ' — OpenAI (no OAuth2) —
        If auth1.Contains(OpenAIIdentifier) Then
            TTS_openAIAvailable = True
            TTS_openAISecondary = False
        End If
        If auth2.Contains(OpenAIIdentifier) Then
            TTS_openAIAvailable = True
            TTS_openAISecondary = True
        End If

        ' — assign TTS URIs based on identifier match —
        If tts1.Contains(GoogleIdentifier) Then TTS_GoogleEndpoint = tts1
        If tts2.Contains(GoogleIdentifier) Then TTS_GoogleEndpoint = tts2

        If tts1.Contains(OpenAIIdentifier) Then TTS_OpenAIEndpoint = tts1
        If tts2.Contains(OpenAIIdentifier) Then TTS_OpenAIEndpoint = tts2

        ' if neither engine auth-configured, bail early
        If Not TTS_googleAvailable AndAlso Not TTS_openAIAvailable Then
            Return
        End If
    End Sub

    Private Shared Function UseSecondaryFor(engine As TTSEngine) As Boolean
        If engine = TTSEngine.Google Then
            Return TTS_googleSecondary
        Else
            Return TTS_openAISecondary
        End If
    End Function


    ' Token-Cache für TTS
    Private Shared ttsAccessToken1 As String = String.Empty
    Private Shared ttsTokenExpiry1 As DateTime = DateTime.MinValue
    Private Shared ttsAccessToken2 As String = String.Empty
    Private Shared ttsTokenExpiry2 As DateTime = DateTime.MinValue

    Private Shared Async Function GetFreshTTSToken(useSecond As Boolean) _
    As System.Threading.Tasks.Task(Of String)

        Try
            Dim token As String
            Dim expiry As DateTime

            If useSecond Then
                token = ttsAccessToken2
                expiry = ttsTokenExpiry2
            Else
                token = ttsAccessToken1
                expiry = ttsTokenExpiry1
            End If

            ' Wenn kein Token oder abgelaufen, neuen holen
            If String.IsNullOrEmpty(token) OrElse DateTime.UtcNow >= expiry Then
                ' Parameter je nach gewählter API
                Dim clientEmail = If(useSecond, INI_OAuth2ClientMail_2, INI_OAuth2ClientMail)
                Dim scopes = If(useSecond, INI_OAuth2Scopes_2, INI_OAuth2Scopes)
                Dim rawKey = If(useSecond, INI_APIKey_2, INI_APIKey)
                Dim authServer = If(useSecond, INI_OAuth2Endpoint_2, INI_OAuth2Endpoint)
                Dim life = If(useSecond, INI_OAuth2ATExpiry_2, INI_OAuth2ATExpiry)

                ' GoogleOAuthHelper konfigurieren
                GoogleOAuthHelper.client_email = clientEmail
                GoogleOAuthHelper.private_key = TranscriptionForm.FormatPrivateKey(rawKey)
                GoogleOAuthHelper.scopes = scopes
                GoogleOAuthHelper.token_uri = authServer
                GoogleOAuthHelper.token_life = life

                ' neuen Token holen
                Dim newToken As String = Await GoogleOAuthHelper.GetAccessToken()
                Dim newExpiry = DateTime.UtcNow.AddSeconds(life - 300)

                If useSecond Then
                    ttsAccessToken2 = newToken
                    ttsTokenExpiry2 = newExpiry
                Else
                    ttsAccessToken1 = newToken
                    ttsTokenExpiry1 = newExpiry
                End If

                token = newToken
            End If

            Return token

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(
            $"Error fetching TTS token: {ex.Message}",
            "TTS Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function

    Public Shared cts As New CancellationTokenSource()

    Private Shared Async Function GenerateOpenAITTSAsync(
        input As String,
        languageCode As String,
        voiceName As String,
        pitch As Double,
        speakingRate As Double
    ) As Task(Of Byte())

        Try

            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Dim apiKey = If(TTS_openAISecondary, DecodedAPI_2, DecodedAPI)

            Debug.WriteLine($"[TTS] OpenAI endpoint = '{TTS_OpenAIEndpoint}'")
            Debug.WriteLine($"[TTS] OpenAI API Key = '{apiKey}'")

            Using client As New System.Net.Http.HttpClient()
                client.DefaultRequestHeaders.Authorization =
                New Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)

                ' build JSON
                Dim j = New JObject From {
                {"model", TTS_OpenAI_Model},
                {"input", input},
                {"voice", voiceName},
                {"response_format", "mp3"},
                {"instructions", ""}
            }

                Dim content = New StringContent(j.ToString(), Encoding.UTF8, "application/json")

                ' POST to the detected OpenAI endpoint
                Dim resp = Await client.PostAsync(TTS_OpenAIEndpoint, content).ConfigureAwait(False)
                If resp.IsSuccessStatusCode Then
                    Return Await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(False)
                Else
                    Dim err = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                    Throw New System.Exception($"OpenAI TTS Error {resp.StatusCode}: {err}")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error in GenerateOpenAITTSAsync: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function



    Public Shared Async Function GenerateAudioFromText(input As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O", Optional nossml As Boolean = False, Optional Pitch As Double = 0, Optional SpeakingRate As Double = 1, Optional CurrentPara As String = "") As Task(Of Byte())

        Try

            If IsNothing(prevExecState) Then
                prevExecState = SetThreadExecutionState(ES_KEEP_CURRENT_SETTING)
                TTS_PreExistingSleepBlocker = False
            Else
                TTS_PreExistingSleepBlocker = True
            End If

            Dim eng = TTS_SelectedEngine

            If eng = TTSEngine.OpenAI Then
                ' strip off “ — Beschreibung” if present
                Dim rawVoice = voiceName.Split(" "c)(0)
                Return Await GenerateOpenAITTSAsync(input,
                                       languageCode,
                                       rawVoice,
                                       Pitch,
                                       SpeakingRate)
            End If

            Using httpClient As New HttpClient()

                Dim AccessToken As String = Await GetFreshTTSToken(UseSecondaryFor(TTSEngine.Google))
                If String.IsNullOrEmpty(AccessToken) Then
                    ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                    Return Nothing
                End If


                If String.IsNullOrEmpty(AccessToken) Then
                    ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                    Return Nothing
                End If

                httpClient.DefaultRequestHeaders.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken)

                Dim requestBody As JObject

                'Debug.WriteLine(input)

                Dim jsonPayload As String

                If input.Trim().StartsWith("{") Then
                    jsonPayload = input
                Else

                    Dim textlabel As String = "text"
                    Dim ssmlPattern As String = "<[^>]+>"  ' Matches any tag-like structure <...>

                    If nossml Then
                        input = Regex.Replace(input, ssmlPattern, String.Empty)
                    Else
                        If Regex.IsMatch(input, ssmlPattern) Then
                            If Not input.Trim().StartsWith("<speak>") Then
                                input = "<speak>" & input & "</speak>"
                            End If
                            textlabel = "ssml"
                        End If
                    End If

                    ' Process as single-speaker plain text
                    requestBody = New JObject From {
                    {"input", New JObject From {{$"{textlabel}", input}}},
                    {"voice", New JObject From {
                        {"languageCode", languageCode},
                        {"name", voiceName}
                    }},
                    {"audioConfig", New JObject From {
                        {"audioEncoding", "MP3"},
                        {"pitch", Pitch},
                        {"speakingRate", SpeakingRate},
                        {"effectsProfileId", New JArray("small-bluetooth-speaker-class-device")}
                    }}
                }
                    jsonPayload = requestBody.ToString()
                End If
                ' Convert payload to JSON
                Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")

                Try
                    ' Make API request

                    If Len(input) > TTSLargeText Then
                        Dim t As New Thread(Sub()
                                                ShowCustomMessageBox("Audio generation has started and runs in the background. Press 'Esc' to abort.).", "", 3, "", True)
                                            End Sub)
                        t.SetApartmentState(ApartmentState.STA)
                        t.Start()
                    End If

                    Dim response As HttpResponseMessage = Await httpClient.PostAsync(TTS_GoogleEndpoint & "text:synthesize", content, cts.Token).ConfigureAwait(False)

                    ' Error Handling: Check if API call failed
                    If response Is Nothing Then
                        ShowCustomMessageBox("Error generating audio: No response from Google TTS API.")
                        Return Nothing
                    End If

                    Dim responseString As String = Await response.Content.ReadAsStringAsync()

                    ' Debug output: Show API response for troubleshooting
                    Debug.WriteLine($"Google TTS API Response: {responseString}")

                    If response.IsSuccessStatusCode Then
                        Dim responseJson As JObject = JObject.Parse(responseString)

                        ' Check if "audioContent" exists in response
                        If responseJson.ContainsKey("audioContent") Then
                            Dim audioBase64 As String = responseJson("audioContent").ToString()
                            Return System.Convert.FromBase64String(audioBase64)
                        Else
                            ShowCustomMessageBox("Error generating audio: 'audioContent' not found in response.")
                            Return Nothing
                        End If
                    Else
                        ShowCustomMessageBox($"Error generating audio: API returned status {response.StatusCode}. Response: {responseString}{If(String.IsNullOrEmpty(CurrentPara), "", "Text: " & CurrentPara) & " [in clipboard]"}).")
                        If Not String.IsNullOrEmpty(CurrentPara) Then SLib.PutInClipboard(response.StatusCode & vbCrLf & vbCrLf & responseString & vbCrLf & vbCrLf & CurrentPara)
                        Return Nothing
                    End If
                Catch ex As TaskCanceledException
                    ShowCustomMessageBox("Audio generation aborted.")
                    Return Nothing
                Catch ex As Exception
                    MessageBox.Show($"Error in GenerateAudioFromText (HTTP): {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Nothing
                End Try

            End Using
        Catch ex As Exception
            MessageBox.Show($"Error in GenerateAudioFromText: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        Finally
            If IsNothing(prevExecState) Then
                SetThreadExecutionState(ES_CONTINUOUS)
            Else
                If Not TTS_PreExistingSleepBlocker Then
                    SetThreadExecutionState(prevExecState)
                    prevExecState = Nothing
                End If
            End If
        End Try

    End Function


        Public Function ParseTextToConversation(text As String) As List(Of Tuple(Of String, String))
            Dim conversation As New List(Of Tuple(Of String, String))
            Dim currentSpeaker As String = ""
            Dim currentText As String = ""

            Dim paragraphs As String() = text.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each para As String In paragraphs
                Dim trimmedText As String = para.Trim()
                If String.IsNullOrEmpty(trimmedText) Then Continue For

                ' Check if the paragraph starts with a speaker tag
                Dim newSpeaker As String = ""
                If hostTags.Any(Function(tag) trimmedText.StartsWith(tag, StringComparison.OrdinalIgnoreCase)) Then
                    newSpeaker = "H"
                    trimmedText = trimmedText.Substring(trimmedText.IndexOf(":"c) + 1).Trim()
                ElseIf guestTags.Any(Function(tag) trimmedText.StartsWith(tag, StringComparison.OrdinalIgnoreCase)) Then
                    newSpeaker = "G"
                    trimmedText = trimmedText.Substring(trimmedText.IndexOf(":"c) + 1).Trim()
                End If

                ' If a new speaker is detected, store the previous entry and start a new one
                If newSpeaker <> "" Then
                    If Not String.IsNullOrEmpty(currentSpeaker) Then
                        conversation.Add(Tuple.Create(currentSpeaker, currentText.Trim()))
                    End If
                    currentSpeaker = newSpeaker
                    currentText = trimmedText
                Else
                    ' Continue the current speaker's dialogue
                    If Not String.IsNullOrEmpty(currentSpeaker) Then
                        currentText &= " " & trimmedText
                    End If
                End If
            Next

            ' Add the last entry
            If Not String.IsNullOrEmpty(currentSpeaker) Then
                conversation.Add(Tuple.Create(currentSpeaker, currentText.Trim()))
            End If

            Return conversation
        End Function


    Async Sub GenerateAndPlayPodcastAudio(
        conversation As List(Of Tuple(Of String, String)),
        filepath As String,
        languagecode As String,
        hostVoice As String,
        guestVoice As String,
        pitch As Double,
        speakingrate As Double,
        nossml As Boolean
    )

        Try

            If IsNothing(prevExecState) Then
                prevExecState = SetThreadExecutionState(ES_KEEP_CURRENT_SETTING)
                TTS_PreExistingSleepBlocker = False
            Else
                TTS_PreExistingSleepBlocker = True
            End If

            Dim outputFiles As New List(Of String)

            ' ensure a valid output path
            If String.IsNullOrWhiteSpace(filepath) Then
                filepath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
            End If

            ' defaults
            If String.IsNullOrEmpty(languagecode) Then languagecode = "en-US"
            If String.IsNullOrEmpty(hostVoice) Then hostVoice = "en-US-Studio-O"
            If String.IsNullOrEmpty(guestVoice) Then guestVoice = "en-US-Casual-K"

            Dim Exited As Boolean = False
            Dim eng = TTS_SelectedEngine

            Using httpClient As New HttpClient()
                ' — set Authorization header once, based on engine —
                If eng = TTSEngine.Google Then
                    Debug.WriteLine($"[TTS] Using Google TTS engine with endpoint '{TTS_GoogleEndpoint}'")
                    ' Google: fetch OAuth token
                    Dim token = Await GetFreshTTSToken(TTS_googleSecondary)
                    If String.IsNullOrEmpty(token) Then
                        ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                        Return
                    End If
                    httpClient.DefaultRequestHeaders.Authorization =
                    New Net.Http.Headers.AuthenticationHeaderValue("Bearer", token)
                Else
                    Debug.WriteLine($"[TTS] Using OpenAI TTS engine with endpoint '{TTS_OpenAIEndpoint}'")
                    ' OpenAI: use API key
                    Dim key = If(TTS_openAISecondary, INI_APIKey_2, INI_APIKey)
                    httpClient.DefaultRequestHeaders.Authorization =
                    New Net.Http.Headers.AuthenticationHeaderValue("Bearer", key)
                End If

                ' start “running in background” message
                Dim t As New Thread(Sub()
                                        ShowCustomMessageBox(
                                        "Audio generation has started and runs in the background. Press 'Esc' to abort.",
                                        "", 3, "", True)
                                    End Sub)
                t.SetApartmentState(ApartmentState.STA)
                t.Start()

                ' process each speaker snippet
                For i = 0 To conversation.Count - 1

                    If (GetAsyncKeyState(Keys.Escape) And &H8000) <> 0 Then Exited = True : Exit For
                    If (GetAsyncKeyState(Keys.Escape) And 1) <> 0 Then Exited = True : Exit For

                    Dim speaker = conversation(i).Item1
                    Dim text = conversation(i).Item2
                    Dim voice = If(speaker = "H", hostVoice, guestVoice)

                    ' handle SSML stripping/wrapping
                    Dim textlabel = "text"
                    If Not nossml Then
                        If Regex.IsMatch(text, "<[^>]+>") AndAlso Not text.Trim().StartsWith("<speak>") Then
                            text = $"<speak>{text}</speak>"
                            textlabel = "ssml"
                        End If
                    Else
                        text = Regex.Replace(text, "<[^>]+>", "")
                    End If

                    Dim audioBytes As Byte()

                    If eng = TTSEngine.Google Then
                        ' — Google path —
                        Dim requestBody = New JObject From {
                        {"input", New JObject From {{textlabel, text}}},
                        {"voice", New JObject From {
                            {"languageCode", languagecode},
                            {"name", voice}
                        }},
                        {"audioConfig", New JObject From {
                            {"audioEncoding", "MP3"},
                            {"pitch", pitch},
                            {"speakingRate", speakingrate},
                            {"effectsProfileId", New JArray("small-bluetooth-speaker-class-device")}
                        }}
                    }

                        Dim content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
                        Dim resp = Await httpClient.PostAsync(TTS_GoogleEndpoint & "text:synthesize", content)
                        Dim respStr = Await resp.Content.ReadAsStringAsync()
                        Dim respJson = JObject.Parse(respStr)

                        If respJson.ContainsKey("audioContent") Then
                            audioBytes = Convert.FromBase64String(respJson("audioContent").ToString())
                        Else
                            ShowCustomMessageBox("Error: no audioContent in Google response.")
                            Continue For
                        End If

                    Else
                        ' — OpenAI path —
                        ' strip off any “ — Beschreibung” from the combo text
                        Dim rawVoice = voice.Split(" "c)(0)
                        audioBytes = Await GenerateOpenAITTSAsync(text, languagecode, rawVoice, pitch, speakingrate)
                    End If

                    Debug.WriteLine($"Generated audio of {audioBytes.Length} for speaker {speaker} ({voice}) with text length {text.Length} characters.")

                    ' save snippet
                    Dim tempFile = Path.Combine(Path.GetTempPath(), $"{AN2}_podcast_temp_{i}.mp3")
                    File.WriteAllBytes(tempFile, audioBytes)
                    outputFiles.Add(tempFile)

                    ' throttle
                    Await System.Threading.Tasks.Task.Delay(1000)
                Next

                ' merge & cleanup
                If Not Exited Then MergeAudioFiles(outputFiles, filepath)
                For Each f In outputFiles : File.Delete(f) : Next
            End Using

            If Exited Then
                ShowCustomMessageBox("Multi-speaker audio generation aborted.")
            Else
                If ShowCustomYesNoBox(
                    $"Your multi-speaker audio sequence has been generated ('{filepath}') and is ready to be played. Play it?",
                    "Yes", "No (file remains available)") = 1 Then
                    PlayAudio(filepath)
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"Error generating podcast audio: {ex.Message}")
        Finally
            If IsNothing(prevExecState) Then
                SetThreadExecutionState(ES_CONTINUOUS)
            Else
                If Not TTS_PreExistingSleepBlocker Then
                    SetThreadExecutionState(prevExecState)
                    prevExecState = Nothing
                End If
            End If
        End Try
    End Sub


    Sub MergeAudioFiles(inputFiles As List(Of String), outputFile As String)
            Try
                Using outputStream As New FileStream(outputFile, FileMode.Create)
                    For Each file In inputFiles
                        Dim mp3Bytes As Byte() = System.IO.File.ReadAllBytes(file)
                        outputStream.Write(mp3Bytes, 0, mp3Bytes.Length)
                    Next
                End Using
                Console.WriteLine("Podcast audio merged successfully!")
            Catch ex As Exception
                Debug.WriteLine($"Error merging audio files: {ex.Message}")
            End Try
        End Sub

        ' Function to save audio to a file
        Public Shared Sub SaveAudioToFile(audioData As Byte(), filePath As String)
            Try
                If audioData IsNot Nothing AndAlso audioData.Length > 0 Then
                    File.WriteAllBytes(filePath, audioData)
                    Debug.WriteLine($"Audio file saved: {filePath}")
                Else
                    Debug.WriteLine("No audio received.")
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error saving file: {ex.Message}")
            End Try
        End Sub

        ' Function to play the generated MP3 audio using NAudio
        Public Shared Sub PlayAudio(filePath As String)


            Dim splash As New SplashScreen($"Playing MP3... press 'Esc' to abort")
            If File.Exists(filePath) Then
                splash.Show()
                splash.Refresh()
            End If

            Try

                If File.Exists(filePath) Then

                    Using mp3Reader As New Mp3FileReader(filePath)
                        Using waveOut As New WaveOutEvent()
                            waveOut.Init(mp3Reader)
                            waveOut.Play()

                            ' Keep playing until the audio ends
                            While waveOut.PlaybackState = PlaybackState.Playing
                                Thread.Sleep(100)
                                System.Windows.Forms.Application.DoEvents()
                                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                                    Exit While
                                End If
                                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                                    Exit While
                                End If
                            End While

                            ' Stop playback
                            waveOut.Stop()
                        End Using ' Automatically disposes waveOut
                    End Using ' Automatically disposes mp3Reader

                    splash.Close()

                Else
                    splash.Close()
                    ShowCustomMessageBox("Audio file not found.")
                End If
            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"Error playing audio: {ex.Message}")
            End Try
        End Sub

        Shared Async Sub GenerateAndPlayAudio(textToSpeak As String, filepath As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O")

            Dim Temporary As Boolean = (filepath = "")

            Dim audioBytes As Byte() = Await System.Threading.Tasks.Task.Run(Function() GenerateAudioFromText(textToSpeak, languageCode, voiceName).Result)

            Try
                If audioBytes IsNot Nothing Then
                    If Temporary Then
                        filepath = System.IO.Path.Combine(ExpandEnvironmentVariables("%TEMP%"), $"{AN2}_temp.mp3")
                    End If
                    SaveAudioToFile(audioBytes, filepath)
                    Dim Result As Integer = 1
                    If Len(textToSpeak) > TTSLargeText Then
                        Result = ShowCustomYesNoBox("Your audio sequence has been generated " & If(Temporary, "", $"('{filepath}') ") & "and is ready to be played. Play it?", "Yes", If(Temporary, "No", "No (file remains available)"))
                    End If
                    If Result = 1 Then
                        PlayAudio(filepath)
                    End If
                    If Temporary Then
                        System.IO.File.Delete(filepath)
                    End If
                End If
            Catch ex As System.Exception

            End Try
        End Sub


        Public Sub ReadPodcast(Text As String)

            Dim NoSSML As Boolean = My.Settings.NoSSML
            Dim Pitch As Double = My.Settings.Pitch
            Dim SpeakingRate As Double = My.Settings.Speakingrate

            ' Create an array of InputParameter objects.
            Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Pitch", Pitch),
                    New SLib.InputParameter("Speaking Rate", SpeakingRate),
                    New SLib.InputParameter("No SSML", NoSSML)
                    }

            Dim conversation As List(Of Tuple(Of String, String)) = ParseTextToConversation(Text)
            Dim hasHost As Boolean = conversation.Any(Function(t) t.Item1 = "H")
            Dim hasGuest As Boolean = conversation.Any(Function(t) t.Item1 = "G")

            If hasHost AndAlso hasGuest Then
            Using frm As New TTSSelectionForm("Select the voice you wish to use for creating your audio file and configure where to save it.", $"{AN} Text-to-Speech - Select Voices", True) ' TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voice you wish to use for creating your audio file and configure where to save it.", $"{AN} Google Text-to-Speech - Select Voices", True)
                If frm.ShowDialog() = DialogResult.OK Then
                    Dim selectedVoices As List(Of String) = frm.SelectedVoices
                    Dim selectedLanguage As String = frm.SelectedLanguage
                    Dim outputPath As String = frm.SelectedOutputPath

                    Debug.WriteLine("Voices=" & selectedVoices(0))
                    Debug.WriteLine("TTS_SelectedEngine=" & TTS_SelectedEngine)

                    ' Call the procedure (the parameters are passed ByRef).
                    If ShowCustomVariableInputForm("Please enter the following parameters to apply when creating your podcast audio file:", $"Create Podcast Audio", params) Then

                        ' After OK is clicked, update your original variables:
                        Pitch = CDbl(params(0).Value)
                        SpeakingRate = CDbl(params(1).Value)
                        NoSSML = CBool(params(2).Value)

                        My.Settings.NoSSML = NoSSML
                        My.Settings.Pitch = Pitch
                        My.Settings.Speakingrate = SpeakingRate
                        My.Settings.Save()

                        GenerateAndPlayPodcastAudio(conversation, outputPath, selectedLanguage, selectedVoices(0).Replace(" (male)", "").Replace(" (female)", ""), selectedVoices(1).Replace(" (male)", "").Replace(" (female)", ""), Pitch, SpeakingRate, NoSSML)
                    End If
                End If
            End Using
        Else
                ' Missing either Host or Guest
                ShowCustomMessageBox($"No conversation was found. Use '{hostTags(0)}' and '{guestTags(0)}' to dedicate content to the host and guest.")
            End If

        End Sub


        Public Async Sub GenerateAndPlayAudioFromSelectionParagraphs(filepath As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O", Optional voiceNameAlt As String = "")

            Dim CurrentPara As String = ""

            Try

                Dim Temporary As Boolean = (filepath = "")
                Dim Alternate As Boolean = True

                If Temporary Then
                    filepath = System.IO.Path.Combine(ExpandEnvironmentVariables("%TEMP%"), $"{AN2}_temp.mp3")
                End If

                If voiceNameAlt = "" Then Alternate = False

                ' Get the current Word selection.
                Dim app As Word.Application = Globals.ThisAddIn.Application
                Dim selection As Selection = app.Selection
                If selection Is Nothing OrElse selection.Paragraphs.Count = 0 Then
                    ShowCustomMessageBox("No text selected.")
                    Return
                End If

                Dim NoSSML As Boolean = My.Settings.NoSSML
                Dim Pitch As Double = My.Settings.Pitch
                Dim SpeakingRate As Double = My.Settings.Speakingrate
                Dim ReadTitleNumbers As Boolean = False
                Dim CleanText As Boolean = False
                Dim CleanTextPrompt As String = My.Settings.CleanTextPrompt
                If String.IsNullOrWhiteSpace(CleanTextPrompt) Then CleanTextPrompt = SP_CleanTextPrompt

                ' Create an array of InputParameter objects.
                Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Pitch", Pitch),
                    New SLib.InputParameter("Speaking Rate", SpeakingRate),
                    New SLib.InputParameter("No SSML", NoSSML),
                    New SLib.InputParameter("Title Numbers", ReadTitleNumbers),
                    New SLib.InputParameter("Clean text", CleanText)
                    }

                ' Call the procedure (the parameters are passed ByRef).
                If Not ShowCustomVariableInputForm("Please enter the following parameters to apply when creating your audio file based on your text:", $"Create Audio", params) Then Return

                Pitch = CDbl(params(0).Value)
                SpeakingRate = CDbl(params(1).Value)
                NoSSML = CBool(params(2).Value)
                ReadTitleNumbers = CBool(params(3).Value)
                CleanText = CBool(params(4).Value)

                My.Settings.NoSSML = NoSSML
                My.Settings.Pitch = Pitch
                My.Settings.Speakingrate = SpeakingRate
                My.Settings.Save()

                If CleanText Then
                    CleanTextPrompt = ShowCustomInputBox("Please enter the prompt to 'clean' the text with (each paragraph will be submitted to this prompt)", "Create Audio", False, CleanTextPrompt).Trim()
                    If CleanTextPrompt = "ESC" Then Return
                    If CleanTextPrompt = "" Then
                        CleanText = False
                    Else
                        My.Settings.CleanTextPrompt = CleanTextPrompt
                        My.Settings.Save()
                    End If
                End If

                Dim totalParagraphs As Integer = selection.Paragraphs.Count
                Dim tempFiles As New List(Of String)
                Dim paragraphIndex As Integer = 0
                Dim sentenceEndPunctuation As String() = {".", "!", "?", ";", ":", ",", ")", "]", "}"}
                Dim bracketedTextPattern As String = "^\s*[\(\[\{][^\)\]\}]*[\)\]\}]\s*$"

                Dim voiceName1 As String = voiceName
                Dim voiceName2 As String = voiceNameAlt
                Dim currentVoiceName As String = voiceName1
                Dim firstTitleEncountered As Boolean = False
                Dim LastTextWasTitle As Boolean = False

                ShowProgressBarInSeparateThread($"{AN} Audio Generation", "Starting audio generation...")
                ProgressBarModule.CancelOperation = False

                ' Process each paragraph in the selection.
                For Each para As Paragraph In selection.Paragraphs
                    ' Allow the user to abort by pressing Escape.
                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Or (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Or ProgressBarModule.CancelOperation Then
                        For Each file In tempFiles
                            Try
                                If IO.File.Exists(file) Then IO.File.Delete(file)
                            Catch ex As Exception
                                Debug.WriteLine($"Error deleting temp file {file}: {ex.Message}")
                            End Try
                        Next
                        ShowCustomMessageBox("Audio generation aborted by user.")
                        ProgressBarModule.CancelOperation = True
                        Return
                    End If

                    ' Get the trimmed paragraph text.
                    Dim paraText As String

                    ' Check if the paragraph has numbering
                    If Not String.IsNullOrEmpty(para.Range.ListFormat.ListString) And ReadTitleNumbers Then
                        ' Include the numbering before the paragraph text
                        paraText = para.Range.ListFormat.ListString.Trim(".") & vbCrLf & para.Range.Text.Trim()
                    Else
                        ' No numbering, just take the paragraph text
                        paraText = para.Range.Text.Trim()
                    End If


                    ' Skip paragraphs that are empty...
                    If String.IsNullOrWhiteSpace(paraText) Or Regex.IsMatch(paraText, bracketedTextPattern) Then Continue For
                    ' ...or that contain only numbers or control characters.
                    If Regex.IsMatch(paraText, "^[\d\p{C}\s]+$") Then Continue For

                    Dim lastChar As String = paraText.Substring(paraText.Length - 1)

                    ' Check if the last character is one of the defined punctuation marks
                    If Not sentenceEndPunctuation.Contains(lastChar) Then
                        ' Append a period
                        paraText = paraText & "."
                    End If

                    ' Determine if this paragraph is part of a bullet list.
                    Dim isBullet As Boolean = False
                    If para.Range.ListFormat IsNot Nothing AndAlso para.Range.ListFormat.ListType <> WdListType.wdListNoNumbering Then
                        isBullet = True
                    End If

                    ' Determine if the paragraph “looks like” a title.
                    Dim isTitle As Boolean = False
                    Dim styleName As String = ""
                    Try
                        styleName = para.Range.Style.NameLocal.ToString().ToLower()
                    Catch ex As Exception
                        Debug.WriteLine("Error retrieving style: " & ex.Message)
                    End Try
                    If styleName.Contains("heading") Then
                        isTitle = True
                    Else
                        Dim lineCount As Long = para.Range.ComputeStatistics(WdStatistic.wdStatisticLines)
                        If lineCount <= 2 Then
                            isTitle = True
                        End If
                        If Not paraText.EndsWith(".") Then
                            isTitle = True
                        End If
                    End If

                    Debug.WriteLine("Para = " & paraText & vbCrLf & vbCrLf)
                    Debug.WriteLine("IsTitle = " & isTitle & vbCrLf)
                    CurrentPara = Left(paraText, 400) & "..."

                    If isTitle AndAlso Alternate Then
                        If Not firstTitleEncountered Then
                            firstTitleEncountered = True
                            ' For the very first title, keep the current voice unchanged.
                        Else
                            If Not LastTextWasTitle Then
                                ' Switch the voice if the last paragraph was not a title.
                                Debug.WriteLine("Switching ...")
                                If currentVoiceName = voiceName1 Then
                                    currentVoiceName = voiceName2
                                Else
                                    currentVoiceName = voiceName1
                                End If
                            End If
                        End If
                        LastTextWasTitle = True
                    Else
                        LastTextWasTitle = False
                    End If

                    ' Set the maximum value if you know the total number of steps.
                    GlobalProgressMax = totalParagraphs

                    ' Update the current progress value and status label.
                    GlobalProgressValue = paragraphIndex + 1
                    GlobalProgressLabel = $"Paragraph {paragraphIndex + 1} of {totalParagraphs} (some may be skipped)"

                    ' For bullet lists, insert a short pause BEFORE the paragraph.
                    If isBullet Then
                        Dim silenceFileBefore As String = Await GenerateSilenceAudioFileAsync(0.1)
                        If Not String.IsNullOrEmpty(silenceFileBefore) Then tempFiles.Add(silenceFileBefore)
                    End If

                    If CleanText Then
                        ' Remove any unwanted characters from the paragraph text.
                        paraText = Await LLM(CleanTextPrompt, "<TEXTTOPROCESS>" & paraText & "</TEXTTOPROCESS>", "", "", 0, False, True)
                        paraText = paraText.Trim().Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "").Trim()
                        CurrentPara = Left(CurrentPara, 100) & $"... [cleaned: {Left(paraText, 400)}...]"
                        Debug.WriteLine("Cleaned Para = " & paraText & vbCrLf & vbCrLf)
                    End If

                    ' Generate the audio for the paragraph via your TTS API.
                    Dim paragraphAudioBytes As Byte() = Await GenerateAudioFromText(paraText, languageCode, currentVoiceName, NoSSML, Pitch, SpeakingRate, CurrentPara)

                    CurrentPara = ""

                    If paragraphAudioBytes IsNot Nothing Then
                        Dim tempParaFile As String = Path.Combine(Path.GetTempPath(), $"{AN2}_temp_para_{paragraphIndex}.mp3")
                        File.WriteAllBytes(tempParaFile, paragraphAudioBytes)
                        tempFiles.Add(tempParaFile)
                    Else
                        ' If audio generation failed, skip this paragraph.
                        Continue For
                    End If

                    ' For bullet lists, insert a short pause AFTER the paragraph.
                    If isBullet Then
                        Dim silenceFileAfterBullet As String = Await GenerateSilenceAudioFileAsync(0.3)
                        If Not String.IsNullOrEmpty(silenceFileAfterBullet) Then tempFiles.Add(silenceFileAfterBullet)
                    End If

                ' After each paragraph, add an extra pause:
                ' • Use a medium pause (0.7 sec) for titles.
                ' • Otherwise use a short pause (0.3 sec).
                If isTitle Then
                    Dim silenceFileTitle As String = Await GenerateSilenceAudioFileAsync(0.7)
                    If Not String.IsNullOrEmpty(silenceFileTitle) Then tempFiles.Add(silenceFileTitle)
                    Else
                    Dim silenceFileRegular As String = Await GenerateSilenceAudioFileAsync(0.3)
                    If Not String.IsNullOrEmpty(silenceFileRegular) Then tempFiles.Add(silenceFileRegular)
                    End If

                    Await System.Threading.Tasks.Task.Delay(1000) ' Delay to not overhwelm the API

                    paragraphIndex += 1
                Next

                ' If no valid paragraphs were found, notify the user.
                If tempFiles.Count = 0 Then
                    ShowCustomMessageBox("No valid paragraphs found For audio generation; skipping empty ones And {...}, [...] And (...).")
                    Return
                End If

                If Not ProgressBarModule.CancelOperation Then
                    ' Merge all the temporary audio files into one final file.
                    MergeAudioFiles(tempFiles, filepath)
                End If

                ' Cleanup temporary files.
                For Each file In tempFiles
                    Try
                        If IO.File.Exists(file) Then IO.File.Delete(file)
                    Catch ex As Exception
                        Debug.WriteLine($"Error deleting temp file {file}: {ex.Message}")
                    End Try
                Next

                If Not ProgressBarModule.CancelOperation Then
                    ProgressBarModule.CancelOperation = True
                    ' Play the merged audio file.
                    PlayAudio(filepath)
                    If Temporary Then
                        System.IO.File.Delete(filepath)
                    End If
                Else
                    ProgressBarModule.CancelOperation = True
                    ShowCustomMessageBox("Audio generation aborted by user.")
                End If

            Catch ex As Exception
                ShowCustomMessageBox($"Error generating audio from selected paragraphs ({ex.Message}{If(String.IsNullOrEmpty(CurrentPara), "", "; Text: " & CurrentPara) & " [in clipboard]"}).")
                If Not String.IsNullOrEmpty(CurrentPara) Then SLib.PutInClipboard(ex.Message & vbCrLf & vbCrLf & CurrentPara)
            End Try
        End Sub

        Private Async Function GenerateSilenceAudioFileAsync(durationSeconds As Double) As Task(Of String)
            Return Await System.Threading.Tasks.Task.Run(Function() GenerateSilenceAudioFile(durationSeconds))
        End Function

        ' Synchronous helper that creates a buffer of silence and encodes it to MP3.
        Private Function GenerateSilenceAudioFile(durationSeconds As Double) As String
            Try
                ' Set audio format parameters.
                Dim sampleRate As Integer = 24000       ' Adjust as needed to match your TTS output.
                Dim channels As Integer = 1
                Dim bitsPerSample As Integer = 16
                Dim blockAlign As Integer = channels * (bitsPerSample \ 8)
                Dim totalSamples As Integer = CInt(sampleRate * durationSeconds)
                Dim totalBytes As Integer = totalSamples * blockAlign

                ' Create a buffer filled with zeros (silence).
                Dim silenceBytes(totalBytes - 1) As Byte
                ' (The array is automatically initialized to zeros.)

                ' Generate a temporary file name.
                Dim tempFile As String = Path.Combine(Path.GetTempPath(), $"{AN2}_silence_{CInt(durationSeconds * 1000)}ms.mp3")

                ' Wrap the silence buffer in a MemoryStream and then a RawSourceWaveStream.
                Using ms As New MemoryStream(silenceBytes)
                    Dim waveFormat As New WaveFormat(sampleRate, bitsPerSample, channels)
                    Using waveStream As New RawSourceWaveStream(ms, waveFormat)
                        ' Encode the silence to MP3.
                        MediaFoundationEncoder.EncodeToMp3(waveStream, tempFile)
                    End Using
                End Using

                Return tempFile
            Catch ex As Exception
                Debug.WriteLine($"Error generating silence audio: {ex.Message}")
                Return Nothing
            End Try
        End Function


    Public Class TTSSelectionForm
        Inherits Form

        ' -- Controls --
        Private lblIntro As Label

        ' engine selector combo:
        Private cmbEngine As Forms.ComboBox


        ' --- Set 1 Controls ---
        Private lblSet1 As Label
        Private cmbLanguage1 As Forms.ComboBox
        Private cmbVoice1A As Forms.ComboBox
        Private btnPlay1A As Forms.Button
        Private cmbVoice1B As Forms.ComboBox
        Private btnPlay1B As Forms.Button

        ' --- Set 2 Controls ---
        Private lblSet2 As Label
        Private cmbLanguage2 As Forms.ComboBox
        Private cmbVoice2A As Forms.ComboBox
        Private btnPlay2A As Forms.Button
        Private cmbVoice2B As Forms.ComboBox
        Private btnPlay2B As Forms.Button

        ' --- Sample text to play ---
        Private lblSampleText As Label
        Private txtSampleText As Forms.TextBox

        ' --- Bottom buttons ---
        Private btnOK As Forms.Button
        Private btnCancel As Forms.Button
        Private btnDesktop As Forms.Button

        ' --- For output path ---
        Private lblOutputPath As Label
        Private txtOutputPath As Forms.TextBox
        Private chkTemporary As Forms.CheckBox

        ' --- For storing voices from Google TTS ---
        ' This class helps us parse the JSON response from the voices API
        Private Class GoogleVoicesList
            <JsonProperty("voices")>
            Public Property Voices As List(Of GoogleVoice)
        End Class

        Private Class GoogleVoice
            <JsonProperty("name")>
            Public Property Name As String

            <JsonProperty("languageCodes")>
            Public Property LanguageCodes As List(Of String)

            <JsonProperty("ssmlGender")>
            Public Property SsmlGender As String
        End Class

        ' We can cache voices once retrieved for each language
        Private voiceCache As New Dictionary(Of String, List(Of GoogleVoice))()

        ' --- Dependencies / external references for Auth, etc. ---
        'Private _context As ISharedContext ' or your actual type
        'Private INI_OAuth2ClientMail As String
        'Private INI_OAuth2Scopes As String
        'Private INI_APIKey As String
        'Private INI_OAuth2Endpoint As String
        'Private INI_OAuth2ATExpiry As Long

        ' --- New parameters/fields for the amended form ---
        Private _twoVoicesRequired As Boolean
        Private _topLabelText As String
        Private _formTitle As String

        ' Radio buttons for voice selection.
        ' In one‐voice mode all four are in one group.
        ' In two‐voice mode we group each voice set separately (using Panels).
        Private rdoVoice1A As RadioButton, rdoVoice1B As RadioButton
        Private rdoVoice2A As RadioButton, rdoVoice2B As RadioButton
        Private pnlVoiceSet1 As Panel, pnlVoiceSet2 As Panel

        ' --- Public properties to return results ---
        ' In one‑voice mode SelectedVoices will contain one item;
        ' in two‑voice mode it will contain two items.
        Public Property SelectedVoices As List(Of String) = New List(Of String)()
        Public Property SelectedOutputPath As String = ""
        Public Property SelectedLanguage As String = ""


        Public Sub New(topLabelText As String,
               formTitle As String,
               twoVoicesRequired As Boolean)

            'context As ISharedContext,
            'clientMail As String,
            'scopes As String,
            'apiKey As String,
            'oauth2Endpoint As String,
            'oauth2Expiry As Long,

            ' Assign external parameters
            '_context = context
            'INI_OAuth2ClientMail = clientMail
            'INI_OAuth2Scopes = scopes
            'INI_APIKey = apiKey
            'INI_OAuth2Endpoint = oauth2Endpoint
            'INI_OAuth2ATExpiry = oauth2Expiry

            ' Store our extra parameters
            _topLabelText = topLabelText
            _formTitle = formTitle
            _twoVoicesRequired = twoVoicesRequired

            ' --- FORM PROPERTIES ---
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            Me.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            Me.Text = _formTitle
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.AutoSize = False
            Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Me.MinimumSize = New System.Drawing.Size(810, 480)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            Me.MaximizeBox = True

            Me.SuspendLayout()
            CreateControls()
            LayoutControls()
            Me.ResumeLayout()

            AddHandlers()

            Dim saved = My.Settings.TTSProvider
            If saved = "OpenAI" AndAlso TTS_openAIAvailable Then
                cmbEngine.SelectedItem = "OpenAI"
            ElseIf saved = "Google" AndAlso TTS_googleAvailable Then
                cmbEngine.SelectedItem = "Google"
            Else
                ' fall back to whichever is first in the list
                cmbEngine.SelectedIndex = 0
            End If

            PopulateLanguageComboBoxes()
            LoadSettingsAndVoices()

            txtSampleText.Text = If(
        String.IsNullOrEmpty(My.Settings.TTSSampleText),
        $"Hello, I am talking using {_formTitle}!",
        My.Settings.TTSSampleText
    )
        End Sub

        Private Sub CreateControls()
            ' --- Intro ---
            lblIntro = New System.Windows.Forms.Label() With {
        .Font = Me.Font,
        .Text = _topLabelText,
        .AutoSize = True,
        .MaximumSize = New System.Drawing.Size(700, 0)
    }

            ' --- Engine selector ---
            cmbEngine = New System.Windows.Forms.ComboBox() With {
    .Font = Me.Font,
    .DropDownStyle = ComboBoxStyle.DropDownList,
    .Width = 150,
    .Margin = New System.Windows.Forms.Padding(0, -4, 0, 10)
}
            cmbEngine.Items.Clear()
            If TTS_googleAvailable Then cmbEngine.Items.Add("Google")
            If TTS_openAIAvailable Then cmbEngine.Items.Add("OpenAI")
            ' default to first available
            cmbEngine.SelectedIndex = 0
            AddHandler cmbEngine.SelectedIndexChanged, AddressOf EngineChanged


            ' --- Voice Set 1 ---
            lblSet1 = New System.Windows.Forms.Label() With {.Font = Me.Font, .Text = "Your default voice set 1:", .AutoSize = True}
            cmbLanguage1 = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            cmbVoice1A = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            btnPlay1A = New System.Windows.Forms.Button() With {
        .Font = New System.Drawing.Font("Segoe UI Symbol", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point),
        .Text = "▶",
        .Size = New System.Drawing.Size(24, cmbVoice1A.PreferredHeight),
        .AutoSize = False
    }
            cmbVoice1B = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            btnPlay1B = New System.Windows.Forms.Button() With {
        .Font = New System.Drawing.Font("Segoe UI Symbol", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point),
        .Text = "▶",
        .Size = New System.Drawing.Size(24, cmbVoice1B.PreferredHeight),
        .AutoSize = False
    }

            ' --- Voice Set 2 ---
            lblSet2 = New System.Windows.Forms.Label() With {.Font = Me.Font, .Text = "Your default voice set 2:", .AutoSize = True}
            cmbLanguage2 = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            cmbVoice2A = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            btnPlay2A = New System.Windows.Forms.Button() With {
        .Font = New System.Drawing.Font("Segoe UI Symbol", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point),
        .Text = "▶",
        .Size = New System.Drawing.Size(24, cmbVoice2A.PreferredHeight),
        .AutoSize = False
    }
            cmbVoice2B = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            btnPlay2B = New System.Windows.Forms.Button() With {
        .Font = New System.Drawing.Font("Segoe UI Symbol", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point),
        .Text = "▶",
        .Size = New System.Drawing.Size(24, cmbVoice2B.PreferredHeight),
        .AutoSize = False
    }

            ' --- Radio Buttons ---
            If Not _twoVoicesRequired Then
                rdoVoice1A = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                rdoVoice1B = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                rdoVoice2A = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                rdoVoice2B = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                Select Case My.Settings.TTSLastRdoOneVoice
                    Case "Voice1A"
                        rdoVoice1A.Checked = True
                    Case "Voice1B"
                        rdoVoice1B.Checked = True
                    Case "Voice2A"
                        rdoVoice2A.Checked = True
                    Case "Voice2B"
                        rdoVoice2B.Checked = True
                    Case Else
                        rdoVoice1A.Checked = True ' Default if no previous selection
                End Select
            Else
                rdoVoice1A = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                rdoVoice2A = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                Select Case My.Settings.TTSLastRdoTwoVoices
                    Case "Voice1"
                        rdoVoice1A.Checked = True
                    Case "Voice2"
                        rdoVoice2A.Checked = True
                    Case Else
                        rdoVoice1A.Checked = True ' Default if no previous selection
                End Select
            End If

            ' --- Sample & Output rows ---
            lblSampleText = New System.Windows.Forms.Label() With {.Font = Me.Font, .Text = "Sample text:", .AutoSize = True}
            txtSampleText = New System.Windows.Forms.TextBox() With {.Font = Me.Font, .Width = 467}
            lblOutputPath = New System.Windows.Forms.Label() With {.Font = Me.Font, .Text = "Output (.mp3):", .AutoSize = True}
            txtOutputPath = New System.Windows.Forms.TextBox() With {.Font = Me.Font, .Width = 330}
            chkTemporary = New System.Windows.Forms.CheckBox() With {.Font = Me.Font, .Text = "Temp only", .AutoSize = True}

            ' --- Bottom Buttons ---
            btnOK = New System.Windows.Forms.Button() With {.Font = Me.Font, .Text = "OK", .AutoSize = True}
            btnCancel = New System.Windows.Forms.Button() With {.Font = Me.Font, .Text = "Cancel", .AutoSize = True}
            btnDesktop = New System.Windows.Forms.Button() With {.Font = Me.Font, .Text = "Save on Desktop", .AutoSize = True}

            ' --- Wire up mutual‑exclusion for all radios ---
            Dim radios As New List(Of System.Windows.Forms.RadioButton)
            For Each rb In New RadioButton() {rdoVoice1A, rdoVoice1B, rdoVoice2A, rdoVoice2B}
                If rb IsNot Nothing Then radios.Add(rb)
            Next
            For Each rb In radios
                AddHandler rb.CheckedChanged, Sub(s, e)
                                                  Dim meRb = DirectCast(s, RadioButton)
                                                  If meRb.Checked Then
                                                      For Each other In radios
                                                          If other IsNot meRb Then other.Checked = False
                                                      Next
                                                  End If
                                              End Sub
            Next
        End Sub

        Private Sub LayoutControls()

            Me.Controls.Clear()

            ' Root: 2 cols, 9 rows, bottom padding = 20px
            Dim root As New System.Windows.Forms.TableLayoutPanel() With {
                      .Dock = System.Windows.Forms.DockStyle.Fill,
                      .AutoSize = True,
                      .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                      .ColumnCount = 2,
                      .RowCount = 9,
                      .Padding = New System.Windows.Forms.Padding(10, 10, 10, 20)
                    }
            root.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            root.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            For i = 0 To 8
                root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Next

            ' Row0: Intro
            root.Controls.Add(lblIntro, 0, 0)
            root.SetColumnSpan(lblIntro, 2)

            ' Row1: Provider
            root.RowStyles.Insert(1, New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            root.Controls.Add(New Label() With {
                    .Font = Me.Font,
                    .Text = "Text-to-Speech Provider:",
                    .AutoSize = True
                }, 0, 1)
            root.Controls.Add(cmbEngine, 1, 1)


            ' Row1: Headings
            root.Controls.Add(lblSet1, 0, 2)
            root.Controls.Add(lblSet2, 1, 2)

            ' Row2: Language (+ two‑voice radio)
            Dim fl2a As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then fl2a.Controls.Add(rdoVoice1A)
            fl2a.Controls.Add(cmbLanguage1)
            root.Controls.Add(fl2a, 0, 3)

            Dim fl2b As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then fl2b.Controls.Add(rdoVoice2A)
            fl2b.Controls.Add(cmbLanguage2)
            root.Controls.Add(fl2b, 1, 3)

            ' Row3: Voice1A + play + single‑voice radio or indent if two‑voice mode
            Dim fl3a As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then
                ' indent by the radio’s width so it lines up with language
                fl3a.Padding = New System.Windows.Forms.Padding(rdoVoice1A.PreferredSize.Width, 0, 0, 0)
            Else
                rdoVoice1A.Margin = New System.Windows.Forms.Padding(0,
        (cmbVoice1A.PreferredHeight - rdoVoice1A.PreferredSize.Height) \ 2, 0, 0)
                fl3a.Controls.Add(rdoVoice1A)
            End If
            fl3a.Controls.Add(cmbVoice1A)
            fl3a.Controls.Add(btnPlay1A)
            root.Controls.Add(fl3a, 0, 4)

            Dim fl3b As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then
                fl3b.Padding = New System.Windows.Forms.Padding(rdoVoice2A.PreferredSize.Width, 0, 0, 0)
            Else
                rdoVoice2A.Margin = New System.Windows.Forms.Padding(0,
        (cmbVoice2A.PreferredHeight - rdoVoice2A.PreferredSize.Height) \ 2, 0, 0)
                fl3b.Controls.Add(rdoVoice2A)
            End If
            fl3b.Controls.Add(cmbVoice2A)
            fl3b.Controls.Add(btnPlay2A)
            root.Controls.Add(fl3b, 1, 4)

            ' Row4: Voice1B + play (same indent logic)
            Dim fl4a As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then
                fl4a.Padding = New System.Windows.Forms.Padding(rdoVoice1A.PreferredSize.Width, 0, 0, 0)
            Else
                rdoVoice1B.Margin = New System.Windows.Forms.Padding(0,
        (cmbVoice1B.PreferredHeight - rdoVoice1B.PreferredSize.Height) \ 2, 0, 0)
                fl4a.Controls.Add(rdoVoice1B)
            End If
            fl4a.Controls.Add(cmbVoice1B)
            fl4a.Controls.Add(btnPlay1B)
            root.Controls.Add(fl4a, 0, 5)

            Dim fl4b As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then
                fl4b.Padding = New System.Windows.Forms.Padding(rdoVoice2A.PreferredSize.Width, 0, 0, 0)
            Else
                rdoVoice2B.Margin = New System.Windows.Forms.Padding(0,
        (cmbVoice2B.PreferredHeight - rdoVoice2B.PreferredSize.Height) \ 2, 0, 0)
                fl4b.Controls.Add(rdoVoice2B)
            End If
            fl4b.Controls.Add(cmbVoice2B)
            fl4b.Controls.Add(btnPlay2B)
            root.Controls.Add(fl4b, 1, 5)

            ' Row5: Sample text (2‑col table for vertical centering)
            Dim tbl5 As New System.Windows.Forms.TableLayoutPanel() With {
      .ColumnCount = 2,
      .RowCount = 1,
      .AutoSize = True,
      .Dock = System.Windows.Forms.DockStyle.Top
    }
            tbl5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            tbl5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            lblSampleText.Dock = System.Windows.Forms.DockStyle.Fill
            lblSampleText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            txtSampleText.Dock = System.Windows.Forms.DockStyle.Fill
            tbl5.Controls.Add(lblSampleText, 0, 0)
            tbl5.Controls.Add(txtSampleText, 1, 0)
            root.Controls.Add(tbl5, 0, 6)
            root.SetColumnSpan(tbl5, 2)

            ' Row6: Output path (3‑col table)
            Dim tbl6 As New System.Windows.Forms.TableLayoutPanel() With {
      .ColumnCount = 3,
      .RowCount = 1,
      .AutoSize = True,
      .Dock = System.Windows.Forms.DockStyle.Top
    }
            tbl6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            tbl6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            tbl6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            lblOutputPath.Dock = System.Windows.Forms.DockStyle.Fill
            lblOutputPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            txtOutputPath.Dock = System.Windows.Forms.DockStyle.Fill
            chkTemporary.Anchor = System.Windows.Forms.AnchorStyles.Left
            tbl6.Controls.Add(lblOutputPath, 0, 0)
            tbl6.Controls.Add(txtOutputPath, 1, 0)
            tbl6.Controls.Add(chkTemporary, 2, 0)
            root.Controls.Add(tbl6, 0, 7)
            root.SetColumnSpan(tbl6, 2)

            Dim pnlButtons As New System.Windows.Forms.FlowLayoutPanel() With {
                  .Dock = System.Windows.Forms.DockStyle.Bottom,
                  .AutoSize = True,
                  .Padding = New Padding(10),
                  .FlowDirection = FlowDirection.LeftToRight
                }
            pnlButtons.Controls.Add(btnOK)
            pnlButtons.Controls.Add(btnCancel)
            pnlButtons.Controls.Add(btnDesktop)

            Me.Controls.Clear()
            Me.Controls.Add(pnlButtons)
            Me.Controls.Add(root)

        End Sub

        Private Sub EngineChanged(sender As Object, e As EventArgs)
            ' set our global
            TTS_SelectedEngine = If(cmbEngine.SelectedItem.ToString() = "OpenAI",
                             TTSEngine.OpenAI,
                             TTSEngine.Google)

            My.Settings.TTSProvider = cmbEngine.SelectedItem.ToString()
            My.Settings.Save()

            ' rebuild the combos
            PopulateLanguageComboBoxes()
            LoadSettingsAndVoices()
        End Sub

        Private Sub AddHandlers()
            AddHandler cmbLanguage1.SelectedIndexChanged, AddressOf cmbLanguage1_SelectedIndexChanged
            AddHandler cmbLanguage2.SelectedIndexChanged, AddressOf cmbLanguage2_SelectedIndexChanged

            AddHandler btnPlay1A.Click, AddressOf btnPlay1A_Click
            AddHandler btnPlay1B.Click, AddressOf btnPlay1B_Click
            AddHandler btnPlay2A.Click, AddressOf btnPlay2A_Click
            AddHandler btnPlay2B.Click, AddressOf btnPlay2B_Click

            AddHandler btnOK.Click, AddressOf btnOK_Click
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            AddHandler btnDesktop.Click, AddressOf btnDesktop_Click
        End Sub

        Private Sub PopulateLanguageComboBoxes()
            cmbLanguage1.Items.Clear()
            cmbLanguage2.Items.Clear()

            If TTS_SelectedEngine = TTSEngine.OpenAI Then
                cmbLanguage1.Items.AddRange(OpenAILanguages)
                cmbLanguage2.Items.AddRange(OpenAILanguages)
            Else
                For Each lang In GoogleTTSsupportedLanguages
                    cmbLanguage1.Items.Add(lang)
                    cmbLanguage2.Items.Add(lang)
                Next
            End If

            If cmbLanguage1.Items.Count > 0 Then cmbLanguage1.SelectedIndex = 0
            If cmbLanguage2.Items.Count > 0 Then cmbLanguage2.SelectedIndex = 0
        End Sub


        Private Async Sub LoadSettingsAndVoices()
            RemoveHandler cmbLanguage1.SelectedIndexChanged, AddressOf cmbLanguage1_SelectedIndexChanged
            RemoveHandler cmbLanguage2.SelectedIndexChanged, AddressOf cmbLanguage2_SelectedIndexChanged

            ' restore last‐used languages
            cmbLanguage1.SelectedItem = My.Settings.TTS1languagecode
            cmbLanguage2.SelectedItem = My.Settings.TTS2languagecode

            Dim tasks As New List(Of System.Threading.Tasks.Task)

            If TTS_SelectedEngine = TTSEngine.OpenAI Then
                ' immediate, sync fill of voices
                PopulateOpenAIVoices(cmbLanguage1.Text, cmbVoice1A, cmbVoice1B)
                PopulateOpenAIVoices(cmbLanguage2.Text, cmbVoice2A, cmbVoice2B)
            Else
                ' Google: async fetch
                If Not String.IsNullOrEmpty(cmbLanguage1.Text) Then
                    tasks.Add(LoadVoicesIntoComboBoxesAsync(cmbLanguage1.Text, cmbVoice1A, cmbVoice1B))
                End If
                If Not String.IsNullOrEmpty(cmbLanguage2.Text) Then
                    tasks.Add(LoadVoicesIntoComboBoxesAsync(cmbLanguage2.Text, cmbVoice2A, cmbVoice2B))
                End If
            End If

            If tasks.Count > 0 Then Await System.Threading.Tasks.Task.WhenAll(tasks)

            ' restore last‐used voice selections
            cmbVoice1A.SelectedItem = My.Settings.TTS1voiceA
            cmbVoice1B.SelectedItem = My.Settings.TTS1voiceB
            cmbVoice2A.SelectedItem = My.Settings.TTS2voiceA
            cmbVoice2B.SelectedItem = My.Settings.TTS2voiceB

            AddHandler cmbLanguage1.SelectedIndexChanged, AddressOf cmbLanguage1_SelectedIndexChanged
            AddHandler cmbLanguage2.SelectedIndexChanged, AddressOf cmbLanguage2_SelectedIndexChanged
        End Sub


        Private Async Sub xxxLoadSettingsAndVoices()
            RemoveHandler cmbLanguage1.SelectedIndexChanged, AddressOf cmbLanguage1_SelectedIndexChanged
            RemoveHandler cmbLanguage2.SelectedIndexChanged, AddressOf cmbLanguage2_SelectedIndexChanged

            cmbLanguage1.SelectedItem = If(IsNothing(My.Settings.TTS1languagecode), "", My.Settings.TTS1languagecode)
            cmbLanguage2.SelectedItem = If(IsNothing(My.Settings.TTS2languagecode), "", My.Settings.TTS2languagecode)

            Dim tasks As New List(Of System.Threading.Tasks.Task)
            If TTS_SelectedEngine = TTSEngine.OpenAI Then
                ' OpenAI: no async fetch needed
                PopulateOpenAIVoices(cmbLanguage1.Text, cmbVoice1A, cmbVoice1B)
                PopulateOpenAIVoices(cmbLanguage2.Text, cmbVoice2A, cmbVoice2B)
            Else
                If Not IsNothing(cmbLanguage1.SelectedItem) AndAlso cmbLanguage1.Text <> "" Then
                    tasks.Add(LoadVoicesIntoComboBoxesAsync(cmbLanguage1.SelectedItem.ToString(), cmbVoice1A, cmbVoice1B))
                End If

                If Not IsNothing(cmbLanguage2.SelectedItem) AndAlso cmbLanguage2.Text <> "" Then
                    tasks.Add(LoadVoicesIntoComboBoxesAsync(cmbLanguage2.SelectedItem.ToString(), cmbVoice2A, cmbVoice2B))
                End If
            End If
            If tasks.Count > 0 Then
                Await System.Threading.Tasks.Task.WhenAll(tasks)
            End If

            ' Now, set the selected voices after both lists are populated
            If Not IsNothing(cmbLanguage1.SelectedItem) AndAlso cmbLanguage1.Text <> "" Then
                cmbVoice1A.SelectedItem = If(IsNothing(My.Settings.TTS1voiceA), "", My.Settings.TTS1voiceA)
                cmbVoice1B.SelectedItem = If(IsNothing(My.Settings.TTS1voiceB), "", My.Settings.TTS1voiceB)
            End If

            If Not IsNothing(cmbLanguage2.SelectedItem) AndAlso cmbLanguage2.Text <> "" Then
                cmbVoice2A.SelectedItem = If(IsNothing(My.Settings.TTS2voiceA), "", My.Settings.TTS2voiceA)
                cmbVoice2B.SelectedItem = If(IsNothing(My.Settings.TTS2voiceB), "", My.Settings.TTS2voiceB)
            End If

            AddHandler cmbLanguage1.SelectedIndexChanged, AddressOf cmbLanguage1_SelectedIndexChanged
            AddHandler cmbLanguage2.SelectedIndexChanged, AddressOf cmbLanguage2_SelectedIndexChanged

        End Sub


        Private Sub PopulateOpenAIVoices(lang As String,
                                 comboA As Forms.ComboBox,
                                 comboB As Forms.ComboBox)
            comboA.Items.Clear() : comboB.Items.Clear()
            For Each v In OpenAIVoices
                Dim disp = $"{v} — {OpenAIDescriptions(v)}"
                comboA.Items.Add(disp)
                comboB.Items.Add(disp)
            Next
            If comboA.Items.Count > 0 Then comboA.SelectedIndex = 0
            If comboB.Items.Count > 0 Then comboB.SelectedIndex = 0
        End Sub

        Private Async Sub cmbLanguage1_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim lang = TryCast(cmbLanguage1.SelectedItem, String)
            If String.IsNullOrEmpty(lang) Then Return

            If TTS_SelectedEngine = TTSEngine.OpenAI Then
                PopulateOpenAIVoices(lang, cmbVoice1A, cmbVoice1B)
            Else
                Await LoadVoicesIntoComboBoxesAsync(lang, cmbVoice1A, cmbVoice1B)
            End If
        End Sub

        Private Async Sub cmbLanguage2_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim lang = TryCast(cmbLanguage2.SelectedItem, String)
            If String.IsNullOrEmpty(lang) Then Return

            If TTS_SelectedEngine = TTSEngine.OpenAI Then
                PopulateOpenAIVoices(lang, cmbVoice2A, cmbVoice2B)
            Else
                Await LoadVoicesIntoComboBoxesAsync(lang, cmbVoice2A, cmbVoice2B)
            End If
        End Sub



        Private Async Sub xxcmbLanguage1_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim selectedLang As String = TryCast(cmbLanguage1.SelectedItem, String)
            If Not String.IsNullOrEmpty(selectedLang) Then
                Await LoadVoicesIntoComboBoxesAsync(selectedLang, cmbVoice1A, cmbVoice1B)
            End If
        End Sub

        Private Async Sub xxcmbLanguage2_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim selectedLang As String = TryCast(cmbLanguage2.SelectedItem, String)
            If Not String.IsNullOrEmpty(selectedLang) Then
                Await LoadVoicesIntoComboBoxesAsync(selectedLang, cmbVoice2A, cmbVoice2B)
            End If
        End Sub


        Private Async Function LoadVoicesIntoComboBoxesAsync(languageCode As String,
                                                           comboA As Forms.ComboBox,
                                                           comboB As Forms.ComboBox) As System.Threading.Tasks.Task
            Try
                Dim voicesForLang As List(Of GoogleVoice) = Await GetVoicesByLanguageAsync(languageCode)
                comboA.Items.Clear()
                comboB.Items.Clear()

                For Each v In voicesForLang
                    Dim displayName As String = $"{v.Name} ({v.SsmlGender.ToLower()})"
                    comboA.Items.Add(displayName)
                    comboB.Items.Add(displayName)
                Next
            Catch ex As System.Exception
                ShowCustomMessageBox("When trying to load the voices from the Google server, an error occurred: " & ex.Message)
            End Try
        End Function

        Private Async Function GetVoicesByLanguageAsync(languageCode As String) As Threading.Tasks.Task(Of List(Of GoogleVoice))
            If voiceCache.ContainsKey(languageCode) Then
                Return voiceCache(languageCode)
            End If

            Dim AccessToken As String = Await GetFreshTTSToken(UseSecondaryFor(TTSEngine.Google))
            If String.IsNullOrEmpty(AccessToken) Then
                ShowCustomMessageBox("Error accessing Google API - authentication failed (no token).")
                Return Nothing
            End If

            ' Build request
            Dim url As String = TTS_GoogleEndpoint & "voices?languageCode=" & languageCode
            Using httpClient As New HttpClient()
                httpClient.DefaultRequestHeaders.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken)

                Dim response As HttpResponseMessage = Await httpClient.GetAsync(url)
                If response.IsSuccessStatusCode Then
                    Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                    Dim voicesList As GoogleVoicesList = JsonConvert.DeserializeObject(Of GoogleVoicesList)(responseContent)

                    If voicesList IsNot Nothing AndAlso voicesList.Voices IsNot Nothing Then
                        voiceCache(languageCode) = voicesList.Voices
                        Return voicesList.Voices
                    Else
                        Return New List(Of GoogleVoice)()
                    End If
                Else
                    ShowCustomMessageBox("Failed to retrieve voices: " & response.StatusCode.ToString())
                    Return Nothing
                End If
            End Using
        End Function

        ' --- Play button event handlers ---
        Private Async Sub btnPlay1A_Click(sender As Object, e As EventArgs)
            Await PlaySelectedVoiceAsync(cmbLanguage1, cmbVoice1A)
        End Sub

        Private Async Sub btnPlay1B_Click(sender As Object, e As EventArgs)
            Await PlaySelectedVoiceAsync(cmbLanguage1, cmbVoice1B)
        End Sub

        Private Async Sub btnPlay2A_Click(sender As Object, e As EventArgs)
            Await PlaySelectedVoiceAsync(cmbLanguage2, cmbVoice2A)
        End Sub

        Private Async Sub btnPlay2B_Click(sender As Object, e As EventArgs)
            Await PlaySelectedVoiceAsync(cmbLanguage2, cmbVoice2B)
        End Sub

        Private Async Function PlaySelectedVoiceAsync(cmbLang As Forms.ComboBox, cmbVoice As Forms.ComboBox) As Threading.Tasks.Task
            Try
                Dim lang As String = TryCast(cmbLang.SelectedItem, String)
                Dim voiceName As String = TryCast(cmbVoice.SelectedItem, String)
                Dim sampleText As String = txtSampleText.Text

                If String.IsNullOrEmpty(lang) OrElse String.IsNullOrEmpty(voiceName) Then
                    ShowCustomMessageBox("Please select both language and voice before playing.")
                    Return
                End If
                voiceName = voiceName.Replace(" (male)", "").Replace(" (female)", "")
                If TTS_SelectedEngine = TTSEngine.OpenAI Then
                    ' remove “ — Beschreibung”
                    voiceName = voiceName.Split(" "c)(0)
                End If
                Await Threading.Tasks.Task.Run(Sub()
                                                   GenerateAndPlayAudio(sampleText, "", lang, voiceName)
                                               End Sub)
            Catch ex As System.Exception
                ShowCustomMessageBox("When trying to play the voice, an error occurred: " & ex.Message)
            End Try
        End Function

        ' --- OK / Cancel / Desktop event handlers ---
        Private Sub btnOK_Click(sender As Object, e As EventArgs)

            TTS_SelectedEngine = If(cmbEngine.SelectedItem.ToString() = "OpenAI",
                         TTSEngine.OpenAI,
                         TTSEngine.Google)

            My.Settings.TTSProvider = cmbEngine.SelectedItem.ToString()
            My.Settings.Save()

            Dim NotAllSelected As Boolean = False

            ' Determine which voice(s) were selected based on radio buttons
            SelectedVoices.Clear()
            If Not _twoVoicesRequired Then
                ' ONE VOICE mode: the four radio buttons are one group.

                If rdoVoice1A.Checked Then
                    If cmbVoice1A.SelectedItem IsNot Nothing AndAlso cmbVoice1A.SelectedItem.ToString() <> "" Then
                        Dim sel As String = cmbVoice1A.SelectedItem.ToString()
                        If TTS_SelectedEngine = TTSEngine.OpenAI Then
                            ' drop the “ — Beschreibung” part
                            sel = sel.Split(" "c)(0)
                        End If
                        SelectedVoices.Add(sel)
                        SelectedLanguage = cmbLanguage1.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                ElseIf rdoVoice1B.Checked Then
                    If cmbVoice1B.SelectedItem IsNot Nothing AndAlso cmbVoice1B.SelectedItem.ToString() <> "" Then

                        Dim sel As String = cmbVoice1B.SelectedItem.ToString()
                        If TTS_SelectedEngine = TTSEngine.OpenAI Then
                            ' drop the “ — Beschreibung” part
                            sel = sel.Split(" "c)(0)
                        End If
                        SelectedVoices.Add(sel)

                        SelectedLanguage = cmbLanguage1.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                ElseIf rdoVoice2A.Checked Then
                    If cmbVoice2A.SelectedItem IsNot Nothing AndAlso cmbVoice2A.SelectedItem.ToString() <> "" Then

                        Dim sel As String = cmbVoice2A.SelectedItem.ToString()
                        If TTS_SelectedEngine = TTSEngine.OpenAI Then
                            ' drop the “ — Beschreibung” part
                            sel = sel.Split(" "c)(0)
                        End If
                        SelectedVoices.Add(sel)

                        SelectedLanguage = cmbLanguage2.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                ElseIf rdoVoice2B.Checked Then
                    If cmbVoice2B.SelectedItem IsNot Nothing AndAlso cmbVoice2B.SelectedItem.ToString() <> "" Then
                        Dim sel As String = cmbVoice2B.SelectedItem.ToString()
                        If TTS_SelectedEngine = TTSEngine.OpenAI Then
                            ' drop the “ — Beschreibung” part
                            sel = sel.Split(" "c)(0)
                        End If
                        SelectedVoices.Add(sel)

                        SelectedLanguage = cmbLanguage2.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                End If
            Else
                ' TWO VOICES mode: one voice from each set.
                If rdoVoice1A.Checked Then
                    If cmbVoice1A.SelectedItem IsNot Nothing AndAlso cmbVoice1A.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice1A.SelectedItem.ToString())
                        SelectedLanguage = cmbLanguage1.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                    If cmbVoice1B.SelectedItem IsNot Nothing AndAlso cmbVoice1B.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice1B.SelectedItem.ToString())
                        SelectedLanguage = cmbLanguage1.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                ElseIf rdoVoice2A.Checked Then
                    If cmbVoice2A.SelectedItem IsNot Nothing AndAlso cmbVoice2A.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice2A.SelectedItem.ToString())
                        SelectedLanguage = cmbLanguage2.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                    If cmbVoice2B.SelectedItem IsNot Nothing AndAlso cmbVoice2B.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice2B.SelectedItem.ToString())
                        SelectedLanguage = cmbLanguage2.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                End If
            End If

            If NotAllSelected Then
                ShowCustomMessageBox("Please complete your voice selection (Or 'Cancel').")
                Return
            End If

            ' Save selected radio button (for one-voice mode)
            If Not _twoVoicesRequired Then
                If rdoVoice1A.Checked Then
                    My.Settings.TTSLastRdoOneVoice = "Voice1A"
                ElseIf rdoVoice1B.Checked Then
                    My.Settings.TTSLastRdoOneVoice = "Voice1B"
                ElseIf rdoVoice2A.Checked Then
                    My.Settings.TTSLastRdoOneVoice = "Voice2A"
                ElseIf rdoVoice2B.Checked Then
                    My.Settings.TTSLastRdoOneVoice = "Voice2B"
                End If
            Else
                ' Save selected radio button (for two-voices mode)
                If rdoVoice1A.Checked Then
                    My.Settings.TTSLastRdoTwoVoices = "Voice1"
                ElseIf rdoVoice2A.Checked Then
                    My.Settings.TTSLastRdoTwoVoices = "Voice2"
                End If
            End If
            ' Save settings as before
            My.Settings.Save()

            ' Determine output path: if Temporary is checked, return blank.

            SelectedOutputPath = txtOutputPath.Text

            If String.IsNullOrWhiteSpace(SelectedOutputPath) Then
                ' Use default path (Desktop) with default filename
                SelectedOutputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
            ElseIf SelectedOutputPath.EndsWith("\") OrElse SelectedOutputPath.EndsWith("/") Then
                ' If only a folder is given, append default filename
                SelectedOutputPath = Path.Combine(SelectedOutputPath, TTSDefaultFile)
            Else
                Dim dir As String = Path.GetDirectoryName(SelectedOutputPath)
                Dim fileName As String = Path.GetFileName(SelectedOutputPath)

                ' If no directory is found, assume Desktop as the base
                If String.IsNullOrWhiteSpace(dir) Then
                    SelectedOutputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName)
                    dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                End If

                ' If no filename is given, use the default filename
                If String.IsNullOrWhiteSpace(fileName) Then
                    SelectedOutputPath = Path.Combine(dir, TTSDefaultFile)
                End If

                ' Ensure the filename has ".mp3" extension
                If Not fileName.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) Then
                    SelectedOutputPath = Path.Combine(dir, fileName & ".mp3")
                End If
            End If

            ' Update the TextBox with the corrected path
            txtOutputPath.Text = SelectedOutputPath

            txtOutputPath.Text = SelectedOutputPath

            SaveSettings()

            If chkTemporary.Checked Then
                SelectedOutputPath = ""
            End If

            Me.DialogResult = DialogResult.OK
            Me.Close()

        End Sub

        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            ' If cancelled, clear any voice selection.
            SelectedVoices.Clear()
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        End Sub

        Private Sub SaveSettings()
            My.Settings.TTS1languagecode = If(cmbLanguage1.SelectedItem?.ToString(), "")
            My.Settings.TTS1voiceA = If(cmbVoice1A.SelectedItem?.ToString(), "")
            My.Settings.TTS1voiceB = If(cmbVoice1B.SelectedItem?.ToString(), "")
            My.Settings.TTS2languagecode = If(cmbLanguage2.SelectedItem?.ToString(), "")
            My.Settings.TTS2voiceA = If(cmbVoice2A.SelectedItem?.ToString(), "")
            My.Settings.TTS2voiceB = If(cmbVoice2B.SelectedItem?.ToString(), "")
            My.Settings.TTSSampleText = If(txtSampleText.Text, "")
            My.Settings.TTSOutputPath = txtOutputPath.Text
            My.Settings.Save()
        End Sub

        ' --- chkTemporary CheckedChanged handler ---
        Private Sub chkTemporary_CheckedChanged(sender As Object, e As EventArgs)
            txtOutputPath.Enabled = Not chkTemporary.Checked
        End Sub

        Private Sub btnDesktop_Click(sender As Object, e As EventArgs)
            ' Get the filename
            Dim fileName As String = Path.GetFileName(txtOutputPath.Text)

            ' Get the user's Desktop path
            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

            ' Construct new file path
            txtOutputPath.Text = Path.Combine(desktopPath, fileName)

        End Sub

    End Class


End Class

