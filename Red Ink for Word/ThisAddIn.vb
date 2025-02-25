' Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See License.txt or https://vischer.com/redink for more information.
'
' 25.2.2025
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
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.

Option Explicit On

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Net
Imports HtmlAgilityPack
Imports System.Net.Http
Imports System.Security.Policy
Imports System.Threading
Imports Markdig
Imports System.Drawing
Imports System.Diagnostics
Imports System.Data
Imports NAudio.Wave
Imports Vosk
Imports Newtonsoft.Json.Linq
Imports NAudio.Wave.SampleProviders
Imports NAudio.CoreAudioApi
Imports Whisper.net
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows
Imports System.Speech.Synthesis
Imports Whisper.net.LibraryLoader
Imports Newtonsoft.Json


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

    Private mainThreadControl As New System.Windows.Forms.Control()
    Public StartupInitialized As Boolean = False
    Private WithEvents wordApp As Word.Application

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        wordApp = Application
        Try
            If wordApp IsNot Nothing Then
                AddHandler wordApp.WindowActivate, AddressOf WordApp_WindowActivate
                AddHandler wordApp.DocumentOpen, AddressOf WordApp_DocumentOpen
                AddHandler wordApp.NewDocument, AddressOf WordApp_NewDocument
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

    Public Const Version As String = "V.250225 Gen2 Beta Test"

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
    Private Const InPlacePrefix As String = "Replace:"
    Private Const MarkupPrefix As String = "Markup:"
    Private Const MarkupPrefixDiff As String = "MarkupDiff:"
    Private Const MarkupPrefixDiffW As String = "MarkupDiffW:"
    Private Const MarkupPrefixWord As String = "MarkupWord:"
    Private Const MarkupPrefixRegex As String = "MarkupRegex:"
    Private Const MarkupPrefixAll As String = "Markup[Diff|DiffW|Word|Regex]:"
    Private Const ClipboardPrefix As String = "Clipboard:"
    Private Const ClipboardPrefix2 As String = "Clip:"
    Private Const BubblesPrefix As String = "Bubbles:"
    Private Const BubbleCutText As String = " (" & ChrW(&H2702) & ")"
    Private Const SearchAllTrigger As String = "(full)"
    Private Const SearchMultiTrigger As String = "All:"

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

    Public Shared DragDropFormLabel As String = ""
    Public Shared DragDropFormFilter As String = ""

    Public Shared TTSDefaultFile As String = $"{AN2}-output.mp3"
    Public Const TTSLargeText As Integer = 2500
    Public Shared hostTags As String() = {"H:", "Host:", "A:", "1:"}
    Public Shared guestTags As String() = {"G:", "Guest:", "Gast:", "B:", "2:"}
    Public Shared GoogleIdentifier As String = "googleapis.com"
    Public Shared TTSSecondAPI As Boolean = False

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

    Public Shared Property INI_Response As String
        Get
            Return _context.INI_Response
        End Get
        Set(value As String)
            _context.INI_Response = value
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

    Public Shared Property INI_Response_2 As String
        Get
            Return _context.INI_Response_2
        End Get
        Set(value As String)
            _context.INI_Response_2 = value
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
    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal Hidesplash As Boolean = False, Optional ByVal AddUserPrompt As String = "") As Task(Of String)
        Return Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, Hidesplash, AddUserPrompt)
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

        If MenusAdded Then Exit Sub

        ' Remove existing context menus from relevant context menus
        If RemoveMenu Then
            RemoveOldContextMenu()
            RemoveMenu = False
        End If

        If Not INI_ContextMenu Then Exit Sub

        If Not VBAModuleWorking() Then Exit Sub

        If INIloaded = False Then Exit Sub

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
                    shortcutDict(Trim(shortcutPair(0))) = Trim(shortcutPair(1))
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
                Exit Sub ' No shortcut assigned
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
        End If

        Dim TranscriptionForm = New TranscriptionForm()
        TranscriptionForm.Show()
    End Sub

    Public Sub ShowChatForm()
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

    Public Async Sub InLanguage1()
        If INILoadFail() Then Exit Sub
        TranslateLanguage = INI_Language1
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, False, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub InLanguage2()
        If INILoadFail() Then Exit Sub
        TranslateLanguage = INI_Language2
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, False, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub InOther()
        If INILoadFail() Then Exit Sub
        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, False, False, False, True, False, INI_KeepFormatCap)
        End If
    End Sub

    Public Async Sub Correct()
        If INILoadFail() Then Exit Sub
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Correct), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Improve()
        If INILoadFail() Then Exit Sub
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Improve), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Friendly()
        If INILoadFail() Then Exit Sub
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Friendly), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Convincing()
        If INILoadFail() Then Exit Sub
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Convincing), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub NoFillers()
        If INILoadFail() Then Exit Sub
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_NoFillers), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Anonymize()
        If INILoadFail() Then Exit Sub

        Dim DoMarkup As Boolean = INI_DoMarkupWord
        Dim DoReplace As Boolean = INI_ReplaceText2
        If Not DoMarkup Or Not DoReplace Then
            Dim result2 As Integer = ShowCustomYesNoBox($"As per your current settings no markup will be applied. For anonymizing a larger text, doing a markup may be a better choice. How do you want to continue?", "Continue as is", "Continue with a markup")
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
        If INILoadFail() Then Exit Sub
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Explain), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, True, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub SuggestTitles()
        If INILoadFail() Then Exit Sub
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_SuggestTitles), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, True, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub Shorten()

        If INILoadFail() Then Exit Sub
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection

        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
        End If

        Dim Textlength As Integer = GetSelectedTextLength()
        Dim UserInput As String
        Dim ShortenPercentValue As Integer = 0
        Do
            UserInput = Trim(SLib.ShowCustomInputBox("Enter the percentage by which your text should be shortened (it has " & Textlength & " words; " & ShortenPercent & "% will cut approx. " & (Textlength * ShortenPercent / 100) & " words)", $"{AN} Shortener", True, CStr(ShortenPercent) & "%"))
            If String.IsNullOrEmpty(UserInput) Then
                Exit Sub
            End If
            UserInput = UserInput.Replace("%", "").Trim()
            If Integer.TryParse(UserInput, ShortenPercentValue) AndAlso ShortenPercentValue >= 1 AndAlso ShortenPercentValue <= 99 Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid percentage between 1 And 99.")
            End If
        Loop
        If ShortenPercentValue = 0 Then Exit Sub
        ShortenLength = (Textlength - (Textlength * (100 - ShortenPercentValue) / 100))
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Improve), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub
    Public Async Sub SwitchParty()
        If INILoadFail() Then Exit Sub
        Dim UserInput As String
        Do
            UserInput = Trim(SLib.ShowCustomInputBox("Please provide the original party name And the New party name, separated by a comma (example: Elvis Presley, Taylor Swift):", $"{AN} Switch Party", True))

            If String.IsNullOrEmpty(UserInput) Then
                Exit Sub
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
        If INILoadFail() Then Exit Sub
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection

        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
        End If

        Dim Textlength As Integer = GetSelectedTextLength()

        Dim UserInput As String
        SummaryLength = 0

        Do
            UserInput = Trim(SLib.ShowCustomInputBox("Enter the number of words your summary shall have (the selected text has " & Textlength & " words; the proposal " & SummaryPercent & "%):", $"{AN} Summarizer", True, CStr(CInt(SummaryPercent * Textlength / 100))))

            If String.IsNullOrEmpty(UserInput) Then
                Exit Sub
            End If

            If Integer.TryParse(UserInput, SummaryLength) AndAlso SummaryLength >= 1 AndAlso SummaryLength <= Textlength Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid word count between 1 and " & Textlength & ".")
            End If
        Loop
        If SummaryLength = 0 Then Exit Sub

        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Summarize), False, False, False, False, False, False, True, False, True, False, 0)
    End Sub

    Public Async Sub CreatePodcast()
        If INILoadFail() Then Exit Sub
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection

        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
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
                    New SLib.InputParameter("Duration", Duration),
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
        If INILoadFail() Then Exit Sub
        Dim Endpoint As String = INI_Endpoint
        Dim Endpoint_2 As String = INI_Endpoint_2
        Dim TTSEndpoint As String = INI_TTSEndpoint
        If Endpoint.Contains(GoogleIdentifier) Then
            If Endpoint.Contains(GoogleIdentifier) Then
                TTSSecondAPI = False
            ElseIf Endpoint_2.Contains(GoogleIdentifier) Then
                TTSSecondAPI = True
            Else
                Exit Sub
            End If
        End If

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Selection = application.Selection
        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
        End If
        SelectedText = Trim(selection.Text)
        If SelectedText.Contains("H: ") And SelectedText.Contains("G: ") Then
            ReadPodcast(SelectedText)
        Else
            If Trim(selection.Text).StartsWith("{") Then
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
                Exit Sub
            Else
                Using frm As New TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voice you wish to use for creating your audio file and configure where to save it.", $"{AN} Google Text-to-Speech - Select Voices", False)
                    If frm.ShowDialog() = DialogResult.OK Then
                        Dim selectedVoices As List(Of String) = frm.SelectedVoices
                        Dim selectedLanguage As String = frm.SelectedLanguage
                        Dim outputPath As String = frm.SelectedOutputPath
                        GenerateAndPlayAudioFromSelectionParagraphs(outputPath, selectedLanguage, selectedVoices(0).Replace(" (male)", "").Replace(" (female)", ""))
                    End If
                End Using
            End If
        End If
    End Sub
    Public Async Sub FreeStyleNM()
        If INILoadFail() Then Exit Sub
        FreeStyle(False)
    End Sub
    Public Async Sub FreeStyleAM()
        If INILoadFail() Then Exit Sub
        FreeStyle(True)
    End Sub
    Public Async Sub FreeStyle(UseSecondAPI)

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

            Dim MarkupInstruct As String = $"start with '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"with '{InPlacePrefix}' for replacing the selection"
            Dim BubblesInstruct As String = $"with '{BubblesPrefix}' for having your text commented"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}' for separate output"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim ExtInstruct As String = $"; inlcude '{ExtTrigger}' for text of a file (txt, docx, pdf)"
            Dim TPMarkupInstruct As String = $"; add '{TPMarkupTriggerInstruct}' if revisions [of user] should be pointed out to the LLM"
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}' for overriding formatting defaults"
            Dim LibInstruct As String = $"; add '{LibTrigger}' for library search"
            Dim NetInstruct As String = $"; add '{NetTrigger}' for internet search"
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; ctrl-v for your last prompt")

            Dim AddOnInstruct As String = TPMarkupInstruct
            AddOnInstruct += NoFormatInstruct.Replace("; add", ", ")
            If INI_Lib Then
                AddOnInstruct += LibInstruct.Replace("; add", ",")
            End If
            If INI_ISearch Then
                AddOnInstruct += NetInstruct.Replace("; add", ", ")
            End If

            Dim lastCommaIndex As Integer = AddOnInstruct.LastIndexOf(","c)
            If lastCommaIndex <> -1 Then
                AddOnInstruct = AddOnInstruct.Substring(0, lastCommaIndex) & ", and" & AddOnInstruct.Substring(lastCommaIndex + 1)
            End If

            Dim application As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Selection = application.Selection

            If selection.Type = WdSelectionType.wdSelectionIP Then NoText = True

            SLib.StoreClipboard()

            If Not String.IsNullOrWhiteSpace(My.Settings.LastPrompt) Then SLib.PutInClipboard(My.Settings.LastPrompt)

            If Not NoText Then
                OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {ClipboardInstruct}, {InplaceInstruct} or {BubblesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{LastPromptInstruct}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False))
            Else
                OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct} or {BubblesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{LastPromptInstruct}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False))
            End If

            SLib.RestoreClipboard()

            ' Command line commands

            SelectedText = ""

            If Not NoText Then

                SelectedText = selection.Text

                If OtherPrompt.StartsWith("codebasis", StringComparison.OrdinalIgnoreCase) Then
                    SLib.WriteToRegistry(RemoveCR(RegPath_CodeBasis), RemoveCR(selection.Text))
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Exit Sub
                End If
                If OtherPrompt.StartsWith("inipath", StringComparison.OrdinalIgnoreCase) Then
                    SLib.WriteToRegistry(RemoveCR(RegPath_IniPath), RemoveCR(selection.Text))
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Exit Sub
                End If
                If OtherPrompt.StartsWith("encode", StringComparison.OrdinalIgnoreCase) Then
                    Dim Key As String = CodeAPIKey(RemoveCR(selection.Text))
                    SLib.PutInClipboard(Key)
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    selection.TypeText(vbCrLf & "Encoded key (also in clipboard):" & vbCrLf & Key)
                    selection.ParagraphFormat.Hyphenation = False ' Turn off hyphenation
                    SLib.PutInClipboard(Key)
                    Exit Sub
                End If

                If OtherPrompt.StartsWith("decode", StringComparison.OrdinalIgnoreCase) Then
                    Dim Key As String = DeCodeAPIKey(RemoveCR(selection.Text))
                    SLib.PutInClipboard(Key)
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    selection.TypeText(vbCrLf & "Decoded key (also in clipboard):" & vbCrLf & Key)
                    selection.ParagraphFormat.Hyphenation = False ' Turn off hyphenation
                    Exit Sub
                End If

            End If
            If OtherPrompt.StartsWith("domain", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"{AN} is running in the domain '{GetDomain()}' and configured to run in {If(String.IsNullOrEmpty(SLib.alloweddomains), "any domain ('alloweddomains' has not been set).", "'" & SLib.alloweddomains & "'.")}", "")
                Exit Sub
            End If
            If OtherPrompt.StartsWith("model", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox("I am using the " & INI_Model & " model as my primary model with a default timeout of " & (INI_Timeout / 1000) & " seconds (" & Format(INI_Timeout / 60000, "0.00") & " minutes)." & If(INI_MaxOutputToken > 0, "The maximum output token length is " & INI_MaxOutputToken & ".", ""))
                Exit Sub
            End If
            If OtherPrompt.StartsWith("terms", StringComparison.OrdinalIgnoreCase) Then
                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                selection.TypeText(vbCrLf & If(INI_UsageRestrictions = "", "No usage restrictions or permissions have been defined in the configuration file.", "The defined usage restrictions or permissions defined in the configuration file are: " & INI_UsageRestrictions) & vbCrLf)
                Exit Sub
            End If
            If OtherPrompt.StartsWith("switch", StringComparison.OrdinalIgnoreCase) Then
                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                If INI_SecondAPI Then
                    ShowCustomMessageBox("You have temporarily switched the two configured models. Primary is now '" & INI_Model & "', and secondary is '" & INI_Model_2 & "'.")
                Else
                    ShowCustomMessageBox("You have defined only one model ('" & INI_Model & "').")
                End If
                Exit Sub
            End If
            If OtherPrompt.StartsWith("version", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox("You are using " & Version & $" of {AN}. (c) by David Rosenthal, VISCHER. Go to https://vischer.com/{AN2} for more information. This copy of {AN} is set to expire on {LicensedTill.ToString("dd-MMM-yyyy")}", AN)
                Exit Sub
            End If
            If OtherPrompt.StartsWith("reset", StringComparison.OrdinalIgnoreCase) Then
                If ShowCustomYesNoBox($"Do you really want to reset your local configuration file and settings (if any) by removing non-mandatory entries? The current configuration file '{AN2}.ini' will NOT be saved to a '.bak' file. If you only want to reload the configuration settings for giving up any temporary changes, use 'reload' instead.", "Yes", "No") = 1 Then
                    INIloaded = False
                    ResetLocalAppConfig(_context)
                    MenusAdded = False
                    AddContextMenu()
                    ShowCustomMessageBox($"Following the reset, the configuration file '{AN2}.ini' has been be reloaded.")
                End If
                Exit Sub
            End If

            If OtherPrompt.StartsWith("speech", StringComparison.OrdinalIgnoreCase) Then
                Transcriptor()
                Exit Sub

            End If

            If OtherPrompt.StartsWith("readlocal", StringComparison.OrdinalIgnoreCase) Then
                SpeakSelectedText()
                Exit Sub

            End If

            If OtherPrompt.StartsWith("voiceslocal", StringComparison.OrdinalIgnoreCase) Then
                SelectVoiceByNumber()
                Exit Sub
            End If

            If OtherPrompt.StartsWith("voices2", StringComparison.OrdinalIgnoreCase) Then
                Using frm As New TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voices you wish to use.", $"{AN} Google Text-to-Speech - Select Voices", True)
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

                Exit Sub
            End If

            If OtherPrompt.StartsWith("voices", StringComparison.OrdinalIgnoreCase) Then
                Using frm As New TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voices you wish to use.", $"{AN} Google Text-to-Speech - Select Voices", False)
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

                Exit Sub
            End If

            If OtherPrompt.StartsWith("createpodcast", StringComparison.OrdinalIgnoreCase) Then
                CreatePodcast()
                Exit Sub
            End If

            If OtherPrompt.StartsWith("readpodcast", StringComparison.OrdinalIgnoreCase) Then
                ReadPodcast(selection.Text)
                Exit Sub
            End If

            If OtherPrompt.StartsWith("read", StringComparison.OrdinalIgnoreCase) Then
                CreateAudio()
                Exit Sub
            End If

            If OtherPrompt.StartsWith("cleanmenu", StringComparison.OrdinalIgnoreCase) Then
                RemoveOldContextMenu()
                RemoveVeryOldContextMenu()
                MenusAdded = False
                AddContextMenu()
                Exit Sub
            End If

            If OtherPrompt.StartsWith("reload", StringComparison.OrdinalIgnoreCase) Then
                INIloaded = False
                InitializeConfig(False, True)
                MenusAdded = False
                AddContextMenu()
                ShowCustomMessageBox($"The configuration file '{AN2}.ini' has been be reloaded.")
                Exit Sub
            End If
            If OtherPrompt.StartsWith("settings", StringComparison.OrdinalIgnoreCase) Then
                ShowSettings()
                Exit Sub
            End If

            If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then

                Dim promptlibresult As (String, Boolean, Boolean, Boolean)

                promptlibresult = ShowPromptSelector(INI_PromptLibPath, Not NoText, Not NoText)

                OtherPrompt = promptlibresult.Item1
                DoMarkup = promptlibresult.Item2
                DoBubbles = promptlibresult.Item3
                DoClipboard = promptlibresult.Item4

                If OtherPrompt = "" Then
                    Exit Sub
                End If
            Else
                If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Exit Sub
            End If

            My.Settings.LastPrompt = OtherPrompt
            My.Settings.Save()

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
            End If


            If OtherPrompt.IndexOf(NetTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NetTrigger, "").Trim()
                DoNet = True
            End If


            If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                doc = GetFileContent()
                OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ExtTrigger), doc, RegexOptions.IgnoreCase)
                ShowCustomMessageBox($"This file will be included in your prompt where you have referred to {ExtTrigger}: " & vbCrLf & vbCrLf & doc)
            End If

            If NoText And DoBubbles Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Ask the LLM to comment on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Exit Sub
                End If
            End If

            If NoText And DoMarkup Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Do the markup on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Exit Sub
                End If
            End If

            If Not DoInplace And DoMarkup Then
                Dim AppendMarkup As Integer = ShowCustomYesNoBox("You have asked for a markup to be created, but according to the configuration, it will not replace your current selection but added to it at the end. Is this really what you want?", "Yes, add markup ", "No, replace text with markup")
                If AppendMarkup = 0 Then
                    Exit Sub
                ElseIf AppendMarkup = 2 Then
                    DoInplace = True
                End If
            End If

            If DoLib Then
                Dim isSuccess As Boolean = Await ConsultLibrary(DoMarkup) ' updates SysPrompt
                If Not isSuccess Then Exit Sub
            ElseIf DoNet Then
                Dim isSuccess As Boolean = Await ConsultInternet(DoMarkup) ' updates SysPrompt
                If Not isSuccess Then Exit Sub
            ElseIf NoText Then
                SysPrompt = SP_FreestyleNoText
            Else
                SysPrompt = SP_FreestyleText
                If DoBubbles Then SysPrompt = SysPrompt & " " & SP_Add_Bubbles
            End If

            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SysPrompt), True, DoKeepFormat, DoKeepParaFormat, DoInplace, DoMarkup, MarkupMethod, DoClipboard, DoBubbles, False, UseSecondAPI, KeepFormatCap, DoTPMarkup, TPMarkupName)

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

    Private Async Function ProcessSelectedText(SysCommand As String, CheckMaxToken As Boolean, KeepFormat As Boolean, ParaFormatInline As Boolean, InPlace As Boolean, DoMarkup As Boolean, MarkupMethod As Integer, PutInClipboard As Boolean, PutInBubbles As Boolean, SelectionMandatory As Boolean, UseSecondAPI As Boolean, FormattingCap As Integer, Optional DoTPMarkup As Boolean = False, Optional TPMarkupname As String = "", Optional CreatePodcast As Boolean = False) As Task(Of String)

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

            Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast)

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
                                                                                FormattingCap, DoTPMarkup, TPMarkupname)
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
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname)
                                    Await System.Threading.Tasks.Task.Delay(500)
                                End If
                            Else

                                Do
                                    textChunk.Start += 1
                                Loop While textChunk.Tables.Count <> 0 And Not textChunk.Start = textChunk.End

                                If textChunk.Tables.Count = 0 AndAlso textChunk.Start < textChunk.End Then
                                    textChunk.Select()
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname)
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
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname)
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
                            Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname)
                        Else
                            Do
                                finalChunk.Start += 1
                            Loop While finalChunk.Tables.Count <> 0 And Not finalChunk.Start = finalChunk.End

                            finalChunk.End = selRange.End

                            If finalChunk.Tables.Count = 0 AndAlso finalChunk.Start < finalChunk.End Then
                                finalChunk.Select()
                                Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname)
                            End If
                        End If
                    End If

                    splash.Close()
                End If

            ElseIf userdialog = 1 Then

                Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast)

            End If

        End If

        If Not PutInClipboard Then
            selection.Collapse(WdCollapseDirection.wdCollapseEnd)
            selection.MoveStart(WdUnits.wdCharacter, 0)
            selection.MoveEnd(WdUnits.wdCharacter, 0)
        End If

        Return ""

    End Function
    Private Async Function TrueProcessSelectedText(SysCommand As String, CheckMaxToken As Boolean, KeepFormat As Boolean, ParaFormatInline As Boolean, InPlace As Boolean, DoMarkup As Boolean, MarkupMethod As Integer, PutInClipboard As Boolean, PutInBubbles As Boolean, SelectionMandatory As Boolean, UseSecondAPI As Boolean, FormattingCap As Integer, Optional DoTPMarkup As Boolean = False, Optional TPMarkupname As String = "", Optional CreatePodcast As Boolean = False) As Task(Of String)

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

            If Not ParaFormatInline And Not NoFormatting And Not NoSelectedText Then

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

            Dim LLMResult = Await LLM(SysCommand & " " & If(NoFormatting, "", If(KeepFormat, " " & SP_Add_KeepHTMLIntact, " " & SP_Add_KeepInlineIntact)), If(NoSelectedText, "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0, UseSecondAPI, False, OtherPrompt)

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

                If CreatePodcast Then

                    Dim IsGoogle As Boolean = False
                    Dim Endpoint As String = INI_Endpoint
                    Dim Endpoint_2 As String = INI_Endpoint_2
                    Dim TTSEndpoint As String = INI_TTSEndpoint
                    If Endpoint.Contains(GoogleIdentifier) Then
                        If Endpoint.Contains(GoogleIdentifier) Then
                            TTSSecondAPI = False
                            IsGoogle = True
                        ElseIf Endpoint_2.Contains(GoogleIdentifier) Then
                            TTSSecondAPI = True
                            IsGoogle = True
                        End If
                    End If

                    If IsGoogle Then
                        Dim FinalText = ShowCustomWindow("The LLM has created the following podcast script for you (you can edit it; you do not have to manually remove the SSML codes, if you do not like them):", LLMResult, "The next step is the production of an audio file. You can choose whether you want to use the original text or your text with any changes you have made. The text will also be put in the clipboard. If you select Cancel, the original text will only be put into the clipboard.", AN, True)

                        If FinalText = "" Then
                            SLib.PutInClipboard(LLMResult)
                        Else
                            FinalText = Trim(FinalText)
                            SLib.PutInClipboard(FinalText)
                            If FinalText.Contains("H: ") And FinalText.Contains("G: ") Then ReadPodcast(FinalText)
                        End If
                    Else
                        Dim FinalText = ShowCustomWindow("The LLM has created the following podcast script for you (you can edit it; you do not have to manually remove the SSML codes, if you do not like them):", LLMResult, $"The next step is the production of an audio file. Since you have not configured {AN} for Google, you unfortunately cannot do that here. However, you can choose whether you want the original text or the text with your changes to put in the clipboard for further use. If you select Cancel, no text will be put in the clipboard.", AN, True)

                        If FinalText <> "" Then
                            SLib.PutInClipboard(LLMResult)
                        Else
                            FinalText = Trim(FinalText)
                            SLib.PutInClipboard(FinalText)
                        End If
                    End If

                ElseIf PutInClipboard Then

                    Dim FinalText = ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made. If you select Cancel, nothing will be put into the clipboard.", AN)

                    If FinalText <> "" Then
                        SLib.PutInClipboard(FinalText)
                    End If

                ElseIf PutInBubbles Then

                    Dim responseItems() As String = LLMResult.Split({"§§§"}, StringSplitOptions.RemoveEmptyEntries)
                    Dim wrongformatresponse As New List(Of String)
                    Dim notfoundresponse As New List(Of String)
                    Dim originalRange As Word.Range = selection.Range.Duplicate ' Save the original selection range
                    Dim BubblecutHappened As Boolean = False

                    Dim splash As New SplashScreen("Adding bubbles to your text... press 'Esc' to abort")
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
                                        notfoundresponse.Add("'" & findText & "' " & ChrW(8594) & $" {AN5}: " & commentText)
                                    End If
                                Else
                                    ' Use chunk-by-chunk search for > 255 characters
                                    If FindLongTextInChunks(findText, 255, selection) Then
                                        ' If found, selection now covers the entire matched text
                                        Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, $"{AN5}: " & commentText)
                                    Else
                                        notfoundresponse.Add("'" & findText & "' " & ChrW(8594) & $" {AN5}: " & commentText)
                                    End If
                                End If

                            Catch ex As Exception
                                notfoundresponse.Add("'" & findText & "' " & ChrW(8594) & $" {AN5}: " & commentText & " [Error: " & ex.Message & "]")
                            End Try

                        Else
                            wrongformatresponse.Add(item)
                        End If

                        selection.SetRange(originalRange.Start, originalRange.End) ' Restore the original selection
                    Next

                    splash.Close()

                    Dim ErrorList As String = ""
                    If notfoundresponse.Count > 0 Then
                        ErrorList += "The following comments could not be assigned to your text (they were not found):" & vbCrLf
                        For Each item In notfoundresponse
                            If Trim(item) <> "" Then ErrorList += Trim("- " & item & vbCrLf)
                        Next
                        ErrorList += vbCrLf
                    End If

                    If wrongformatresponse.Count > 0 Then
                        ErrorList += "The following responses could not be identified as bubble comments:" & vbCrLf
                        For Each item In wrongformatresponse
                            If Trim(item) <> "" Then ErrorList += Trim("- " & item & vbCrLf)
                        Next
                        ErrorList += vbCrLf
                    End If
                    If Not String.IsNullOrWhiteSpace(ErrorList) Then
                        If BubblecutHappened Then
                            ErrorList = $"Some of the sections to which the bubble comments relate were too long for selecting. Only the initial part has been selected. This is indicated by '{BubbleCutText}' in the bubble comments, as applicable." & vbCrLf & vbCrLf & ErrorList
                        End If

                        ErrorList = ShowCustomWindow("Errors when implementing the 'bubbles' feedback of the LLM:", ErrorList, "The above error list will be included in a final comment at the end of your selection (it will also be included in the clipboard). You can have the original list included, or you can now make changes and have this version used. If you select Cancel, nothing will be put added to the document.", AN)

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

                ElseIf KeepFormat And Not NoFormatting Then
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
                                    rng = selection.Range
                                End If
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
                                    SLib.InsertTextWithBoldMarkers(selection, LLMResult)
                                    rng = selection.Range
                                End If
                                Dim SaveRng As Range = rng.Duplicate
                                CompareAndInsert(SelectedText, LLMResult, rng.Duplicate, MarkupMethod = 3, "This is the markup of the text inserted:")
                                If Not ParaFormatInline And Not NoFormatting Then
                                    ApplyParagraphFormat(rng)
                                End If
                                RestoreSpecialTextElements(SaveRng)
                            Else
                                CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                                RestoreSpecialTextElements(rng)
                            End If
                        Else
                            SLib.InsertTextWithBoldMarkers(selection, LLMResult)
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

            ' Attempt to find the chunk
            With selection.Find
                .Text = currentChunk
                .Forward = True
                .Wrap = Word.WdFindWrap.wdFindStop
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
            Exit Sub
        End If

        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Selection = app.Selection

            If selection Is Nothing OrElse selection.Range Is Nothing Then
                MessageBox.Show("Error in MarkupSelectedTextWithRegex: No text selected (anymore). Can't proceed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
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
            app.UserName = AN

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

                    'Dim regex As New Regex(regexPair.Pattern)
                    'Dim matches = regex.Matches(selectedRange.Text)
                    'For Each match As Match In matches
                    'If match.Success Then
                    'Dim matchRange As Range = selectedRange.Duplicate
                    'matchRange.Start = selectedRange.Start + match.Index
                    'matchRange.End = matchRange.Start + match.Length
                    'matchRange.Text = regexPair.Replacement
                    'End If
                    'Next

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
            app.UserName = originalUserName

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
                            .Text = oldText
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
            Exit Sub
        End If

        Dim insertionStart As Integer = selection.Range.Start

        ' Extract the range from the selection
        Dim range As Microsoft.Office.Interop.Word.Range = selection.Range

        ' Convert Markdown to HTML using Markdig
        Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()
        Dim htmlResult As String = Markdown.ToHtml(Result, markdownPipeline).Trim

        ' Load the HTML into HtmlDocument
        Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
        htmlDoc.LoadHtml(htmlResult)

        Dim nodesToRemove = htmlDoc.DocumentNode.SelectNodes("//*[not(node()) and not(normalize-space())]")
        If nodesToRemove IsNot Nothing Then
            For Each emptyNode In nodesToRemove
                emptyNode.Remove()
            Next
        End If

        ' Parse and insert HTML content into the Word range
        ParseHtmlNode(htmlDoc.DocumentNode, range)

        If TrailingCR Then
            range.Text = vbCrLf
            range.Collapse(False)
        End If

        Dim insertionEnd As Integer = range.End

        Dim doc As Microsoft.Office.Interop.Word.Document = selection.Document
        selection.SetRange(insertionStart, insertionEnd)
        selection.Select()

    End Sub

    Private Shared Sub ParseHtmlNode(node As HtmlNode, range As Microsoft.Office.Interop.Word.Range)
        For Each childNode As HtmlNode In node.ChildNodes
            Select Case childNode.Name.ToLower()
                Case "#text"
                    ' Insert plain text
                    range.Text = HtmlEntity.DeEntitize(childNode.InnerText)
                    range.Font.Reset()
                    range.Collapse(False)

                Case "strong", "b"
                    ' Bold text
                    Dim boldRange As Range = range.Duplicate
                    boldRange.Text = HtmlEntity.DeEntitize(childNode.InnerText)
                    boldRange.Font.Bold = True
                    boldRange.Collapse(False)
                    range.SetRange(boldRange.End, boldRange.End)

                Case "em", "i"
                    ' Italic text
                    Dim italicRange As Range = range.Duplicate
                    italicRange.Text = HtmlEntity.DeEntitize(childNode.InnerText)
                    italicRange.Font.Italic = True
                    italicRange.Collapse(False)
                    range.SetRange(italicRange.End, italicRange.End)

                Case "u"
                    ' Underlined text
                    Dim underlineRange As Range = range.Duplicate
                    underlineRange.Text = HtmlEntity.DeEntitize(childNode.InnerText)
                    underlineRange.Font.Underline = WdUnderline.wdUnderlineSingle
                    underlineRange.Collapse(False)
                    range.SetRange(underlineRange.End, underlineRange.End)

                Case "br"
                    ' Line break
                    range.Text = vbCr
                    range.Collapse(False)

                Case "yyp", "yydiv"
                    If Not String.IsNullOrWhiteSpace(childNode.InnerText) Then
                        ParseHtmlNode(childNode, range)
                        ' Remove the extra vbCr insertion
                        ' (Or insert it only if absolutely needed)
                    End If

                Case "p", "div"
                    If Not String.IsNullOrWhiteSpace(childNode.InnerText) Then
                        ParseHtmlNode(childNode, range)
                        If Not childNode.NextSibling Is Nothing Then
                            range.Text = vbCr ' Only add a line break if there's a sibling
                            range.Collapse(False)
                        End If
                    End If

                Case "a"
                    ' Hyperlink
                    Dim hyperlinkRange As Range = range.Duplicate
                    hyperlinkRange.Text = HtmlEntity.DeEntitize(childNode.InnerText)
                    range.Document.Hyperlinks.Add(hyperlinkRange, childNode.GetAttributeValue("href", ""))
                    hyperlinkRange.Collapse(False)
                    range.SetRange(hyperlinkRange.End, hyperlinkRange.End)


                Case "ul"
                    ' 1) Listen-Start merken
                    Dim startOfList As Integer = range.Start

                    ' 2) Li-Elemente einfügen
                    For Each listItem As HtmlNode In childNode.SelectNodes("li")
                        ' Wir wollen den Inhalt des Li rekursiv behandeln,
                        ' weil es darin z.B. <strong> oder <em> geben kann.
                        ParseHtmlNode(listItem, range)

                        ' Absatzende
                        range.Text = vbCr
                        range.Collapse(False)
                    Next

                    ' 3) Range für die ganze Liste definieren
                    Dim bulletListRange As Range = range.Document.Range(startOfList, range.End)

                    ' 4) Bullet-Liste anwenden
                    bulletListRange.ListFormat.ApplyBulletDefault()
                    bulletListRange.ListFormat.ListIndent()
                    With bulletListRange.ParagraphFormat
                        .LeftIndent = 14.18
                        .FirstLineIndent = -14.18
                    End With

                Case "ol"
                    ' 1) Listen-Start merken
                    Dim startOfList As Integer = range.Start

                    ' 2) Li-Elemente einfügen
                    For Each listItem As HtmlNode In childNode.SelectNodes("li")
                        ' Auch hier wieder rekursiv behandeln
                        ParseHtmlNode(listItem, range)

                        ' Absatzende
                        range.Text = vbCr
                        range.Collapse(False)
                    Next

                    ' 3) Range für die ganze Liste definieren
                    Dim numberedListRange As Range = range.Document.Range(startOfList, range.End)

                    ' 4) Nummerierte Liste anwenden
                    numberedListRange.ListFormat.ApplyNumberDefault()
                    numberedListRange.ListFormat.ListIndent()

                    With numberedListRange.ParagraphFormat
                        .LeftIndent = 14.18
                        .FirstLineIndent = -14.18
                    End With


                Case "h1"
                    ' Heading 1
                    Dim h1Range As Range = range.Duplicate
                    h1Range.Text = HtmlEntity.DeEntitize(childNode.InnerText) & vbCr
                    h1Range.Font.Size = 16
                    h1Range.Font.Bold = True
                    h1Range.Collapse(False)
                    range.SetRange(h1Range.End, h1Range.End)

                Case "h2"
                    ' Heading 2
                    Dim h2Range As Range = range.Duplicate
                    h2Range.Text = HtmlEntity.DeEntitize(childNode.InnerText) & vbCr
                    h2Range.Font.Size = 14
                    h2Range.Font.Bold = True
                    h2Range.Collapse(False)
                    range.SetRange(h2Range.End, h2Range.End)

                Case "h3"
                    ' Heading 3
                    Dim h3Range As Range = range.Duplicate
                    h3Range.Text = HtmlEntity.DeEntitize(childNode.InnerText) & vbCr
                    h3Range.Font.Size = 12
                    h3Range.Font.Bold = True
                    h3Range.Collapse(False)
                    range.SetRange(h3Range.End, h3Range.End)

                Case "code"
                    ' Inline code
                    Dim codeRange As Range = range.Duplicate
                    codeRange.Text = HtmlEntity.DeEntitize(childNode.InnerText)
                    codeRange.Font.Name = "Courier New"
                    codeRange.Font.Size = 10
                    codeRange.Collapse(False)
                    range.SetRange(codeRange.End, codeRange.End)

                Case "pre"
                    ' Code block
                    Dim preRange As Range = range.Duplicate
                    preRange.Text = HtmlEntity.DeEntitize(childNode.InnerText) & vbCr
                    preRange.Font.Name = "Courier New"
                    preRange.Font.Size = 10
                    preRange.Collapse(False)
                    range.SetRange(preRange.End, preRange.End)

                Case "hr"
                    ' Horizontal rule
                    Dim hrRange As Range = range.Duplicate
                    hrRange.Text = vbCr & "――――――――――――――――――――――――――――――――――" & vbCr
                    hrRange.Collapse(False)
                    range.SetRange(hrRange.End, hrRange.End)

                Case Else
                    ' Handle other tags recursively
                    ParseHtmlNode(childNode, range)
            End Select
        Next
    End Sub

    Private Shared Sub DebugParseHtmlNode(node As HtmlNode, range As Microsoft.Office.Interop.Word.Range)
        For Each childNode As HtmlNode In node.ChildNodes
            Select Case childNode.Name.ToLower()
                Case "#text"
                    ' Insert plain text
                    range.Text = childNode.InnerText
                    range.Collapse(False)

                Case "strong", "b"
                    ' Bold text
                    Dim boldRange As Range = range.Duplicate
                    boldRange.Text = childNode.InnerText
                    boldRange.Font.Bold = True
                    boldRange.Collapse(False)
                    range.SetRange(boldRange.End, boldRange.End)

                Case "em", "i"
                    ' Italic text
                    Dim italicRange As Range = range.Duplicate
                    italicRange.Text = childNode.InnerText
                    italicRange.Font.Italic = True
                    italicRange.Collapse(False)
                    range.SetRange(italicRange.End, italicRange.End)

                Case "u"
                    ' Underlined text
                    Dim underlineRange As Range = range.Duplicate
                    underlineRange.Text = childNode.InnerText
                    underlineRange.Font.Underline = WdUnderline.wdUnderlineSingle
                    underlineRange.Collapse(False)
                    range.SetRange(underlineRange.End, underlineRange.End)

                Case "s", "del", "strike"
                    ' Strikethrough text
                    Dim strikeRange As Range = range.Duplicate
                    strikeRange.Text = childNode.InnerText
                    strikeRange.Font.StrikeThrough = True
                    strikeRange.Collapse(False)
                    range.SetRange(strikeRange.End, strikeRange.End)

                Case "br"
                    ' Line break
                    range.Text = vbCr
                    range.Collapse(False)

                Case "p", "div"
                    ' Paragraph handling
                    ParseHtmlNode(childNode, range)
                    range.Text = vbCr
                    range.Collapse(False)

                Case "a"
                    ' Hyperlink
                    Dim hyperlinkRange As Range = range.Duplicate
                    hyperlinkRange.Text = childNode.InnerText
                    range.Document.Hyperlinks.Add(hyperlinkRange, childNode.GetAttributeValue("href", ""))
                    hyperlinkRange.Collapse(False)
                    range.SetRange(hyperlinkRange.End, hyperlinkRange.End)

                Case "ul"
                    ' Unordered list using Word bullet format
                    Dim listRange As Range = range.Duplicate
                    listRange.ListFormat.ApplyBulletDefault()
                    For Each listItem As HtmlNode In childNode.SelectNodes("li")
                        ParseHtmlNode(listItem, listRange)
                    Next
                    ' Turn off bullet formatting after the list
                    listRange.ListFormat.RemoveNumbers()

                Case "ol"
                    ' Ordered list using Word numbered format
                    Dim listRange As Range = range.Duplicate
                    listRange.ListFormat.ApplyNumberDefault()
                    For Each listItem As HtmlNode In childNode.SelectNodes("li")
                        ParseHtmlNode(listItem, listRange)
                    Next
                    ' Turn off numbering formatting after the list
                    listRange.ListFormat.RemoveNumbers()

                Case "li"
                    ' Handle list item
                    Dim listItemRange As Range = range.Duplicate
                    listItemRange.Text = childNode.InnerText & vbCr
                    listItemRange.Collapse(False)
                    For Each subNode As HtmlNode In childNode.ChildNodes
                        If subNode.Name.ToLower() = "ul" OrElse subNode.Name.ToLower() = "ol" Then
                            ParseHtmlNode(subNode, listItemRange)
                        End If
                    Next

                Case "h1"
                    ' Heading 1
                    Dim h1Range As Range = range.Duplicate
                    h1Range.Text = childNode.InnerText & vbCr
                    h1Range.Font.Size = 16
                    h1Range.Font.Bold = True
                    h1Range.Collapse(False)
                    range.SetRange(h1Range.End, h1Range.End)

                Case "h2"
                    ' Heading 2
                    Dim h2Range As Range = range.Duplicate
                    h2Range.Text = childNode.InnerText & vbCr
                    h2Range.Font.Size = 14
                    h2Range.Font.Bold = True
                    h2Range.Collapse(False)
                    range.SetRange(h2Range.End, h2Range.End)

                Case "h3"
                    ' Heading 3
                    Dim h3Range As Range = range.Duplicate
                    h3Range.Text = childNode.InnerText & vbCr
                    h3Range.Font.Size = 12
                    h3Range.Font.Bold = True
                    h3Range.Collapse(False)
                    range.SetRange(h3Range.End, h3Range.End)

                Case "code"
                    ' Inline code
                    Dim codeRange As Range = range.Duplicate
                    codeRange.Text = childNode.InnerText
                    codeRange.Font.Name = "Courier New"
                    codeRange.Font.Size = 10
                    codeRange.Collapse(False)
                    range.SetRange(codeRange.End, codeRange.End)

                Case "pre"
                    ' Code block
                    Dim preRange As Range = range.Duplicate
                    preRange.Text = childNode.InnerText & vbCr
                    preRange.Font.Name = "Courier New"
                    preRange.Font.Size = 10
                    preRange.Collapse(False)
                    range.SetRange(preRange.End, preRange.End)

                Case "hr"
                    ' Horizontal rule
                    Dim hrRange As Range = range.Duplicate
                    hrRange.Text = vbCr & "――――――――――――――――――――――――――――――――――" & vbCr
                    hrRange.Collapse(False)
                    range.SetRange(hrRange.End, hrRange.End)

                Case Else
                    ' Handle other tags recursively
                    ParseHtmlNode(childNode, range)
            End Select
        Next
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
            wordApp.UserName = AN

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
                RevisedAuthor:=AN
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
            wordApp.UserName = originalAuthor

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
                    Exit Sub
                End If
                tagText = inputText.Substring(pos - 1, tagEndPos - pos)
                pos = tagEndPos + 9
            ElseIf inputText.Substring(pos - 1, System.Math.Min(11, lenText - pos + 1)) = "[DEL_START]" Then
                pos += 11
                tagType = 2 ' Delete formatting
                tagEndPos = inputText.IndexOf("[DEL_END]", pos - 1) + 1
                If tagEndPos = -1 Then
                    MessageBox.Show("Error in ParseText: Missing [DEL_END] tag.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
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
            Exit Sub
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
        If String.IsNullOrEmpty(RegexPattern) Then Exit Sub

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
            Exit Sub
        End If

        ' Validate all regex patterns first
        For Each pattern As String In patterns
            Try
                Dim regexTest As New Regex(pattern, regexOptions)
            Catch ex As ArgumentException
                ShowCustomMessageBox($"Your regex pattern '{pattern}' is invalid ({ex.Message}). Aborting without any replacements done.")
                Exit Sub
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
                    Exit Sub
                Else
                    ShowCustomMessageBox($"No matches found for '{pattern}' {DocRef}.")
                    Exit Sub
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
            Exit Sub
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

    Public Async Sub ContextSearch()

        Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
        Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.ActiveDocument
        Dim SearchAll As Boolean = False
        Dim SearchMulti As Boolean = False

        Dim lastcontextsearch As String = If(String.IsNullOrWhiteSpace(My.Settings.LastContextSearch), "", My.Settings.LastContextSearch)

        SearchContext = Trim(ShowCustomInputBox($"Enter the search term (start with '{SearchMultiTrigger}' if you want to find and highlight all occurrences at once and add '{SearchAllTrigger}' if you want to search the entire document):", "Context Search", True, lastcontextsearch))
        If String.IsNullOrWhiteSpace(SearchContext) Or SearchContext = "ESC" Then Exit Sub

        My.Settings.LastContextSearch = SearchContext
        My.Settings.Save()

        If SearchContext.IndexOf(SearchAllTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            SearchContext = SearchContext.Replace(SearchAllTrigger, "").Trim()
            SearchAll = True
        End If
        If SearchContext.StartsWith(SearchMultiTrigger, StringComparison.OrdinalIgnoreCase) Then
            SearchContext = SearchContext.Replace(SearchMultiTrigger, "").Trim()
            SearchMulti = True
        End If

        SearchContext = SearchContext.Replace("  ", "")

        Dim SearchText As String

        If Not String.IsNullOrWhiteSpace(selection.Text) And Len(selection.Text) > 3 And Not SearchAll Then
            SearchText = selection.Text
        ElseIf selection.Start < selection.Document.Content.End And Not SearchAll Then
            SearchText = selection.Document.Range(selection.Start, selection.Document.Content.End).Text
            selection.SetRange(selection.Start, selection.Document.Content.End)
        Else
            SearchText = selection.Document.Content.Text
            selection.SetRange(0, selection.Document.Content.End)
            SearchAll = True
        End If

        Dim LLMResult As String = Await LLM(InterpolateAtRuntime(If(SearchMulti, SP_ContextSearchMulti, SP_ContextSearch)), "<TEXTTOSEARCH>" & SearchText & "</TEXTTOSEARCH>", "", "", 0)

        LLMResult = LLMResult.Replace("<TEXTTOSEARCH>", "").Replace("</TEXTTOSEARCH>", "")

        If SearchMulti Then

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
                doc.Application.UserName = AN

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
                        selection.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow
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
                doc.Application.UserName = originalAuthor

            Else
                ShowCustomMessageBox($"The LLM has found no hits for the context '{SearchContext}'{If(SearchAll, " (searched full document)", "")}.", "Context Search")
            End If

        Else
            If Not String.IsNullOrWhiteSpace(LLMResult) Then
                Dim FindText As String = Trim(LLMResult)
                FindText = FindText.TrimEnd(ControlChars.Lf)
                FindText = FindText.TrimEnd(ControlChars.Cr)

                If FindLongTextInChunks(FindText, 255, selection) And selection IsNot Nothing Then
                    wordApp.ActiveWindow.ScrollIntoView(selection.Range, True)
                Else
                    ShowCustomMessageBox($"The LLM did not find the context '{SearchContext}'{If(SearchAll, " (searched full document)", " (search direction was down)")}.", "Context Search")
                End If
            End If
        End If
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

        Dim Settings As New Dictionary(Of String, String) From {
                {"Temperature", "Temperature of {model}"},
                {"Timeout", "Timeout of {model}"},
                {"Temperature_2", "Temperature of {model2}"},
                {"Timeout_2", "Timeout of {model2}"},
                {"DoubleS", "Convert '" & ChrW(223) & "' to 'ss'"},
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

    Private Async Function OldStartHttpListener() As Task(Of String)
        Dim prefix As String = "http://127.0.0.1:12334/"
        Try
            ' Initialize the listener once.
            If httpListener Is Nothing Then
                httpListener = New HttpListener()
                httpListener.Prefixes.Add(prefix)
                httpListener.Start()
                Debug.WriteLine("HttpListener started.")
            End If

            While Not isShuttingDown
                ' If for some reason the listener is not listening (disposed), restart it.
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
                    ' Use asynchronous call to wait for an incoming request.
                    Dim context As HttpListenerContext = Await httpListener.GetContextAsync()
                    Dim result As String = Await HandleHttpRequest(context)
                Catch ex As System.ObjectDisposedException
                    Debug.WriteLine("HttpListener was disposed. Restarting listener...")
                    ' Continue to the next iteration so that the above block restarts the listener.
                    Continue While
                Catch ex As System.Exception
                    Debug.WriteLine("Error httplistener handling request: " & ex.Message)
                End Try
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
                If textToInsert IsNot Nothing Then
                    ' Get the active Word document and the selection
                    Dim app As Word.Application = Globals.ThisAddIn.Application
                    Dim selection As Word.Selection = app.Selection

                    ' Insert the text at the current cursor position
                    selection.TypeText(textToInsert)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"Error in ProcessRequestInAddIn: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return result

    End Function

    Public Class TranscriptionForm
        Inherits Form

        ' --- UI Components ---
        Private RichTextBox1 As Forms.RichTextBox
        Private StartButton As Forms.Button
        Private StopButton As Forms.Button
        Private ClearButton As Forms.Button
        Private LoadButton As Forms.Button
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

        ' --- Vosk & Audio State ---
        Private recognizer As VoskRecognizer
        Private waveIn As WaveInEvent
        Private capturing As Boolean = False
        Private partialText As String = ""
        Private finalText As New StringBuilder()
        Private Const VoskTooltip = "Only for Vosk: Set similarity threshold for speaker identification (0.5-0.7 for real-time speaker tracking, 1.0-1.5 for meetings/interviews)"
        Private Const VoskToggle = "Iden"

        Private waveInputs As New List(Of WaveInEvent)()
        Private waveProviders As New List(Of BufferedWaveProvider)()
        Private waveMixer As MixingSampleProvider
        Private MultiSource As Boolean = False
        Private audioThread As Thread
        Private stopProcessing As Boolean = False

        Private WhisperRecognizer As WhisperProcessor
        Private audioBuffer As New List(Of Single)
        Private STTCanceled As Boolean = False
        Private cts As CancellationTokenSource = New CancellationTokenSource()
        Private Const WhisperTooltip = "Only for Whisper: Select if text shall be translated to English and the threshold for detecting voice (default = 0.6, increase for noisy environments)"
        Private Const WhisperToggle = "Trans"

        Private STTModel As String = "whisper"

        Public Sub New()
            ' Initialize UI Components
            InitializeComponents()

            ' Load available Vosk models
            Dim modelPath As String = Globals.ThisAddIn.INI_SpeechModelPath
            Dim modelsexist As Boolean = False
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
            AddHandler Me.cultureComboBox.DropDown, AddressOf cultureComboBox_MouseMove

            LoadAudioDevices()

            AddHandler Me.deviceComboBox.MouseMove, AddressOf deviceComboBox_MouseMove

            LoadAndPopulateProcessComboBox(Globals.ThisAddIn.INI_PromptLibPath_Transcript, processCombobox)

            Dim index As Integer = Me.cultureComboBox.SelectedIndex
            If index >= 0 Then
                If Me.cultureComboBox.Items(index).startswith("ggml") Then
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

        Private Sub cultureComboBox_MouseMove(sender As Object, e As MouseEventArgs)
            Dim index As Integer = Me.cultureComboBox.SelectedIndex
            If index >= 0 Then
                ToolTip.SetToolTip(Me.cultureComboBox, Me.cultureComboBox.Items(index).ToString())
                If Me.cultureComboBox.Items(index).startswith("ggml") Then
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


        Private Sub InitializeComponents()
            ' Set standard font for UI
            Me.Font = New System.Drawing.Font("Segoe UI", 9)

            ' --- UI Elements ---
            Me.RichTextBox1 = New RichTextBox() With {
            .Font = New System.Drawing.Font("Segoe UI", 10),
            .Multiline = True,
            .ScrollBars = RichTextBoxScrollBars.Vertical,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        }

            Me.StartButton = New Forms.Button() With {.Text = "Start", .Enabled = True, .AutoSize = True}
            Me.StopButton = New Forms.Button() With {.Text = "Stop", .Enabled = False, .AutoSize = True}
            Me.ClearButton = New Forms.Button() With {.Text = "Clear", .AutoSize = True}
            Me.LoadButton = New Forms.Button() With {.Text = "Load", .AutoSize = True}
            Me.QuitButton = New Forms.Button() With {.Text = "Quit", .AutoSize = True}
            Me.ProcessButton = New Forms.Button() With {.Text = "Process:", .AutoSize = True}

            Me.cultureComboBox = New Forms.ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

            Me.deviceComboBox = New Forms.ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

            Me.processCombobox = New Forms.ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

            Me.Label1 = New Label() With {
            .Text = "Model:",
            .AutoSize = True
        }

            Me.Label2 = New Label() With {
            .Text = "Source:",
            .AutoSize = True
        }

            Me.SpeakerIdent = New System.Windows.Forms.CheckBox() With {
            .Text = VoskToggle,
            .AutoSize = True,
            .Checked = My.Settings.LastSpeakerEnabled
        }

            ' Declare the ToolTip at the class level


            ' Initialize the TextBox and set the ToolTip
            Me.SpeakerDistance = New System.Windows.Forms.TextBox() With {
                    .AutoSize = False,
                    .Width = 31,
                    .Text = If(My.Settings.LastSpeakerDistance <= 0, "1.0", My.Settings.LastSpeakerDistance.ToString)
                }


            ' Status Label (New)
            Me.StatusLabel = New Label() With {
            .Text = "Transcribing:",
            .Font = New System.Drawing.Font("Segoe UI", 9, FontStyle.Regular),
            .ForeColor = Color.Black,
            .AutoSize = True
        }

            ' Partial Text Label
            Me.PartialTextLabel = New Label() With {
                .Font = New System.Drawing.Font("Segoe UI", 9, FontStyle.Italic),
                .ForeColor = Color.DimGray,
                .Text = "...",
                .AutoSize = False,
                .Height = 60,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }


            ' Panel for Buttons
            Me.ButtonPanel = New Panel() With {
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .Height = 40
        }

            ' --- Layout ---

            Me.ClientSize = New Drawing.Size(800, 580)
            Me.Text = $"{AN} Transcriptor (editable text, audio will not be stored)"
            Me.FormBorderStyle = FormBorderStyle.Sizable

            ' Define padding values
            Dim horizontalPadding As Integer = 10 ' Space between elements in a row
            Dim verticalPadding As Integer = 10   ' Space between rows

            ' Absolute position of Label1
            Label1.Location = New System.Drawing.Point(12, 15)

            ' Calculate actual width of Label1
            Dim label1Width As Integer = Label1.PreferredSize.Width

            ' Position cultureComboBox relative to Label1, ensuring enough space
            cultureComboBox.Location = New System.Drawing.Point(Label1.Left + label1Width + horizontalPadding + horizontalPadding, Label1.Top - 3)
            cultureComboBox.Size = New Size(250, 21)

            ' Calculate actual width of Label2
            Dim label2Width As Integer = Label2.PreferredSize.Width

            ' Position Label2 relative to cultureComboBox
            Label2.Location = New System.Drawing.Point(cultureComboBox.Left + cultureComboBox.Width + horizontalPadding + horizontalPadding, Label1.Top)

            ' Position deviceComboBox relative to Label2, ensuring enough space
            deviceComboBox.Location = New System.Drawing.Point(Label2.Left + label2Width + horizontalPadding + horizontalPadding, Label1.Top - 3)
            deviceComboBox.Size = New Size(250, 21)

            SpeakerIdent.Location = New System.Drawing.Point(deviceComboBox.Left + deviceComboBox.Width + horizontalPadding + horizontalPadding, Label1.Top - 1)
            SpeakerDistance.Location = New System.Drawing.Point(SpeakerIdent.Left + SpeakerIdent.PreferredSize.Width + 10, Label1.Top - 1)



            ' Position StatusLabel below the first row with consistent padding
            StatusLabel.Location = New System.Drawing.Point(12, Label1.Bottom + verticalPadding)

            ' Position PartialTextLabel below StatusLabel
            PartialTextLabel.Location = New System.Drawing.Point(12, StatusLabel.Bottom + verticalPadding)
            PartialTextLabel.Size = New Size(770, 60) ' Wider, not shrinking

            ' Position RichTextBox1 below PartialTextLabel with extra spacing
            RichTextBox1.Location = New System.Drawing.Point(12, PartialTextLabel.Bottom + verticalPadding)
            RichTextBox1.Size = New Size(770, 350)

            ' Position ButtonPanel below RichTextBox1
            ButtonPanel.Location = New System.Drawing.Point(12, RichTextBox1.Bottom + verticalPadding)
            ButtonPanel.Size = New Size(770, 45)

            ' Define padding values
            Dim buttonPadding As Integer = 10 ' Space between buttons inside the panel
            Dim buttonTopMargin As Integer = 5 ' Vertical margin from the top of ButtonPanel

            StartButton.Location = New System.Drawing.Point(0, buttonTopMargin)
            StopButton.Location = New System.Drawing.Point(StartButton.Right + buttonPadding, buttonTopMargin)
            ClearButton.Location = New System.Drawing.Point(StopButton.Right + buttonPadding, buttonTopMargin)
            LoadButton.Location = New System.Drawing.Point(ClearButton.Right + buttonPadding, buttonTopMargin)
            QuitButton.Location = New System.Drawing.Point(LoadButton.Right + buttonPadding, buttonTopMargin)
            ProcessButton.Location = New System.Drawing.Point(QuitButton.Right + buttonPadding + buttonPadding, buttonTopMargin)

            processCombobox.Location = New System.Drawing.Point(ProcessButton.Right + buttonPadding, buttonTopMargin)
            processCombobox.Size = New Size(250, 21)

            ButtonPanel.Size = New Size(770, buttonTopMargin + verticalPadding + StartButton.Height)
            ButtonPanel.Padding = New Padding(buttonPadding)

            Me.ClientSize = New Drawing.Size(800, ButtonPanel.Bottom + verticalPadding)

            ' Ensure the form cannot be resized closer than 20 pixels to the right of SpeakerDistance
            Dim minWidth As Integer = SpeakerDistance.Left + SpeakerDistance.Width + 40

            ' Set Minimum Size
            Me.MinimumSize = New Size(minWidth, Me.Height)


            ' Add elements to Form
            Me.Controls.Add(Label1)
            Me.Controls.Add(cultureComboBox)
            Me.Controls.Add(Label2)
            Me.Controls.Add(deviceComboBox)
            Me.Controls.Add(SpeakerIdent)
            Me.Controls.Add(SpeakerDistance)
            Me.Controls.Add(StatusLabel)
            Me.Controls.Add(PartialTextLabel)
            Me.Controls.Add(RichTextBox1)
            Me.Controls.Add(ButtonPanel)

            ' Add buttons to panel
            ButtonPanel.Controls.Add(StartButton)
            ButtonPanel.Controls.Add(StopButton)
            ButtonPanel.Controls.Add(ClearButton)
            ButtonPanel.Controls.Add(LoadButton)
            ButtonPanel.Controls.Add(QuitButton)
            ButtonPanel.Controls.Add(ProcessButton)
            ButtonPanel.Controls.Add(processCombobox)

            ' Icon
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        End Sub

        Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
            Dim minWidth As Integer = SpeakerDistance.Left + SpeakerDistance.Width + 40
            If Me.Width < minWidth Then
                Me.Width = minWidth ' Force minimum width dynamically
            End If
        End Sub


        Private Sub StopRecording()

            STTCanceled = True

            CancelTranscription()

            audioBuffer.Clear()

            If WhisperRecognizer IsNot Nothing Then
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper stopped...")
                WhisperRecognizer.DisposeAsync()
                WhisperRecognizer = Nothing
            End If

            If waveIn IsNot Nothing Then
                waveIn.StopRecording()
                RemoveHandler waveIn.DataAvailable, AddressOf OnAudioDataAvailable
            End If

            If MultiSource Then
                MultiSource = False

                stopProcessing = True

                If audioThread IsNot Nothing AndAlso audioThread.IsAlive Then
                    audioThread.Join() ' Wait for the audio thread to finish
                End If

                For Each waveInDevice As WaveInEvent In waveInputs
                    waveInDevice.StopRecording()
                    waveInDevice.Dispose()
                Next
                waveInputs.Clear()

                For Each bufferProvider As BufferedWaveProvider In waveProviders
                    bufferProvider.ClearBuffer()
                Next
                waveProviders.Clear()

            Else
                If waveIn IsNot Nothing Then
                    waveIn.StopRecording()
                    waveIn.Dispose()
                End If
            End If
        End Sub

        Private Sub StopButton_Click(sender As Object, e As EventArgs)
            If capturing Then
                StopRecording()
                capturing = False

                Me.StartButton.Enabled = True
                Me.StopButton.Enabled = False
                Me.LoadButton.Enabled = True
                Me.cultureComboBox.Enabled = True
                Me.deviceComboBox.Enabled = True
                Me.SpeakerIdent.Enabled = True
                Me.SpeakerDistance.Enabled = True
                If STTModel = "vosk" Then
                    Addline(PartialTextLabel.Text)
                End If
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")
            End If
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
                    StopRecording()
                    capturing = False
                    Me.StartButton.Enabled = True
                    Me.StopButton.Enabled = False
                    Me.LoadButton.Enabled = True
                    If STTModel = "vosk" Then
                        Addline(PartialTextLabel.Text)
                    End If
                    PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")
                End If
            End If
        End Sub

        Private Sub QuitButton_Click(sender As Object, e As EventArgs)
            If capturing Then
                StopRecording()
                capturing = False
                Me.StartButton.Enabled = True
                Me.LoadButton.Enabled = True
                Me.StopButton.Enabled = False
                Addline(PartialTextLabel.Text)
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")
                capturing = False
            End If
            Me.Close()
        End Sub

        Private Sub LoadButton_Click(sender As Object, e As EventArgs)
            If capturing Then
                Return
            End If

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

            Try

                If Me.cultureComboBox.SelectedItem.ToString().StartsWith("ggml") Then
                    STTModel = "whisper"
                Else
                    STTModel = "vosk"
                End If

                Select Case STTModel

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
                        PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper Is listening And working... (no Partial results shown, please wait)")

                    Case Else
                        splash.Close()
                        ShowCustomMessageBox($"No valid model selected. Please Select a model.")
                        Return

                End Select

                My.Settings.LastAudioSource = Me.deviceComboBox.SelectedItem.ToString()
                My.Settings.LastSpeechModel = Me.cultureComboBox.SelectedItem.ToString()
                My.Settings.LastSpeakerEnabled = Me.SpeakerIdent.Checked
                similarityThreshold = Double.Parse(Me.SpeakerDistance.Text)
                If similarityThreshold = 0 Then similarityThreshold = 1.0
                If similarityThreshold < 0.2 Then similarityThreshold = 0.2
                If similarityThreshold > 2.5 Then similarityThreshold = 2.5
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
                splash.Close()

                Select Case STTModel
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
                Dim capabilities As WaveInCapabilities = WaveInEvent.GetCapabilities(i)
                deviceComboBox.Items.Add($"{i}: {capabilities.ProductName}")
            Next
            deviceComboBox.Items.Add($"{i}: Combine all devices ({waveIn.DeviceCount})")

            ' Select default device (if available)
            Dim lastAudioSource As String = My.Settings.LastAudioSource
            If Not String.IsNullOrEmpty(lastAudioSource) AndAlso deviceComboBox.Items.Contains(lastAudioSource) Then
                deviceComboBox.SelectedItem = lastAudioSource
            ElseIf deviceComboBox.Items.Count > 0 Then
                deviceComboBox.SelectedIndex = 0
            End If

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

                Debug.WriteLine("Vosk Recognizer initialized")

            End If

            recognizer.SetMaxAlternatives(0) ' Forces earlier finalization
            recognizer.SetWords(True) ' Enable word timestamps
            recognizer.SetPartialWords(True) ' Partial words emitted faster
        End Sub

        Private Sub StartWhisper(Optional language As String = "auto")
            Dim modelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), Me.cultureComboBox.SelectedItem.ToString())

            ' Load the model using WhisperFactory
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

        Private Sub StartButton_Click(sender As Object, e As EventArgs)

            If capturing Then
                Return
            End If

            Dim splash As New SplashScreen($"Loading model...")
            splash.Show()
            splash.Refresh()

            Try

                If Me.cultureComboBox.SelectedItem.ToString().StartsWith("ggml") Then
                    STTModel = "whisper"
                Else
                    STTModel = "vosk"
                End If

                Select Case STTModel

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
                If similarityThreshold = 0 Then similarityThreshold = 1.0
                If similarityThreshold < 0.5 Then similarityThreshold = 0.5
                If similarityThreshold > 2.5 Then similarityThreshold = 2.5
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
                splash.Close()

            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"There has been an error starting the transcription engine (Error: {ex.Message}).")

            End Try
        End Sub


        Private Function StartRecording() As Boolean

            Dim selectedDeviceIndex As Integer = deviceComboBox.SelectedIndex

            Debug.WriteLine($"Selected device index: {selectedDeviceIndex}")
            Debug.WriteLine($"Device count: {waveIn.DeviceCount}")

            If selectedDeviceIndex < waveIn.DeviceCount Then
                waveIn = New WaveInEvent() With {
                        .DeviceNumber = selectedDeviceIndex,
                        .WaveFormat = New WaveFormat(16000, 1)
                    }
                AddHandler waveIn.DataAvailable, AddressOf OnAudioDataAvailable
                waveIn.StartRecording()

                Return True

            Else

                Dim deviceCount As Integer = waveIn.DeviceCount
                If deviceCount = 0 Then

                    ShowCustomMessageBox($"No audio devices to mix found.")
                    Return False
                End If

                MultiSource = True

                ' Create a list to hold audio sources for the mixer
                Dim sources As New List(Of ISampleProvider)()

                ' Change waveInputs to store both WaveInEvent and WasapiCapture
                Dim waveInputs As New List(Of IDisposable)() ' Supports both WaveInEvent and WasapiCapture

                For i As Integer = 0 To deviceCount - 1


                    Dim deviceCaps As WaveInCapabilities = WaveInEvent.GetCapabilities(i)
                    Dim selectedWaveFormat As WaveFormat = New WaveFormat(16000, 1) ' Desired format (16kHz Mono)
                    Dim useWasapi As Boolean = False ' Initialize useWasapi as False
                    Dim bufferProvider As BufferedWaveProvider = Nothing
                    Dim floatProvider As ISampleProvider = Nothing

                    ' Check if the device supports 16 kHz, 1 channel before opening it
                    If Not IsFormatSupported(i, 16000, 1) Then
                        ' Get the device's default format instead
                        selectedWaveFormat = GetDeviceDefaultFormat(i)
                        Debug.WriteLine($"Device {i} does NOT support 16kHz Mono. Using default: {selectedWaveFormat.SampleRate} Hz, {selectedWaveFormat.Channels} channels.")
                    Else
                        Debug.WriteLine($"Device {i} supports 16kHz Mono.")
                    End If

                    Dim waveInDevice As WaveInEvent = Nothing
                    Try
                        useWasapi = False ' Explicitly reset before trying WaveInEvent

                        ' Attempt to open with WaveInEvent first (Shared Mode)
                        waveInDevice = New WaveInEvent() With {
                                    .DeviceNumber = i,
                                    .WaveFormat = selectedWaveFormat
                                }

                        bufferProvider = New BufferedWaveProvider(waveInDevice.WaveFormat) With {
                                    .BufferLength = 16384,
                                    .DiscardOnBufferOverflow = True
                                    }

                        ' Event handler for processing incoming audio data
                        AddHandler waveInDevice.DataAvailable, Sub(senderObj As Object, eventArgs As WaveInEventArgs)
                                                                   bufferProvider.AddSamples(eventArgs.Buffer, 0, eventArgs.BytesRecorded)
                                                               End Sub

                        waveInputs.Add(waveInDevice) ' Store for cleanup
                        waveProviders.Add(bufferProvider)
                        floatProvider = New Pcm16BitToSampleProvider(bufferProvider)

                    Catch ex As System.Exception
                        Debug.WriteLine($"WaveInEvent failed for device {i}: {ex.Message}")
                        useWasapi = True ' If WaveInEvent fails, fallback to WASAPI
                    End Try

                    If useWasapi Then
                        ' Fallback: Use WASAPI in shared mode
                        Debug.WriteLine($"Using WASAPI Capture for device {i} instead of WaveInEvent.")

                        Dim enumerator As New MMDeviceEnumerator()
                        Dim captureDevices As MMDeviceCollection = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)

                        If i < captureDevices.Count Then
                            Dim captureDevice As MMDevice = captureDevices(i)
                            Dim wasapiCapture As New WasapiCapture(captureDevice) ' Shared mode by default
                            wasapiCapture.StartRecording()

                            bufferProvider = New BufferedWaveProvider(wasapiCapture.WaveFormat) With {
                                    .BufferLength = 16384,
                                    .DiscardOnBufferOverflow = True
}

                            ' Event handler for processing incoming audio data
                            AddHandler wasapiCapture.DataAvailable, Sub(senderObj As Object, eventArgs As WaveInEventArgs)
                                                                        bufferProvider.AddSamples(eventArgs.Buffer, 0, eventArgs.BytesRecorded)
                                                                    End Sub
                            floatProvider = New Pcm16BitToSampleProvider(bufferProvider)

                            ' Store WasapiCapture separately since it's not a WaveInEvent
                            waveInputs.Add(wasapiCapture)
                            waveProviders.Add(bufferProvider)
                        End If
                    End If

                    ' Resample if needed
                    If selectedWaveFormat.SampleRate <> 16000 OrElse selectedWaveFormat.Channels <> 1 Then
                        floatProvider = New WdlResamplingSampleProvider(floatProvider, 16000)
                        floatProvider = New MonoToStereoSampleProvider(floatProvider) ' Convert to mono if needed
                    End If

                    ' Add to mixer sources
                    sources.Add(floatProvider)

                    Debug.WriteLine($"Added input device {i}: {deviceCaps.ProductName}")

                Next


                ' Start recording on all devices
                For Each waveInDevice As WaveInEvent In waveInputs
                    waveInDevice.StartRecording()
                Next

                ' Create the mixer with all floating-point audio sources
                waveMixer = New MixingSampleProvider(sources) With {
                            .ReadFully = True
                        }

                ' Start processing audio buffer in a separate thread
                stopProcessing = False
                audioThread = New Thread(AddressOf ProcessAudio)
                audioThread.Start()

                Return True

            End If


        End Function


        Function IsFormatSupported(deviceIndex As Integer, sampleRate As Integer, channels As Integer) As Boolean
            Try
                Using testWaveIn As New WaveInEvent() With {
                        .DeviceNumber = deviceIndex,
                        .WaveFormat = New WaveFormat(sampleRate, channels)
                         }
                    ' If no exception occurs, the format is supported
                    Return True
                End Using
            Catch ex As System.Exception
                Return False ' Format not supported
            End Try
        End Function


        Function GetDeviceDefaultFormat(deviceIndex As Integer) As WaveFormat
            Try
                Dim enumerator As New MMDeviceEnumerator()
                Dim devices As MMDeviceCollection = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)

                If deviceIndex < devices.Count Then
                    Dim device As MMDevice = devices(deviceIndex)
                    Dim defaultFormat As WaveFormat = device.AudioClient.MixFormat
                    Return defaultFormat
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"Error getting default format for device {deviceIndex}: {ex.Message}")
            End Try

            ' Fallback format
            Return New WaveFormat(44100, 2)
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

        Private Sub ProcessAudio()
            'Dim buffer(4096 - 1) As Byte
            Dim buffer(16384 - 1) As Byte
            Dim waveProvider As New SampleToWaveProvider(waveMixer) ' Convert ISampleProvider to WaveProvider

            While Not stopProcessing
                Dim bytesRead As Integer = waveProvider.Read(buffer, 0, buffer.Length)
                If bytesRead > 0 Then
                    OnAudioDataAvailableMix(buffer, bytesRead)
                End If
                'Thread.Sleep(10) ' Allow some time for new audio to be buffered
            End While
        End Sub

        ' --- Process Audio ---

        Private Async Sub OnAudioDataAvailable(sender As Object, e As WaveInEventArgs)

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

        Private Async Sub OnAudioDataAvailableMix(buffer As Byte(), bytesRecorded As Integer)

            Select Case STTModel

                Case "vosk"
                    If recognizer IsNot Nothing AndAlso capturing Then
                        Dim jsonResult As String = If(recognizer.AcceptWaveform(buffer, bytesRecorded),
                                                  recognizer.Result,
                                                  recognizer.PartialResult)
                        ProcessTranscriptionJson(jsonResult)
                    End If

                Case "whisper"

                    If WhisperRecognizer Is Nothing Then Return

                    Try
                        ' Convert audio buffer to float array
                        Dim samples As Single() = ConvertAudioToFloat(buffer)

                        ' Append to buffer
                        audioBuffer.AddRange(samples)
                        ' Only process when buffer has enough data 
                        If audioBuffer.Count < 32000 Then Return ' Adjust threshold based on sample rate
                        ' Copy buffered audio and clear buffer
                        Dim processSamples = audioBuffer.ToArray()
                        audioBuffer.Clear()

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
                StopRecording()
                capturing = False
                Me.StartButton.Enabled = True
                Me.StopButton.Enabled = False
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
                StopRecording()
                capturing = False
                Me.StartButton.Enabled = True
                Me.StopButton.Enabled = False
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

            SyncLock finalText
                finalText.AppendLine(completedline)
            End SyncLock

            RichTextBox1.Invoke(Sub()
                                    RichTextBox1.AppendText(completedline & vbCrLf)
                                    PartialTextLabel.Text = ""
                                    If String.IsNullOrWhiteSpace(RichTextBox1.SelectedText) Then
                                        RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                        RichTextBox1.ScrollToCaret()
                                    End If
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

        If String.IsNullOrWhiteSpace(UserInput) Then Exit Sub

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
            Exit Sub
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


    Public Shared cts As New CancellationTokenSource()

    Public Shared Async Function GenerateAudioFromText(input As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O", Optional nossml As Boolean = False, Optional Pitch As Double = 0, Optional SpeakingRate As Double = 1) As Task(Of Byte())
        Try
            Using httpClient As New HttpClient()

                If TTSSecondAPI Then
                    DecodedAPI_2 = Await GetFreshAccessToken(_context, INI_OAuth2ClientMail_2, INI_OAuth2Scopes_2, INI_APIKey_2, INI_OAuth2Endpoint_2, INI_OAuth2ATExpiry_2, True)
                Else
                    DecodedAPI = Await GetFreshAccessToken(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, False)
                End If
                Dim AccessToken As String = If(TTSSecondAPI, DecodedAPI_2, DecodedAPI)
                If String.IsNullOrEmpty(AccessToken) Then
                    ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                    Return Nothing
                End If

                httpClient.DefaultRequestHeaders.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken)

                Dim requestBody As JObject

                'Debug.WriteLine(input)

                Dim jsonPayload As String

                If Trim(input).StartsWith("{") Then
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

                ' Allow for Esc to cancel
                Dim keyListenerTask As System.Threading.Tasks.Task = System.Threading.Tasks.Task.Run(Sub()
                                                                                                         While Not cts.Token.IsCancellationRequested
                                                                                                             If Console.KeyAvailable Then
                                                                                                                 Dim key As ConsoleKeyInfo = Console.ReadKey(True)
                                                                                                                 If key.Key = ConsoleKey.Escape Then
                                                                                                                     cts.Cancel()
                                                                                                                     Exit While
                                                                                                                 End If
                                                                                                             End If
                                                                                                             Thread.Sleep(100) ' Reduce CPU usage
                                                                                                         End While
                                                                                                     End Sub)

                Try
                    ' Make API request

                    If Len(input) > TTSLargeText Then
                        Dim t As New Thread(Sub()
                                                ShowCustomMessageBox("Audio generation has started and runs in the background. Press 'Esc' to abort.).", "", 3, "", True)
                                            End Sub)
                        t.SetApartmentState(ApartmentState.STA)
                        t.Start()
                    End If

                    Dim response As HttpResponseMessage = Await httpClient.PostAsync(INI_TTSEndpoint & "text:synthesize", content, cts.Token).ConfigureAwait(False)

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
                            Return Convert.FromBase64String(audioBase64)
                        Else
                            ShowCustomMessageBox("Error generating audio: 'audioContent' not found in response.")
                            Return Nothing
                        End If
                    Else
                        ShowCustomMessageBox($"Error generating audio: API returned status {response.StatusCode}. Response: {responseString}")
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


    Async Sub GenerateAndPlayPodcastAudio(conversation As List(Of Tuple(Of String, String)), filepath As String, languagecode As String, hostVoice As String, guestVoice As String, pitch As Double, speakingrate As Double, nossml As Boolean)
        Try
            Dim outputFiles As New List(Of String)

            If String.IsNullOrWhiteSpace(filepath) Then filepath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)

            If languagecode = "" Then languagecode = "en-US"
            If hostVoice = "" Then hostVoice = "en-US-Studio-O"
            If guestVoice = "" Then guestVoice = "en-US-Casual-K"

            Dim Exited As Boolean = False


            Using httpClient As New HttpClient()
                If TTSSecondAPI Then
                    DecodedAPI_2 = Await GetFreshAccessToken(_context, INI_OAuth2ClientMail_2, INI_OAuth2Scopes_2, INI_APIKey_2, INI_OAuth2Endpoint_2, INI_OAuth2ATExpiry_2, True)
                Else
                    DecodedAPI = Await GetFreshAccessToken(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, False)
                End If
                Dim AccessToken As String = If(TTSSecondAPI, DecodedAPI_2, DecodedAPI)
                If String.IsNullOrEmpty(AccessToken) Then
                    ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                    Exit Sub
                End If

                httpClient.DefaultRequestHeaders.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken)

                Dim t As New Thread(Sub()
                                        ShowCustomMessageBox("Audio generation has started and runs in the background. Press 'Esc' to abort.).", "", 3, "", True)
                                    End Sub)
                t.SetApartmentState(ApartmentState.STA)
                t.Start()

                ' Process each speaker separately
                For i As Integer = 0 To conversation.Count - 1

                    System.Windows.Forms.Application.DoEvents()
                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                        Exited = True
                        Exit For
                    End If
                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                        Exited = True
                        Exit For
                    End If

                    Dim speaker As String = conversation(i).Item1
                    Dim text As String = conversation(i).Item2

                    ' Select voice based on speaker
                    Dim voiceName As String = If(speaker = "H", hostVoice, guestVoice)

                    Dim textlabel As String = "text"
                    Dim ssmlPattern As String = "<[^>]+>"

                    If nossml Then
                        text = Regex.Replace(text, ssmlPattern, String.Empty)
                    Else
                        If Regex.IsMatch(text, ssmlPattern) AndAlso Not text.Trim().StartsWith("<speak>") Then
                            If Not text.Trim().StartsWith("<speak>") Then
                                text = "<speak>" & text & "</speak>"
                            End If
                            textlabel = "ssml"
                        End If
                    End If

                    ' Create request
                    Dim requestBody As New JObject From {
                    {"input", New JObject From {{$"{textlabel}", text}}},
                    {"voice", New JObject From {
                        {"languageCode", languagecode},
                        {"name", voiceName}
                    }},
                    {"audioConfig", New JObject From {
                        {"audioEncoding", "MP3"},
                        {"pitch", pitch},
                        {"speakingRate", speakingrate},
                        {"effectsProfileId", New JArray("small-bluetooth-speaker-class-device")}
                    }}
                }


                    Dim jsonPayload As String = requestBody.ToString()
                    Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")

                    Dim response As HttpResponseMessage = Await httpClient.PostAsync(INI_TTSEndpoint & "text:synthesize", content)
                    Dim responseString As String = Await response.Content.ReadAsStringAsync()
                    Dim responseJson As JObject = JObject.Parse(responseString)

                    If responseJson.ContainsKey("audioContent") Then
                        Dim audioBase64 As String = responseJson("audioContent").ToString()
                        Dim audioBytes As Byte() = Convert.FromBase64String(audioBase64)

                        ' Save each audio snippet separately
                        Dim tempFile As String = Path.Combine(ExpandEnvironmentVariables("%TEMP%"), $"{AN2}_podcast_temp_{i}.mp3")
                        File.WriteAllBytes(tempFile, audioBytes)
                        outputFiles.Add(tempFile)
                    End If

                    Await System.Threading.Tasks.Task.Delay(1000) ' Delay to not overhwelm the API

                Next

                If Not Exited Then
                    ' Merge all audio files into one
                    MergeAudioFiles(outputFiles, filepath)
                End If
                ' Cleanup temp files
                For Each file In outputFiles
                    System.IO.File.Delete(file)
                Next
            End Using
            If Exited Then
                ShowCustomMessageBox("Multi-speaker audio generation aborted.")
                Exit Sub
            Else
                Try
                    Dim Result As Integer = ShowCustomYesNoBox($"Your multi-speaker audio sequence has been generated ('{filepath}') and is ready to be played. Play it?", "Yes", "No (file remains available)")
                    If Result = 1 Then
                        PlayAudio(filepath)
                    End If
                Catch ex As System.Exception

                End Try
            End If
            Exit Sub
        Catch ex As Exception
            Debug.WriteLine($"Error generating podcast audio: {ex.Message}")
            Exit Sub
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
            Using frm As New TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voice you wish to use for creating your audio file and configure where to save it.", $"{AN} Google Text-to-Speech - Select Voices", true)
                If frm.ShowDialog() = DialogResult.OK Then
                    Dim selectedVoices As List(Of String) = frm.SelectedVoices
                    Dim selectedLanguage As String = frm.SelectedLanguage
                    Dim outputPath As String = frm.SelectedOutputPath

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


    Public Async Sub GenerateAndPlayAudioFromSelectionParagraphs(filepath As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O")
        Try

            Dim Temporary As Boolean = (filepath = "")

            If Temporary Then
                filepath = System.IO.Path.Combine(ExpandEnvironmentVariables("%TEMP%"), $"{AN2}_temp.mp3")
            End If

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

            ' Create an array of InputParameter objects.
            Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Pitch", Pitch),
                    New SLib.InputParameter("Speaking Rate", SpeakingRate),
                    New SLib.InputParameter("No SSML", NoSSML),
                    New SLib.InputParameter("Title Numbers", ReadTitleNumbers)
                    }

            ' Call the procedure (the parameters are passed ByRef).
            If Not ShowCustomVariableInputForm("Please enter the following parameters to apply when creating your audio file based on your text:", $"Create Audio", params) Then Return


            Pitch = CDbl(params(0).Value)
            SpeakingRate = CDbl(params(1).Value)
            NoSSML = CBool(params(2).Value)
            ReadTitleNumbers = CBool(params(3).Value)

            My.Settings.NoSSML = NoSSML
            My.Settings.Pitch = Pitch
            My.Settings.Speakingrate = SpeakingRate
            My.Settings.Save()

            Dim totalParagraphs As Integer = selection.Paragraphs.Count
            Dim tempFiles As New List(Of String)
            Dim paragraphIndex As Integer = 0
            Dim sentenceEndPunctuation As String() = {".", "!", "?", ";", ":", ",", ")", "]", "}"}
            Dim bracketedTextPattern As String = "^\s*[\(\[\{][^\)\]\}]*[\)\]\}]\s*$"

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
                    ' Also treat very short paragraphs (one or two lines) as titles.
                    Dim lines() As String = paraText.Split({vbLf}, StringSplitOptions.RemoveEmptyEntries)
                    If lines.Length <= 2 Then isTitle = True
                End If

                ' Set the maximum value if you know the total number of steps.
                GlobalProgressMax = totalParagraphs

                ' Update the current progress value and status label.
                GlobalProgressValue = paragraphIndex + 1
                GlobalProgressLabel = $"Paragraph {paragraphIndex + 1} of {totalParagraphs} (some may be skipped)"

                ' For bullet lists, insert a short pause BEFORE the paragraph.
                If isBullet Then
                    Dim silenceFileBefore As String = Await GenerateSilenceAudioFileAsync(0.3)
                    If Not String.IsNullOrEmpty(silenceFileBefore) Then tempFiles.Add(silenceFileBefore)
                End If

                ' Generate the audio for the paragraph via your TTS API.
                Dim paragraphAudioBytes As Byte() = Await GenerateAudioFromText(paraText, languageCode, voiceName, NoSSML, Pitch, SpeakingRate)
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
                ' • Use a medium pause (1.0 sec) for titles.
                ' • Otherwise use a short pause (0.5 sec).
                If isTitle Then
                    Dim silenceFileTitle As String = Await GenerateSilenceAudioFileAsync(1.0)
                    If Not String.IsNullOrEmpty(silenceFileTitle) Then tempFiles.Add(silenceFileTitle)
                Else
                    Dim silenceFileRegular As String = Await GenerateSilenceAudioFileAsync(0.5)
                    If Not String.IsNullOrEmpty(silenceFileRegular) Then tempFiles.Add(silenceFileRegular)
                End If

                Await System.Threading.Tasks.Task.Delay(1000) ' Delay to not overhwelm the API

                paragraphIndex += 1
            Next

            ' If no valid paragraphs were found, notify the user.
            If tempFiles.Count = 0 Then
                ShowCustomMessageBox("No valid paragraphs found for audio generation; skipping empty ones and {...}, [...] and (...).")
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
            ShowCustomMessageBox($"Error generating audio from selected paragraphs ({ex.Message}).")
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
        Private _context As ISharedContext ' or your actual type
        Private INI_OAuth2ClientMail As String
        Private INI_OAuth2Scopes As String
        Private INI_APIKey As String
        Private INI_OAuth2Endpoint As String
        Private INI_OAuth2ATExpiry As Long

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

        ' --- Constructor ---
        Public Sub New(context As ISharedContext,
                   clientMail As String,
                   scopes As String,
                   apiKey As String,
                   oauth2Endpoint As String,
                   oauth2Expiry As Long,
                   topLabelText As String,
                   formTitle As String,
                   twoVoicesRequired As Boolean)
            ' Assign external parameters
            _context = context
            INI_OAuth2ClientMail = clientMail
            INI_OAuth2Scopes = scopes
            INI_APIKey = apiKey
            INI_OAuth2Endpoint = oauth2Endpoint
            INI_OAuth2ATExpiry = oauth2Expiry

            ' Store our extra parameters
            _topLabelText = topLabelText
            _formTitle = formTitle
            _twoVoicesRequired = twoVoicesRequired

            ' Form settings
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Me.Text = _formTitle
            Me.Font = New Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.Size = New Size(660, 500)
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False

            ' Create controls (all existing ones plus new ones for radio buttons and output path)
            CreateControls()
            LayoutControls()
            AddHandlers()

            PopulateLanguageComboBoxes()
            LoadSettingsAndVoices()

            txtSampleText.Text = If(String.IsNullOrEmpty(My.Settings.TTSSampleText),
                                  $"Hello, I am talking using {_formTitle}!",
                                  My.Settings.TTSSampleText)
        End Sub

        ' --- CreateControls ---
        Private Sub CreateControls()
            ' Top label (autosized)
            lblIntro = New Label() With {
            .Text = _topLabelText,
            .AutoSize = True
        }

            ' Column 1
            lblSet1 = New Label() With {
            .Text = "Your default voice set 1:",
            .AutoSize = True
        }
            cmbLanguage1 = New Forms.ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Width = 200
        }
            cmbVoice1A = New Forms.ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Width = 200
        }
            btnPlay1A = New Forms.Button() With {
            .Width = 50
        }
            cmbVoice1B = New Forms.ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Width = 200
        }
            btnPlay1B = New Forms.Button() With {
            .Width = 50
        }

            ' Column 2
            lblSet2 = New Label() With {
            .Text = "Your default voice set 2:",
            .AutoSize = True
        }
            cmbLanguage2 = New Forms.ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Width = 200
        }
            cmbVoice2A = New Forms.ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Width = 200
        }
            btnPlay2A = New Forms.Button() With {
            .Width = 50
        }
            cmbVoice2B = New Forms.ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Width = 200
        }
            btnPlay2B = New Forms.Button() With {
            .Width = 50
        }

            ' Sample text
            lblSampleText = New Label() With {
            .Text = "Sample text:",
            .AutoSize = True
        }
            txtSampleText = New Forms.TextBox() With {
            .Text = "",
            .Width = 467
        }

            ' Bottom buttons
            btnOK = New Forms.Button() With {
            .Text = "OK",
            .AutoSize = True
        }
            btnCancel = New Forms.Button() With {
            .Text = "Cancel",
            .AutoSize = True
        }
            btnDesktop = New Forms.Button() With {
            .Text = "Save on Desktop",
            .AutoSize = True
            }

            ' Add the basic controls to the form
            Me.Controls.Add(lblIntro)
            Me.Controls.Add(lblSet1)
            Me.Controls.Add(cmbLanguage1)
            Me.Controls.Add(cmbVoice1A)
            Me.Controls.Add(btnPlay1A)
            Me.Controls.Add(cmbVoice1B)
            Me.Controls.Add(btnPlay1B)

            Me.Controls.Add(lblSet2)
            Me.Controls.Add(cmbLanguage2)
            Me.Controls.Add(cmbVoice2A)
            Me.Controls.Add(btnPlay2A)
            Me.Controls.Add(cmbVoice2B)
            Me.Controls.Add(btnPlay2B)

            Me.Controls.Add(lblSampleText)
            Me.Controls.Add(txtSampleText)
            Me.Controls.Add(btnOK)
            Me.Controls.Add(btnCancel)
            Me.Controls.Add(btnDesktop)

            ' --- Create radio buttons for voice selection ---
            If Not _twoVoicesRequired Then
                ' ONE VOICE mode: all four radio buttons belong to one group.
                rdoVoice1A = New RadioButton() With {.AutoSize = True}
                rdoVoice1B = New RadioButton() With {.AutoSize = True}
                rdoVoice2A = New RadioButton() With {.AutoSize = True}
                rdoVoice2B = New RadioButton() With {.AutoSize = True}
                ' Add them directly to the form.
                Me.Controls.Add(rdoVoice1A)
                Me.Controls.Add(rdoVoice1B)
                Me.Controls.Add(rdoVoice2A)
                Me.Controls.Add(rdoVoice2B)
                ' Default selection
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
                ' TWO VOICES mode: only one radio button per voice set.
                ' Create a single radio button for Voice Set 1
                rdoVoice1A = New RadioButton() With {.AutoSize = True}
                Me.Controls.Add(rdoVoice1A)
                rdoVoice1A.Checked = True

                ' Create a single radio button for Voice Set 2
                rdoVoice2A = New RadioButton() With {.AutoSize = True}
                Me.Controls.Add(rdoVoice2A)
                Select Case My.Settings.TTSLastRdoTwoVoices
                    Case "Voice1"
                        rdoVoice1A.Checked = True
                    Case "Voice2"
                        rdoVoice2A.Checked = True
                    Case Else
                        rdoVoice1A.Checked = True ' Default if no previous selection
                End Select
            End If



            ' --- Change play buttons to display Webdings character 52 ---
            Dim webdingsFont As New Drawing.Font("Webdings", 9.0F)
            btnPlay1A.Font = webdingsFont
            btnPlay1B.Font = webdingsFont
            btnPlay2A.Font = webdingsFont
            btnPlay2B.Font = webdingsFont
            btnPlay1A.Text = ChrW(52)
            btnPlay1B.Text = ChrW(52)
            btnPlay2A.Text = ChrW(52)
            btnPlay2B.Text = ChrW(52)

            ' --- Create output path controls ---
            lblOutputPath = New Label() With {
            .Text = "Output (.mp3):",
            .AutoSize = True
        }
            txtOutputPath = New Forms.TextBox() With {
            .Text = My.Settings.TTSOutputPath,
            .Width = 330
        }
            chkTemporary = New Forms.CheckBox() With {
            .Text = "Temporary only",
            .AutoSize = True
        }
            ' When checked, the output path text box is disabled.
            AddHandler chkTemporary.CheckedChanged, AddressOf chkTemporary_CheckedChanged

            Me.Controls.Add(lblOutputPath)
            Me.Controls.Add(txtOutputPath)
            Me.Controls.Add(chkTemporary)
        End Sub

        ' --- LayoutControls ---
        Private Sub LayoutControls()
            Dim marginLeft As Integer = 20
            Dim spacingY As Integer = 8
            Dim currentTop As Integer = 20

            ' Top label (autosized)
            lblIntro.Left = marginLeft
            lblIntro.Top = currentTop
            currentTop = lblIntro.Bottom + spacingY + 4

            ' Layout for Voice Set 1 voice choices (two rows)
            Dim radioWidth As Integer = 25 ' space reserved for radio buttons
            Dim controlSpacingX As Integer = 5

            ' Voice Set 1 label and language combo
            lblSet1.Left = marginLeft
            lblSet1.Top = currentTop
            currentTop = lblSet1.Bottom + 5

            cmbLanguage1.Left = marginLeft + radioWidth
            cmbLanguage1.Top = currentTop
            If _twoVoicesRequired Then
                rdoVoice1A.Left = marginLeft
                rdoVoice1A.Top = currentTop + 6
            End If

            ' First row for set 1
            currentTop = cmbLanguage1.Bottom + spacingY
            cmbVoice1A.Left = marginLeft + radioWidth
            cmbVoice1A.Top = currentTop
            btnPlay1A.Left = cmbVoice1A.Right + controlSpacingX
            btnPlay1A.Top = cmbVoice1A.Top
            If Not _twoVoicesRequired Then
                rdoVoice1A.Left = marginLeft
                rdoVoice1A.Top = currentTop + 6
            End If

            ' Second row for set 1
            currentTop = cmbVoice1A.Bottom + spacingY
            cmbVoice1B.Left = marginLeft + radioWidth
            cmbVoice1B.Top = currentTop
            btnPlay1B.Left = cmbVoice1B.Right + controlSpacingX
            btnPlay1B.Top = cmbVoice1B.Top
            If Not _twoVoicesRequired Then
                rdoVoice1B.Left = marginLeft
                rdoVoice1B.Top = currentTop + 6
            End If

            ' Now, Voice Set 2: position it to the right of set 1.
            Dim set2Left As Integer = btnPlay1A.Right + 30

            currentTop = lblSet1.Top ' align with set 1 label
            lblSet2.Left = set2Left
            lblSet2.Top = currentTop
            currentTop = lblSet2.Bottom + 5

            cmbLanguage2.Left = set2Left + radioWidth
            cmbLanguage2.Top = currentTop
            If _twoVoicesRequired Then
                rdoVoice2A.Left = set2Left
                rdoVoice2A.Top = currentTop + 6
            End If

            ' First row for set 2
            currentTop = cmbLanguage2.Bottom + spacingY
            If Not _twoVoicesRequired Then
                rdoVoice2A.Left = set2Left
                rdoVoice2A.Top = currentTop + 6
            End If
            cmbVoice2A.Left = set2Left + radioWidth
            cmbVoice2A.Top = currentTop
            btnPlay2A.Left = cmbVoice2A.Right + controlSpacingX
            btnPlay2A.Top = cmbVoice2A.Top

            ' Second row for set 2
            currentTop = cmbVoice2A.Bottom + spacingY
            If Not _twoVoicesRequired Then
                rdoVoice2B.Left = set2Left
                rdoVoice2B.Top = currentTop + 6
            End If
            cmbVoice2B.Left = set2Left + radioWidth
            cmbVoice2B.Top = currentTop
            btnPlay2B.Left = cmbVoice2B.Right + controlSpacingX
            btnPlay2B.Top = cmbVoice2B.Top

            ' Next, the Sample Text row
            Dim sampleLeft As Integer = marginLeft
            currentTop = Math.Max(cmbVoice1B.Bottom, cmbVoice2B.Bottom) + 20
            lblSampleText.Left = sampleLeft
            lblSampleText.Top = currentTop
            txtSampleText.Left = lblSampleText.Right + 35
            txtSampleText.Top = lblSampleText.Top - 3

            ' Now, add the Output Path row (below sample text)
            currentTop = txtSampleText.Bottom + 20
            lblOutputPath.Left = marginLeft
            lblOutputPath.Top = currentTop
            txtOutputPath.Left = lblSampleText.Right + 35
            txtOutputPath.Top = currentTop - 3
            chkTemporary.Left = txtOutputPath.Right + 10
            chkTemporary.Top = currentTop

            ' Finally, the bottom OK / Cancel buttons
            currentTop = txtOutputPath.Bottom + 30
            btnOK.Left = marginLeft
            btnOK.Top = currentTop
            btnCancel.Left = btnOK.Right + 10
            btnCancel.Top = btnOK.Top
            btnDesktop.Left = btnCancel.Right + 10
            btnDesktop.Top = btnOK.Top

            ' Adjust overall form height if needed.
            Me.Height = btnCancel.Bottom + 60
        End Sub

        ' --- AddHandlers ---
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

        Private Async Sub PopulateLanguageComboBoxes()

            ' Populate combos
            cmbLanguage1.Items.Clear()
            cmbLanguage2.Items.Clear()

            For Each lang In GoogleTTSsupportedLanguages
                cmbLanguage1.Items.Add(lang)
                cmbLanguage2.Items.Add(lang)
            Next
        End Sub
        Private Async Sub LoadSettingsAndVoices()
            RemoveHandler cmbLanguage1.SelectedIndexChanged, AddressOf cmbLanguage1_SelectedIndexChanged
            RemoveHandler cmbLanguage2.SelectedIndexChanged, AddressOf cmbLanguage2_SelectedIndexChanged

            cmbLanguage1.SelectedItem = If(IsNothing(My.Settings.TTS1languagecode), "", My.Settings.TTS1languagecode)
            cmbLanguage2.SelectedItem = If(IsNothing(My.Settings.TTS2languagecode), "", My.Settings.TTS2languagecode)

            Dim tasks As New List(Of System.Threading.Tasks.Task)

            If Not IsNothing(cmbLanguage1.SelectedItem) AndAlso cmbLanguage1.Text <> "" Then
                tasks.Add(LoadVoicesIntoComboBoxesAsync(cmbLanguage1.SelectedItem.ToString(), cmbVoice1A, cmbVoice1B))
            End If

            If Not IsNothing(cmbLanguage2.SelectedItem) AndAlso cmbLanguage2.Text <> "" Then
                tasks.Add(LoadVoicesIntoComboBoxesAsync(cmbLanguage2.SelectedItem.ToString(), cmbVoice2A, cmbVoice2B))
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


        Private Async Sub cmbLanguage1_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim selectedLang As String = TryCast(cmbLanguage1.SelectedItem, String)
            If Not String.IsNullOrEmpty(selectedLang) Then
                Await LoadVoicesIntoComboBoxesAsync(selectedLang, cmbVoice1A, cmbVoice1B)
            End If
        End Sub

        Private Async Sub cmbLanguage2_SelectedIndexChanged(sender As Object, e As EventArgs)
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

            If TTSSecondAPI Then
                DecodedAPI_2 = Await GetFreshAccessToken(_context, INI_OAuth2ClientMail_2, INI_OAuth2Scopes_2, INI_APIKey_2, INI_OAuth2Endpoint_2, INI_OAuth2ATExpiry_2, True)
            Else
                DecodedAPI = Await GetFreshAccessToken(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, False)
            End If
            Dim AccessToken As String = If(TTSSecondAPI, DecodedAPI_2, DecodedAPI)
            If String.IsNullOrEmpty(AccessToken) Then
                ShowCustomMessageBox("Error accessing Google API - authentication failed (no token).")
                Return Nothing
            End If

            ' Build request
            Dim url As String = INI_TTSEndpoint & "voices?languageCode=" & languageCode
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
                Await Threading.Tasks.Task.Run(Sub()
                                                   GenerateAndPlayAudio(sampleText, "", lang, voiceName)
                                               End Sub)
            Catch ex As System.Exception
                ShowCustomMessageBox("When trying to play the voice, an error occurred: " & ex.Message)
            End Try
        End Function

        ' --- OK / Cancel / Desktop event handlers ---
        Private Sub btnOK_Click(sender As Object, e As EventArgs)

            Dim NotAllSelected As Boolean = False

            ' Determine which voice(s) were selected based on radio buttons
            SelectedVoices.Clear()
            If Not _twoVoicesRequired Then
                ' ONE VOICE mode: the four radio buttons are one group.
                If rdoVoice1A.Checked Then
                    If cmbVoice1A.SelectedItem IsNot Nothing AndAlso cmbVoice1A.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice1A.SelectedItem.ToString())
                        SelectedLanguage = cmbLanguage1.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                ElseIf rdoVoice1B.Checked Then
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
                ElseIf rdoVoice2B.Checked Then
                    If cmbVoice2B.SelectedItem IsNot Nothing AndAlso cmbVoice2B.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice2B.SelectedItem.ToString())
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
                ShowCustomMessageBox("Please complete your voice selection (or 'Cancel').")
                Exit Sub
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
