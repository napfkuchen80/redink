' Red Ink for Outlook
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See License.txt or https://vischer.com/redink for more information.
'
' 22.6.2025
'
' The compiled version of Red Ink also ...
'
' Includes DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Appache-2.0 license (http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex).
' Includes Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json
' Includes HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier,Jeff Klawiter,Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/
' Includes Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/
' Includes PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig
' Includes MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license (https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig
' Includes NAudio in unchanged form; Copyright (c) 2020 Mark Heath; licensed under a proprietary open source license (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio
' Includes Vosk in unchanged form; Copyright (c) 2022 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes Grpc.core in unchanged form; Copyright (c) 2023 The gRPC Authors; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/grpc/grpc
' Includes Google Speech V1 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes Google Protobuf in unchanged form; Copyright (c) 2025 Google Inc.; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/protocolbuffers/protobuf
' Includes MarkdownToRTF in modified form; Copyright (c) 2025 Gustavo Hennig; original licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/GustavoHennig/MarkdownToRtf
' Includes Nito.AsyncEx in unchanged form; Copyright (c) 2021 Stephen Cleary; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/StephenCleary/AsyncEx
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.

Option Explicit On

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports Markdig
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic.FileIO
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports Nito.AsyncEx


Module Module1
    ' Correct attribute declaration for DllImport
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function
End Module

Public Class ThisAddIn

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function FindWindow(
                                ByVal lpClassName As String,
                                ByVal lpWindowName As String
                            ) As IntPtr
    End Function

    Public StartupInitialized As Boolean = False
    Private mainThreadControl As New System.Windows.Forms.Control()
    Private WithEvents outlookExplorer As Outlook.Explorer
    Private ReadOnly uiCtx As System.Threading.SynchronizationContext =
        System.Threading.SynchronizationContext.Current

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

        mainThreadControl.CreateControl()

        outlookExplorer = Application.ActiveExplorer()

        If outlookExplorer IsNot Nothing Then
            AddHandler outlookExplorer.Activate, AddressOf Explorer_Activate
        Else
            mainThreadControl.BeginInvoke(CType(AddressOf DelayedStartupTasks, MethodInvoker))
            StartupInitialized = True
        End If
    End Sub

    Private Sub Explorer_Activate()
        StartupInitialized = True
        RemoveHandler outlookExplorer.Activate, AddressOf Explorer_Activate
        DelayedStartupTasks()
    End Sub

    Private Sub DelayedStartupTasks()
        Try
            InitializeConfig(True, True)
            UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, "Outlook", INI_UpdatePath)
            Dim result = Globals.Ribbons.Ribbon1.UpdateRibbon()
            result = Globals.Ribbons.Ribbon2.UpdateRibbon()
            mainThreadControl.CreateControl()
            StartupHttpListener()
        Catch ex As System.Exception
            ' Handling errors gracefully
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ShutdownHttpListener()
    End Sub

    ' Hardcoded config values

    Public Const AN As String = "Red Ink"
    Public Const AN2 As String = "red_ink"

    Public Const Version As String = "V.220625 Gen2 Beta Test"

    ' Hardcoded configuration

    Public Const ShortenPercent As Integer = 20
    Public Const SummaryPercent As Integer = 20
    Private Const NetTrigger As String = "(net)"
    Private Const LibTrigger As String = "(Lib)"
    Private Const MarkupPrefix As String = "Markup:"
    Private Const MarkupPrefixDiff As String = "MarkupDiff:"
    Private Const MarkupPrefixDiffW As String = "MarkupDiffW:"
    Private Const MarkupPrefixWord As String = "MarkupWord:"
    Private Const MarkupPrefixAll As String = "Markup[Diff|DiffW|Word]:"
    Private Const ClipboardPrefix As String = "Clipboard:"
    Private Const InsertPrefix As String = "Insert:"
    Private Const NoFormatTrigger As String = "(noformat)"
    Private Const NoFormatTrigger2 As String = "(nf)"
    Private Const KFTrigger As String = "(keepformat)"
    Private Const KFTrigger2 As String = "(kf)"
    Private Const KPFTrigger As String = "(keepparaformat)"
    Private Const KPFTrigger2 As String = "(kpf)"
    Private Const InPlacePrefix As String = "Replace:"
    Private Const ObjectTrigger2 As String = "(clip)"

    Private Const ESC_KEY As Integer = &H1B

    Private Const SecondAPICode As String = "(2nd)"

    ' Variables that are available to InterpolateAtRuntime

    Public TranslateLanguage As String = ""
    Public OtherPrompt As String = ""
    Public ShortenLength, SummaryLength As Long
    Public DateTimeNow As String

    Public InspectorOpened As Boolean = False

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

    Public Shared ReadOnly Property RDV As String = "Outlook (" & Version & ")"
    Public Shared ReadOnly Property InitialConfigFailed As Boolean = False
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


    Public Shared Property INI_ISearch_Apply_SP As String
        Get
            Return _context.INI_ISearch_Apply_SP
        End Get
        Set(value As String)
            _context.INI_ISearch_Apply_SP = value
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


    Public Shared Property INI_Lib_Apply_SP As String
        Get
            Return _context.INI_Lib_Apply_SP
        End Get
        Set(value As String)
            _context.INI_Lib_Apply_SP = value
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


    '───────────────────────────────────────────────────────────────────────────
    ' Runs a Sub on the UI thread and *waits* for it to complete.
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUi(uiAction As System.Action) _
        As System.Threading.Tasks.Task

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Object)()

        mainThreadControl.Invoke(New MethodInvoker(
        Sub()
            Try
                uiAction.Invoke()
                tcs.SetResult(Nothing)
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task        ' completes only after uiAction finished
    End Function


    '───────────────────────────────────────────────────────────────────────────
    ' Runs a Function(Of T) on the UI thread and waits for its return value.
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUi(Of T)(uiFunc As System.Func(Of T)) _
        As System.Threading.Tasks.Task(Of T)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)()

        mainThreadControl.Invoke(New MethodInvoker(
        Sub()
            Try
                tcs.SetResult(uiFunc.Invoke())
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task        ' completes only after uiFunc returns
    End Function


    '───────────────────────────────────────────────────────────────────────────
    ' SwitchToUiTask  –  runs an *async* function (returns Task(Of T)) on the
    ' Outlook UI thread and gives you a Task(Of T) you can Await from any thread.
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUiTask(Of T)(
        uiFunc As System.Func(Of System.Threading.Tasks.Task(Of T))) _
        As System.Threading.Tasks.Task(Of T)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)()

        mainThreadControl.BeginInvoke(New MethodInvoker(
        Sub()
            Try
                Dim inner As System.Threading.Tasks.Task(Of T) = uiFunc.Invoke()
                inner.ContinueWith(
                    Sub(taskObj)
                        If taskObj.IsFaulted Then
                            tcs.SetException(taskObj.Exception.InnerExceptions)
                        ElseIf taskObj.IsCanceled Then
                            tcs.SetCanceled()
                        Else
                            tcs.SetResult(taskObj.Result)
                        End If
                    End Sub,
                    System.Threading.Tasks.TaskScheduler.Default)
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function


    Public Sub InitializeConfig(FirstTime As Boolean, Reload As Boolean)
        _context.InitialConfigFailed = False
        _context.RDV = "Outlook (" & Version & ")"
        SharedMethods.InitializeConfig(_context, FirstTime, Reload)
    End Sub
    Private Function INIValuesMissing()
        Return SharedMethods.INIValuesMissing(_context)
    End Function
    Public Shared Async Function PostCorrection(inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
        Return Await SharedMethods.PostCorrection(_context, inputText, UseSecondAPI)
    End Function

    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional HideSplash As Boolean = False, Optional ByVal AddUserPrompt As String = "", Optional ByVal FileObject As String = "") As Task(Of String)
        Return Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, HideSplash, AddUserPrompt, FileObject)
    End Function

    Private Function ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String))
        SharedMethods.ShowSettingsWindow(Settings, SettingsTips, _context)
    End Function
    Private Function ShowPromptSelector(filePath As String, enableMarkup As Boolean, enableBubbles As Boolean) As (String, Boolean, Boolean, Boolean)
        Return SharedMethods.ShowPromptSelector(filePath, enableMarkup, enableBubbles, _context)
    End Function

#End Region

    Enum Operation
        Insert = 1
        Delete = 2
        Equal = 3
    End Enum

    Public Sub MainMenu(RI_Command As String)

        If Not INIloaded Then
            If Not StartupInitialized Then
                Try
                    DelayedStartupTasks()
                    RemoveHandler outlookExplorer.Activate, AddressOf Explorer_Activate
                Catch ex As System.Exception
                End Try
                If Not INIloaded Then Exit Sub
            Else
                InitializeConfig(False, False)
                If Not INIloaded Then
                    Exit Sub
                End If
            End If
        End If

        Try
            ' Use fully qualified names to avoid ambiguity
            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = outlookApp.ActiveInspector
            Dim Textlength As Long

            If inspector Is Nothing Then

                InspectorOpened = False

                OpenInspectorAndReapplySelection(RI_Command = "Sumup")

                If Not InspectorOpened Then Exit Sub

                inspector = outlookApp.ActiveInspector
                If inspector Is Nothing Then

                    System.Windows.Forms.MessageBox.Show("Error in MainMenu: No active email item found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            If inspector.CurrentItem.Class = Microsoft.Office.Interop.Outlook.OlObjectClass.olMail Then
                Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)
                Dim wordEditor As Microsoft.Office.Interop.Word.Document = DirectCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

                InitializeConfig(False, False)

                If GPTSetupError OrElse INIValuesMissing() Or Not INIloaded Then Return

                Select Case RI_Command

                    Case "Translate"
                        TranslateLanguage = ""
                        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True)
                        If String.IsNullOrEmpty(TranslateLanguage) Then Return
                        Command_InsertAfter(InterpolateAtRuntime(SP_Translate), False, INI_KeepFormat1, INI_ReplaceText1)
                    Case "PrimLang"
                        TranslateLanguage = INI_Language1
                        Command_InsertAfter(InterpolateAtRuntime(SP_Translate), False, INI_KeepFormat1, INI_ReplaceText1)
                    Case "Correct"
                        Command_InsertAfter(InterpolateAtRuntime(SP_Correct), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Summarize"

                        Textlength = GetSelectedTextLength()

                        If Textlength = 0 Then
                            SLib.ShowCustomMessageBox("Please select the text to be processed.")
                            Exit Sub
                        End If

                        Dim UserInput As String
                        SummaryLength = 0

                        Do
                            UserInput = Trim(SLib.ShowCustomInputBox("Enter the number of words your summary shall have (the selected text has " & Textlength & " words; the proposal " & SummaryPercent & "%):", $"{AN} Summarizer", True, CStr(SummaryPercent * Textlength / 100)))

                            If String.IsNullOrEmpty(UserInput) Then
                                Exit Sub
                            End If

                            If Integer.TryParse(UserInput, SummaryLength) AndAlso SummaryLength >= 1 AndAlso SummaryLength <= Textlength Then
                                Exit Do
                            Else
                                SLib.ShowCustomMessageBox("Please enter a valid word count between 1 and " & Textlength & ".")
                            End If
                        Loop
                        If SummaryLength = 0 Then Exit Sub
                        'SummaryLength = (Textlength - (Textlength * SummaryPercent / 100))'

                        Command_InsertAfter(InterpolateAtRuntime(SP_Summarize), False)
                    Case "Improve"
                        Command_InsertAfter(InterpolateAtRuntime(SP_Improve), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "NoFillers"
                        Command_InsertAfter(InterpolateAtRuntime(SP_NoFillers), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Friendly"
                        Command_InsertAfter(InterpolateAtRuntime(SP_Friendly), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Convincing"
                        Command_InsertAfter(InterpolateAtRuntime(SP_Convincing), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Shorten"
                        Textlength = GetSelectedTextLength()
                        If Textlength = 0 Then
                            SLib.ShowCustomMessageBox("Please select the text to be processed.")
                            Exit Sub
                        End If
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
                                SLib.ShowCustomMessageBox("Please enter a valid percentage between 1 And 99.")
                            End If
                        Loop
                        ShortenLength = (Textlength - (Textlength * (100 - ShortenPercent) / 100))
                        Command_InsertAfter(InterpolateAtRuntime(SP_Shorten), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Sumup"

                        Dim selectedText As String = mailItem.Body
                        ShowSumup(selectedText)

                        'FreeStyle_InsertBefore(SP_MailSumup, False)
                    Case "Answers"
                        FreeStyle_InsertBefore(SP_MailReply, True)
                    Case "Freestyle"
                        FreeStyle_InsertAfter()
                    Case Else
                        System.Windows.Forms.MessageBox.Show("Error in MainMenu: Invalid internal command.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Select

            Else
                SLib.ShowCustomMessageBox($"Please open an email for editing for using {AN}.")
            End If
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show("Error in MainMenu: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Sub OpenInspectorAndReapplySelection(Sumup As Boolean)
        Try
            ' Grab Outlook instances
            Dim oApp As Outlook.Application = Globals.ThisAddIn.Application
            Dim oExplorer As Outlook.Explorer = oApp.ActiveExplorer()

            If oExplorer Is Nothing Then
                If Sumup Then
                    ShowCustomMessageBox("You can only use this function when you have selected an e-mail.")
                Else
                    ShowCustomMessageBox("You can only use this function when you are editing an e-mail.")
                End If
                Return
            End If

            ' Check for inline response
            Dim inlineResponse As Object = oExplorer.ActiveInlineResponse
            If inlineResponse Is Nothing Then

                ' Get the current selection in the explorer
                Dim selection As Outlook.Selection = oExplorer.Selection

                ' Check if any item is selected
                If selection.Count = 0 Then
                    ShowCustomMessageBox("No email is selected.")
                    Return
                End If

                If selection.Count > 1 Then
                    If Not Sumup Then
                        ShowCustomMessageBox("Multiple emails selected. Please select only one email when not using Sumup mode.")
                        Return
                    Else
                        ' Combine texts from all selected emails.
                        Dim mailItems As New List(Of Microsoft.Office.Interop.Outlook.MailItem)
                        For Each item As Object In selection
                            If TypeOf item Is Microsoft.Office.Interop.Outlook.MailItem Then
                                mailItems.Add(CType(item, Microsoft.Office.Interop.Outlook.MailItem))
                            End If
                        Next

                        If mailItems.Count = 0 Then
                            ShowCustomMessageBox("None of the selected items are emails.")
                            Return
                        End If

                        ' Order the emails: latest email first (descending order by ReceivedTime)
                        mailItems = mailItems.OrderByDescending(Function(m) m.ReceivedTime).ToList()

                        Const PR_LAST_VERB_EXECUTED As String = "http://schemas.microsoft.com/mapi/proptag/0x10810003"

                        Dim selectedText As String = String.Empty
                        Dim count As Integer = 1
                        For Each mail As Microsoft.Office.Interop.Outlook.MailItem In mailItems

                            Dim lastVerb As Integer = 0

                            Try
                                lastVerb = mail.PropertyAccessor.GetProperty(PR_LAST_VERB_EXECUTED)
                            Catch comEx As COMException
                                ' Property nicht gesetzt → noch nicht beantwortet
                                lastVerb = 0
                            Catch ex As System.Exception
                                ' Sicherstellen, dass System.Exception voll qualifiziert ist
                                lastVerb = 0
                            End Try


                            If lastVerb <> 102 AndAlso lastVerb <> 103 Then
                                Dim tag As String = count.ToString("D4") ' Format count with four digits
                                Dim latestBody As String = GetLatestMailBody(mail.Body)
                                selectedText &= "<EMAIL" & tag & ">" & latestBody & "</EMAIL" & tag & ">"
                                count += 1
                            End If
                        Next

                        ShowSumup2(selectedText)
                        Return
                    End If
                Else
                    ' Only one email is selected.
                    If Sumup Then
                        Dim selectedItem As Object = selection(1)
                        If TypeOf selectedItem Is Outlook.MailItem Then
                            Dim mail As Outlook.MailItem = CType(selectedItem, Outlook.MailItem)
                            Dim selectedText As String = mail.Body
                            ShowSumup(selectedText)
                            Return
                        Else
                            ShowCustomMessageBox("The selected item is not an email.")
                            Return
                        End If
                    Else
                        ShowCustomMessageBox("You can only use this function when you are editing one (single) e-mail.")
                        Return
                    End If
                End If

            End If

            ' Ensure it is a MailItem
            Dim mailItem As MailItem = TryCast(inlineResponse, MailItem)
            If mailItem Is Nothing Then
                ShowCustomMessageBox("You can only use this function when you are editing an e-mail (currently, there is no valid e-mail item).")
                Return
            End If

            ' Capture the user's current selection range (or caret) from the inline editor
            Dim oldSelStart As Integer = 0
            Dim oldSelEnd As Integer = 0
            If Not GetSelectionOrCaretRangeFromInlineEditor(oExplorer, oldSelStart, oldSelEnd) Then
                ' If this fails entirely (no Word editor, etc.), we can just open the window without reapplying.
                ' But no error is shown for "empty selection" anymore – only true failures (e.g., no WordEditor).
                ' We'll just continue and open the Inspector, albeit we can't set the cursor position.
            End If

            ' Open the Inspector modelessly
            Dim inspector As Inspector = mailItem.GetInspector
            If inspector Is Nothing Then
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: Failed to open the ActiveInspector.",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            inspector.Display(False) ' modeless - do not block

            ' A short delay to let the new WordEditor initialize
            System.Threading.Thread.Sleep(500)

            ' Ensure it's still open
            If inspector.CurrentItem Is Nothing Then
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: The Inspector window was closed before processing could complete.",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Reapply the original selection (or caret position) to the new Inspector's WordEditor
            Try
                Dim wordDoc As Word.Document = TryCast(inspector.WordEditor, Word.Document)
                If wordDoc IsNot Nothing Then
                    Dim wordSel As Word.Selection = wordDoc.Application.Selection

                    ' Only reapply if we successfully retrieved the inline offsets
                    If oldSelStart <> 0 OrElse oldSelEnd <> 0 Then
                        wordSel.SetRange(Start:=oldSelStart, End:=oldSelEnd)
                        wordSel.Select()
                    End If
                End If

            Catch ex As System.Exception
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: Failed to restore the original selection: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            ' Bring the new Inspector window to the foreground

            InspectorOpened = True

            inspector.Activate()

            ' Clean up COM references
            Marshal.ReleaseComObject(inspector)
            Marshal.ReleaseComObject(oExplorer)

            Return

        Catch ex As System.Exception
            MessageBox.Show("Error in OpenInspectorAndReapplySelection: " & ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetLatestMailBody(ByVal fullBody As String) As String
        Try
            ' Define an array of candidate markers that are common indicators of quoted messages,
            ' including localized variants.
            Dim markers() As String = {
            "-----Original Message-----",
            "-----Ursprüngliche Nachricht-----",
            "-----Vorherige Nachricht-----",
            "-----Mensaje original-----",
            "-----Messaggio originale-----",
            "-----Courrier original-----",
            "On ",
            "wrote:"
        }

            ' Regular expression to detect header lines with a proper email address
            Dim emailPattern As String = "^(From:|Von:|De:|Da:)\s+[\w\.-]+@[\w\.-]+\.\w+"

            ' Split the email body into lines
            Dim lines() As String = fullBody.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
            Dim sb As New StringBuilder()

            For i As Integer = 0 To lines.Length - 1
                Dim currentLine As String = lines(i)
                Dim trimmedLine As String = currentLine.TrimStart()
                Dim isChainMarker As Boolean = False

                ' First, check each line against our list of known chain markers.
                For Each marker As String In markers
                    If trimmedLine.StartsWith(marker, StringComparison.InvariantCultureIgnoreCase) Then
                        ' Only consider short lines (heuristically less than 100 characters) as markers.
                        If trimmedLine.Length < 100 Then
                            isChainMarker = True
                            Exit For
                        End If
                    End If
                Next

                ' If none of the above markers was found, try to detect headers indicating a quoted message.
                If Not isChainMarker Then
                    ' Check for email header markers using a regex pattern (with an @ symbol)
                    If Regex.IsMatch(trimmedLine, emailPattern, RegexOptions.IgnoreCase) Then
                        isChainMarker = True
                    Else
                        ' Additional check: headers with a name or parenthesized comment following the marker.
                        Dim headerMarkers() As String = {"From:", "Von:", "De:", "Da:"}
                        For Each header As String In headerMarkers
                            If trimmedLine.StartsWith(header, StringComparison.InvariantCultureIgnoreCase) Then
                                ' Extract the text after the header marker.
                                Dim remainingText As String = trimmedLine.Substring(header.Length).Trim()
                                ' Check if the remaining text contains a comma (e.g., "Doe, John") or a parenthesized group.
                                If remainingText.Contains(",") OrElse (remainingText.Contains("(") AndAlso remainingText.Contains(")")) Then
                                    isChainMarker = True
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If

                ' If a marker is confidently detected, assume the latest mail ends here.
                If isChainMarker Then
                    Return sb.ToString().TrimEnd()
                End If

                ' Otherwise, add the current line to the accumulated result.
                sb.AppendLine(currentLine)
            Next

            ' No clear marker found; return the full email content.
            Return fullBody
        Catch ex As System.Exception
            ' In case of any error, return the full email body
            ' (Alternatively, you could log the exception as needed)
            Return fullBody
        End Try
    End Function


    Private Function GetSelectionOrCaretRangeFromInlineEditor(oExplorer As Outlook.Explorer, ByRef selStart As Integer, ByRef selEnd As Integer) As Boolean
        Try
            Dim inlineWordEditor As Object = oExplorer.ActiveInlineResponseWordEditor
            If inlineWordEditor Is Nothing Then
                ' No inline Word editor, so we can't read a selection/caret
                Return False
            End If

            Dim wordSel As Word.Selection =
            TryCast(inlineWordEditor.Application.Selection, Word.Selection)
            If wordSel Is Nothing Then
                Return False
            End If

            ' Even if there's no highlighted text, there's always a caret position
            ' So we record them (could be equal if there's no actual selection)
            selStart = wordSel.Start
            selEnd = wordSel.End

            Return True

        Catch ex As System.Exception
            MessageBox.Show("Failed to retrieve the selection: " & ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Async Sub ShowSumup(selectedtext As String)

        Dim LLMResult As String = ""

        LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup), "<MAILCHAIN>" & selectedtext & "</MAILCHAIN>", "", "", 0)

        If INI_PostCorrection <> "" Then
            LLMResult = Await PostCorrection(LLMResult)
        End If

        'Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

        Dim builder As New MarkdownPipelineBuilder()

        builder.UsePipeTables()
        builder.UseGridTables()
        builder.UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras()
        builder.UseFootnotes()
        builder.UseDefinitionLists()
        builder.UseAbbreviations()
        builder.UseAutoLinks()
        builder.UseTaskLists()
        builder.UseEmojiAndSmiley()
        builder.UseMathematics()
        builder.UseFigures()
        builder.UseAdvancedExtensions()
        builder.UseGenericAttributes()

        Dim markdownPipeline As MarkdownPipeline = builder.Build()

        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        ShowHTMLCustomMessageBox(htmlText, $"{AN} Sum-up")

    End Sub

    Private Async Sub ShowSumup2(selectedtext As String)

        Dim LLMResult As String = ""

        DateTimeNow = DateTime.Now.ToString("yyyy-MMM-dd HH:mm")

        LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup2), selectedtext, "", "", 0)

        If INI_PostCorrection <> "" Then
            LLMResult = Await PostCorrection(LLMResult)
        End If

        ' Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

        Dim builder As New MarkdownPipelineBuilder()

        builder.UsePipeTables()
        builder.UseGridTables()
        builder.UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras()
        builder.UseFootnotes()
        builder.UseDefinitionLists()
        builder.UseAbbreviations()
        builder.UseAutoLinks()
        builder.UseTaskLists()
        builder.UseEmojiAndSmiley()
        builder.UseMathematics()
        builder.UseFigures()
        builder.UseAdvancedExtensions()
        builder.UseGenericAttributes()

        Dim markdownPipeline As MarkdownPipeline = builder.Build()

        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        ShowHTMLCustomMessageBox(htmlText, $"{AN} Sum-up (of unanswered mails)")

    End Sub



    Private Async Sub FreeStyle_InsertBefore(Command As String, Optional AskForPrompt As Boolean = False)
        Try
            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Inspector = outlookApp.ActiveInspector()

            ' Ensure the inspector is open and the item is a MailItem
            If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is MailItem Then
                SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
                Return
            End If

            Dim mailItem As MailItem = DirectCast(inspector.CurrentItem, MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = OlBodyFormat.olFormatPlain Then
                SLib.ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

            If wordEditor Is Nothing Then
                SLib.ShowCustomMessageBox("Unable to access the necessary email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If

            ' Get the selected text
            Dim selectedText As String = wordEditor.Application.Selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                selectedText = wordEditor.Content.Text
            End If

            OtherPrompt = ""
            Dim LLMResult As String = ""

            If AskForPrompt Then
                ' Prompt for additional instructions
                OtherPrompt = SLib.ShowCustomInputBox("Please provide additional instructions for drafting an answer (or leave it empty for the most likely substantive response):", $"{AN} Answers", False)

                ' Call your LLM function with the selected text
                LLMResult = Await LLM(InterpolateAtRuntime(SP_MailReply), "<MAILCHAIN>" & selectedText & "</MAILCHAIN>", "", "", 0)
            Else
                LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup), "<MAILCHAIN>" & selectedText & "</MAILCHAIN>", "", "", 0)
            End If
            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            'LLMResult = LLMResult.Replace("**", "")  ' Remove bold markers

            ' Convert Markdown to HTML using Markdig
            ' Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

            Dim builder As New MarkdownPipelineBuilder()

            builder.UsePipeTables()
            builder.UseGridTables()
            builder.UseSoftlineBreakAsHardlineBreak()
            builder.UseListExtras()
            builder.UseFootnotes()
            builder.UseDefinitionLists()
            builder.UseAbbreviations()
            builder.UseAutoLinks()
            builder.UseTaskLists()
            builder.UseEmojiAndSmiley()
            builder.UseMathematics()
            builder.UseFigures()
            builder.UseAdvancedExtensions()
            builder.UseGenericAttributes()

            Dim markdownPipeline As MarkdownPipeline = builder.Build()

            Dim convertedHtml As String = Markdown.ToHtml(LLMResult, markdownPipeline)

            If mailItem.BodyFormat = OlBodyFormat.olFormatHTML Then
                ' Ensure consistent font and style for HTML emails
                Dim defaultStyle As String = "<div style='font-family:Arial, sans-serif; font-size:11pt;'>" ' Adjust as needed
                Dim formattedResult As String = defaultStyle & convertedHtml & "</div><br/><br/>"

                ' Append the formatted result to the HTML body
                mailItem.HTMLBody = formattedResult & mailItem.HTMLBody
            Else
                ' Convert HTML to plain text for non-HTML formats (optional)
                Dim doc As New HtmlAgilityPack.HtmlDocument()
                doc.LoadHtml(convertedHtml)
                Dim plainTextResult As String = doc.DocumentNode.InnerText

                ' Standard handling for Plain Text and Rich Text
                mailItem.Body = plainTextResult & vbCrLf & vbCrLf & mailItem.Body
            End If

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle_InsertBefore: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub Command_InsertAfter(ByVal SysCommand As String, Optional ByVal DoMarkup As Boolean = False, Optional KeepFormat As Boolean = False, Optional Inplace As Boolean = False, Optional MarkupMethod As Integer = 3)
        Try
            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = outlookApp.ActiveInspector()

            ' Ensure the inspector is open and the item is a MailItem
            If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                ShowCustomMessageBox("Please open an email to use this function.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

            If wordEditor Is Nothing Then
                ShowCustomMessageBox("Unable to access the email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If

            ' Get the selected text and range
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim range As Microsoft.Office.Interop.Word.Range = selection.Range.Duplicate ' Duplicate to preserve original
            Dim SelectedText As String

            If INI_KeepFormatCap > 0 Then If Len(selection.Text) > INI_KeepFormatCap Then KeepFormat = False

            If KeepFormat Then
                SelectedText = SLib.GetRangeHtml(selection.Range)
            Else
                SelectedText = selection.Text
            End If

            If String.IsNullOrWhiteSpace(SelectedText) Then
                ShowCustomMessageBox($"Please select the text to be processed.")
                Return
            End If

            If DoMarkup And MarkupMethod = 2 And Len(SelectedText) > INI_MarkupDiffCap Then
                Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Diff markup method at {INI_MarkupDiffCap} chars (your selection has {Len(SelectedText)} chars). {If(KeepFormat, "This may be because HTML codes have been inserted to keep the formatting (you can turn this off in the settings). ", "")}. How do you want to continue?", "Use Diff in Window compare instead", "Use Diff")
                Select Case MarkupChange
                    Case 1
                        MarkupMethod = 3
                    Case 2
                        MarkupMethod = 2
                    Case Else
                        Exit Sub
                End Select
            End If

            Dim trailingCR As Boolean = SelectedText.EndsWith(vbCrLf)

            ' Call your LLM function with the selected text
            Dim LLMResult As String = Await LLM(SysCommand & If(KeepFormat, " " & SP_Add_KeepHTMLIntact, ""), "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>", "", "", 0)

            LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            ' Replace the selected text with the processed result
            If Not String.IsNullOrWhiteSpace(LLMResult) Then
                If KeepFormat Then

                    Dim Plaintext As String = ""

                    SelectedText = selection.Text
                    SLib.InsertTextWithFormat(LLMResult, range, Inplace)
                    If DoMarkup Then
                        LLMResult = SLib.RemoveHTML(LLMResult)
                        If MarkupMethod <> 3 Then
                            range.Text = vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf
                        End If
                        range.Collapse(WdCollapseDirection.wdCollapseEnd)
                        selection.SetRange(range.Start, selection.End)

                        CompareAndInsertText(SelectedText, LLMResult, MarkupMethod = 3, "This is the markup of the text inserted:", True)
                    End If

                Else

                    If Inplace Then
                        If Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                        If Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)
                        If DoMarkup And MarkupMethod <> 3 Then
                            selection.TypeText(LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                        Else
                            selection.TypeText(LLMResult)
                        End If
                    Else
                        selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        If DoMarkup And MarkupMethod <> 3 Then
                            'selection.TypeText(vbCrLf & LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                            SLib.InsertTextWithMarkdown(selection, vbCrLf & LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                        Else
                            'selection.TypeText(vbCrLf & LLMResult & vbCrLf)
                            SLib.InsertTextWithMarkdown(selection, vbCrLf & LLMResult & vbCrLf)

                        End If
                    End If

                    ' Use Find to locate the nearest line break backward and adjust selection
                    range = selection.Range
                    With range.Find
                        .Text = vbCrLf
                        .Forward = False
                        .MatchWildcards = False
                        If .Execute() Then
                            selection.SetRange(range.Start, selection.End)
                        End If
                    End With

                    ' Perform markup comparison and insertion if necessary
                    If DoMarkup Then
                        If MarkupMethod = 2 Or MarkupMethod = 3 Then
                            CompareAndInsertText(SelectedText, LLMResult, MarkupMethod = 3, "This is the markup of the text inserted:", True)
                        Else
                            CompareAndInsertTextCompareDocs(SelectedText, LLMResult)
                        End If

                    End If

                End If

            Else
                ShowCustomMessageBox("The LLM did not return any content to insert.")

            End If

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If range IsNot Nothing Then Marshal.ReleaseComObject(range) : range = Nothing
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in Command_InsertAfter: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub FreeStyle_InsertAfter()
        Try

            Dim DoMarkup As Boolean = False
            Dim DoInplace As Boolean = False
            Dim DoClipboard As Boolean = False
            Dim NoText As Boolean = False
            Dim MarkupMethod As Integer = INI_MarkupMethodOutlook
            Dim KeepFormatCap = INI_KeepFormatCap ' currently not used
            Dim DoKeepFormat As Boolean = INI_KeepFormat2 ' currently not used
            Dim DoKeepParaFormat As Boolean = INI_KeepParaFormatInline ' currently not used
            Dim DoFileObject As Boolean = False
            Dim FileObject As String = ""

            Dim UseSecondAPI As Boolean = False

            Dim MarkupInstruct As String = $"start with '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"use '{InPlacePrefix}' for replacing the selection"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}' for separate output"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}' for overriding formatting defaults"
            Dim SecondAPIInstruct As String = If(INI_SecondAPI, $"'{SecondAPICode}' to use the secondary model ({INI_Model_2})", "")
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
            Dim ObjectInstruct As String = $"; add '{ObjectTrigger2}' for including a clipboard object"

            Dim AddOnInstruct As String = "; add " & SecondAPIInstruct

            Dim lastCommaIndex As Integer = AddOnInstruct.LastIndexOf(","c)
            If lastCommaIndex <> -1 Then
                AddOnInstruct = AddOnInstruct.Substring(0, lastCommaIndex) & ", and" & AddOnInstruct.Substring(lastCommaIndex + 1)
            End If

            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = outlookApp.ActiveInspector()

            ' Ensure the inspector is open and the item is a MailItem
            If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                SLib.ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

            If wordEditor Is Nothing Then
                SLib.ShowCustomMessageBox("Unable to access the necessary email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If

            ' Get the selected text
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim selectedText As String = selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                NoText = True
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

            ' Prompt for the text to process

            If Not NoText Then
                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {InplaceInstruct}, {ClipboardInstruct}){PromptLibInstruct}{AddOnInstruct}{LastPromptInstruct}:", $"{AN} Freestyle", False, "", My.Settings.LastPrompt)
            Else
                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct}){PromptLibInstruct}{AddOnInstruct}{LastPromptInstruct}:", $"{AN} Freestyle", False, "", My.Settings.LastPrompt)
            End If

            If String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt <> "ESC" AndAlso INI_PromptLib Then

                Dim promptlibresult As (String, Boolean, Boolean, Boolean)

                promptlibresult = ShowPromptSelector(INI_PromptLibPath, Not NoText, Nothing)

                OtherPrompt = promptlibresult.Item1
                DoMarkup = promptlibresult.Item2
                DoClipboard = promptlibresult.Item4

                If OtherPrompt = "" Then
                    Exit Sub
                End If
            Else
                If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Exit Sub
            End If

            My.Settings.LastPrompt = OtherPrompt
            My.Settings.Save()

            ' Check if otherPrompt starts with "Markup:" (case-insensitive)

            If OtherPrompt.StartsWith(ClipboardPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix.Length).Trim()
                DoClipboard = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                DoMarkup = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefixWord, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixWord.Length).Trim()
                DoMarkup = True
                MarkupMethod = 1
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiffW, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiffW.Length).Trim()
                DoMarkup = True
                MarkupMethod = 3
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiff, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiff.Length).Trim()
                DoMarkup = True
                MarkupMethod = 2
            ElseIf OtherPrompt.StartsWith(InPlacePrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(InPlacePrefix.Length).Trim()
                DoMarkup = False
                MarkupMethod = 3
                DoInplace = True
            End If

            ' Formatting Trigger (currently not used)

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
            If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a file object follows)").Trim()
                FileObject = "clipboard"
            End If

            If INI_SecondAPI Then
                If OtherPrompt.Contains(SecondAPICode) Then
                    UseSecondAPI = True
                    OtherPrompt = OtherPrompt.Replace(SecondAPICode, "").Trim()
                End If
            End If

            If DoMarkup And MarkupMethod = 2 And Len(selectedText) > INI_MarkupDiffCap Then
                Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Diff markup method at {INI_MarkupDiffCap} chars (your selection has {Len(selectedText)} chars). How do you want to continue?", "Use Diff in Window compare instead", "Use Diff")
                Select Case MarkupChange
                    Case 1
                        MarkupMethod = 3
                    Case 2
                        MarkupMethod = 2
                    Case Else
                        Exit Sub
                End Select
            End If

            Dim trailingCR As Boolean = selectedText.EndsWith(vbCrLf)

            ' Call your LLM function with the selected text

            Dim LLMResult As String

            If Not NoText Then
                LLMResult = Await LLM(InterpolateAtRuntime(SP_FreestyleText), "<TEXTTOPROCESS>" & selectedText & "</TEXTTOPROCESS>", "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

                LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")
            Else
                LLMResult = Await LLM(InterpolateAtRuntime(SP_FreestyleNoText), "", "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)
            End If

            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            OtherPrompt = ""

            If DoClipboard Then
                Dim FinalText As String = SLib.ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made. If you select Cancel, nothing will be put into the clipboard (without formatting).", AN, True)

                If FinalText <> "" Then
                    SLib.PutInClipboard(FinalText)
                End If
            Else
                ' Collapse the selection to the end

                If Not DoInplace Then
                    selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                Else
                    If Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                    If Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)
                End If

                ' Insert the result as a new paragraph
                If DoMarkup And MarkupMethod <> 3 Then
                    SLib.InsertTextWithMarkdown(selection, vbCrLf & LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                Else
                    If DoInplace Then
                        SLib.InsertTextWithMarkdown(selection, LLMResult)
                    Else
                        SLib.InsertTextWithMarkdown(selection, vbCrLf & LLMResult & vbCrLf)
                    End If
                End If

                ' Use Find to locate the nearest line break backward and adjust selection
                Dim range As Microsoft.Office.Interop.Word.Range = selection.Range
                With range.Find
                    .Text = vbCrLf
                    .Forward = False
                    .MatchWildcards = False
                    If .Execute() Then
                        selection.SetRange(range.Start, selection.End)
                    End If
                End With

                ' Perform markup comparison and insertion if necessary
                If DoMarkup Then
                    If MarkupMethod = 2 Or MarkupMethod = 3 Then
                        CompareAndInsertText(selectedText, LLMResult, MarkupMethod = 3, "This is the markup of the text inserted:", True)
                    Else
                        CompareAndInsertTextCompareDocs(selectedText, LLMResult)
                    End If
                End If
            End If

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle_InsertAfter: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub CompareAndInsertTextCompareDocs(input1 As String, input2 As String)

        Dim splash As New SplashScreen("Creating markup using the Word compare functionality (ignore any flickering and press 'No' if prompted) ...")
        splash.Show()
        splash.Refresh()
        Try
            ' Get the active inspector (compose mail window)
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Inspector = outlookApp.ActiveInspector

            ' Ensure the current item is a MailItem and in compose mode
            If TypeOf inspector.CurrentItem Is MailItem Then
                Dim mailItem As MailItem = CType(inspector.CurrentItem, MailItem)
                Dim editor As Object = inspector.WordEditor

                ' Cast the WordEditor to Word.Document
                Dim wordDoc As Document = CType(editor, Document)

                ' Create a new temporary Word application for comparison
                Dim wordApp As New Microsoft.Office.Interop.Word.Application()
                wordApp.Visible = False

                ' Create temporary documents for input1 and input2
                Dim tempDoc1 As Document = wordApp.Documents.Add()
                Dim tempDoc2 As Document = wordApp.Documents.Add()

                ' Insert the input texts into the temporary documents
                tempDoc1.Content.Text = input1
                tempDoc2.Content.Text = input2

                ' Perform the comparison
                Dim compareResult As Document = wordApp.CompareDocuments(tempDoc1, tempDoc2,
                                                            WdCompareDestination.wdCompareDestinationNew,
                                                            WdGranularity.wdGranularityWordLevel,
                                                            False, False, False, False, False, False)

                ' Convert tracked changes to static formatting
                For Each revision As Revision In compareResult.Revisions
                    Select Case revision.Type
                        Case WdRevisionType.wdRevisionInsert
                            ' Insertions: Apply blue color and underline
                            revision.Range.Font.Color = WdColor.wdColorBlue
                            revision.Range.Font.Underline = WdUnderline.wdUnderlineSingle
                        Case WdRevisionType.wdRevisionDelete
                            ' Deletions: Apply red color and strikethrough
                            revision.Range.Font.Color = WdColor.wdColorRed
                            revision.Range.Font.StrikeThrough = True
                    End Select
                    revision.Accept() ' Accept the revision to make the formatting static
                Next

                ' Copy the comparison result to clipboard
                compareResult.Content.Copy()

                ' Paste the comparison result into the Outlook compose window at the current selection
                wordDoc.Application.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting)

                ' Clean up
                tempDoc1.Close(False)
                tempDoc2.Close(False)
                compareResult.Close(False)
                wordApp.Quit(False)

            Else
                MessageBox.Show("Error in CompareAndInsertTextCompareDocs: The mail compose window is not open (anymore).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            ' Release COM objects in reverse order of creation
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in CompareAndInsertTextCompareDocs: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            splash.Close()

        End Try
    End Sub

    Private Sub CompareAndInsertText(text1 As String, text2 As String, Optional ShowInWindow As Boolean = False, Optional TextforWindow As String = "A text with these changes will be inserted ('Esc' to abort):", Optional DoNotWait As Boolean = False)
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
        'sText = Regex.Replace(sText, "\[DEL_START\](.*?)\s*{vbCrLf}\s*(.*?)\[DEL_END\]", Function(m) $"[DEL_START]{m.Groups(1).Value}{m.Groups(2).Value}[DEL_END] ")

        ' Include instances of line breaks surrounded by [INS_START] and [INS_END] without the [INS...] text
        sText = sText.Replace("[INS_START]{vbCrLf}[INS_END] ", "{vbCrLf}")

        ' Replace placeholders with actual line breaks
        sText = sText.Replace("{vbCrLf}", vbCrLf)

        ' Adjust overlapping tags
        sText = sText.Replace("[DEL_END] [INS_START]", "[DEL_END][INS_START]")
        sText = sText.Replace("[INS_START][INS_END] ", "")

        ' Insert formatted text into the Word editor
        If Not ShowInWindow Then
            InsertFormattedText(sText & vbCrLf)
        Else
            Dim htmlContent As String = ConvertMarkupToRTF(TextforWindow & "\r\r" & sText)
            System.Threading.Tasks.Task.Run(Sub()
                                                ShowRTFCustomMessageBox(htmlContent)
                                            End Sub)
        End If

    End Sub


    Private Function ConvertRtfToPlainText(rtfContent As String) As String
        If String.IsNullOrWhiteSpace(rtfContent) Then Return String.Empty

        ' Remove RTF headers and control words
        Dim plainText As String = Regex.Replace(rtfContent, "{\\.*?}|\\[a-z]+[0-9]*|[{}]", String.Empty)

        ' Decode escaped characters (e.g., \'xx)
        plainText = Regex.Replace(plainText, "\\'([0-9a-fA-F]{2})", Function(m)
                                                                        Dim hex = Convert.ToByte(m.Groups(1).Value, 16)
                                                                        Return Chr(hex)
                                                                    End Function)

        ' Replace RTF line breaks (\par) with actual line breaks
        plainText = Regex.Replace(plainText, "\\par", Environment.NewLine, RegexOptions.IgnoreCase)

        ' Trim the result
        plainText = plainText.Trim()

        Return plainText
    End Function

    Private Sub InsertFormattedText(inputText As String)
        Dim objInspector As Microsoft.Office.Interop.Outlook.Inspector
        Dim objWordDoc As Microsoft.Office.Interop.Word.Document
        Dim objSelection As Object
        Dim objRange As Object
        Dim TextArray() As String = {}
        Dim FormatArray() As Integer = {}
        Dim i As Integer

        ' Store original font properties
        Dim originalFontColor As Integer = 0
        Dim originalUnderline As Integer = 0
        Dim originalStrikeThrough As Boolean = False
        Dim originalBold As Boolean = False
        Dim originalItalic As Boolean = False

        ' Check if there is an active inspector (open email)
        objInspector = TryCast(Globals.ThisAddIn.Application.ActiveInspector, Microsoft.Office.Interop.Outlook.Inspector)
        If objInspector Is Nothing Then
            MessageBox.Show("Error in InsertFormattedText: No open mail item found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' Get the Word editor and the current selection
        objWordDoc = TryCast(objInspector.WordEditor, Microsoft.Office.Interop.Word.Document)
        If objWordDoc Is Nothing Then
            MessageBox.Show("Error in InsertFormattedText: Unable to access the necessary mail editor for this mail.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        objSelection = objWordDoc.Application.Selection

        ' Store original font properties
        If objSelection.Font IsNot Nothing Then
            With objSelection.Font
                originalFontColor = .Color
                originalUnderline = .Underline
                originalStrikeThrough = .StrikeThrough
                originalBold = .Bold
                originalItalic = .Italic
            End With
        End If

        Dim splash As New SplashScreen("Creating your markup ... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        ' Parse the input text into chunks with formatting information
        ParseText(inputText, TextArray, FormatArray)

        ' Reset formatting before starting
        If objSelection.Font IsNot Nothing Then objSelection.Font.Reset()

        ' Insert each text chunk with the appropriate formatting
        For i = 0 To TextArray.Length - 1

            System.Windows.Forms.Application.DoEvents()

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                Exit For
            End If

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                ' Exit the loop
                Exit For
            End If


            ' Reset formatting to original before each insertion
            If objSelection.Font IsNot Nothing Then
                With objSelection.Font
                    .Color = originalFontColor
                    .Underline = originalUnderline
                    .StrikeThrough = originalStrikeThrough
                    .Bold = originalBold
                    .Italic = originalItalic
                End With
            End If

            ' Insert the text at the current cursor position
            objSelection.Collapse(0) ' Collapse to insertion point
            objSelection.TypeText(TextArray(i))

            ' Define the range for the inserted text
            objRange = objSelection.Range
            objRange.Start = objSelection.Start - TextArray(i).Length
            objRange.End = objSelection.Start

            ' Apply formatting based on the tag
            Select Case FormatArray(i)
                Case 1 ' [INS_START]...[INS_END]: Blue underline
                    If objRange.Font IsNot Nothing Then
                        With objRange.Font
                            .Color = RGB(0, 0, 255)
                            .Underline = True
                            .StrikeThrough = False
                        End With
                    End If
                Case 2 ' [DEL_START]...[DEL_END]: Red strikethrough
                    If objRange.Font IsNot Nothing Then
                        With objRange.Font
                            .Color = RGB(255, 0, 0)
                            .StrikeThrough = True
                            .Underline = False
                        End With
                    End If
                Case Else ' Normal text
                    ' Already reset to original formatting
            End Select
        Next

        ' Ensure formatting is reset after all insertions
        If objSelection.Font IsNot Nothing Then
            With objSelection.Font
                .Color = originalFontColor
                .Underline = originalUnderline
                .StrikeThrough = originalStrikeThrough
                .Bold = originalBold
                .Italic = originalItalic
            End With
        End If

        splash.Close()

        ' Release COM objects in reverse order of creation
        If objInspector IsNot Nothing Then Marshal.ReleaseComObject(objInspector) : objInspector = Nothing
        If objWordDoc IsNot Nothing Then Marshal.ReleaseComObject(objWordDoc) : objWordDoc = Nothing

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


    Private Function GetSelectedTextLength() As Integer
        Try
            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = outlookApp.ActiveInspector()

            ' Ensure the inspector is open and the item is a MailItem
            If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                Return 0
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
            DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                Return 0
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document =
            TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

            If wordEditor Is Nothing Then
                Return 0
            End If

            ' Get the selected text
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim selectedText As String = selection.Text

            If String.IsNullOrWhiteSpace(selectedText) Then
                Return 0
            End If

            ' Split on whitespace to count words;
            ' filter out empty entries in case of multiple spaces/newlines
            Dim words = selectedText.Split(New Char() {" "c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf},
                                       StringSplitOptions.RemoveEmptyEntries)
            Return words.Length

            ' Release COM objects in reverse order of creation
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception  ' Explicitly referencing System.Exception per your guideline
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
                        {"DoMarkupOutlook", "Also do a markup (other commands)"},
                        {"MarkupMethodOutlook", "Markup method (1 = Word, 2 = Diff, 3 = DiffW)"},
                        {"MarkupDiffCap", "Maximum characters for Diff Markup"},
                        {"PreCorrection", "Additional instruction for prompts"},
                        {"PostCorrection", "Prompt to apply after queries"},
                        {"Language1", "Default translation language"},
                        {"PromptLibPath", "Prompt library file"}
                    }

        Dim SettingsTips As New Dictionary(Of String, String) From {
                        {"Temperature", "The higher, the more creative the LLM will be (0.0-2.0)"},
                        {"Timeout", "In milliseconds"},
                        {"Temperature_2", "The higher, the more creative the LLM will be (0.0-2.0)"},
                        {"Timeout_2", "In milliseconds"},
                        {"DoubleS", "For Switzerland"},
                        {"KeepFormat1", "If selected, the original's text basic formatting of a translated text will be retained (by HTML encoding, takes time!)"},
                        {"ReplaceText1", "If selected, the response of the LLM for translations will replace the original text"},
                        {"KeepFormat2", "If selected, the original's text basic formatting of other text (other than translations) will be retained (by HTML encoding, takes time!)"},
                        {"ReplaceText2", "If selected, the response of the LLM for other commands (than translate) will replace the original text"},
                        {"DoMarkupOutlook", "Whether a markup should be done for functions that change only parts of a text"},
                        {"MarkupMethodOutlook", "Markup method to use: 1 = Compare using the Word compare function, 2 = Simple Differ, 3 = Simple Diff shown in a window"},
                        {"MarkupDiffCap", "The maximum size of the text that should be processed using the Diff method (to avoid you having to wait too long)"},
                        {"PreCorrection", "Add prompting text that will be added to all basic requests (e.g., for special language tasks)"},
                        {"PostCorrection", "Add a prompt that will be applied to each result before it is further processed (slow!)"},
                        {"Language1", "The language (in English) that will be used for the quick access button in the ribbon"},
                        {"PromptLibPath", "The filename (including path, support environmental variables) for your prompt library (if any)"}
                    }

        ShowSettingsWindow(Settings, SettingsTips)

        Globals.Ribbons.Ribbon1.UpdateRibbon()
        Globals.Ribbons.Ribbon2.UpdateRibbon()

    End Sub



    ' WebExtension integration

    Private httpListener As HttpListener
    Private listenerThread As Thread
    Private isShuttingDown As Boolean = False
    Private listenerTask As System.Threading.Tasks.Task

    Private Sub StartupHttpListener()
        ' fire-and-forget – no raw Thread needed
        listenerTask = StartHttpListener()          ' <— this compiles
    End Sub


    Private Sub ShutdownHttpListener()
        ' Cleanly stop the listener if it's running.
        isShuttingDown = True
        If httpListener IsNot Nothing AndAlso httpListener.IsListening Then
            httpListener.Stop()
            httpListener.Close()
        End If
    End Sub


    Private Async Function StartHttpListener() As System.Threading.Tasks.Task
        Const prefix As String = "http://127.0.0.1:12333/"
        Dim consecutiveFailures As Integer = 0

        While Not isShuttingDown
            Try
                ' ensure listener is alive
                If httpListener Is Nothing Then
                    httpListener = New HttpListener()
                    httpListener.Prefixes.Add(prefix)
                    httpListener.Start()
                    Debug.WriteLine("HttpListener started.")
                ElseIf Not httpListener.IsListening Then
                    httpListener.Close()
                    httpListener = Nothing
                    Continue While   ' next loop will recreate
                End If

                ' wait for a request
                Dim ctx As HttpListenerContext = Await httpListener.GetContextAsync()

                ' fire-and-forget handler
                Call HandleHttpRequest(ctx) _                        ' ignore returned Task
                        .ContinueWith( _                                 ' line-continuation
                            Sub(tResult As System.Threading.Tasks.Task)
                                If tResult.IsFaulted AndAlso
                                   tResult.Exception IsNot Nothing Then

                                    Debug.WriteLine("HandleHttpRequest error: " &
                                                    tResult.Exception.GetBaseException().Message)
                                End If
                            End Sub, _                                   ' ← underscore
                            System.Threading.Tasks.TaskScheduler.Default)

                consecutiveFailures = 0                        ' success
            Catch ex As ObjectDisposedException
                consecutiveFailures += 1
            Catch ex As System.Exception
                consecutiveFailures += 1
                Debug.WriteLine("Listener error: " & ex.Message)
            End Try

            ' if we hit 10 failures in a row, recycle the listener
            If consecutiveFailures >= 10 AndAlso Not isShuttingDown Then
                Debug.WriteLine("Restarting HttpListener after 10 failures.")
                Try
                    If httpListener IsNot Nothing Then httpListener.Close()
                Catch
                End Try
                httpListener = Nothing
                consecutiveFailures = 0
                Await System.Threading.Tasks.Task.Delay(5000)  ' back-off pause
            End If
        End While
    End Function
    ' ---------------------------------------------------------------------------

    Private Async Function HandleHttpRequest(
            ctx As System.Net.HttpListenerContext) As System.Threading.Tasks.Task

        Dim req = ctx.Request
        Dim res = ctx.Response

        If req.HttpMethod = "OPTIONS" Then
            res.AddHeader("Access-Control-Allow-Origin", "*")
            res.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
            res.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization")
            res.StatusCode = 204 : res.Close() : Return
        End If

        Dim body As String = ""
        If req.HasEntityBody Then
            Using rdr As New IO.StreamReader(req.InputStream, System.Text.Encoding.UTF8)
                body = Await rdr.ReadToEndAsync().ConfigureAwait(False)
            End Using
        End If

        Dim responseText As String =
            Await ProcessRequestInAddIn(body, req.RawUrl).ConfigureAwait(False)

        Dim buf = System.Text.Encoding.UTF8.GetBytes(responseText)
        res.ContentLength64 = buf.Length
        res.ContentType = "text/plain; charset=utf-8"
        res.AddHeader("Access-Control-Allow-Origin", "*")
        Using os = res.OutputStream
            Await os.WriteAsync(buf, 0, buf.Length).ConfigureAwait(False)
        End Using
        res.Close()
    End Function
    ' ---------------------------------------------------------------------------

    ' --------------- LLM helper (runs off the UI thread) -----------------------


    ' ----------------------------------------
    ' 1) Feld für den Scheduler (Klassen-/Modul-Ebene)
    ' ----------------------------------------
    Private Shared llmScheduler As System.Threading.Tasks.TaskScheduler

    ' ----------------------------------------
    ' 2) STA-Thread mit eigener WinForms-Message-Loop initialisieren
    ' ----------------------------------------
    Private Sub EnsureLlmUiThread()
        If llmScheduler IsNot Nothing Then
            Return
        End If

        ' TaskCompletionSource liefert uns den Scheduler, sobald der STA-Thread
        ' seinen SynchronizationContext gesetzt hat.
        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of System.Threading.Tasks.TaskScheduler)()

        ' Thread-Start
        Dim th As New System.Threading.Thread(Sub()
                                                  ' 1) SyncContext für WinForms in diesem Thread setzen
                                                  System.Threading.SynchronizationContext.SetSynchronizationContext(
                                                  New System.Windows.Forms.WindowsFormsSynchronizationContext())

                                                  ' 2) Scheduler aus dem aktuellen Context erzeugen
                                                  tcs.SetResult(System.Threading.Tasks.TaskScheduler.FromCurrentSynchronizationContext())

                                                  ' 3) Message-Loop starten (Application.Run pumpt Meldungen)
                                                  System.Windows.Forms.Application.Run()
                                              End Sub)

        th.SetApartmentState(System.Threading.ApartmentState.STA)
        th.IsBackground = True
        th.Start()

        ' blockierend warten, bis wir den Scheduler erhalten haben
        llmScheduler = tcs.Task.Result
    End Sub

    ' ----------------------------------------
    ' 3) Neues RunLlmAsync (Drop-in für Deine alte Methode)
    ' ----------------------------------------

    ''' <summary>
    ''' Führt Deinen LLM-Call (mit HTTP + UI-Dialogs) komplett
    ''' auf einem eigenen STA-Thread mit Message-Loop aus.
    ''' </summary>
    Public Function RunLlmAsync(
    sysPrompt As String,
    userPrompt As String
) As Task(Of String)

        ' Stelle sicher, dass unser STA-Thread und Scheduler bereit sind
        EnsureLlmUiThread()

        ' Wir packen alles in einen LongRunning-Task, der
        ' auf genau diesem STA-Thread ausgeführt wird:
        Return System.Threading.Tasks.Task.Factory.StartNew(Of String)(
        Function() As String
            ' AsyncContext.Run pumpt WinForms- & COM-Messages
            ' solange bis Dein LLM-Task wirklich fertig ist.
            Return AsyncContext.Run(
                Async Function() As Task(Of String)
                    ' Hier kommt Dein bisheriger LLM-Aufruf hin:
                    ' er darf HTTP machen und beliebige WinForms-Dialogs öffnen.
                    Return Await LLM(sysPrompt, userPrompt, "", "", 0)
                End Function)
        End Function,
        CancellationToken.None,
        TaskCreationOptions.LongRunning,
        llmScheduler)
    End Function



    ' ---------------------------------------------------------------------------

    ' --------------- Compare & insert helper (runs on UI) ----------------------
    '───────────────────────────────────────────────────────────────────────────
    ' Shows the compare window on the UI thread and BLOCKS the calling
    ' Task until the user dismisses it.  Returns True if the user accepted,
    ' False if Esc was pressed (like the original code).
    '───────────────────────────────────────────────────────────────────────────
    Private Function CompareAndInsertSyncConfirm(
        originalText As String,
        llmResult As String) _
        As System.Threading.Tasks.Task(Of Boolean)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Boolean)

        ' marshal to UI thread with BeginInvoke so listener thread never blocks
        mainThreadControl.BeginInvoke(New MethodInvoker(
        Sub()

            ' 1) show compare window (modal for this thread)
            CompareAndInsertText(originalText, llmResult, True)

            ' 2) pump one message cycle so the Esc keystroke is processed
            System.Windows.Forms.Application.DoEvents()

            ' 3) read Esc status exactly like the old code
            Dim escNow As Boolean =
                (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0
            Dim escDown As Boolean =
                (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0

            Dim accepted As Boolean = Not (escNow Or escDown)

            tcs.SetResult(accepted)      ' unblock the awaiting thread
        End Sub))

        Return tcs.Task       ' caller awaits without blocking the listener
    End Function
    ' ---------------------------------------------------------------------------


    '───────────────────────────────────────────────────────────────────────────
    ' Waits asynchronously until the preview window (ShowRTFCustomMessageBox)
    ' is either closed with OK  → returns True
    '                     or Esc is pressed       → returns False
    ' Works even though the preview window is created on its own worker thread.
    '───────────────────────────────────────────────────────────────────────────
    Private Function WaitForPreviewDecisionAsync() _
        As System.Threading.Tasks.Task(Of Boolean)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Boolean)()

        ' Run once on the UI thread: attach handlers when the form appears
        SwitchToUi(Sub()

                       ' 1) keep checking until the preview form exists
                       Dim previewForm As System.Windows.Forms.Form = Nothing
                       Dim searchTimer As New System.Windows.Forms.Timer() With {.Interval = 100}

                       AddHandler searchTimer.Tick,
            Sub()

                If previewForm Is Nothing OrElse previewForm.IsDisposed Then
                    previewForm = System.Windows.Forms.Application.OpenForms _
                                      .Cast(Of System.Windows.Forms.Form)() _
                                      .FirstOrDefault(Function(f) f.Name = "ShowRTFCustomMessageBox" _
                                                          OrElse f.Text.StartsWith(AN))

                    If previewForm Is Nothing Then Return  ' continue searching
                    ' found the form ⇒ attach handlers ------------------------
                    previewForm.KeyPreview = True

                    AddHandler previewForm.KeyDown,
                        Sub(_s, e)
                            If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                tcs.TrySetResult(False)     ' user cancelled
                            End If
                        End Sub

                    AddHandler previewForm.FormClosed,
                        Sub(_s, _e)
                            tcs.TrySetResult(True)          ' user accepted (OK)
                        End Sub
                End If

                ' stop searching once the Task is completed
                If tcs.Task.IsCompleted Then searchTimer.Stop()
            End Sub

                       searchTimer.Start()
                   End Sub).Wait()   ' Wait only for handler-attachment setup

        Return tcs.Task    ' listener awaits without polling or blocking UI
    End Function




    ' ---------------- MAIN REQUEST DISPATCH ------------------------------------
    Private Async Function ProcessRequestInAddIn(
            body As String,
            rawUrl As String) As System.Threading.Tasks.Task(Of String)

        Dim j = Newtonsoft.Json.Linq.JObject.Parse(body)
        Dim cmd = j("Command")?.ToString()
        Dim textBody = j("Text")?.ToString()
        Dim sourceUrl = j("URL")?.ToString()

        Select Case cmd
        ' -------------------------------------------------------------------
            Case "redink_sendtooutlook"
                If String.IsNullOrWhiteSpace(textBody) Then Return ""
                ' All Outlook automation on UI thread
                Await SwitchToUi(Sub()
                                     Dim olApp = Globals.ThisAddIn.Application
                                     Dim insp = olApp.ActiveInspector()
                                     If insp Is Nothing Then Exit Sub
                                     If Not TypeOf insp.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then Exit Sub
                                     Dim mail = CType(insp.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)
                                     If mail.Sent Then Exit Sub
                                     Dim doc = CType(insp.WordEditor, Microsoft.Office.Interop.Word.Document)
                                     Dim rng = doc.Application.Selection.Range
                                     doc.Application.ScreenUpdating = False
                                     rng.Text = textBody & " (" & sourceUrl & ")"
                                     doc.Application.ScreenUpdating = True
                                     ' release COM objects
                                     System.Runtime.InteropServices.Marshal.ReleaseComObject(rng)
                                     System.Runtime.InteropServices.Marshal.ReleaseComObject(doc)
                                 End Sub)
                Return ""
        ' -------------------------------------------------------------------
            Case "redink_translate"
                ' ─── 1  guard clauses ─────────────────────────────────────────
                If String.IsNullOrWhiteSpace(textBody) Then Return ""

                ' Ask the user for a target language (UI thread)
                Dim targetLang As String = Await SwitchToUi(Function()
                                                                Return SLib.ShowCustomInputBox(
                   "Enter your target language:",
                   AN & " Translate (for Browser)",
                   True, INI_Language1)
                                                            End Function)

                If String.IsNullOrWhiteSpace(targetLang) OrElse targetLang = "ESC" Then
                    Return ""                                   ' user cancelled
                End If

                ' ─── 2  call the LLM on the UI thread, get Task(Of String) ─────
                Dim llmOut As String = Await RunLlmAsync(
                    InterpolateAtRuntime(SP_Translate),
                    $"<TEXTTOPROCESS>{textBody}</TEXTTOPROCESS>")

                ' ─── 3  clean up the wrapper tags / markdown ──────────────────
                llmOut = llmOut.Replace("<TEXTTOPROCESS>", "") _
                   .Replace("</TEXTTOPROCESS>", "") _
                   .Replace("**", "").Trim()

                If llmOut = "" Then Return ""                  ' safety net

                ' Optional: copy to clipboard so the user can paste manually
                Await SwitchToUi(Sub() SLib.PutInClipboard(llmOut))

                ' ─── 4  SEND the translation back to the caller ───────────────
                Return llmOut

            ' -------------------------------------------------------------------
            Case "redink_correct"

                If String.IsNullOrWhiteSpace(textBody) Then Return ""

                ' 1)  Run the LLM on the UI thread
                Dim llmOut As String = Await RunLlmAsync(
                                                    InterpolateAtRuntime(SP_Correct),
                                                    $"<TEXTTOPROCESS>{textBody}</TEXTTOPROCESS>")
                llmOut = llmOut.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

                If llmOut = "" Then Return ""

                ' 2)  Show the compare / preview window (synchronous)
                Await SwitchToUi(Sub()
                                     CompareAndInsertText(textBody, llmOut, True)
                                 End Sub)

                ' 3)  
                Dim accepted As Boolean = Await WaitForPreviewDecisionAsync()

                If Not accepted Then Return ""          ' Esc pressed → abort

                Return llmOut

        ' -------------------------------------------------------------------
            Case "redink_freestyle"

                '─── A  gather prompt on UI thread ──────────────────────────────
                Dim noText As Boolean = String.IsNullOrWhiteSpace(textBody)

                Dim promptCaption As String = AN & " Freestyle (for Browser)"

                Dim sb As New System.Text.StringBuilder()
                If noText Then
                    sb.Append("Please provide the prompt you wish to execute ")
                Else
                    sb.Append("Please provide the prompt you wish to execute using the selected text ")
                End If

                sb.Append("(" & MarkupPrefix & " for markups, " & InsertPrefix & " for direct insert)")
                If INI_PromptLib Then sb.Append(" or press 'OK' for the prompt library")
                If Not String.IsNullOrWhiteSpace(My.Settings.LastPrompt) Then sb.Append("; ctrl-p for your last prompt")
                sb.Append(":")
                Dim promptMsg As String = sb.ToString()

                OtherPrompt = Await SwitchToUi(Function()
                                                   Return SLib.ShowCustomInputBox(promptMsg, promptCaption, False, "", My.Settings.LastPrompt)
                                               End Function)

                Dim doMarkupFlag As Boolean = False
                Dim doInsertFlag As Boolean = False

                '─── prompt library branch ─────────────────────────────────────
                If String.IsNullOrEmpty(otherPrompt) AndAlso otherPrompt <> "ESC" AndAlso INI_PromptLib Then
                    Dim sel = Await SwitchToUi(Function()
                                                   Return ShowPromptSelector(INI_PromptLibPath, Not noText, Nothing)
                                               End Function)                         ' (prompt, doMarkup, doInsert, canceled)

                    otherPrompt = sel.Item1
                    doMarkupFlag = sel.Item2
                    doInsertFlag = Not sel.Item4         ' library’s “canceled” → insert = False
                End If

                ' user cancelled
                If String.IsNullOrWhiteSpace(otherPrompt) OrElse otherPrompt = "ESC" Then
                    Return ""
                End If

                ' remember last prompt
                My.Settings.LastPrompt = otherPrompt
                My.Settings.Save()

                '─── decode prefix flags ───────────────────────────────────────
                If otherPrompt.StartsWith(InsertPrefix, StringComparison.OrdinalIgnoreCase) Then
                    otherPrompt = otherPrompt.Substring(InsertPrefix.Length).Trim()
                    doInsertFlag = True
                ElseIf otherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) AndAlso Not noText Then
                    otherPrompt = otherPrompt.Substring(MarkupPrefix.Length).Trim()
                    doMarkupFlag = True
                    doInsertFlag = True          ' old logic: markup implies insert
                End If

                '─── B  call the LLM on UI thread (async) ──────────────────────
                Dim llmResult As String
                If noText Then
                    llmResult = Await RunLlmAsync(
            InterpolateAtRuntime(SP_FreestyleNoText), "")
                Else
                    llmResult = Await RunLlmAsync(
            InterpolateAtRuntime(SP_FreestyleText),
            $"<TEXTTOPROCESS>{textBody}</TEXTTOPROCESS>")
                End If

                llmResult = llmResult.Replace("<TEXTTOPROCESS>", "") _
                         .Replace("</TEXTTOPROCESS>", "") _
                         .Trim()

                If String.IsNullOrEmpty(llmResult) Then Return ""

                '─── C  present / insert / clipboard exactly like old code ─────

                ' A) markup path (implies insert)  -----------------------------
                If doMarkupFlag Then
                    Await SwitchToUi(Sub()
                                         CompareAndInsertText(textBody, llmResult, True)
                                     End Sub)

                    Dim accepted As Boolean = Await WaitForPreviewDecisionAsync()

                    If Not accepted Then Return ""          ' Esc pressed → abort

                    Return llmResult                    ' user accepted
                End If

                ' B) plain insert path  ----------------------------------------
                If doInsertFlag Then
                    'Await InsertTextIntoCurrentMailAsync(llmResult)
                    Return llmResult                        ' send text back
                End If

                ' C) clipboard-only path  --------------------------------------
                Dim finalTxt As String = Await SwitchToUi(Function()
                                                              Return SLib.ShowCustomWindow(
                                                                  "The LLM has provided the following result (you can edit it):",
                                                                  llmResult,
                                                                  "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made. If you select Cancel, nothing will be put into the clipboard (without formatting).",
                                                                  AN, True, True)
                                                          End Function)

                If Not String.IsNullOrWhiteSpace(finalTxt) Then
                    Await SwitchToUi(Sub() SLib.PutInClipboard(finalTxt))
                End If

                Return ""                                   ' old behaviour: no body sent

        End Select
    End Function



End Class
