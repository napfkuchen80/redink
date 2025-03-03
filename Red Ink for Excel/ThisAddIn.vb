' Red Ink for Excel
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See License.txt or https://vischer.com/redink for more information.
'
' 3.3.2025
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
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.

Option Explicit On

Imports Microsoft.Office.Core
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.IO
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Tools


Module Module1
    ' Correct attribute declaration for DllImport
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function

End Module


#Region "BridgeSubs"

<ComVisible(True)>
Public Class BridgeSubs
    Public Async Function DoInLanguage1() As Task(Of Boolean)
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function DoInLanguage2() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage2()
    End Function

    Public Async Function DoInOtherFormulas() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.InOtherFormulas()
    End Function

    Public Async Function DoCorrect() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Correct()
    End Function

    Public Async Function DoImprove() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Improve()
    End Function

    Public Async Function DoShorten() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Shorten()
    End Function

    Public Async Function DoAnonymize() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Anonymize()
    End Function

    Public Async Function DoSwitchParty() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.SwitchParty()
    End Function

    Public Async Function DoFreestyleNM() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function DoFreestyleAM() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleAM()
    End Function

    Public Sub DoAdjustHeight(Optional Silent As Boolean = False)
        Globals.ThisAddIn.AdjustHeight(Silent)
    End Sub
    Public Sub DoRegexSearchReplace()
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Public Sub DoAdjustLegacyNotes()
        Globals.ThisAddIn.AdjustLegacyNotes()
    End Sub

    Public Sub DoAddContextMenu()
        Globals.ThisAddIn.AddContextMenu()
    End Sub

    Public Function GetLLMConfig(UseSecondAPI As Boolean) As String
        Dim Result As String = Globals.ThisAddIn.GetAPIConfiguration(UseSecondAPI)
        Return Result
    End Function

    Public Function SignJWT(jwtUnsigned As String, privateKeyPem As String) As String
        Return SLib.SignJWT(jwtUnsigned, privateKeyPem)
    End Function

    Public Function GetFileTextContent(ByVal filePath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
        Try
            ' Normalize and check the path
            filePath = Path.GetFullPath(filePath)
            If Not File.Exists(filePath) Then
                Return If(ReturnErrorInsteadOfEmpty, "Error: File not found", "")
            End If

            ' Determine file type by extension
            Dim extension As String = Path.GetExtension(filePath).ToLower()

            Select Case extension
                Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm"
                    Return ReadTextFile(filePath, ReturnErrorInsteadOfEmpty)

                Case ".rtf"
                    Return ReadRtfAsText(filePath, ReturnErrorInsteadOfEmpty)

                Case ".doc", ".docx"
                    Return ReadWordDocument(filePath, ReturnErrorInsteadOfEmpty)

                Case ".pdf"
                    Return ReadPdfAsText(filePath, ReturnErrorInsteadOfEmpty)

                Case Else
                    Return If(ReturnErrorInsteadOfEmpty, "Error: File type not supported (not txt, rtf, doc, docx, pdf, ini, csv, log, json, xml, html or htm)", "")
            End Select
        Catch ex As UnauthorizedAccessException
            Return If(ReturnErrorInsteadOfEmpty, "Error: Unauthorized access", "")
        Catch ex As IOException
            Return If(ReturnErrorInsteadOfEmpty, "Error: IO Error: " & ex.Message, "")
        Catch ex As System.Exception
            Return If(ReturnErrorInsteadOfEmpty, "Error: Unexpected error: " & ex.Message, "")
        End Try
    End Function



End Class

#End Region

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        InitializeAddInFeatures()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        RemoveOldContextMenu()
    End Sub

    ' Hardcoded config values

    Public Const Version As String = "V.030325 Gen2 Beta Test"

    Public Const AN As String = "Red Ink"
    Public Const AN2 As String = "redink"

    Private Const ShortenPercent As Integer = 20
    Private Const TextPrefix As String = "TextOnly:"
    Private Const TextPrefix2 As String = "Text:"
    Private Const CellByCellPrefix As String = "CellByCell:"
    Private Const CellByCellPrefix2 As String = "CBC:"
    Private Const RIMenu = AN
    Private Const MinHelperVersion = 1           ' Minimum version of the helper file that is required

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

    Public Shared ReadOnly Property RDV As String = "Excel (" & Version & ")"
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
        _context.RDV = "Excel (" & Version & ")"
        SharedMethods.InitializeConfig(_context, FirstTime, Reload)
    End Sub
    Private Function INIValuesMissing()
        Return SharedMethods.INIValuesMissing(_context)
    End Function
    Public Shared Async Function PostCorrection(inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
        Return Await SharedMethods.PostCorrection(_context, inputText, UseSecondAPI)
    End Function
    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal Hidesplash As Boolean = True, Optional ByVal AddUserPrompt As String = "") As Task(Of String)
        Return Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, Hidesplash, AddUserPrompt)
    End Function
    Private Function ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String))
        SharedMethods.ShowSettingsWindow(Settings, SettingsTips, _context)
    End Function
    Private Function ShowPromptSelector(filePath As String, enableMarkup As Boolean, enableBubbles As Boolean) As (String, Boolean, Boolean, Boolean)
        Return SharedMethods.ShowPromptSelector(filePath, enableMarkup, enableBubbles, _context)
    End Function

#End Region

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

        If RemoveMenu Then
            RemoveOldContextMenu()
            RemoveMenu = False
        End If

        If Not INI_ContextMenu Then Exit Sub

        If Not VBAModuleWorking() Then Exit Sub

        If INIloaded = False Then Exit Sub

        MenusAdded = True

        ' List of relevant context menus
        Dim contextMenus As String() = {"Cell", "Row", "Column", "List Range Popup", "PivotTable Context Menu", "Text Box", "Drawing Object", "Chart"}
        Dim application As Excel.Application = Globals.ThisAddIn.Application

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
            Dim excelHelpersMenu As CommandBarPopup
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
                    subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption) ' Access the value
                End If
            End If
            If Not String.IsNullOrWhiteSpace(INI_Language2) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language2
                subControl.OnAction = "CallInLanguage2"
                subControl.FaceId = 6112

                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
                End If
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "To Other (text)"
            subControl.OnAction = "CallInOther"
            subControl.FaceId = 6112
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "To Other (cells)"
            subControl.OnAction = "CallInOtherFormulas"
            subControl.FaceId = 6112
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Correct"
            subControl.OnAction = "CallCorrect"
            subControl.FaceId = 329

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Write Neatly"
            subControl.OnAction = "CallNeatly"
            subControl.FaceId = 162

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Shorten"
            subControl.OnAction = "CallShorten"
            subControl.FaceId = 292
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Anonymize"
            subControl.OnAction = "CallAnonymize"
            subControl.FaceId = 7502
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Switch Party"
            subControl.OnAction = "CallSwitchParty"
            subControl.FaceId = 327
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Freestyle"
            subControl.OnAction = "CallFreestyleNM"
            subControl.FaceId = 346
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            If INI_SecondAPI Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "Freestyle (" & INI_Model_2 & ")"
                subControl.OnAction = "CallFreestyleAM"
                subControl.FaceId = 346
                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
                End If

            End If

            ' Create new submenu "Excel helpers"
            excelHelpersMenu = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
            excelHelpersMenu.Caption = "Excel Helpers"

            subSubControl = CType(excelHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Adjust Cell Height"
            subSubControl.OnAction = "CallAdjustHeight"
            subSubControl.FaceId = 1647

            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut: " & shortcutDict(subSubControl.Caption)
            End If


            subSubControl = CType(excelHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Adjust Size of Notes"
            subSubControl.OnAction = "CallAdjustLegacyNotes"
            subSubControl.FaceId = 1996

            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut: " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(excelHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Regex Search && Replace"
            subSubControl.OnAction = "CallRegexSearchReplace"
            subSubControl.FaceId = 288
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut: " & shortcutDict(subSubControl.Caption)
            End If

            If Not String.IsNullOrWhiteSpace(INI_ShortcutsWordExcel) Then
                ' Assign shortcuts using the dictionary
                If Not String.IsNullOrWhiteSpace(INI_Language1) Then AssignShortcut("To " & INI_Language1, "CallInLanguage1", shortcutDict)
                If Not String.IsNullOrWhiteSpace(INI_Language2) Then AssignShortcut("To " & INI_Language2, "CallInLanguage2", shortcutDict)
                AssignShortcut("To Other (text)", "CallInOther", shortcutDict)
                AssignShortcut("To Other (cells)", "CallInOther", shortcutDict)
                AssignShortcut("Correct", "CallCorrect", shortcutDict)
                AssignShortcut("Write Neatly", "CallImprove", shortcutDict)
                AssignShortcut("Shorten", "CallShorten", shortcutDict)
                AssignShortcut("Anonymize", "CallAnonymize", shortcutDict)
                AssignShortcut("Switch Party", "CallSwitchParty", shortcutDict)
                AssignShortcut("Freestyle", "CallFreestyleNM", shortcutDict)

                ' Assign shortcuts for second API if applicable
                If INI_SecondAPI Then
                    AssignShortcut("Freestyle (" & INI_Model_2 & ")", "CallFreestyleAM", shortcutDict)
                End If

                ' Assign shortcuts for submenu "Excel helpers"
                AssignShortcut("Adjust Cell Height", "CallAdjustheight", shortcutDict)
                AssignShortcut("Adjust Size of Notes", "CallAdjustLegacyNotes", shortcutDict)
                AssignShortcut("Regex Search & Replace", "CallRegexSearchReplace", shortcutDict)
                AssignShortcut("Regex Search && Replace", "CallRegexSearchReplace", shortcutDict)
            End If
        Catch ex As System.Exception

        End Try
    End Sub
    Public Sub AssignShortcut(ByVal controlName As String, ByVal macro As String, ByRef shortcutDict As Dictionary(Of String, String))
        Dim shortcutKey As String
        Dim keyCombination As String
        Try
            ' Check if there is a shortcut assigned for this control
            If shortcutDict.ContainsKey(controlName) Then
                shortcutKey = shortcutDict(controlName)
            Else
                Exit Sub ' No shortcut assigned
            End If

            ' Build the key combination string from the shortcutKey text
            keyCombination = BuildKeyCodeFromText(shortcutKey)

            If Not String.IsNullOrEmpty(keyCombination) Then
                ' Assign the shortcut key to the macro in Excel using Application.OnKey
                Globals.ThisAddIn.Application.OnKey(keyCombination, macro)
            End If
        Catch ex As System.Exception
            ' Handle exceptions gracefully
            ' Debug.WriteLine("Error in AssignShortcut: " & ex.Message)
        End Try
    End Sub

    Public Function BuildKeyCodeFromText(ByVal shortcutKey As String) As String
        Dim parts() As String
        Dim keysCollection As New List(Of String)()
        Dim keyCombination As String = ""

        Try
            parts = shortcutKey.Split("-"c)

            For Each part As String In parts
                Select Case part.Trim().ToUpper()
                    Case "CTRL"
                        keysCollection.Add("^") ' Control key representation in Excel
                    Case "SHIFT"
                        keysCollection.Add("+") ' Shift key representation in Excel
                    Case "ALT"
                        keysCollection.Add("%") ' Alt key representation in Excel

                ' Map digits directly
                    Case "0" To "9"
                        keysCollection.Add(part.Trim())

                ' Map function keys directly
                    Case "F1" To "F12"
                        keysCollection.Add(part.Trim())

                ' Letters mapped directly
                    Case "A" To "Z"
                        keysCollection.Add(part.Trim().ToUpper())

                    Case Else
                        ' Unknown key
                        Return ""
                End Select
            Next

            ' Combine the keys into a single shortcut string for VBA
            keyCombination = String.Join("", keysCollection)

            Return keyCombination

        Catch ex As System.Exception
            ' Handle errors gracefully
            ' Debug.WriteLine("Error in BuildKeyCodeFromText: " & ex.Message)
            Return ""
        End Try
    End Function


    Public Sub RemoveOldContextMenu()
        Dim application As Excel.Application = Globals.ThisAddIn.Application

        ' Array of relevant context menus
        Dim contextMenus As String() = {"Cell", "Row", "Column", "List Range Popup", "PivotTable Context Menu", "Text Box", "Drawing Object", "Chart"}

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
                                'Debug.WriteLine($"Error removing control: {ex.Message}")
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
    Public Context As String
    Public SysPrompt As String
    Public OldParty, NewParty As String
    Public SelectedText As String

    Public Structure CellState
        Public CellAddress As String
        Public OldValue As Object
        Public HadFormula As Boolean
        Public OldFormula As String
    End Structure

    Public Shared undoStates As New List(Of CellState)

    Public Async Function InLanguage1() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = INI_Language1
        Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, True, False)
    End Function
    Public Async Function InLanguage2() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = INI_Language2
        Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, True, False)
    End Function
    Public Async Function InOther() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            If selectedRange IsNot Nothing Then
                selectedRange.Select()
            End If

            Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, True, False)
        End If
    End Function
    Public Async Function InOtherFormulas() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, True, True, False)
        End If
    End Function
    Public Async Function Correct() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await ProcessSelectedRange(SP_Correct, True, False, False, True, False)
    End Function
    Public Async Function Improve() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If

        Context = Trim(SLib.ShowCustomInputBox("Please provide the context that should be taken into account, if any:", $"{AN} Write Neatly", True))

        If String.IsNullOrWhiteSpace(Context) Then
            Context = "n/a"
        End If

        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If

        Dim result As Boolean = Await ProcessSelectedRange(SP_WriteNeatly, True, False, False, True, False)

    End Function
    Public Async Function Anonymize() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await ProcessSelectedRange(SP_Anonymize, True, False, False, True, False)
    End Function
    Public Async Function Shorten() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)

        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If

        Dim totalLength As Integer = 0
        Dim maxLength As Integer = 0
        Dim cellCount As Integer = 0

        For Each cell As Excel.Range In selectedRange.Cells
            If Not CellProtected(cell) AndAlso Not cell.HasFormula Then
                Dim cellText As String = CStr(cell.Value)
                If Not String.IsNullOrEmpty(cellText) Then
                    Dim textLength As Integer = cellText.Length
                    totalLength += textLength
                    If textLength > maxLength Then
                        maxLength = textLength
                    End If
                    cellCount += 1
                End If
            End If
        Next

        Dim averageLength As Double = If(cellCount > 0, totalLength / cellCount, 0)

        Dim UserInput As String
        Dim ShortenPercentValue As Integer = 0
        Do
            UserInput = Trim(SLib.ShowCustomInputBox($"Enter the percentage by which the text of each selected cell should be shortened (the cells have have of average {averageLength:n1} words and {maxLength} at max; {ShortenPercent}% will cut approx. " & (averageLength * ShortenPercent / 100) & " words in average):", $"{AN} Shortener", True, CStr(ShortenPercent) & "%"))
            If String.IsNullOrEmpty(UserInput) Then
                Return False
            End If
            UserInput = UserInput.Replace("%", "").Trim()
            If Integer.TryParse(UserInput, ShortenPercentValue) AndAlso ShortenPercentValue >= 1 AndAlso ShortenPercentValue <= 99 Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid percentage between 1 and 99.")
            End If
        Loop
        If ShortenPercentValue = 0 Then Return False
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If


        Dim result As Boolean = Await ProcessSelectedRange(SP_Shorten, True, False, False, True, False, ShortenPercentValue)
    End Function
    Public Async Function SwitchParty() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)

        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If

        Dim UserInput As String
        Do
            UserInput = Trim(SLib.ShowCustomInputBox("Please provide the original party name and the new party name, separated by a comma (example: Elvis Presley, Taylor Swift):", $"{AN} Switch Party", True))

            If String.IsNullOrEmpty(UserInput) Then
                Return False
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
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If

        Dim result As Boolean = Await ProcessSelectedRange(SP_SwitchParty, True, False, False, True, False)

    End Function

    Public Async Function FreestyleNM() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await Freestyle(False)
    End Function

    Public Async Function FreestyleAM() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await Freestyle(True)
    End Function

    Public Async Function Freestyle(ByVal UseSecondAPI As Boolean) As Task(Of Boolean)

        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        Dim NoSelectedCells As Boolean = False
        Dim DoClipboard As Boolean = False
        Dim DoFormulas As Boolean = True

        Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; ctrl-v for your last prompt")

        If selectedRange Is Nothing Then
            NoSelectedCells = True
        End If

        Dim DoRange As Boolean = True
        Dim CBCInstruct As String = $"with '{CellByCellPrefix}' or '{CellByCellPrefix2} if the instruction should be executed cell-by-cell"
        Dim TextInstruct As String = $"use '{TextPrefix}' or '{TextPrefix2}' if the instruction should apply cell-by-cell, but only to text cells"

        Dim PromptLibInstruct As String = ""
        If INI_PromptLib Then
            PromptLibInstruct = " or press 'OK' for the prompt library"
        End If

        SLib.StoreClipboard()

        If Not String.IsNullOrWhiteSpace(My.Settings.LastPrompt) Then SLib.PutInClipboard(My.Settings.LastPrompt)

        If Not NoSelectedCells Then
            OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected cells (start {CBCInstruct}; {TextInstruct})" & PromptLibInstruct & LastPromptInstruct & ":", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False))
        Else
            OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute {PromptLibInstruct} (the result will be shown to you before inserting anything into your worksheet){LastPromptInstruct}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False))
            DoRange = True
        End If

        SLib.RestoreClipboard()

        If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then

            Dim promptlibresult As (String, Boolean, Boolean, Boolean)

            promptlibresult = ShowPromptSelector(INI_PromptLibPath, Nothing, Nothing)

            OtherPrompt = promptlibresult.Item1
            DoClipboard = promptlibresult.Item4

            If OtherPrompt = "" Then
                Return False
            End If
        Else
            If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return False
        End If

        My.Settings.LastPrompt = OtherPrompt
        My.Settings.Save()

        If OtherPrompt.StartsWith(CellByCellPrefix, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(CellByCellPrefix.Length).Trim()
            DoRange = False
        End If
        If OtherPrompt.StartsWith(CellByCellPrefix2, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(CellByCellPrefix2.Length).Trim()
            DoRange = False
        End If
        If OtherPrompt.StartsWith(TextPrefix, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(TextPrefix.Length).Trim()
            DoRange = False
            DoFormulas = False
        End If
        If OtherPrompt.StartsWith(TextPrefix2, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(TextPrefix2.Length).Trim()
            DoRange = False
            DoFormulas = False
        End If
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If

        If Not NoSelectedCells Then
            If DoRange Then
                Dim result As Boolean = Await ProcessSelectedRange(SP_RangeOfCells, True, DoRange, DoFormulas, False, UseSecondAPI, 0, True)
            Else
                Dim result As Boolean = Await ProcessSelectedRange(SP_FreestyleText, True, DoRange, DoFormulas, False, UseSecondAPI)
            End If
        Else
            Dim result As Boolean = Await ProcessSelectedRange(SP_RangeOfCells, True, DoRange, DoFormulas, False, UseSecondAPI, 0, True)
        End If

    End Function


    ' ProcessSelectedRang
    '
    ' This function processes the selected range of cells in Excel. It takes the following parameters:
    ' - SysCommand: The system command to be executed
    ' - CheckMaxToken: A boolean value indicating whether the maximum token count should be checked
    ' - DoRange: A boolean value indicating whether the selected range should be processed
    ' - DoFormulas: A boolean value indicating whether formulas should be processed
    ' - SelectionMandatory: A boolean value indicating whether a selection is mandatory
    ' - UseSecondAPI: A boolean value indicating whether the second API should be used
    ' - Optional: ShortenPercentValue: A percentage value by which the text should be shortened (for calculating the word count for each cell individually)

    Private Async Function ProcessSelectedRange(ByVal SysCommand As String, CheckMaxToken As Boolean, DoRange As Boolean, DoFormulas As Boolean, SelectionMandatory As Boolean, ByVal UseSecondAPI As Boolean, Optional ShortenPercentValue As Integer = 0, Optional Freestyle As Boolean = False) As Task(Of Boolean)

        Dim excelApp As Excel.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)

        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        Dim NoSelectedCells As Boolean = False
        Dim DoShorten As Boolean = False

        ' Get the used range of the active sheet
        Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim usedRange As Excel.Range = activeSheet.UsedRange

        ' Check if a selection has been made
        If selectedRange Is Nothing Then
            NoSelectedCells = True
        Else
            ' If the entire row, column, or sheet is selected, limit to used range
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' If the intersection results in no cells, set NoSelectedCells to True
            If selectedRange Is Nothing Then
                NoSelectedCells = True
                If Freestyle Or Not SelectionMandatory Then
                    DoRange = True
                    Freestyle = True
                    SysCommand = SP_RangeOfCells
                End If
            End If
        End If

        ' Check if cells are selected and show message if mandatory selection is required
        If NoSelectedCells AndAlso SelectionMandatory Then
            ShowCustomMessageBox("Please select cells (with content) to be processed.")
            Return False
        End If

        ' Check if all selected cells are blocked
        If AreAllCellsBlocked(selectedRange) Then
            ShowCustomMessageBox($"{AN} cannot do anything because the cells are blocked.")
            Return False
        End If

        If ShortenPercentValue > 0 Then
            DoShorten = True
        End If

        Dim MaxToken As Integer = If(UseSecondAPI, INI_MaxOutputToken_2, INI_MaxOutputToken)
        If Not NoSelectedCells And CheckMaxToken And MaxToken > 0 Then

            SelectedText = GetSelectedText(selectedRange, DoFormulas)

            Dim EstimatedTokens As Integer = EstimateTokenCount(SelectedText)

            If EstimatedTokens > MaxToken Then
                ShowCustomMessageBox("The content of the selected cells is larger than the maximum output your LLM can supposedly generate. Therefore, the output may be shorter than expected based on maximum tokens supported, which is " & MaxToken & " tokens. Your input (with formatting information, as the case may be) has an estimated to be " & EstimatedTokens & " tokens). Therefore, check whether the output is complete.", AN, 15)
            End If

        End If

        If Not DoShorten Then
            SysCommand = InterpolateAtRuntime(SysCommand)
        End If

        undoStates.Clear()

        If Not DoRange Then

            Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")
            splash.Show()
            splash.Refresh()

            'Application.ScreenUpdating = False ' Prevent UI updates during processing
            Try
                For Each cell As Excel.Range In selectedRange.Cells

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                        ' Exit the loop
                        Exit For
                    End If
                    Try
                        If Not IsNothing(cell.Value) AndAlso Not CellProtected(cell) Then
                            If cell.HasFormula AndAlso DoFormulas Then
                                ' Handle formulas
                                SelectedText = cell.Formula

                                If DoShorten Then
                                    Dim Textlength As Integer = SelectedText.Length
                                    ShortenLength = (Textlength - (Textlength * (100 - ShortenPercentValue) / 100))
                                    SysCommand = InterpolateAtRuntime(SysCommand)
                                End If

                                Await System.Threading.Tasks.Task.Delay(500)

                                Dim LLMResult As String = Await LLM(SysCommand & " " & SP_Add_KeepFormulasIntact, If(NoSelectedCells, "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0, UseSecondAPI, True)

                                LLMResult = Trim(LLMResult)

                                If Not String.IsNullOrEmpty(LLMResult) Then
                                    LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                                End If
                                If Not String.IsNullOrWhiteSpace(LLMResult) Then
                                    Dim state As New CellState With {
                                                                    .CellAddress = cell.Address,
                                                                    .OldValue = cell.Value,
                                                                    .HadFormula = cell.HasFormula,
                                                                    .OldFormula = If(cell.HasFormula, cell.Formula, "")
                                                                }
                                    Try
                                        cell.Formula = LLMResult ' Replace cell formula
                                        undoStates.Add(state)
                                    Catch ex As Exception
                                        If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                            Try
                                                cell.FormulaLocal = LLMResult
                                                undoStates.Add(state)
                                            Catch ex2 As Exception
                                                If ex2.Message.Contains("HRESULT: 0x800A03EC") Then
                                                    Try
                                                        cell.FormulaLocal = Trim(ConvertFormulaToLocale(LLMResult, excelApp))
                                                        undoStates.Add(state)
                                                    Catch ex3 As Exception
                                                        If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                                            ShowCustomMessageBox($"Error: Excel rejected the formula '{LLMResult}' that {AN} tried to assign to the cell {cell.Address(False, False)}.")
                                                        Else
                                                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{LLMResult}' in cell {cell.Address(False, False)}: {ex.Message}")
                                                        End If
                                                    End Try
                                                Else
                                                    ShowCustomMessageBox($"An error occurred when trying to insert the formula '{LLMResult}' in cell {cell.Address(False, False)}: {ex.Message}")
                                                End If
                                            End Try
                                        Else
                                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{LLMResult}' in cell {cell.Address(False, False)}: {ex.Message}")
                                        End If
                                    End Try
                                End If
                            ElseIf Not cell.HasFormula Then
                                ' Handle plain text cells
                                SelectedText = CStr(cell.Value)

                                Dim regex As New Regex("((\r\n)|\n|\r){2,}$")
                                'Dim trailingCR As Boolean = regex.IsMatch(SelectedText)
                                'Dim trailingCR As Boolean = (SelectedText.EndsWith(vbCrLf) Or SelectedText.EndsWith(vbLf) Or SelectedText.EndsWith(vbCr))

                                If DoShorten Then
                                    Dim Textlength As Integer = SelectedText.Length
                                    ShortenLength = (Textlength - (Textlength * (100 - ShortenPercentValue) / 100))
                                    SysCommand = InterpolateAtRuntime(SysCommand)
                                End If

                                Await System.Threading.Tasks.Task.Delay(500)

                                Dim LLMResult As String = Await LLM(SysCommand, If(NoSelectedCells, "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0, UseSecondAPI)

                                If Not String.IsNullOrEmpty(LLMResult) Then
                                    LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                                End If

                                'If Not trailingCR And LLMResult.EndsWith(ControlChars.CrLf) Then LLMResult = LLMResult.TrimEnd(ControlChars.CrLf)
                                'If Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                                'If Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)

                                LLMResult = Trim(LLMResult).TrimEnd(ControlChars.CrLf).TrimEnd(ControlChars.Lf).TrimEnd(ControlChars.Cr)
                                LLMResult = Trim(LLMResult).TrimEnd(ControlChars.CrLf).TrimEnd(ControlChars.Lf).TrimEnd(ControlChars.Cr)
                                LLMResult = Trim(LLMResult).TrimEnd(ControlChars.CrLf).TrimEnd(ControlChars.Lf).TrimEnd(ControlChars.Cr)

                                If Not String.IsNullOrWhiteSpace(LLMResult) Then
                                    Dim state As New CellState With {
                                                                    .CellAddress = cell.Address,
                                                                    .OldValue = cell.Value,
                                                                    .HadFormula = cell.HasFormula,
                                                                    .OldFormula = If(cell.HasFormula, cell.Formula, "")
                                                                }
                                    cell.Value = LLMResult ' Set the result back to the cell
                                    undoStates.Add(state)
                                End If
                            End If
                        End If

                    Catch ex As Exception
                        ' Log the error and continue with the next cell
                        Debug.WriteLine($"ProcessSelectedRange Error processing cell {cell.Address}: {ex.Message}")
                    End Try
                Next
            Finally
                'Application.ScreenUpdating = True ' Re-enable UI updates
            End Try

            splash.Close()

        Else
            Try

                If NoSelectedCells Then
                    activeSheet.Application.ActiveCell.Select()
                    selectedRange = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
                    If selectedRange Is Nothing Then
                        SelectedText = ""
                        Try
                            SelectedText = $"Current cell = {activeSheet.Application.ActiveCell.Address(False, False)} Text = '{activeSheet.Application.ActiveCell.Text}' Formula = '{activeSheet.Application.ActiveCell.Formula}' (use this for your output unless instructed otherwise)"
                            Debug.WriteLine("NoSelectedCell - SelectedText = " & SelectedText)
                        Catch
                        End Try
                    Else
                        NoSelectedCells = False
                    End If
                End If

                If Not NoSelectedCells Then
                    SelectedText = ConvertRangeToString(selectedRange, DoFormulas)
                End If

                Dim LLMResult As String = Await LLM(SysCommand, If(NoSelectedCells, SelectedText, "<RANGEOFCELLS>" & SelectedText & "</RANGEOFCELLS>"), "", "", 0, UseSecondAPI, False, OtherPrompt)

                LLMResult = LLMResult.Replace("<RANGEOFCELLS>", "").Replace("</RANGEOFCELLS>", "")

                OtherPrompt = ""

                If Not String.IsNullOrEmpty(LLMResult) Then
                    LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                End If

                Dim instructions As New List(Of String)
                instructions = ParseLLMResponse(LLMResult)

                If instructions.Count > 0 Then

                    Dim FinalText = ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, $"Shall {AN} implement this, if possible (don't worry, formulas will be automatically converted in to the locale of the your Excel application)?", AN, True)

                    ' Handle the user's response
                    If Not String.IsNullOrWhiteSpace(FinalText) Then
                        Debug.WriteLine("Finaltext=" & FinalText)
                        instructions = ParseLLMResponse(FinalText)
                        ApplyLLMInstructions(instructions)
                        PutInClipboard(FinalText)
                        ShowCustomMessageBox("Implementation of the instructions completed (to the extent possible). They are also in the clipboard.")
                    End If

                Else

                    Dim FinalText = ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, "If you chose OK, it will be put in the clipboard.", AN)

                    If Not String.IsNullOrWhiteSpace(FinalText) Then PutInClipboard(FinalText)

                End If

            Catch ex As Exception
                MessageBox.Show("Error in Range: " & ex.Message)
            End Try

        End If

        Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()

    End Function

    ' Helpers for the Range Functionality

    Private Function ConvertRangeToString(ByVal CellRange As Excel.Range, ByVal IncludeFormulas As Boolean) As String
        Dim output As String = String.Empty


        Dim rowCount As Integer = CellRange.Rows.Count
        Dim colCount As Integer = CellRange.Columns.Count

        If rowCount = 1 And colCount = 1 Then
            output = "Current Cell: " & CellRange.Address & vbCrLf
        End If

        ' Loop through each cell in the range
        For Each cell As Excel.Range In CellRange.Cells
            ' Get the cell address (e.g., "A1", "B4")
            output &= "Cell " & cell.Address(False, False) & ": " & vbCrLf

            ' Get the cell value
            Dim cellValue As String = If(cell.Value IsNot Nothing, cell.Value.ToString(), "(empty)")
            output &= "  Value: " & cellValue & vbCrLf

            ' Check if the cell contains a formula, and include it
            If IncludeFormulas Then
                If cell.HasFormula Then
                    Dim cellFormula As String = cell.Formula
                    output &= "  Formula: " & cellFormula & vbCrLf
                Else
                    output &= "  Formula: none" & vbCrLf
                End If
            End If

            ' Add a separator line between cells
            output &= New String("-"c, 40) & vbCrLf
        Next

        Return output
    End Function
    Public Function ParseLLMResponse(ByVal Response As String) As List(Of String)
        Dim instructions As New List(Of String)()
        Dim startPos As Integer, instructionEnd As Integer
        Dim tempInstruction As String
        Dim cellPattern As String

        ' Ensure we remove any newlines that might affect parsing
        Response = Response.Replace(vbCrLf, " ").Replace(vbLf, " ")

        ' Pattern for finding Cell
        cellPattern = "[Cell:"

        ' Start parsing the response
        startPos = Response.IndexOf(cellPattern)

        Do While startPos >= 0
            ' Find next cell occurrence to extract the block between this and next [Cell:]
            instructionEnd = Response.IndexOf(cellPattern, startPos + cellPattern.Length)

            ' If there's no further [Cell:], capture till the end of the string
            If instructionEnd = -1 Then instructionEnd = Response.Length

            ' Extract the instruction block between the current and next [Cell:]
            tempInstruction = Response.Substring(startPos, instructionEnd - startPos)
            instructions.Add(tempInstruction)

            ' Move to the next instruction start, exit if at the end
            startPos = Response.IndexOf(cellPattern, instructionEnd)
        Loop

        Return instructions
    End Function
    Sub ApplyLLMInstructions(ByVal instructions As List(Of String))

        Dim instruction As String
        Dim cellAddress As String
        Dim formulaOrValue As String
        Dim formulaOrValueLocale As String = ""
        Dim cleanedValue As String
        Dim ii As Integer

        ' Get the active Excel application and sheet
        Dim excelApp As Excel.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)
        Dim activeSheet As Worksheet = CType(excelApp.ActiveSheet, Worksheet)

        ii = 0

        undoStates.Clear()

        Dim splash As New SplashScreen("Implementing... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        Debug.WriteLine("Instructions: " & String.Join(Environment.NewLine, instructions))

        ' Loop through the parsed instructions and ask for confirmation before applying
        For Each instruction In instructions

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

            cellAddress = GetCellFromInstruction(instruction)
            formulaOrValue = GetFormulaOrValueFromInstruction(instruction)

            If Not String.IsNullOrWhiteSpace(cellAddress) AndAlso Not String.IsNullOrWhiteSpace(formulaOrValue) Then
                ii += 1
                Debug.WriteLine($"Processing: Cell='{cellAddress}', Value='{formulaOrValue}'")

                If formulaOrValue.StartsWith("=") Then formulaOrValueLocale = Trim(ConvertFormulaToLocale(formulaOrValue, excelApp))

                Try
                    If activeSheet IsNot Nothing AndAlso activeSheet.Range(cellAddress) IsNot Nothing Then
                        Dim targetRange As Range
                        Try
                            ' Ensure the address is valid before accessing it
                            If Regex.IsMatch(cellAddress, "^[A-Z]+\d+$") Then
                                targetRange = activeSheet.Range(cellAddress)

                                ' Handle merged cells properly
                                If targetRange.MergeCells Then
                                    targetRange = targetRange.MergeArea.Cells(1, 1)
                                End If

                                If formulaOrValue.StartsWith("=") Then
                                    Dim state As New CellState With {
                                                                    .CellAddress = targetRange.Address,
                                                                    .OldValue = targetRange.Value,
                                                                    .HadFormula = targetRange.HasFormula,
                                                                    .OldFormula = If(targetRange.HasFormula, targetRange.Formula, "")
                                                                }
                                    ' Fix cell format issues
                                    targetRange.NumberFormat = "General"
                                    targetRange.Value = ""
                                    Try
                                        targetRange.Formula = formulaOrValue
                                        undoStates.Add(state)
                                    Catch ex As Exception
                                        If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                            Try
                                                targetRange.FormulaLocal = formulaOrValue
                                                undoStates.Add(state)
                                            Catch ex2 As Exception
                                                If ex2.Message.Contains("HRESULT: 0x800A03EC") Then
                                                    Try
                                                        targetRange.FormulaLocal = formulaOrValueLocale
                                                        undoStates.Add(state)
                                                    Catch ex3 As Exception
                                                        If ex3.Message.Contains("HRESULT: 0x800A03EC") Then
                                                            ShowCustomMessageBox($"Error: Excel rejected the formula '{formulaOrValue}' that {AN} tried to assign to the cell {cellAddress}.")
                                                        Else
                                                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{formulaOrValue}' in cell {cellAddress}: {ex.Message}")
                                                        End If
                                                    End Try
                                                Else
                                                    ShowCustomMessageBox($"An error occurred when trying to insert the formula '{formulaOrValue}' in cell {cellAddress}: {ex.Message}")
                                                End If
                                            End Try
                                        Else
                                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{formulaOrValue}' in cell {cellAddress}: {ex.Message}")
                                        End If
                                    End Try
                                Else
                                    Dim state As New CellState With {
                                                                    .CellAddress = targetRange.Address,
                                                                    .OldValue = targetRange.Value,
                                                                    .HadFormula = targetRange.HasFormula,
                                                                    .OldFormula = If(targetRange.HasFormula, targetRange.Formula, "")
                                                                }
                                    ' Assign values properly
                                    If IsNumeric(formulaOrValue) Then
                                        targetRange.Value = formulaOrValue
                                    Else
                                        ' Remove unwanted apostrophes
                                        cleanedValue = formulaOrValue.Trim("'")
                                        targetRange.NumberFormat = "@" ' Ensure it's stored as text
                                        targetRange.Value = cleanedValue
                                    End If
                                    undoStates.Add(state)
                                End If
                            Else
                                Debug.WriteLine($"Invalid cell address: {cellAddress}")
                            End If
                        Catch ex As Exception
                            If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                ShowCustomMessageBox($"Error: Excel rejected the formula '{formulaOrValue}' that {AN} tried to assign to the cell {cellAddress}.")
                            Else
                                ShowCustomMessageBox($"An error occurred when trying to insert the formula '{formulaOrValue}' in cell {cellAddress}: {ex.Message}")
                            End If
                        End Try
                    Else
                        Debug.WriteLine($"Invalid or missing cell address: {cellAddress}")
                    End If
                Catch ex As Exception
                    If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                        ShowCustomMessageBox($"Error: Excel rejected the formula '{formulaOrValue}' that {AN} tried to assign to the cell {cellAddress}.")
                    Else
                        ShowCustomMessageBox($"An error occurred when trying to insert the formula '{formulaOrValue}' in cell {cellAddress}: {ex.Message}")
                    End If
                End Try
            End If
        Next
        splash.Close()

    End Sub


    Public Function ConvertFormulaToLocale(ByVal englishFormula As String, ByVal excelApp As Excel.Application) As String
        Dim wb As Workbook = Nothing
        Dim ws As Worksheet = Nothing
        Dim localizedFormula As String = ""

        ' Disable screen updating to prevent Excel from flashing
        Dim previousScreenUpdating As Boolean = excelApp.ScreenUpdating
        Dim previousDisplayAlerts As Boolean = excelApp.DisplayAlerts

        Try
            excelApp.ScreenUpdating = False ' Hide UI updates
            excelApp.DisplayAlerts = False ' Prevent pop-ups (e.g., when closing the temp workbook)

            ' Create a temporary worksheet
            wb = excelApp.Workbooks.Add()
            ws = CType(wb.Sheets(1), Worksheet)

            ' Set the formula using English syntax
            Dim tempRange As Excel.Range = ws.Range("A1")
            tempRange.Formula = englishFormula

            ' Retrieve the formula in the local Excel language
            localizedFormula = tempRange.FormulaLocal

            ' Close workbook without saving
            wb.Close(False)
        Catch ex As Exception
            localizedFormula = englishFormula ' Fallback to English if an error occurs
        Finally
            ' Restore Excel's UI settings
            excelApp.ScreenUpdating = previousScreenUpdating
            excelApp.DisplayAlerts = previousDisplayAlerts

            ' Release COM objects
            ReleaseObject(ws)
            ReleaseObject(wb)
        End Try

        Return localizedFormula
    End Function


    ' Helper function to release COM objects
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Function GetCellFromInstruction(ByVal instruction As String) As String
        Dim startPos As Integer = instruction.IndexOf("[Cell: ") + 7
        Dim endPos As Integer = instruction.IndexOf("]", startPos)
        If startPos > 6 AndAlso endPos > startPos Then
            Return instruction.Substring(startPos, endPos - startPos).Trim()
        End If
        Return String.Empty
    End Function
    Function GetFormulaOrValueFromInstruction(ByVal instruction As String) As String
        Dim startPos As Integer = -1

        If instruction.Contains("[Formula: ") Then
            startPos = instruction.IndexOf("[Formula: ") + 10
        ElseIf instruction.Contains("[Value: ") Then
            startPos = instruction.IndexOf("[Value: ") + 8
        End If

        If startPos > -1 Then
            Dim endPos As Integer = instruction.IndexOf("]", startPos)
            If endPos > startPos Then
                Return instruction.Substring(startPos, endPos - startPos).Trim()
            End If
        End If

        Return String.Empty
    End Function

    ' Excel Helpers

    Public Sub AdjustHeight(Optional Silent As Boolean = False)

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active Excel worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the current selection
            Dim selectedRange As Excel.Range = Globals.ThisAddIn.Application.Selection
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' Check if the selection is empty or null
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then
                Dim result As Integer = 0
                If Not Silent Then
                    result = ShowCustomYesNoBox("No cells are selected. Would you like to perform the operation on the entire worksheet?", "Yes", "No", "Adjust Height")
                End If
                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                Else
                    If Not Silent Then ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Perform AutoFit on the rows of the selected range to ensure initial proper height
            selectedRange.Rows.AutoFit()


            ' Prepare dictionaries for tracking row heights
            Dim rowOriginalHeights As New Dictionary(Of Integer, Double)()
            Dim rowMaxHeights As New Dictionary(Of Integer, Double)()

            ' Initialize these dictionaries for each row in the selection
            For Each oneRow As Excel.Range In selectedRange.Rows
                Dim rowIndex As Integer = oneRow.Row
                Dim currentHeight As Double = activeSheet.Rows(rowIndex).RowHeight
                rowOriginalHeights(rowIndex) = currentHeight
                ' Start the max at whatever the row is currently
                rowMaxHeights(rowIndex) = currentHeight
            Next

            splash.Show()
            splash.Refresh()


            ' Iterate through each cell in the selection
            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    ' Exit the loop
                    Exit For
                End If

                If cell Is Nothing Then Continue For

                ' We'll always enable wrapping so AutoFit will compute multi-line height
                cell.WrapText = True

                Dim wasMerged As Boolean = cell.MergeCells
                Dim mergeArea As Excel.Range = If(wasMerged, cell.MergeArea, cell)

                ' Temporarily store the row index for dictionary look-up
                Dim rowIndex As Integer = mergeArea.Row

                ' We'll measure how tall Excel wants to make this cell
                Dim newHeight As Double = 0

                If wasMerged Then

                    ' Store the original column widths for each column
                    Dim firstColIndex As Integer = mergeArea.Column
                    Dim totalCols As Integer = mergeArea.Columns.Count
                    Dim originalWidths As New List(Of Double)

                    For iCol As Integer = 0 To totalCols - 1
                        Dim colWidth As Double =
                        activeSheet.Columns(firstColIndex + iCol).ColumnWidth
                        originalWidths.Add(colWidth)
                    Next

                    ' Sum the widths so we can set it on the first column after unmerging
                    Dim combinedWidth As Double = originalWidths.Sum()

                    ' Unmerge
                    mergeArea.UnMerge()

                    ' Set only the first column to the combined width so AutoFit sees the "true" width
                    activeSheet.Columns(firstColIndex).ColumnWidth = combinedWidth

                    ' Autofit (note: must do autofit on entire row(s) that the cell spans)
                    mergeArea.Rows.AutoFit()

                    ' Capture the new row height
                    newHeight = mergeArea.RowHeight

                    ' Restore the original column widths
                    For iCol As Integer = 0 To totalCols - 1
                        activeSheet.Columns(firstColIndex + iCol).ColumnWidth = originalWidths(iCol)
                    Next

                    ' Re-merge
                    Dim remergeRange As Excel.Range = activeSheet.Range(
                    activeSheet.Cells(mergeArea.Row, firstColIndex),
                    activeSheet.Cells(mergeArea.Row, firstColIndex + totalCols - 1)
                )
                    remergeRange.Merge()

                Else
                    ' If not merged, simply use AutoFit
                    mergeArea.Rows.AutoFit()
                    newHeight = mergeArea.RowHeight
                End If

                ' Store the maximum needed height for this row so far
                If rowMaxHeights.ContainsKey(rowIndex) Then
                    ' Compare existing max with newly measured height
                    If newHeight > rowMaxHeights(rowIndex) Then
                        rowMaxHeights(rowIndex) = newHeight
                    End If
                End If

            Next


            ' Now set each row’s height to the maximum of:
            For Each rowIndex As Integer In rowMaxHeights.Keys.ToList()

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    ' Exit the loop
                    Exit For
                End If
                Dim finalHeight As Double = Math.Max(rowMaxHeights(rowIndex), rowOriginalHeights(rowIndex))
                If finalHeight > 409 Then
                    finalHeight = 409
                End If

                activeSheet.Rows(rowIndex).RowHeight = finalHeight
            Next

        Catch ex As System.Exception
            MessageBox.Show($"Error in AdjustHeight: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try

    End Sub

    Public Sub AdjustLegacyNotes()

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active Excel worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the current selection
            Dim selectedRange As Excel.Range = Globals.ThisAddIn.Application.Selection
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' Check if the selection is empty or null
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then
                Dim result As Integer = ShowCustomYesNoBox(
                "No cells are selected. Would you like to perform the operation on the entire worksheet?",
                "Yes",
                "No",
                "Adjust Legacy Notes"
            )

                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                Else
                    ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Perform AutoFit on the rows of the selected range to ensure initial proper height
            selectedRange.Rows.AutoFit()


            splash.Show()
            splash.Refresh()

            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    ' Exit the loop
                    Exit For
                End If

                If cell Is Nothing Then Continue For

                If cell.Comment IsNot Nothing Then

                    ' Ensure the note box dimensions are at least 70 wide and 20 high, and no more than 200 wide
                    Dim comment As Excel.Comment = cell.Comment
                    With comment.Shape

                        .TextFrame.AutoSize = True
                        Dim MinimumHeight As Double = .Height

                        .TextFrame.AutoSize = False

                        ' Enforce width constraints
                        If .Width < 70 Then
                            .Width = 70
                        End If
                        If .Width > 250 Then
                            .Width = 250
                        End If

                        ' Dynamically calculate and set height
                        Dim textLength As Integer = Len(comment.Text)
                        Dim lines As Integer = CInt(Math.Ceiling(textLength / (250 / (.TextFrame.Characters.Font.Size - 2)))) ' Approximation based on average char width
                        Dim lineHeight As Double = .TextFrame.Characters.Font.Size + 2 ' Approximate height per line in points
                        Dim requiredHeight As Double = Math.Max(MinimumHeight, (lines * lineHeight)) + 10

                        If lines > 1 Then .Width = 250

                        .Height = requiredHeight

                    End With
                End If

            Next

        Catch ex As System.Exception
            MessageBox.Show($"Error in AdjustLegacyNotes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try

    End Sub


    Private Shared LastRegexPattern As String = String.Empty
    Private Shared LastRegexOptions As String = String.Empty
    Private Shared LastRegexReplace As String = String.Empty

    Public Sub RegexSearchReplace()

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the selected range
            Dim selectedRange As Excel.Range = Globals.ThisAddIn.Application.Selection
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            Dim processEntireSheet As Boolean = False

            ' If no range is selected, ask to process the entire worksheet
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then

                Dim result As Integer = ShowCustomYesNoBox("No cells are selected. Would you like to perform the operation on the entire worksheet?", "Yes", "No", "Regex Search & Replace")

                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                    processEntireSheet = True
                Else
                    ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Step 1: Get regex patterns
            Dim regexPattern As String = ShowCustomInputBox("Step 1: Enter your Regex pattern(s), one per line (more info about Regex: vischerlnk.com/regexinfo):", "Regex Search & Replace", False, LastRegexPattern)?.Trim()
            If String.IsNullOrEmpty(regexPattern) Then Exit Sub

            ' Step 2: Get regex options
            Dim optionsInput As String = ShowCustomInputBox("Enter regex option(s) (i for IgnoreCase, m for Multiline, s for Singleline, c for Compiled, r for RightToLeft, e for ExplicitCapture):", "Regex Search & Replace", True, LastRegexOptions)

            Dim regexOptions As RegexOptions = RegexOptions.None

            If Not String.IsNullOrEmpty(optionsInput) Then
                If optionsInput.Contains("i") Then regexOptions = regexOptions Or RegexOptions.IgnoreCase
                If optionsInput.Contains("m") Then regexOptions = regexOptions Or RegexOptions.Multiline
                If optionsInput.Contains("s") Then regexOptions = regexOptions Or RegexOptions.Singleline
                If optionsInput.Contains("c") Then regexOptions = regexOptions Or RegexOptions.Compiled
                If optionsInput.Contains("r") Then regexOptions = regexOptions Or RegexOptions.RightToLeft
                If optionsInput.Contains("e") Then regexOptions = regexOptions Or RegexOptions.ExplicitCapture
            End If

            ' Step 3: Get replacement text
            Dim replacementText As String = ShowCustomInputBox("Step 2: Enter your replacement text(s), one on each line, matching to your pattern(s):", "Regex Search & Replace", False, LastRegexReplace)

            ' Update the last-used regex pattern and options
            LastRegexPattern = regexPattern
            LastRegexOptions = optionsInput
            LastRegexReplace = replacementText

            ' Split patterns and replacements into lines
            Dim patterns() As String = regexPattern.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
            Dim replacements() As String = If(Not String.IsNullOrEmpty(replacementText), replacementText.Split(New String() {Environment.NewLine}, StringSplitOptions.None), Nothing)

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

            splash.Show()
            splash.Refresh()

            ' Perform replacements
            Dim totalReplacements As Integer = 0

            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    ' Exit the loop
                    Exit For
                End If
                If cell.Value2 IsNot Nothing AndAlso TypeOf cell.Value2 Is String Then
                    Dim cellText As String = cell.Value2.ToString()

                    For i As Integer = 0 To patterns.Length - 1
                        Dim regex As New Regex(patterns(i), regexOptions)
                        Dim replacement As String = If(replacements IsNot Nothing, replacements(i), Nothing)

                        ' Perform replacement
                        Dim newText As String = regex.Replace(cellText, replacement)
                        If newText <> cellText Then
                            totalReplacements += 1
                            cell.Value2 = newText
                        End If
                    Next
                End If
            Next

            ShowCustomMessageBox($"{totalReplacements} replacement(s) made in the selected cells.")

        Catch ex As System.Exception
            MessageBox.Show($"Error in RegexSearchReplace: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try
    End Sub

    ' Other Helpers 

    Function GetSelectedText(selectedRange As Excel.Range, DoFormulas As Boolean) As String
        Dim selectedTextBuilder As New StringBuilder()

        For Each cell As Excel.Range In selectedRange.Cells
            If Not IsNothing(cell.Value) AndAlso Not CellProtected(cell) Then
                If cell.HasFormula Then
                    If DoFormulas Then
                        selectedTextBuilder.AppendLine(cell.Formula)
                    End If
                Else
                    selectedTextBuilder.AppendLine(CStr(cell.Value))
                End If
            End If
        Next

        Return selectedTextBuilder.ToString()
    End Function
    Private Function AreAllCellsBlocked(ByVal rng As Excel.Range) As Boolean
        Dim allLocked As Boolean = True ' Assume all cells are locked by default

        If rng Is Nothing Then Return False

        ' Check if the worksheet is protected
        If rng.Worksheet.ProtectContents Then
            ' Iterate through each cell in the range
            For Each cell As Excel.Range In rng.Cells
                ' Check if the cell is locked and cannot be changed
                If Not CellProtected(cell) Then
                    allLocked = False
                    Exit For
                End If
            Next
        Else
            ' Worksheet is not protected, so cells can be modified
            allLocked = False
        End If

        ' Return True if all cells are locked and the worksheet is protected
        Return allLocked
    End Function
    Private Function OldCellProtected(ByVal cell As Excel.Range) As Boolean
        ' Check if the cell is locked and the worksheet is protected
        If cell.Worksheet.ProtectContents Then
            If cell.Locked AndAlso Not cell.Worksheet.Protection.AllowEditRanges.Cast(Of Excel.AllowEditRange).Any(Function(r) r.Range.Address = cell.Address) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Function CellProtected(ByVal cell As Excel.Range) As Boolean
        ' If the worksheet is not protected, no cell is actually protected
        If Not cell.Worksheet.ProtectContents Then
            Return False
        End If

        ' If the cell is not locked, it is not protected
        If Not cell.Locked Then
            Return False
        End If

        ' Check whether cell is in any AllowEditRange
        For Each aer As Excel.AllowEditRange In cell.Worksheet.Protection.AllowEditRanges
            ' If Intersect is not Nothing, the cell is within that allow-edit range
            If cell.Application.Intersect(aer.Range, cell) IsNot Nothing Then
                ' The cell can be edited => not protected
                Return False
            End If
        Next

        ' If it is locked, sheet is protected, and no allow-edit range covers the cell => it is effectively protected
        Return True
    End Function


    Public Sub UndoAction()
        Try
            Dim app As Excel.Application = Globals.ThisAddIn.Application

            ' Process each saved state to restore the previous value or formula.
            For Each state In undoStates
                Dim rng As Excel.Range = app.Range(state.CellAddress)
                If state.HadFormula Then
                    rng.Formula = state.OldFormula
                Else
                    rng.Value = state.OldValue
                End If
            Next

            ' Clear the undo state after restoring.
            undoStates.Clear()

            Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()

        Catch ex As System.Exception
            MessageBox.Show("Error during undo (" & ex.Message & ").")
        End Try
    End Sub

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

        Dim xlApp As Microsoft.Office.Interop.Excel.Application = Me.Application

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
                {"PreCorrection", "Additional instruction for prompts"},
                {"PostCorrection", "Prompt to apply after queries"},
                {"Language1", "Default translation language 1"},
                {"Language2", "Default translation language 2"},
                {"PromptLibPath", "Prompt library file"}
            }

        Dim SettingsTips As New Dictionary(Of String, String) From {
                {"Temperature", "The higher, the more creative the LLM will be (0.0-2.0)"},
                {"Timeout", "In milliseconds"},
                {"Temperature_2", "The higher, the more creative the LLM will be (0.0-2.0)"},
                {"Timeout_2", "In milliseconds"},
                {"DoubleS", "For Switzerland"},
                {"PreCorrection", "Add prompting text that will be added to all basic requests (e.g., for special language tasks)"},
                {"PostCorrection", "Add a prompt that will be applied to each result before it is further processed (slow!)"},
                {"Language1", "The language (in English) that will be used for the first quick access button in the ribbon"},
                {"Language2", "The language (in English) that will be used for the second quick access button in the ribbon"},
                {"PromptLibPath", "The filename (including path, support environmental variables) for your prompt library (if any)"}
                }
        ShowSettingsWindow(Settings, SettingsTips)

        Dim splash As New SplashScreen("Updating menu following your changes ...")
        splash.Show()
        splash.Refresh()

        AddContextMenu()

        splash.Close()

    End Sub

    Public Function GetAPIConfiguration(UseSecondAPI As Boolean) As String
        Dim config As New List(Of String)()

        If UseSecondAPI Then
            config.Add("INI_OAuth2§§" & INI_OAuth2_2.ToString)
            config.Add("INI_OAuth2ClientMail§§" & INI_OAuth2ClientMail_2)
            config.Add("INI_OAuth2Scopes§§" & INI_OAuth2Scopes_2)
            config.Add("INI_OAuth2Endpoint§§" & INI_OAuth2Endpoint_2)
            config.Add("INI_OAuth2ATExpiry§§" & INI_OAuth2ATExpiry_2.ToString)
            config.Add("INI_APIKey§§" & INI_APIKey_2)
            config.Add("INI_Temperature§§" & INI_Temperature_2.ToString)
            config.Add("INI_Timeout§§" & INI_Timeout_2)
            config.Add("INI_MaxOutputToken§§" & INI_MaxOutputToken_2.ToString)
            config.Add("INI_Model§§" & INI_Model_2)
            config.Add("INI_Endpoint§§" & INI_Endpoint_2)
            config.Add("INI_HeaderA§§" & INI_HeaderA_2)
            config.Add("INI_HeaderB§§" & INI_HeaderB_2)
            config.Add("INI_APICall§§" & INI_APICall_2)
            config.Add("INI_Response§§" & INI_Response_2)
            config.Add("DecodedAPI§§" & DecodedAPI_2)
        Else
            config.Add("INI_OAuth2§§" & INI_OAuth2.ToString)
            config.Add("INI_OAuth2ClientMail§§" & INI_OAuth2ClientMail)
            config.Add("INI_OAuth2Scopes§§" & INI_OAuth2Scopes)
            config.Add("INI_OAuth2Endpoint§§" & INI_OAuth2Endpoint)
            config.Add("INI_OAuth2ATExpiry§§" & INI_OAuth2ATExpiry.ToString)
            config.Add("INI_APIKey§§" & INI_APIKey)
            config.Add("INI_Temperature§§" & INI_Temperature.ToString)
            config.Add("INI_Timeout§§" & INI_Timeout)
            config.Add("INI_MaxOutputToken§§" & INI_MaxOutputToken.ToString)
            config.Add("INI_Model§§" & INI_Model)
            config.Add("INI_Endpoint§§" & INI_Endpoint)
            config.Add("INI_HeaderA§§" & INI_HeaderA)
            config.Add("INI_HeaderB§§" & INI_HeaderB)
            config.Add("INI_APICall§§" & INI_APICall)
            config.Add("INI_Response§§" & INI_Response)
            config.Add("DecodedAPI§§" & DecodedAPI)
        End If

        ' Join the list into a single string with a delimiter
        Return String.Join("@@@", config)
    End Function

End Class
