' Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See License.txt or https://vischer.com/redink for more information.
'
' 26.2.2025
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

Imports System.Reflection.Emit
Imports SharedLibrary.SharedLibrary.SharedContext
Imports Microsoft.Office.Interop.Word
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Management
Imports System.IO
Imports Microsoft.Win32
Imports HtmlAgilityPack
Imports Newtonsoft.Json
Imports System.Text.RegularExpressions
Imports System.Text
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.OpenSsl
Imports Org.BouncyCastle.Security
Imports System.Net.Http
Imports Org.BouncyCastle.Utilities.IO.Pem
Imports UglyToad.PdfPig
Imports UglyToad.PdfPig.Content
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Org.BouncyCastle.Crypto
Imports System.Deployment.Application
Imports Microsoft.Office.Interop
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Markdig.Extensions


Namespace SharedLibrary

    Public Class SharedContext
        Implements ISharedContext

        Public Interface ISharedContext

#Region "Shared Properties"

            Property INI_APIKey As String
            Property INI_APIKeyBack As String
            Property INI_Temperature As String
            Property INI_Timeout As Long
            Property INI_MaxOutputToken As Integer
            Property INI_Model As String
            Property INI_Endpoint As String
            Property INI_HeaderA As String
            Property INI_HeaderB As String
            Property INI_APICall As String
            Property INI_Response As String
            Property INI_DoubleS As Boolean
            Property INI_PreCorrection As String
            Property INI_PostCorrection As String
            Property INI_APIEncrypted As Boolean
            Property INI_APIKeyPrefix As String
            Property INI_MarkupMethodOutlook As Integer
            Property INI_MarkupDiffCap As Integer
            Property INI_MarkupRegexCap As Integer

            Property INI_OpenSSLPath As String
            Property INI_OAuth2 As Boolean
            Property INI_OAuth2ClientMail As String
            Property INI_OAuth2Scopes As String
            Property INI_OAuth2Endpoint As String
            Property INI_OAuth2ATExpiry As Long
            Property INI_SecondAPI As Boolean
            Property INI_APIKey_2 As String
            Property INI_APIKeyBack_2 As String
            Property INI_Temperature_2 As String
            Property INI_Timeout_2 As Long
            Property INI_MaxOutputToken_2 As Integer
            Property INI_Model_2 As String
            Property INI_Endpoint_2 As String
            Property INI_HeaderA_2 As String
            Property INI_HeaderB_2 As String
            Property INI_APICall_2 As String
            Property INI_Response_2 As String
            Property INI_APIEncrypted_2 As Boolean
            Property INI_APIKeyPrefix_2 As String
            Property INI_OAuth2_2 As Boolean
            Property INI_OAuth2ClientMail_2 As String
            Property INI_OAuth2Scopes_2 As String
            Property INI_OAuth2Endpoint_2 As String
            Property INI_OAuth2ATExpiry_2 As Long
            Property INI_APIDebug As Boolean
            Property INI_UsageRestrictions As String
            Property INI_Language1 As String
            Property INI_Language2 As String
            Property INI_KeepFormat1 As Boolean
            Property INI_KeepFormat2 As Boolean
            Property INI_KeepFormatCap As Integer
            Property INI_KeepParaFormatInline As Boolean
            Property INI_ReplaceText1 As Boolean
            Property INI_ReplaceText2 As Boolean
            Property INI_DoMarkupOutlook As Boolean
            Property INI_DoMarkupWord As Boolean
            Property INI_RoastMe As Boolean

            Property DecodedAPI As String
            Property DecodedAPI_2 As String
            Property TokenExpiry As DateTime
            Property TokenExpiry_2 As DateTime

            Property Codebasis As String

            Property GPTSetupError As Boolean
            Property INIloaded As Boolean

            Property RDV As String
            Property InitialConfigFailed As Boolean
            Property INI_ContextMenu As Boolean
            Property INI_UpdateCheckInterval As Integer
            Property INI_UpdatePath As String
            Property INI_SpeechModelPath As String
            Property INI_TTSEndpoint As String
            Property SP_Translate As String
            Property SP_Correct As String
            Property SP_Improve As String
            Property SP_Explain As String
            Property SP_SuggestTitles As String
            Property SP_Friendly As String
            Property SP_Convincing As String
            Property SP_NoFillers As String
            Property SP_Podcast As String
            Property SP_Shorten As String
            Property SP_Summarize As String
            Property SP_MailReply As String
            Property SP_MailSumup As String
            Property SP_FreestyleText As String
            Property SP_FreestyleNoText As String
            Property SP_SwitchParty As String
            Property SP_Anonymize As String
            Property SP_ContextSearch As String
            Property SP_ContextSearchMulti As String
            Property SP_WriteNeatly As String
            Property SP_RangeOfCells As String
            Property SP_Add_KeepFormulasIntact As String
            Property SP_Add_KeepHTMLIntact As String
            Property SP_Add_KeepInlineIntact As String
            Property SP_Add_Bubbles As String
            Property SP_Add_Revisions As String
            Property SP_MarkupRegex As String
            Property SP_ChatWord As String
            Property SP_Add_ChatWord_Commands As String
            Property INI_ChatCap As Integer

            Property INI_ISearch As Boolean
            Property INI_ISearch_Approve As Boolean
            Property INI_ISearch_URL As String
            Property INI_ISearch_ResponseURLStart As String
            Property INI_ISearch_ResponseMask1 As String
            Property INI_ISearch_ResponseMask2 As String
            Property INI_ISearch_Name As String
            Property INI_ISearch_Tries As Integer
            Property INI_ISearch_Results As Integer
            Property INI_ISearch_MaxDepth As Integer
            Property INI_ISearch_Timeout As Long
            Property INI_ISearch_SearchTerm_SP As String
            Property INI_ISearch_Apply_SP_Markup As String
            Property INI_ISearch_Apply_SP As String
            Property INI_Placeholder_03 As String
            Property INI_Lib As Boolean
            Property INI_Lib_File As String
            Property INI_Lib_Timeout As Long
            Property INI_Lib_Find_SP As String
            Property INI_Lib_Apply_SP As String
            Property INI_Lib_Apply_SP_Markup As String
            Property INI_Placeholder_01 As String
            Property INI_Placeholder_02 As String
            Property INI_MarkupMethodHelper As Integer
            Property INI_MarkupMethodWord As Integer
            Property INI_ShortcutsWordExcel As String
            Property INI_PromptLib As Boolean
            Property INI_PromptLibPath As String
            Property INI_PromptLibPath_Transcript As String
            Property PromptLibrary() As List(Of String)
            Property PromptTitles() As List(Of String)
            Property MenusAdded As Boolean



#End Region

        End Interface

#Region "Shared Properties 2"

        Public Sub New()
            ' Initialize the PromptTitles and PromptLibrary properties
            PromptTitles = New List(Of String)()
            PromptLibrary = New List(Of String)()
        End Sub

        Public Property INI_APIKey As String Implements ISharedContext.INI_APIKey
        Public Property INI_APIKeyBack As String Implements ISharedContext.INI_APIKeyBack
        Public Property INI_Temperature As String Implements ISharedContext.INI_Temperature
        Public Property INI_Timeout As Long Implements ISharedContext.INI_Timeout
        Public Property INI_MaxOutputToken As Integer Implements ISharedContext.INI_MaxOutputToken
        Public Property INI_Model As String Implements ISharedContext.INI_Model
        Public Property INI_Endpoint As String Implements ISharedContext.INI_Endpoint
        Public Property INI_HeaderA As String Implements ISharedContext.INI_HeaderA
        Public Property INI_HeaderB As String Implements ISharedContext.INI_HeaderB
        Public Property INI_APICall As String Implements ISharedContext.INI_APICall
        Public Property INI_Response As String Implements ISharedContext.INI_Response
        Public Property INI_DoubleS As Boolean Implements ISharedContext.INI_DoubleS
        Public Property INI_PreCorrection As String Implements ISharedContext.INI_PreCorrection
        Public Property INI_PostCorrection As String Implements ISharedContext.INI_PostCorrection
        Public Property INI_APIEncrypted As Boolean Implements ISharedContext.INI_APIEncrypted
        Public Property INI_APIKeyPrefix As String Implements ISharedContext.INI_APIKeyPrefix
        Public Property INI_MarkupMethodOutlook As Integer Implements ISharedContext.INI_MarkupMethodOutlook
        Public Property INI_MarkupDiffCap As Integer Implements ISharedContext.INI_MarkupDiffCap
        Public Property INI_MarkupRegexCap As Integer Implements ISharedContext.INI_MarkupRegexCap
        Public Property INI_OpenSSLPath As String Implements ISharedContext.INI_OpenSSLPath
        Public Property INI_OAuth2 As Boolean Implements ISharedContext.INI_OAuth2
        Public Property INI_OAuth2ClientMail As String Implements ISharedContext.INI_OAuth2ClientMail
        Public Property INI_OAuth2Scopes As String Implements ISharedContext.INI_OAuth2Scopes
        Public Property INI_OAuth2Endpoint As String Implements ISharedContext.INI_OAuth2Endpoint
        Public Property INI_OAuth2ATExpiry As Long Implements ISharedContext.INI_OAuth2ATExpiry
        Public Property INI_SecondAPI As Boolean Implements ISharedContext.INI_SecondAPI
        Public Property INI_APIKey_2 As String Implements ISharedContext.INI_APIKey_2
        Public Property INI_APIKeyBack_2 As String Implements ISharedContext.INI_APIKeyBack_2
        Public Property INI_Temperature_2 As String Implements ISharedContext.INI_Temperature_2
        Public Property INI_Timeout_2 As Long Implements ISharedContext.INI_Timeout_2
        Public Property INI_MaxOutputToken_2 As Integer Implements ISharedContext.INI_MaxOutputToken_2
        Public Property INI_Model_2 As String Implements ISharedContext.INI_Model_2
        Public Property INI_Endpoint_2 As String Implements ISharedContext.INI_Endpoint_2
        Public Property INI_HeaderA_2 As String Implements ISharedContext.INI_HeaderA_2
        Public Property INI_HeaderB_2 As String Implements ISharedContext.INI_HeaderB_2
        Public Property INI_APICall_2 As String Implements ISharedContext.INI_APICall_2
        Public Property INI_Response_2 As String Implements ISharedContext.INI_Response_2
        Public Property INI_APIEncrypted_2 As Boolean Implements ISharedContext.INI_APIEncrypted_2
        Public Property INI_APIKeyPrefix_2 As String Implements ISharedContext.INI_APIKeyPrefix_2
        Public Property INI_OAuth2_2 As Boolean Implements ISharedContext.INI_OAuth2_2
        Public Property INI_OAuth2ClientMail_2 As String Implements ISharedContext.INI_OAuth2ClientMail_2
        Public Property INI_OAuth2Scopes_2 As String Implements ISharedContext.INI_OAuth2Scopes_2
        Public Property INI_OAuth2Endpoint_2 As String Implements ISharedContext.INI_OAuth2Endpoint_2
        Public Property INI_OAuth2ATExpiry_2 As Long Implements ISharedContext.INI_OAuth2ATExpiry_2
        Public Property INI_APIDebug As Boolean Implements ISharedContext.INI_APIDebug
        Public Property INI_UsageRestrictions As String Implements ISharedContext.INI_UsageRestrictions
        Public Property INI_Language1 As String Implements ISharedContext.INI_Language1
        Public Property INI_Language2 As String Implements ISharedContext.INI_Language2
        Public Property INI_KeepFormat1 As Boolean Implements ISharedContext.INI_KeepFormat1
        Public Property INI_KeepFormat2 As Boolean Implements ISharedContext.INI_KeepFormat2
        Public Property INI_KeepFormatCap As Integer Implements ISharedContext.INI_KeepFormatCap
        Public Property INI_KeepParaFormatInline As Boolean Implements ISharedContext.INI_KeepParaFormatInline
        Public Property INI_ReplaceText1 As Boolean Implements ISharedContext.INI_ReplaceText1
        Public Property INI_ReplaceText2 As Boolean Implements ISharedContext.INI_ReplaceText2
        Public Property INI_DoMarkupOutlook As Boolean Implements ISharedContext.INI_DoMarkupOutlook
        Public Property INI_DoMarkupWord As Boolean Implements ISharedContext.INI_DoMarkupWord
        Public Property INI_RoastMe As Boolean Implements ISharedContext.INI_RoastMe
        Public Property DecodedAPI As String Implements ISharedContext.DecodedAPI
        Public Property DecodedAPI_2 As String Implements ISharedContext.DecodedAPI_2
        Public Property TokenExpiry As DateTime Implements ISharedContext.TokenExpiry
        Public Property TokenExpiry_2 As DateTime Implements ISharedContext.TokenExpiry_2
        Public Property Codebasis As String Implements ISharedContext.Codebasis

        Public Property GPTSetupError As Boolean Implements ISharedContext.GPTSetupError
        Public Property INIloaded As Boolean Implements ISharedContext.INIloaded
        Public Property RDV As String Implements ISharedContext.RDV
        Public Property InitialConfigFailed As Boolean Implements ISharedContext.InitialConfigFailed
        Public Property INI_ContextMenu As Boolean Implements ISharedContext.INI_ContextMenu
        Public Property INI_UpdateCheckInterval As Integer Implements ISharedContext.INI_UpdateCheckInterval
        Public Property INI_UpdatePath As String Implements ISharedContext.INI_UpdatePath
        Public Property INI_SpeechModelPath As String Implements ISharedContext.INI_SpeechModelPath
        Public Property INI_TTSEndpoint As String Implements ISharedContext.INI_TTSEndpoint
        Public Property SP_Translate As String Implements ISharedContext.SP_Translate
        Public Property SP_Correct As String Implements ISharedContext.SP_Correct
        Public Property SP_Improve As String Implements ISharedContext.SP_Improve
        Public Property SP_Explain As String Implements ISharedContext.SP_Explain
        Public Property SP_SuggestTitles As String Implements ISharedContext.SP_SuggestTitles
        Public Property SP_Friendly As String Implements ISharedContext.SP_Friendly
        Public Property SP_Convincing As String Implements ISharedContext.SP_Convincing
        Public Property SP_NoFillers As String Implements ISharedContext.SP_NoFillers
        Public Property SP_Podcast As String Implements ISharedContext.SP_Podcast
        Public Property SP_Shorten As String Implements ISharedContext.SP_Shorten
        Public Property SP_Summarize As String Implements ISharedContext.SP_Summarize
        Public Property SP_MailReply As String Implements ISharedContext.SP_MailReply
        Public Property SP_MailSumup As String Implements ISharedContext.SP_MailSumup
        Public Property SP_FreestyleText As String Implements ISharedContext.SP_FreestyleText
        Public Property SP_FreestyleNoText As String Implements ISharedContext.SP_FreestyleNoText
        Public Property SP_SwitchParty As String Implements ISharedContext.SP_SwitchParty
        Public Property SP_Anonymize As String Implements ISharedContext.SP_Anonymize
        Public Property SP_ContextSearch As String Implements ISharedContext.SP_ContextSearch
        Public Property SP_ContextSearchMulti As String Implements ISharedContext.SP_ContextSearchMulti
        Public Property SP_RangeOfCells As String Implements ISharedContext.SP_RangeOfCells
        Public Property SP_WriteNeatly As String Implements ISharedContext.SP_WriteNeatly
        Public Property SP_Add_KeepFormulasIntact As String Implements ISharedContext.SP_Add_KeepFormulasIntact
        Public Property SP_Add_KeepHTMLIntact As String Implements ISharedContext.SP_Add_KeepHTMLIntact
        Public Property SP_Add_KeepInlineIntact As String Implements ISharedContext.SP_Add_KeepInlineIntact
        Public Property SP_Add_Bubbles As String Implements ISharedContext.SP_Add_Bubbles
        Public Property SP_Add_Revisions As String Implements ISharedContext.SP_Add_Revisions
        Public Property SP_MarkupRegex As String Implements ISharedContext.SP_MarkupRegex
        Public Property SP_ChatWord As String Implements ISharedContext.SP_ChatWord

        Public Property SP_Add_ChatWord_Commands As String Implements ISharedContext.SP_Add_ChatWord_Commands
        Public Property INI_ChatCap As Integer Implements ISharedContext.INI_ChatCap
        Public Property INI_ISearch As Boolean Implements ISharedContext.INI_ISearch
        Public Property INI_ISearch_Approve As Boolean Implements ISharedContext.INI_ISearch_Approve
        Public Property INI_ISearch_URL As String Implements ISharedContext.INI_ISearch_URL
        Public Property INI_ISearch_ResponseURLStart As String Implements ISharedContext.INI_ISearch_ResponseURLStart
        Public Property INI_ISearch_ResponseMask1 As String Implements ISharedContext.INI_ISearch_ResponseMask1
        Public Property INI_ISearch_ResponseMask2 As String Implements ISharedContext.INI_ISearch_ResponseMask2
        Public Property INI_ISearch_Name As String Implements ISharedContext.INI_ISearch_Name
        Public Property INI_ISearch_Tries As Integer Implements ISharedContext.INI_ISearch_Tries
        Public Property INI_ISearch_Results As Integer Implements ISharedContext.INI_ISearch_Results
        Public Property INI_ISearch_MaxDepth As Integer Implements ISharedContext.INI_ISearch_MaxDepth
        Public Property INI_ISearch_Timeout As Long Implements ISharedContext.INI_ISearch_Timeout
        Public Property INI_ISearch_SearchTerm_SP As String Implements ISharedContext.INI_ISearch_SearchTerm_SP
        Public Property INI_Placeholder_03 As String Implements ISharedContext.INI_Placeholder_03
        Public Property INI_ISearch_Apply_SP As String Implements ISharedContext.INI_ISearch_Apply_SP
        Public Property INI_ISearch_Apply_SP_Markup As String Implements ISharedContext.INI_ISearch_Apply_SP_Markup
        Public Property INI_Lib As Boolean Implements ISharedContext.INI_Lib
        Public Property INI_Lib_File As String Implements ISharedContext.INI_Lib_File
        Public Property INI_Lib_Timeout As Long Implements ISharedContext.INI_Lib_Timeout
        Public Property INI_Lib_Find_SP As String Implements ISharedContext.INI_Lib_Find_SP
        Public Property INI_Lib_Apply_SP_Markup As String Implements ISharedContext.INI_Lib_Apply_SP_Markup
        Public Property INI_Lib_Apply_SP As String Implements ISharedContext.INI_Lib_Apply_SP
        Public Property INI_Placeholder_01 As String Implements ISharedContext.INI_Placeholder_01
        Public Property INI_Placeholder_02 As String Implements ISharedContext.INI_Placeholder_02
        Public Property INI_MarkupMethodHelper As Integer Implements ISharedContext.INI_MarkupMethodHelper
        Public Property INI_MarkupMethodWord As Integer Implements ISharedContext.INI_MarkupMethodWord
        Public Property INI_ShortcutsWordExcel As String Implements ISharedContext.INI_ShortcutsWordExcel
        Public Property INI_PromptLib As Boolean Implements ISharedContext.INI_PromptLib
        Public Property INI_PromptLibPath As String Implements ISharedContext.INI_PromptLibPath
        Public Property INI_PromptLibPath_Transcript As String Implements ISharedContext.INI_PromptLibPath_Transcript
        Public Property PromptLibrary() As List(Of String) Implements ISharedContext.PromptLibrary
        Public Property PromptTitles() As List(Of String) Implements ISharedContext.PromptTitles
        Public Property MenusAdded As Boolean Implements ISharedContext.MenusAdded



#End Region

    End Class

    Public Class InputParameter
        Public Property Name As String
        Public Property Value As Object
        ' We use this property to keep track of the dynamically created control.
        Public Property InputControl As Control

        Public Sub New(ByVal name As String, ByVal value As Object)
            Me.Name = name
            Me.Value = value
        End Sub
    End Class

    Module WinAPI
        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Public Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
        End Function
    End Module

    Public Module ProgressBarModule
        ' Global variables to control the progress form.
        Public GlobalProgressValue As Integer = 0
        Public GlobalProgressMax As Integer = 100
        Public GlobalProgressLabel As String = "Initializing..."
        Public CancelOperation As Boolean = False

        ' Call this procedure to launch the progress form.
        Public Sub ShowProgressBarInSeparateThread(headerText As String, initialLabel As String)
            Dim t As New Thread(Sub()
                                    ' Create and show the progress form modally.
                                    Dim progressForm As New ProgressForm(headerText, initialLabel)
                                    progressForm.ShowDialog()
                                End Sub)
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        End Sub
    End Module

    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        Private _hwnd As IntPtr

        Public Sub New(handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property
    End Class

    Public Class SharedMethods

        ' Amend the following two values to hard code the encryption key and permitted domains (otherwise the values are taken from the registry at the path below)

        Private Const Int_CodeBasis As String = ""
        Public Const alloweddomains As String = ""

        Public Const AN As String = "Red Ink" ' 
        Public Const AN2 As String = "redink" ' 
        Public Const AN3 As String = "Red Ink" ' Name used for Visual Studio Project 
        Public Const AN4 As String = "https://vischer.com/redink"  ' Name of sub-directory on Website of vischer.com/...  
        Public Const MaxUseDate As Date = #9/30/2025#

        Private Const ISearch_MaxTries = 30          ' maximum number of search hits to be tried
        Private Const ISearch_MaxMaxDepth = 10       ' maximum number of search levels to crawl a website
        Private Const ISearch_MaxResults = 15        ' maximum number of search results
        Private Const ISearch_MaxSearchTimeout = 60  ' maximum seconds per search hit
        Private Const ISearch_DefTries = 10          ' default maximum number of search hits to be tried
        Private Const ISearch_DefResults = 4         ' default maximum number of search results
        Private Const ISearch_DefMaxDepth = 2        ' maximum number of search levels to crawl a website
        Private Const ISearch_DefSearchTimeout = 3   ' default maximum seconds per search hit

        Public Const RegPath_Base As String = "HKEY_CURRENT_USER\Software\" & AN3 & "\"
        Public Const RegPath_CodeBasis As String = "CodeBasis"
        Public Const RegPath_IniPath As String = "IniPath"
        Public Const RegPath_IniPrio As Boolean = False ' True if the registry path shall have priority over the default path

        Private Const RegexSeparator1 As String = "|||"  ' Set also in Word Addin
        Private Const RegexSeparator2 As String = "§§§"  ' Set also in Word Addin

        Public Shared RemoveMenu As Boolean = False

        Public Shared LicenseText As String =
            $"{AN} for Word, Excel And Outlook has been created In VB.net (handlers: VBA) and uses:" & vbCrLf & vbCrLf &
            "1. DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Appache-2.0 license " &
            "(http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex)." & vbCrLf &
            "2. Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license " &
            "(https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json" & vbCrLf &
            "3. HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier, Jeff Klawiter, " &
            "Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/" & vbCrLf &
            "4. Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; " &
            "licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/" & vbCrLf &
            "5. PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license " &
            "(https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig" & vbCrLf &
            "6. MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license " &
            "(https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig" & vbCrLf &
            "7. NAudio in unchanged form; Copyright (c) 2020 Mark Heath; licensed under a proprietary open source license (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio" &
            "(https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig" & vbCrLf &
            "8. Vosk in unchanged form; Copyright (c) 2022 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/" & vbCrLf &
            "9. Whisper.net In unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net" & vbCrLf &
            "10. Various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; " &
            "Copyright (c) 2016- Microsoft Corp." & vbCrLf & vbCrLf & "Disclaimer:" & vbCrLf & vbCrLf &
            "THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS 'As Is' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE." & vbCrLf & vbCrLf &
            "See the Red Ink license file (https://apps.vischer.com/redink/license.txt) for more information."

        Public Shared DefaultINIPaths As New Dictionary(Of String, String) From {
            {"Word", "%AppData%\Microsoft\Word\" & AN2 & ".ini"},
            {"Excel", "%AppData%\Microsoft\Excel\" & AN2 & ".ini"},
            {"Outlook", "%AppData%\Microsoft\Outlook\" & AN2 & ".ini"}
        }

        Public Shared HelperPaths As New Dictionary(Of String, String) From {
            {"Word", "%AppData%\Microsoft\Word\STARTUP\" & AN2 & "_helper.dotm"},
            {"Excel", "%AppData%\Microsoft\Excel\XLSTART\" & AN2 & "_helper.xlam"}
        }

        Public Shared ExcelHelper As String = AN2 & "_helper.xlam"
        Public Shared WordHelper As String = AN2 & "_helper.dotm"

        Public Shared ExcelHelperUrl As String = "https://apps.vischer.com/redink/" & ExcelHelper
        Public Shared WordHelperUrl As String = "https://apps.vischer.com/redink/" & WordHelper

        Const Default_SP_Translate As String = "You are a translator that precisely complies with its instructions step by step. Translate in to {TranslateLanguage} the text that is provided to you and is marked as 'Texttoprocess'. When you translate, do not add any other comments and the translation should be of about the same length. Whenever there is a line feed or carriage return in text provided to you, it is essential that you also include such line feed or carriage return in the output you generate. The carriage returns and line feeds in the output must match exactly those in the original text provided to you. Accordingly, if there are two carriage returns or line feeds in succession in the text provided to you, there must also be two carriage returns or line feeds in the text you generate. Remove any double spaces that follow punctuation marks. Before translating, check whether the text is drafted in a formal or informal manner, and maintain such style. If and when asked to translate to a language where the translation of 'you' is translated differently depending on whether it is formal or not, such as German or French, go by default for a formal translation (e.g., 'Sie' or 'vous'), unless the text is clearly very informal, for example, because the text is addressed to a person by their first name or signed only with the first name of a person. {INI_PreCorrection}"
        Const Default_SP_Correct As String = "You are a legal professional with very good language skills that precisely complies with its instructions step by step. Amend the text that is provided to you, in its original language, and is marked as 'Texttoprocess' to only correct spelling, missing words, clearly unnecessary words, strange or archaic language and poor style. When doing so, do not significantly change the length of the text. Whenever there is a line feed or carriage return in text provided to you, it is essential that you also include such line feed or carriage return in the output you generate. The carriage returns and line feeds in the output must match exactly those in the original text provided to you. Accordingly, if there are two carriage returns or line feeds in succession in the text provided to you, there must also be two carriage returns or line feeds in the text you generate. {INI_PreCorrection}"
        Const Default_SP_Improve As String = "You are a legal professional with very good language skills that precisely complies with its instructions step by step. Amend the text that is provided to you, in its original language, and is marked as 'Texttoprocess' to be much more concise, to the point, better structured and easier to understand and in better, professional style. Change passive voice to active voice, where this makes sense. Remove rendundancies and filler words, except where this is necessary for easy reading and style. When doing so, do not significantly change the length of the text. Also, do not change the overall meaning, tone or content of the text. {INI_PreCorrection}"
        Const Default_SP_Shorten As String = "You are a legal professional and editor with excellent language, logical and rhetorical skills that precisely complies with its instructions step by step. Shorten the text that is provided to you, in its original language,  and is marked as 'Texttoprocess'. Shorten it as much as necessary to ensure that the output generated by you has {ShortenLength} words. In a first step try to remove redundancies, and if this is not sufficient to fulfill the instruction, then remove less important information or combine information. However, preserve the original tone, the original message of the texttoprocess (but not the <texttoprocesstag>) and any material information. {INI_PreCorrection}"
        Const Default_SP_Summarize As String = "You are a legal professional with excellent language, logical and rhetorical skills that precisely complies with its instructions step by step. Create a very short summary of the that is provided to you, in its original language, and is marked as 'Texttoprocess'. Ensure that your output has {SummaryLength} words. Use the same language style as in the original text, but do not add any information or other thoughts to it. {INI_PreCorrection}"
        Const Default_SP_FreestyleText As String = "You are a legal professional with excellent language, logical and rhetorical skills that precisely complies with its instructions step by step. Perform the instruction '{OtherPrompt}' using the language of the command and the text provided to you and marked as 'texttoprocess'. {INI_PreCorrection} However, do not include the text of your instruction in your output."
        Const Default_SP_FreestyleNoText As String = "You are a legal professional with excellent language, logical and rhetorical skills that precisely complies with its instructions step by step. Perform the instruction '{OtherPrompt}' using the language of the command. {INI_PreCorrection} However, do not include the text of your instruction in your output."
        Const Default_SP_MailReply As String = "You are a legal professional with excellent legal, language, logical and rhetorical skills that precisely complies with its instructions step by step. Your task is to read the text that is provided to you and marked as 'mailchain', which contains an e-mail chain. The first mail you get is the e-mail to which you shall draft a response for me. When drafting the response for me, comply with the following instructions and information (= key instructions): {OtherPrompt}. \n\nThese are the further rules that every answer should follow: 1. Draft it in the same language as the first mail you get has been written (do not consider headers, the subject line or the footer. 2. The top (and latest) e-mail you are provided with in the mailchain is from the person who wrote to me. This will be the person to whom I want to respond to. You will draft an e-mail to respond to that person, i.e. the author of the top and latest e-mail. Please keep that in mind when drafting a response and make sure that you . 3. Please read the entire mail chain and distinguish exactly who has written what and what the party, to whom I will respond, has written when drafting the response. However, on the substance, focus on the key instructions I provided to you above, if any. 4. In your proposed response use the same style, type of language and way of e-mail drafting as I do. 5. Do not process and never consider or include signatures and mail footers. 6. Provide your output in the Markdown format. 7. When drafting a reply, use full salutations and closing formulas that are adequate in view of the tone of the mailchain. 8. Finally, when drafting the response, it is very important that you comply with all instructions and careful check your response for compliance with all instructions before you provide it. {INI_PreCorrection}"
        Const Default_SP_MailSumup As String = "You are a highly skilled legal professional who strictly follows instructions step by step; analyze the body of the provided ""mailchain"" to determine its predominant language (ignoring sender, recipient, subject, etc.), strictly use this language for the output, generate a concise, structured Markdown-formatted summary (in bold, but not header formatting) including a one-sentence key takeaway followed by a breakdown of key points distinguishing different authors, ensuring the summary is very short and concise while retaining all critical information and getting an understanding of the conversation. {INI_PreCorrection}"
        Const Default_SP_SwitchParty As String = "You are a legal professional And editor with excellent language, logical And rhetorical skills that precisely complies with its instructions step by step. Your task is to swap parties in a text and adapt the text to still read correctly. To do so, rewrite the text that is provided to you and is marked as 'TEXTTOPROCESS' as if '{OldParty}' were '{NewParty}' preserving all other information, but ensure that in particular all pronouns, titles, possessive forms and the use of plural and singular are appropriately adjusted. If {OldParty} or {NewParty} is not a name, treat it based on its meaning, even if it starts with a capital letter. \n {INI_PreCorrection}"
        Const Default_SP_Anonymize As String = "You are very careful editor And legal professional that precisely complies with its instructions step by step. Fully anonymize the text that Is provided to you And Is marked as 'TEXTTOPROCESS'. Do so only by replacing any names, companies, businesses, parties, organizations, proprietary product names, unknown abbreviations, personal addresses, e-mail accounts, phone numbers, IDs, credit card information, account numbers and other identifying information by the expression '[redacted]' and before providing the result, check whether there is no information left that could directly or indirectly identify any person, company, business, party or organization, including information that could link to them by doing an Internet search, and if so, redact it as well. {INI_PreCorrection}"
        Const Default_SP_RangeOfCells As String = "You are an expert In analyzing And explaining Excel files To non-experts And In drafting Excel formulas For use within Excel. You precisely comply With your instructions. Perform the instruction '{OtherPrompt}' using the range of cells provided You between the tags <RANGEOFCELLS> ... </RANGEOFCELLS>. When providing your advice, follow this exact format for each suggestion: \n 1. Use the delimiter ""[Cell: X]"" for each cell reference (e.g., [Cell: A1]). 2. For formulas, use '[Formula: =expression]' (e.g., [Formula: =SUM(A1:A10)]). 3. For values, use ""[Value: 'text']"" (e.g., [Value: 'New value']). 4. Each instruction should start with the ""[Cell: X]"" marker followed by a [Formula: ...] or [Value: ...] in the next line. 5. Ensure that each instruction is on a new line. 6. If a formula or value is not required for a cell, leave that part out or indicate it as empty. {INI_PreCorrection}"
        Const Default_SP_WriteNeatly As String = "You are a legal professional with very good language skills that precisely complies with its instructions step by step. Amend the text that is provided to you, in its original language, and is marked as 'Texttoprocess' to be a coherent, concise and easy to understand text based the text and keywords in the provided text, without changing or adding any meaning or information to it, but taking into account the following context, if any: '{Context}' {INI_PreCorrection}"
        Const Default_SP_Add_KeepFormulasIntact As String = "Beware, the text contains an Excel formula. Unless expressly instructed otherwise, make sure that the formula still works as intended."
        Const Default_SP_Add_KeepHTMLIntact As String = "When completing your task, leave any HTML tags within 'TEXTTOPROCESS' fully intact in the output."
        Const Default_SP_Add_KeepInlineIntact As String = "Do not remove any text that appears between {{ and }}; these placeholders contain content that is part of the text."
        Const Default_SP_Add_Bubbles As String = "Provide your response to the instruction not in a single, combined text, but split up your response in portions so that each portion relates to one particular portion of the texttoprocess. When doing so, follow strictly these rules: \n1. For each such portion of the the texttoprocess, provide your response in the the form of a comment to the portion of the text to which it relates. \n2. Provide each portion of your response by first quoting verbatim the relevant portion of the texttoprocess followed by the relevant comment for that portion of the texttorpocess. When doing so, follow strictly this syntax: ""text1@@comment1§§§text2@@comment2§§§text3@@comment3"". It is important that you provide your output exactly in this form: First provide the quoted text, then the separator @@ and then your comment. After that, add the separator §§§ and continue with the second portion and comment in the same way, and so on. Make sure to use these separators exactly as instructed. If you do not comply, your answer will be invalid. \n3. Make sure you quote the portion of the Texttoprocess exactly as it has been provided to you; do not change anything to the quoted portion of the Texttoprocess, do not add or remove any characters, do not add quotation marks.\n4. Keep the quoted text as short as possible (ensuring that it is still unique in the texttoprocess) and that the comment for such portion is drafted meaningful. \n5. Limit your output to those sections of the texttoprocess where you actually do have something meaningful to say. Unless expressly instructed otherwise, you are not allowed to refer to sections of the texttoprocess for which you have no comment or remark. For example, 'No comment' or the like is a bad, invalid response. If there is a paragraph or section for which you have no meaningfull comment, skip it in your output. \n6. Follow these rules strictly, because your output will otherwise not be valid."
        Const Default_SP_Add_Revisions As String = "Where the instructions refer to markups, changes, insertions, deletions or revisions in the text, they are found within the tags <ins>...</ins> for insertions and within the tags <del> ... </del> for deletions."
        Public Shared Default_SP_MarkupRegex As String = $"You are an expert text comparison system and want you to give the instructions necessary to change an original text using search & replace commands to match the new text. I will below provide two blocks of text: one labeled <ORIGINALTEXT> ... </ORIGINALTEXT> and one labeled <NEWTEXT> ... </NEWTEXT>. With the two texts, do the following: \n1. You must identify every difference between them, including punctuation changes, word replacements, insertions, or deletions. Be very exact. You must find every tiny bit that is different. \n2. Develop a profound strategy on how and in which sequence to most efficiently and exactly apply these replacements, insertions and deletions to the old text using a search-and-replace function. This means you can search for certain text and all occurrences of such text will be replaced with the text string you provide. If the text string is empty (''), then the occurrences of the text will be deleted. When developing the strategy, you must consider the following: (a) Every occurrence of the search text will be replaced, not just the first one. This means that if you wish to change only one occurrence, you have to provide more context (i.e. more words) so that the search term will only find the one occurrence you are aiming at. (b) If there are several identical words or sentences that need to be change in the same manner, you can combine them, but only do so, if there are no further changes that involve these sections of the text. (c) Consider that if you run a search, it will also apply to text you have already changed earlier. This can result in problems, so you need to avoid this. (d) Consider that if you replace certain words, this may also trigger changes that are not wanted. For example, if in the sentence 'Their color is blue and the sun is shining on his neck.' you wish to change the first appearance of 'is' to 'are', you may not use the search term 'is' because it will also find the second appearance of 'is' and it will find 'his'. Instead, you will have to search for 'is blue' and replace it with 'are blue'. Hence, alway provide sufficient context where this is necessary to avoid unwanted changes. (e) You should avoid searching and replacing for the same text multiple times, as this will result in multiplication of words. If all occurrences of one term needs to be replaced with another term, you need to provide this only once. (f) Pay close attention to upper and lower case letters, as well as punctuation marks and spaces. The search and replace function is sensitive to that. (g) When building search terms, keep in mind that the system only matches whole word; wildcards and special characters are not supported. \n3. Implement the strategy by producing a list of search terms and replacement texts (or empty strings for deletions). Your list must be strictly in this format, with no additional commentary or line breaks beyond the separators: SearchTerm1{RegexSeparator1}ReplacementforSearchTerm1{RegexSeparator2}SearchTerm2{RegexSeparator1}ReplacementforSearchTerm2{RegexSeparator2}SearchTerm3{RegexSeparator1}ReplacementforSearchTerm3... For example, if SearchTerm3 indicates a text to be deleted, the ReplacementforSearchTerm3 would be empty. - Use '{RegexSeparator1}' to separate the search term from its replacement. - Use '{RegexSeparator2}' to separate one find/replace pair from the next. - Do not include numeric placeholders (like 'Search Term 1') or any extraneous text. When generating the search and replacement terms, it is mandatory that you include the search and replacement terms exactly as they exist in the underlying text. Never change, correct or modify it. You must strictly comply with this. Otherwise your output will be unusable and invalid. \nNow, here are the texts:"
        'Public Shared Default_SP_MarkupRegex As String = $"You are an expert text comparison system. I will provide two blocks of text: one labeled <ORIGINALTEXT> ... </ORIGINALTEXT> and one labeled <NEWTEXT> ... </NEWTEXT>. You must identify every difference between them, including punctuation changes, word replacements, insertions, or deletions. Then, for each distinct difference, produce: 1. A unique Regex pattern that matches ONLY that specific changed string in the original text (no placeholder text like 'Regex Pattern 1'; provide the actual pattern). 2. The replacement text exactly as it appears in the new text. Ensure the Regex patterns do not match identical text in other parts of the document. Your output must be strictly in this format, with no additional commentary or line breaks beyond the separators: RegexThatMatchesChange1{RegexSeparator1}ReplacementForChange1{RegexSeparator2}RegexThatMatchesChange2{RegexSeparator1}ReplacementForChange2{RegexSeparator2}RegexThatMatchesChange3{RegexSeparator1}ReplacementForChange3... - Use '{RegexSeparator1}' to separate the Regex from its replacement. - Use '{RegexSeparator2}' to separate one find/replace pair from the next. - Do not include numeric placeholders (like 'Regex Pattern 1') or any extraneous text. Now, here are the texts:"
        Const Default_SP_ChatWord As String = "You are a helpful AI, you are running inside Microsoft Word, and may be shown with content from the document that the user has opened currently (you will be told later in this prompt). When responding to the user, do so in the language of the question, unless the user instructs you otherwise. Before generating any output, keep in mind the following:\n\n 1. You have a legal professional background, are very intelligent, creative and precise. You have a good feeling for adequate wording and how to express ideas, and you have a lot of ideas on how to achieve things. You are easy going. \n\n 2. You exist within the application Microsoft Word. If the user allows you to interact with his document, then you can do so and you will automatically get additional instructions how to do so. \n\n 3. You always remain polite, but you adapt to the communications style of the user, and try to provide the type of help the user expresses. If the user gives commands, execute the commands without big discussion, except if something is not clear. If the user wants you to analyse his text, do so, be a concise, critical, eloquent, wise and to the point discussion partner and, if the user wants, go into details. If the user's input seems uncoordinated, too generic or really unclear, ask back and offer the kind of help you can really give, and try to find out what the user wants so you can help. If it despite several tries is not clear what the users wants, you might offer him certain help, but be not too fortcoming with offering ideas what you can do. In any event, follow the KISS principle: Unless it is necessary to complete a task, keep it always short and simple. \n\n 4. Your task is to help the user with his text. You may be asked to do this to answer some general questions to help the user brainstorm, draft his text, sort his ideas etc., or you may be asked to do specific stuff with his text. \n\n 5. If you are given access to the user's text (which is upon the user to decide using two checkboxes), you will be presented to it further below as 'content'. \n\n 6. You will also be given the name of the document that contains the 'content'. This is important because you may have to deal with several different documents, and can distinguish them based on their names. Try to do so and remember them. \n\n. 7. If you need to remember something, make sure you provide it as part of your output. You can only remember things that are contained in your output or the output of the user. Accordingly, if the user asks you to remember something from a particular content (i.e. other than what the user tells you or you have provided as an output), then repeat it, and if necessary with the name of the document, if it is meaningful. \n\n 8. Do not remove or add carriage returns or line feeds from a text unless this is necessary for fulfilling your task. Also, do not use double spaces following punctuation marks (double spaces following punctuation marks are only permitted if included in the original text). \n\n 9. The user can decide by clicking a checkbox 'Grant write access' whether he gives you the ability to change his content, search within the content or insert new text. If further below you are informed of the commands (e.g., [#INSERT ...#]) to do so, you know that he has done so and you may provide him assistance in explaining what you can do, if you believe he should know. \n\n 10. Be precise and follow instructions exactly. Otherwise your answers may be invalid."
        Const Default_SP_Add_ChatWord_Commands As String = "To help the user, you can now directly interact with the document or selection content provided to you (this comes from the user). Unless stated otherwise, this is the text of the user to which the user will when asking you to do things with his document, such as finding, replacing, deleting or inserting text you generate, or making changes to the text or implementing the suggestions you have made. Try to help the user to improve his content or answer questions concerning it. You are now authorized to do so if this is required to fulfill a request of the user. Proactively offer the user this possibility, if this helps to solve the user's issues. But never ask whether you should find, replace, delete or insert text if you actually do issue such as a command. Beware: You either ask whether you should issue a command to find, replace, delete or insert text, or ask so, but never both. If you are unsure, ask before doing something. \n\nYou can fulfill the users instructions by including commands in your output that will let the system search, modify and delete such content as per your instructions.\n\nTo do so, you must follow these instructions exactly: 1. You can optionally insert one or more of these commands for Word: - [#FIND: @@searchterm@@#] for finding, highlighting, marking or showing text to the user. The searchterm must be enclosed in @@ without quotes or other punctuation. - [#REPLACE: @@searchterm@@ §§newtext§§#] for search-and-replace. The searchterm must be in @@, the replacement text in §§, both without quotes. 2. If there are multiple occurrences of the search term in the document, you must provide additional context in the search term to uniquely identify the correct occurrence. Context may include a nearby phrase, word, or sentence fragment. Consider the entire text and other possible matches of what you wish to find and replace in order to find, replace or even delete content that you were not intending. 3. Ensure that the replacement term preserves necessary context to avoid accidental changes or deletions to other text. For example, if replacing only the second occurrence of ""example"" in ""This is an example. Another example follows."", the instruction could be [#REPLACE: @@Another example@@ §§Another sample@@#]. 4. If you provide multiple replacement commands, you must consider the changes already made by earlier commands when drafting later ones. For example, if the first command replaces ""example"" with ""sample"" and the second occurrence of ""example"" is in the same text, the search term for the second replacement must reflect the updated text. 5. You also have a command [#INSERTAFTER: @@searchtext@@ §§newtext§§#], which appends new text (newtext) immediately after searchtext. Use this if the user wants to add or expand text in the document. Your search term will be the text immediately preceeding the point where you want to insert the text for achieving your goal. If, HOWEVER, you are asked or required to insert newtext immediately before the text of the search term, then use the command [#INSERTBEFORE: @@searchtext@@ §§newtext§§#]. Inserting 'before' works as inserting 'after', with the exception that the newtext will be inserted before the text found and not after. 6. If your task is to insert a particular text in the user's empty document or with no instruction as to the location of the new text, use the command [#INSERT: @@newtext@@#] instead of INSERTBEFORE or INSERTAFTER. In this case, 'newtext' is the text you are asked to insert into the user's content (not the text you provide as your response. Never include what you wish to tell the user into newtext. The INSERT command is reserved exclusively for inserting text into the user's content. 7. If you want to delete text, do so by executing a [#REPLACE: @@searchtext@@ §§§§#] command, leaving the replacement text empty. 8. If content to be searched for contains carriage returns (often shown as '\r') or line feeds (often shown as '\n'), make sure your search term also contains the \r and \n in the same place. If you do not include the carriage returns ('\r') and line feed characters ('\n') in your search terms, your command will not work and your response is invalid. 9. Before issuing any commands, think carefully about the order of the commands you issue. They will be executed in the order you produce them. Build a logical sequence to avoid following commands affecting the outcome of preceeding commands. Keep in mind that replaced or deleted text will remain visible to the system. For example, if you replace 'whirlpool' with 'table' and issue second command to replace 'pool' with 'chair', it will also find all occurences of 'whirlpool', even despite your previous command of replacing 'whirlpool'. To solve such issues, only issue commands that are certainly not conflicting. Then explain to the user what other changes you wish to do, but ask the user to first accept the changes if the user agrees, and wait for approval to continue issuing your commands. 10. No other commands are allowed. Keep in mind that you cannot change and formatting or deal with it; if you are asked to do things you can't do, tell the user so. 11. In your visible answer to the user, never show these commands in the same line. Provide any commands only after your user-facing text, each on its own line. 12. If you do not need to find, replace, delete or insert text, do not produce a command. If you are unsure what to do, ask the user and interact. You can also make proposals explaining what you want to do and ask the user if this is what the user wants. If the user gives you a direct instruction, however, you can comply. 13. Use the exact syntax for the commands. If you deviate in any way (e.g. quotes, extra spaces, or missing delimiters), the response is invalid. 14. If you provide searchterms in your commands, be very precise. If you do not exactly quote the text as it is contained in the content, your command will not be executed. 15. The user does not see these commands, so do not repeat them in your text. Do not include them in the middle of your output. Always place them on separate lines at the end of your output. 16. Never repeat the text of your output in the commands and vice versa. However, if you issue commands, provide the user a summary of what you have done with his document and ask him to check. 17. If you include commands in your output, do not ask the user whether you shall implement the changes you suggest. Only ask the user whether you shall implement a change in the document if you have not already done so; keep in mind that any command you include will usually be executed when you provider your answer (unless something goes wrong, which is always possible, which is why every command should be checked). Asking the user whether you may issue commands if you already issue them is contradictory. If you are not sure, ask the user and issue commands only once the user has approved so. 18. Keep your response to the user and the commands for finding, replacing, inserting and deleting text completely separate.\n\n\nNow here are some examples: - Good example if the user wants to find, highlight or show to the user ""example"" with context: Text to user: ""I located the correct ""example"" in the sentence ""This is an example.""."" Then on a new line: [#FIND: @@This is an example@@#]. - Good example for replacing the second occurrence of ""example"": Text to user: ""I recommend replacing the second occurrence of ""example"" in ""This is an example. Another example follows.""."" Then on a new line: [#REPLACE: @@Another example@@ §§Another sample§§#]. - Good example for sequential replacements: Text to user: ""I suggest replacing ""example"" step by step: First, replace ""example"" in ""This is an example."" with ""sample."" Then, replace ""Another example follows."" with ""Another sample follows.""."" On separate lines: [#REPLACE: @@This is an example@@ §§This is a sample§§#] [#REPLACE: @@Another example follows@@ §§Another sample follows§§#]. - Good example for insertion: Text to user: ""I suggest adding a summary after the phrase ""Introduction:""."" Then on a new line: [#INSERTAFTER: @@Introduction:@@ §§Here is a short summary.§§#]. - If you have to delete a text containing carriage returns such as ""This is line1.\rThis is line 2.\r\r"", a good example is: [#REPLACE: @@This is line 1.\rThis is line 2.\r\r@@ §§§§#] \n\n--- A bad and invalid response is: [#REPLACE: @@This is line 1.This is line 2.@@ §§§§#] (because the search term in your command is missing the three carriage returns that are contained in the user content - the search term will not work without the three carriage returns; always include the same carriage returns and line feeds from the original content in your command search terms). --- Another bad and invalid response: [#REPLACE: @@example@@ §§sample@@#] (because it ends with a '@@' instead of a '§§', which is a mistake; you may never use an '@@' at the end of a command that replaces or inserts text). \n\nYou must follow these instructions strictly."
        Const Default_INI_ISearch_SearchTerm_SP As String = "You are an advanced language model tasked with generating precise and direct search terms required to fulfill the given instruction. Analyze the instruction and any additional text provided within <TEXTTOPROCESS> and </TEXTTOPROCESS> tags, if present, to output only the specific search terms needed to retrieve the required information. If no additional text is provided, base your search terms solely on the instruction. The search terms should be formatted as they would appear in a search engine query, without any additional explanations or context. Instruction: {OtherPrompt}, Current Date: {CurrentDate}. Provide only the search terms, formatted for direct input into a search engine. Avoid any additional text or explanations."
        Const Default_INI_ISearch_Apply_SP As String = "You are a legal professional with excellent legal, language and logical skills and you precisely comply with your instructions step by step. You will execute the following instruction in the language of the command using (1) the knowledge and Information contained in the internet search results provided within the <SEARCHRESULT1> … </SEARCHRESULT1>, <SEARCHRESULT2> … </SEARCHRESULT2> etc. tags, and (2) the text provided within the <TEXTTOPROCESS> and </TEXTTOPROCESS> tags, if present. {INI_PreCorrection} \n Instruction: '{OtherPrompt}'\n {SearchResult} \n"
        Const Default_INI_ISearch_Apply_SP_Markup As String = "You are a legal professional With excellent legal, Language And logical skills And you precisely comply With your instructions Step by Step. You will execute the following instruction In the language Of the command Using the knowledge And Information contained In the internet search results provided within the <SEARCHRESULT1> … </SEARCHRESULT1>, <SEARCHRESULT2> … </SEARCHRESULT2> etc. tags, And applying it directly To text provided within the <TEXTTOPROCESS> And </TEXTTOPROCESS> tags (amending it, as per the instruction). {INI_PreCorrection} \n Instruction: '{OtherPrompt} \n {SearchResult} \n"

        Const Default_SP_ContextSearch As String = "You are a very careful editor and legal professional that precisely complies with its instructions step by step. Your task is to help the user find within a text a particular sentence, section, or word based on contextual information. To do so, follow precisely these instructions:\n\n1.Study the Search Context\nYou will be provided with a Search Context (between {SearchContext}) that describes what the user is looking for. Understand the bigger picture: \n(i) What does the context refer to or mean? \n(ii) What synonyms, related terms, or references might appear in that subject matter? \n(iii) How could it be expressed with variations in phrasing? \n\n2. Read the Text\nYou will be provided with a text to search (between the tags <TEXTTOSEARCH> and </TEXTTOSEARCH>). Read it thoroughly and keep in mind all synonyms, related terms, or references identified in step 1.\n\n3. Find the First Relevant Portion\n Go through the text and locate the first portion (word, part of a sentence, entire sentence, paragraph, or multiple paragraphs) that matches or closely relates to the Search Context—either directly by wording or indirectly by meaning or context or consequences.\n\n4.Provide a Distinguishing Snippet\nWhen you find a match, extract the relevant snippet verbatim from the text. Include enough text before and/or after the main hit to ensure that the snippet is distinct from any earlier identical occurrences.\n Example: If the text is ‘There is an example, and yet another example.’ and only the second ‘example’ matches, your snippet should be ‘another example’ (verbatim, without quotes) so it cannot be confused with the first occurrence.\n\n5. Preserve Text Exactly\n Output the matched snippet exactly as it appears in the original text—no additions, no omissions, no extra punctuation, spacing, or formatting.\n\n6. Output the Snippet Only Provide nothing else in your output: no commentary, headings, explanation, quotation marks, additional carriage returns, or linefeeds. \n\n7. Avoid Invalid Output\nAny deviation from these instructions renders your output invalid. You must comply precisely.\n\n Now here is the Search Context: {SearchContext}"
        Const Default_SP_ContextSearchMulti As String = "You are a very careful editor and legal professional that precisely complies with its instructions step by step. Your task is to help the user find within a text all words, sentences, or sections that match particular contextual information. To do so, follow these instructions precisely:\n\n1. Study the Search Context\nYou will be provided with a Search Context (between {SearchContext}) that describes what the user is looking for. Understand the bigger picture:\n(i) What does the context refer to or mean?\n(ii) What synonyms, related terms, or references might appear in that subject matter?\n(iii) How could it be expressed with variations in phrasing?\n\n2. Read the Text\nYou will be provided with a text to search (between the tags <TEXTTOSEARCH> and </TEXTTOSEARCH>). Read it thoroughly and keep in mind all synonyms, related terms, or indirect references identified in step 1.\n\n3. Find All Relevant Portions\nGo through the text and locate every portion (word, part of a sentence, entire sentence, paragraph, or multiple paragraphs) that matches or relates to the Search Context—either directly by wording or indirectly by meaning or context or consequences. There might be multiple hits.\n\n4. Output Each Match Separately\nFor each match you find:\n(a) Extract the relevant snippet verbatim from the text.\n(b) Include enough text before and/or after it to ensure the snippet is distinct from any earlier identical occurrences in the text.\n(c) Separate each snippet from the next one with @@@.\n(d) Example: If the text is ‘There is an example, and yet another example.’ and only the second ‘example’ matches, output ‘another example’, making sure it cannot be confused with the first occurrence.\n\n5. Preserve Text Exactly\nOutput each matched snippet exactly as it appears in the original text—no additions, no omissions, no extra punctuation, spacing, or formatting.\n\n6. Output the Snippets Only\nProvide nothing else in your output: no commentary, headings, explanation, quotation marks, additional carriage returns, or linefeeds.\n\n7. Include All Matches\nContinue finding and listing all matches until none remain. Example format with three matches:\n Matchtext1@@@Matchtext2@@@Matchtext3\n\n8. Avoid Invalid Output\nAny deviation from these instructions renders your output invalid. You must comply precisely.\n\nNow here is the Search Context: {SearchContext}"

        Const Default_SP_Podcast As String = "You are professional podcaster and very experience script author. Create a lively and engaging text deep dive dialogue with a host and a guest based on the text you will be provided below between the tags <TEXTTOPROCESS> and </TEXTTOPROCESS>. You shall create an engaging deep dive discussion about the text that is exciting, entertaining and educational to listen to. Always keep this in mind. \n\n When creating the dialogue, it is important that you strictly follow these rules: \n\n1. The dialogue must be in **{Language}**. \n\n2. If any words or sentences appear that are not in {Language}, use SSML '<lang>' tags to ensure correct pronunciation. \n\n3. The dialogue should be a **natural, fast-paced** exchange between the charismatic host {HostName} and the insightful guest {GuestName}, avoiding exaggerated speech or unnecessary dramatization. \n\n4. Cover all key points in the text **in a natural flow**—do not sound robotic or overly formal. Summarize only if necessary, while keeping all critical information. \n\n5. Keep the tone **conversational and engaging**, similar to a professional yet relaxed podcast. Do not overuse enthusiasm—keep it authentic and balanced. \n\n6. When generating the dialogue, keep in mind the following context and background information: {DialogueContext}. \n\n7. Adapt the style to the target audience: {TargetAudience}. \n\n8. Format strictly: Start host lines with 'H:' and guest lines with 'G:', each on a new paragraph. \n\n9. Keep the dialogue dynamic—avoid long monologues or unnatural phrasing. Use short, engaging sentences with occasional rhetorical questions or casual expressions to make it feel real. \n\n10. Your instruction with regard to the duration of the dialogue is: {Duration}. Make sure, you create a script that will result in speech of this duration (e.g., if the instruction is 10 minutes, then create text for ten minutes of discussion, and not only five minutes, which would be wrong, hence, you may need to do a deeper dive). \n\n11. Use SSML to improve pronunciation and pacing: '<say-as interpret-as=\""characters\"">' for abbreviations and acronyms of up to three letters or with numbers (e.g., <say-as interpret-as=\""characters\"">KI</say-as> where there are abbreviations acronyms of up to three or with numbers where you are not sure how they are spoken; abbreviations and acronyms of four or more letters, read them normally), '<lang xml:lang=\""en-US\"">' for foreign words (e.g., <lang xml:lang=\""en-US\"">Artificial Intelligence</lang>), and '<say-as>' for numbers, dates, and symbols. \n\n12. Apply '<emphasis level=\""moderate\"">' or '<emphasis level=\""strong\"">'only to **key words or very important points that should stand out naturally**—avoid artificial exaggeration. \n\n13. Use '<prosody rate=\""medium\"">' to **maintain a natural speaking rhythm** and prevent robotic speech—do not use 'slow' unless necessary for dramatic effect. \n\n14. When a dash ('-') appears, replace it with '<break time=\""500ms\"">' to introduce a natural pause and prevent rushed pronunciation. \n\n15. The final dialogue should sound like two real people having an **authentic and fluid conversation**, completely in the language in rule no. 1, without artificial slowness, exaggeration, or awkward phrasing. Keep in mind that your output will be spoken, not read. \n\16. It is important that you really comply with these rules, otherwise the output will be invalid. 17. Finally, here are additional instructions (if any) that override any other instructions given so far and are to be followed precisely: {ExtraInstructions} {INI_PreCorrection}\n\n\n"
        Const Default_SP_Explain As String = "You are a great thinker, a specialist in all fields, a philosoph and a teacher. Create me an advanced prompt for an advanced large language model that will analyze a Text (the Texttoprocess) it is provided between the tags <TEXTTOPROCESS> and </TEXTTOPROCESS>. Step 1: Thorougly analyze the text you have been given, its logic, identify any errors and fallacies of the author, understand the substance the author discusses and the way the author argues. Do not yet create any output. Once you have completed step 1, go to Step 2: Start your output with a one word summary (in bold, as a title) and a further title that captures all relevant substance and bottomline of the text (do not refer to it as a summary or title, just provide it as the title of your analysis). Then explain in simple, short and consise terms what the author wants to say and expressly list any explicit or implicit 'Calls to Action' are. Now, insofar the author makes arguments, provide me a description of the logic and approach the author takes in making the point, including any errors, ambiguities, contradictions and fallacies you can identify. Finally, insofar the author discusses a special fied of knowledge, provide in detail the necessary background knowledge a layman needs to know to fully understand the text, the special terms and concepts used by the text, including technology, methods and art and sciences discussed in it. When acronyms, terms or other references could have different meanings and it is not absolutely clear what they are in the present context, express such uncertainty. If you make assumptions, say so, explain why and only where they are clear. Provide the output well structured, concise, short and simple, easy to understand and provide it in the original language of the Texttoprocess. {INI_PreCorrection}"
        Const Default_SP_SuggestTitles As String = "You are a legal professional and a clever, astute and well-educated copy editor. You are in the following given a text, enclosed between <TEXTTOPROCESS> and </TEXTTOPROCESS>. Your goal is to read and analyze the content, then create multiple sets of possible titles in the same language as the original text, with three (3) distinct titles each for: (1) professional memo, (2) blog/news post, (3) informal, (4) humorous, and (5) ambiguous, cryptic but ingenious. The titles must be clever, easy to read, well-aligned with the text, and suitable for the stated purpose. Provide more than average results. Use the structure:\nProfessional Memo Titles:\n1) ...\n2) ...\n3) ...\nBlog or News Post Titles:\n1) ...\n2) ...\n3) ...\nInformal Titles:\n1) ...\n2) ...\n3) ...\nHumorous Titles:\n1) ...\n2) ...\n3) ...\nFood for Thought Titles:\n1) ...\n2) ...\n3) ...\n. It is mandatory that you provide your output and all titles provide in the original language of the Texttoprocess."
        Const Default_SP_Friendly As String = "You are a legal professional with exceptional language skills who follows instructions meticulously step by step. Your task is to refine the text labeled 'Texttoprocess' (in its original language) to make it more friendly, while otherwise preserving its substance, wording and style. Use rhetorical techniques and wording that is typically well received and generates a positive attitude by the recipient, but stay straightforward, and do neither exaggerate nor brownnose. Whenever there is a line feed or carriage return in text provided to you, it is essential that you also include such line feed or carriage return in the output you generate. The carriage returns and line feeds in the output must match exactly those in the original text provided to you. Accordingly, if there are two carriage returns or line feeds in succession in the text provided to you, there must also be two carriage returns or line feeds in the text you generate. Also, only provide the revised text, never provide any explanations or comments on how you have fulfilled your instructions.  {INI_PreCorrection}"
        Const Default_SP_Convincing As String = "You are a legal professional with exceptional language skills who follows instructions meticulously  step by step. Your task is to refine the text labeled 'Texttoprocess' (in its original language) to make it more convincing. Make it more persuasive and concise by the way you amend the language, but preserve its original substance and style. Do not alter the underlying content and arguments, but use rhetorical and language techniques to make the text more convincing, but do not exaggerate and do not brownnose. Whenever there is a line feed or carriage return in text provided to you, it is essential that you also include such line feed or carriage return in the output you generate. The carriage returns and line feeds in the output must match exactly those in the original text provided to you. Accordingly, if there are two carriage returns or line feeds in succession in the text provided to you, there must also be two carriage returns or line feeds in the text you generate. Also, only provide the revised text, never provide any explanations or comments on how you have fulfilled your instructions. {INI_PreCorrection}"
        Const Default_SP_NoFillers As String = "You are a legal professional with exceptional language skills who follows instructions meticulously step by step. Amend the text that is provided to you, in its original language, and is labeled as 'Texttoprocess' as follows: 1. Remove any and all filler words and any and all other words that do not add any meaning or are not necessary for understanding and easily reading the text. 2. Remove any other redundant language or other redunancies. 3. Change passive voice to active voice but only where this is easily possible without changing the entire sentence. 4. Ensure that the text is easy to read, concise and clear. 5. Do not alter the text's overall flow, readability, content, meaning, tone and style. 6. Do not change or remove words where you are not sure whether they are necessary for good reading and content; the text should remain easily readable and not appear choppy or abbreviated. 7. Before you provide me with the revised text, compare its meaning with the the original text and ensure that it remains the same. Otherwise adapt the output to ensure that the meaning of the revised text stays the same as with the original text. 8. Never remove or add line breaks, carriage returns or vertical tabs from the text you are provided. 9. Also, only provide the revised text, never provide any explanations or comments on how you have fulfilled your instructions.{INI_PreCorrection}"

        Public Shared LicensedTill As Date = CDate("1.1.2000")

        Public Shared Function GetDefaultINIPath(ByVal key As String) As String
            For Each entry In DefaultINIPaths
                If key.Contains(entry.Key) Then
                    Return ExpandEnvironmentVariables(entry.Value)
                End If
            Next
            Return ExpandEnvironmentVariables(DefaultINIPaths.Values.First())
        End Function

        Public Class SplashScreen

            Inherits Form

            Public Sub New(Optional customText As String = "Please wait ...", Optional formWidth As Integer = 300, Optional formHeight As Integer = 100)
                ' Set the form properties
                Me.Text = $"{AN}"
                Me.FormBorderStyle = FormBorderStyle.None
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.Top -= 40
                Me.BackColor = ColorTranslator.FromWin32(&H8000000F)

                ' Set a predefined font for consistency
                Dim standardFont As New System.Drawing.Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)

                ' Create the PictureBox
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                Dim pictureBox As New PictureBox()
                pictureBox.Image = bmp
                pictureBox.SizeMode = PictureBoxSizeMode.Zoom
                pictureBox.SetBounds(10, 10, 30, 30)

                ' Create the Label with updated font
                Dim label As New System.Windows.Forms.Label()
                label.Text = customText
                label.Font = standardFont
                label.AutoSize = True
                label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

                ' Dynamically calculate the label width
                Dim labelSize As Size = TextRenderer.MeasureText(label.Text, standardFont)
                label.SetBounds(pictureBox.Right + 10, 15, labelSize.Width, labelSize.Height)

                ' Adjust the form size dynamically based on the provided dimensions
                Dim contentWidth As Integer = pictureBox.Width + label.Width + 40 ' Add padding for spacing
                Dim contentHeight As Integer = Math.Max(pictureBox.Height + 20, label.Height + 30) ' Align to bottom of logo
                Me.ClientSize = New System.Drawing.Size(Math.Max(formWidth, contentWidth), contentHeight)


                ' Add the controls to the form
                Me.Controls.Add(pictureBox)
                Me.Controls.Add(label)
            End Sub

        End Class

        Public Class Diff
            Public Enum Operation
                Equal
                Insert
                Delete
            End Enum

            Public Property Op As Operation
            Public Property Text As String

            Public Sub New(op As Operation, text As String)
                Me.Op = op
                Me.Text = text
            End Sub
        End Class


        Private Shared clipboardData As Object = Nothing ' Variable to store clipboard content

        Public Shared Sub StoreClipboard()
            If Clipboard.ContainsText() Then
                clipboardData = Clipboard.GetText()
            ElseIf Clipboard.ContainsImage() Then
                clipboardData = Clipboard.GetImage()
            ElseIf Clipboard.ContainsData(DataFormats.Serializable) Then
                clipboardData = Clipboard.GetData(DataFormats.Serializable)
            ElseIf Clipboard.ContainsData(DataFormats.FileDrop) Then
                clipboardData = Clipboard.GetData(DataFormats.FileDrop)
            Else
                clipboardData = Nothing ' No supported data format found
            End If
        End Sub

        Public Shared Sub RestoreClipboard()
            If clipboardData Is Nothing Then Exit Sub

            If TypeOf clipboardData Is String Then
                Clipboard.SetText(CStr(clipboardData))
            ElseIf TypeOf clipboardData Is Image Then
                Clipboard.SetImage(CType(clipboardData, Image))
            ElseIf TypeOf clipboardData Is Object Then
                Clipboard.SetData(DataFormats.Serializable, clipboardData)
            End If
        End Sub

        Public Shared Sub InsertTextWithBoldMarkers(selection As Object, gptResult As String)
            ' Save the starting position of the insertion
            Dim startPosition As Integer = selection.Start

            ' Split the text by "**" to identify bold and regular sections
            Dim parts() As String
            parts = Split(gptResult, "**")

            ' Iterate through the parts and add text with appropriate formatting
            For i As Integer = 0 To UBound(parts)
                If i Mod 2 = 1 Then
                    ' Odd-index parts are bold
                    selection.Font.Bold = True
                Else
                    ' Even-index parts are normal text
                    selection.Font.Bold = False
                End If

                ' Insert the text part
                If parts(i) <> "" Then
                    selection.TypeText(parts(i))
                End If
            Next

            ' Reset bold formatting to normal after insertion
            selection.Font.Bold = False

            ' Save the end position of the insertion
            Dim endPosition As Integer = selection.Start

            ' Select the entire inserted text
            selection.SetRange(startPosition, endPosition)
        End Sub

        Public Shared Sub InsertTextWithFormat(formattedText As String, ByRef range As Microsoft.Office.Interop.Word.Range, ReplaceSelection As Boolean)

            Try
                Dim htmlHeader As String = "<html><head><meta charset=""UTF-8""></head><body><!--StartFragment-->"
                Dim htmlFooter As String = "<!--EndFragment--></body></html>"

                Dim cleanedHtml As String = $"{htmlHeader}{formattedText}{htmlFooter}"

                cleanedHtml = CreateProperHtml(cleanedHtml).Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, "")

                ' Combine into the full HTML content
                Dim dummyHtml As String = $"Version:0.9{vbCrLf}" &
                      $"StartHTML:00000000{vbCrLf}" &
                      $"EndHTML:00000000{vbCrLf}" &
                      $"StartFragment:00000000{vbCrLf}" &
                      $"EndFragment:00000000{vbCrLf}" & cleanedHtml

                ' Calculate offsets
                Dim startHtmlOffset As Integer = dummyHtml.IndexOf("<html>")
                Dim endHtmlOffset As Integer = dummyHtml.Length
                Dim startFragmentOffset As Integer = dummyHtml.IndexOf("<!--StartFragment-->") + "<!--StartFragment-->".Length
                Dim endFragmentOffset As Integer = dummyHtml.IndexOf("<!--EndFragment-->")

                ' Replace placeholders
                Dim finalHtml As String = dummyHtml.Replace("StartHTML:00000000", $"StartHTML:{startHtmlOffset:D8}") _
                               .Replace("EndHTML:00000000", $"EndHTML:{endHtmlOffset:D8}") _
                               .Replace("StartFragment:00000000", $"StartFragment:{startFragmentOffset:D8}") _
                               .Replace("EndFragment:00000000", $"EndFragment:{endFragmentOffset:D8}")

                ' Perform the clipboard operation on an STA thread
                Dim clipboardThread As New Threading.Thread(Sub()
                                                                ClipboardSet(finalHtml)
                                                            End Sub)
                clipboardThread.SetApartmentState(Threading.ApartmentState.STA)
                clipboardThread.Start()
                clipboardThread.Join()

                ' Move the selection to the specified range
                range.Select()

                ' Replace the selected text if needed
                If ReplaceSelection Then
                    range.Text = ""
                    range.Application.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdFormatSurroundingFormattingWithEmphasis)
                Else
                    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    range.Select()
                    range.Application.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdFormatSurroundingFormattingWithEmphasis)
                End If

                ' After the paste
                System.Threading.Thread.Sleep(100)

                range = range.Application.Selection.Range


            Catch ex As System.Exception
                MessageBox.Show("InsertTextWithFormat Error: " & ex.Message)
            End Try

        End Sub


        Public Shared Function RemoveHTML(html As String) As String

            If String.IsNullOrEmpty(html) Then
                Return String.Empty
            End If

            ' Replace <br> and </p> with vbCrLf.
            ' Handle variations like <br>, <br/>, <br />, and </p> in a case-insensitive manner
            html = Regex.Replace(html, "</p>", vbCrLf, RegexOptions.IgnoreCase)
            html = Regex.Replace(html, "<br\s*/?>", vbCrLf, RegexOptions.IgnoreCase)

            ' Load into HtmlAgilityPack to remove remaining tags and handle entities
            Dim doc As New HtmlAgilityPack.HtmlDocument()
            doc.LoadHtml(html)

            ' Get the inner text (this strips out all remaining HTML tags)
            Dim textContent As String = doc.DocumentNode.InnerText

            ' Decode HTML entities (including special characters and umlauts)
            ' HtmlEntity.DeEntitize converts HTML encoded characters to their decoded form
            textContent = HtmlEntity.DeEntitize(textContent)

            ' Remove extra line breaks or whitespace caused by replaced tags
            ' Convert multiple consecutive line breaks into a single one 
            textContent = Regex.Replace(textContent, "(?<!\\)\\[rnt]", Function(m)
                                                                           Select Case m.Value
                                                                               Case "\n" : Return vbLf
                                                                               Case "\r" : Return vbCr
                                                                               Case "\t" : Return vbTab
                                                                               Case Else : Return m.Value
                                                                           End Select
                                                                       End Function)

            ' Trim leading and trailing whitespace
            textContent = textContent.Trim()

            Return textContent
        End Function



        Public Shared Function ConvertMarkupToRTF(inputText As String) As String
            ' Define the RTF header with font and color tables
            Dim rtfHeader As String =
                    "{\rtf1\ansi\deff0" &
                    "{\fonttbl{\f0\fnil\fcharset0 Calibri;}}" &
                    "{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red255\green0\blue0;}" &
                    "\f0\fs20\cf1 "

            ' Replace custom markup with RTF formatting
            Dim rtfContent As String = inputText.Replace(vbCrLf, "\r\n").Replace(vbCr, "\r").Replace(vbLf, "\n")

            ' Convert [DEL_START] ... [DEL_END] to red + strikethrough
            rtfContent = Regex.Replace(rtfContent, "\[DEL_START\](.*?)\[DEL_END\]", "{\cf3\strike $1}{\strike0}", RegexOptions.Singleline)

            ' Convert [INS_START] ... [INS_END] to blue + underline
            rtfContent = Regex.Replace(rtfContent, "\[INS_START\](.*?)\[INS_END\]", "{\cf2\ul $1}{\ul0}", RegexOptions.Singleline)

            ' Convert newlines to RTF paragraph breaks  yyyyyy
            rtfContent = Regex.Replace(rtfContent, "(?<!\\)\\r\\n", "\par ")
            rtfContent = Regex.Replace(rtfContent, "(?<!\\)\\r", "\par ")
            rtfContent = Regex.Replace(rtfContent, "(?<!\\)\\n", "\par ")

            ' Add RTF footer
            Dim rtfFooter As String = "}"

            ' Combine and return the full RTF string
            Return rtfHeader & rtfContent & rtfFooter
        End Function


        Public Shared Sub ClipboardSet(finalHtml As String)
            Try
                ' Encode the final HTML as UTF-8
                Dim utf8Bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(finalHtml)
                Dim utf8String As String = System.Text.Encoding.UTF8.GetString(utf8Bytes)

                ' Create a DataObject with the UTF-8 content
                Dim dataObject As New DataObject()
                dataObject.SetText(utf8String, TextDataFormat.Html)
                Clipboard.SetDataObject(dataObject, True)
            Catch ex As System.Exception
                MessageBox.Show("Error setting clipboard data: " & ex.Message)
            End Try
        End Sub


        Public Shared Function CreateProperHtml(inputHtml As String) As String
            ' Create a new HtmlDocument

            inputHtml = inputHtml.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

            Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()

            ' Load the input HTML string into the document
            htmlDoc.LoadHtml(inputHtml)

            Dim headTag = htmlDoc.DocumentNode.SelectSingleNode("//head")
            If headTag Is Nothing Then
                headTag = HtmlAgilityPack.HtmlNode.CreateNode("<head></head>")
                Dim htmlTag = htmlDoc.DocumentNode.SelectSingleNode("//html")
                If htmlTag Is Nothing Then
                    htmlTag = HtmlAgilityPack.HtmlNode.CreateNode("<html></html>")
                    htmlDoc.DocumentNode.AppendChild(htmlTag)
                End If
                htmlTag.PrependChild(headTag)
            End If

            If Not headTag.InnerHtml.Contains("charset") Then
                headTag.InnerHtml = "<meta charset=""UTF-8"">" & headTag.InnerHtml
            End If

            ' Process text nodes
            For Each textNode As HtmlAgilityPack.HtmlNode In htmlDoc.DocumentNode.DescendantsAndSelf().Where(Function(node) node.NodeType = HtmlAgilityPack.HtmlNodeType.Text)
                ' Only encode text that does not already contain valid HTML entities
                Dim originalText = textNode.InnerHtml
                If Not originalText.Contains("&") Then
                    textNode.InnerHtml = System.Net.WebUtility.HtmlEncode(originalText)
                End If
            Next

            ' Return the cleaned-up, well-formed HTML
            Return htmlDoc.DocumentNode.OuterHtml
        End Function

        Public Shared Function GetRangeHtml(ByVal range As Range) As String
            Dim htmlContent As String = String.Empty
            Dim tempFile As String = System.IO.Path.GetTempFileName()

            Try
                ' Save the range as a filtered HTML file
                range.ExportFragment(FileName:=tempFile, Format:=WdSaveFormat.wdFormatFilteredHTML)

                ' Read the HTML content
                htmlContent = System.IO.File.ReadAllText(tempFile)
            Finally
                ' Delete the temporary file
                If System.IO.File.Exists(tempFile) Then
                    System.IO.File.Delete(tempFile)
                End If
            End Try

            htmlContent = SimplifyHtml(htmlContent)

            Return htmlContent
        End Function

        Public Shared Function SimplifyHtml(htmlContent As String) As String
            ' Load the HTML content into an HtmlDocument
            Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
            htmlDoc.LoadHtml(htmlContent)

            ' Process the document to remove irrelevant tags and attributes
            CleanHtmlNode(htmlDoc.DocumentNode)

            ' Get the simplified HTML
            Dim simplifiedHtml As String = htmlDoc.DocumentNode.OuterHtml

            ' Remove real line breaks
            simplifiedHtml = simplifiedHtml.Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, "")

            ' Return the simplified HTML
            Return simplifiedHtml
        End Function

        Public Shared Sub CleanHtmlNode(node As HtmlNode)
            If node.NodeType = HtmlNodeType.Element Then
                ' Define the allowed tags
                Dim allowedTags As HashSet(Of String) = New HashSet(Of String) From {"b", "strong", "i", "em", "u", "font", "span", "p", "ul", "ol", "li", "br"}

                ' Define the allowed attributes
                Dim allowedAttributes As HashSet(Of String) = New HashSet(Of String) From {"style", "class"}

                ' Remove attributes that are not in the allowed list
                For Each attr In node.Attributes.ToList()
                    If Not allowedAttributes.Contains(attr.Name.ToLower()) Then
                        node.Attributes.Remove(attr.Name)
                    End If
                Next

                ' If the node is not an allowed tag, replace it with its inner content
                If Not allowedTags.Contains(node.Name.ToLower()) Then
                    Dim parentNode = node.ParentNode
                    Dim innerNodes = node.ChildNodes.ToList()
                    For Each innerNode In innerNodes
                        If innerNode.Name.ToLower() = "p" OrElse innerNode.Name.ToLower() = "br" Then
                            parentNode.InsertBefore(HtmlNode.CreateNode(innerNode.OuterHtml), node)
                        Else
                            parentNode.InsertBefore(innerNode, node)
                        End If
                    Next
                    parentNode.RemoveChild(node)
                End If
            End If

            ' Recursively process child nodes
            For Each childNode In node.ChildNodes.ToList()
                CleanHtmlNode(childNode)
            Next
        End Sub

        Public Class GoogleOAuthHelper
            ' Public variables
            Public Shared client_email As String = ""
            Public Shared private_key As String = ""
            Public Shared scopes As String = ""
            Public Shared token_uri As String = ""
            Public Shared token_life As Long = 0

            ' Base64Url encoding
            Private Shared Function Base64UrlEncode(input As String) As String
                Return Convert.ToBase64String(Encoding.UTF8.GetBytes(input)).
                Replace("+", "-").
                Replace("/", "_").
                Replace("=", "")
            End Function

            Private Shared Function Base64UrlEncode(inputBytes As Byte()) As String
                Return Convert.ToBase64String(inputBytes).
                Replace("+", "-").
                Replace("/", "_").
                Replace("=", "")
            End Function

            ' Sign data using BouncyCastle
            Private Shared Function SignData(data As Byte()) As Byte()
                Dim rsaKey As RsaPrivateCrtKeyParameters
                Dim formattedPrivateKey As String = private_key.Replace("\n", Environment.NewLine)

                Using reader As New StringReader(formattedPrivateKey)
                    Dim pemReader = New Org.BouncyCastle.OpenSsl.PemReader(reader)
                    rsaKey = DirectCast(pemReader.ReadObject(), RsaPrivateCrtKeyParameters)
                End Using

                Dim signer = SignerUtilities.GetSigner("SHA256withRSA")
                signer.Init(True, rsaKey)
                signer.BlockUpdate(data, 0, data.Length)
                Return signer.GenerateSignature()
            End Function

            ' Generate JWT
            Public Shared Function GenerateJWT() As String
                Dim issuedAt As Long = DateTimeOffset.UtcNow.ToUnixTimeSeconds()
                Dim expiry As Long = issuedAt + 3600 ' 1 hour expiry

                Dim header = New With {.alg = "RS256", .typ = "JWT"}
                Dim payload = New With {
                                        .iss = client_email,
                                        .scope = scopes,
                                        .aud = token_uri,
                                        .exp = expiry,
                                        .iat = issuedAt
                                    }

                Dim headerBase64 = Base64UrlEncode(JsonConvert.SerializeObject(header))
                Dim payloadBase64 = Base64UrlEncode(JsonConvert.SerializeObject(payload))
                Dim unsignedToken = $"{headerBase64}.{payloadBase64}"
                Dim signature = SignData(Encoding.UTF8.GetBytes(unsignedToken))
                Dim signatureBase64 = Base64UrlEncode(signature)

                Return $"{unsignedToken}.{signatureBase64}"
            End Function

            ' Get Access Token
            Public Shared Async Function GetAccessToken() As Task(Of String)
                Dim jwt = GenerateJWT()
                Dim requestBody As String = JsonConvert.SerializeObject(New With {
                                .grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer",
                                .assertion = jwt
                                            })

                Using client As New HttpClient()
                    Dim content = New StringContent(requestBody, Encoding.UTF8, "application/json")
                    Dim response = Await client.PostAsync(token_uri, content)

                    If response.IsSuccessStatusCode Then
                        Dim responseBody = Await response.Content.ReadAsStringAsync()
                        Dim tokenData = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(responseBody)
                        Return tokenData("access_token")
                    Else
                        ShowCustomMessageBox($"Error getting access token: {response.ReasonPhrase}")
                        Return ""
                    End If
                End Using
            End Function
        End Class

        Public Shared Async Function PostCorrection(context As ISharedContext, inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
            Dim OutputText As String = inputText
            If Not String.IsNullOrEmpty(context.INI_PostCorrection) Then

                ' Wait not to overload the API
                Await System.Threading.Tasks.Task.Delay(500)

                OutputText = Await LLM(context, context.INI_PostCorrection, "<TEXTTOPROCESS>" & inputText & "</TEXTTOPROCESS>", "", "", 0, UseSecondAPI)
            End If
            Return OutputText
        End Function

        Public Shared Async Function LLM(context As ISharedContext, ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal Hidesplash As Boolean = False, Optional ByVal AddUserPrompt As String = "") As Task(Of String)

            Dim splash As New SplashScreen("Waiting for the LLM to respond...")

            If Not Hidesplash Then
                splash.Show()
                splash.Refresh()
            End If

            Try

                ' Configure TLS
                If (System.Net.ServicePointManager.SecurityProtocol And System.Net.SecurityProtocolType.Tls12) = 0 Then
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.ServicePointManager.SecurityProtocol Or System.Net.SecurityProtocolType.Tls12
                End If
                System.Net.ServicePointManager.DefaultConnectionLimit = 100 ' Adjust based on expected load

                ' Initialize API variables
                Dim Endpoint As String
                Dim HeaderA As String
                Dim HeaderB As String
                Dim APICall As String
                Dim TemperatureValue As String
                Dim ModelValue As String
                Dim TimeoutValue As Long
                Dim ResponseKey As String
                Dim DoubleS As Boolean

                If UseSecondAPI Then

                    If context.INI_OAuth2_2 Then
                        context.DecodedAPI_2 = Await GetFreshAccessToken(context, context.INI_OAuth2ClientMail_2, context.INI_OAuth2Scopes_2, context.INI_APIKey_2, context.INI_OAuth2Endpoint_2, context.INI_OAuth2ATExpiry_2, True)
                        If context.DecodedAPI_2 = "" Then Exit Function
                    End If

                Else
                    If context.INI_OAuth2 Then
                        context.DecodedAPI = Await GetFreshAccessToken(context, context.INI_OAuth2ClientMail, context.INI_OAuth2Scopes, context.INI_APIKey, context.INI_OAuth2Endpoint, context.INI_OAuth2ATExpiry, False)
                        If context.DecodedAPI = "" Then Exit Function
                    End If
                End If

                If UseSecondAPI Then

                    Endpoint = Replace(Replace(context.INI_Endpoint_2, "{model}", context.INI_Model_2), "{apikey}", context.DecodedAPI_2)
                    HeaderA = Replace(Replace(context.INI_HeaderA_2, "{model}", context.INI_Model_2), "{apikey}", context.DecodedAPI_2)
                    HeaderB = Replace(Replace(context.INI_HeaderB_2, "{model}", context.INI_Model_2), "{apikey}", context.DecodedAPI_2)
                    APICall = context.INI_APICall_2
                    ResponseKey = context.INI_Response_2
                    DoubleS = context.INI_DoubleS

                    TemperatureValue = If(String.IsNullOrEmpty(Temperature) OrElse Temperature = "Default", context.INI_Temperature_2, Temperature)
                    ModelValue = If(String.IsNullOrEmpty(Model) OrElse Model = "Default", context.INI_Model_2, Model)
                    TimeoutValue = If(Timeout = 0, context.INI_Timeout_2, Timeout)
                Else

                    Endpoint = Replace(Replace(context.INI_Endpoint, "{model}", context.INI_Model), "{apikey}", context.DecodedAPI)
                    HeaderA = Replace(Replace(context.INI_HeaderA, "{model}", context.INI_Model), "{apikey}", context.DecodedAPI)
                    HeaderB = Replace(Replace(context.INI_HeaderB, "{model}", context.INI_Model), "{apikey}", context.DecodedAPI)
                    APICall = context.INI_APICall
                    ResponseKey = context.INI_Response
                    DoubleS = context.INI_DoubleS
                    TemperatureValue = If(String.IsNullOrEmpty(Temperature) OrElse Temperature = "Default", context.INI_Temperature, Temperature)
                    ModelValue = If(String.IsNullOrEmpty(Model) OrElse Model = "Default", context.INI_Model, Model)
                    TimeoutValue = If(Timeout = 0, context.INI_Timeout, Timeout)
                End If

                ' Replace placeholders in the request body
                Dim requestBody As String = APICall
                requestBody = requestBody.Replace("{model}", ModelValue)
                requestBody = requestBody.Replace("{promptsystem}", CleanString(promptSystem))
                requestBody = requestBody.Replace("{promptuser}", CleanString(promptUser))
                requestBody = requestBody.Replace("{userinstruction}", CleanString(AddUserPrompt))
                requestBody = requestBody.Replace("{temperature}", TemperatureValue)

                Dim Returnvalue As String = ""

                ' Configure HttpClient with timeout
                Using handler As New System.Net.Http.HttpClientHandler()
                    handler.UseProxy = True
                    handler.Proxy = System.Net.WebRequest.DefaultWebProxy
                    handler.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
                    Using client As New System.Net.Http.HttpClient(handler)
                        client.Timeout = TimeSpan.FromMilliseconds(TimeoutValue)

                        ' Add headers
                        If Not String.IsNullOrEmpty(HeaderA) AndAlso Not String.IsNullOrEmpty(HeaderB) Then
                            client.DefaultRequestHeaders.Add(HeaderA, HeaderB)
                        End If

                        If context.INI_APIDebug Then
                            Debug.WriteLine($"SENT TO API:{Environment.NewLine}{requestBody}")
                        End If


                        ' Send the request
                        Try
                            Dim requestContent As New System.Net.Http.StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")
                            Dim response As System.Net.Http.HttpResponseMessage = Await client.PostAsync(Endpoint, requestContent)

                            ' Check response status
                            If Not response.IsSuccessStatusCode Then
                                Dim errorContent As String = Await response.Content.ReadAsStringAsync()
                                ShowCustomMessageBox($"HTTP Error {response.StatusCode} when accessing the LLM endpoint: {errorContent}")
                                Return ""
                            End If

                            ' Read response
                            Dim responseText As String = Await response.Content.ReadAsStringAsync()

                            If context.INI_APIDebug Then
                                Debug.WriteLine($"RECEIVED FROM API:{Environment.NewLine}{responseText}")
                            End If

                            ' Process the response
                            Dim jsonObject As JObject = JObject.Parse(responseText)

                            ' Extract the "error" segment
                            Dim text As String = FindJsonProperty(jsonObject, "error")
                            'Dim text As String = ExtractJSONValue(responseText, "error")

                            If Not String.IsNullOrEmpty(text) Then
                                text = FindJsonProperty(jsonObject, "message")
                                'text = ExtractJSONValue(responseText, "message")
                                ShowCustomMessageBox($"The LLM API generated the following error message: {Environment.NewLine}{text}{Environment.NewLine}{responseText}" & If(text.Contains("429"), vbCrLf & text & vbCrLf & "Note: The error '429' means that the LLM got too many requests or in too quick succession. Please wait a few seconds and try again or slow down or reduce your request.", ""))
                                Return ""
                            Else
                                text = FindJsonProperty(jsonObject, ResponseKey)

                                text = text & ExtractCitations(jsonObject)

                                ' Previous implementation
                                ' text = ExtractJSONValue(responseText, ResponseKey)
                                ' text = ConvertEscapeCharacters(text)
                                If DoubleS Then
                                    text = text.Replace(ChrW(223), "ss") ' Replace German sharp-S if needed
                                End If
                                Returnvalue = text
                            End If
                        Catch ex As System.Net.Http.HttpRequestException
                            ShowCustomMessageBox($"An HTTP request exception occurred: {ex.Message} when accessing the LLM endpoint (2).")
                        Catch ex As TaskCanceledException
                            ShowCustomMessageBox($"The request timed out. Please try again or increase the timeout setting.")
                        End Try
                    End Using ' Dispose HttpClient
                End Using ' Dispose HttpClientHandler
                Return Returnvalue

            Catch ex As System.Exception
                ShowCustomMessageBox($"An unexpected error occurred when accessing the LLM endpoint: {ex.Message}")
                Return ""
            Finally
                If Not Hidesplash Then
                    splash.Close()
                End If
            End Try
        End Function



        Public Shared Function ExtractCitations(ByRef jsonObj As JObject) As String
                Try
                    Dim OriginalJsonObj As JObject = jsonObj.DeepClone()
                    Dim citationList As New List(Of String)
                    Dim sourceUris As New HashSet(Of String)

                    ' 1. Attempt extraction from candidates path (if present)
                    Dim candidateCitations As JToken = jsonObj.SelectToken("candidates[0].content.parts[0].citations")
                    If candidateCitations IsNot Nothing Then
                        If candidateCitations.Type = JTokenType.Array Then
                            For Each citation As JObject In candidateCitations
                                ProcessCitationObject(citation, citationList, sourceUris)
                            Next
                        ElseIf candidateCitations.Type = JTokenType.Object Then
                            ProcessCitationObject(CType(candidateCitations, JObject), citationList, sourceUris)
                        End If
                    End If

                    ' 2. Check for top-level citations (outside of candidates)
                    Dim topLevelCitations As JToken = jsonObj.SelectToken("citations")
                    If topLevelCitations IsNot Nothing Then
                        If topLevelCitations.Type = JTokenType.Array Then
                            For Each citation As JToken In topLevelCitations
                                If citation.Type = JTokenType.String Then
                                    citationList.Add(citation.ToString())
                                ElseIf citation.Type = JTokenType.Object Then
                                    ProcessCitationObject(CType(citation, JObject), citationList, sourceUris)
                                End If
                            Next
                        ElseIf topLevelCitations.Type = JTokenType.Object Then
                            ' Handle Format 2 (fullNote/shortNote) in a top-level object
                            Dim fullNote As String = topLevelCitations("fullNote")?.ToString()
                            If Not String.IsNullOrEmpty(fullNote) Then
                                citationList.Add(fullNote)
                            End If
                            Dim shortNote As String = topLevelCitations("shortNote")?.ToString()
                            If Not String.IsNullOrEmpty(shortNote) Then
                                citationList.Add(shortNote)
                            End If
                            ' In case no fullNote exists, fallback to checking for a URL
                            Dim url As String = topLevelCitations("url")?.ToString()
                            If Not String.IsNullOrEmpty(url) Then
                                citationList.Add(url)
                            End If
                        End If
                    End If

                    ' 3. Check citation metadata sources
                    Dim metadataSources As JToken = jsonObj.SelectToken("citationMetadata.citationSources")
                    If metadataSources IsNot Nothing AndAlso metadataSources.Type = JTokenType.Array Then
                        For Each source As JObject In metadataSources
                            ProcessMetadataSource(source, citationList, sourceUris)
                        Next
                    End If

                    ' 4. Check legacy formats
                    ExtractLegacyCitations(jsonObj, citationList, sourceUris)

                    Debug.WriteLine("Total citations count: " & citationList.Count.ToString())

                    ' 5. Build output: if any citation was found, format them;
                    ' otherwise, fall back to the simple citations extractor.
                    If citationList.Count > 0 Then
                        Debug.WriteLine("Citations: " & String.Join(", ", citationList))
                        Return FormatCitations(citationList)
                    Else
                        Dim result As String = ExtractSimpleCitations(OriginalJsonObj)
                        Debug.WriteLine("Fallback Result = " & result)
                        Return result
                    End If

                Catch ex As Exception
                Debug.WriteLine("Error parsing citations: " & ex.Message)
            End Try

                Return String.Empty
            End Function

            Private Shared Sub ProcessCitationObject(citation As JObject, ByRef citationList As List(Of String), ByRef sourceUris As HashSet(Of String))
                Try
                    ' Format 1: Check for a "source" property (MLA/Chicago style)
                    Dim source = citation.SelectToken("source")
                    If source IsNot Nothing Then
                        AddSource(source, citationList, sourceUris)
                        ' Optionally include an inline citation if available
                        Dim inlineCitation = citation("inlineCitation")?.ToString()
                        If Not String.IsNullOrEmpty(inlineCitation) Then
                            citationList.Add("Inline: " & inlineCitation)
                        End If
                        Return
                    End If

                    ' Format 2: Check for a "fullNote" property (full note/short note format)
                    Dim fullNote As String = citation("fullNote")?.ToString()
                    If Not String.IsNullOrEmpty(fullNote) Then
                        citationList.Add(fullNote)
                        Return
                    End If

                    ' Format 3: IEEE style with "referenceEntry"
                    Dim refEntry As String = citation("referenceEntry")?.ToString()
                    If Not String.IsNullOrEmpty(refEntry) Then
                        Dim ieeeUri As String = ExtractIeeeUri(refEntry)
                        If Not String.IsNullOrEmpty(ieeeUri) AndAlso sourceUris.Add(ieeeUri) Then
                            citationList.Add($"{refEntry} | Source: {ieeeUri}")
                        Else
                            citationList.Add(refEntry)
                        End If
                        Return
                    End If

                    ' Format 4: Harvard style with "referenceList.entry" and optionally "textualCitation"
                    Dim refListToken As JToken = citation.SelectToken("referenceList.entry")
                    If refListToken IsNot Nothing Then
                        Dim refList As String = refListToken.ToString()
                        Dim harvardUri As String = ExtractHarvardUri(refList)
                        Dim textualCitation As String = citation("textualCitation")?.ToString()
                        Dim formattedCitation As String = (If(Not String.IsNullOrEmpty(textualCitation), textualCitation, "") & " | " & refList).Trim(" "c, "|"c)
                        citationList.Add(formattedCitation)
                        Return
                    End If

                    ' Fallback: If the citation object has a "url" property directly, extract it.
                    Dim url As String = citation("url")?.ToString()
                    If Not String.IsNullOrEmpty(url) AndAlso sourceUris.Add(url) Then
                        citationList.Add(url)
                    End If

                Catch ex As Exception
                Debug.WriteLine("Error processing citation object: " & ex.Message)
            End Try
            End Sub

            Private Shared Sub ProcessMetadataSource(source As JObject, ByRef citationList As List(Of String), ByRef sourceUris As HashSet(Of String))
                Try
                    Dim uri As String = source("uri")?.ToString()
                    If Not String.IsNullOrEmpty(uri) AndAlso sourceUris.Add(uri) Then
                        Dim title As String = source("title")?.ToString()
                        If String.IsNullOrWhiteSpace(title) Then title = "No title"
                        Dim authors As String = String.Join(", ", source.SelectTokens("authors[*].given").Select(Function(t) t.ToString()))
                        Dim doi As String = source("doi")?.ToString()
                        citationList.Add($"Source: {title} | Authors: {If(authors, "Unknown")} | DOI: {If(doi, "N/A")} | URL: {uri}")
                    End If
                Catch ex As Exception
                Debug.WriteLine("Error processing metadata source: " & ex.Message)
            End Try
            End Sub

            Private Shared Sub AddSource(source As JToken, ByRef citationList As List(Of String), ByRef sourceUris As HashSet(Of String))
                Try
                    Dim uri As String = source("uri")?.ToString()
                    If String.IsNullOrEmpty(uri) OrElse sourceUris.Contains(uri) Then Return

                    Dim sb As New StringBuilder()
                    sb.Append("Source: ")

                    ' Build title with container if available
                    Dim title As String = source("title")?.ToString()
                    Dim container As String = source("containerTitle")?.ToString()
                    If Not String.IsNullOrEmpty(container) Then
                        sb.Append($"{title}. In: {container}")
                    Else
                        sb.Append(title)
                    End If

                    ' Add authors
                    Dim authors = source.SelectTokens("authors[*]")
                    If authors IsNot Nothing AndAlso authors.Any() Then
                        sb.Append(" | Authors: ")
                        For Each author In authors
                            Dim given As String = author("given")?.ToString()
                            Dim family As String = author("family")?.ToString()
                            If Not String.IsNullOrEmpty(family) Then
                                sb.Append($"{family}, {given}; ")
                            End If
                        Next
                        If sb.Length > 2 Then
                            sb.Length -= 2 ' Remove last semicolon and space
                        End If
                    End If

                    ' Add publication info
                    Dim pubDate As String = source("publicationDate")?.ToString()
                    If Not String.IsNullOrEmpty(pubDate) Then
                        sb.Append($" | Published: {pubDate}")
                    End If

                    ' Add DOI if available
                    Dim doi As String = source("doi")?.ToString()
                    If Not String.IsNullOrEmpty(doi) Then
                        sb.Append($" | DOI: {doi}")
                    End If
                    sb.Append($" | URL: {uri}")

                    citationList.Add(sb.ToString())
                    sourceUris.Add(uri)
                Catch ex As Exception
                Debug.WriteLine("Error adding source: " & ex.Message)
            End Try
            End Sub

            Private Shared Sub ExtractLegacyCitations(jsonObj As JObject, ByRef citationList As List(Of String), ByRef sourceUris As HashSet(Of String))
                Try
                    ' Old format v0.9 compatibility: look for any "sources" with a URL.
                    Dim legacyCitations = jsonObj.SelectTokens("$..sources[?(@.url)]")
                    For Each legacySource In legacyCitations
                        Dim url As String = legacySource("url")?.ToString()
                        If Not String.IsNullOrEmpty(url) AndAlso sourceUris.Add(url) Then
                            citationList.Add($"Legacy source: {url}")
                        End If
                    Next
                Catch ex As Exception
                Debug.WriteLine("Error processing legacy citations: " & ex.Message)
            End Try
            End Sub

            Private Shared Function ExtractIeeeUri(refEntry As String) As String
                Try
                    Dim doiMatch = Regex.Match(refEntry, "doi:\s*(\S+)")
                    If doiMatch.Success Then
                        ' Trim any trailing punctuation
                        Return $"https://doi.org/{doiMatch.Groups(1).Value.TrimEnd("."c)}"
                    End If
                Catch ex As Exception
                    Debug.WriteLine("DOI extraction error: " & ex.Message)
                End Try
                Return String.Empty
            End Function

            Private Shared Function ExtractHarvardUri(refEntry As String) As String
                Try
                    Dim uriMatch = Regex.Match(refEntry, "Available at:\s*(\S+)\s*\(")
                    If uriMatch.Success Then
                        Return uriMatch.Groups(1).Value
                    End If
                Catch ex As Exception
                    Debug.WriteLine("Harvard URI extraction error: " & ex.Message)
                End Try
                Return String.Empty
            End Function

            Private Shared Function FormatCitations(citationList As List(Of String)) As String
                Dim sb As New StringBuilder()
                sb.AppendLine(vbCrLf & "References:")
                For i As Integer = 0 To citationList.Count - 1
                    sb.AppendLine($"[{i + 1}] {citationList(i)}")
                Next
                Return sb.ToString()
            End Function

            Private Shared Function ExtractSimpleCitations(ByRef jsonObj As JObject) As String
                Try
                    Dim citations As JToken = jsonObj.SelectToken("citations")
                    Dim citationList As New List(Of String)

                    If citations IsNot Nothing Then
                        If citations.Type = JTokenType.Array Then
                            For Each citation As JToken In citations
                                If citation.Type = JTokenType.String Then
                                    citationList.Add(citation.ToString())
                                ElseIf citation.Type = JTokenType.Object Then
                                    ' Try to extract URL or fullNote from the object
                                    Dim url As JToken = citation.SelectToken("url")
                                    If url IsNot Nothing Then
                                        citationList.Add(url.ToString())
                                    Else
                                        Dim fullNote As String = citation("fullNote")?.ToString()
                                        If Not String.IsNullOrEmpty(fullNote) Then
                                            citationList.Add(fullNote)
                                        End If
                                    End If
                                End If
                            Next
                        ElseIf citations.Type = JTokenType.Object Then
                            Dim fullNote As String = citations("fullNote")?.ToString()
                            If Not String.IsNullOrEmpty(fullNote) Then
                                citationList.Add(fullNote)
                            Else
                                Dim url As String = citations("url")?.ToString()
                                If Not String.IsNullOrEmpty(url) Then
                                    citationList.Add(url)
                                End If
                            End If
                        End If
                    End If

                    Dim simpleCitationOutput As New StringBuilder()
                    simpleCitationOutput.AppendLine(vbCrLf)
                    For i As Integer = 0 To citationList.Count - 1
                        simpleCitationOutput.AppendLine("[" & (i + 1).ToString() & "] " & citationList(i))
                    Next

                    Return simpleCitationOutput.ToString()

                Catch ex As Exception
                Debug.WriteLine("Error parsing JSON for simple citations: " & ex.Message)
            End Try

                Return String.Empty
            End Function



            Private Shared Function xxxExtractSimpleCitations(ByRef jsonObj As JObject) As String

            Try

                ' Extract citations
                Dim citations As JToken = jsonObj.SelectToken("citations")
                Dim citationList As New List(Of String)

                If citations IsNot Nothing Then
                    If citations.Type = JTokenType.Array Then
                        ' Check if the citations array contains simple strings or objects
                        For Each citation As JToken In citations
                            If citation.Type = JTokenType.String Then
                                ' Simple URL format
                                citationList.Add(citation.ToString())
                            ElseIf citation.Type = JTokenType.Object Then
                                ' Extract "url" from the object format
                                Dim url As JToken = citation.SelectToken("url")
                                If url IsNot Nothing Then
                                    citationList.Add(url.ToString())
                                End If
                            End If
                        Next
                    End If
                End If

                Dim simpleCitationOutput As String = ""

                ' Append citations if found
                If citationList.Count > 0 Then
                    simpleCitationOutput = vbCrLf & vbCrLf
                    For i As Integer = 0 To citationList.Count - 1
                        simpleCitationOutput &= "[" & (i + 1).ToString() & "] " & citationList(i) & vbCrLf
                    Next
                End If

                Return simpleCitationOutput

            Catch ex As Exception
                MessageBox.Show("Error parsing JSON for simple citations: " & ex.Message)
                Return String.Empty
            End Try
        End Function


        Public Shared Function FindJsonProperty(token As JToken, searchtext As String) As String
            If token.Type = JTokenType.Object Then
                For Each prop As JProperty In CType(token, JObject).Properties()
                    If prop.Name = searchtext Then
                        Return prop.Value.ToString()
                    End If
                    Dim result As String = FindJsonProperty(prop.Value, searchtext)
                    If Not String.IsNullOrEmpty(result) Then Return result
                Next
            ElseIf token.Type = JTokenType.Array Then
                For Each item As JToken In CType(token, JArray)
                    Dim result As String = FindJsonProperty(item, searchtext)
                    If Not String.IsNullOrEmpty(result) Then Return result
                Next
            End If
            Return Nothing
        End Function


        Public Shared Async Function GetFreshAccessToken(context As ISharedContext, ByVal clientEmail As String, ByVal ClientScopes As String, ByVal PrivateKey As String, ByVal AuthServer As String, ByVal TLife As Long, ByVal SecondAPI As Boolean) As Task(Of String)
            Try

                Dim accessToken As String = String.Empty
                Dim currentexpiry As DateTime
                If SecondAPI Then
                    accessToken = context.DecodedAPI_2
                    currentexpiry = context.TokenExpiry_2
                Else
                    accessToken = context.DecodedAPI
                    currentexpiry = context.TokenExpiry
                End If

                PrivateKey = PrivateKey.Replace("\n", "")

                Dim formattedKey As String = String.Empty

                For i As Integer = 0 To PrivateKey.Length - 1 Step 64
                    If i + 64 <= PrivateKey.Length Then
                        formattedKey &= PrivateKey.Substring(i, 64) & vbLf
                    Else
                        formattedKey &= PrivateKey.Substring(i) & vbLf
                    End If
                Next

                GoogleOAuthHelper.client_email = clientEmail
                GoogleOAuthHelper.private_key = "-----BEGIN PRIVATE KEY-----" & vbLf & formattedKey & "-----END PRIVATE KEY-----" & vbLf
                GoogleOAuthHelper.scopes = ClientScopes
                GoogleOAuthHelper.token_uri = AuthServer
                GoogleOAuthHelper.token_life = TLife

                If String.IsNullOrEmpty(accessToken) OrElse DateTime.UtcNow >= currentexpiry Then
                    ' Token is missing or expired, fetch a new one
                    accessToken = Await GoogleOAuthHelper.GetAccessToken()
                    If SecondAPI Then
                        context.TokenExpiry_2 = DateTime.UtcNow.AddSeconds(GoogleOAuthHelper.token_life - 300) ' Set expiry 5 minutes before actual
                    Else
                        context.TokenExpiry = DateTime.UtcNow.AddSeconds(GoogleOAuthHelper.token_life - 300) ' Set expiry 5 minutes before actual
                    End If
                End If
                Return accessToken
                Exit Function

            Catch ex As System.Exception
                ' Handle exceptions explicitly with System.Exception
                MessageBox.Show("Error while fetching an access token: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If SecondAPI Then
                    context.DecodedAPI_2 = String.Empty
                Else
                    context.DecodedAPI = String.Empty
                End If
                Return String.Empty
            End Try

        End Function


        Public Shared Function CleanString(ByVal input As String) As String
            Dim cleanedString As String = ""

            If Not IsEmptyOrBlank(input) Then

                For Each currentChar As Char In input
                    Dim charCode As Integer = AscW(currentChar)

                    Select Case charCode
                        Case 8
                            cleanedString &= "\b"
                        Case 9
                            cleanedString &= "\t"
                        Case 10
                            cleanedString &= "\n"
                        Case 12
                            cleanedString &= "\f"
                        Case 13
                            cleanedString &= "\n"  '\r
                        Case 34
                            cleanedString &= "\"""
                        Case 92
                            cleanedString &= "\\"
                        Case 0 To 31
                            cleanedString &= "\u" & charCode.ToString("X4")
                        Case Else
                            cleanedString &= currentChar
                    End Select
                Next

            End If

            ' Condense multiple spaces to a single space
            While cleanedString.Contains("  ")
                cleanedString = cleanedString.Replace("  ", " ")
            End While

            Return cleanedString
        End Function

        Public Shared Function ConvertEscapeCharacters(ByVal inputText As String) As String

            ' Handle basic escape sequences
            inputText = Regex.Replace(inputText, "(?<!\\)\\n\\n", Environment.NewLine & Environment.NewLine)
            inputText = Regex.Replace(inputText, "(?<!\\)\\n", vbLf)
            inputText = Regex.Replace(inputText, "(?<!\\)\\r", vbCr)
            inputText = Regex.Replace(inputText, "(?<!\\)\\t", vbTab)

            ' Convert Unicode escape sequences (e.g., \u000B)
            Dim unicodePattern As String = "\\u([0-9A-Fa-f]{4})"
            inputText = Regex.Replace(inputText, unicodePattern, Function(m)
                                                                     Dim unicodeValue As Integer = Convert.ToInt32(m.Groups(1).Value, 16)
                                                                     Return ChrW(unicodeValue)
                                                                 End Function)

            inputText = inputText.Replace("\\", "\")
            inputText = inputText.Replace("\""", """")

            Return inputText
        End Function


        Public Shared Function ExtractJSONValue(jsonString As String, objectName As String) As String

            Try

                Dim jsonObject As JObject = JObject.Parse(jsonString)
                Return FindJsonProperty(jsonObject, objectName)

                ' old code prior to using the library

                Dim searchKey As String = $"""{objectName}""" ' Enclose objectName in double quotes
                Dim keyPos As Integer = jsonString.IndexOf(searchKey)

                If keyPos = -1 Then
                    ' Key not found
                    Return ""
                End If

                ' Move past the key
                Dim pos As Integer = keyPos + searchKey.Length

                ' Skip any whitespace
                While pos < jsonString.Length AndAlso (jsonString(pos) = " "c OrElse jsonString(pos) = vbTab OrElse jsonString(pos) = vbCr OrElse jsonString(pos) = vbLf)
                    pos += 1
                End While

                ' Check for colon ':'
                If pos >= jsonString.Length OrElse jsonString(pos) <> ":"c Then
                    ' Invalid JSON format
                    Return ""
                End If

                pos += 1 ' Move past the ':'

                ' Skip any whitespace after the colon
                While pos < jsonString.Length AndAlso (jsonString(pos) = " "c OrElse jsonString(pos) = vbTab OrElse jsonString(pos) = vbCr OrElse jsonString(pos) = vbLf)
                    pos += 1
                End While

                ' Get the first character of the value
                If pos >= jsonString.Length Then Return ""
                Dim valueStartChar As Char = jsonString(pos)

                Select Case valueStartChar
                    Case """"c
                        ' String value
                        Dim valueStartPos As Integer = pos + 1 ' Skip the opening quote
                        Dim valueEndPos As Integer = valueStartPos
                        While valueEndPos < jsonString.Length
                            If jsonString(valueEndPos) = """"c Then
                                ' Check if the quote is escaped
                                Dim backslashCount As Integer = 0
                                Dim tempPos As Integer = valueEndPos - 1
                                While tempPos >= valueStartPos AndAlso jsonString(tempPos) = "\"c
                                    backslashCount += 1
                                    tempPos -= 1
                                End While
                                If backslashCount Mod 2 = 0 Then
                                    ' Even number of backslashes, so the quote is not escaped
                                    Exit While
                                End If
                            End If
                            valueEndPos += 1
                        End While
                        ' Extract the string value
                        Dim valueString As String = jsonString.Substring(valueStartPos, valueEndPos - valueStartPos)
                        ' Replace escaped characters
                        ' valueString = valueString.Replace("\""", """").Replace("\\", "\")
                        Return valueString

                    Case "{"c
                        ' Object value
                        Dim valueStartPos As Integer = pos
                        Dim braceCount As Integer = 1
                        Dim valueEndPos As Integer = pos + 1
                        While valueEndPos < jsonString.Length AndAlso braceCount > 0
                            Dim c As Char = jsonString(valueEndPos)
                            If c = "{"c Then
                                braceCount += 1
                            ElseIf c = "}"c Then
                                braceCount -= 1
                            ElseIf c = """"c Then
                                ' Skip strings inside the object
                                valueEndPos += 1
                                While valueEndPos < jsonString.Length
                                    If jsonString(valueEndPos) = """"c Then
                                        ' Check if the quote is escaped
                                        Dim backslashCount As Integer = 0
                                        Dim tempPos As Integer = valueEndPos - 1
                                        While tempPos >= valueStartPos AndAlso jsonString(tempPos) = "\"c
                                            backslashCount += 1
                                            tempPos -= 1
                                        End While
                                        If backslashCount Mod 2 = 0 Then Exit While
                                    End If
                                    valueEndPos += 1
                                End While
                            End If
                            valueEndPos += 1
                        End While
                        Dim valueString As String = jsonString.Substring(valueStartPos, valueEndPos - valueStartPos)
                        Return valueString

                    Case "["c
                        ' Array value
                        Dim valueStartPos As Integer = pos
                        Dim bracketCount As Integer = 1
                        Dim valueEndPos As Integer = pos + 1
                        While valueEndPos < jsonString.Length AndAlso bracketCount > 0
                            Dim c As Char = jsonString(valueEndPos)
                            If c = "["c Then
                                bracketCount += 1
                            ElseIf c = "]"c Then
                                bracketCount -= 1
                            ElseIf c = """"c Then
                                ' Skip strings inside the array
                                valueEndPos += 1
                                While valueEndPos < jsonString.Length
                                    If jsonString(valueEndPos) = """"c Then
                                        Dim backslashCount As Integer = 0
                                        Dim tempPos As Integer = valueEndPos - 1
                                        While tempPos >= valueStartPos AndAlso jsonString(tempPos) = "\"c
                                            backslashCount += 1
                                            tempPos -= 1
                                        End While
                                        If backslashCount Mod 2 = 0 Then Exit While
                                    End If
                                    valueEndPos += 1
                                End While
                            End If
                            valueEndPos += 1
                        End While
                        Dim valueString As String = jsonString.Substring(valueStartPos, valueEndPos - valueStartPos)
                        Return valueString

                    Case "t"c
                        ' Check for "true"
                        If jsonString.Substring(pos, Math.Min(4, jsonString.Length - pos)) = "true" Then
                            Return "true"
                        End If

                    Case "f"c
                        ' Check for "false"
                        If jsonString.Substring(pos, Math.Min(5, jsonString.Length - pos)) = "false" Then
                            Return "false"
                        End If

                    Case "n"c
                        ' Check for "null"
                        If jsonString.Substring(pos, Math.Min(4, jsonString.Length - pos)) = "null" Then
                            Return "null"
                        End If

                    Case Else
                        ' Number value
                        Dim valueStartPos As Integer = pos
                        Dim valueEndPos As Integer = pos
                        While valueEndPos < jsonString.Length
                            Dim c As Char = jsonString(valueEndPos)
                            If c = ","c OrElse c = "}"c OrElse c = "]"c OrElse c = vbCr OrElse c = vbLf OrElse c = " "c OrElse c = vbTab Then
                                Exit While
                            End If
                            valueEndPos += 1
                        End While
                        Dim valueString As String = jsonString.Substring(valueStartPos, valueEndPos - valueStartPos)
                        Return valueString
                End Select

                ' If none of the cases match
                Return ""

            Catch ex As System.Exception
                MessageBox.Show($"Error in ExtractJSONValue: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return ""
            End Try
        End Function

        Public Shared Function CountWords(ByVal selectedText As String) As Integer
            Try
                ' Trim leading and trailing whitespace
                selectedText = selectedText.Trim()

                ' If the text is empty, return 0
                If String.IsNullOrWhiteSpace(selectedText) Then
                    Return 0
                End If

                ' Split the text into words using spaces
                Dim words() As String = selectedText.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)

                ' Count non-empty words
                Dim wordCount As Integer = words.Count(Function(word) Not String.IsNullOrWhiteSpace(word))

                Return wordCount
            Catch ex As Exception
                ' In case of an error, return a default value of 1
                Return 1
            End Try
        End Function

        Public Shared Sub InitializeConfig(ByRef context As ISharedContext, FirstTime As Boolean, Reload As Boolean)

            If context.INIloaded And Not Reload Then Exit Sub

            context.GPTSetupError = True

            context.INIloaded = False

            Dim IniFilePath As String = ""
            Dim RegFilePath As String = ""
            Dim DefaultPath As String = ""
            Dim DefaultPath2 As String = ""

            Try

                ' Determine the configuration file path

                RegFilePath = GetFromRegistry(RegPath_Base, RegPath_IniPath, True)
                DefaultPath = GetDefaultINIPath(context.RDV)
                DefaultPath2 = GetDefaultINIPath("Word")

                If Not String.IsNullOrWhiteSpace(RegFilePath) And RegPath_IniPrio Then
                    IniFilePath = System.IO.Path.Combine(ExpandEnvironmentVariables(RegFilePath), $"{AN2}.ini")
                ElseIf System.IO.File.Exists(DefaultPath) Then
                    IniFilePath = DefaultPath
                ElseIf System.IO.File.Exists(DefaultPath2) Then
                    IniFilePath = DefaultPath2
                ElseIf Not String.IsNullOrWhiteSpace(RegFilePath) Then
                    IniFilePath = System.IO.Path.Combine(ExpandEnvironmentVariables(RegFilePath), $"{AN2}.ini")
                Else
                    IniFilePath = DefaultPath
                End If

                IniFilePath = RemoveCR(IniFilePath)

                ' Check if the configuration file exists

                If Not System.IO.File.Exists(IniFilePath) Then
                    If FirstTime Then
                        Using frm As New InitialConfig(context)
                            frm.ShowDialog()
                        End Using
                        IniFilePath = DefaultPath
                        If context.InitialConfigFailed And Not System.IO.File.Exists(IniFilePath) Then
                            ShowCustomMessageBox($"You have aborted the setup wizard and no configuration file has been found ('{IniFilePath}'). You will have to retry or configure it manually to use {AN}, even if you see the menus (they will disappear once {AN} has been de-installed or de-activated).")
                            Exit Sub
                        End If
                        If Not System.IO.File.Exists(IniFilePath) Then
                            ShowCustomMessageBox($"The configuration file is (still) not found ('{IniFilePath}'). There may be an error in the setup assistant. Please configure the configuration file manually.")
                            Exit Sub
                        End If
                    Else
                        ShowCustomMessageBox($"The configuration file has not been found ('{IniFilePath}').")
                        Exit Sub
                    End If
                End If

                Dim iniContent As String = ""
                Dim configDict As New Dictionary(Of String, String)

                ' Read and parse the .ini file
                iniContent = System.IO.File.ReadAllText(IniFilePath)
                Dim iniLines As String() = iniContent.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                For Each line As String In iniLines
                    Dim trimmedLine = line.Trim()
                    If Not String.IsNullOrEmpty(trimmedLine) AndAlso Not trimmedLine.StartsWith(";") Then ' Skip comments and empty lines
                        Dim keyValue = trimmedLine.Split(New Char() {"="c}, 2)
                        If keyValue.Length = 2 Then
                            configDict(keyValue(0).Trim()) = keyValue(1).Trim()
                        End If
                    End If
                Next

                ' Assign and validate configuration values
                context.INI_APIKey = If(configDict.ContainsKey("APIKey"), configDict("APIKey"), "")
                context.INI_Endpoint = If(configDict.ContainsKey("Endpoint"), configDict("Endpoint"), "")
                context.INI_HeaderA = If(configDict.ContainsKey("HeaderA"), configDict("HeaderA"), "")
                context.INI_HeaderB = If(configDict.ContainsKey("HeaderB"), configDict("HeaderB"), "")
                context.INI_Response = If(configDict.ContainsKey("Response"), configDict("Response"), "")
                context.INI_APICall = If(configDict.ContainsKey("APICall"), configDict("APICall"), "")
                context.INI_Timeout = If(configDict.ContainsKey("Timeout"), CLng(configDict("Timeout")), 0)
                context.INI_MaxOutputToken = If(configDict.ContainsKey("MaxOutputToken"), CInt(configDict("MaxOutputToken")), 0)
                context.INI_Temperature = If(configDict.ContainsKey("Temperature"), configDict("Temperature"), "")
                context.INI_Model = If(configDict.ContainsKey("Model"), configDict("Model"), "")

                context.SP_Translate = If(configDict.ContainsKey("SP_Translate"), configDict("SP_Translate"), Default_SP_Translate)
                context.SP_Correct = If(configDict.ContainsKey("SP_Correct"), configDict("SP_Correct"), Default_SP_Correct)
                context.SP_Improve = If(configDict.ContainsKey("SP_Improve"), configDict("SP_Improve"), Default_SP_Improve)
                context.SP_Explain = If(configDict.ContainsKey("SP_Explain"), configDict("SP_Explain"), Default_SP_Explain)
                context.SP_SuggestTitles = If(configDict.ContainsKey("SP_SuggestTitles"), configDict("SP_SuggestTitles"), Default_SP_SuggestTitles)
                context.SP_Friendly = If(configDict.ContainsKey("SP_Friendly"), configDict("SP_Friendly"), Default_SP_Friendly)
                context.SP_Convincing = If(configDict.ContainsKey("SP_Convincing"), configDict("SP_Convincing"), Default_SP_Convincing)
                context.SP_NoFillers = If(configDict.ContainsKey("SP_NoFillers"), configDict("SP_NoFillers"), Default_SP_NoFillers)
                context.SP_Podcast = If(configDict.ContainsKey("SP_Podcast"), configDict("SP_Podcast"), Default_SP_Podcast)
                context.SP_Shorten = If(configDict.ContainsKey("SP_Shorten"), configDict("SP_Shorten"), Default_SP_Shorten)
                context.SP_Summarize = If(configDict.ContainsKey("SP_Summarize"), configDict("SP_Summarize"), Default_SP_Summarize)
                context.SP_FreestyleText = If(configDict.ContainsKey("SP_FreestyleText"), configDict("SP_FreestyleText"), Default_SP_FreestyleText)
                context.SP_FreestyleNoText = If(configDict.ContainsKey("SP_FreestyleNoText"), configDict("SP_FreestyleNoText"), Default_SP_FreestyleNoText)
                context.SP_MailReply = If(configDict.ContainsKey("SP_MailReply"), configDict("SP_MailReply"), Default_SP_MailReply)
                context.SP_MailSumup = If(configDict.ContainsKey("SP_MailSumup"), configDict("SP_MailSumup"), Default_SP_MailSumup)
                context.SP_SwitchParty = If(configDict.ContainsKey("SP_SwitchParty"), configDict("SP_SwitchParty"), Default_SP_SwitchParty)
                context.SP_Anonymize = If(configDict.ContainsKey("SP_Anonymize"), configDict("SP_Anonymize"), Default_SP_Anonymize)
                context.SP_ContextSearch = If(configDict.ContainsKey("SP_ContextSearch"), configDict("SP_ContextSearch"), Default_SP_ContextSearch)
                context.SP_ContextSearchMulti = If(configDict.ContainsKey("SP_ContextSearchMulti"), configDict("SP_ContextSearchMulti"), Default_SP_ContextSearchMulti)
                context.SP_RangeOfCells = If(configDict.ContainsKey("SP_RangeOfCells"), configDict("SP_RangeOfCells"), Default_SP_RangeOfCells)
                context.SP_WriteNeatly = If(configDict.ContainsKey("SP_WriteNeatly"), configDict("SP_WriteNeatly"), Default_SP_WriteNeatly)
                context.SP_Add_KeepFormulasIntact = If(configDict.ContainsKey("SP_Add_KeepFormulasIntact"), configDict("SP_Add_KeepFormulasIntact"), Default_SP_Add_KeepFormulasIntact)
                context.SP_Add_KeepHTMLIntact = If(configDict.ContainsKey("SP_Add_KeepHTMLIntact"), configDict("SP_Add_KeepHTMLIntact"), Default_SP_Add_KeepHTMLIntact)
                context.SP_Add_KeepInlineIntact = If(configDict.ContainsKey("SP_Add_KeepInlineIntact"), configDict("SP_Add_KeepInlineIntact"), Default_SP_Add_KeepInlineIntact)
                context.SP_Add_Bubbles = If(configDict.ContainsKey("SP_Add_Bubbles"), configDict("SP_Add_Bubbles"), Default_SP_Add_Bubbles)
                context.SP_Add_Revisions = If(configDict.ContainsKey("SP_Add_Revisions"), configDict("SP_Add_Revisions"), Default_SP_Add_Revisions)
                context.SP_MarkupRegex = If(configDict.ContainsKey("SP_MarkupRegex"), configDict("SP_MarkupRegex"), Default_SP_MarkupRegex)
                context.SP_ChatWord = If(configDict.ContainsKey("SP_ChatWord"), configDict("SP_ChatWord"), Default_SP_ChatWord)
                context.SP_Add_ChatWord_Commands = If(configDict.ContainsKey("SP_Add_ChatWord_Commands"), configDict("SP_Add_ChatWord_Commands"), Default_SP_Add_ChatWord_Commands)


                ' Required For Excel Helper
                context.INI_OpenSSLPath = If(configDict.ContainsKey("OpenSSLPath"), configDict("OpenSSLPath"), "%APPDATA%\Microsoft\OpenSSL_Runtime\openssl.exe")

                ' Optional values
                context.INI_PreCorrection = If(configDict.ContainsKey("PreCorrection"), configDict("PreCorrection"), "")
                context.INI_PostCorrection = If(configDict.ContainsKey("PostCorrection"), configDict("PostCorrection"), "")
                context.INI_APIKeyPrefix = If(configDict.ContainsKey("APIKeyPrefix"), configDict("APIKeyPrefix"), "")
                context.INI_UsageRestrictions = If(configDict.ContainsKey("UsageRestrictions"), configDict("UsageRestrictions"), "")
                context.INI_Language1 = If(configDict.ContainsKey("Language1"), configDict("Language1"), "English")
                context.INI_Language2 = If(configDict.ContainsKey("Language2"), configDict("Language2"), "German")
                context.INI_KeepFormatCap = If(configDict.ContainsKey("KeepFormatCap"), CInt(configDict("KeepFormatCap")), 5000)
                context.INI_MarkupMethodHelper = If(configDict.ContainsKey("MarkupMethodHelper"), CInt(configDict("MarkupMethodHelper")), 3)
                context.INI_MarkupMethodWord = If(configDict.ContainsKey("MarkupMethodWord"), CInt(configDict("MarkupMethodWord")), 3)
                context.INI_MarkupMethodOutlook = If(configDict.ContainsKey("MarkupMethodWord"), CInt(configDict("MarkupMethodWord")), 3)
                context.INI_MarkupDiffCap = If(configDict.ContainsKey("MarkupDiffCap"), CInt(configDict("MarkupDiffCap")), 3000)
                context.INI_MarkupRegexCap = If(configDict.ContainsKey("MarkupRegexCap"), CInt(configDict("MarkupRegexCap")), 30000)
                context.INI_ChatCap = If(configDict.ContainsKey("ChatCap"), CInt(configDict("ChatCap")), 50000)

                ' Boolean parameters
                context.INI_DoubleS = ParseBoolean(configDict, "DoubleS")
                context.INI_KeepFormat1 = ParseBoolean(configDict, "KeepFormat1")
                context.INI_ReplaceText1 = ParseBoolean(configDict, "ReplaceText1", True)
                context.INI_KeepFormat2 = ParseBoolean(configDict, "KeepFormat2")
                context.INI_KeepParaFormatInline = ParseBoolean(configDict, "KeepParaFormatInline")
                context.INI_ReplaceText2 = ParseBoolean(configDict, "ReplaceText2")
                context.INI_DoMarkupOutlook = ParseBoolean(configDict, "DoMarkupOutlook", True)
                context.INI_DoMarkupWord = ParseBoolean(configDict, "DoMarkupWord", True)
                context.INI_RoastMe = ParseBoolean(configDict, "RoastMe", False)
                context.INI_APIDebug = ParseBoolean(configDict, "APIDebug")
                context.INI_APIEncrypted = ParseBoolean(configDict, "APIKeyEncrypted")

                context.INI_ShortcutsWordExcel = If(configDict.ContainsKey("ShortcutsWordExcel"), configDict("ShortcutsWordExcel"), "")
                context.INI_ContextMenu = ParseBoolean(configDict, "ContextMenu", True)

                ' Other parameters

                context.INI_UpdateCheckInterval = If(configDict.ContainsKey("UpdateCheckInterval"), CInt(configDict("UpdateCheckInterval")), 0)
                context.INI_UpdatePath = If(configDict.ContainsKey("UpdatePath"), configDict("UpdatePath"), "")
                context.INI_SpeechModelPath = If(configDict.ContainsKey("SpeechModelPath"), configDict("SpeechModelPath"), "")
                context.INI_TTSEndpoint = If(configDict.ContainsKey("TTSEndpoint"), configDict("TTSEndpoint"), "")

                context.INI_PromptLibPath = If(configDict.ContainsKey("PromptLib"), configDict("PromptLib"), "")
                context.INI_PromptLibPath_Transcript = If(configDict.ContainsKey("PromptLib_Transcript"), configDict("PromptLib_Transcript"), "")

                ' Process Internet search if enabled
                context.INI_ISearch = ParseBoolean(configDict, "ISearch", True)
                If context.INI_ISearch Then
                    context.INI_ISearch_Approve = ParseBoolean(configDict, "ISearch_Approve", False)
                    context.INI_ISearch_URL = If(configDict.ContainsKey("ISearch_URL"), configDict("ISearch_URL"), "https://duckduckgo.com/html/?q=")
                    context.INI_ISearch_ResponseMask1 = If(configDict.ContainsKey("ISearch_ResponseMask1"), configDict("ISearch_ResponseMask1"), "duckduckgo.com/l/?uddg=")
                    context.INI_ISearch_ResponseMask2 = If(configDict.ContainsKey("ISearch_ResponseMask2"), configDict("ISearch_ResponseMask2"), "&")
                    context.INI_ISearch_Name = If(configDict.ContainsKey("ISearch_Name"), configDict("ISearch_Name"), "DuckDuckGo")
                    context.INI_ISearch_Tries = If(configDict.ContainsKey("ISearch_Tries"), CInt(configDict("ISearch_Tries")), ISearch_DefTries)
                    context.INI_ISearch_Results = If(configDict.ContainsKey("ISearch_Results"), CInt(configDict("ISearch_Results")), ISearch_DefResults)
                    context.INI_ISearch_MaxDepth = If(configDict.ContainsKey("ISearch_MaxDepth"), CInt(configDict("ISearch_MaxDepth")), ISearch_DefMaxDepth)
                    context.INI_ISearch_Timeout = If(configDict.ContainsKey("ISearch_Timeout"), CLng(configDict("ISearch_Timeout")), ISearch_DefSearchTimeout)
                    context.INI_ISearch_SearchTerm_SP = If(configDict.ContainsKey("ISearch_SearchTerm_SP"), configDict("ISearch_SearchTerm_SP"), Default_INI_ISearch_SearchTerm_SP)
                    context.INI_ISearch_Apply_SP = If(configDict.ContainsKey("ISearch_Apply_SP"), configDict("ISearch_Apply_SP"), Default_INI_ISearch_Apply_SP)
                    context.INI_ISearch_Apply_SP_Markup = If(configDict.ContainsKey("ISearch_Apply_SP_Markup"), configDict("ISearch_Apply_SP_Markup"), Default_INI_ISearch_Apply_SP_Markup)
                    If context.INI_ISearch_Tries > ISearch_MaxTries Then context.INI_ISearch_Tries = ISearch_MaxTries
                    If context.INI_ISearch_Results > ISearch_MaxResults Then context.INI_ISearch_Results = ISearch_MaxResults
                    If context.INI_ISearch_MaxDepth > ISearch_MaxMaxDepth Then context.INI_ISearch_MaxDepth = ISearch_MaxMaxDepth
                    If context.INI_ISearch_Timeout > ISearch_MaxSearchTimeout Then context.INI_ISearch_Timeout = ISearch_MaxSearchTimeout
                    If context.INI_ISearch_Results > ISearch_MaxResults Then context.INI_ISearch_Results = ISearch_MaxResults
                End If

                ' Process RAG if enabled
                context.INI_Lib = ParseBoolean(configDict, "Lib")
                If context.INI_Lib Then
                    context.INI_Lib_File = If(configDict.ContainsKey("Lib_File"), configDict("Lib_File"), "")
                    context.INI_Lib_Timeout = If(configDict.ContainsKey("Lib_Timeout"), CLng(configDict("Lib_Timeout")), 60000)
                    context.INI_Lib_Find_SP = If(configDict.ContainsKey("Lib_Find_SP"), configDict("Lib_Find_SP"), "You are a legal professional with very good legal, language and logical skills and text handling capabilities, and you precisely comply with any instructions step by step. You are given an instruction from the user: {OtherPrompt}. If present, the user also provides text between <TEXTTOPROCESS> and </TEXTTOPROCESS>. A library of text elements is included between <LIBRARY> and </LIBRARY>, with each element separated by the string '@@@'. Identify and return only those library elements that are directly applicable to the user’s instruction. If multiple elements apply, separate them with '---'. If no elements apply, return an empty result. Return only the applicable library text elements, without any commentary or explanation.\n<LIBRARY>{LibraryText}</LIBRARY>")
                    context.INI_Lib_Apply_SP = If(configDict.ContainsKey("Lib_Apply_SP"), configDict("Lib_Apply_SP"), "You are a legal professional with very good legal, language and logical skills and text handling capabilities, and you precisely comply with any instructions step by step. You have the following instruction: {OtherPrompt}. You have the relevant library text between <LIBRESULT> and </LIBRESULT>. (If multiple library elements apply, they are separated by '---'.) The user’s text, if any, appears between <TEXTTOPROCESS> and </TEXTTOPROCESS>. Use the library elements intelligently to comply with the user’s instruction, such as drafting or improving a clause. If no text is provided, create a suitable text from scratch, relying on the library elements and the instruction. Present a clean, final version of the text without markup or extra commentary.\n<LIBRESULT>{LibResult}</LIBRESULT>")
                    context.INI_Lib_Apply_SP_Markup = If(configDict.ContainsKey("Lib_Apply_SP_Markup"), configDict("Lib_Apply_SP_Markup"), "You are a legal professional with very good legal, language and logical skills and text handling capabilities, and you precisely comply with any instructions step by step. You have the following instruction: {OtherPrompt}. The user-provided text appears between <TEXTTOPROCESS> and </TEXTTOPROCESS>. The relevant library text is between <LIBRESULT> and </LIBRESULT>. (If multiple library elements apply, they are separated by '---'.) Use these library elements intelligently to fulfill the user’s instruction, such as drafting or modifying a clause based on a sample in the library. Fullfill the instruction by amending or expanding the portions of the user’s text that (for example, by including a clause or amending an existing one), always take into account the full text provided by the user and only use the additional information from the library and the instruction. Provide your final text without additional commentary.\n<LIBRESULT>{LibResult}</LIBRESULT>")

                End If

                ' Process SecondAPI configuration if enabled
                context.INI_Endpoint_2 = "" ' necessary for googleapi check (should not be null)
                context.INI_SecondAPI = ParseBoolean(configDict, "SecondAPI")
                If context.INI_SecondAPI Then
                    context.INI_APIKey_2 = If(configDict.ContainsKey("APIKey_2"), configDict("APIKey_2"), "")
                    context.INI_Endpoint_2 = If(configDict.ContainsKey("Endpoint_2"), configDict("Endpoint_2"), "")
                    context.INI_HeaderA_2 = If(configDict.ContainsKey("HeaderA_2"), configDict("HeaderA_2"), "")
                    context.INI_HeaderB_2 = If(configDict.ContainsKey("HeaderB_2"), configDict("HeaderB_2"), "")
                    context.INI_Response_2 = If(configDict.ContainsKey("Response_2"), configDict("Response_2"), "")
                    context.INI_APICall_2 = If(configDict.ContainsKey("APICall_2"), configDict("APICall_2"), "")
                    context.INI_Timeout_2 = If(configDict.ContainsKey("Timeout_2"), CLng(configDict("Timeout_2")), 0)
                    context.INI_MaxOutputToken_2 = If(configDict.ContainsKey("MaxOutputToken_2"), CInt(configDict("MaxOutputToken_2")), 0)
                    context.INI_Temperature_2 = If(configDict.ContainsKey("Temperature_2"), configDict("Temperature_2"), "")
                    context.INI_Model_2 = If(configDict.ContainsKey("Model_2"), configDict("Model_2"), "")
                    context.INI_APIEncrypted_2 = ParseBoolean(configDict, "APIKeyEncrypted_2")
                    context.INI_APIKeyPrefix_2 = If(configDict.ContainsKey("APIKeyPrefix_2"), configDict("APIKeyPrefix_2"), "")
                End If

                ' Process OAuth2 configuration if enabled
                context.INI_OAuth2 = ParseBoolean(configDict, "OAuth2")
                If context.INI_OAuth2 Then
                    context.INI_OAuth2ClientMail = If(configDict.ContainsKey("OAuth2ClientMail"), configDict("OAuth2ClientMail"), "")
                    context.INI_OAuth2Scopes = If(configDict.ContainsKey("OAuth2Scopes"), configDict("OAuth2Scopes"), "")
                    context.INI_OAuth2Endpoint = If(configDict.ContainsKey("OAuth2Endpoint"), configDict("OAuth2Endpoint"), "")
                    context.INI_OAuth2ATExpiry = If(configDict.ContainsKey("OAuth2ATExpiry"), CLng(configDict("OAuth2ATExpiry")), 3600)

                End If

                If context.INI_SecondAPI Then
                    context.INI_OAuth2_2 = ParseBoolean(configDict, "OAuth2_2")
                    If context.INI_OAuth2_2 Then
                        context.INI_OAuth2ClientMail_2 = If(configDict.ContainsKey("OAuth2ClientMail_2"), configDict("OAuth2ClientMail_2"), "")
                        context.INI_OAuth2Scopes_2 = If(configDict.ContainsKey("OAuth2Scopes_2"), configDict("OAuth2Scopes_2"), "")
                        context.INI_OAuth2Endpoint_2 = If(configDict.ContainsKey("OAuth2Endpoint_2"), configDict("OAuth2Endpoint_2"), "")
                        context.INI_OAuth2ATExpiry_2 = If(configDict.ContainsKey("OAuth2ATExpiry_2"), CLng(configDict("OAuth2ATExpiry_2")), 3600)
                    End If
                End If

                If context.INI_APIEncrypted Or context.INI_APIEncrypted_2 Then
                    If IsEmptyOrBlank(Int_CodeBasis) Then
                        context.Codebasis = GetFromRegistry(RegPath_Base, RegPath_CodeBasis, False)
                    Else
                        context.Codebasis = Int_CodeBasis
                    End If
                End If

                context.INI_APIKeyBack = context.INI_APIKey
                context.INI_APIKeyBack_2 = context.INI_APIKey_2

                LicensedTill = If(configDict.ContainsKey("LicensedTill"), CDate(configDict("LicensedTill")), MaxUseDate)

                If DateTime.Now.AddDays(+7) > LicensedTill Then
                    ShowCustomMessageBox($"Your configured license for {AN} for {context.RDV} will expire in {DateDiff(DateInterval.Day, DateTime.Now, LicensedTill) + 1} days. Please contact your administrator or update the configuration file.", AN, 10)
                End If

                If Now > LicensedTill Then
                    ShowCustomMessageBox($"Your configured license for {AN} for {context.RDV} has expired. Please renew and configure the license to continue using {AN}.")
                    Exit Sub
                End If

                If INIValuesMissing(context) Then
                    Exit Sub
                End If

                ' Additional configurations for OAuth2
                context.TokenExpiry = DateAdd(DateInterval.Year, -1, DateTime.Now)
                context.DecodedAPI = ""
                context.INI_APIKeyBack = context.INI_APIKey

                context.TokenExpiry_2 = DateAdd(DateInterval.Year, -1, DateTime.Now)
                context.DecodedAPI_2 = ""
                context.INI_APIKeyBack_2 = context.INI_APIKey_2

                ' Set PromptLib if Path is configured
                If context.INI_PromptLibPath = "" Then context.INI_PromptLib = False Else context.INI_PromptLib = True

                ' Check and decrypt API keys
                If context.INI_OAuth2 Then
                    context.INI_APIKey = Trim(Replace(RealAPIKey(context.INI_APIKey, False, True, context), "\n", ""))
                    If String.IsNullOrWhiteSpace(context.INI_APIKey) Then
                        ShowCustomMessageBox("Internal error: Could not determine private key (likely a decryption error).")
                        Exit Sub
                    End If
                Else
                    context.DecodedAPI = RealAPIKey(context.INI_APIKey, False, False, context)
                    If String.IsNullOrWhiteSpace(context.DecodedAPI) Then
                        ShowCustomMessageBox("Internal error: Could not determine API key (likely a decryption error).")
                        Exit Sub
                    End If
                End If

                ' Decrypt second API keys
                If context.INI_SecondAPI Then
                    If context.INI_OAuth2_2 Then
                        context.INI_APIKey_2 = Trim(Replace(RealAPIKey(context.INI_APIKey_2, True, True, context), "\n", ""))
                        If String.IsNullOrWhiteSpace(context.INI_APIKey_2) Then
                            ShowCustomMessageBox("Internal error: Could not determine private key (likely a decryption error).")
                            Exit Sub
                        End If
                    Else
                        context.DecodedAPI_2 = RealAPIKey(context.INI_APIKey_2, True, False, context)
                        If String.IsNullOrWhiteSpace(context.DecodedAPI_2) Then
                            MessageBox.Show("Internal error: Could not determine API key for second API (likely a decryption error).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                    End If
                End If

                context.GPTSetupError = False
                context.INIloaded = True

            Catch ex As System.Exception
                MessageBox.Show($"Error in InitializeConfig: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ' Helper function to parse boolean values
        Public Shared Function ParseBoolean(configDict As Dictionary(Of String, String), key As String, Optional defaultvalue As Boolean = False) As Boolean
            If configDict.ContainsKey(key) Then
                Dim value = configDict(key).Trim().ToLower()
                Return value = "yes" OrElse value = "true" OrElse value = "ja" OrElse value = "wahr"
            End If
            Return defaultvalue
        End Function



        Public Shared Function SignJWT(jwtUnsigned As String, privateKeyPem As String) As String
            Try

                Dim privateKey As AsymmetricCipherKeyPair
                If Left(privateKeyPem, 3) <> "---" Then
                    privateKeyPem = "-----BEGIN PRIVATE KEY-----" & vbLf & ConvertToPemFormat(privateKeyPem) & vbLf & "-----END PRIVATE KEY-----"
                End If

                ' Read the private key properly
                Dim privateKeyObject As Object
                Using reader As New StringReader(privateKeyPem)
                    Dim pemReader = New Org.BouncyCastle.OpenSsl.PemReader(reader)
                    privateKeyObject = pemReader.ReadObject()
                End Using

                ' Determine if we have a key pair or just a private key
                Dim privateKeyParams As RsaKeyParameters
                If TypeOf privateKeyObject Is AsymmetricCipherKeyPair Then
                    Dim keyPair = CType(privateKeyObject, AsymmetricCipherKeyPair)
                    privateKeyParams = CType(keyPair.Private, RsaKeyParameters)
                ElseIf TypeOf privateKeyObject Is RsaPrivateCrtKeyParameters Then
                    privateKeyParams = CType(privateKeyObject, RsaPrivateCrtKeyParameters)
                Else
                    Throw New ApplicationException("Invalid private key format.")
                End If

                ' Convert unsigned JWT to bytes
                Dim unsignedDataBytes = Encoding.UTF8.GetBytes(jwtUnsigned)

                ' Use SHA256 for the signature
                Dim signer = SignerUtilities.GetSigner("SHA256withRSA")
                signer.Init(True, privateKeyParams)
                signer.BlockUpdate(unsignedDataBytes, 0, unsignedDataBytes.Length)
                Dim signatureBytes = signer.GenerateSignature()

                ' Base64 encode the signature
                Dim base64Signature = Convert.ToBase64String(signatureBytes)

                Return base64Signature
            Catch ex As Exception
                Throw New ApplicationException("Error signing JWT: " & ex.Message, ex)
            End Try
        End Function

        Private Shared Function ConvertToPemFormat(rawKey As String) As String
            Dim sb As New StringBuilder()
            Dim index As Integer = 0
            While index < rawKey.Length
                Dim chunk As String = rawKey.Substring(index, Math.Min(64, rawKey.Length - index))
                sb.AppendLine(chunk)
                index += 64
            End While
            Return sb.ToString().Trim()
        End Function



        Public Shared Function INIValuesMissing(ByVal context As ISharedContext) As Boolean
            Dim missingSettings As New Dictionary(Of String, String)
            Dim usercompleted As Boolean = False

            Do

                missingSettings.Clear()

                ' Check for missing values
                If String.IsNullOrEmpty(context.INI_APIKey) Then missingSettings.Add("APIKey", "APIKey (Model 1)")
                If String.IsNullOrEmpty(context.INI_Temperature) Then missingSettings.Add("Temperature", "Temperature (Model 1)")
                If context.INI_Timeout = 0 Then missingSettings.Add("Timeout", "Timeout (Model 1)")
                If String.IsNullOrEmpty(context.INI_Model) Then missingSettings.Add("Model", "Model (Model 1)")
                If String.IsNullOrEmpty(context.INI_Endpoint) Then missingSettings.Add("Endpoint", "Endpoint (Model 1)")
                If String.IsNullOrEmpty(context.INI_APICall) Then missingSettings.Add("APICall", "APICall (Model 1)")
                If String.IsNullOrEmpty(context.INI_Response) Then missingSettings.Add("Response", "Response (Model 1)")

                If context.INI_SecondAPI Then
                    If String.IsNullOrEmpty(context.INI_APIKey_2) Then missingSettings.Add("APIKey_2", "APIKey (Model 2)")
                    If String.IsNullOrEmpty(context.INI_Temperature_2) Then missingSettings.Add("Temperature_2", "Temperature (Model 2)")
                    If context.INI_Timeout_2 = 0 Then missingSettings.Add("Timeout_2", "Timeout (Model 2)")
                    If String.IsNullOrEmpty(context.INI_Model_2) Then missingSettings.Add("Model_2", "Model (Model 2)")
                    If String.IsNullOrEmpty(context.INI_Endpoint_2) Then missingSettings.Add("Endpoint_2", "Endpoint (Model 2)")
                    If String.IsNullOrEmpty(context.INI_APICall_2) Then missingSettings.Add("APICall_2", "APICall (Model 2)")
                    If String.IsNullOrEmpty(context.INI_Response_2) Then missingSettings.Add("Response_2", "Response (Model 2)")
                End If

                If context.INI_OAuth2 Then
                    If String.IsNullOrEmpty(context.INI_OAuth2ClientMail) Then missingSettings.Add("OAuth2ClientMail", "OAuth2Client Mail (Model 1)")
                    If String.IsNullOrEmpty(context.INI_OAuth2Scopes) Then missingSettings.Add("OAuth2Scopes", "OAuth2Scopes (Model 1)")
                    If String.IsNullOrEmpty(context.INI_OAuth2Endpoint) Then missingSettings.Add("OAuth2Endpoint", "OAuth2Endpoint (Model 1)")
                    If context.INI_OAuth2ATExpiry < 0 Then missingSettings.Add("OAuth2ATExpiry", "OAuth2ATExpiry (Model 1)")
                End If

                If context.INI_OAuth2_2 Then
                    If String.IsNullOrEmpty(context.INI_OAuth2ClientMail_2) Then missingSettings.Add("OAuth2ClientMail_2", "OAuth2ClientMail (Model 2)")
                    If String.IsNullOrEmpty(context.INI_OAuth2Scopes_2) Then missingSettings.Add("OAuth2Scopes_2", "OAuth2Scopes (Model 2)")
                    If String.IsNullOrEmpty(context.INI_OAuth2Endpoint_2) Then missingSettings.Add("OAuth2Endpoint_2", "OAuth2Endpoint (Model 2)")
                    If context.INI_OAuth2ATExpiry_2 < 0 Then missingSettings.Add("OAuth2ATExpiry_2", "OAuth2ATExpiry (Model 2)")
                End If

                If context.INI_ISearch And context.RDV.Substring(0, 4) = "Word" Then
                    If String.IsNullOrEmpty(context.INI_ISearch_URL) Then missingSettings.Add("ISearch_URL", "Search URL")
                    If String.IsNullOrEmpty(context.INI_ISearch_ResponseMask1) Then missingSettings.Add("ISearch_ResponseMask1", "Response Mask 1")
                    If String.IsNullOrEmpty(context.INI_ISearch_ResponseMask2) Then missingSettings.Add("ISearch_ResponseMask2", "Response Mask 2")
                    If String.IsNullOrEmpty(context.INI_ISearch_Name) Then missingSettings.Add("ISearch_Name", "ISearch_Name")
                    If context.INI_ISearch_Tries = 0 Then missingSettings.Add("ISearch_Tries", "ISearch_Tries")
                    If context.INI_ISearch_Results = 0 Then missingSettings.Add("ISearch_Results", "ISearch_Results")
                End If

                If context.INI_Lib And context.RDV.Substring(0, 4) = "Word" Then
                    If String.IsNullOrEmpty(context.INI_Lib_File) Then missingSettings.Add("Lib_File", "Lib_File")
                    If String.IsNullOrEmpty(context.INI_Lib_Find_SP) Then missingSettings.Add("Lib_Find_SP", "Lib_Find_SP")
                    If String.IsNullOrEmpty(context.INI_Lib_Apply_SP) Then missingSettings.Add("Lib_Apply_SP", "Lib_Apply_SP")
                    If String.IsNullOrEmpty(context.INI_Lib_Apply_SP_Markup) Then missingSettings.Add("Lib_Apply_SP_Markup", "Lib_Apply_SP_Markup")
                End If

                If context.INI_APIEncrypted Or context.INI_APIEncrypted_2 Then
                    If String.IsNullOrEmpty(context.Codebasis) Then missingSettings.Add("Codebasis", "CodeBasis (for decryption)")
                End If

                ' If there are missing settings, prompt user to complete them
                If missingSettings.Count > 0 Then
                    usercompleted = MissingSettingsWindow(missingSettings, context)
                    If Not usercompleted Then
                        ShowCustomMessageBox($"You have not provided all required parameters, which is why {AN} will not operate properly. Update '{AN2}.ini' (all values are described in the manual) before you continue or retry and add the parameters.")
                        Return True
                        Exit Do
                    End If
                Else
                    Return False
                    Exit Do
                End If
            Loop

        End Function


        Public Shared Function RenameFileToBak(filePath As String) As Boolean
            Try
                ' Rename the file to a .bak file
                Dim bakFilePath As String = filePath & ".bak"
                If File.Exists(bakFilePath) Then
                    File.Delete(bakFilePath)
                End If
                File.Move(filePath, bakFilePath)
                Return True
            Catch ex As Exception
                MessageBox.Show($"Error renaming file to .bak: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try

        End Function

        Public Shared Sub WriteToRegistry(ByVal regPath As String, ByVal regValue As String)
            Try
                ' Remove carriage returns from the value
                regValue = RemoveCR(regValue)

                ' Split the registry path into hive and subkey
                Dim hiveName As String = regPath.Split("\")(0)
                Dim subKeyPath As String = String.Join("\", regPath.Split("\").Skip(1))
                Dim registryHive As RegistryKey

                ' Determine the appropriate registry hive
                Select Case hiveName.ToUpper()
                    Case "HKEY_CURRENT_USER"
                        registryHive = Registry.CurrentUser
                    Case "HKEY_LOCAL_MACHINE"
                        registryHive = Registry.LocalMachine
                    Case Else
                        Throw New ArgumentException("Unsupported registry hive: " & hiveName)
                End Select

                ' Write the value to the registry
                Using subKey As RegistryKey = registryHive.CreateSubKey(subKeyPath, True)
                    If subKey Is Nothing Then
                        Throw New Exception("Unable to open or create the registry key at: " & regPath)
                    End If
                    subKey.SetValue("", regValue, RegistryValueKind.String)
                End Using

                ShowCustomMessageBox($"Written value '{regValue}' to the registry at '{regPath}.'")

            Catch ex As Exception
                MessageBox.Show($"Error: Unable to write to the registry at '{regPath}'. {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Public Shared Function GetFromRegistry(registryPath As String, valueName As String, Optional suppressErrors As Boolean = False) As String
            Try
                ' Split the registry path into hive and subkey
                Dim hiveName As String = registryPath.Split("\"c)(0)
                Dim subKeyPath As String = registryPath.Substring(hiveName.Length + 1)

                ' Determine the registry hive
                Dim hive As RegistryKey = Nothing
                Select Case hiveName.ToUpper()
                    Case "HKEY_CURRENT_USER"
                        hive = Registry.CurrentUser
                    Case "HKEY_LOCAL_MACHINE"
                        hive = Registry.LocalMachine
                    Case "HKEY_CLASSES_ROOT"
                        hive = Registry.ClassesRoot
                    Case "HKEY_USERS"
                        hive = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        hive = Registry.CurrentConfig
                    Case Else
                        If Not suppressErrors Then
                            MessageBox.Show("Error in GetFromRegistry - invalid registry hive: " & hiveName, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        Return ""
                End Select

                ' Open the subkey and retrieve the value
                Using subKey As RegistryKey = hive.OpenSubKey(subKeyPath)
                    If subKey IsNot Nothing Then
                        Return RemoveCR(subKey.GetValue(valueName, Nothing)?.ToString())
                    Else
                        If Not suppressErrors Then
                            MessageBox.Show("Error in GetFromRegistry - Registry key not found: " & subKeyPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        Return ""
                    End If
                End Using

            Catch ex As System.Exception
                If Not suppressErrors Then
                    MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return ""
            End Try
        End Function

        Public Shared Function RemoveCR(ByVal inputtext As String) As String
            If inputtext IsNot Nothing Then
                inputtext = inputtext.Trim()
                inputtext = inputtext.Replace(vbCr, "")
                inputtext = inputtext.Replace(vbLf, "")
                inputtext = inputtext.Replace(vbCrLf, "")
                inputtext = inputtext.Trim()
            Else
                inputtext = ""
            End If
            Return inputtext
        End Function

        Public Shared Function IsEmptyOrBlank(ByVal str As String) As Boolean
            ' Check if the string is empty or consists only of whitespace
            Return String.IsNullOrWhiteSpace(str)
        End Function

        Public Shared Function ExpandEnvironmentVariables(ByVal filePath As String) As String
            ' Start with the input path
            Dim expandedPath As String = Environment.ExpandEnvironmentVariables(filePath)

            Try

                ' Remove any preceding and trailing quotation marks
                expandedPath = expandedPath.Trim(""""c)

                ' Expand known variables using Environment.GetEnvironmentVariable and ensure proper path format
                expandedPath = Regex.Replace(expandedPath, "%APPDATA%", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%USERPROFILE%", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%WINDIR%", Path.Combine(Environment.GetEnvironmentVariable("WINDIR")), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%TEMP%", Path.Combine(Path.GetTempPath()), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%HOMEPATH%", Path.Combine(Environment.GetEnvironmentVariable("HOMEPATH")), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%APPSTARTUPPATH%", Path.Combine(System.Windows.Forms.Application.StartupPath), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%DESKTOP%", Path.Combine(Environment.SpecialFolder.Desktop), RegexOptions.IgnoreCase)

                ' Clean up any potential double backslashes
                expandedPath = Regex.Replace(expandedPath, "\\{2,}", "\\")

                ' Return the expanded path
                If expandedPath = "" Then Return "" Else Return Path.GetFullPath(expandedPath)

            Catch ex As System.Exception
                ' Return Nothing on failure
                Return ""
            End Try

        End Function

        Public Shared Function RealAPIKey(ByVal APIInput As String, ByVal SecondAPI As Boolean, ByVal IgnorePrefix As Boolean, ByVal context As ISharedContext) As String

            APIInput = Trim(RemoveCR(APIInput))

            Dim Prefix As String = ""
            Dim Result As String = APIInput

            ' Determine the prefix based on whether it's the second API and IgnorePrefix is false
            If Not SecondAPI Then
                If Not IgnorePrefix Then
                    Prefix = context.INI_APIKeyPrefix

                    If Not String.IsNullOrWhiteSpace(Prefix) Then
                        ' Remove the prefix if present
                        If APIInput.StartsWith(Prefix) Then
                            APIInput = APIInput.Substring(Prefix.Length)
                        End If
                    End If
                End If

                Result = APIInput

                ' Decode the API key if encryption is enabled for the main API
                If context.INI_APIEncrypted Then
                    Result = DecodeString(APIInput, context.Codebasis)
                End If
            Else
                If Not IgnorePrefix Then
                    Prefix = context.INI_APIKeyPrefix_2

                    If Not String.IsNullOrWhiteSpace(Prefix) Then
                        ' Remove the prefix if present
                        If APIInput.StartsWith(Prefix) Then
                            APIInput = APIInput.Substring(Prefix.Length)
                        End If
                    End If
                End If

                Result = APIInput

                ' Decode the API key if encryption is enabled for the second API
                If context.INI_APIEncrypted_2 Then
                    Result = DecodeString(APIInput, context.Codebasis)
                End If
            End If

            ' Remove any carriage return characters
            Result = RemoveCR(Result)

            ' Add the prefix back and return the final result

            Result = Prefix & Result

            Return Result
        End Function

        Public Shared Function DecodeBase64(ByVal base64String As String) As Byte()
            Try
                ' Normalize the input: remove whitespaces and line breaks
                base64String = base64String.Replace(vbCrLf, "").Replace(vbLf, "").Replace(vbCr, "").Replace(" ", "")

                ' Convert URL-safe Base64 to standard Base64 if input is URL-safe
                base64String = base64String.Replace("-", "+").Replace("_", "/")

                ' Add padding
                While (base64String.Length Mod 4) <> 0
                    base64String &= "="
                End While

                ' Decode the Base64 string
                Return Convert.FromBase64String(base64String)
            Catch ex As System.Exception
                ' Return Nothing on failure
                Return Nothing
            End Try
        End Function

        Public Shared Function DecodeString(ByVal encodedText As String, ByVal pTerm As String) As String
            ' Remove literal "\n" if present
            encodedText = encodedText.Replace("\n", "")
            ' Also ensure actual newline characters are removed
            encodedText = encodedText.Replace(vbCr, "").Replace(vbLf, "")
            ' Remove spaces if any
            encodedText = encodedText.Replace(" ", "")

            Dim encryptedBytes As Byte() = DecodeBase64(encodedText)
            If encryptedBytes Is Nothing Then
                Return "Error: Invalid Base64 input"
            End If

            Dim pTermBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(pTerm)
            Dim decryptedBytes(encryptedBytes.Length - 1) As Byte

            For i As Integer = 0 To encryptedBytes.Length - 1
                decryptedBytes(i) = encryptedBytes(i) Xor pTermBytes(i Mod pTermBytes.Length)
            Next

            ' Convert decrypted bytes to string
            ' If UTF8 fails due to unexpected characters, try ASCII or verify the original encoding.
            Try
                Return System.Text.Encoding.UTF8.GetString(decryptedBytes)
            Catch
                Return System.Text.Encoding.ASCII.GetString(decryptedBytes)
            End Try
        End Function

        Public Shared Function CodeString(ByVal inputText As String, ByVal pTerm As String) As String
            Dim inputBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(inputText)
            Dim pTermBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(pTerm)
            Dim encryptedBytes(inputBytes.Length - 1) As Byte

            Dim inputLength As Integer = inputBytes.Length
            Dim pTermLength As Integer = pTermBytes.Length

            ' Encrypt each byte with XOR operation
            For i As Integer = 0 To inputBytes.Length - 1
                encryptedBytes(i) = inputBytes(i) Xor pTermBytes(i Mod pTermLength)
            Next

            ' Convert encrypted bytes to Base64
            Return Convert.ToBase64String(encryptedBytes)
        End Function

        Public Shared Function GetDomain() As String
            Try
                ' Initialize a WMI query to get the Domain property from Win32_ComputerSystem
                Dim searcher As New ManagementObjectSearcher("SELECT Domain FROM Win32_ComputerSystem")
                Dim strDomain As String = String.Empty

                ' Execute the query and retrieve the result
                For Each queryObj As ManagementObject In searcher.Get()
                    If queryObj("Domain") IsNot Nothing Then
                        strDomain = queryObj("Domain").ToString()
                    End If
                Next

                ' If the domain is not retrieved, return an appropriate message
                If String.IsNullOrEmpty(strDomain) Then
                    MessageBox.Show($"Error in GetDomain - unable to determine the domain name or workgroup.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    strDomain = ""
                End If

                Return strDomain
            Catch ex As System.Exception
                MessageBox.Show($"Error in GetDomain - Error retrieving domain or workgroup: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return String.Empty
            End Try
        End Function

        Public Shared Function WrongDomain()
            Dim strDomain As String = GetDomain() ' Current domain of the computer
            Dim domainList() As String
            Dim domainFound As Boolean = False

            If Not String.IsNullOrEmpty(alloweddomains) Then
                ' Convert the list of allowed domains into an array
                domainList = alloweddomains.Split(","c)

                ' Check if the current domain is in the allowed list
                For Each domain In domainList
                    If strDomain.Equals(domain.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        domainFound = True
                        Exit For
                    End If
                Next

                ' If the domain is not in the list of allowed domains
                If Not domainFound Then
                    ShowCustomMessageBox($"This copy of {AN} may not be executed in this network environment (which is '{strDomain}'). The domain has to be added to the code by your administrator.")
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function

        Public Shared Function ShowCustomInputBox(prompt As String, title As String, SimpleInput As Boolean, Optional DefaultValue As String = "") As String

            Dim inputForm As New Form()
            inputForm.Opacity = 0
            Dim promptLabel As New System.Windows.Forms.Label()
            Dim inputTextBox As New TextBox()
            Dim okButton As New Button()
            Dim cancelButton As New Button()

            ' Form attributes
            inputForm.Text = title
            inputForm.FormBorderStyle = FormBorderStyle.FixedDialog
            inputForm.StartPosition = FormStartPosition.CenterScreen
            inputForm.MaximizeBox = False
            inputForm.MinimizeBox = False
            inputForm.ShowInTaskbar = False
            inputForm.TopMost = True


            ' Set the icon
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            inputForm.Icon = Icon.FromHandle(bmp.GetHicon())

            ' Set predefined font for consistent layout
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            inputForm.Font = standardFont

            ' Prompt label
            promptLabel.Text = prompt
            promptLabel.Font = standardFont
            promptLabel.AutoSize = True
            promptLabel.MaximumSize = New Size(500, 0) ' Increased maximum width for text wrapping
            promptLabel.Location = New System.Drawing.Point(20, 20) ' Margin around the prompt
            inputForm.Controls.Add(promptLabel)

            ' Input TextBox
            Dim textBoxHeight As Integer = If(SimpleInput, 25, 100)
            inputTextBox.Multiline = Not SimpleInput
            inputTextBox.WordWrap = True
            inputTextBox.ScrollBars = If(SimpleInput, ScrollBars.None, ScrollBars.Vertical)
            inputTextBox.Location = New System.Drawing.Point(20, promptLabel.Bottom + 20) ' Margin below the prompt
            inputTextBox.Width = 495
            inputTextBox.Height = textBoxHeight
            inputTextBox.Text = DefaultValue ' Set default value if provided
            inputForm.Controls.Add(inputTextBox)


            ' Add KeyDown handler for Enter key if SimpleInput is True
            If SimpleInput Then
                AddHandler inputTextBox.KeyDown, Sub(sender, e)
                                                     If e.KeyCode = Keys.Enter Then
                                                         inputForm.DialogResult = DialogResult.OK
                                                         inputForm.Close()
                                                         e.SuppressKeyPress = True ' Prevent the ding sound
                                                     End If
                                                 End Sub
            Else
                AddHandler inputTextBox.KeyDown, Sub(sender, e)
                                                     If e.KeyCode = Keys.Enter AndAlso e.Modifiers = Keys.Control Then
                                                         inputForm.DialogResult = DialogResult.OK
                                                         inputForm.Close()
                                                         e.SuppressKeyPress = True ' Prevent the ding sound
                                                     End If
                                                 End Sub
                AddHandler inputTextBox.KeyDown, Sub(sender, e)
                                                     If e.KeyCode = Keys.Escape Then
                                                         inputForm.DialogResult = DialogResult.Cancel
                                                         inputForm.Close()
                                                         e.SuppressKeyPress = True ' Prevent the ding sound
                                                     End If
                                                 End Sub
            End If



            ' Buttons
            okButton.Text = "OK"
            cancelButton.Text = "Cancel"

            ' Measure and adjust button sizes based on font
            Dim okButtonSize As Size = TextRenderer.MeasureText(okButton.Text, standardFont)
            Dim cancelButtonSize As Size = TextRenderer.MeasureText(cancelButton.Text, standardFont)
            Dim buttonWidth As Integer = Math.Max(okButtonSize.Width, cancelButtonSize.Width) + 20
            Dim buttonHeight As Integer = Math.Max(okButtonSize.Height, cancelButtonSize.Height) + 10
            okButton.Size = New Size(buttonWidth, buttonHeight)
            cancelButton.Size = New Size(buttonWidth, buttonHeight)

            ' Button positions
            okButton.Location = New System.Drawing.Point(20, inputTextBox.Bottom + 20) ' Margin below the input box
            cancelButton.Location = New System.Drawing.Point(okButton.Right + 10, inputTextBox.Bottom + 20) ' Margin between buttons
            inputForm.Controls.Add(okButton)
            inputForm.Controls.Add(cancelButton)

            ' Button click handlers
            AddHandler okButton.Click, Sub(sender, e)
                                           inputForm.DialogResult = DialogResult.OK
                                           inputForm.Close()
                                       End Sub
            AddHandler cancelButton.Click, Sub(sender, e)
                                               inputForm.DialogResult = DialogResult.Cancel
                                               inputForm.Close()
                                           End Sub

            ' Adjust form size dynamically
            Dim formWidth As Integer = Math.Max(500, Math.Max(promptLabel.Width + 40, inputTextBox.Width + 40)) ' Adjusted width for the form
            Dim formHeight As Integer = cancelButton.Bottom + 30
            inputForm.ClientSize = New Size(formWidth, formHeight)

            ' Show dialog

            inputForm.TopMost = True
            inputForm.BringToFront()
            inputForm.Focus()

            Dim Result As DialogResult

            If title.Contains("Browser") Then
                Dim outlookApp As Object = CreateObject("Outlook.Application")

                If outlookApp IsNot Nothing Then
                    Dim explorer As Object = outlookApp.GetType().InvokeMember(
                        "ActiveExplorer",
                        Reflection.BindingFlags.GetProperty,
                        Nothing,
                        outlookApp,
                        Nothing
                    )
                    If explorer IsNot Nothing Then
                        ' WindowState = 1 => Normal window (OlWindowState.olNormalWindow)
                        explorer.GetType().InvokeMember(
                            "WindowState",
                            Reflection.BindingFlags.SetProperty,
                            Nothing,
                            explorer,
                            New Object() {1}
                        )
                        explorer.GetType().InvokeMember(
                            "Activate",
                            Reflection.BindingFlags.InvokeMethod,
                            Nothing,
                            explorer,
                            Nothing
                        )
                    End If
                End If
                inputForm.Opacity = 1
                Dim outlookHwnd As IntPtr = FindWindow("rctrl_renwnd32", Nothing) ' or however you get it
                Result = inputForm.ShowDialog(New WindowWrapper(outlookHwnd))
            Else
                inputForm.Opacity = 1
                Result = inputForm.ShowDialog()
            End If

            If Result = DialogResult.OK Then
                Return inputTextBox.Text
            Else
                If Not SimpleInput Then
                    Return "ESC"
                Else
                    Return ""
                End If
            End If
        End Function


        Public Shared Function ShowCustomYesNoBox(ByVal bodyText As String, ByVal button1Text As String, ByVal button2Text As String, Optional header As String = AN, Optional autoCloseSeconds As Integer? = Nothing, Optional Defaulttext As String = "") As Integer
            Dim messageForm As New Form()
            messageForm.Opacity = 0
            Dim bodyLabel As New System.Windows.Forms.Label()
            Dim button1 As New Button()
            Dim button2 As New Button()
            Dim countdownLabel As New System.Windows.Forms.Label()
            Dim truncatedLabel As New System.Windows.Forms.Label()
            Dim Truncated As Boolean = False

            If Len(bodyText) > 10000 Then
                bodyText = Left(bodyText, 10000)
                Truncated = True
            End If

            ' Form attributes
            messageForm.Text = header
            messageForm.FormBorderStyle = FormBorderStyle.FixedDialog
            messageForm.StartPosition = FormStartPosition.CenterScreen
            messageForm.MaximizeBox = False
            messageForm.MinimizeBox = False
            messageForm.ShowInTaskbar = False
            messageForm.TopMost = True

            ' Set the icon
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            messageForm.Icon = Icon.FromHandle(bmp.GetHicon())

            ' Set predefined font for consistent layout
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            messageForm.Font = standardFont

            ' Body Label
            bodyLabel.Text = bodyText
            bodyLabel.Font = standardFont

            Dim maxLabelWidth As Integer = 450 ' Maximum width for the label
            Dim maxScreenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height - 100

            ' Measure text size considering maximum width and word wrap
            bodyLabel.MaximumSize = New Size(maxLabelWidth, maxScreenHeight \ 2)
            bodyLabel.AutoSize = True

            ' Calculate required width for the form based on the label's rendered size
            Dim textSize As Size = TextRenderer.MeasureText(bodyText, standardFont, New Size(maxLabelWidth, Integer.MaxValue), TextFormatFlags.WordBreak)
            Dim requiredWidth As Integer = Math.Min(textSize.Width, maxLabelWidth)

            ' Adjust label size for word wrap
            bodyLabel.Size = New Size(requiredWidth, textSize.Height)
            bodyLabel.Location = New System.Drawing.Point(20, 20) ' Add margin
            messageForm.Controls.Add(bodyLabel)

            ' Adjust form width to fit the label and buttons
            Dim formWidth As Integer = Math.Max(requiredWidth + 40, button1.Width + button2.Width + 60) ' Include margin and button widths
            messageForm.ClientSize = New Size(formWidth, messageForm.ClientSize.Height)


            ' Button1
            button1.Text = button1Text
            Dim button1Size As Size = TextRenderer.MeasureText(button1.Text, standardFont)
            button1.Size = New Size(button1Size.Width + 20, button1Size.Height + 10)
            button1.Location = New System.Drawing.Point(20, bodyLabel.Bottom + 20)
            messageForm.Controls.Add(button1)

            ' Button2
            button2.Text = button2Text
            Dim button2Size As Size = TextRenderer.MeasureText(button2.Text, standardFont)
            button2.Size = New Size(button2Size.Width + 20, button2Size.Height + 10)
            button2.Location = New System.Drawing.Point(button1.Right + 10, bodyLabel.Bottom + 20)
            messageForm.Controls.Add(button2)

            If Truncated Then
                truncatedLabel.Text = "(text has been truncated)"
                truncatedLabel.Font = standardFont
                truncatedLabel.AutoSize = True
                truncatedLabel.Location = New System.Drawing.Point(bodyLabel.Right - TextRenderer.MeasureText(truncatedLabel.Text, standardFont).Width, bodyLabel.Bottom + 20)
                messageForm.Controls.Add(truncatedLabel)
            End If

            ' Countdown Label
            countdownLabel.Font = standardFont
            countdownLabel.Text = $"(closes in 0 seconds{Defaulttext})"
            countdownLabel.AutoSize = True
            countdownLabel.Location = New System.Drawing.Point(button2.Right + 10, bodyLabel.Bottom + 25)
            messageForm.Controls.Add(countdownLabel)

            ' Adjust the height of the form dynamically
            Dim totalHeight As Integer = bodyLabel.Bottom + 20 ' Start with the bottom of the body label plus margin

            ' Include buttons and any additional controls
            Dim buttonsBottom As Integer = Math.Max(button1.Bottom, button2.Bottom)
            totalHeight = Math.Max(totalHeight, buttonsBottom + 20)

            ' Include truncated label or countdown label if present
            If Truncated Then
                totalHeight = Math.Max(totalHeight, truncatedLabel.Bottom + 20)
            End If

            If autoCloseSeconds.HasValue Then
                totalHeight = Math.Max(totalHeight, countdownLabel.Bottom + 20)
            End If

            ' Set the form height
            totalHeight = Math.Min(totalHeight, maxScreenHeight) ' Ensure it doesn't exceed the screen's max height
            messageForm.ClientSize = New Size(Math.Max(countdownLabel.Right + 20, bodyLabel.Right + 20), totalHeight)

            ' Result variable
            Dim result As Integer = 0

            ' Button click handlers
            AddHandler button1.Click, Sub(sender, e)
                                          result = 1
                                          messageForm.Close()
                                      End Sub
            AddHandler button2.Click, Sub(sender, e)
                                          result = 2
                                          messageForm.Close()
                                      End Sub

            ' Timer for auto-close functionality
            If autoCloseSeconds.HasValue Then
                Dim remainingTime As Integer = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"

                ' Ensure the form handle is created.
                Dim dummyHandle = messageForm.Handle

                Dim timer As New System.Timers.Timer(1000) ' 1 second interval
                timer.AutoReset = True
                AddHandler timer.Elapsed, Sub(sender As Object, e As System.Timers.ElapsedEventArgs)
                                              messageForm.BeginInvoke(Sub()
                                                                          remainingTime -= 1
                                                                          If remainingTime > 0 Then
                                                                              countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"
                                                                          Else
                                                                              timer.Stop()
                                                                              result = 3
                                                                              messageForm.Close()
                                                                          End If
                                                                      End Sub)
                                          End Sub
                timer.Start()
                messageForm.Opacity = 1
                messageForm.ShowDialog()
            Else
                countdownLabel.Text = ""
                messageForm.Opacity = 1
                messageForm.ShowDialog()
            End If


            Return result
        End Function

        Public Shared Sub ShowCustomMessageBox(ByVal bodyText As String, Optional header As String = AN, Optional autoCloseSeconds As Integer? = Nothing, Optional Defaulttext As String = " - execution continues meanwhile", Optional SeparateThread As Boolean = False)
            Dim messageForm As New Form()
            messageForm.Opacity = 0
            Dim bodyLabel As New System.Windows.Forms.Label()
            Dim okButton As New Button()
            Dim countdownLabel As New System.Windows.Forms.Label()
            Dim truncatedLabel As New System.Windows.Forms.Label()
            Dim Truncated As Boolean = False

            If String.IsNullOrWhiteSpace(header) Then header = AN

            If Len(bodyText) > 15000 Then
                bodyText = Left(bodyText, 15000) + "(...)"
                Truncated = True
            End If

            ' Form attributes
            messageForm.Text = header
            messageForm.FormBorderStyle = FormBorderStyle.FixedDialog
            messageForm.StartPosition = FormStartPosition.CenterScreen
            messageForm.MaximizeBox = False
            messageForm.MinimizeBox = False
            messageForm.ShowInTaskbar = False
            messageForm.TopMost = True

            ' Set the icon
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            messageForm.Icon = Icon.FromHandle(bmp.GetHicon())

            ' Set predefined font for consistent layout
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            messageForm.Font = standardFont

            ' Body Label
            bodyLabel.Text = bodyText
            bodyLabel.Font = standardFont
            bodyLabel.AutoSize = True

            ' Calculate label width and decide layout
            Dim maxLabelWidth As Integer = 450
            Dim maxScreenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width - 100
            Dim maxScreenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height - 100

            bodyLabel.MaximumSize = New Size(maxLabelWidth, maxScreenHeight / 2)
            Dim MaxConsiderSize As New Size(maxLabelWidth, Integer.MaxValue)

            If TextRenderer.MeasureText(bodyText, standardFont, MaxConsiderSize).Width <= maxLabelWidth And Not TextRenderer.MeasureText(bodyText, standardFont, MaxConsiderSize, TextFormatFlags.WordBreak).Height > ((maxLabelWidth / 16) * 9) Then
                ' Single line, set label width to fit content
                bodyLabel.MaximumSize = New Size(maxLabelWidth, Integer.MaxValue)
            Else
                ' Multi-line, set label width to maxLabelWidth and adjust height
                bodyLabel.MaximumSize = New Size(maxLabelWidth, Integer.MaxValue)

                If TextRenderer.MeasureText(bodyText, standardFont, MaxConsiderSize, TextFormatFlags.WordBreak).Height > (maxScreenHeight \ 2) Then

                    ' Recalculate maxLabelWidth for 16:9, ensuring it fits within the screen width
                    Dim proposedWidth As Integer = (maxScreenHeight \ 2) * 16 \ 9
                    If proposedWidth > maxScreenWidth Then
                        proposedWidth = maxScreenWidth
                    End If

                    bodyLabel.MaximumSize = New Size(proposedWidth, maxScreenHeight \ 2)
                    MaxConsiderSize = New Size(proposedWidth, Integer.MaxValue)

                    If TextRenderer.MeasureText(bodyText, standardFont, MaxConsiderSize, TextFormatFlags.WordBreak).Height > (maxScreenHeight \ 2) Then
                        Truncated = True
                    End If
                End If
            End If
            bodyLabel.Location = New System.Drawing.Point(20, 20) ' Margin around the label
            bodyLabel.AutoEllipsis = True
            messageForm.Controls.Add(bodyLabel)

            ' OK Button
            okButton.Text = "OK"

            ' Measure and adjust button size based on font
            Dim okButtonSize As Size = TextRenderer.MeasureText(okButton.Text, standardFont)
            okButton.Size = New Size(okButtonSize.Width + 20, okButtonSize.Height + 10)

            ' Button position aligned to the left
            okButton.Location = New System.Drawing.Point(20, bodyLabel.Bottom + 20) ' Align to the left below text
            messageForm.Controls.Add(okButton)

            ' Countdown Label
            countdownLabel.Font = standardFont
            countdownLabel.AutoSize = True
            countdownLabel.Location = New System.Drawing.Point(okButton.Right + 10, bodyLabel.Bottom + 25) ' Positioned to the right of the OK button
            messageForm.Controls.Add(countdownLabel)

            ' Button click handler
            Dim userClicked As Boolean = False
            AddHandler okButton.Click, Sub(sender, e)
                                           userClicked = True
                                           messageForm.Close()
                                       End Sub

            ' Adjust form size dynamically based on label size
            Dim formHeight As Integer = bodyLabel.Bottom + okButton.Height + 50
            Dim formWidth As Integer = Math.Max(bodyLabel.Width + 40, okButton.Right + 20 + TextRenderer.MeasureText("(closes in 99 seconds{defaulttext})", standardFont).Width)
            'Dim formWidth As Integer = Math.Min(bodyLabel.Width + 40, maxScreenWidth)
            formWidth = Math.Min(formWidth, maxScreenWidth)

            If formHeight > maxScreenHeight Then
                formHeight = maxScreenHeight
            End If

            messageForm.ClientSize = New Size(formWidth, formHeight)

            ' Recalculate button position based on adjusted form size
            okButton.Location = New System.Drawing.Point(20, bodyLabel.Bottom + 20)
            countdownLabel.Location = New System.Drawing.Point(okButton.Right + 10, bodyLabel.Bottom + 25)

            ' Timer for auto-close functionality
            If autoCloseSeconds.HasValue Then
                Dim remainingTime As Integer = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"

                Dim timer As New System.Windows.Forms.Timer()
                timer.Interval = 1000 ' Tick every second
                AddHandler timer.Tick, Sub(sender, e)
                                           remainingTime -= 1
                                           If remainingTime > 0 Then
                                               countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"
                                           Else
                                               timer.Stop()
                                               If Not userClicked Then
                                                   messageForm.Close()
                                               End If
                                           End If
                                       End Sub
                timer.Start()

                ' Show the message box non-modally if timer is used
                messageForm.Opacity = 1
                If SeparateThread Then
                    messageForm.ShowDialog()
                Else
                    messageForm.Show()
                    System.Windows.Forms.Application.DoEvents()
                End If
            Else
                ' Show the message box modally if no timer is used
                messageForm.Opacity = 1
                messageForm.ShowDialog()
            End If
        End Sub


        Public Class ProgressForm
            Inherits Form

            Private WithEvents progressBar As ProgressBar
            Private WithEvents lblHeader As System.Windows.Forms.Label
            Private WithEvents lblStatus As System.Windows.Forms.Label
            Private WithEvents btnCancel As Button
            Private WithEvents uiTimer As System.Windows.Forms.Timer

            ' Constructor receives the header text and the initial status text.
            Public Sub New(headerText As String, initialLabel As String)
                ' Set form properties.
                Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                Me.Text = headerText
                Me.Width = 400
                Me.Height = 220
                Me.FormBorderStyle = FormBorderStyle.FixedDialog
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.MaximizeBox = False
                Me.MinimizeBox = False
                Me.ShowInTaskbar = False
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                Me.Icon = Icon.FromHandle(bmp.GetHicon())

                ' Create and add header label.
                lblHeader = New System.Windows.Forms.Label()
                lblHeader.Text = "Progress ..."
                lblHeader.AutoSize = True
                lblHeader.Top = 10
                lblHeader.Left = 10
                lblHeader.Font = standardFont
                Me.Controls.Add(lblHeader)

                ' Create and add progress bar.
                progressBar = New ProgressBar()
                progressBar.Minimum = 0
                progressBar.Maximum = ProgressBarModule.GlobalProgressMax
                progressBar.Width = 360
                progressBar.Height = 25
                progressBar.Top = 40
                progressBar.Left = 10
                Me.Controls.Add(progressBar)

                ' Create and add status label.
                lblStatus = New System.Windows.Forms.Label()
                lblStatus.Text = initialLabel
                lblStatus.AutoSize = True
                lblStatus.Top = 75
                lblStatus.Left = 10
                lblStatus.Width = 360
                lblStatus.Font = standardFont
                Me.Controls.Add(lblStatus)

                ' Create and add Cancel button.
                btnCancel = New Button()
                btnCancel.Text = "Cancel"
                btnCancel.Top = 120
                btnCancel.Left = 10
                btnCancel.Font = standardFont
                btnCancel.Visible = True
                btnCancel.AutoSize = True
                AddHandler btnCancel.Click, AddressOf btnCancel_Click
                Me.Controls.Add(btnCancel)


                ' Create a timer to update the UI controls based on global variables.
                uiTimer = New System.Windows.Forms.Timer()
                uiTimer.Interval = 250 ' Update every 250 ms.
                AddHandler uiTimer.Tick, AddressOf Timer_Tick
                uiTimer.Start()
            End Sub

            ' Timer tick event updates the progress bar and status label.
            Private Sub Timer_Tick(sender As Object, e As EventArgs)
                Try
                    ' Update the progress bar maximum and value.
                    progressBar.Maximum = ProgressBarModule.GlobalProgressMax
                    progressBar.Value = Math.Min(ProgressBarModule.GlobalProgressValue, progressBar.Maximum)

                    ' Update the status text.
                    lblStatus.Text = ProgressBarModule.GlobalProgressLabel

                    ' If the cancel flag is set, close the form with a Cancel result.
                    If ProgressBarModule.CancelOperation Then
                        Me.DialogResult = DialogResult.Cancel
                        Me.Close()
                    End If
                Catch ex As Exception
                    ' It is possible to get an exception if the form is closing.
                    Debug.WriteLine("Timer error: " & ex.Message)
                End Try
            End Sub

            ' When the Cancel button is clicked, set the global cancel flag.
            Private Sub btnCancel_Click(sender As Object, e As EventArgs)
                ProgressBarModule.CancelOperation = True
            End Sub

            ' Stop the timer when the form is closed.
            Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
                uiTimer.Stop()
                ProgressBarModule.CancelOperation = True
                MyBase.OnFormClosed(e)
            End Sub
        End Class


        Public Shared Sub ShowRTFCustomMessageBox(ByVal bodyText As String, Optional header As String = AN, Optional autoCloseSeconds As Integer? = Nothing, Optional Defaulttext As String = " - execution continues meanwhile")

            Dim RTFMessageForm As New Form()
            Dim bodyLabel As New System.Windows.Forms.RichTextBox()
            Dim okButton As New Button()
            Dim countdownLabel As New System.Windows.Forms.Label()
            Dim truncatedLabel As New System.Windows.Forms.Label()
            Dim Truncated As Boolean = False

            If String.IsNullOrWhiteSpace(header) Then header = AN

            ' Form attributes
            RTFMessageForm.Opacity = 0
            RTFMessageForm.Text = header
            RTFMessageForm.FormBorderStyle = FormBorderStyle.Sizable ' Allow resizing
            RTFMessageForm.StartPosition = FormStartPosition.CenterScreen
            RTFMessageForm.MaximizeBox = True
            RTFMessageForm.MinimizeBox = True
            RTFMessageForm.ShowInTaskbar = False
            RTFMessageForm.TopMost = True
            RTFMessageForm.KeyPreview = True ' Enable form-level key event handling

            ' Set minimum size to prevent controls from being cut off
            RTFMessageForm.MinimumSize = New Size(650, 335) ' Adjust based on your needs

            ' Set the icon
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            RTFMessageForm.Icon = Icon.FromHandle(bmp.GetHicon())

            ' Set predefined font for consistent layout
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

            ' Body Label (now being a RTF box)
            bodyLabel.ReadOnly = True ' Make it behave like a label
            bodyLabel.BorderStyle = BorderStyle.None ' Remove borders for a clean look
            bodyLabel.BackColor = RTFMessageForm.BackColor ' Match the form's background color
            bodyLabel.TabStop = False ' Disable tab focus
            bodyLabel.Clear()
            bodyLabel.Rtf = bodyText
            bodyLabel.Width = 600
            bodyLabel.Height = 200
            bodyLabel.Location = New System.Drawing.Point(20, 20) ' Margin around the label
            bodyLabel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom ' Enable resizing

            RTFMessageForm.Controls.Add(bodyLabel)

            ' OK Button
            okButton.Text = "OK"
            Dim okButtonSize As Size = TextRenderer.MeasureText(okButton.Text, standardFont)
            okButton.AutoSize = True
            'okButton.Size = New Size(okButtonSize.Width + 40, okButtonSize.Height + 20) ' Make the button larger
            okButton.Location = New System.Drawing.Point(20, RTFMessageForm.ClientSize.Height - okButton.Height - 20) ' Adjust to the bottom left corner
            okButton.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left ' Keep position fixed at the bottom left
            RTFMessageForm.Controls.Add(okButton)

            ' Countdown Label
            countdownLabel.Font = standardFont
            countdownLabel.AutoSize = True
            countdownLabel.Location = New System.Drawing.Point(okButton.Right + 10, RTFMessageForm.ClientSize.Height - okButton.Height - 15) ' Positioned to the right of the OK button
            countdownLabel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left ' Keep position fixed
            RTFMessageForm.Controls.Add(countdownLabel)


            ' Button click handler
            Dim userClicked As Boolean = False
            AddHandler okButton.Click, Sub(sender, e)
                                           userClicked = True
                                           RTFMessageForm.Close()
                                           RTFMessageForm = Nothing
                                       End Sub
            AddHandler RTFMessageForm.KeyDown, Sub(sender, e)
                                                   If e.KeyCode = Keys.Escape Then
                                                       userClicked = True
                                                       RTFMessageForm.Close()
                                                       RTFMessageForm = Nothing
                                                       e.SuppressKeyPress = True ' Prevent the ding sound
                                                   End If
                                               End Sub
            AddHandler RTFMessageForm.Shown, Sub(sender, e)
                                                 RTFMessageForm.Activate()
                                             End Sub


            ' Adjust initial form size dynamically
            Dim formWidth As Integer = Math.Max(650, bodyLabel.Width + 40)
            Dim formHeight As Integer = Math.Max(335, bodyLabel.Height + okButton.Height + 80)

            ' Adjust control positions after setting form size
            okButton.Location = New System.Drawing.Point(20, RTFMessageForm.ClientSize.Height - okButton.Height - 20)
            countdownLabel.Location = New System.Drawing.Point(okButton.Right + 10, okButton.Top)

            RTFMessageForm.ClientSize = New Size(formWidth, formHeight)

            ' Recalculate button position based on adjusted form size
            okButton.Location = New System.Drawing.Point(20, bodyLabel.Bottom + 20)
            countdownLabel.Location = New System.Drawing.Point(okButton.Right + 10, bodyLabel.Bottom + 25)

            ' Timer for auto-close functionality

            If autoCloseSeconds.HasValue AndAlso autoCloseSeconds > 0 Then
                Dim remainingTime As Integer = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"

                Dim timer As New System.Windows.Forms.Timer()
                timer.Interval = 1000 ' Tick every second
                AddHandler timer.Tick, Sub(sender, e)
                                           remainingTime -= 1
                                           If remainingTime > 0 Then
                                               countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"
                                           Else
                                               timer.Stop()
                                               If Not userClicked Then
                                                   RTFMessageForm.Close()
                                               End If
                                           End If
                                       End Sub
                timer.Start()

                ' Show the message box non-modally if timer is used
                RTFMessageForm.Opacity = 1
                RTFMessageForm.Show()
                RTFMessageForm.BringToFront()
                RTFMessageForm.Activate()
                System.Windows.Forms.Application.DoEvents()
            Else

                ' Show the message box modally if no timer is used
                RTFMessageForm.Opacity = 1
                RTFMessageForm.TopMost = True
                RTFMessageForm.ShowDialog()
            End If
        End Sub

        Public Shared Sub ShowHTMLCustomMessageBox(ByVal bodyText As String, Optional header As String = AN, Optional Defaulttext As String = " - execution continues meanwhile")
            Dim t As New Threading.Thread(Sub()
                                              Dim HTMLMessageForm As New System.Windows.Forms.Form()
                                              Dim htmlBrowser As New System.Windows.Forms.WebBrowser()
                                              Dim okButton As New System.Windows.Forms.Button()

                                              If String.IsNullOrWhiteSpace(header) Then header = AN

                                              ' Form attributes
                                              HTMLMessageForm.Opacity = 0
                                              HTMLMessageForm.Text = header
                                              HTMLMessageForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
                                              HTMLMessageForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                                              HTMLMessageForm.MaximizeBox = True
                                              HTMLMessageForm.MinimizeBox = True
                                              HTMLMessageForm.ShowInTaskbar = True
                                              HTMLMessageForm.TopMost = False
                                              HTMLMessageForm.KeyPreview = True
                                              HTMLMessageForm.MinimumSize = New System.Drawing.Size(650, 335)

                                              ' (Optional) Remove the icon setting if you suspect it might be causing issues.
                                              Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
                                              HTMLMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

                                              Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

                                              ' HTML Browser
                                              htmlBrowser.AllowNavigation = False
                                              htmlBrowser.WebBrowserShortcutsEnabled = False
                                              htmlBrowser.ScrollBarsEnabled = True
                                              htmlBrowser.ScriptErrorsSuppressed = True
                                              htmlBrowser.DocumentText = bodyText
                                              htmlBrowser.Width = 600
                                              htmlBrowser.Height = 200
                                              htmlBrowser.Location = New System.Drawing.Point(20, 20)
                                              htmlBrowser.Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right Or System.Windows.Forms.AnchorStyles.Bottom
                                              htmlBrowser.BackColor = HTMLMessageForm.BackColor

                                              AddHandler htmlBrowser.DocumentCompleted, Sub(sender, e)
                                                                                            If htmlBrowser.Document IsNot Nothing AndAlso htmlBrowser.Document.Body IsNot Nothing Then
                                                                                                htmlBrowser.Document.Body.Style = "font-family: 'Segoe UI'; font-size: 9pt; margin: 0px;"
                                                                                            End If
                                                                                        End Sub

                                              HTMLMessageForm.Controls.Add(htmlBrowser)

                                              ' OK Button
                                              okButton.Text = "OK"
                                              Dim okButtonSize As System.Drawing.Size = TextRenderer.MeasureText(okButton.Text, standardFont)
                                              okButton.AutoSize = True
                                              okButton.Location = New System.Drawing.Point(20, HTMLMessageForm.ClientSize.Height - okButton.Height - 20)
                                              okButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left
                                              HTMLMessageForm.Controls.Add(okButton)

                                              AddHandler okButton.Click, Sub(sender, e)
                                                                             HTMLMessageForm.Close()
                                                                             HTMLMessageForm = Nothing
                                                                         End Sub

                                              AddHandler HTMLMessageForm.KeyDown, Sub(sender, e)
                                                                                      If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                                                          HTMLMessageForm.Close()
                                                                                          HTMLMessageForm = Nothing
                                                                                          e.SuppressKeyPress = True
                                                                                      End If
                                                                                  End Sub

                                              AddHandler HTMLMessageForm.Shown, Sub(sender, e)
                                                                                    HTMLMessageForm.Activate()
                                                                                End Sub

                                              AddHandler htmlBrowser.DocumentCompleted, Sub(sender, e)
                                                                                            If htmlBrowser.Document IsNot Nothing AndAlso htmlBrowser.Document.Body IsNot Nothing Then
                                                                                                ' Get the form's background color
                                                                                                Dim formBackColor As System.Drawing.Color = HTMLMessageForm.BackColor
                                                                                                ' Convert it to an RGB string (you can also use a hex string if you prefer)
                                                                                                Dim rgbValue As String = $"rgb({formBackColor.R}, {formBackColor.G}, {formBackColor.B})"
                                                                                                ' Apply the style to the HTML document's body, including any other styles you need
                                                                                                htmlBrowser.Document.Body.Style = $"background-color: {rgbValue}; font-family: 'Segoe UI'; font-size: 9pt; margin: 0px;"
                                                                                            End If
                                                                                        End Sub

                                              Dim formWidth As Integer = Math.Max(650, htmlBrowser.Width + 40)
                                              Dim formHeight As Integer = Math.Max(335, htmlBrowser.Height + okButton.Height + 80)
                                              HTMLMessageForm.ClientSize = New System.Drawing.Size(formWidth, formHeight)
                                              okButton.Location = New System.Drawing.Point(20, htmlBrowser.Bottom + 20)

                                              ' Optionally, if you wish to set an owner, you could get Outlook's handle here.
                                              ' Otherwise, simply run the dialog:
                                              HTMLMessageForm.Opacity = 1
                                              HTMLMessageForm.ShowDialog()
                                          End Sub)
            t.SetApartmentState(Threading.ApartmentState.STA)
            t.Start()
        End Sub



        Public Class InputParameter
            Public Property Name As String
            Public Property Value As Object
            ' We use this property to keep track of the dynamically created control.
            Public Property InputControl As Control

            Public Sub New(ByVal name As String, ByVal value As Object)
                Me.Name = name
                Me.Value = value
            End Sub
        End Class


        Public Shared Function ShowCustomVariableInputForm(ByVal prompt As String, ByVal header As String, ByRef params() As InputParameter) As Boolean
            ' Create a new form and set its basic properties.
            If String.IsNullOrWhiteSpace(header) Then header = AN
            Dim inputForm As New Form()
            inputForm.Font = New Drawing.Font("Segoe UI", 9)
            inputForm.Text = header
            ' Set the icon from resources:
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            inputForm.Icon = Icon.FromHandle(bmp.GetHicon())
            inputForm.StartPosition = FormStartPosition.CenterScreen
            inputForm.FormBorderStyle = FormBorderStyle.FixedDialog
            inputForm.MaximizeBox = False
            inputForm.MinimizeBox = False

            ' Create a prompt label at the top (autosize).
            Dim promptLabel As New System.Windows.Forms.Label()
            promptLabel.MaximumSize = New Size(400, 0) ' 400 for width; 0 means no limit on height.
            promptLabel.AutoSize = True
            promptLabel.Text = prompt
            promptLabel.Location = New Drawing.Point(15, 10)
            inputForm.Controls.Add(promptLabel)

            ' We'll use currentY to keep track of the vertical position.
            Dim currentY As Integer = promptLabel.Bottom + 10
            Dim widthneeded As Integer = 0

            ' For each parameter, create a label and an input control.
            For Each param As InputParameter In params
                Dim paramLabel As New System.Windows.Forms.Label()
                Dim textSize As Size = TextRenderer.MeasureText(param.Name & ":", paramLabel.Font)
                Dim labelwidth As Integer = textSize.Width
                If labelwidth > widthneeded Then widthneeded = labelwidth
            Next

            For Each param As InputParameter In params
                ' Create a label for this parameter.
                Dim paramLabel As New System.Windows.Forms.Label()
                paramLabel.AutoSize = True
                paramLabel.Text = param.Name & ":"
                paramLabel.Location = New Drawing.Point(15, currentY)
                inputForm.Controls.Add(paramLabel)

                ' Create the input control.
                Dim inputControl As Control = Nothing
                If TypeOf param.Value Is Boolean Then
                    ' For Boolean values use a CheckBox.
                    Dim chk As New System.Windows.Forms.CheckBox()
                    chk.Checked = Convert.ToBoolean(param.Value)
                    chk.AutoSize = True
                    ' Place the checkbox at a fixed X position.
                    chk.Location = New Drawing.Point(widthneeded + 50, currentY)
                    inputControl = chk
                Else
                    ' For other types use a TextBox.
                    Dim txt As New TextBox()
                    txt.Text = param.Value.ToString()
                    ' Set width depending on the type.
                    If TypeOf param.Value Is Integer OrElse TypeOf param.Value Is Double Then
                        txt.Width = 100
                    ElseIf TypeOf param.Value Is String Then
                        txt.Width = 350
                    Else
                        txt.Width = 350
                    End If
                    txt.Location = New Drawing.Point(widthneeded + 50, currentY - 3)
                    inputControl = txt
                End If
                inputForm.Controls.Add(inputControl)
                ' Save a reference to the control so we can retrieve its value later.
                param.InputControl = inputControl

                ' Increase currentY for the next parameter
                currentY += Math.Max(paramLabel.Height, inputControl.Height) + 5
            Next

            currentY += 10

            ' Add OK and Cancel buttons.
            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.AutoSize = True
            btnOK.DialogResult = DialogResult.OK
            btnOK.Location = New Drawing.Point(15, currentY)
            inputForm.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "Cancel"
            btnCancel.AutoSize = True
            btnCancel.DialogResult = DialogResult.Cancel
            btnCancel.Location = New Drawing.Point(btnOK.Right + 15, currentY)
            inputForm.Controls.Add(btnCancel)

            ' Set AcceptButton and CancelButton.
            inputForm.AcceptButton = btnOK
            inputForm.CancelButton = btnCancel

            ' Adjust the form’s client size dynamically based on the contained controls.
            Dim maxWidth As Integer = 0
            For Each ctrl As Control In inputForm.Controls
                If ctrl.Right > maxWidth Then
                    maxWidth = ctrl.Right
                End If
            Next
            promptLabel.MaximumSize = New Size(maxWidth - 15 - 15, 0)
            inputForm.ClientSize = New Size(maxWidth + 15, btnOK.Bottom + 15)

            Dim Returnvalue As Boolean = False

            ' Show the form modally.
            If inputForm.ShowDialog() = DialogResult.OK Then
                ' For each parameter update its Value property from the associated control.
                For Each param As InputParameter In params
                    If TypeOf param.Value Is Boolean Then
                        Dim chk As System.Windows.Forms.CheckBox = CType(param.InputControl, System.Windows.Forms.CheckBox)
                        param.Value = chk.Checked
                    ElseIf TypeOf param.Value Is Integer Then
                        Dim txt As TextBox = CType(param.InputControl, TextBox)
                        Dim newVal As Integer
                        If Integer.TryParse(txt.Text, newVal) Then
                            param.Value = newVal
                        Else
                            ShowCustomMessageBox($"Invalid value for {param.Name}. Will use the original value ('{param.Value}').")
                        End If
                    ElseIf TypeOf param.Value Is Double Then
                        Dim txt As TextBox = CType(param.InputControl, TextBox)
                        Dim newVal As Double
                        If Double.TryParse(txt.Text, newVal) Then
                            param.Value = newVal
                        Else
                            ShowCustomMessageBox($"Invalid value for {param.Name}. Will use the original value ('{param.Value}').")
                        End If
                    ElseIf TypeOf param.Value Is String Then
                        Dim txt As TextBox = CType(param.InputControl, TextBox)
                        param.Value = txt.Text
                    Else
                        ' For any other type, simply assign the text.
                        Dim txt As TextBox = CType(param.InputControl, TextBox)
                        param.Value = txt.Text
                    End If
                Next
                Returnvalue = True
            Else
                Returnvalue = False
            End If

            inputForm.Dispose()
            Return returnvalue
        End Function



        Public Shared Function ShowCustomWindow(introLine As String, ByVal bodyText As String, finalRemark As String, header As String, Optional NoRTF As Boolean = False, Optional Getfocus As Boolean = False) As String

            Dim OriginalText As String = bodyText
            Dim styledForm As New System.Windows.Forms.Form()
            Dim introLabel As New System.Windows.Forms.Label()
            Dim bodyTextBox As New System.Windows.Forms.RichTextBox()
            Dim finalRemarkLabel As New System.Windows.Forms.Label()
            Dim editedButton As New System.Windows.Forms.Button()
            Dim originalButton As New System.Windows.Forms.Button()
            Dim cancelButton As New System.Windows.Forms.Button()
            Dim toolStrip As New System.Windows.Forms.ToolStrip()

            ' Get screen dimensions
            Dim screenWidth As Integer = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width
            Dim screenHeight As Integer = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height

            ' Calculate maximum dimensions maintaining a 16:9 aspect ratio
            Dim maxWidth As Integer = screenWidth \ 2
            Dim maxHeight As Integer = Math.Min(screenHeight \ 2, (maxWidth * 9) \ 16)
            maxWidth = Math.Min(maxWidth, (maxHeight * 16) \ 9)

            ' Set minimum dimensions
            Dim minWidth As Integer = 400
            Dim minHeight As Integer = 300

            ' Form attributes
            styledForm.Text = header
            styledForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            styledForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            styledForm.MaximizeBox = False
            styledForm.MinimizeBox = False
            styledForm.ShowInTaskbar = False
            styledForm.TopMost = True
            styledForm.CancelButton = cancelButton

            ' Set the icon
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo) ' Add your logo to resources
            styledForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Set predefined font for consistent layout
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            styledForm.Font = standardFont

            ' Intro line label
            introLabel.Text = introLine
            introLabel.Font = standardFont
            introLabel.AutoSize = True
            introLabel.MaximumSize = New System.Drawing.Size(maxWidth - 40, 0)
            introLabel.Location = New System.Drawing.Point(10, 10)
            styledForm.Controls.Add(introLabel)

            ' Final remark label
            If Not String.IsNullOrEmpty(finalRemark) Then
                finalRemarkLabel.Text = finalRemark
                finalRemarkLabel.Font = standardFont
                finalRemarkLabel.AutoSize = True
                finalRemarkLabel.MaximumSize = New System.Drawing.Size(maxWidth - 40, 0)
            End If

            ' Edited Button
            editedButton.Text = "OK, use edited text"
            Dim buttonSize As System.Drawing.Size = System.Windows.Forms.TextRenderer.MeasureText(editedButton.Text, standardFont)
            Dim buttonWidth As Integer = buttonSize.Width + 20
            Dim buttonHeight As Integer = buttonSize.Height + 10
            editedButton.Size = New System.Drawing.Size(buttonWidth, buttonHeight)

            ' Original Button
            originalButton.Text = "OK, use original text"
            Dim originalButtonSize As System.Drawing.Size = System.Windows.Forms.TextRenderer.MeasureText(originalButton.Text, standardFont)
            Dim originalButtonWidth As Integer = originalButtonSize.Width + 20
            originalButton.Size = New System.Drawing.Size(originalButtonWidth, buttonHeight)

            ' Cancel Button
            cancelButton.Text = "Cancel"
            Dim cancelButtonSize As System.Drawing.Size = System.Windows.Forms.TextRenderer.MeasureText(cancelButton.Text, standardFont)
            Dim cancelButtonWidth As Integer = cancelButtonSize.Width + 20
            cancelButton.Size = New System.Drawing.Size(cancelButtonWidth, buttonHeight)

            ' Calculate space needed for static elements
            Dim staticElementsHeight As Integer = introLabel.Height + 20 + buttonHeight + 60
            If Not String.IsNullOrEmpty(finalRemark) Then
                staticElementsHeight += finalRemarkLabel.Height + 10
            End If

            ' Calculate remaining space for body text
            Dim bodyTextHeight As Integer = Math.Max(minHeight - staticElementsHeight, maxHeight - staticElementsHeight - 20)

            ' Body text box
            bodyTextBox.Font = New System.Drawing.Font("Segoe UI", 10.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            bodyTextBox.Multiline = True
            bodyTextBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
            bodyTextBox.WordWrap = True
            bodyTextBox.Location = New System.Drawing.Point(10, introLabel.Bottom + 10)
            bodyTextBox.Width = maxWidth - 40
            bodyTextBox.Height = bodyTextHeight - 20

            ' ToolStrip for text formatting
            toolStrip.Dock = System.Windows.Forms.DockStyle.None
            toolStrip.Location = New System.Drawing.Point(bodyTextBox.Right - 130, bodyTextBox.Top - 35)

            Dim boldButton As New System.Windows.Forms.ToolStripButton("B") With {.Font = New System.Drawing.Font(standardFont, FontStyle.Bold)}
            AddHandler boldButton.Click, Sub(sender, e)
                                             If bodyTextBox.SelectionLength > 0 Then
                                                 bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor FontStyle.Bold)
                                             End If
                                         End Sub

            Dim italicButton As New System.Windows.Forms.ToolStripButton("I") With {.Font = New System.Drawing.Font(standardFont, FontStyle.Italic)}
            AddHandler italicButton.Click, Sub(sender, e)
                                               If bodyTextBox.SelectionLength > 0 Then
                                                   bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor FontStyle.Italic)
                                               End If
                                           End Sub

            Dim underlineButton As New System.Windows.Forms.ToolStripButton("U") With {.Font = New System.Drawing.Font(standardFont, FontStyle.Underline)}
            AddHandler underlineButton.Click, Sub(sender, e)
                                                  If bodyTextBox.SelectionLength > 0 Then
                                                      bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor FontStyle.Underline)
                                                  End If
                                              End Sub

            Dim bulletButton As New System.Windows.Forms.ToolStripButton("•") With {.Font = New System.Drawing.Font(standardFont, FontStyle.Regular)}
            AddHandler bulletButton.Click, Sub(sender, e)
                                               bodyTextBox.SelectionIndent = If(bodyTextBox.SelectionIndent = 20, 0, 20)
                                               bodyTextBox.SelectionBullet = Not bodyTextBox.SelectionBullet
                                               bodyTextBox.BulletIndent = If(bodyTextBox.BulletIndent = 15, 0, 15)
                                           End Sub

            toolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {boldButton, italicButton, underlineButton, bulletButton})
            styledForm.Controls.Add(toolStrip)

            ' Format text: Bold between ** and ** or {{ and }}
            bodyTextBox.Text = bodyText

            bodyTextBox.AppendText(" ")
            bodyTextBox.Select(bodyTextBox.TextLength - 1, 1)
            bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, FontStyle.Bold)
            bodyTextBox.DeselectAll()

            bodyText = bodyTextBox.Rtf

            Dim pattern As String = "(\*\*.*?\*\*|\{\{.*?\}\})"
            Dim matches As MatchCollection = Regex.Matches(bodyText, pattern)
            Dim newbody = bodyText

            ' Process matches in reverse order
            For Each match As Match In matches.Cast(Of Match).Reverse()
                Dim startIndex As Integer = match.Index
                Dim boldText As String = match.Value
                Dim plainText As String = boldText.Substring(2, boldText.Length - 4) ' Remove ** or {{ }}

                ' Apply bold formatting
                newbody = newbody.Remove(startIndex, boldText.Length).Insert(startIndex, "\b\f1 " & plainText & "\b0\f0 ")

            Next

            bodyTextBox.Rtf = newbody

            ' Deselect any text
            bodyTextBox.Select(0, 0)

            styledForm.Controls.Add(bodyTextBox)

            ' Position final remark label and buttons
            If Not String.IsNullOrEmpty(finalRemark) Then
                finalRemarkLabel.Location = New System.Drawing.Point(10, bodyTextBox.Bottom + 10)
                styledForm.Controls.Add(finalRemarkLabel)
                editedButton.Location = New System.Drawing.Point((maxWidth - buttonWidth - originalButtonWidth - cancelButtonWidth - 40) / 2, finalRemarkLabel.Bottom + 20)
                originalButton.Location = New System.Drawing.Point(editedButton.Right + 10, finalRemarkLabel.Bottom + 20)
                cancelButton.Location = New System.Drawing.Point(originalButton.Right + 10, finalRemarkLabel.Bottom + 20)
            Else
                editedButton.Location = New System.Drawing.Point((maxWidth - buttonWidth - originalButtonWidth - cancelButtonWidth - 40) / 2, bodyTextBox.Bottom + 20)
                originalButton.Location = New System.Drawing.Point(editedButton.Right + 10, bodyTextBox.Bottom + 20)
                cancelButton.Location = New System.Drawing.Point(originalButton.Right + 10, bodyTextBox.Bottom + 20)
            End If

            styledForm.Controls.Add(editedButton)
            styledForm.Controls.Add(originalButton)
            styledForm.Controls.Add(cancelButton)

            ' Event handlers for buttons
            Dim returnValue As String = String.Empty

            AddHandler editedButton.Click, Sub(sender, e)
                                               If NoRTF Then
                                                   returnValue = bodyTextBox.Text
                                               Else
                                                   returnValue = bodyTextBox.Rtf
                                               End If
                                               styledForm.DialogResult = DialogResult.OK
                                               styledForm.Close()
                                           End Sub

            AddHandler originalButton.Click, Sub(sender, e)
                                                 If NoRTF Then
                                                     returnValue = OriginalText
                                                 Else
                                                     returnValue = bodyText
                                                 End If
                                                 styledForm.DialogResult = DialogResult.OK
                                                 styledForm.Close()
                                             End Sub

            AddHandler cancelButton.Click, Sub(sender, e)
                                               returnValue = String.Empty
                                               styledForm.DialogResult = DialogResult.Cancel
                                               styledForm.Close()
                                           End Sub

            ' Adjust form size dynamically
            Dim formWidth As Integer = Math.Max(minWidth, Math.Min(maxWidth, bodyTextBox.Width + 20))
            Dim formHeight As Integer = Math.Max(minHeight, Math.Min(maxHeight, staticElementsHeight + bodyTextBox.Height + 100))
            styledForm.ClientSize = New System.Drawing.Size(formWidth, formHeight)

            ' Show dialog and return appropriate value

            styledForm.TopMost = True
            styledForm.BringToFront()
            styledForm.Activate()
            styledForm.Focus()
            If Getfocus Then
                Dim outlookHwnd As IntPtr = FindWindow("rctrl_renwnd32", Nothing)
                Dim Result = styledForm.ShowDialog(New WindowWrapper(outlookHwnd))
            Else
                styledForm.ShowDialog()
            End If
            Return returnValue
        End Function

        Public Shared Function MissingSettingsWindow(Settings As Dictionary(Of String, String), context As ISharedContext) As Boolean

            ' Create the form
            Dim settingsForm As New System.Windows.Forms.Form()
            settingsForm.Text = $"{AN} Settings"
            settingsForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            settingsForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            settingsForm.MaximizeBox = False
            settingsForm.MinimizeBox = False
            settingsForm.ShowInTaskbar = False
            settingsForm.TopMost = True

            ' Set the icon
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            settingsForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Set a predefined font for consistent layout
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            settingsForm.Font = standardFont

            ' Add description label
            Dim descriptionLabel As New System.Windows.Forms.Label()
            descriptionLabel.Text = "Complete missing mandatory values:"
            descriptionLabel.AutoSize = True
            descriptionLabel.Location = New System.Drawing.Point(10, 20)
            settingsForm.Controls.Add(descriptionLabel)

            ' Define controls for labels and inputs
            Dim labelControls As New Dictionary(Of String, System.Windows.Forms.Label)
            Dim settingControls As New Dictionary(Of String, System.Windows.Forms.Control)

            ' Dynamically calculate label width
            Dim maxLabelWidth As Integer = 0

            ' Calculate maximum label width
            For Each setting In Settings
                Dim textSize As System.Drawing.Size = TextRenderer.MeasureText(setting.Value & ":", standardFont)
                maxLabelWidth = Math.Max(maxLabelWidth, textSize.Width)
            Next

            Dim controlXOffset As Integer = maxLabelWidth + 20
            Dim defaultControlWidth As Integer = 240
            Dim lineSpacing As Integer = CInt(TextRenderer.MeasureText("Sample", standardFont).Height * 1.5)
            Dim yPos As Integer = descriptionLabel.Bottom + 20

            ' Add labels and input controls
            For Each setting In Settings
                Dim label As New System.Windows.Forms.Label()
                If context.INI_SecondAPI Then
                    label.Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", context.INI_Model_2) & ":"
                Else
                    label.Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", "2nd model (none)") & ":"
                End If
                label.AutoSize = True
                label.Font = standardFont
                label.Location = New System.Drawing.Point(10, yPos)
                settingsForm.Controls.Add(label)
                labelControls.Add(setting.Key, label)

                If IsBooleanSetting(setting.Key) Then
                    Dim checkBox As New System.Windows.Forms.CheckBox()
                    checkBox.Checked = Boolean.Parse(GetSettingValue(setting.Key, context))
                    checkBox.Location = New System.Drawing.Point(controlXOffset, yPos)
                    settingsForm.Controls.Add(checkBox)
                    settingControls.Add(setting.Key, checkBox)
                Else
                    Dim textBox As New System.Windows.Forms.TextBox()
                    textBox.Text = GetSettingValue(setting.Key, context)
                    textBox.Size = New System.Drawing.Size(defaultControlWidth, 20)
                    textBox.Location = New System.Drawing.Point(controlXOffset, yPos)
                    settingsForm.Controls.Add(textBox)
                    settingControls.Add(setting.Key, textBox)
                End If

                yPos += lineSpacing
            Next

            ' Add buttons
            Dim buttonYPos As Integer = yPos + 20
            Dim buttonSpacing As Integer = 10

            Dim okButton As New System.Windows.Forms.Button()
            okButton.Text = "Save and continue"
            Dim okButtonSize As System.Drawing.Size = TextRenderer.MeasureText(okButton.Text, standardFont)
            okButton.Size = New System.Drawing.Size(okButtonSize.Width + 20, okButtonSize.Height + 10)
            okButton.Location = New System.Drawing.Point(10, buttonYPos)
            settingsForm.Controls.Add(okButton)

            Dim okButtonToolTip As New System.Windows.Forms.ToolTip()
            okButtonToolTip.SetToolTip(okButton, $"Will save the exisiting values and those you have entered into a local copy of '{AN2}.ini' (overwriting any existing such file).")

            Dim cancelButton As New System.Windows.Forms.Button()
            cancelButton.Text = "Cancel"
            Dim cancelButtonSize As System.Drawing.Size = TextRenderer.MeasureText(cancelButton.Text, standardFont)
            cancelButton.Size = New System.Drawing.Size(cancelButtonSize.Width + 20, cancelButtonSize.Height + 10)
            cancelButton.Location = New System.Drawing.Point(okButton.Right + buttonSpacing, buttonYPos)
            settingsForm.Controls.Add(cancelButton)

            Dim cancelButtonToolTip As New System.Windows.Forms.ToolTip()
            cancelButtonToolTip.SetToolTip(cancelButton, $"{AN} will not operate properly until you have provided the necessary configuration parameters. You can retry later.")

            ' Flag to track whether the user completed the form
            Dim userCompleted As Boolean = False

            ' Attach handlers to buttons
            AddHandler okButton.Click, Sub(sender, e)
                                           For Each settingKey In settingControls.Keys
                                               Dim control = settingControls(settingKey)
                                               If TypeOf control Is System.Windows.Forms.TextBox Then
                                                   ' Handle TextBox settings
                                                   Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                   SetSettingValue(settingKey, textValue, context)
                                               ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                   ' Handle CheckBox settings
                                                   Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                   SetSettingValue(settingKey, boolValue.ToString(), context)
                                               Else
                                                   MessageBox.Show($"Error in MissingSettingsWindow - unsupported control type for setting '{settingKey}' in MissingSettingsWindow.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                               End If
                                           Next
                                           UpdateAppConfig(context) ' Save the configuration
                                           userCompleted = True
                                           settingsForm.Close()
                                       End Sub

            AddHandler cancelButton.Click, Sub(sender, e)
                                               settingsForm.Close()
                                           End Sub

            ' Adjust form size dynamically
            settingsForm.ClientSize = New System.Drawing.Size(controlXOffset + defaultControlWidth + 40, cancelButton.Bottom + 20)

            ' Show the form and wait for user input
            settingsForm.ShowDialog()

            ' Return whether the user completed the form
            Return userCompleted
        End Function

        Public Shared Sub ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String), ByRef context As ISharedContext)

            InitializeConfig(context, False, False)

            If context.INIloaded = False Then Exit Sub

            ' Create the form
            Dim settingsForm As New System.Windows.Forms.Form()
            settingsForm.Text = $"{AN} Settings"
            settingsForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            settingsForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            settingsForm.MaximizeBox = False
            settingsForm.MinimizeBox = False
            settingsForm.ShowInTaskbar = False
            settingsForm.TopMost = True

            ' Set the icon
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            settingsForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Set a predefined font for consistent layout
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            settingsForm.Font = standardFont

            ' Add description label
            Dim descriptionLabel As New System.Windows.Forms.Label()
            descriptionLabel.Text = "You can temporarily change the following values (save to keep them):"
            descriptionLabel.AutoSize = True
            descriptionLabel.Location = New System.Drawing.Point(10, 20)
            settingsForm.Controls.Add(descriptionLabel)

            ' Define controls for labels and inputs
            Dim labelControls As New Dictionary(Of String, System.Windows.Forms.Label)
            Dim settingControls As New Dictionary(Of String, System.Windows.Forms.Control)

            ' Dynamically calculate label width
            Dim maxLabelWidth As Integer = 0


            ' Calculate maximum label width
            For Each setting In Settings
                Dim textSize As System.Drawing.Size
                If context.INI_SecondAPI Then
                    textSize = TextRenderer.MeasureText(setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", context.INI_Model_2) & ":", standardFont)
                Else
                    textSize = TextRenderer.MeasureText(setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", "2nd model (none)") & ":", standardFont)
                End If

                maxLabelWidth = Math.Max(maxLabelWidth, textSize.Width)
            Next

            Dim controlXOffset As Integer = maxLabelWidth + 20
            Dim defaultControlWidth As Integer = 350  '240
            Dim lineSpacing As Integer = CInt(TextRenderer.MeasureText("Sample", standardFont).Height * 1.5)
            Dim yPos As Integer = descriptionLabel.Bottom + 20

            ' Add labels and input controls
            For Each setting In Settings
                Dim label As New System.Windows.Forms.Label()
                Dim ToolTip As New System.Windows.Forms.ToolTip()
                If context.INI_SecondAPI Then
                    label.Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", context.INI_Model_2) & ":"
                Else
                    label.Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", "2nd model (none)") & ":"
                End If
                label.AutoSize = True
                label.Font = standardFont
                label.Location = New System.Drawing.Point(10, yPos)
                settingsForm.Controls.Add(label)
                labelControls.Add(setting.Key, label)
                Dim ToolTipText As String = SettingsTips(setting.Key)
                ToolTip.SetToolTip(label, ToolTipText)

                If IsBooleanSetting(setting.Key) Then
                    Dim checkBox As New System.Windows.Forms.CheckBox()
                    checkBox.Checked = Boolean.Parse(GetSettingValue(setting.Key, context))
                    checkBox.Location = New System.Drawing.Point(controlXOffset, yPos)
                    If setting.Key.Contains("_2") And Not context.INI_SecondAPI Then
                        checkBox.Enabled = False
                    Else
                        checkBox.Enabled = True
                    End If
                    settingsForm.Controls.Add(checkBox)
                    settingControls.Add(setting.Key, checkBox)
                    ToolTip.SetToolTip(checkBox, ToolTipText)
                Else
                    Dim textBox As New System.Windows.Forms.TextBox()
                    textBox.Text = GetSettingValue(setting.Key, context)
                    textBox.Size = New System.Drawing.Size(defaultControlWidth, 20)
                    textBox.Location = New System.Drawing.Point(controlXOffset, yPos)
                    If setting.Key.Contains("_2") And Not context.INI_SecondAPI Then
                        textBox.Enabled = False
                    Else
                        textBox.Enabled = True
                    End If
                    settingsForm.Controls.Add(textBox)
                    settingControls.Add(setting.Key, textBox)
                    ToolTip.SetToolTip(textBox, ToolTipText)
                End If

                If setting.Key.Contains("_2") Then

                End If
                yPos += lineSpacing
            Next

            ' Add buttons
            Dim buttonYPos As Integer = yPos + 20
            Dim buttonSpacing As Integer = 10

            Dim switchButton As New System.Windows.Forms.Button()
            switchButton.Text = "Switch Model"
            Dim switchButtonSize As System.Drawing.Size = TextRenderer.MeasureText(switchButton.Text, standardFont)
            switchButton.Size = New System.Drawing.Size(switchButtonSize.Width + 20, switchButtonSize.Height + 10)
            switchButton.Location = New System.Drawing.Point(10, buttonYPos)
            switchButton.Enabled = context.INI_SecondAPI
            settingsForm.Controls.Add(switchButton)

            Dim SwitchButtonToolTip As New System.Windows.Forms.ToolTip()
            SwitchButtonToolTip.SetToolTip(switchButton, "Will accept the current settings and switch the primary model with the secondary model.")

            Dim expertConfigButton As New System.Windows.Forms.Button()
            expertConfigButton.Text = "Expert Config"
            Dim expertButtonSize As System.Drawing.Size = TextRenderer.MeasureText(expertConfigButton.Text, standardFont)
            expertConfigButton.Size = New System.Drawing.Size(expertButtonSize.Width + 20, expertButtonSize.Height + 10)
            expertConfigButton.Location = New System.Drawing.Point(switchButton.Right + buttonSpacing, buttonYPos)
            settingsForm.Controls.Add(expertConfigButton)

            Dim expertConfigButtonToolTip As New System.Windows.Forms.ToolTip()
            expertConfigButtonToolTip.SetToolTip(expertConfigButton, $"Will accept the current settings and in a separate window let you amend all configuration variables from '{AN2}.ini'.")


            Dim saveConfigButton As New System.Windows.Forms.Button()
            saveConfigButton.Text = "Save Configuration"
            Dim saveButtonSize As System.Drawing.Size = TextRenderer.MeasureText(saveConfigButton.Text, standardFont)
            saveConfigButton.Size = New System.Drawing.Size(saveButtonSize.Width + 20, saveButtonSize.Height + 10)
            saveConfigButton.Location = New System.Drawing.Point(expertConfigButton.Right + buttonSpacing, buttonYPos)
            settingsForm.Controls.Add(saveConfigButton)

            Dim saveConfigToolTip As New System.Windows.Forms.ToolTip()
            saveConfigToolTip.SetToolTip(saveConfigButton, $"Will save the current configuration to a local copy of '{AN2}.ini' (overwriting any existing such file).")

            Dim CentralConfigAvailable As Boolean = System.IO.File.Exists(System.IO.Path.Combine(ExpandEnvironmentVariables(GetFromRegistry(RegPath_Base, RegPath_IniPath, True)), $"{AN2}.ini"))
            Dim delLocalConfigButton As New System.Windows.Forms.Button()
            If CentralConfigAvailable Then
                delLocalConfigButton.Text = "Give Up Local Config"
            Else
                delLocalConfigButton.Text = "Reset Optional Values"
            End If
            Dim delLocalButtonSize As System.Drawing.Size = TextRenderer.MeasureText(delLocalConfigButton.Text, standardFont)
            delLocalConfigButton.Size = New System.Drawing.Size(delLocalButtonSize.Width + 20, delLocalButtonSize.Height + 10)
            delLocalConfigButton.Location = New System.Drawing.Point(saveConfigButton.Right + buttonSpacing, buttonYPos)
            settingsForm.Controls.Add(delLocalConfigButton)

            Dim delLocalConfigToolTip As New System.Windows.Forms.ToolTip()
            If CentralConfigAvailable Then
                If Left(context.RDV, 4) = "Word" Then
                    delLocalConfigToolTip.SetToolTip(delLocalConfigButton, $"This will deactivate the local configuration in '{AN2}.ini' (by renaming it to '.bak', overwriting any existing such file) and have the central configuration file applied going forward.")
                Else
                    delLocalConfigToolTip.SetToolTip(delLocalConfigButton, $"This will deactivate the local configuration in '{AN2}.ini' (by renaming it to '.bak', overwriting any existing such file), and have the configuration file of your 'Word' add-in (if available) and otherwise the central one applied going forward.")
                End If
            Else
                delLocalConfigToolTip.SetToolTip(delLocalConfigButton, $"This will reset all parameters that are not mandatory by removing them from your local configuration file '{AN2}.ini'. A copy will be saved beforhand to '.bak', overwriting any existing such file.")
            End If

            Dim okButton As New System.Windows.Forms.Button()
            okButton.Text = "OK"
            Dim okButtonSize As System.Drawing.Size = TextRenderer.MeasureText(okButton.Text, standardFont)
            okButton.Size = New System.Drawing.Size(okButtonSize.Width + 20, okButtonSize.Height + 10)
            okButton.Location = New System.Drawing.Point(10, buttonYPos + 50)
            settingsForm.Controls.Add(okButton)

            Dim cancelButton As New System.Windows.Forms.Button()
            cancelButton.Text = "Cancel"
            Dim cancelButtonSize As System.Drawing.Size = TextRenderer.MeasureText(cancelButton.Text, standardFont)
            cancelButton.Size = New System.Drawing.Size(cancelButtonSize.Width + 20, cancelButtonSize.Height + 10)
            cancelButton.Location = New System.Drawing.Point(okButton.Right + buttonSpacing, buttonYPos + 50)
            settingsForm.Controls.Add(cancelButton)

            Dim aboutButton As New System.Windows.Forms.Button()
            aboutButton.Text = $"About {AN}"
            Dim aboutButtonSize As System.Drawing.Size = TextRenderer.MeasureText(aboutButton.Text, standardFont)
            aboutButton.Size = New System.Drawing.Size(aboutButtonSize.Width + 20, aboutButtonSize.Height + 10)
            aboutButton.Location = New System.Drawing.Point(cancelButton.Right + buttonSpacing, cancelButton.Top)
            settingsForm.Controls.Add(aboutButton)

            Dim RightSide As Integer = aboutButton.Right

            Dim updateButton As New System.Windows.Forms.Button()
            updateButton.Text = "Check for Updates"
            If Not String.IsNullOrWhiteSpace(context.INI_UpdatePath) Then
                updateButton.Text = "Do local update"
            End If
            Dim updateButtonSize As System.Drawing.Size = TextRenderer.MeasureText(updateButton.Text, standardFont)
            updateButton.Size = New System.Drawing.Size(updateButtonSize.Width + 20, updateButtonSize.Height + 10)
            updateButton.Location = New System.Drawing.Point(aboutButton.Right + buttonSpacing, cancelButton.Top)
            If ApplicationDeployment.IsNetworkDeployed Or Not String.IsNullOrWhiteSpace(context.INI_UpdatePath) Then
                settingsForm.Controls.Add(updateButton)
                RightSide = updateButton.Right
            End If

            Dim FilePath As String = ""
            Dim IsExcel As Boolean = True
            If context.RDV.Contains("Word") Then
                FilePath = ExpandEnvironmentVariables(HelperPaths("Word"))
                IsExcel = False
            ElseIf context.RDV.Contains("Excel") Then
                FilePath = ExpandEnvironmentVariables(HelperPaths("Excel"))
            End If
            Debug.WriteLine("Filepath=" & FilePath)

            Dim helperButton As New System.Windows.Forms.Button()
            If Not String.IsNullOrEmpty(FilePath) Then
                If File.Exists(FilePath) Then
                    helperButton.Text = "Remove Helper"
                Else
                    helperButton.Text = "Install Helper"
                End If
                Dim HelperButtonSize As System.Drawing.Size = TextRenderer.MeasureText(helperButton.Text, standardFont)
                helperButton.Size = New System.Drawing.Size(HelperButtonSize.Width + 20, HelperButtonSize.Height + 10)
                helperButton.Location = New System.Drawing.Point(RightSide + buttonSpacing, cancelButton.Top)
                settingsForm.Controls.Add(helperButton)
            End If
            Dim CapturedContext As ISharedContext = context

            ' Attach handlers to buttons
            AddHandler switchButton.Click, Sub(sender, e)
                                               If CapturedContext.INI_SecondAPI Then
                                                   For Each settingKey In settingControls.Keys
                                                       Dim control = settingControls(settingKey)
                                                       If TypeOf control Is System.Windows.Forms.TextBox Then
                                                           ' Handle TextBox settings
                                                           Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                           SetSettingValue(settingKey, textValue, CapturedContext)
                                                       ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                           ' Handle CheckBox settings
                                                           Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                           SetSettingValue(settingKey, boolValue.ToString(), CapturedContext)
                                                       Else
                                                           MessageBox.Show($"Error in ShowSettingsWindow - unsupported control type for setting '{settingKey}' in ShowSettingsWindow (Switch).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                       End If
                                                   Next
                                                   SwitchModels(CapturedContext)
                                                   RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                                   switchButton.Enabled = CapturedContext.INI_SecondAPI
                                               End If
                                               CapturedContext.MenusAdded = False
                                           End Sub

            AddHandler expertConfigButton.Click, Sub(sender, e)
                                                     For Each settingKey In settingControls.Keys
                                                         Dim control = settingControls(settingKey)
                                                         If TypeOf control Is System.Windows.Forms.TextBox Then
                                                             ' Handle TextBox settings
                                                             Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                             SetSettingValue(settingKey, textValue, CapturedContext)
                                                         ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                             ' Handle CheckBox settings
                                                             Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                             SetSettingValue(settingKey, boolValue.ToString(), CapturedContext)
                                                         Else
                                                             MessageBox.Show($"Error in ShowSettingsWindow - unsupported control type for setting '{settingKey}' in ShowSettingsWindow (ExpertConfig).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                         End If
                                                     Next
                                                     ShowExpertConfiguration(CapturedContext, settingsForm)
                                                     RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                                     switchButton.Enabled = CapturedContext.INI_SecondAPI
                                                     CapturedContext.MenusAdded = False
                                                 End Sub

            AddHandler saveConfigButton.Click, Sub(sender, e)
                                                   For Each settingKey In settingControls.Keys
                                                       Dim control = settingControls(settingKey)
                                                       If TypeOf control Is System.Windows.Forms.TextBox Then
                                                           ' Handle TextBox settings
                                                           Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                           SetSettingValue(settingKey, textValue, CapturedContext)
                                                       ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                           ' Handle CheckBox settings
                                                           Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                           SetSettingValue(settingKey, boolValue.ToString(), CapturedContext)
                                                       Else
                                                           MessageBox.Show($"Error in ShowSettingsWindow - unsupported control type for setting '{settingKey}' in ShowSettingsWindow (Save).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                       End If
                                                   Next
                                                   UpdateAppConfig(CapturedContext)
                                                   CapturedContext.MenusAdded = False
                                               End Sub

            AddHandler delLocalConfigButton.Click, Sub(sender, e)
                                                       If CentralConfigAvailable Then
                                                           If ShowCustomYesNoBox($"Do you really want to deactivate your local configuration file? The file '{AN2}.ini' will be renamed to '.bak' overwriting any existing such file", "Yes", "No") = 1 Then
                                                               If RenameFileToBak(GetDefaultINIPath(CapturedContext.RDV)) Then
                                                                   ShowCustomMessageBox("Local configuration deactivated. The central configuration will be applied going forward.", "OK")
                                                                   InitializeConfig(CapturedContext, False, True)
                                                               End If
                                                           End If
                                                       Else
                                                           If ShowCustomYesNoBox($"Do you really want to reset your local configuration file by removing non-mandatory entries? The current configuration file '{AN2}.ini' will beforehand be saved to a '.bak' file overwriting any existing such file.", "Yes", "No") = 1 Then
                                                               If RenameFileToBak(GetDefaultINIPath(CapturedContext.RDV)) Then
                                                                   ResetLocalAppConfig(CapturedContext)
                                                               End If
                                                           End If
                                                       End If
                                                       RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                                       switchButton.Enabled = CapturedContext.INI_SecondAPI
                                                       CapturedContext.MenusAdded = False
                                                   End Sub

            AddHandler helperButton.Click, Async Sub(sender, e)
                                               If helperButton.Text = "Remove Helper" Then
                                                   If ShowCustomYesNoBox($"Do you really want to remove the helper file '{FilePath}' from your system? It will be unloaded and deleted. You can re-install it later.", "Yes", "No") = 1 Then
                                                       If IsExcel Then UnloadExcelAddin(ExcelHelper) Else UnloadWordAddin(WordHelper)
                                                       Try
                                                           System.IO.File.Delete(FilePath)
                                                       Catch ex As System.Exception
                                                       End Try
                                                       If System.IO.File.Exists(FilePath) Then
                                                           ShowCustomMessageBox($"The helper file could not be deleted. Try to manually delete the file '{FilePath}' after having closed the application.")
                                                       Else
                                                           ShowCustomMessageBox("The helper file was successfully deleted.")
                                                           helperButton.Text = "Install Helper"
                                                           CapturedContext.MenusAdded = False
                                                           RemoveMenu = True
                                                       End If
                                                   End If
                                               Else
                                                   If ShowCustomYesNoBox($"Do you really want to download the helper file from https://apps.vischer.com and have it installed to '{FilePath}'? Next time you start the application, it will be automatically loaded.", "Yes", "No") = 1 Then
                                                       Dim DownloadUrl As String = ""
                                                       If IsExcel Then DownloadUrl = ExcelHelperUrl Else DownloadUrl = WordHelperUrl
                                                       Try
                                                           Using client As New HttpClient()
                                                               ' Increase timeout to prevent cutoffs
                                                               client.Timeout = TimeSpan.FromMinutes(10)

                                                               ' Disable automatic decompression (prevents incomplete downloads)
                                                               client.DefaultRequestHeaders.AcceptEncoding.Clear()

                                                               ' Start the download request
                                                               Using response As HttpResponseMessage = Await client.GetAsync(DownloadUrl, HttpCompletionOption.ResponseHeadersRead)
                                                                   response.EnsureSuccessStatusCode()

                                                                   ' Create file stream for writing
                                                                   Using fileStream As FileStream = New FileStream(FilePath, FileMode.Create, FileAccess.Write, FileShare.None)
                                                                       Using httpStream As Stream = Await response.Content.ReadAsStreamAsync()
                                                                           Dim buffer(8192) As Byte ' 8KB buffer
                                                                           Dim bytesRead As Integer
                                                                           Do
                                                                               bytesRead = Await httpStream.ReadAsync(buffer, 0, buffer.Length)
                                                                               If bytesRead = 0 Then Exit Do ' Stop when finished
                                                                               Await fileStream.WriteAsync(buffer, 0, bytesRead) ' Write chunk to file
                                                                           Loop
                                                                       End Using
                                                                   End Using
                                                               End Using
                                                           End Using
                                                           ShowCustomMessageBox($"Download to '{FilePath}' completed. You must restart the application for it to be loaded.")
                                                           helperButton.Text = "Remove Helper"
                                                       Catch ex As System.Exception
                                                           ShowCustomMessageBox($"Error when downloading from '{DownloadUrl}' to '{FilePath}'. You may have to download and install the helper file manually.")
                                                       End Try

                                                   End If
                                               End If
                                               RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                               switchButton.Enabled = CapturedContext.INI_SecondAPI
                                               CapturedContext.MenusAdded = False
                                           End Sub

            AddHandler aboutButton.Click, Sub(sender, e)
                                              ShowAboutWindow(settingsForm, CapturedContext)
                                          End Sub

            If ApplicationDeployment.IsNetworkDeployed Or Not String.IsNullOrWhiteSpace(CapturedContext.INI_UpdatePath) Then

                AddHandler updateButton.Click, Sub(sender, e)
                                                   Dim updater As New UpdateHandler()
                                                   updater.CheckAndInstallUpdates(CapturedContext.RDV, CapturedContext.INI_UpdatePath)
                                               End Sub

            End If

            AddHandler okButton.Click, Sub(sender, e)
                                           For Each settingKey In settingControls.Keys
                                               Dim control = settingControls(settingKey)
                                               If TypeOf control Is System.Windows.Forms.TextBox Then
                                                   ' Handle TextBox settings
                                                   Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                   SetSettingValue(settingKey, textValue, CapturedContext)
                                               ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                   ' Handle CheckBox settings
                                                   Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                   SetSettingValue(settingKey, boolValue.ToString(), CapturedContext)
                                               Else
                                                   MessageBox.Show($"Error in ShowSettingsWindow - unsupported control type for setting '{settingKey}' in ShowSettingsWindow (OK).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                               End If
                                           Next
                                           CapturedContext.MenusAdded = False
                                           settingsForm.Close()
                                       End Sub

            AddHandler cancelButton.Click, Sub(sender, e)
                                               settingsForm.Close()
                                           End Sub

            ' Adjust form size dynamically
            settingsForm.ClientSize = New System.Drawing.Size(controlXOffset + defaultControlWidth + 40, cancelButton.Bottom + 20)

            ' Show the form
            settingsForm.ShowDialog()

        End Sub


        Public Shared Sub UnloadExcelAddin(addinName As String)
            Dim excelApp As Excel.Application = Nothing
            Try

                ' Start or get running instance of Excel
                excelApp = TryCast(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)
                If excelApp Is Nothing Then
                    excelApp = New Excel.Application()
                    excelApp.Visible = False
                End If

                For Each addin As Excel.AddIn In excelApp.AddIns2
                    If addin.FullName.ToLower().Contains(addinName.ToLower()) Then
                        Debug.WriteLine("Unloading add-in: " & addin.FullName)
                        addin.Installed = False  ' Unload the add-in
                        Marshal.ReleaseComObject(excelApp)
                        excelApp = Nothing
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                        Debug.WriteLine("Waiting for Excel to release file lock...")
                        Thread.Sleep(1000)
                        Exit For
                    End If
                Next

            Catch ex As Exception
                Debug.WriteLine("Error unloading Excel add-In: " & ex.Message)
            Finally
                If excelApp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            End Try
        End Sub

        Public Shared Sub UnloadWordAddin(addInName As String)
            Try
                ' Attempt to get the active (running) Word Application instance.
                Dim wordApp As Microsoft.Office.Interop.Word.Application = CType(Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)

                ' Iterate through all loaded AddIns in Word.
                For Each addIn As AddIn In wordApp.AddIns
                    ' Compare names in a case-insensitive manner (if desired).
                    Debug.WriteLine("Addin: " & addIn.Name)
                    If addIn.Name.Equals(addInName, StringComparison.OrdinalIgnoreCase) Then
                        ' Unload the add-in from the current Word session.
                        addIn.Installed = False
                        addIn.Delete()
                        Debug.WriteLine("Deleted!")
                        Exit For
                    End If
                Next

            Catch ex As System.Exception
                Debug.WriteLine("Error unloading Word add-in: " & ex.Message)
            End Try
        End Sub




        Public Shared Sub RefreshFormValues(settingControls As Dictionary(Of String, System.Windows.Forms.Control),
                              labelControls As Dictionary(Of String, System.Windows.Forms.Label), ByRef context As ISharedContext, Settings As Dictionary(Of String, String))
            ' Update the labels and input controls dynamically
            For Each setting In Settings
                ' Update label text
                If labelControls.ContainsKey(setting.Key) Then
                    If context.INI_SecondAPI Then
                        labelControls(setting.Key).Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", context.INI_Model_2) & ":"
                    Else
                        labelControls(setting.Key).Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", "of 2nd model (none)") & ":"
                    End If

                    ' Update input controls
                    If TypeOf settingControls(setting.Key) Is System.Windows.Forms.TextBox Then
                        settingControls(setting.Key).Text = GetSettingValue(setting.Key, context)
                    ElseIf TypeOf settingControls(setting.Key) Is System.Windows.Forms.CheckBox Then
                        DirectCast(settingControls(setting.Key), System.Windows.Forms.CheckBox).Checked = Boolean.Parse(GetSettingValue(setting.Key, context))
                    End If
                End If
            Next
        End Sub

        Public Shared Function IsBooleanSetting(settingKey As String) As Boolean
            ' Determine if a setting is a Boolean based on its key
            Dim booleanSettings As New List(Of String) From {
        "DoubleS", "KeepFormat1", "ReplaceText1",
        "KeepFormat2", "KeepParaFormatInline", "ReplaceText2", "DoMarkupOutlook", "DoMarkupWord",
        "APIDebug", "ISearch_Approve", "ISearch", "Lib"
            }
            Return booleanSettings.Contains(settingKey)
        End Function


        Public Shared Function GetSettingValue(settingName As String, ByRef context As ISharedContext) As String
            ' Return the value of the setting based on its name
            Select Case settingName
                Case "APIKey"
                    Return context.INI_APIKeyBack
                Case "Temperature"
                    Return context.INI_Temperature
                Case "Timeout"
                    Return context.INI_Timeout.ToString() ' Convert Long to String
                Case "Model"
                    Return context.INI_Model
                Case "Endpoint"
                    Return context.INI_Endpoint
                Case "HeaderA"
                    Return context.INI_HeaderA
                Case "HeaderB"
                    Return context.INI_HeaderB
                Case "APICall"
                    Return context.INI_APICall
                Case "Response"
                    Return context.INI_Response
                Case "APIKey_2"
                    Return context.INI_APIKeyBack_2
                Case "Temperature_2"
                    Return context.INI_Temperature_2
                Case "Timeout_2"
                    Return context.INI_Timeout_2.ToString() ' Convert Long to String
                Case "Model_2"
                    Return context.INI_Model_2
                Case "Endpoint_2"
                    Return context.INI_Endpoint_2
                Case "HeaderA_2"
                    Return context.INI_HeaderA_2
                Case "HeaderB_2"
                    Return context.INI_HeaderB_2
                Case "APICall_2"
                    Return context.INI_APICall_2
                Case "Response_2"
                    Return context.INI_Response_2
                Case "OAuth2ClientMail"
                    Return context.INI_OAuth2ClientMail
                Case "OAuth2Scopes"
                    Return context.INI_OAuth2Scopes
                Case "OAuth2Endpoint"
                    Return context.INI_OAuth2Endpoint
                Case "OAuth2ATExpiry"
                    Return context.INI_OAuth2ATExpiry.ToString() ' Convert to String
                Case "OAuth2ClientMail_2"
                    Return context.INI_OAuth2ClientMail_2
                Case "OAuth2Scopes_2"
                    Return context.INI_OAuth2Scopes_2
                Case "OAuth2Endpoint_2"
                    Return context.INI_OAuth2Endpoint_2
                Case "OAuth2ATExpiry_2"
                    Return context.INI_OAuth2ATExpiry_2.ToString() ' Convert to String
                Case "Codebasis"
                    Return context.Codebasis
                Case "DoubleS"
                    Return context.INI_DoubleS.ToString()
                Case "KeepFormat1"
                    Return context.INI_KeepFormat1.ToString()
                Case "ReplaceText1"
                    Return context.INI_ReplaceText1.ToString()
                Case "KeepFormat2"
                    Return context.INI_KeepFormat2.ToString()
                Case "KeepFormatCap"
                    Return context.INI_KeepFormatCap.ToString()
                Case "KeepParaFormatInline"
                    Return context.INI_KeepParaFormatInline.ToString()
                Case "ReplaceText2"
                    Return context.INI_ReplaceText2.ToString()
                Case "DoMarkupOutlook"
                    Return context.INI_DoMarkupOutlook.ToString()
                Case "DoMarkupWord"
                    Return context.INI_DoMarkupWord.ToString()
                Case "MarkupMethodHelper"
                    Return context.INI_MarkupMethodHelper.ToString()
                Case "MarkupMethodWord"
                    Return context.INI_MarkupMethodWord.ToString()
                Case "MarkupMethodOutlook"
                    Return context.INI_MarkupMethodOutlook.ToString()
                Case "MarkupDiffCap"
                    Return context.INI_MarkupDiffCap.ToString()
                Case "MarkupRegexCap"
                    Return context.INI_MarkupRegexCap.ToString()
                Case "ChatCap"
                    Return context.INI_ChatCap.ToString()
                Case "PreCorrection"
                    Return context.INI_PreCorrection.ToString()
                Case "PostCorrection"
                    Return context.INI_PostCorrection.ToString()
                Case "Language1"
                    Return context.INI_Language1.ToString()
                Case "Language2"
                    Return context.INI_Language2.ToString()
                Case "ShortcutsWordExcel"
                    Return context.INI_ShortcutsWordExcel
                Case "PromptLibPath"
                    Return context.INI_PromptLibPath
                Case "PromptLibPath_Transcript"
                    Return context.INI_PromptLibPath_Transcript
                Case "SpeechModelPath"
                    Return context.INI_SpeechModelPath
                Case "APIDebug"
                    Return context.INI_APIDebug.ToString()
                Case "ISearch"
                    Return context.INI_ISearch.ToString()
                Case "ISearch_Approve"
                    Return context.INI_ISearch_Approve.ToString()
                Case "ISearch_URL"
                    Return context.INI_ISearch_URL
                Case "ISearch_ResponseURLStart"
                    Return context.INI_ISearch_ResponseURLStart
                Case "ISearch_ResponseMask1"
                    Return context.INI_ISearch_ResponseMask1
                Case "ISearch_ResponseMask2"
                    Return context.INI_ISearch_ResponseMask2
                Case "ISearch_Name"
                    Return context.INI_ISearch_Name
                Case "ISearch_Tries"
                    Return context.INI_ISearch_Tries.ToString() ' Convert Integer to String
                Case "ISearch_Results"
                    Return context.INI_ISearch_Results.ToString() ' Convert Integer to String
                Case "ISearch_MaxDepth"
                    Return context.INI_ISearch_MaxDepth.ToString() ' Convert Integer to String
                Case "ISearch_Timeout"
                    Return context.INI_ISearch_Timeout.ToString() ' Convert Long to String
                Case "ISearch_SearchTerm_SP"
                    Return context.INI_ISearch_SearchTerm_SP
                Case "ISearch_Apply_SP"
                    Return context.INI_ISearch_Apply_SP
                Case "ISearch_Apply_SP_Markup"
                    Return context.INI_ISearch_Apply_SP_Markup
                Case "Lib"
                    Return context.INI_Lib.ToString()
                Case "Lib_File"
                    Return context.INI_Lib_File
                Case "Lib_Timeout"
                    Return context.INI_Lib_Timeout.ToString() ' Convert Long to String
                Case "Lib_Find_SP"
                    Return context.INI_Lib_Find_SP
                Case "Lib_Apply_SP"
                    Return context.INI_Lib_Apply_SP
                Case "Lib_Apply_SP_Markup"
                    Return context.INI_Lib_Apply_SP_Markup
                Case Else
                    Return ""
            End Select
        End Function


        Public Shared Sub SetSettingValue(settingName As String, value As String, ByRef context As ISharedContext)
            ' Set the value of the setting based on its name

            Select Case Trim(settingName)
                Case "APIKey"
                    context.INI_APIKeyBack = value
                Case "APIKeyPrefix"
                    context.INI_APIKeyPrefix = value
                Case "Temperature"
                    context.INI_Temperature = value
                Case "Timeout"
                    context.INI_Timeout = Long.Parse(value) ' Parse String to Long
                Case "Model"
                    context.INI_Model = value
                Case "Endpoint"
                    context.INI_Endpoint = value
                Case "HeaderA"
                    context.INI_HeaderA = value
                Case "HeaderB"
                    context.INI_HeaderB = value
                Case "APICall"
                    context.INI_APICall = value
                Case "Response"
                    context.INI_Response = value
                Case "APIKey_2"
                    context.INI_APIKeyBack_2 = value
                Case "APIKeyPrefix_2"
                    context.INI_APIKeyPrefix_2 = value
                Case "Temperature_2"
                    context.INI_Temperature_2 = value
                Case "Timeout_2"
                    context.INI_Timeout_2 = Long.Parse(value) ' Parse String to Long
                Case "Model_2"
                    context.INI_Model_2 = value
                Case "Endpoint_2"
                    context.INI_Endpoint_2 = value
                Case "HeaderA_2"
                    context.INI_HeaderA_2 = value
                Case "HeaderB_2"
                    context.INI_HeaderB_2 = value
                Case "APICall_2"
                    context.INI_APICall_2 = value
                Case "Response_2"
                    context.INI_Response_2 = value
                Case "OAuth2ClientMail"
                    context.INI_OAuth2ClientMail = value
                Case "OAuth2Scopes"
                    context.INI_OAuth2Scopes = value
                Case "OAuth2Endpoint"
                    context.INI_OAuth2Endpoint = value
                Case "OAuth2ATExpiry"
                    context.INI_OAuth2ATExpiry = Long.Parse(value) ' Parse String to Long
                Case "OAuth2ClientMail_2"
                    context.INI_OAuth2ClientMail_2 = value
                Case "OAuth2Scopes_2"
                    context.INI_OAuth2Scopes_2 = value
                Case "OAuth2Endpoint_2"
                    context.INI_OAuth2Endpoint_2 = value
                Case "OAuth2ATExpiry_2"
                    context.INI_OAuth2ATExpiry_2 = Long.Parse(value)
                Case "Codebasis"
                    context.Codebasis = value
                Case "DoubleS"
                    context.INI_DoubleS = Boolean.Parse(value)
                Case "KeepFormat1"
                    context.INI_KeepFormat1 = Boolean.Parse(value)
                Case "ReplaceText1"
                    context.INI_ReplaceText1 = Boolean.Parse(value)
                Case "KeepFormat2"
                    context.INI_KeepFormat2 = Boolean.Parse(value)
                Case "KeepFormatCap"
                    context.INI_KeepFormatCap = Integer.Parse(value)
                Case "KeepParaFormatInline"
                    context.INI_KeepParaFormatInline = Boolean.Parse(value)
                Case "ReplaceText2"
                    context.INI_ReplaceText2 = Boolean.Parse(value)
                Case "DoMarkupOutlook"
                    context.INI_DoMarkupOutlook = Boolean.Parse(value)
                Case "DoMarkupWord"
                    context.INI_DoMarkupWord = Boolean.Parse(value)
                Case "MarkupMethodHelper"
                    context.INI_MarkupMethodHelper = Integer.Parse(value)
                Case "MarkupMethodWord"
                    context.INI_MarkupMethodWord = Integer.Parse(value)
                Case "MarkupMethodOutlook"
                    context.INI_MarkupMethodOutlook = Integer.Parse(value)
                Case "MarkupDiffCap"
                    context.INI_MarkupDiffCap = Integer.Parse(value)
                Case "MarkupRegexCap"
                    context.INI_MarkupRegexCap = Integer.Parse(value)
                Case "ChatCap"
                    context.INI_ChatCap = Integer.Parse(value)
                Case "PreCorrection"
                    context.INI_PreCorrection = value
                Case "PostCorrection"
                    context.INI_PostCorrection = value
                Case "Language1"
                    context.INI_Language1 = value
                Case "Language2"
                    context.INI_Language2 = value
                Case "ShortcutsWordExcel"
                    context.INI_ShortcutsWordExcel = value
                Case "PromptLibPath"
                    context.INI_PromptLibPath = value
                Case "PromptLibPath_Transcript"
                    context.INI_PromptLibPath_Transcript = value
                Case "SpeechModelPath"
                    context.INI_SpeechModelPath = value
                Case "APIDebug"
                    context.INI_APIDebug = Boolean.Parse(value)
                Case "ISearch"
                    context.INI_ISearch = Boolean.Parse(value)
                Case "ISearch_Approve"
                    context.INI_ISearch_Approve = Boolean.Parse(value)
                Case "ISearch_URL"
                    context.INI_ISearch_URL = value
                Case "ISearch_ResponseURLStart"
                    context.INI_ISearch_ResponseURLStart = value
                Case "ISearch_ResponseMask1"
                    context.INI_ISearch_ResponseMask1 = value
                Case "ISearch_ResponseMask2"
                    context.INI_ISearch_ResponseMask2 = value
                Case "ISearch_Name"
                    context.INI_ISearch_Name = value
                Case "ISearch_Tries"
                    context.INI_ISearch_Tries = Integer.Parse(value) ' Parse String to Integer
                Case "ISearch_Results"
                    context.INI_ISearch_Results = Integer.Parse(value) ' Parse String to Integer
                Case "ISearch_MaxDepth"
                    context.INI_ISearch_MaxDepth = Integer.Parse(value) ' Parse String to Integer
                Case "ISearch_Timeout"
                    context.INI_ISearch_Timeout = Long.Parse(value) ' Parse String to Long
                Case "ISearch_SearchTerm_SP"
                    context.INI_ISearch_SearchTerm_SP = value
                Case "ISearch_Apply_SP"
                    context.INI_ISearch_Apply_SP = value
                Case "ISearch_Apply_SP_Markup"
                    context.INI_ISearch_Apply_SP_Markup = value
                Case "Lib"
                    context.INI_Lib = Boolean.Parse(value)
                Case "Lib_File"
                    context.INI_Lib_File = value
                Case "Lib_Timeout"
                    context.INI_Lib_Timeout = Long.Parse(value) ' Parse String to Long
                Case "Lib_Find_SP"
                    context.INI_Lib_Find_SP = value
                Case "Lib_Apply_SP"
                    context.INI_Lib_Apply_SP = value
                Case "Lib_Apply_SP_Markup"
                    context.INI_Lib_Apply_SP_Markup = value

                Case Else
                    MessageBox.Show($"Error in SetSettingValue - could not save the value for '{settingName}'.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Select

            If context.INI_PromptLibPath = "" Then context.INI_PromptLib = False Else context.INI_PromptLib = True

        End Sub


        Public Shared Sub SwitchModels(ByRef context As ISharedContext)
            ' Switch the content of variables with a _2 suffix with their corresponding variables without the _2 suffix
            Dim temp As String
            Dim tempb As Boolean
            Dim templ As Long
            Dim tempi As Integer
            Dim tempt As DateTime

            temp = context.INI_Model
            context.INI_Model = context.INI_Model_2
            context.INI_Model_2 = temp

            temp = context.INI_APIKeyPrefix
            context.INI_APIKeyPrefix = context.INI_APIKeyPrefix_2
            context.INI_APIKeyPrefix_2 = temp

            temp = context.INI_APIKey
            context.INI_APIKey = context.INI_APIKey_2
            context.INI_APIKey_2 = temp

            tempb = context.INI_APIEncrypted
            context.INI_APIEncrypted = context.INI_APIEncrypted_2
            context.INI_APIEncrypted_2 = tempb

            temp = context.INI_Temperature
            context.INI_Temperature = context.INI_Temperature_2
            context.INI_Temperature_2 = temp

            templ = context.INI_Timeout
            context.INI_Timeout = context.INI_Timeout_2
            context.INI_Timeout_2 = templ

            tempi = context.INI_MaxOutputToken
            context.INI_MaxOutputToken = context.INI_MaxOutputToken_2
            context.INI_MaxOutputToken_2 = tempi

            temp = context.INI_Endpoint
            context.INI_Endpoint = context.INI_Endpoint_2
            context.INI_Endpoint_2 = temp

            temp = context.INI_HeaderA
            context.INI_HeaderA = context.INI_HeaderA_2
            context.INI_HeaderA_2 = temp

            temp = context.INI_HeaderB
            context.INI_HeaderB = context.INI_HeaderB_2
            context.INI_HeaderB_2 = temp

            temp = context.INI_Response
            context.INI_Response = context.INI_Response_2
            context.INI_Response_2 = temp

            temp = context.INI_APICall
            context.INI_APICall = context.INI_APICall_2
            context.INI_APICall_2 = temp

            temp = context.INI_OAuth2ClientMail
            context.INI_OAuth2ClientMail = context.INI_OAuth2ClientMail_2
            context.INI_OAuth2ClientMail_2 = temp

            temp = context.INI_OAuth2Scopes
            context.INI_OAuth2Scopes = context.INI_OAuth2Scopes_2
            context.INI_OAuth2Scopes_2 = temp

            temp = context.INI_OAuth2Endpoint
            context.INI_OAuth2Endpoint = context.INI_OAuth2Endpoint_2
            context.INI_OAuth2Endpoint_2 = temp

            templ = context.INI_OAuth2ATExpiry
            context.INI_OAuth2ATExpiry = context.INI_OAuth2ATExpiry_2
            context.INI_OAuth2ATExpiry_2 = templ

            temp = context.DecodedAPI
            context.DecodedAPI = context.DecodedAPI_2
            context.DecodedAPI_2 = temp

            temp = context.INI_APIKeyBack
            context.INI_APIKeyBack = context.INI_APIKeyBack_2
            context.INI_APIKeyBack_2 = temp

            tempt = context.TokenExpiry
            context.TokenExpiry = context.TokenExpiry_2
            context.TokenExpiry_2 = tempt

            tempb = context.INI_OAuth2
            context.INI_OAuth2 = context.INI_OAuth2_2
            context.INI_OAuth2_2 = tempb



        End Sub

        Public Shared Sub UpdateAppConfig(ByRef context As ISharedContext)
            Try


                Dim IniFilePath As String = ""
                Dim RegFilePath As String = ""
                Dim DefaultPath As String = ""
                Dim DefaultPath2 As String = ""
                Dim TempIniFilePath As String = ""

                ' Determine the configuration file path

                RegFilePath = GetFromRegistry(RegPath_Base, RegPath_IniPath, True)
                DefaultPath = GetDefaultINIPath(context.RDV)
                DefaultPath2 = GetDefaultINIPath("Word")

                If Not String.IsNullOrWhiteSpace(RegFilePath) And RegPath_IniPrio Then
                    IniFilePath = System.IO.Path.Combine(ExpandEnvironmentVariables(RegFilePath), $"{AN2}.ini")
                ElseIf System.IO.File.Exists(DefaultPath) Then
                    IniFilePath = DefaultPath
                ElseIf System.IO.File.Exists(DefaultPath2) Then
                    IniFilePath = DefaultPath2
                ElseIf Not String.IsNullOrWhiteSpace(RegFilePath) Then
                    IniFilePath = System.IO.Path.Combine(ExpandEnvironmentVariables(RegFilePath), $"{AN2}.ini")
                Else
                    IniFilePath = DefaultPath
                End If

                IniFilePath = RemoveCR(IniFilePath)

                ' Validate IniFilePath
                If Not System.IO.File.Exists(IniFilePath) Then
                    ShowCustomMessageBox($"The configuration file '{IniFilePath}' was not found.")
                    Exit Sub
                End If

                ' Create a temporary file for the updated configuration
                TempIniFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(IniFilePath), $"{AN2}_temp.ini")

                ' Define all expected keys and their default or in-memory values
                Dim expectedKeys As New Dictionary(Of String, String) From {
                    {"APIKey", context.INI_APIKeyBack},
                    {"Endpoint", context.INI_Endpoint},
                    {"HeaderA", context.INI_HeaderA},
                    {"HeaderB", context.INI_HeaderB},
                    {"Response", context.INI_Response},
                    {"APICall", context.INI_APICall},
                    {"Timeout", context.INI_Timeout.ToString()},
                    {"MaxOutputToken", context.INI_MaxOutputToken.ToString()},
                    {"Temperature", context.INI_Temperature},
                    {"Model", context.INI_Model},
                    {"PreCorrection", context.INI_PreCorrection},
                    {"PostCorrection", context.INI_PostCorrection},
                    {"APIKeyPrefix", context.INI_APIKeyPrefix},
                    {"UsageRestrictions", context.INI_UsageRestrictions},
                    {"Language1", context.INI_Language1},
                    {"Language2", context.INI_Language2},
                    {"DoubleS", context.INI_DoubleS.ToString()},
                    {"KeepFormat1", context.INI_KeepFormat1.ToString()},
                    {"ReplaceText1", context.INI_ReplaceText1.ToString()},
                    {"KeepFormat2", context.INI_KeepFormat2.ToString()},
                    {"KeepFormatCap", context.INI_KeepFormatCap.ToString()},
                    {"KeepParaFormatInline", context.INI_KeepParaFormatInline.ToString()},
                    {"ReplaceText2", context.INI_ReplaceText2.ToString()},
                    {"DoMarkupOutlook", context.INI_DoMarkupOutlook.ToString()},
                    {"DoMarkupWord", context.INI_DoMarkupWord.ToString()},
                    {"MarkupMethodOutlook", context.INI_MarkupMethodOutlook.ToString()},
                    {"MarkupDiffCap", context.INI_MarkupDiffCap.ToString()},
                    {"MarkupRegexCap", context.INI_MarkupRegexCap.ToString()},
                    {"ChatCap", context.INI_ChatCap.ToString()},
                    {"APIDebug", context.INI_APIDebug.ToString()},
                    {"APIKeyEncrypted", context.INI_APIEncrypted.ToString()},
                    {"SecondAPI", context.INI_SecondAPI.ToString()},
                    {"APIKey_2", context.INI_APIKeyBack_2},
                    {"Endpoint_2", context.INI_Endpoint_2},
                    {"HeaderA_2", context.INI_HeaderA_2},
                    {"HeaderB_2", context.INI_HeaderB_2},
                    {"Response_2", context.INI_Response_2},
                    {"APICall_2", context.INI_APICall_2},
                    {"Timeout_2", context.INI_Timeout_2.ToString()},
                    {"MaxOutputToken_2", context.INI_MaxOutputToken_2.ToString()},
                    {"Temperature_2", context.INI_Temperature_2},
                    {"Model_2", context.INI_Model_2},
                    {"APIKeyEncrypted_2", context.INI_APIEncrypted_2.ToString()},
                    {"APIKeyPrefix_2", context.INI_APIKeyPrefix_2},
                    {"OAuth2", context.INI_OAuth2.ToString()},
                    {"OAuth2ClientMail", context.INI_OAuth2ClientMail},
                    {"OAuth2Scopes", context.INI_OAuth2Scopes},
                    {"OAuth2Endpoint", context.INI_OAuth2Endpoint},
                    {"OAuth2ATExpiry", context.INI_OAuth2ATExpiry.ToString()},
                    {"OAuth2_2", context.INI_OAuth2_2.ToString()},
                    {"OAuth2ClientMail_2", context.INI_OAuth2ClientMail_2},
                    {"OAuth2Scopes_2", context.INI_OAuth2Scopes_2},
                    {"OAuth2Endpoint_2", context.INI_OAuth2Endpoint_2},
                    {"OAuth2ATExpiry_2", context.INI_OAuth2ATExpiry_2.ToString()},
                    {"ISearch", context.INI_ISearch.ToString()},
                    {"ISearch_Approve", context.INI_ISearch_Approve.ToString()},
                    {"ISearch_URL", context.INI_ISearch_URL},
                    {"ISearch_ResponseMask1", context.INI_ISearch_ResponseMask1},
                    {"ISearch_ResponseMask2", context.INI_ISearch_ResponseMask2},
                    {"ISearch_Name", context.INI_ISearch_Name},
                    {"ISearch_Tries", context.INI_ISearch_Tries.ToString()},
                    {"ISearch_Results", context.INI_ISearch_Results.ToString()},
                    {"ISearch_MaxDepth", context.INI_ISearch_MaxDepth.ToString()},
                    {"ISearch_Timeout", context.INI_ISearch_Timeout.ToString()},
                    {"ISearch_SearchTerm_SP", context.INI_ISearch_SearchTerm_SP},
                    {"ISearch_Apply_SP", context.INI_ISearch_Apply_SP},
                    {"ISearch_Apply_SP_Markup", context.INI_ISearch_Apply_SP_Markup},
                    {"Lib", context.INI_Lib.ToString()},
                    {"Lib_File", context.INI_Lib_File},
                    {"Lib_Timeout", context.INI_Lib_Timeout.ToString()},
                    {"Lib_Find_SP", context.INI_Lib_Find_SP},
                    {"Lib_Apply_SP", context.INI_Lib_Apply_SP},
                    {"Lib_Apply_SP_Markup", context.INI_ISearch_Apply_SP_Markup},
                    {"MarkupMethodHelper", context.INI_MarkupMethodHelper.ToString()},
                    {"MarkupMethodWord", context.INI_MarkupMethodWord.ToString()},
                    {"ShortcutsWordExcel", context.INI_ShortcutsWordExcel},
                    {"ContextMenu", context.INI_ContextMenu},
                    {"UpdateCheckInterval", context.INI_UpdateCheckInterval},
                    {"UpdatePath", context.INI_UpdatePath},
                    {"SpeechModelPath", context.INI_SpeechModelPath},
                    {"TTSEndpoint", context.INI_TTSEndpoint},
                    {"PromptLib", context.INI_PromptLibPath},
                    {"PromptLib_Transcript", context.INI_PromptLibPath_Transcript},
                    {"SP_Translate", context.SP_Translate},
                    {"SP_Correct", context.SP_Correct},
                    {"SP_Improve", context.SP_Improve},
                    {"SP_Explain", context.SP_Explain},
                    {"SP_SuggestTitles", context.SP_SuggestTitles},
                    {"SP_Friendly", context.SP_Friendly},
                    {"SP_Convincing", context.SP_Convincing},
                    {"SP_NoFillers", context.SP_NoFillers},
                    {"SP_Podcast", context.SP_Podcast},
                    {"SP_Shorten", context.SP_Shorten},
                    {"SP_Summarize", context.SP_Summarize},
                    {"SP_MailReply", context.SP_MailReply},
                    {"SP_MailSumup", context.SP_MailSumup},
                    {"SP_FreestyleText", context.SP_FreestyleText},
                    {"SP_FreestyleNoText", context.SP_FreestyleNoText},
                    {"SP_SwitchParty", context.SP_SwitchParty},
                    {"SP_Anonymize", context.SP_Anonymize},
                    {"SP_ContextSearch", context.SP_ContextSearch},
                    {"SP_ContextSearchMulti", context.SP_ContextSearchMulti},
                    {"SP_RangeOfCells", context.SP_RangeOfCells},
                    {"SP_WriteNeatly", context.SP_WriteNeatly},
                    {"SP_Add_KeepFormulasIntact", context.SP_Add_KeepFormulasIntact},
                    {"SP_Add_KeepHTMLIntact", context.SP_Add_KeepHTMLIntact},
                    {"SP_Add_KeepInlineIntact", context.SP_Add_KeepInlineIntact},
                    {"SP_Add_Bubbles", context.SP_Add_Bubbles},
                    {"SP_Add_Revisions", context.SP_Add_Revisions},
                    {"SP_MarkupRegex", context.SP_MarkupRegex},
                    {"SP_ChatWord", context.SP_ChatWord},
                    {"SP_Add_ChatWord_Commands", context.SP_Add_ChatWord_Commands}
                }

                Dim KeysToSkipWhenDefault As New Dictionary(Of String, String) From {
                    {"ISearch_SearchTerm_SP", Default_INI_ISearch_SearchTerm_SP},
                    {"ISearch_Apply_SP", Default_INI_ISearch_Apply_SP},
                    {"ISearch_Apply_SP_Markup", Default_INI_ISearch_Apply_SP_Markup},
                    {"SP_Translate", Default_SP_Translate},
                    {"SP_Correct", Default_SP_Correct},
                    {"SP_Improve", Default_SP_Improve},
                    {"SP_Explain", Default_SP_Explain},
                    {"SP_SuggestTitles", Default_SP_SuggestTitles},
                    {"SP_Friendly", Default_SP_Friendly},
                    {"SP_Convincing", Default_SP_Convincing},
                    {"SP_NoFillers", Default_SP_NoFillers},
                    {"SP_Podcast", Default_SP_Podcast},
                    {"SP_Shorten", Default_SP_Shorten},
                    {"SP_Summarize", Default_SP_Summarize},
                    {"SP_MailReply", Default_SP_MailReply},
                    {"SP_MailSumup", Default_SP_MailSumup},
                    {"SP_FreestyleText", Default_SP_FreestyleText},
                    {"SP_FreestyleNoText", Default_SP_FreestyleNoText},
                    {"SP_SwitchParty", Default_SP_SwitchParty},
                    {"SP_Anonymize", Default_SP_Anonymize},
                    {"SP_ContextSearch", Default_SP_ContextSearch},
                    {"SP_ContextSearchMultiple", Default_SP_ContextSearchMulti},
                    {"SP_RangeOfCells", Default_SP_RangeOfCells},
                    {"SP_WriteNeatly", Default_SP_WriteNeatly},
                    {"SP_Add_KeepFormulasIntact", Default_SP_Add_KeepFormulasIntact},
                    {"SP_Add_KeepHTMLIntact", Default_SP_Add_KeepHTMLIntact},
                    {"SP_Add_KeepInlineIntact", Default_SP_Add_KeepInlineIntact},
                    {"SP_Add_Bubbles", Default_SP_Add_Bubbles},
                    {"SP_Add_Revisions", Default_SP_Add_Revisions},
                    {"SP_MarkupRegex", Default_SP_MarkupRegex},
                    {"SP_ChatWord", Default_SP_ChatWord},
                    {"SP_Add_ChatWord_Commands", Default_SP_Add_ChatWord_Commands}
                }


                ' Read the original ini file content
                Dim originalContent As String = System.IO.File.ReadAllText(IniFilePath)
                Dim updatedContent As New StringBuilder()
                Dim foundKeys As New HashSet(Of String)()

                ' Split into lines and process each line
                Dim iniLines As String() = originalContent.Split({vbCrLf}, StringSplitOptions.None)
                For Each line As String In iniLines
                    Dim trimmedLine As String = line.Trim()

                    ' Preserve comments and empty lines
                    If String.IsNullOrEmpty(trimmedLine) OrElse trimmedLine.StartsWith(";") Then
                        updatedContent.AppendLine(line)
                        Continue For
                    End If

                    ' Process key-value pairs
                    Dim keyValue As String() = trimmedLine.Split(New Char() {"="c}, 2)
                    If keyValue.Length = 2 Then
                        Dim key As String = keyValue(0).Trim()
                        Dim value As String = keyValue(1).Trim()

                        ' Update values for known keys
                        If expectedKeys.ContainsKey(key) Then
                            value = expectedKeys(key)
                            foundKeys.Add(key)
                        End If

                        ' Write the updated key-value pair
                        updatedContent.AppendLine($"{key} = {value}")
                    Else
                        ' Preserve lines that are not key-value pairs
                        updatedContent.AppendLine(line)
                    End If
                Next

                ' Add missing keys to the updated content, but now skip certain keys if their value matches the default
                For Each key In expectedKeys.Keys.Except(foundKeys)
                    Dim value As String = expectedKeys(key)
                    ' Check if the key exists in KeysToSkipWhenDefault and if the value matches the default
                    If KeysToSkipWhenDefault.ContainsKey(key) AndAlso KeysToSkipWhenDefault(key) = value Then
                        ' Skip adding this key as its value matches the default
                        Continue For
                    End If

                    ' Write the key-value pair to the updated content
                    updatedContent.AppendLine($"{key} = {value}")
                Next

                ' Write the updated content to the temporary ini file
                System.IO.File.WriteAllText(TempIniFilePath, updatedContent.ToString())

                ' Replace the original file with the updated file
                System.IO.File.Delete(DefaultPath)
                System.IO.File.Move(TempIniFilePath, DefaultPath)

                context.INIloaded = False

                If IniFilePath = DefaultPath Then
                    ShowCustomMessageBox("Your configuration file has been updated.")
                Else
                    ShowCustomMessageBox("Your configuration has been saved To a local configuration file (which will be used going forward until deleted).")
                End If

                InitializeConfig(context, False, True)

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error updating configuration file: {ex.Message}")
            End Try
        End Sub


        Public Shared Sub ResetLocalAppConfig(ByRef context As ISharedContext)
            Try
                ' Determine the path to the existing .ini file
                Dim IniFilePath As String = System.IO.Path.Combine(GetDefaultINIPath(context.RDV))
                Dim TempIniFilePath As String = ""

                ' Validate IniFilePath
                If Not System.IO.File.Exists(IniFilePath) Then
                    ShowCustomMessageBox($"The configuration file '{IniFilePath}' was not found.")
                    Exit Sub
                End If

                ' Create a temporary file for the updated configuration
                TempIniFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(IniFilePath), $"{AN2}_temp.ini")

                ' Define all expected keys and their default or in-memory values
                Dim expectedKeys As New Dictionary(Of String, String) From {
                    {"APIKey", context.INI_APIKeyBack},
                    {"Endpoint", context.INI_Endpoint},
                    {"HeaderA", context.INI_HeaderA},
                    {"HeaderB", context.INI_HeaderB},
                    {"Response", context.INI_Response},
                    {"APICall", context.INI_APICall},
                    {"Timeout", context.INI_Timeout.ToString()},
                    {"MaxOutputToken", context.INI_MaxOutputToken.ToString()},
                    {"Temperature", context.INI_Temperature},
                    {"Model", context.INI_Model},
                    {"APIKeyPrefix", context.INI_APIKeyPrefix},
                    {"APIKeyEncrypted", context.INI_APIEncrypted.ToString()},
                    {"SecondAPI", context.INI_SecondAPI.ToString()},
                    {"APIKey_2", context.INI_APIKeyBack_2},
                    {"Endpoint_2", context.INI_Endpoint_2},
                    {"HeaderA_2", context.INI_HeaderA_2},
                    {"HeaderB_2", context.INI_HeaderB_2},
                    {"Response_2", context.INI_Response_2},
                    {"APICall_2", context.INI_APICall_2},
                    {"Timeout_2", context.INI_Timeout_2.ToString()},
                    {"MaxOutputToken_2", context.INI_MaxOutputToken_2.ToString()},
                    {"Temperature_2", context.INI_Temperature_2},
                    {"Model_2", context.INI_Model_2},
                    {"APIKeyEncrypted_2", context.INI_APIEncrypted_2.ToString()},
                    {"APIKeyPrefix_2", context.INI_APIKeyPrefix_2},
                    {"OAuth2", context.INI_OAuth2.ToString()},
                    {"OAuth2ClientMail", context.INI_OAuth2ClientMail},
                    {"OAuth2Scopes", context.INI_OAuth2Scopes},
                    {"OAuth2Endpoint", context.INI_OAuth2Endpoint},
                    {"OAuth2ATExpiry", context.INI_OAuth2ATExpiry.ToString()},
                    {"OAuth2_2", context.INI_OAuth2_2.ToString()},
                    {"OAuth2ClientMail_2", context.INI_OAuth2ClientMail_2},
                    {"OAuth2Scopes_2", context.INI_OAuth2Scopes_2},
                    {"OAuth2Endpoint_2", context.INI_OAuth2Endpoint_2},
                    {"OAuth2ATExpiry_2", context.INI_OAuth2ATExpiry_2.ToString()},
                    {"SpeechModelPath", context.INI_SpeechModelPath},
                    {"TTSEndpoint", context.INI_TTSEndpoint},
                    {"PromptLib", context.INI_PromptLibPath},
                    {"PromptLib_Transcript", context.INI_PromptLibPath_Transcript}
                }

                ' Read the original ini file content
                Dim originalContent As String = System.IO.File.ReadAllText(IniFilePath)
                Dim updatedContent As New StringBuilder()
                Dim foundKeys As New HashSet(Of String)()

                ' Split into lines and process each line
                Dim iniLines As String() = originalContent.Split({vbCrLf}, StringSplitOptions.None)
                For Each line As String In iniLines
                    Dim trimmedLine As String = line.Trim()

                    ' Preserve comments and empty lines
                    If String.IsNullOrEmpty(trimmedLine) OrElse trimmedLine.StartsWith(";") Then
                        updatedContent.AppendLine(line)
                        Continue For
                    End If

                    ' Process key-value pairs
                    Dim keyValue As String() = trimmedLine.Split(New Char() {"="c}, 2)
                    If keyValue.Length = 2 Then
                        Dim key As String = keyValue(0).Trim()
                        Dim value As String = keyValue(1).Trim()

                        ' Retain keys that are in the expectedKeys dictionary
                        If expectedKeys.ContainsKey(key) Then
                            value = expectedKeys(key)
                            foundKeys.Add(key)
                            updatedContent.AppendLine($"{key} = {value}")
                        End If
                    End If
                Next

                ' Add missing keys to the updated content
                For Each key In expectedKeys.Keys.Except(foundKeys)
                    updatedContent.AppendLine($"{key} = {expectedKeys(key)}")
                Next

                ' Write the updated content to the temporary ini file
                System.IO.File.WriteAllText(TempIniFilePath, updatedContent.ToString())

                ' Replace the original file with the updated file
                System.IO.File.Delete(IniFilePath)
                System.IO.File.Move(TempIniFilePath, IniFilePath)

                context.INIloaded = False

                ShowCustomMessageBox("Configuration file has been updated.")

                InitializeConfig(context, False, True)

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error resetting configuration file: {ex.Message}")

            End Try
        End Sub


        Public Shared Function ShowVariableConfigurationWindow(
                                        variableNames As List(Of String),
                                        variableValues As Dictionary(Of String, Object),
                                        Optional ownerForm As Form = Nothing
                                    ) As Dictionary(Of String, Object)

            Dim baseWidth As Integer = 400
            Dim baseHeight As Integer = 300

            Dim configForm As New Form() With {
                        .Text = "Configure Parameters",
                        .FormBorderStyle = FormBorderStyle.Sizable,
                        .StartPosition = FormStartPosition.CenterScreen,
                        .MaximizeBox = True,
                        .MinimizeBox = True,
                        .Font = New System.Drawing.Font("Segoe UI", 9.0F),
                        .ClientSize = New Size(baseWidth * 2, baseHeight * 2) ' Double the "base" size
                            }

            ' Safely set the icon (if resource is valid)
            Try
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                configForm.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch ex As Exception
                MessageBox.Show($"Error loading icon: {ex.Message}",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
            End Try

            ' Create the DataGridView and allow it to fill space / resize with the form
            Dim dgv As New DataGridView() With {
                            .Dock = DockStyle.Fill,
                            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                            .AllowUserToResizeColumns = True,
                            .AllowUserToAddRows = False,
                            .AllowUserToDeleteRows = False,
                            .ReadOnly = False
                        }

            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
            dgv.ColumnHeadersHeight += 5

            ' Increase the header font size by 5 points over the form's base font
            dgv.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font(configForm.Font.FontFamily,
                                                      configForm.Font.Size,
                                                      configForm.Font.Style)

            ' Define columns
            Dim colVariableName As New DataGridViewTextBoxColumn() With {
                                .HeaderText = "Variable Name",
                                .ReadOnly = True
                            }
            Dim colValue As New DataGridViewTextBoxColumn() With {
                                .HeaderText = "Value"
                            }
            Dim colType As New DataGridViewTextBoxColumn() With {
                        .HeaderText = "Type",
                        .Visible = False,
                        .ReadOnly = True
                    }
            dgv.Columns.AddRange({colVariableName, colValue, colType})

            ' Populate the DataGridView
            For Each variableName As String In variableNames
                Dim valueObj As Object = If(variableValues.ContainsKey(variableName), variableValues(variableName), Nothing)
                Dim displayValue As String = If(valueObj?.ToString(), "")
                Dim typeName As String = If(valueObj?.GetType().ToString(), "String")
                dgv.Rows.Add(variableName, displayValue, typeName)
            Next

            ' Create Save / Cancel buttons
            Dim saveButton As New Button() With {.Text = "Save", .AutoSize = True}
            Dim cancelButton As New Button() With {.Text = "Cancel", .AutoSize = True}

            ' FlowLayoutPanel for buttons
            Dim buttonPanel As New FlowLayoutPanel() With {
                        .Dock = DockStyle.Bottom,
                        .FlowDirection = FlowDirection.RightToLeft,
                        .AutoSize = True
                    }
            buttonPanel.Controls.Add(cancelButton)
            buttonPanel.Controls.Add(saveButton)

            ' Main TableLayoutPanel
            Dim mainPanel As New TableLayoutPanel() With {
                        .Dock = DockStyle.Fill,
                        .RowCount = 2,
                        .ColumnCount = 1
                    }
            ' The DataGridView row fills all remaining space
            mainPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            ' The button panel auto-sizes
            mainPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            mainPanel.Controls.Add(dgv, 0, 0)
            mainPanel.Controls.Add(buttonPanel, 0, 1)

            ' Add to the form
            configForm.Controls.Add(mainPanel)

            ' Event handlers
            Dim result As DialogResult = DialogResult.Cancel

            AddHandler saveButton.Click,
                        Sub(sender, e)
                            Try
                                For Each row As DataGridViewRow In dgv.Rows
                                    If Not row.IsNewRow Then
                                        Dim varName As String = CStr(row.Cells(0).Value)
                                        Dim varValue As String = CStr(row.Cells(1).Value)
                                        If Not String.IsNullOrEmpty(varName) Then
                                            variableValues(varName) = varValue
                                        End If
                                    End If
                                Next
                                result = DialogResult.OK
                                configForm.Close()
                            Catch ex As Exception
                                MessageBox.Show($"Error saving values: {ex.Message}",
                                                "Error",
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error)
                            End Try
                        End Sub

            AddHandler cancelButton.Click,
                        Sub(sender, e)
                            result = DialogResult.Cancel
                            configForm.Close()
                        End Sub

            ' Show the form (modal) in front of the owner form if available
            If ownerForm IsNot Nothing Then
                configForm.ShowDialog(ownerForm)
            Else
                configForm.ShowDialog()
            End If

            ' Return updated dictionary or original
            Return If(result = DialogResult.OK,
              variableValues,
              Nothing)

            'Return If(result = DialogResult.OK,
            'variableValues,
            ' New Dictionary(Of String, Object)(variableValues))

        End Function



        Public Shared Sub ShowExpertConfiguration(ByRef context As ISharedContext, ownerform As Form)
            ' Dictionary to store variable names and their current values
            Dim variableValues As New Dictionary(Of String, Object)

            ' Populate the dictionary with all the required variables
            variableValues.Add("APIKey", context.INI_APIKeyBack) ' Use Context.INI_APIKeyBack, display as Context.INI_APIKey
            variableValues.Add("Temperature", context.INI_Temperature)
            variableValues.Add("Timeout", context.INI_Timeout)
            variableValues.Add("MaxOutputToken", context.INI_MaxOutputToken)
            variableValues.Add("Model", context.INI_Model)
            variableValues.Add("Endpoint", context.INI_Endpoint)
            variableValues.Add("HeaderA", context.INI_HeaderA)
            variableValues.Add("HeaderB", context.INI_HeaderB)
            variableValues.Add("APICall", context.INI_APICall)
            variableValues.Add("Response", context.INI_Response)
            variableValues.Add("DoubleS", context.INI_DoubleS)
            variableValues.Add("PreCorrection", context.INI_PreCorrection)
            variableValues.Add("PostCorrection", context.INI_PostCorrection)
            variableValues.Add("APIEncrypted", context.INI_APIEncrypted)
            variableValues.Add("APIKeyPrefix", context.INI_APIKeyPrefix)
            variableValues.Add("MarkupMethodOutlook", context.INI_MarkupMethodOutlook)
            variableValues.Add("MarkupDiffCap", context.INI_MarkupDiffCap)
            variableValues.Add("MarkupRegexCap", context.INI_MarkupRegexCap)
            variableValues.Add("ChatCap", context.INI_ChatCap)
            variableValues.Add("OAuth2", context.INI_OAuth2)
            variableValues.Add("OAuth2ClientMail", context.INI_OAuth2ClientMail)
            variableValues.Add("OAuth2Scopes", context.INI_OAuth2Scopes)
            variableValues.Add("OAuth2Endpoint", context.INI_OAuth2Endpoint)
            variableValues.Add("OAuth2ATExpiry", context.INI_OAuth2ATExpiry)
            variableValues.Add("SecondAPI", context.INI_SecondAPI)
            variableValues.Add("APIKey_2", context.INI_APIKeyBack_2) ' Use Context.INI_APIKeyBack_2, display as Context.INI_APIKey_2
            variableValues.Add("Temperature_2", context.INI_Temperature_2)
            variableValues.Add("Timeout_2", context.INI_Timeout_2)
            variableValues.Add("MaxOutputToken_2", context.INI_MaxOutputToken_2)
            variableValues.Add("Model_2", context.INI_Model_2)
            variableValues.Add("Endpoint_2", context.INI_Endpoint_2)
            variableValues.Add("HeaderA_2", context.INI_HeaderA_2)
            variableValues.Add("HeaderB_2", context.INI_HeaderB_2)
            variableValues.Add("APICall_2", context.INI_APICall_2)
            variableValues.Add("Response_2", context.INI_Response_2)
            variableValues.Add("APIEncrypted_2", context.INI_APIEncrypted_2)
            variableValues.Add("APIKeyPrefix_2", context.INI_APIKeyPrefix_2)
            variableValues.Add("OAuth2_2", context.INI_OAuth2_2)
            variableValues.Add("OAuth2ClientMail_2", context.INI_OAuth2ClientMail_2)
            variableValues.Add("OAuth2Scopes_2", context.INI_OAuth2Scopes_2)
            variableValues.Add("OAuth2Endpoint_2", context.INI_OAuth2Endpoint_2)
            variableValues.Add("OAuth2ATExpiry_2", context.INI_OAuth2ATExpiry_2)
            variableValues.Add("APIDebug", context.INI_APIDebug)
            variableValues.Add("UsageRestrictions", context.INI_UsageRestrictions)
            variableValues.Add("Language1", context.INI_Language1)
            variableValues.Add("Language2", context.INI_Language2)
            variableValues.Add("KeepFormat1", context.INI_KeepFormat1)
            variableValues.Add("KeepFormat2", context.INI_KeepFormat2)
            variableValues.Add("KeepFormatCap", context.INI_KeepFormatCap)
            variableValues.Add("KeepParaFormatInline", context.INI_KeepParaFormatInline)
            variableValues.Add("ReplaceText1", context.INI_ReplaceText1)
            variableValues.Add("ReplaceText2", context.INI_ReplaceText2)
            variableValues.Add("DoMarkupOutlook", context.INI_DoMarkupOutlook)
            variableValues.Add("DoMarkupWord", context.INI_DoMarkupWord)
            variableValues.Add("ISearch", context.INI_ISearch)
            variableValues.Add("ISearch_Approve", context.INI_ISearch_Approve)
            variableValues.Add("ISearch_URL", context.INI_ISearch_URL)
            variableValues.Add("ISearch_ResponseMask1", context.INI_ISearch_ResponseMask1)
            variableValues.Add("ISearch_ResponseMask2", context.INI_ISearch_ResponseMask2)
            variableValues.Add("ISearch_Name", context.INI_ISearch_Name)
            variableValues.Add("ISearch_Tries", context.INI_ISearch_Tries)
            variableValues.Add("ISearch_Results", context.INI_ISearch_Results)
            variableValues.Add("ISearch_MaxDepth", context.INI_ISearch_MaxDepth)
            variableValues.Add("ISearch_Timeout", context.INI_ISearch_Timeout)
            variableValues.Add("ISearch_SearchTerm_SP", context.INI_ISearch_SearchTerm_SP)
            variableValues.Add("ISearch_Apply_SP", context.INI_ISearch_Apply_SP)
            variableValues.Add("ISearch_Apply_SP_Markup", context.INI_ISearch_Apply_SP_Markup)
            variableValues.Add("Lib", context.INI_Lib)
            variableValues.Add("Lib_File", context.INI_Lib_File)
            variableValues.Add("Lib_Timeout", context.INI_Lib_Timeout)
            variableValues.Add("Lib_Find_SP", context.INI_Lib_Find_SP)
            variableValues.Add("Lib_Apply_SP", context.INI_Lib_Apply_SP)
            variableValues.Add("Lib_Apply_SP_Markup", context.INI_Lib_Apply_SP_Markup)
            variableValues.Add("MarkupMethodHelper", context.INI_MarkupMethodHelper)
            variableValues.Add("MarkupMethodWord", context.INI_MarkupMethodWord)
            variableValues.Add("ContextMenu", context.INI_ContextMenu)
            variableValues.Add("UpdateCheckInterval", context.INI_UpdateCheckInterval)
            variableValues.Add("UpdatePath", context.INI_UpdatePath)
            variableValues.Add("SpeechModelPath", context.INI_SpeechModelPath)
            variableValues.Add("TTSEndpoint", context.INI_TTSEndpoint)
            variableValues.Add("ShortcutsWordExcel", context.INI_ShortcutsWordExcel)
            variableValues.Add("PromptLib", context.INI_PromptLibPath)
            variableValues.Add("PromptLib_Transcript", context.INI_PromptLibPath_Transcript)
            variableValues.Add("SP_Translate", context.SP_Translate)
            variableValues.Add("SP_Correct", context.SP_Correct)
            variableValues.Add("SP_Improve", context.SP_Improve)
            variableValues.Add("SP_Explain", context.SP_Explain)
            variableValues.Add("SP_SuggestTitles", context.SP_SuggestTitles)
            variableValues.Add("SP_Friendly", context.SP_Friendly)
            variableValues.Add("SP_Convincing", context.SP_Convincing)
            variableValues.Add("SP_NoFillers", context.SP_NoFillers)
            variableValues.Add("SP_Podcast", context.SP_Podcast)
            variableValues.Add("SP_Shorten", context.SP_Shorten)
            variableValues.Add("SP_Summarize", context.SP_Summarize)
            variableValues.Add("SP_MailReply", context.SP_MailReply)
            variableValues.Add("SP_MailSumup", context.SP_MailSumup)
            variableValues.Add("SP_FreestyleText", context.SP_FreestyleText)
            variableValues.Add("SP_FreestyleNoText", context.SP_FreestyleNoText)
            variableValues.Add("SP_SwitchParty", context.SP_SwitchParty)
            variableValues.Add("SP_Anonymize", context.SP_Anonymize)
            variableValues.Add("SP_ContextSearch", context.SP_ContextSearch)
            variableValues.Add("SP_ContextSearchMulti", context.SP_ContextSearchMulti)
            variableValues.Add("SP_RangeOfCells", context.SP_RangeOfCells)
            variableValues.Add("SP_WriteNeatly", context.SP_WriteNeatly)
            variableValues.Add("SP_Add_KeepFormulasIntact", context.SP_Add_KeepFormulasIntact)
            variableValues.Add("SP_Add_KeepHTMLIntact", context.SP_Add_KeepHTMLIntact)
            variableValues.Add("SP_Add_KeepInlineIntact", context.SP_Add_KeepInlineIntact)
            variableValues.Add("SP_Add_Bubbles", context.SP_Add_Bubbles)
            variableValues.Add("SP_Add_Revisions", context.SP_Add_Revisions)
            variableValues.Add("SP_MarkupRegex", context.SP_MarkupRegex)
            variableValues.Add("SP_ChatWord", context.SP_ChatWord)
            variableValues.Add("SP_Add_ChatWord_Commands", context.SP_Add_ChatWord_Commands)

            ' Extract variable names from the dictionary
            Dim variableNames As New List(Of String)(variableValues.Keys)

            ' Call the ShowVariableConfigurationWindow function and get the updated values
            Dim updatedValues = ShowVariableConfigurationWindow(variableNames, variableValues, ownerform)

            If Not IsNothing(updatedValues) Then

                ' Check if the Save button was pressed (updatedValues differs from variableValues)
                If Not updatedValues.Equals(variableValues) Then
                    ' Update the original variables with the returned values
                    If updatedValues.ContainsKey("APIKey") Then context.INI_APIKeyBack = updatedValues("APIKey")
                    If updatedValues.ContainsKey("Temperature") Then context.INI_Temperature = updatedValues("Temperature")
                    If updatedValues.ContainsKey("Timeout") Then context.INI_Timeout = CLng(updatedValues("Timeout"))
                    If updatedValues.ContainsKey("MaxOutputToken") Then context.INI_MaxOutputToken = CInt(updatedValues("MaxOutputToken"))
                    If updatedValues.ContainsKey("Model") Then context.INI_Model = updatedValues("Model")
                    If updatedValues.ContainsKey("Endpoint") Then context.INI_Endpoint = updatedValues("Endpoint")
                    If updatedValues.ContainsKey("HeaderA") Then context.INI_HeaderA = updatedValues("HeaderA")
                    If updatedValues.ContainsKey("HeaderB") Then context.INI_HeaderB = updatedValues("HeaderB")
                    If updatedValues.ContainsKey("APICall") Then context.INI_APICall = updatedValues("APICall")
                    If updatedValues.ContainsKey("Response") Then context.INI_Response = updatedValues("Response")
                    If updatedValues.ContainsKey("DoubleS") Then context.INI_DoubleS = CBool(updatedValues("DoubleS"))
                    If updatedValues.ContainsKey("PreCorrection") Then context.INI_PreCorrection = updatedValues("PreCorrection")
                    If updatedValues.ContainsKey("PostCorrection") Then context.INI_PostCorrection = updatedValues("PostCorrection")
                    If updatedValues.ContainsKey("APIEncrypted") Then context.INI_APIEncrypted = CBool(updatedValues("APIEncrypted"))
                    If updatedValues.ContainsKey("APIKeyPrefix") Then context.INI_APIKeyPrefix = updatedValues("APIKeyPrefix")
                    If updatedValues.ContainsKey("MarkupMethodOutlook") Then context.INI_MarkupMethodOutlook = CInt(updatedValues("MarkupMethodOutlook"))
                    If updatedValues.ContainsKey("MarkupDiffCap") Then context.INI_MarkupDiffCap = CInt(updatedValues("MarkupDiffCap"))
                    If updatedValues.ContainsKey("MarkupRegexCap") Then context.INI_MarkupRegexCap = CInt(updatedValues("MarkupRegexCap"))
                    If updatedValues.ContainsKey("ChatCap") Then context.INI_ChatCap = CInt(updatedValues("ChatCap"))
                    If updatedValues.ContainsKey("OAuth2") Then context.INI_OAuth2 = CBool(updatedValues("OAuth2"))
                    If updatedValues.ContainsKey("OAuth2ClientMail") Then context.INI_OAuth2ClientMail = updatedValues("OAuth2ClientMail")
                    If updatedValues.ContainsKey("OAuth2Scopes") Then context.INI_OAuth2Scopes = updatedValues("OAuth2Scopes")
                    If updatedValues.ContainsKey("OAuth2Endpoint") Then context.INI_OAuth2Endpoint = updatedValues("OAuth2Endpoint")
                    If updatedValues.ContainsKey("OAuth2ATExpiry") Then context.INI_OAuth2ATExpiry = CLng(updatedValues("OAuth2ATExpiry"))
                    If updatedValues.ContainsKey("SecondAPI") Then context.INI_SecondAPI = CBool(updatedValues("SecondAPI"))
                    If updatedValues.ContainsKey("APIKey_2") Then context.INI_APIKeyBack_2 = updatedValues("APIKey_2")
                    If updatedValues.ContainsKey("Temperature_2") Then context.INI_Temperature_2 = updatedValues("Temperature_2")
                    If updatedValues.ContainsKey("Timeout_2") Then context.INI_Timeout_2 = CLng(updatedValues("Timeout_2"))
                    If updatedValues.ContainsKey("MaxOutputToken_2") Then context.INI_MaxOutputToken_2 = CInt(updatedValues("MaxOutputToken_2"))
                    If updatedValues.ContainsKey("Model_2") Then context.INI_Model_2 = updatedValues("Model_2")
                    If updatedValues.ContainsKey("Endpoint_2") Then context.INI_Endpoint_2 = updatedValues("Endpoint_2")
                    If updatedValues.ContainsKey("HeaderA_2") Then context.INI_HeaderA_2 = updatedValues("HeaderA_2")
                    If updatedValues.ContainsKey("HeaderB_2") Then context.INI_HeaderB_2 = updatedValues("HeaderB_2")
                    If updatedValues.ContainsKey("APICall_2") Then context.INI_APICall_2 = updatedValues("APICall_2")
                    If updatedValues.ContainsKey("Response_2") Then context.INI_Response_2 = updatedValues("Response_2")
                    If updatedValues.ContainsKey("APIEncrypted_2") Then context.INI_APIEncrypted_2 = CBool(updatedValues("APIEncrypted_2"))
                    If updatedValues.ContainsKey("APIKeyPrefix_2") Then context.INI_APIKeyPrefix_2 = updatedValues("APIKeyPrefix_2")
                    If updatedValues.ContainsKey("OAuth2_2") Then context.INI_OAuth2_2 = CBool(updatedValues("OAuth2_2"))
                    If updatedValues.ContainsKey("OAuth2ClientMail_2") Then context.INI_OAuth2ClientMail_2 = updatedValues("OAuth2ClientMail_2")
                    If updatedValues.ContainsKey("OAuth2Scopes_2") Then context.INI_OAuth2Scopes_2 = updatedValues("OAuth2Scopes_2")
                    If updatedValues.ContainsKey("OAuth2Endpoint_2") Then context.INI_OAuth2Endpoint_2 = updatedValues("OAuth2Endpoint_2")
                    If updatedValues.ContainsKey("OAuth2ATExpiry_2") Then context.INI_OAuth2ATExpiry_2 = CLng(updatedValues("OAuth2ATExpiry_2"))
                    If updatedValues.ContainsKey("APIDebug") Then context.INI_APIDebug = CBool(updatedValues("APIDebug"))
                    If updatedValues.ContainsKey("UsageRestrictions") Then context.INI_UsageRestrictions = updatedValues("UsageRestrictions")
                    If updatedValues.ContainsKey("Language1") Then context.INI_Language1 = updatedValues("Language1")
                    If updatedValues.ContainsKey("Language2") Then context.INI_Language2 = updatedValues("Language2")
                    If updatedValues.ContainsKey("KeepFormat1") Then context.INI_KeepFormat1 = CBool(updatedValues("KeepFormat1"))
                    If updatedValues.ContainsKey("KeepFormat2") Then context.INI_KeepFormat2 = CBool(updatedValues("KeepFormat2"))
                    If updatedValues.ContainsKey("KeepFormatCap") Then context.INI_KeepFormatCap = CInt(updatedValues("KeepFormatCap"))
                    If updatedValues.ContainsKey("KeepParaFormatInline") Then context.INI_KeepParaFormatInline = CBool(updatedValues("KeepParaFormatInline"))
                    If updatedValues.ContainsKey("ReplaceText1") Then context.INI_ReplaceText1 = CBool(updatedValues("ReplaceText1"))
                    If updatedValues.ContainsKey("ReplaceText2") Then context.INI_ReplaceText2 = CBool(updatedValues("ReplaceText2"))
                    If updatedValues.ContainsKey("DoMarkupOutlook") Then context.INI_DoMarkupOutlook = CBool(updatedValues("DoMarkupOutlook"))
                    If updatedValues.ContainsKey("DoMarkupWord") Then context.INI_DoMarkupWord = CBool(updatedValues("DoMarkupWord"))
                    If updatedValues.ContainsKey("SP_Translate") Then context.SP_Translate = updatedValues("SP_Translate")
                    If updatedValues.ContainsKey("SP_Correct") Then context.SP_Correct = updatedValues("SP_Correct")
                    If updatedValues.ContainsKey("SP_Improve") Then context.SP_Improve = updatedValues("SP_Improve")
                    If updatedValues.ContainsKey("SP_Explain") Then context.SP_Explain = updatedValues("SP_Explain")
                    If updatedValues.ContainsKey("SP_SuggestTitles") Then context.SP_SuggestTitles = updatedValues("SP_SuggestTitles")
                    If updatedValues.ContainsKey("SP_Friendly") Then context.SP_Friendly = updatedValues("SP_Friendly")
                    If updatedValues.ContainsKey("SP_Convincing") Then context.SP_Convincing = updatedValues("SP_Convincing")
                    If updatedValues.ContainsKey("SP_NoFillers") Then context.SP_NoFillers = updatedValues("SP_NoFillers")
                    If updatedValues.ContainsKey("SP_Podcast") Then context.SP_Podcast = updatedValues("SP_Podcast")
                    If updatedValues.ContainsKey("SP_Shorten") Then context.SP_Shorten = updatedValues("SP_Shorten")
                    If updatedValues.ContainsKey("SP_Summarize") Then context.SP_Summarize = updatedValues("SP_Summarize")
                    If updatedValues.ContainsKey("SP_MailReply") Then context.SP_MailReply = updatedValues("SP_MailReply")
                    If updatedValues.ContainsKey("SP_MailSumup") Then context.SP_MailSumup = updatedValues("SP_MailSumup")
                    If updatedValues.ContainsKey("SP_FreestyleText") Then context.SP_FreestyleText = updatedValues("SP_FreestyleText")
                    If updatedValues.ContainsKey("SP_FreestyleNoText") Then context.SP_FreestyleNoText = updatedValues("SP_FreestyleNoText")
                    If updatedValues.ContainsKey("SP_SwitchParty") Then context.SP_SwitchParty = updatedValues("SP_SwitchParty")
                    If updatedValues.ContainsKey("SP_Anonymize") Then context.SP_Anonymize = updatedValues("SP_Anonymize")
                    If updatedValues.ContainsKey("SP_ContextSearch") Then context.SP_ContextSearch = updatedValues("SP_ContextSearch")
                    If updatedValues.ContainsKey("SP_ContextSearchMulti") Then context.SP_ContextSearchMulti = updatedValues("SP_ContextSearchMulti")
                    If updatedValues.ContainsKey("SP_RangeOfCells") Then context.SP_RangeOfCells = updatedValues("SP_RangeOfCells")
                    If updatedValues.ContainsKey("SP_WriteNeatly") Then context.SP_WriteNeatly = updatedValues("SP_WriteNeatly")
                    If updatedValues.ContainsKey("SP_Add_KeepFormulasIntact") Then context.SP_Add_KeepFormulasIntact = updatedValues("SP_Add_KeepFormulasIntact")
                    If updatedValues.ContainsKey("SP_Add_KeepHTMLIntact") Then context.SP_Add_KeepHTMLIntact = updatedValues("SP_Add_KeepHTMLIntact")
                    If updatedValues.ContainsKey("SP_Add_KeepInlineIntact") Then context.SP_Add_KeepInlineIntact = updatedValues("SP_Add_KeepInlineIntact")
                    If updatedValues.ContainsKey("SP_Add_Bubbles") Then context.SP_Add_Bubbles = updatedValues("SP_Add_Bubbles")
                    If updatedValues.ContainsKey("SP_Add_Revisions") Then context.SP_Add_Revisions = updatedValues("SP_Add_Revisions")
                    If updatedValues.ContainsKey("SP_MarkupRegex") Then context.SP_MarkupRegex = updatedValues("SP_MarkupRegex")
                    If updatedValues.ContainsKey("SP_ChatWord") Then context.SP_ChatWord = updatedValues("SP_ChatWord")
                    If updatedValues.ContainsKey("SP_Add_ChatWord_Commands") Then context.SP_Add_ChatWord_Commands = updatedValues("SP_Add_ChatWord_Commands")
                    If updatedValues.ContainsKey("ISearch") Then context.INI_ISearch = CBool(updatedValues("ISearch"))
                    If updatedValues.ContainsKey("ISearch_Approve") Then context.INI_ISearch_Approve = CBool(updatedValues("ISearch_Approve"))
                    If updatedValues.ContainsKey("ISearch_URL") Then context.INI_ISearch_URL = updatedValues("ISearch_URL")
                    If updatedValues.ContainsKey("ISearch_ResponseMask1") Then context.INI_ISearch_ResponseMask1 = updatedValues("ISearch_ResponseMask1")
                    If updatedValues.ContainsKey("ISearch_ResponseMask2") Then context.INI_ISearch_ResponseMask2 = updatedValues("ISearch_ResponseMask2")
                    If updatedValues.ContainsKey("ISearch_Name") Then context.INI_ISearch_Name = updatedValues("ISearch_Name")
                    If updatedValues.ContainsKey("ISearch_Tries") Then context.INI_ISearch_Tries = CInt(updatedValues("ISearch_Tries"))
                    If updatedValues.ContainsKey("ISearch_Results") Then context.INI_ISearch_Results = CInt(updatedValues("ISearch_Results"))
                    If updatedValues.ContainsKey("ISearch_MaxDepth") Then context.INI_ISearch_MaxDepth = CInt(updatedValues("ISearch_MaxDepth"))
                    If updatedValues.ContainsKey("ISearch_Timeout") Then context.INI_ISearch_Timeout = CLng(updatedValues("ISearch_Timeout"))
                    If updatedValues.ContainsKey("ISearch_SearchTerm_SP") Then context.INI_ISearch_SearchTerm_SP = updatedValues("ISearch_SearchTerm_SP")
                    If updatedValues.ContainsKey("ISearch_Apply_SP") Then context.INI_ISearch_Apply_SP = updatedValues("ISearch_Apply_SP")
                    If updatedValues.ContainsKey("ISearch_Apply_SP_Markup") Then context.INI_ISearch_Apply_SP_Markup = updatedValues("ISearch_Apply_SP_Markup")
                    If updatedValues.ContainsKey("Lib") Then context.INI_Lib = CBool(updatedValues("Lib"))
                    If updatedValues.ContainsKey("Lib_File") Then context.INI_Lib_File = updatedValues("Lib_File")
                    If updatedValues.ContainsKey("Lib_Timeout") Then context.INI_Lib_Timeout = CLng(updatedValues("Lib_Timeout"))
                    If updatedValues.ContainsKey("Lib_Find_SP") Then context.INI_Lib_Find_SP = updatedValues("Lib_Find_SP")
                    If updatedValues.ContainsKey("Lib_Apply_SP") Then context.INI_Lib_Apply_SP = updatedValues("Lib_Apply_SP")
                    If updatedValues.ContainsKey("Lib_Apply_SP_Markup") Then context.INI_Lib_Apply_SP_Markup = updatedValues("Lib_Apply_SP_Markup")
                    If updatedValues.ContainsKey("ISearch_Apply_SP_Markup") Then context.INI_ISearch_Apply_SP_Markup = updatedValues("ISearch_Apply_SP_Markup")
                    If updatedValues.ContainsKey("MarkupMethodHelper") Then context.INI_MarkupMethodHelper = CInt(updatedValues("MarkupMethodHelper"))
                    If updatedValues.ContainsKey("MarkupMethodWord") Then context.INI_MarkupMethodWord = CInt(updatedValues("MarkupMethodWord"))
                    If updatedValues.ContainsKey("ShortcutsWordExcel") Then context.INI_ShortcutsWordExcel = updatedValues("ShortcutsWordExcel")
                    If updatedValues.ContainsKey("ContextMenu") Then context.INI_ContextMenu = updatedValues("ContextMenu")
                    If updatedValues.ContainsKey("UpdateCheckInterval") Then context.INI_UpdateCheckInterval = CInt(updatedValues("UpdateCheckInterval"))
                    If updatedValues.ContainsKey("UpdatePath") Then context.INI_UpdatePath = updatedValues("UpdatePath")
                    If updatedValues.ContainsKey("SpeechModelPath") Then context.INI_SpeechModelPath = updatedValues("SpeechModelPath")
                    If updatedValues.ContainsKey("TTSEndpoint") Then context.INI_TTSEndpoint = updatedValues("TTSEndpoint")
                    If updatedValues.ContainsKey("PromptLib") Then context.INI_PromptLibPath = updatedValues("PromptLib")
                    If updatedValues.ContainsKey("PromptLib_Transcript") Then context.INI_PromptLibPath_Transcript = updatedValues("PromptLib_Transcript")

                    ' Call UpdateAppConfig after all updates
                    UpdateAppConfig(context)
                End If
            End If
        End Sub


        Public Shared Sub ShowAboutWindow(owner As System.Windows.Forms.Form, context As ISharedContext)
            ' Example of using the same font and appearance as ShowWindowsSettings
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            ' Adjusted dimensions 
            Dim formWidth As Integer = CInt(450)

            ' Calculate height based on text content
            Dim ExpireText As String = $"{vbCrLf}{vbCrLf}(expires on {LicensedTill.ToString("dd-MMM-yyyy")})"
            Dim testRichTextBox As New System.Windows.Forms.RichTextBox() With {
        .Font = standardFont,
        .Text = $"{AN}{vbCrLf}{context.RDV}{ExpireText}{vbCrLf}{vbCrLf}Created by David Rosenthal{vbCrLf}david.rosenthal@vischer.com{vbCrLf}{vbCrLf}VISCHER AG, Zürich, Switzerland{vbCrLf}Swiss Law & Tax{vbCrLf}{vbCrLf}All rights reserved.{vbCrLf}{vbCrLf}{AN4}"
    }
            Dim graphics As System.Drawing.Graphics = testRichTextBox.CreateGraphics()
            Dim textSize As System.Drawing.SizeF = graphics.MeasureString(testRichTextBox.Text, standardFont, formWidth - 40)
            Dim formHeight As Integer = CInt(textSize.Height + 240 + 20) ' Add padding for margins, logo, buttons, and 1–2 extra lines
            graphics.Dispose()
            testRichTextBox.Dispose()

            ' Create the form
            Dim aboutForm As New System.Windows.Forms.Form() With {
                        .FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
                        .StartPosition = System.Windows.Forms.FormStartPosition.CenterParent,
                        .ClientSize = New System.Drawing.Size(formWidth, formHeight),
                        .BackColor = owner.BackColor,
                        .Font = standardFont,
                        .MaximizeBox = False,
                        .MinimizeBox = False,
                        .ControlBox = False,
                        .ShowInTaskbar = False
                    }

            ' Add a logo
            Dim logoSize As Integer = 120
            Dim logo As New System.Windows.Forms.PictureBox() With {
                        .Image = My.Resources.Red_Ink_Logo,
                        .SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom,
                        .Size = New System.Drawing.Size(logoSize, logoSize),
                        .Location = New System.Drawing.Point((formWidth - logoSize) \ 2, 20)
                    }
            aboutForm.Controls.Add(logo)

            ' Add a RichTextBox for text
            Dim aboutTextBox As New System.Windows.Forms.RichTextBox() With {
                        .ReadOnly = True,
                        .BorderStyle = System.Windows.Forms.BorderStyle.None,
                        .BackColor = owner.BackColor,
                        .Font = standardFont,
                        .DetectUrls = True
                    }

            Dim topOffset As Integer = logo.Bottom + 10
            Dim bottomPadding As Integer = 100
            Dim availableHeight As Integer = formHeight - topOffset - bottomPadding
            aboutTextBox.Size = New System.Drawing.Size(formWidth - 40, availableHeight)
            aboutTextBox.Location = New System.Drawing.Point(20, topOffset)
            aboutForm.Controls.Add(aboutTextBox)

            Dim aboutContent As String =
        $"{AN}<P>{context.RDV}{ExpireText}<P><P>Created by David Rosenthal<P>david.rosenthal@vischer.com<P><P>VISCHER AG, Zürich, Switzerland<P>Swiss Law & Tax<P><P>All rights reserved.<P><P>{AN4}"

            ' Replace <P> with vbCrLf
            Dim plainText As New System.Text.StringBuilder()

            While aboutContent.Contains("<P>")
                Dim index = aboutContent.IndexOf("<P>")
                plainText.Append(aboutContent.Substring(0, index))
                plainText.Append(vbCrLf)
                aboutContent = aboutContent.Substring(index + 3)
            End While
            plainText.Append(aboutContent)

            ' Set the text and apply formatting
            aboutTextBox.Text = plainText.ToString()

            ' Center the text
            aboutTextBox.SelectAll()
            aboutTextBox.SelectionAlignment = HorizontalAlignment.Center
            aboutTextBox.DeselectAll()

            ' Hide the blinking cursor
            aboutTextBox.SelectionStart = aboutTextBox.Text.Length
            aboutTextBox.SelectionLength = 0
            aboutTextBox.ScrollToCaret() ' Ensures the caret is out of visible range

            ' Add a handler for link clicks
            AddHandler aboutTextBox.LinkClicked,
        Sub(sender, e)
            Try
                Process.Start(New ProcessStartInfo(e.LinkText) With {.UseShellExecute = True})
            Catch ex As System.Exception
                MessageBox.Show("Error in ShowAboutWindow - unable to open the link.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

            ' Add a "License" button
            Dim licenseButton As New System.Windows.Forms.Button() With {
                        .Text = "3rd Party Software Used",
                        .Size = New System.Drawing.Size(300, 30),
                        .Location = New System.Drawing.Point((formWidth - 300) \ 2, aboutTextBox.Bottom + 10)
                    }
            AddHandler licenseButton.Click, Sub(sender, e) ShowRTFCustomMessageBox(ConvertMarkupToRTF(LicenseText), AN)
            aboutForm.Controls.Add(licenseButton)

            ' Add an OK button
            Dim okButton As New System.Windows.Forms.Button() With {
                        .Text = "OK",
                        .Size = New System.Drawing.Size(80, 30),
                        .Location = New System.Drawing.Point((formWidth - 80) \ 2, formHeight - 40)
                    }
            AddHandler okButton.Click, Sub(sender, e) aboutForm.Close()
            aboutForm.Controls.Add(okButton)

            ' Show the form
            aboutForm.ShowDialog(owner)
        End Sub



        Public Shared Function ShowPromptSelector(filePath As String, enableMarkup As Boolean, enableBubbles As Boolean, Context As ISharedContext) As (String, Boolean, Boolean, Boolean)

            filePath = ExpandEnvironmentVariables(filePath)

            Dim LoadResult = LoadPrompts(filePath, Context)
            Dim NoBubbles As Boolean = False
            Dim NoMarkup As Boolean = False

            If enableMarkup = Nothing Then
                NoMarkup = True
                enableMarkup = False
            End If

            If enableBubbles = Nothing Then
                NoBubbles = True
                enableBubbles = False
            End If

            If LoadResult <> 0 Then Return ("", False, False, False)

            ' Create the form
            Dim settingsForm As New Form With {
                        .Text = "Select Prompt",
                        .Width = 900,
                        .Height = 650,
                        .StartPosition = FormStartPosition.CenterScreen,
                        .Padding = New Padding(10)
                    }

            ' Set icon
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            settingsForm.Icon = Icon.FromHandle(bmp.GetHicon())

            ' Set a predefined font
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            settingsForm.Font = standardFont

            ' Create a table layout panel for structured arrangement
            Dim layout As New TableLayoutPanel With {
                        .Dock = DockStyle.Fill,
                        .ColumnCount = 2,
                        .RowCount = 3,
                        .Padding = New Padding(10)
                    }

            ' Configure column and row styles
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 70))
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 50))

            settingsForm.Controls.Add(layout)

            ' Create listbox for prompt titles
            Dim titleListBox As New ListBox With {
                        .Dock = DockStyle.Fill,
                        .Margin = New Padding(10)
                    }
            titleListBox.Items.AddRange(Context.PromptTitles.ToArray())
            layout.Controls.Add(titleListBox, 0, 0)

            ' Create textbox for prompt content
            Dim promptTextBox As New TextBox With {
                            .Dock = DockStyle.Fill,
                            .Multiline = True,
                            .ReadOnly = True,
                            .ScrollBars = ScrollBars.Vertical,
                            .Margin = New Padding(10)
                        }
            layout.Controls.Add(promptTextBox, 1, 0)

            ' Ensure equal sizes for selector and preview
            AddHandler settingsForm.Resize, Sub()
                                                Dim equalHeight = layout.GetRowHeights()(0)
                                                titleListBox.Height = equalHeight
                                                promptTextBox.Height = equalHeight
                                            End Sub

            ' Preselect the first prompt
            If Context.PromptTitles.Count > 0 Then
                titleListBox.SelectedIndex = 0
                promptTextBox.Text = Context.PromptLibrary(0).Replace("\n", vbCrLf)
            End If

            ' Handle title selection
            AddHandler titleListBox.SelectedIndexChanged, Sub()
                                                              Dim selectedIndex = titleListBox.SelectedIndex
                                                              If selectedIndex >= 0 Then
                                                                  Dim selectedPrompt = Context.PromptLibrary(selectedIndex).Replace("\n", vbCrLf)
                                                                  promptTextBox.Text = selectedPrompt
                                                              End If
                                                          End Sub

            ' Handle Enter key to confirm selection
            AddHandler titleListBox.KeyDown, Sub(sender, e)
                                                 If e.KeyCode = Keys.Enter Then
                                                     settingsForm.DialogResult = DialogResult.OK
                                                     settingsForm.Close()
                                                 End If
                                             End Sub

            ' Create a panel for checkboxes
            Dim checkboxPanel As New FlowLayoutPanel With {
                        .FlowDirection = FlowDirection.TopDown,
                        .Dock = DockStyle.Fill,
                        .Margin = New Padding(10)
                    }
            layout.Controls.Add(checkboxPanel, 0, 1)

            ' Checkboxes
            Dim markupCheckbox As New System.Windows.Forms.CheckBox With {
                        .Text = "The output shall be provided as a markup",
                        .AutoSize = True,
                        .Enabled = enableMarkup,
                        .Visible = Not NoMarkup
                    }

            Dim clipboardCheckbox As New System.Windows.Forms.CheckBox With {
                        .Text = "The output shall be shown and put in the clipboard",
                        .AutoSize = True
                    }

            Dim bubblesCheckbox As New System.Windows.Forms.CheckBox With {
                        .Text = "The output shall be provided in the form of bubbles",
                        .AutoSize = True,
                        .Enabled = enableBubbles,
                        .Visible = Not NoBubbles
                    }

            checkboxPanel.Controls.Add(markupCheckbox)
            checkboxPanel.Controls.Add(clipboardCheckbox)
            checkboxPanel.Controls.Add(bubblesCheckbox)

            ' Ensure mutual exclusivity of checkboxes
            AddHandler markupCheckbox.CheckedChanged, Sub()
                                                          If markupCheckbox.Checked Then
                                                              bubblesCheckbox.Checked = False
                                                              clipboardCheckbox.Checked = False
                                                          End If
                                                      End Sub

            AddHandler bubblesCheckbox.CheckedChanged, Sub()
                                                           If bubblesCheckbox.Checked Then
                                                               markupCheckbox.Checked = False
                                                               clipboardCheckbox.Checked = False
                                                           End If
                                                       End Sub

            AddHandler clipboardCheckbox.CheckedChanged, Sub()
                                                             If clipboardCheckbox.Checked Then
                                                                 markupCheckbox.Checked = False
                                                                 bubblesCheckbox.Checked = False
                                                             End If
                                                         End Sub

            ' File path label
            Dim filePathLabel As New System.Windows.Forms.Label With {
                            .Text = $"Source: {filePath}",
                            .AutoSize = True,
                            .MaximumSize = New Size(layout.Width, 0),
                            .Margin = New Padding(10)
                        }
            layout.Controls.Add(filePathLabel, 1, 1)

            ' Add OK, Cancel, and Edit buttons
            Dim buttonPanel As New FlowLayoutPanel With {
                            .FlowDirection = FlowDirection.LeftToRight,
                            .Dock = DockStyle.Bottom,
                            .Padding = New Padding(10)
                        }
            layout.Controls.Add(buttonPanel, 0, 2)
            layout.SetColumnSpan(buttonPanel, 2)

            Dim okButton As New Button With {
                        .Text = "OK",
                        .AutoSize = True,
                        .DialogResult = DialogResult.OK
                    }

            Dim cancelButton As New Button With {
                        .Text = "Cancel",
                        .AutoSize = True,
                        .DialogResult = DialogResult.Cancel
                    }

            Dim editButton As New Button With {
                        .Text = "Edit",
                        .AutoSize = True,
                        .Anchor = AnchorStyles.Right
                    }
            buttonPanel.Controls.Add(okButton)
            buttonPanel.Controls.Add(cancelButton)
            buttonPanel.Controls.Add(editButton)

            ' Align edit button to the right
            Dim spacer As New Panel With {
                        .Dock = DockStyle.Fill
                    }
            buttonPanel.Controls.SetChildIndex(editButton, buttonPanel.Controls.Count - 1)

            ' Handle Edit button click
            AddHandler editButton.Click, Sub()
                                             Dim editorForm As New Form With {
                                                 .Text = "Edit Prompt Library",
                                                 .Width = 800,
                                                 .Height = 600,
                                                 .StartPosition = FormStartPosition.CenterParent,
                                                 .Padding = New Padding(10)
                                             }

                                             ' Set icon for editor
                                             editorForm.Icon = Icon.FromHandle(bmp.GetHicon())

                                             Dim descriptionLabel As New System.Windows.Forms.Label With {
                                                 .Text = $"You can now edit your prompts (stored at {filePath}). Make sure that on each line, the description and the prompt is separated by a '|'; you can use ';' for indicating comments.",
                                                 .Dock = DockStyle.Top,
                                                 .Font = standardFont,
                                                 .AutoSize = True,
                                                 .MaximumSize = New Size(editorForm.Width - 20, 0),
                                                 .Margin = New Padding(10, 20, 20, 20)
                                             }

                                             Dim editorTextBox As New TextBox With {
                                                 .Multiline = True,
                                                 .Dock = DockStyle.Fill,
                                                 .ScrollBars = ScrollBars.Both,
                                                 .Font = standardFont,
                                                 .Margin = New Padding(20),
                                                 .Height = 400
                                             }

                                             ' Load file content into editor
                                             editorTextBox.Text = System.IO.File.ReadAllText(filePath)
                                             editorTextBox.SelectionStart = 0
                                             editorTextBox.SelectionLength = 0

                                             Dim editorButtonPanel As New FlowLayoutPanel With {
                                                 .FlowDirection = FlowDirection.LeftToRight,
                                                 .Dock = DockStyle.Bottom,
                                                 .Padding = New Padding(10),
                                                 .AutoSize = True
                                             }

                                             Dim saveButton As New Button With {
                                                 .Text = "Save",
                                                 .Font = standardFont,
                                                 .AutoSize = True
                                             }

                                             Dim cancelEditButton As New Button With {
                                                 .Text = "Cancel",
                                                 .Font = standardFont,
                                                 .AutoSize = True
                                             }

                                             AddHandler cancelEditButton.Click, Sub()
                                                                                    editorForm.Close()
                                                                                End Sub

                                             AddHandler saveButton.Click, Sub()
                                                                              System.IO.File.WriteAllText(filePath, editorTextBox.Text)
                                                                              editorForm.Close()

                                                                              ' Reload prompts after saving
                                                                              LoadPrompts(filePath, Context)
                                                                              titleListBox.Items.Clear()
                                                                              titleListBox.Items.AddRange(Context.PromptTitles.ToArray())
                                                                              If Context.PromptTitles.Count > 0 Then
                                                                                  titleListBox.SelectedIndex = 0
                                                                                  promptTextBox.Text = Context.PromptLibrary(0).Replace("\n", vbCrLf)
                                                                              End If
                                                                              titleListBox.Focus()
                                                                          End Sub

                                             editorButtonPanel.Controls.Add(saveButton)
                                             editorButtonPanel.Controls.Add(cancelEditButton)

                                             editorForm.Controls.Add(editorTextBox)
                                             editorForm.Controls.Add(descriptionLabel)
                                             editorForm.Controls.Add(editorButtonPanel)
                                             editorForm.ShowDialog()
                                             titleListBox.Focus()
                                         End Sub

            ' Show the form
            Dim result As DialogResult = settingsForm.ShowDialog()

            If result = DialogResult.OK Then
                Dim selectedIndex = titleListBox.SelectedIndex
                If selectedIndex >= 0 Then
                    Return (
                        Context.PromptLibrary(selectedIndex),
                        markupCheckbox.Checked,
                        bubblesCheckbox.Checked,
                        clipboardCheckbox.Checked
                    )
                End If
            End If

            ' Return defaults if cancelled or no selection
            Return ("", False, False, False)
        End Function


        Public Shared Function LoadPrompts(filePath As String, context As ISharedContext) As Integer

            ' Initialize the return code to 0 (no error)
            Dim returnCode As Integer = 0

            filePath = ExpandEnvironmentVariables(filePath)

            Try
                ' Verify the file exists
                If Not System.IO.File.Exists(filePath) Then
                    ShowCustomMessageBox("The prompt library file was not found.")
                    Return 1
                End If

                context.PromptTitles.Clear()
                context.PromptLibrary.Clear()

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
                            context.PromptTitles.Add(title)
                            context.PromptLibrary.Add(prompt)
                        End If
                    End If
                Next

                ' Check if no prompts were found
                If context.PromptLibrary.Count = 0 Then
                    returnCode = 3
                    ShowCustomMessageBox("No prompts have been found in the configured prompt library file.")
                End If

            Catch ex As System.IO.FileNotFoundException
                returnCode = 1
                ShowCustomMessageBox("The prompt library file was not found: " & ex.Message)

            Catch ex As IndexOutOfRangeException
                returnCode = 2
                ShowCustomMessageBox("The format of the prompt library file is not correct (is a '|' or text thereafter missing?): " & ex.Message)

            Catch ex As Exception
                returnCode = 99
                ShowCustomMessageBox("An unexpected error occurred while loading prompts: " & ex.Message)
            End Try

            Return returnCode
        End Function

        Public Shared Sub PutInClipboard(text As String)
            Dim thread As New Threading.Thread(Sub()
                                                   ' Check if the text is RTF formatted
                                                   If text.StartsWith("{\rtf") Then
                                                       ' Set RTF content to the clipboard
                                                       Clipboard.SetData(DataFormats.Rtf, text)
                                                   Else
                                                       ' Set plain text to the clipboard
                                                       Clipboard.SetText(text)
                                                   End If
                                               End Sub)

            ' Ensure the thread is STA (Single Thread Apartment), as required by the clipboard
            thread.SetApartmentState(Threading.ApartmentState.STA)
            thread.Start()
            thread.Join()

        End Sub


        Public Class InfoBox

            Inherits Form

            Private Shared InfoBox As InfoBox
            Private timer As System.Windows.Forms.Timer
            Private label As System.Windows.Forms.Label

            Private Sub New(ByVal text As String, ByVal duration As Integer)
                ' Set form properties
                Me.Text = ""
                Me.FormBorderStyle = FormBorderStyle.None
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.BackColor = ColorTranslator.FromWin32(&H8000000F)
                Me.TopMost = True

                ' Create and add the App logo PictureBox
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                Dim iconPictureBox As New PictureBox()
                iconPictureBox.Image = bmp
                iconPictureBox.SizeMode = PictureBoxSizeMode.Zoom
                iconPictureBox.Size = New Size(32, 32) ' Icon size
                iconPictureBox.Location = New System.Drawing.Point(10, 10) ' Top-left corner
                Me.Controls.Add(iconPictureBox)

                ' Initialize label
                label = New System.Windows.Forms.Label()
                label.Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                label.TextAlign = ContentAlignment.MiddleLeft
                label.MaximumSize = New Size(450, 240)
                label.Width = 450
                label.Height = 240
                label.Text = text
                label.AutoSize = True
                label.AutoEllipsis = True
                'SetWrappedText(label, text)  ' not necessary, if autoellipsis is set

                ' Adjust form size dynamically to accommodate PictureBox and label
                Dim contentRight As Integer = iconPictureBox.Right + 10
                Me.ClientSize = New Size(Math.Max(contentRight + label.Width + 10, iconPictureBox.Width + 20), Math.Max(label.Height + 20, iconPictureBox.Height + 20))

                ' Position label below the icon
                label.Location = New System.Drawing.Point(contentRight, 10)
                Me.Controls.Add(label)


                ' Initialize and start timer if duration > 0
                If duration > 0 Then
                    timer = New System.Windows.Forms.Timer()
                    timer.Interval = duration * 1000
                    AddHandler timer.Tick, AddressOf Timer_Tick
                    timer.Start()
                End If
            End Sub

            Private Sub SetWrappedText(lbl As System.Windows.Forms.Label, text As String)
                ' Set the wrapped text in the label
                lbl.Text = text

                Using g As Graphics = lbl.CreateGraphics()
                    ' Measure the size of the text
                    Dim size As SizeF = g.MeasureString(text, lbl.Font, lbl.Width)

                    ' Check if the text exceeds the maximum label height
                    Dim lineHeight As Single = lbl.Font.GetHeight(g)
                    Dim maxLines As Integer = Math.Floor(lbl.MaximumSize.Height / lineHeight)
                    Dim textLines As Integer = Math.Ceiling(size.Height / lineHeight)

                    If textLines > maxLines Then
                        ' Truncate and add ellipsis if exceeding the maximum visible lines
                        Dim visibleText As String = text.Substring(0, Math.Min(text.Length, lbl.Width * maxLines \ lbl.Font.Size)) & " (...)"
                        lbl.Text = visibleText
                    End If
                End Using
            End Sub


            Private Sub Timer_Tick(ByVal sender As Object, ByVal e As EventArgs)
                Me.Close()
            End Sub

            Public Shared Sub ShowInfoBox(ByVal text As String, Optional ByVal duration As Integer = 0)
                ' Close current InfoBox if open
                If InfoBox IsNot Nothing Then
                    InfoBox.Close()
                End If

                ' If text is empty, return without creating a new form
                If String.IsNullOrEmpty(text) Then
                    Return
                End If

                ' Create a new InfoBox instance and display it
                InfoBox = New InfoBox(text, duration)
                InfoBox.Show()
                InfoBox.Refresh()
                System.Windows.Forms.Application.DoEvents()
            End Sub

        End Class



        Public Shared Function ReadTextFile(filePath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Try
                ' Normalize and check the path
                filePath = Path.GetFullPath(filePath)
                If Not File.Exists(filePath) Then
                    Return If(ReturnErrorInsteadOfEmpty, "Error: File not found.", "")
                End If

                ' Use StreamReader for reading
                Using reader As New StreamReader(filePath, System.Text.Encoding.UTF8, True)
                    Dim content As String = reader.ReadToEnd()
                    Return content
                End Using
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading file: {ex.Message}", "")
            End Try
        End Function

        Public Shared Function ReadRtfAsText(ByVal rtfPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Try
                Dim rtfContent As String = File.ReadAllText(rtfPath)
                Using rtb As New RichTextBox()
                    rtb.Visible = False
                    rtb.Rtf = rtfContent
                    Return rtb.Text
                End Using
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading RTF: {ex.Message}", "")
            End Try
        End Function

        Public Shared Function ReadWordDocument(ByVal docPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Dim app As Microsoft.Office.Interop.Word.Application = Nothing
            Dim doc As Document = Nothing

            Try
                Try
                    ' Try to attach to an existing Word instance.
                    app = CType(Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
                Catch ex As System.Exception
                    ' If Word is not running, create a new Word application.
                    app = New Microsoft.Office.Interop.Word.Application With {.Visible = False}
                End Try

                ' Open the Word document in read-only mode
                doc = app.Documents.Open(docPath, ReadOnly:=True, Visible:=False)

                ' Extract the content text
                Dim text As String = doc.Content.Text

                ' Close the document without saving changes
                doc.Close(SaveChanges:=False)

                ' Return the extracted text
                Return text

            Catch ex As System.Exception
                ' Ensure the document is closed in case of an error
                If doc IsNot Nothing Then
                    doc.Close(SaveChanges:=False)
                End If

                ' Return the error message (or empty string if ReturnErrorInsteadOfEmpty=False)
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading Word document: {ex.Message}", "")

            Finally
                ' Only quit the application if it was newly created
                If app IsNot Nothing AndAlso app.Visible = False Then
                    app.Quit()
                End If
            End Try
        End Function


        Public Shared Function ReadPdfAsText(ByVal pdfPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Try
                Dim sb As New StringBuilder()

                ' Open the PDF document
                Using document As PdfDocument = PdfDocument.Open(pdfPath)
                    ' Loop through each page in the document
                    For Each page In document.GetPages()
                        ' Extract the text from the page using the GetText() method
                        sb.AppendLine(page.Text) ' Use Text property to extract the text
                    Next
                End Using

                ' Return the extracted text
                Return sb.ToString()
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading PDF: {ex.Message}", "")
            End Try
        End Function


        Public Shared Function EstimateTokenCount(text As String) As Integer
            ' Trim the text and handle edge cases
            If String.IsNullOrWhiteSpace(text) Then Return 0

            ' Estimate tokens: Average of 4 characters per token for English Language
            Dim charCount As Integer = text.Length
            Dim estimatedTokens As Integer = Math.Ceiling(charCount / 4.0)

            Return estimatedTokens
        End Function


    End Class


    Public Class AppConfigurationVariable
        Public Property DisplayName As String         ' z.B. "API Key"
        Public Property VarName As String            ' z.B. "INI_APIKey"
        Public Property VarType As String            ' "String", "Integer" etc.
        Public Property ValidationRule As String     ' z.B. "NotEmpty", "0.0-2.0", "Hyperlink" ...
        Public Property DefaultValue As String       ' z.B. "default-api-key"

        ' Zur Laufzeit gespeicherter Wert, damit bei Wechsel des RadioButtons nichts verloren geht
        Public Property CurrentValue As String


    End Class


    Public Class InitialConfig

        Inherits Form

        Private _context As ISharedContext

        ' Radiobuttons
        Private rbOpenAI As RadioButton
        Private rbAzure As RadioButton
        Private rbGemini As RadioButton
        Private rbVertex As RadioButton

        ' Checkboxen für "Use this configuration for app"
        Private chkWord As System.Windows.Forms.CheckBox
        Private chkOutlook As System.Windows.Forms.CheckBox
        Private chkExcel As System.Windows.Forms.CheckBox

        ' Panels/Controls dynamisch
        Private panelConfig As Panel

        ' Label, das den aktuellen Provider anzeigt, z.B. "Configuration for OpenAI:"
        Private lblCurrentProvider As System.Windows.Forms.Label

        ' Arrays mit Konfigurationen für jede RadioButton-Auswahl
        Private configOpenAI As List(Of AppConfigurationVariable)
        Private configAzure As List(Of AppConfigurationVariable)
        Private configGemini As List(Of AppConfigurationVariable)
        Private configVertex As List(Of AppConfigurationVariable)

        ' Liste von TextBoxen/ComboBoxen/etc. in der aktuellen Anzeige
        Private currentConfigControls As New List(Of Control)

        ' Buttons
        Private btnOK As Button
        Private btnCancel As Button

        Private invisibleLabel As New System.Windows.Forms.Label() With {
                .Size = New System.Drawing.Size(1, 10),
                .Visible = True
            }

        Private Const OverallWidth As Integer = 900

        Private lblUseThisConfig As System.Windows.Forms.Label


        '   Konstruktor – erhält das ISharedContext-Objekt per ByRef

        Public Sub New(ByRef context As ISharedContext)
            _context = context
            Me.Size = New System.Drawing.Size(OverallWidth + 20, 800) ' Gesamtgröße des Fensters
            Me.AutoScroll = False
            Me.AutoSize = True
            Me.InitializeComponent()
            Me.FormBorderStyle = FormBorderStyle.Fixed3D ' Set the form border style to Fixed3D for a 3D effect
        End Sub

        '   Formular erstellen und alle geforderten Controls hinzufügen

        Private Sub InitializeComponent()
            ' Form-Eigenschaften
            Me.Text = $"{SharedMethods.AN} Initial Configuration Wizard"
            Me.FormBorderStyle = FormBorderStyle.None
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = ColorTranslator.FromWin32(&H8000000F)
            Me.ControlBox = False  ' Keine Min/Max/Schließen-Buttons
            Me.AutoScroll = True

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Font = standardFont

            ' PictureBox (Logo)
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            Dim pictureBox As New PictureBox() With {
            .Image = bmp,
            .SizeMode = PictureBoxSizeMode.Zoom
        }
            pictureBox.SetBounds(10, 10, 50, 50)
            Me.Controls.Add(pictureBox)

            ' Label "Welcome to {AN}" neben dem Logo
            Dim lblWelcome As New System.Windows.Forms.Label() With {
                .Text = $"Welcome to {SharedMethods.AN}",
                .AutoSize = True,
                .Font = New System.Drawing.Font("Segoe UI", 12.0F, FontStyle.Bold, GraphicsUnit.Point)
            }

            lblWelcome.Location = New System.Drawing.Point(pictureBox.Right + 10, pictureBox.Top + (pictureBox.Height \ 2) - (lblWelcome.Height \ 2))
            Me.Controls.Add(lblWelcome)

            ' LinkLabel mit dem Text über fehlende .ini
            Dim lblInfo As New LinkLabel() With {
            .AutoSize = True,
            .MaximumSize = New Size(OverallWidth, 0),
            .Text = $"No configuration file '{SharedMethods.AN2}.ini' was found, in which all settings " &
                    "can be made locally or centrally. Therefore, you can make the basic settings here, " &
                    "which will then be saved in such a file. You can then expand this manually. " &
                    $"How this works is explained in the instructions, which you can find at {SharedMethods.AN4}"
        }
            lblInfo.Location = New System.Drawing.Point(10, pictureBox.Bottom + 15)
            AddHandler lblInfo.LinkClicked, AddressOf LinkLabel_LinkClicked
            ' Wir markieren den Link-Bereich (ungefähr) am Ende des Textes
            lblInfo.Links.Add(New LinkLabel.Link() With {
            .LinkData = $"{SharedMethods.AN4}",
            .Start = lblInfo.Text.IndexOf($"{SharedMethods.AN4}", StringComparison.Ordinal),
            .Length = $"{SharedMethods.AN4}".Length
        })
            Me.Controls.Add(lblInfo)

            ' Label + RadioButtons "Which AI provider do you use?"
            Dim lblWhichAI As New System.Windows.Forms.Label() With {
            .Text = "Select your LLM API provider:",
            .AutoSize = True,
            .Font = New System.Drawing.Font(standardFont, FontStyle.Bold)
        }
            lblWhichAI.Location = New System.Drawing.Point(10, lblInfo.Bottom + 20)
            Me.Controls.Add(lblWhichAI)

            rbOpenAI = New RadioButton() With {
            .Text = "OpenAI",
            .AutoSize = True,
            .Checked = True
        }
            rbOpenAI.Location = New System.Drawing.Point(lblWhichAI.Right + 10, lblInfo.Bottom + 18)
            AddHandler rbOpenAI.CheckedChanged, AddressOf RadioButton_CheckedChanged
            Me.Controls.Add(rbOpenAI)

            rbAzure = New RadioButton() With {
            .Text = "Microsoft Azure Open AI Services",
            .AutoSize = True
        }
            rbAzure.Location = New System.Drawing.Point(rbOpenAI.Right + 20, rbOpenAI.Top)
            AddHandler rbAzure.CheckedChanged, AddressOf RadioButton_CheckedChanged
            Me.Controls.Add(rbAzure)

            rbGemini = New RadioButton() With {
            .Text = "Google Gemini",
            .AutoSize = True
        }
            rbGemini.Location = New System.Drawing.Point(rbAzure.Right + 20, rbOpenAI.Top)
            AddHandler rbGemini.CheckedChanged, AddressOf RadioButton_CheckedChanged
            Me.Controls.Add(rbGemini)

            rbVertex = New RadioButton() With {
            .Text = "Google Vertex",
            .AutoSize = True
        }
            rbVertex.Location = New System.Drawing.Point(rbGemini.Right + 20, rbOpenAI.Top)
            AddHandler rbVertex.CheckedChanged, AddressOf RadioButton_CheckedChanged
            Me.Controls.Add(rbVertex)

            ' Zweite LinkLabel-Zeile (darunter)
            Dim lblMoreInfo As New LinkLabel() With {
            .AutoSize = True,
            .MaximumSize = New Size(OverallWidth - 20, 0),
            .Text = $"Note: More on how to obtain access to one of these providers is on {SharedMethods.AN4}. Getting an API access is not expensive. You can use the below form also for other providers. If this does not work or you need to configure more, abort and do it manually before restarting your application."
        }
            lblMoreInfo.Location = New System.Drawing.Point(30, rbOpenAI.Bottom + 5)
            AddHandler lblMoreInfo.LinkClicked, AddressOf LinkLabel_LinkClicked
            lblMoreInfo.Links.Add(New LinkLabel.Link() With {
            .LinkData = $"{SharedMethods.AN4}",
            .Start = lblMoreInfo.Text.IndexOf($"{SharedMethods.AN4}", StringComparison.Ordinal),
            .Length = $"{SharedMethods.AN4}".Length
        })
            Me.Controls.Add(lblMoreInfo)

            ' Label für "Configuration for <AI Provider>:"
            lblCurrentProvider = New System.Windows.Forms.Label() With {
            .AutoSize = True,
            .Font = New System.Drawing.Font(standardFont, FontStyle.Bold),
            .Location = New System.Drawing.Point(10, lblMoreInfo.Bottom + 20)
        }
            Me.Controls.Add(lblCurrentProvider)

            ' Panel, in das wir die dynamischen Eingabefelder platzieren
            panelConfig = New Panel() With {
                .AutoScroll = True,
                .Location = New System.Drawing.Point(10, lblCurrentProvider.Bottom + 5),
                .Width = OverallWidth
            }
            AddHandler panelConfig.SizeChanged, AddressOf PanelConfig_SizeChanged
            Me.Controls.Add(panelConfig)

            ' Erstellt die Konfig-Daten
            PrepareConfigData()

            ' Checkboxen: "Use this configuration for app:"
            lblUseThisConfig = New System.Windows.Forms.Label() With {
                .Text = $"Use this configuration for {SharedMethods.AN}:",
                .Font = New System.Drawing.Font(Me.Font, FontStyle.Bold),
                .AutoSize = True
            }
            lblUseThisConfig.Location = New System.Drawing.Point(10, panelConfig.Bottom + 10)
            Me.Controls.Add(lblUseThisConfig)

            chkWord = New System.Windows.Forms.CheckBox() With {
            .Text = "for Word",
            .AutoSize = True,
            .Checked = _context.RDV.StartsWith("Word")
            }
            chkWord.Location = New System.Drawing.Point(lblUseThisConfig.Right + 10, lblUseThisConfig.Top)
            Me.Controls.Add(chkWord)

            chkOutlook = New System.Windows.Forms.CheckBox() With {
            .Text = "for Outlook (as separate config)",
            .AutoSize = True,
            .Checked = _context.RDV.StartsWith("Outlook")
        }
            chkOutlook.Location = New System.Drawing.Point(chkWord.Right + 17, lblUseThisConfig.Top)
            Me.Controls.Add(chkOutlook)

            chkExcel = New System.Windows.Forms.CheckBox() With {
            .Text = "for Excel (as separate config)",
            .AutoSize = True,
            .Checked = _context.RDV.StartsWith("Excel")
        }
            chkExcel.Location = New System.Drawing.Point(chkOutlook.Right + 17, lblUseThisConfig.Top)
            Me.Controls.Add(chkExcel)

            ' Buttons "OK" und "Cancel"
            btnOK = New Button() With {
            .Text = "OK, save this configuration and continue",
            .AutoSize = True
        }
            btnOK.Location = New System.Drawing.Point(10, lblUseThisConfig.Bottom + 20)
            AddHandler btnOK.Click, AddressOf btnOK_Click
            Me.Controls.Add(btnOK)

            btnCancel = New Button() With {
            .Text = "Cancel",
            .AutoSize = True
        }
            btnCancel.Location = New System.Drawing.Point(btnOK.Right + 10, btnOK.Top)
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            Me.Controls.Add(btnCancel)

            ' Invisible label to ensure margin

            invisibleLabel.Location = New System.Drawing.Point(10, btnCancel.Bottom + 10)
            Me.Controls.Add(invisibleLabel)

            ' Anzeige initial füllen
            LoadConfigForSelectedRadioButton()
        End Sub

        Private Sub PanelConfig_SizeChanged(sender As Object, e As EventArgs)
            Dim panel As Panel = CType(sender, Panel)

            ' Adjust controls below panelConfig dynamically
            lblUseThisConfig.Location = New System.Drawing.Point(10, panel.Bottom + 20)
            chkWord.Location = New System.Drawing.Point(lblUseThisConfig.Right + 10, lblUseThisConfig.Top)
            chkOutlook.Location = New System.Drawing.Point(chkWord.Right + 20, lblUseThisConfig.Top)
            chkExcel.Location = New System.Drawing.Point(chkOutlook.Right + 20, lblUseThisConfig.Top)
            btnOK.Location = New System.Drawing.Point(10, lblUseThisConfig.Bottom + 20)
            btnCancel.Location = New System.Drawing.Point(btnOK.Right + 10, btnOK.Top)
            invisibleLabel.Location = New System.Drawing.Point(10, btnCancel.Bottom + 10)
            Me.Height = invisibleLabel.Bottom + 20

        End Sub

        ' Erzeugt die vier Konfigurationslisten für die AI-Provider

        Private Sub PrepareConfigData()
            ' Wir verwenden dieselben 13 Variablen, können aber bei Bedarf
            ' unterschiedliche Defaults pro Provider setzen.

            configOpenAI = CreateDefaultConfigSet("OpenAI")
            configAzure = CreateDefaultConfigSet("Microsoft Azure Open AI Services")
            configGemini = CreateDefaultConfigSet("Google Gemini")
            configVertex = CreateDefaultConfigSet("Google Vertex")
        End Sub


        ' Erzeugt eine Liste aller 13 Konfigurations-Variablen mit Standardwerten

        Private Function CreateDefaultConfigSet(providerName As String) As List(Of AppConfigurationVariable)
            Dim list As New List(Of AppConfigurationVariable)()
            Select Case providerName
                Case "OpenAI"
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "API Key:",
                        .VarName = "INI_APIKey",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = ""
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Temperature:",
                        .VarName = "INI_Temperature",
                        .VarType = "String",
                        .ValidationRule = "0.0-2.0",
                        .DefaultValue = "0.2"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Timeout (ms):",
                        .VarName = "INI_Timeout",
                        .VarType = "Integer",
                        .ValidationRule = ">0",
                        .DefaultValue = "100000"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Model:",
                        .VarName = "INI_Model",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "gpt-4o"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Endpoint:",
                        .VarName = "INI_Endpoint",
                        .VarType = "String",
                        .ValidationRule = "Hyperlink",
                        .DefaultValue = "https://api.openai.com/v1/chat/completions"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "HeaderA:",
                        .VarName = "INI_HeaderA",
                        .VarType = "String",
                        .ValidationRule = "",
                        .DefaultValue = "Authorization"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "HeaderB:",
                        .ValidationRule = "",
                        .VarName = "INI_HeaderB",
                        .VarType = "String",
                        .DefaultValue = "Bearer {apikey}"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "APICall:",
                        .VarName = "INI_APICall",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "{""model"":   ""{model}"",  ""messages"": [{""role"": ""system"",""content"": ""{promptsystem}""},{""role"": ""user"",""content"": ""{promptuser}""}],""temperature"": {temperature}}"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Response tag:",
                        .VarName = "INI_Response",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "content"
                    })
                Case "Microsoft Azure Open AI Services"
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "API Key:",
                        .VarName = "INI_APIKey",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = ""
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Temperature:",
                        .VarName = "INI_Temperature",
                        .VarType = "String",
                        .ValidationRule = "0.0-2.0",
                        .DefaultValue = "0.2"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Timeout (ms):",
                        .VarName = "INI_Timeout",
                        .VarType = "Integer",
                        .ValidationRule = ">0",
                        .DefaultValue = "100000"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Model:",
                        .VarName = "INI_Model",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "gpt-4o"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Endpoint:",
                        .VarName = "INI_Endpoint",
                        .VarType = "String",
                        .ValidationRule = "Hyperlink",
                        .DefaultValue = "https://[your endpoint]/openai/deployments/[your deployment-id]/chat/completions?api-version=2024-06-01"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "HeaderA:",
                        .VarName = "INI_HeaderA",
                        .VarType = "String",
                        .ValidationRule = "",
                        .DefaultValue = "api-key"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "HeaderB:",
                        .VarName = "INI_HeaderB",
                        .VarType = "String",
                        .ValidationRule = "",
                        .DefaultValue = "{apikey}"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "APICall:",
                        .VarName = "INI_APICall",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "{""messages"": [{""role"": ""system"",""content"": ""{promptsystem}""},{""role"": ""user"", ""content"": ""{promptuser}""}],""temperature"": {temperature}}"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Response tag:",
                        .VarName = "INI_Response",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "content"
                    })
                Case "Google Gemini"
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "API Key:",
                        .VarName = "INI_APIKey",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = ""
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Temperature:",
                        .VarName = "INI_Temperature",
                        .VarType = "String",
                        .ValidationRule = "0.0-2.0",
                        .DefaultValue = "0.2"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Timeout (ms):",
                        .VarName = "INI_Timeout",
                        .VarType = "Integer",
                        .ValidationRule = ">0",
                        .DefaultValue = "100000"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Model:",
                        .VarName = "INI_Model",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "gemini-1.5-pro-latest"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Endpoint:",
                        .VarName = "INI_Endpoint",
                        .VarType = "String",
                        .ValidationRule = "Hyperlink",
                        .DefaultValue = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={apikey}"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "HeaderA:",
                        .VarName = "INI_HeaderA",
                        .VarType = "String",
                        .ValidationRule = "",
                        .DefaultValue = "X-Goog-Api-Key"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "HeaderB:",
                        .VarName = "INI_HeaderB",
                        .VarType = "String",
                        .ValidationRule = "",
                        .DefaultValue = "{apikey}"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "APICall:",
                        .VarName = "INI_APICall",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "{""contents"": [{""role"": ""user"",""parts"": [{ ""text"": ""{promptsystem} {promptuser}"" }]}], ""generationConfig"": {""temperature"": {temperature}}}"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Response tag:",
                        .VarName = "INI_Response",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "text"
                    })

                Case "Google Vertex"
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Private Key (barebones, not PEM):",
                        .VarName = "INI_APIKey",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = ""
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Temperature:",
                        .VarName = "INI_Temperature",
                        .VarType = "String",
                        .ValidationRule = "0.0-2.0",
                        .DefaultValue = "0.2"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Timeout (ms):",
                        .VarName = "INI_Timeout",
                        .VarType = "Integer",
                        .ValidationRule = ">0",
                        .DefaultValue = "100000"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Model:",
                        .VarName = "INI_Model",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "gemini-1.5-pro-002"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Endpoint:",
                        .VarName = "INI_Endpoint",
                        .VarType = "String",
                        .ValidationRule = "Hyperlink",
                        .DefaultValue = "https://europe-west6-aiplatform.googleapis.com/v1/projects/[your project ID]/locations/europe-west6/publishers/google/models/{model}:generateContent"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "HeaderA:",
                        .VarName = "INI_HeaderA",
                        .VarType = "String",
                        .ValidationRule = "",
                        .DefaultValue = "Authorization"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "HeaderB:",
                        .VarName = "INI_HeaderB",
                        .VarType = "String",
                        .ValidationRule = "",
                        .DefaultValue = "Bearer {apikey}"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "APICall:",
                        .VarName = "INI_APICall",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "{""contents"": [{""role"": ""user"", ""parts"":[{""text"": ""{promptsystem} {promptuser}""}]}], ""generationConfig"": {""temperature"": {temperature}}}"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "Response tag:",
                        .VarName = "INI_Response",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "text"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "OAuth2 'client_mail':",
                        .VarName = "INI_OAuth2ClientMail",
                        .VarType = "String",
                        .ValidationRule = "E-Mail",
                        .DefaultValue = "[service account mail]]@[your project ID].iam.gserviceaccount.com"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "OAuth2 'scopes':",
                        .VarName = "INI_OAuth2Scopes",
                        .VarType = "String",
                        .ValidationRule = "NotEmpty",
                        .DefaultValue = "https://www.googleapis.com/auth/cloud-platform"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "OAuth2 Endpoint:",
                        .VarName = "INI_OAuth2Endpoint",
                        .VarType = "String",
                        .ValidationRule = "Hyperlink",
                        .DefaultValue = "https://oauth2.googleapis.com/token"
                    })
                    list.Add(New AppConfigurationVariable With {
                        .DisplayName = "OAuth2 Access Token Expiry (ms):",
                        .VarName = "INI_OAuth2ATExpiry",
                        .VarType = "Integer",
                        .ValidationRule = ">0",
                        .DefaultValue = "3600"
                    })

            End Select
            ' Damit der initiale CurrentValue = DefaultValue ist
            For Each item In list
                item.CurrentValue = item.DefaultValue
            Next

            Return list
        End Function

        '   Wird aufgerufen, wenn sich der Radiobutton ändert.
        '   Speichert Werte aus der Anzeige und lädt neue.
        Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs)
            If Not CType(sender, RadioButton).Checked Then
                Return
            End If

            ' Zuerst aktuelle Werte in das passende Config-Set zurückspeichern
            SaveCurrentInputToConfig()

            ' Dann neues Config-Set laden
            LoadConfigForSelectedRadioButton()
        End Sub



        Private Sub SaveCurrentInputToConfig()
            Dim selectedList = GetSelectedConfigList()

            If selectedList Is Nothing OrElse currentConfigControls.Count = 0 Then
                Exit Sub
            End If

            ' currentConfigControls hat immer "Label" + "TextBox" (oder was auch immer),
            ' also in Zweierschritten durchgehen
            For i As Integer = 0 To currentConfigControls.Count - 1
                Dim ctrl = currentConfigControls(i)
                If TypeOf ctrl Is System.Windows.Forms.Label Then
                    ' DisplayName-Label -> Nächste Control ist die Eingabe
                    Dim labelText = CType(ctrl, System.Windows.Forms.Label).Text
                    ' Finde zugehörige Variable:
                    Dim configVar = selectedList.FirstOrDefault(Function(x) x.DisplayName = labelText)
                    If configVar IsNot Nothing Then
                        ' Nächstes Control:
                        If i + 1 < currentConfigControls.Count Then
                            Dim inputControl = currentConfigControls(i + 1)
                            If TypeOf inputControl Is TextBox Then
                                configVar.CurrentValue = CType(inputControl, TextBox).Text
                            End If
                        End If
                    End If
                End If
            Next
        End Sub


        '   Lädt die Eingabefelder für den aktuell ausgewählten RadioButton neu.
        Private Sub LoadConfigForSelectedRadioButton()
            Dim selectedList = GetSelectedConfigList()
            If selectedList Is Nothing Then
                Return
            End If

            ' Panel leeren
            panelConfig.Controls.Clear()
            currentConfigControls.Clear()

            ' Überschrift anpassen
            Dim providerName As String = ""
            If rbOpenAI.Checked Then providerName = rbOpenAI.Text
            If rbAzure.Checked Then providerName = rbAzure.Text
            If rbGemini.Checked Then providerName = rbGemini.Text
            If rbVertex.Checked Then providerName = rbVertex.Text

            lblCurrentProvider.Text = "Configuration for " & providerName & ":"

            ' Dynamisch alle Felder erstellen:
            Dim yPos As Integer = 0

            ' Determine the maximum width needed for the labels
            Dim maxLabelWidth As Integer = 0
            For Each configVar In selectedList
                Dim lbl As New System.Windows.Forms.Label() With {
                    .Text = configVar.DisplayName,
                    .AutoSize = True,
                    .Font = New System.Drawing.Font(Me.Font, FontStyle.Regular)
                }
                maxLabelWidth = Math.Max(maxLabelWidth, lbl.PreferredWidth)
            Next

            ' Create and position the labels and textboxes
            For Each configVar In selectedList
                ' Label
                Dim lbl As New System.Windows.Forms.Label() With {
                    .Text = configVar.DisplayName,
                    .AutoSize = True,
                    .Font = New System.Drawing.Font(Me.Font, FontStyle.Regular)
                }
                lbl.Location = New System.Drawing.Point(0, yPos)
                panelConfig.Controls.Add(lbl)
                currentConfigControls.Add(lbl)

                ' TextBox
                Dim txt As New TextBox() With {
                    .Width = panelConfig.Width - maxLabelWidth - 30,
                    .Text = configVar.CurrentValue
                }
                txt.Location = New System.Drawing.Point(maxLabelWidth + 10, yPos - 2)
                panelConfig.Controls.Add(txt)
                currentConfigControls.Add(txt)

                yPos += lbl.Height + 8
            Next
            panelConfig.Height = yPos + 2
        End Sub


        '   Gibt die Konfig-Liste zurück, die zum ausgewählten RadioButton passt.
        Private Function GetSelectedConfigList() As List(Of AppConfigurationVariable)
            If rbOpenAI.Checked Then Return configOpenAI
            If rbAzure.Checked Then Return configAzure
            If rbGemini.Checked Then Return configGemini
            If rbVertex.Checked Then Return configVertex
            Return Nothing
        End Function


        '   OK-Button: Validieren, wenn ok -> Werte in context.* speichern und UpdateAppConfig() aufrufen
        Private Sub btnOK_Click(sender As Object, e As EventArgs)
            Try
                ' Eingaben aus dem aktuellen Panel nochmal sichern
                SaveCurrentInputToConfig()

                ' Validierung
                If Not ValidateAllConfigs() Then
                    Return
                End If

                ' Falls Validierung bestanden: Ausgewählte Liste übernehmen
                Dim finalList = GetSelectedConfigList()
                If finalList Is Nothing Then
                    SharedMethods.ShowCustomMessageBox("No AI provider selected.")
                    Return
                End If

                ' Zuweisung an _context
                For Each cv In finalList
                    Select Case cv.VarName
                        Case "INI_APIKey" : _context.INI_APIKey = cv.CurrentValue
                        Case "INI_Temperature" : _context.INI_Temperature = cv.CurrentValue
                        Case "INI_Timeout" : _context.INI_Timeout = CInt(cv.CurrentValue)
                        Case "INI_Model" : _context.INI_Model = cv.CurrentValue
                        Case "INI_Endpoint" : _context.INI_Endpoint = cv.CurrentValue
                        Case "INI_HeaderA" : _context.INI_HeaderA = cv.CurrentValue
                        Case "INI_HeaderB" : _context.INI_HeaderB = cv.CurrentValue
                        Case "INI_APICall" : _context.INI_APICall = cv.CurrentValue
                        Case "INI_Response" : _context.INI_Response = cv.CurrentValue
                        Case "INI_OAuth2ClientMail" : _context.INI_OAuth2ClientMail = cv.CurrentValue
                        Case "INI_OAuth2Scopes" : _context.INI_OAuth2Scopes = cv.CurrentValue
                        Case "INI_OAuth2Endpoint" : _context.INI_OAuth2Endpoint = cv.CurrentValue
                        Case "INI_OAuth2ATExpiry" : _context.INI_OAuth2ATExpiry = CInt(cv.CurrentValue)
                    End Select
                Next

                If rbVertex.Checked Then _context.INI_OAuth2 = True

                _context.INIloaded = False

                Dim providerName As String = String.Empty

                For Each control As Control In Me.Controls
                    If TypeOf control Is RadioButton Then
                        Dim radioButton As RadioButton = CType(control, RadioButton)
                        If radioButton.Checked Then
                            providerName = radioButton.Text
                            Exit For
                        End If
                    End If
                Next

                If chkWord.Checked Then CreateAppConfig("Word", providerName)
                If chkExcel.Checked Then CreateAppConfig("Excel", providerName)
                If chkOutlook.Checked Then CreateAppConfig("Outlook", providerName)

                ' Fenster schließen
                Me.DialogResult = DialogResult.OK
                _context.InitialConfigFailed = False
                Me.Close()

            Catch ex As System.Exception
                MessageBox.Show("Error in LoadConfigForSelectedRadioButton: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub


        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            Me.DialogResult = DialogResult.Cancel
            _context.InitialConfigFailed = True
            Me.Close()
        End Sub


        '   Führt die Validierung aller Felder aller 4 Config-Sets durch oder nur des aktuell selektierten?
        '   In der Anforderung hieß es: "Wird OK gedrückt, dann führe die Gültigkeitsprüfung der Werte durch..."
        '   -> Wir validieren NUR das aktuell aktive Set.
        Private Function ValidateAllConfigs() As Boolean

            Dim selectedList = GetSelectedConfigList()

            ' Check if at least one relevant checkbox is checked
            If _context.RDV.StartsWith("Word") AndAlso Not chkWord.Checked Then
                SharedMethods.ShowCustomMessageBox("At least the 'for Word' checkbox needs to be checked.")
                Return False
            ElseIf _context.RDV.StartsWith("Outlook") AndAlso Not chkOutlook.Checked Then
                SharedMethods.ShowCustomMessageBox("At least the 'for Outlook' checkbox needs to be checked.")
                Return False
            ElseIf _context.RDV.StartsWith("Excel") AndAlso Not chkExcel.Checked Then
                SharedMethods.ShowCustomMessageBox("At least the 'for Excel' checkbox needs to be checked.")
                Return False
            End If

            For Each cv In selectedList
                Dim valRule = cv.ValidationRule
                Dim valValue = cv.CurrentValue

                Debug.WriteLine("Validating: valrule=" & valRule & ", valValue='" & valValue & "'")

                ' NotEmpty
                If valRule.Contains("NotEmpty") Then
                    If String.IsNullOrWhiteSpace(valValue) Then
                        SharedMethods.ShowCustomMessageBox("Value For '" & cv.DisplayName & "' cannot be empty.")
                        Return False
                    End If
                End If

                ' E-Mail
                If valRule.Contains("E-Mail") Then
                    ' Minimale Plausibilitätsprüfung
                    If Not valValue.Contains("@") Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a valid e-mail address.")
                        Return False
                    End If
                End If

                ' Hyperlink
                If valRule.Contains("Hyperlink") Then
                    If Not (valValue.StartsWith("http://") OrElse valValue.StartsWith("https://")) Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a valid URL (http/https).")
                        Return False
                    End If
                End If

                ' >0
                If valRule.Contains(">0") Then
                    Dim intVal As Integer
                    If Not Integer.TryParse(valValue, intVal) OrElse intVal <= 0 Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be an integer larger than 0.")
                        Return False
                    End If
                End If

                ' 0.0-2.0
                If valRule.Contains("0.0-2.0") Then
                    Dim dblVal As Double
                    If Not Double.TryParse(valValue, dblVal) Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a floating number between 0.0 and 2.0.")
                        Return False
                    End If
                    If dblVal < 0.0 OrElse dblVal > 2.0 Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be in [0.0 .. 2.0].", "Validation Error")
                        Return False
                    End If
                End If
            Next

            Return True
        End Function


        '   Öffnet Links im Browser
        Private Sub LinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
            Try
                Dim link = e.Link.LinkData.ToString()
                System.Diagnostics.Process.Start(link)
            Catch ex As System.Exception
                MessageBox.Show("Could not open link. Error: " & ex.Message)
            End Try
        End Sub

        Private Sub CreateAppConfig(App As String, provider As String)
            Try
                ' Define the file path
                Dim filepath = SharedMethods.GetDefaultINIPath(App)

                Debug.WriteLine($"Creating {SharedMethods.AN} configuration file: " & filepath)

                ' Open a StreamWriter to create the file
                Using writer As New System.IO.StreamWriter(filepath)
                    ' Write the header
                    writer.WriteLine($"; {SharedMethods.AN} configuration file (automatically generated)")
                    writer.WriteLine(";")
                    writer.WriteLine($"; Go to {SharedMethods.AN4} on how to find the instructions to manually add or change the configuration settings")

                    ' Write an empty line
                    writer.WriteLine()

                    ' Write provider information
                    writer.WriteLine($"; Minimum configuration for {provider}")

                    ' Write another empty line
                    writer.WriteLine()

                    ' Loop through the dictionary and write each configuration value
                    Dim MinimumConfigValues As New Dictionary(Of String, String) From {
                            {"APIKey", _context.INI_APIKey},
                            {"Endpoint", _context.INI_Endpoint},
                            {"HeaderA", _context.INI_HeaderA},
                            {"HeaderB", _context.INI_HeaderB},
                            {"Response", _context.INI_Response},
                            {"APICall", _context.INI_APICall},
                            {"Timeout", _context.INI_Timeout.ToString()},
                            {"Temperature", _context.INI_Temperature},
                            {"Model", _context.INI_Model},
                            {"OAuth2", _context.INI_OAuth2.ToString()},
                            {"OAuth2ClientMail", _context.INI_OAuth2ClientMail},
                            {"OAuth2Scopes", _context.INI_OAuth2Scopes},
                            {"OAuth2Endpoint", _context.INI_OAuth2Endpoint},
                            {"OAuth2ATExpiry", _context.INI_OAuth2ATExpiry.ToString()}
                        }

                    For Each kvp In MinimumConfigValues
                        writer.WriteLine($"{kvp.Key} = {kvp.Value}")
                    Next
                End Using

            Catch ex As System.Exception
                ' Handle errors by showing a custom message box
                SharedMethods.ShowCustomMessageBox($"Error creating configuration file: {ex.Message}")
            End Try
        End Sub



    End Class


    Public Class UpdateHandler

        Public Sub CheckAndInstallUpdates(appname As String, LocalPath As String)
            Try
                ' Ensure the application is ClickOnce deployed
                If ApplicationDeployment.IsNetworkDeployed And String.IsNullOrWhiteSpace(LocalPath) Then
                    Dim deployment As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
                    Dim currentDate As Date = Date.Now

                    ' Check for updates
                    If deployment.CheckForUpdate() Then
                        Dim dialogResult As Integer = SharedMethods.ShowCustomYesNoBox($"An update is available online ({deployment.UpdateLocation.AbsoluteUri}). Do you want to download and install it now? The update will take effect the next time you restart the application. Note: If you run this within a corporate environment, your firewall may block this.", "Yes", "No")

                        If dialogResult = 1 Then
                            ' Download and apply the update
                            deployment.Update()

                            If dialogResult = 1 Or dialogResult = 2 Then
                                ' Update the last check time
                                Select Case Left(appname, 4)
                                    Case "Word"
                                        My.Settings.LastUpdateCheckWord = currentDate
                                    Case "Exce"
                                        My.Settings.LastUpdateCheckExcel = currentDate
                                    Case "Outl"
                                        My.Settings.LastUpdateCheckOutlook = currentDate
                                End Select
                                My.Settings.Save()
                            End If

                            ' Notify the user
                            SharedMethods.ShowCustomMessageBox("The update process has been performed. Restart the application to see whether it was successul.", $"{SharedMethods.AN} Updater")
                        End If
                    Else
                        SharedMethods.ShowCustomMessageBox($"No updates are currently available ({deployment.UpdateLocation.AbsoluteUri}).", $"{SharedMethods.AN} Updater")
                    End If

                    Select Case Left(appname, 4)
                        Case "Word"
                            My.Settings.LastUpdateCheckWord = currentDate
                        Case "Exce"
                            My.Settings.LastUpdateCheckExcel = currentDate
                        Case "Outl"
                            My.Settings.LastUpdateCheckOutlook = currentDate
                    End Select
                    My.Settings.Save()
                Else
                    If LocalPath = "" Then
                        SharedMethods.ShowCustomMessageBox($"This version of {SharedMethods.AN} has not been configured with an update path ('UpdatedPath = '). The configuration should refer to the main directory where the installation sources 'word', 'excel' and 'outlook' are stored. You may have to discuss this with your administrator.", $"{SharedMethods.AN} Updater")
                    Else
                        LocalPath = SharedMethods.ExpandEnvironmentVariables(LocalPath)
                        Dim dialogResult As Integer = SharedMethods.ShowCustomYesNoBox($"This will initiate the installer for this add-in. If there is a new version at '{LocalPath}', it will be installed. Do you want to proceed?", "Yes", "No")
                        If dialogResult = 1 Then
                            Dim vstoFilePath As String = ""
                            Debug.WriteLine(appname)
                            Select Case Left(appname, 4)
                                Case "Word"
                                    vstoFilePath = System.IO.Path.Combine(LocalPath, $"word\{SharedMethods.AN3} for Word.vsto")
                                Case "Exce"
                                    vstoFilePath = System.IO.Path.Combine(LocalPath, $"excel\{SharedMethods.AN3} for Excel.vsto")
                                Case "Outl"
                                    vstoFilePath = System.IO.Path.Combine(LocalPath, $"outlook\{SharedMethods.AN3} for Outlook.vsto")
                            End Select

                            If System.IO.File.Exists(vstoFilePath) Then
                                Process.Start(vstoFilePath)
                                SharedMethods.ShowCustomMessageBox("The update process has been performed. Restart the application to see whether it was successul.", $"{SharedMethods.AN} Updater")
                            Else
                                SharedMethods.ShowCustomMessageBox($"Installer '{vstoFilePath}' not found. Check 'UpdatePath =' in the '{SharedMethods.AN2}.ini''.", $"{SharedMethods.AN} Updater")
                            End If
                        End If
                    End If
                End If
            Catch ex As DeploymentException
                ' Handle exceptions related to update checking and applying
                SharedMethods.ShowCustomMessageBox("An error occurred while checking for or installing updates: " & ex.Message, $"{SharedMethods.AN} Updater")
            End Try
        End Sub


        Public Shared Sub PeriodicCheckForUpdates(checkIntervalInDays As Integer, appname As String, LocalPath As String)
            Try
                ' Get the last update check time from settings

                If checkIntervalInDays = 0 Then Exit Sub

                Dim lastCheck As Date

                Select Case Left(appname, 4)
                    Case "Word"
                        lastCheck = My.Settings.LastUpdateCheckWord
                    Case "Exce"
                        lastCheck = My.Settings.LastUpdateCheckExcel
                    Case "Outl"
                        lastCheck = My.Settings.LastUpdateCheckOutlook
                    Case Else
                        Exit Sub
                End Select

                Dim currentDate As Date = Date.Now

                ' Calculate the number of days elapsed since the last check
                Dim elapsedDays As Double = (currentDate - lastCheck).TotalDays

                ' Check for updates if the interval has passed
                If elapsedDays >= checkIntervalInDays Or checkIntervalInDays < 0 Then
                    ' Ensure the application is ClickOnce deployed

                    If ApplicationDeployment.IsNetworkDeployed And String.IsNullOrWhiteSpace(LocalPath) Then
                        Dim deployment As ApplicationDeployment = ApplicationDeployment.CurrentDeployment

                        Dim Dialogresult As Integer = 0

                        ' Check if an update is available
                        If deployment.CheckForUpdate() Then
                            Dialogresult = SharedMethods.ShowCustomYesNoBox("An update is available online. Do you want to download and install it now? Note: If you run this within a corporate environment, your firewall may block this.", "Yes", If(checkIntervalInDays < 0, "No (configured to check on next startup)", "No, check again in " & checkIntervalInDays & " days"))

                            If Dialogresult = 1 Then
                                ' Download and apply the update
                                deployment.Update()

                                ' Notify the user to restart
                                If Dialogresult = 1 Or Dialogresult = 2 Then
                                    ' Update the last check time
                                    Select Case Left(appname, 4)
                                        Case "Word"
                                            My.Settings.LastUpdateCheckWord = currentDate
                                        Case "Exce"
                                            My.Settings.LastUpdateCheckExcel = currentDate
                                        Case "Outl"
                                            My.Settings.LastUpdateCheckOutlook = currentDate
                                    End Select
                                    My.Settings.Save()
                                End If

                                SharedMethods.ShowCustomMessageBox("The update process has been performed. Restart the application to see whether it was successul.", $"{SharedMethods.AN} Updater")
                            End If
                        End If
                        If Dialogresult = 1 Or Dialogresult = 2 Then
                            ' Update the last check time
                            Select Case Left(appname, 4)
                                Case "Word"
                                    My.Settings.LastUpdateCheckWord = currentDate
                                Case "Exce"
                                    My.Settings.LastUpdateCheckExcel = currentDate
                                Case "Outl"
                                    My.Settings.LastUpdateCheckOutlook = currentDate
                            End Select
                            My.Settings.Save()
                        End If
                    ElseIf Not String.IsNullOrWhiteSpace(LocalPath) Then
                        LocalPath = SharedMethods.ExpandEnvironmentVariables(LocalPath)
                        Dim dialogResult As Integer = SharedMethods.ShowCustomYesNoBox($"Do you want to check for updates? If yes, the installer for this add-in will run. If there is a new version at '{LocalPath}', it will be installed. Do you want to proceed?", "Yes", If(checkIntervalInDays < 0, "No (configured to check on next startup)", "No, check again in " & checkIntervalInDays & " days"))

                        If dialogResult = 1 Then
                            Dim vstoFilePath As String = ""
                            Debug.WriteLine(appname)
                            Select Case Left(appname, 4)
                                Case "Word"
                                    vstoFilePath = System.IO.Path.Combine(LocalPath, $"word\{SharedMethods.AN3} for Word.vsto")
                                Case "Exce"
                                    vstoFilePath = System.IO.Path.Combine(LocalPath, $"excel\{SharedMethods.AN3} for Excel.vsto")
                                Case "Outl"
                                    vstoFilePath = System.IO.Path.Combine(LocalPath, $"outlook\{SharedMethods.AN3} for Outlook.vsto")
                            End Select

                            If System.IO.File.Exists(vstoFilePath) Then
                                Process.Start(vstoFilePath)
                                SharedMethods.ShowCustomMessageBox("The update process has been performed. Restart the application to see whether it was successul.", $"{SharedMethods.AN} Updater")
                            Else
                                SharedMethods.ShowCustomMessageBox($"Installer '{vstoFilePath}' not found. Check 'UpdatePath =' in the '{SharedMethods.AN2}.ini''.", $"{SharedMethods.AN} Updater")
                            End If
                        End If
                        If dialogResult = 1 Or dialogResult = 2 Then
                            ' Update the last check time
                            Select Case Left(appname, 4)
                                Case "Word"
                                    My.Settings.LastUpdateCheckWord = currentDate
                                Case "Exce"
                                    My.Settings.LastUpdateCheckExcel = currentDate
                                Case "Outl"
                                    My.Settings.LastUpdateCheckOutlook = currentDate
                            End Select
                            My.Settings.Save()
                        End If

                    End If
                End If
            Catch ex As DeploymentException
                ' Handle exceptions related to update checking and applying
                SharedMethods.ShowCustomMessageBox("An error occurred while checking for or installing updates: " & ex.Message, $"{SharedMethods.AN} Updater")
            End Try
        End Sub


    End Class


End Namespace

