' Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See License.txt or https://vischer.com/redink for more information.
'
' 6.8.2025
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
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes Grpc.core in unchanged form; Copyright (c) 2023 The gRPC Authors; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/grpc/grpc
' Includes Google Speech V1 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes Google Protobuf in unchanged form; Copyright (c) 2025 Google Inc.; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/protocolbuffers/protobuf
' Includes MarkdownToRTF in modified form; Copyright (c) 2025 Gustavo Hennig; original licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/GustavoHennig/MarkdownToRtf
' Includes Nito.AsyncEx in unchanged form; Copyright (c) 2021 Stephen Cleary; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/StephenCleary/AsyncEx
' Includes NetOffice libraries in unchanged form; Copyright (c) 2020 Sebastian Lange, Erika LeBlanc; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/netoffice/NetOffice-NuGet
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.

Option Explicit On

Imports System.Collections.Concurrent
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Linq.Expressions
Imports System.Net
Imports System.Net.Http
Imports System.Net.WebSockets
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Security.Policy
Imports System.Speech.Synthesis
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Validation
Imports DocumentFormat.OpenXml.Wordprocessing
Imports DocumentFormat.OpenXml.Office2010.PowerPoint
Imports DocumentFormat.OpenXml.Office2010.Drawing
Imports Google.Api.Gax.Grpc
Imports Google.Apis.Auth.OAuth2.Responses
Imports Google.Cloud.Speech.V1
Imports Google.Cloud.Speech.V1.LanguageCodes
Imports Google.Protobuf
Imports Google.Rpc.Context.AttributeContext.Types
Imports Grpc.Core
Imports HtmlAgilityPack
Imports Markdig
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NAudio
Imports NAudio.CoreAudioApi
Imports NAudio.Wave
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.MarkdownToRtf
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports VBScript_RegExp_55
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




Public Module WordSearchHelper

    Private deletionsCache As System.Collections.Generic.List(Of (Integer, Integer)) = Nothing

    Private ReadOnly RX_U As New _
    System.Text.RegularExpressions.Regex("\\u([0-9A-Fa-f]{4,6})",
        System.Text.RegularExpressions.RegexOptions.Compiled Or
        System.Text.RegularExpressions.RegexOptions.CultureInvariant)

    Private ReadOnly RX_ELL As New _
    System.Text.RegularExpressions.Regex("\.\.\.",
        System.Text.RegularExpressions.RegexOptions.Compiled Or
        System.Text.RegularExpressions.RegexOptions.CultureInvariant)



    Public Function FindLongTextAnchoredFast(
    ByRef sel As Microsoft.Office.Interop.Word.Selection,
    ByVal findText As String,
    Optional ByVal skipDeleted As Boolean = False,
    Optional ByVal nWords As Integer = 4,
    Optional ByVal cancel As System.Threading.CancellationToken = Nothing,
    Optional ByVal timeoutSeconds As Integer = 10) As Boolean

        Debug.WriteLine("Skipdeleted=" & skipDeleted)

        Dim _dbgLastSlice As String = ""
        Dim _dbgLastIdx As Integer = -1
        Dim _dbgNeedle As String = ""

        Dim t0 As System.DateTime = System.DateTime.UtcNow
        Dim doc As Microsoft.Office.Interop.Word.Document = sel.Document
        Dim area As Microsoft.Office.Interop.Word.Range =
        If(sel.Range.Start = sel.Range.End, doc.Content.Duplicate, sel.Range.Duplicate)

        '──────── 0) REINE Ctrl-F-Suche (Literal, kein Format) ────────────
        If findText.Length <= 255 Then
            Dim rngPlain As Microsoft.Office.Interop.Word.Range = area.Duplicate
            With rngPlain.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Font.Reset() : .ParagraphFormat.Reset()
                .Text = findText
                .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .MatchCase = False : .MatchWholeWord = False
                .MatchWildcards = False : .Format = False
                .IgnoreSpace = False          'exakt wie Ctrl-F
            End With
            Dim hitPlain As Boolean
            Try : hitPlain = rngPlain.Find.Execute()
            Catch : hitPlain = False         'COM-Fehler (z. B. >255) → weiter
            End Try
            If hitPlain Then
                sel.SetRange(rngPlain.Start, rngPlain.End)
                Return True
            End If
        End If

        '──────── 1) QUICK-Literal  (masked, IgnoreSpace=True, ≤255) ──────
        Dim litPat As String = EscapeForWordWildcard(findText)
        If litPat.Length <= 255 Then
            Dim rngLit As Microsoft.Office.Interop.Word.Range = area.Duplicate
            With rngLit.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Font.Reset() : .ParagraphFormat.Reset()
                .Text = litPat
                .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .MatchCase = False : .MatchWildcards = False
                .Format = False : .IgnoreSpace = True
            End With
            If rngLit.Find.Execute() Then
                sel.SetRange(rngLit.Start, rngLit.End) : Return True
            End If
        End If

        '──────── 2) Start/End-Vorbereitung  ──────────────────────────────
        Dim raw() As String = WordSearchHelper.RawWords(findText)
        Dim canonNeedle As String = Canonicalise(findText, True)
        If canonNeedle = "" Then Return False

        '---- NEU: Kompletter Wildcard-Suchlauf, falls < 255 ----
        Dim fullWildcardPattern As String = BuildWildcardProbe(raw)
        If fullWildcardPattern.Length <= 255 Then
            Dim rngFull As Microsoft.Office.Interop.Word.Range = area.Duplicate
            With rngFull.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Font.Reset() : .ParagraphFormat.Reset()
                .Text = fullWildcardPattern
                .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .MatchCase = False : .MatchWildcards = True
                .Format = False : .IgnoreSpace = True
            End With
            If rngFull.Find.Execute() Then
                sel.SetRange(rngFull.Start, rngFull.End)
                Return True
            End If
        End If

        ' Fallback auf Anker-Logik, falls kompletter Suchlauf nicht möglich war
        If raw.Length < 2 Then Return False ' Anker-Logik braucht min. 2 Wörter

        ' nWords so wählen, dass Start- und End-Anker nicht überlappen
        nWords = System.Math.Min(nWords, raw.Length \ 2)
        If nWords < 1 Then nWords = 1

        ' Sicherstellen, dass der Start-Anker nicht zu lang ist
        Do While nWords > 1 AndAlso BuildWildcardProbe(raw.Take(nWords).ToArray()).Length > 255
            nWords -= 1
        Loop

        Dim startPat As String = BuildWildcardProbe(raw.Take(nWords).ToArray())
        Dim endWords() As String = raw.Skip(raw.Length - nWords).ToArray()
        Dim endPat As String = BuildWildcardProbe(endWords)

        Dim occur As Integer = CountOccurrences(findText, System.String.Join(" "c, endWords))
        If startPat = endPat Then occur = System.Math.Max(2, occur)

        deletionsCache = Nothing

        '──────── 3) Start-Wildcard-Find (formatneutral) ──────────────────
        Using sRng As New RangeProxy(area.Duplicate)
            Dim fS As Microsoft.Office.Interop.Word.Find = sRng.Range.Find
            With fS
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Font.Reset() : .ParagraphFormat.Reset()
                .Text = startPat
                .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .MatchCase = False : .MatchWildcards = True
                .Format = False : .IgnoreSpace = True
            End With

            Dim okS As Boolean : Try : okS = fS.Execute() : Catch : okS = False : End Try
            While okS
                If (System.DateTime.UtcNow - t0).TotalSeconds > timeoutSeconds Then
                    Throw New System.Exception("Timeout while searching.")
                End If
                cancel.ThrowIfCancellationRequested()

                Dim posStart As Integer = sRng.Range.Start
                Dim searchFrom As Integer = sRng.Range.End

                '──────── End-Wildcard-Find ───────────────────────────────
                Dim eRng As Microsoft.Office.Interop.Word.Range = doc.Range(searchFrom, area.End)
                Dim fE As Microsoft.Office.Interop.Word.Find = eRng.Find
                With fE
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Font.Reset() : .ParagraphFormat.Reset()
                    .Text = endPat
                    .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .MatchCase = False : .MatchWildcards = True
                    .Format = False : .IgnoreSpace = True
                End With
                Dim okE As Boolean : Try : okE = fE.Execute() : Catch : okE = False : End Try
                For i As Integer = 2 To occur
                    If Not okE Then Exit For
                    eRng.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    Try : okE = fE.Execute() : Catch : okE = False : End Try
                Next

                If okE Then
                    Dim sliceTxt As String, back As System.Collections.Generic.IReadOnlyList(Of Integer)
                    VisibleSlice(doc, posStart, eRng.End - posStart, skipDeleted, sliceTxt, back)

                    Debug.WriteLine(sliceTxt & vbCrLf)

                    Dim canSlice As String = Canonicalise(sliceTxt, True)
                    Dim idx As Integer = canSlice.IndexOf(canonNeedle, System.StringComparison.Ordinal)

                    _dbgLastSlice = canSlice
                    _dbgLastIdx = idx
                    _dbgNeedle = canonNeedle

                    If idx >= 0 Then
                        sel.SetRange(back(idx), back(idx + canonNeedle.Length - 1) + 1)
                        Return True
                    End If
                End If

                sRng.CollapseEndPlusOne()
                If sRng.Range.Start >= area.End Then Exit While
                Try : okS = fS.Execute() : Catch : okS = False : End Try
            End While
        End Using

        Dim elapsedSec As Double = (System.DateTime.UtcNow - t0).TotalSeconds

        System.Diagnostics.Debug.WriteLine("===== FindLongTextAnchoredFast: FINAL DEBUG =====")
        System.Diagnostics.Debug.WriteLine("  findText        = '" & findText & "'")
        System.Diagnostics.Debug.WriteLine("  lastIdx         = " & _dbgLastIdx)
        System.Diagnostics.Debug.WriteLine("  needle.Length   = " & _dbgNeedle.Length)
        System.Diagnostics.Debug.WriteLine("  slice.Length    = " & _dbgLastSlice.Length)
        System.Diagnostics.Debug.WriteLine("  contains?       = " & _dbgLastSlice.Contains(_dbgNeedle).ToString())
        ' show first and last 200 chars of the slice
        Dim previewLen As Integer = 200
        Dim startEx As String = If(_dbgLastSlice.Length <= previewLen,
                                _dbgLastSlice,
                                _dbgLastSlice.Substring(0, previewLen) & "…")
        Dim endEx As String = If(_dbgLastSlice.Length <= previewLen,
                                "",
                                "…" & _dbgLastSlice.Substring(_dbgLastSlice.Length - previewLen))
        System.Diagnostics.Debug.WriteLine("  slice excerpt start: '" & startEx & "'")
        If endEx <> "" Then
            System.Diagnostics.Debug.WriteLine("  slice excerpt end:   '" & endEx & "'")
        End If
        System.Diagnostics.Debug.WriteLine("===============================================")

        Return False
    End Function


    Private Function BuildWildcardProbe(ByVal words() As String) As String
        Dim sb As New System.Text.StringBuilder(words.Length * 14)
        Dim i As Integer = 0
        While i < words.Length
            If i > 0 Then sb.Append(" "c)

            Dim w As String = words(i)

            '— komplette Placeholder-Sequenz —
            If w.Contains("["c) Then
                'bis zum Wort mit schließendem ]
                While i < words.Length AndAlso Not words(i).Contains("]"c)
                    i += 1
                End While
                'Literal „\[“ + beliebig Zeichen + „\]“ (+ evtl. Rest)
                sb.Append("\[*\]")
                If i < words.Length Then
                    Dim rest As String = words(i).Substring(words(i).IndexOf("]"c) + 1)
                    If rest <> "" Then sb.Append(EscapeForWordWildcard(rest))
                End If
            Else
                'Hyphen-Familie → ?
                w = w.Replace("-"c, "?"c) _
                 .Replace(ChrW(&H2010), "?"c) _
                 .Replace(ChrW(&H2011), "?"c) _
                 .Replace(ChrW(&H2013), "?"c) _
                 .Replace(ChrW(&H2014), "?"c) _
                 .Replace(ChrW(&HAD), "?"c)
                sb.Append(EscapeForWordWildcard(w))
            End If
            i += 1
        End While
        Return sb.ToString()
    End Function


    'Zählt (kanonisiert, case-insensitiv) wie oft subTxt in txt vorkommt
    Private Function CountOccurrences(ByVal txt As String, ByVal subTxt As String) As Integer
        txt = Canonicalise(txt, True)
        subTxt = Canonicalise(subTxt, True)
        Dim cnt As Integer = 0
        Dim pos As Integer = txt.IndexOf(subTxt, System.StringComparison.OrdinalIgnoreCase)
        While pos <> -1
            cnt += 1
            pos = txt.IndexOf(subTxt, pos + subTxt.Length, System.StringComparison.OrdinalIgnoreCase)
        End While
        Return cnt
    End Function

    'RawWords  –  nur \u-Escape → Zeichen, Ellipsis → U+2026, Normalisierung
    Private Function RawWords(ByVal src As String) As String()
        src = RX_U.Replace(src, Function(m) _
        System.Char.ConvertFromUtf32(System.Convert.ToInt32(m.Groups(1).Value, 16)))
        src = RX_ELL.Replace(src, ChrW(&H2026))
        src = src.Normalize(System.Text.NormalizationForm.FormKC)
        Return src.Split(New Char() {" "c, ChrW(9), ChrW(10), ChrW(13)},
                     System.StringSplitOptions.RemoveEmptyEntries)
    End Function

    Private Sub VisibleSlice(
    ByVal doc As Microsoft.Office.Interop.Word.Document,
    ByVal absStart As Integer,
    ByVal sliceLen As Integer,
    ByVal skipDeleted As Boolean,
    ByRef visOut As String,
    ByRef mapBack As System.Collections.Generic.IReadOnlyList(Of Integer))

        '---- Slice-Grenzen ------------------------------------------------
        Dim rawEnd As Integer =
        System.Math.Min(doc.Content.End, absStart + sliceLen + 500)
        Dim rawRng As Microsoft.Office.Interop.Word.Range = doc.Range(absStart, rawEnd)
        Dim rawTxt As String = rawRng.Text

        If Not skipDeleted Then
            ' auf sliceLen kürzen
            Dim outTxt As String = If(rawTxt.Length > sliceLen, rawTxt.Substring(0, sliceLen), rawTxt)
            visOut = outTxt
            ' Back-Map 1:1 aufbauen
            Dim backList = New System.Collections.Generic.List(Of Integer)(outTxt.Length)
            For i As Integer = 0 To outTxt.Length - 1
                backList.Add(absStart + i)
            Next
            mapBack = backList
            Return
        End If

        '---- Deletions-Cache einmalig aufbauen ----------------------------
        If deletionsCache Is Nothing Then
            deletionsCache = New System.Collections.Generic.List(Of (Integer, Integer))()
            For Each rev As Microsoft.Office.Interop.Word.Revision In doc.Revisions
                If rev.Type =
               Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionInsert _
               OrElse rev.Type =
               Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionMovedTo Then _
               Continue For
                deletionsCache.Add((rev.Range.Start, rev.Range.End))
            Next
            deletionsCache.Sort(Function(a, b) a.Item1.CompareTo(b.Item1))
        End If

        '---- Nur Intervalle aufnehmen, die in unseren Slice fallen --------
        Dim intervals As New System.Collections.Generic.List(Of (Integer, Integer))()
        For Each iv In deletionsCache
            If iv.Item2 <= absStart Then Continue For   'liegt davor
            If iv.Item1 >= rawEnd Then Exit For         'liegt dahinter (Liste ist sortiert)
            Dim sRel As Integer = System.Math.Max(iv.Item1, absStart) - absStart
            Dim eRel As Integer = System.Math.Min(iv.Item2, rawEnd) - absStart
            intervals.Add((sRel, eRel))
        Next

        '---- Überlappungen mergen ----------------------------------------
        Dim merged As New System.Collections.Generic.List(Of (Integer, Integer))()
        For Each iv In intervals
            If merged.Count = 0 OrElse iv.Item1 > merged(merged.Count - 1).Item2 Then
                merged.Add(iv)
            Else
                Dim last = merged(merged.Count - 1)
                merged(merged.Count - 1) =
                (last.Item1, System.Math.Max(last.Item2, iv.Item2))
            End If
        Next

        '---- Sichtbaren Text + Back-Map bauen ----------------------------
        Dim sb As New System.Text.StringBuilder(rawTxt.Length)
        Dim back As New System.Collections.Generic.List(Of Integer)(rawTxt.Length)
        Dim pos As Integer = 0

        For Each iv In merged
            If iv.Item1 > pos Then
                sb.Append(rawTxt, pos, iv.Item1 - pos)
                For i As Integer = pos To iv.Item1 - 1
                    back.Add(absStart + i)
                Next
            End If
            pos = iv.Item2
        Next
        If pos < rawTxt.Length Then
            sb.Append(rawTxt, pos, rawTxt.Length - pos)
            For i As Integer = pos To rawTxt.Length - 1
                back.Add(absStart + i)
            Next
        End If

        '---- Auf gewünschte Länge kürzen ---------------------------------
        If sb.Length > sliceLen Then
            sb.Length = sliceLen
            back = back.GetRange(0, sliceLen)
        End If

        visOut = sb.ToString()
        mapBack = back
    End Sub


    '══════════════════════════════════════════════════════════════════════
    '  Normalisierungen
    '══════════════════════════════════════════════════════════════════════
    Private Function PrepareNeedle(ByVal txt As String) As String
        txt = RX_U.Replace(txt,
        Function(m) System.Char.ConvertFromUtf32(
            System.Convert.ToInt32(m.Groups(1).Value, 16)))
        Return RX_ELL.Replace(txt, ChrW(&H2026))
    End Function

    Private Function Canonicalise(ByVal src As String, ByVal collapseWS As Boolean) As String
        src = PrepareNeedle(src).Normalize(System.Text.NormalizationForm.FormKC)

        Dim sb As New System.Text.StringBuilder(src.Length)
        Dim pendingSpace As Boolean = False

        For Each ch As Char In src
            If IsDocNoise(ch) Then Continue For

            Dim code As Integer = AscW(ch)
            Dim isHyphenOrWs As Boolean = System.Char.IsWhiteSpace(ch) OrElse code = &HA0

            If Not isHyphenOrWs Then
                Select Case code
                    Case &H2010, &H2011, &H2013, &H2014, &HAD, 45 ' Hyphen family
                        isHyphenOrWs = True
                End Select
            End If

            If isHyphenOrWs Then
                pendingSpace = True
            Else
                If pendingSpace AndAlso collapseWS Then sb.Append(" "c)
                pendingSpace = False
                sb.Append(CanonizeDocChar(ch))
            End If
        Next
        Return sb.ToString().Trim()
    End Function


    Private Function IsDocNoise(ByVal ch As Char) As Boolean
        Dim code As Integer = AscW(ch)
        If code < 32 AndAlso code <> 9 AndAlso code <> 10 AndAlso code <> 13 Then Return True
        Select Case code
            Case &HA0, &H200B, &H200C, &H200D, &H2060,
             &H200E To &H200F, &H202A To &H202E,
             1, 19, 20, 21, &HFFFA, &HFFFB, &HFFFC
                Return True
        End Select
        Return False
    End Function

    Private Function CanonizeDocChar(ByVal ch As Char) As String
        Select Case AscW(ch)
            Case &HDF, &H1E9E : Return "SS"  ' ß / ẞ
            Case Else
                Return System.Char.ToUpperInvariant(ch)
        End Select
    End Function

    'Maskiert alle Word-Wildcard-Sonderzeichen, damit sie literale Bedeutung behalten
    Private Function EscapeForWordWildcard(ByVal s As String) As String
        If s = "" Then Return ""
        Dim sb As New System.Text.StringBuilder(s.Length * 2)
        For Each ch As Char In s
            Select Case ch
                Case "?"c, "*"c, "@"c, "["c, "]"c, "("c, ")"c,
                 "{"c, "}"c, "\"c, "<"c, ">"c
                    sb.Append("\"c)      '\  ist das Literal-Escape in Word
            End Select
            sb.Append(ch)
        Next
        Return sb.ToString()
    End Function

    Private NotInheritable Class RangeProxy
        Implements System.IDisposable

        Friend ReadOnly Range As Microsoft.Office.Interop.Word.Range
        Private ReadOnly ptr As Object

        Friend Sub New(ByVal r As Microsoft.Office.Interop.Word.Range)
            Range = r
            ptr = r                    'COM-Pointer merken
        End Sub

        Friend Sub CollapseEndPlusOne()
            Range.Collapse(
            Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
            Range.SetRange(Range.Start + 1, Range.Start + 1)
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If ptr IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ptr)
            End If
        End Sub
    End Class

End Module




Public Class ThisAddIn


    <System.Runtime.InteropServices.DllImport("user32.dll",
    SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Public Shared Function GetAsyncKeyState(ByVal vKey As System.Int32) As System.Int16
    End Function

    ' Convenience constant (optional, but self-documenting)
    Private Const VK_ESCAPE As System.Int32 = &H1B



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

        If System.Threading.SynchronizationContext.Current Is Nothing Then
            System.Threading.SynchronizationContext.SetSynchronizationContext(
        New System.Windows.Forms.WindowsFormsSynchronizationContext())
        End If

        wordApp = Application
        Try
            If wordApp IsNot Nothing Then
                AddHandler wordApp.WindowActivate, AddressOf WordApp_WindowActivate
                AddHandler wordApp.DocumentOpen, AddressOf WordApp_DocumentOpen
                AddHandler wordApp.NewDocument, AddressOf WordApp_NewDocument
                AddHandler wordApp.ProtectedViewWindowOpen, AddressOf WordApp_ProtectedViewWindowOpen
                AddHandler wordApp.ProtectedViewWindowBeforeEdit, AddressOf WordApp_ProtectedViewWindowBeforeEdit
                AddHandler wordApp.ProtectedViewWindowActivate, AddressOf WordApp_ProtectedViewWindowActivate
                AddHandler wordApp.DocumentChange, AddressOf WordApp_DocumentChange
            Else
                mainThreadControl.BeginInvoke(CType(AddressOf DelayedStartupTasks, MethodInvoker))
                StartupInitialized = True
            End If
            If wordApp.Documents.Count > 0 Then
                'Run everything on the Office UI thread
                mainThreadControl.BeginInvoke(
                                            Sub()
                                                'Detach the one-shot startup hooks
                                                RemoveStartupHandlers()          'sets StartupInitialized = True
                                                DelayedStartupTasks()
                                            End Sub)
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
            pvWin As Microsoft.Office.Interop.Word.ProtectedViewWindow)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    ' Fires just before the user clicks “Edit” in Protected View.
    Private Sub WordApp_ProtectedViewWindowBeforeEdit(
            pvWin As Microsoft.Office.Interop.Word.ProtectedViewWindow,
            ByRef Cancel As Boolean)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    ' Fires when the Protected View window is activated.
    Private Sub WordApp_ProtectedViewWindowActivate(
            pvWin As Microsoft.Office.Interop.Word.ProtectedViewWindow)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    Private Sub WordApp_DocumentChange()
        If Not StartupInitialized Then
            RemoveStartupHandlers()
            DelayedStartupTasks()
        End If
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

    Public Const Version As String = "V.060825 Gen2 Beta Test"


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
    Private Const SameAsReplaceTrigger As String = "(sar)"
    Private Const KFTrigger As String = "(keepformat)"
    Private Const KFTrigger2 As String = "(kf)"
    Private Const KPFTrigger As String = "(keepparaformat)"
    Private Const KPFTrigger2 As String = "(kpf)"
    Private Const ObjectTrigger As String = "(file)"
    Private Const ObjectTrigger2 As String = "(clip)"
    Private Const InPlacePrefix As String = "Replace:"
    Private Const NewdocPrefix As String = "Newdoc:"
    Private Const AddPrefix As String = "Append:"
    Private Const AddPrefix2 As String = "Add:"
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
    Private Const SlidesPrefix As String = "Slides:"
    Private Const BubbleCutText As String = " (" & ChrW(&H2702) & ")"
    Private Const SearchNextTrigger As String = "Next:"
    Private Const BoWTrigger As String = "(bow)"
    Private Const ChunkTrigger As String = "(iterate)"
    Private Const EmbedTrigger As String = "(embed)"
    Private Const RefreshTrigger As String = "(refresh)"

    Private Const RegexSeparator1 As String = "|||"  ' Set also in SharedLibrary
    Private Const RegexSeparator2 As String = "§§§"  ' Set also in SharedLibrary 
    Private Const RIMenu = AN
    Private Const OldRIMenu = AN & " " & ChrW(&HD83D) & ChrW(&HDC09)
    Private Const MinHelperVersion = 1 ' Minimum version of the helper file that is required

    Public Const SearchChunkSize As Integer = 1 ' Size of chunks used for search (in characters)
    Public Const IgnoreMarkups As Boolean = False ' Whether to ignore markups in the text when doing a search

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

    Private Const Code_JsonTemplateFormatter As String = "Public Module JsonTemplateFormatter" & vbCrLf & "''' <summary>" & vbCrLf & "''' Hauptfunktion für JSON-String + Template" & vbCrLf & "''' </summary>" & vbCrLf & "Public Function FormatJsonWithTemplate(json As String, ByVal template As String) As String" & vbCrLf & "    Dim jObj As JObject" & vbCrLf & "    Try" & vbCrLf & "        jObj = JObject.Parse(json)" & vbCrLf & "    Catch ex As Newtonsoft.Json.JsonReaderException" & vbCrLf & "        Return $""[Fehler beim Parsen des JSON: {ex.Message}]""" & vbCrLf & "    End Try" & vbCrLf & "    NormalizeSources(jObj)" & vbCrLf & "    Return FormatJsonWithTemplate(jObj, template)" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "''' <summary>" & vbCrLf & "''' Hauptfunktion für direkten JObject + Template" & vbCrLf & "''' </summary>" & vbCrLf & "Public Function FormatJsonWithTemplate(jObj As JObject, ByVal template As String) As String" & vbCrLf & "    If String.IsNullOrWhiteSpace(template) Then Return """"" & vbCrLf & "    NormalizeSources(jObj)" & vbCrLf & "    ' Normalize CRLF / Platzhalter für Zeilenumbruch" & vbCrLf & "    template = template _" & vbCrLf & "        .Replace(""\\N"", vbCrLf) _" & vbCrLf & "        .Replace(""\\n"", vbCrLf) _" & vbCrLf & "        .Replace(""\\R"", vbCrLf) _" & vbCrLf & "        .Replace(""\\r"", vbCrLf)" & vbCrLf & "    template = Regex.Replace(template, ""<cr>"", vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "    Dim hasLoop = Regex.IsMatch(template, ""\\{\\%\\s*for\\s+([^\\s\\%]+)\\s*\\%\\}"", RegexOptions.Singleline)" & vbCrLf & "    Dim hasPh = Regex.IsMatch(template, ""\\{([^}]+)\\}"")" & vbCrLf & "    ' === Einfache Fallbehandlung ===" & vbCrLf & "    If Not hasLoop AndAlso Not hasPh Then" & vbCrLf & "        ' Template enthält keine Platzhalter → als einfacher JSONPath behandeln" & vbCrLf & "        Return FindJsonProperty(jObj, template)" & vbCrLf & "    End If" & vbCrLf & "    ' === Schleifen-Blöcke ===" & vbCrLf & "    Dim loopRegex = New Regex(""\\{\\%\\s*for\\s+([^%\\s]+)\\s*\\%\\}(.*?)\\{\\%\\s*endfor\\s*\\%\\}"", RegexOptions.Singleline Or RegexOptions.IgnoreCase)" & vbCrLf & "    Dim mLoop = loopRegex.Match(template)" & vbCrLf & "    While mLoop.Success" & vbCrLf & "        Dim fullBlock = mLoop.Value" & vbCrLf & "        Dim rawPath = mLoop.Groups(1).Value.Trim()" & vbCrLf & "        Dim innerTpl = mLoop.Groups(2).Value" & vbCrLf & "        Dim path = If(rawPath.StartsWith(""$""), rawPath, ""$."" & rawPath)" & vbCrLf & "        Dim tokens = jObj.SelectTokens(path)" & vbCrLf & "        Dim items = tokens.SelectMany(Function(t)" & vbCrLf & "            If t.Type = JTokenType.Array Then" & vbCrLf & "                Return CType(t, JArray).OfType(Of JObject)()" & vbCrLf & "            ElseIf t.Type = JTokenType.Object Then" & vbCrLf & "                Return {CType(t, JObject)}" & vbCrLf & "            Else" & vbCrLf & "                Return Enumerable.Empty(Of JObject)()" & vbCrLf & "            End If" & vbCrLf & "        End Function)" & vbCrLf & "        Dim rendered = items.Select(Function(o) FormatJsonWithTemplate(o, innerTpl)).ToArray()" & vbCrLf & "        template = template.Replace(fullBlock, If(rendered.Any, String.Join(vbCrLf & vbCrLf, rendered), """"))" & vbCrLf & "        mLoop = loopRegex.Match(template)" & vbCrLf & "    End While" & vbCrLf & "    ' === Platzhalter (non-gierig) ===" & vbCrLf & "    Dim phRegex = New Regex(""\\{(.+?)\\}"", RegexOptions.Singleline)" & vbCrLf & "    Dim result = template" & vbCrLf & "    For Each mPh As Match In phRegex.Matches(template)" & vbCrLf & "        Dim fullPh = mPh.Value" & vbCrLf & "        Dim content = mPh.Groups(1).Value" & vbCrLf & "        ' HTML- oder No-CR-Flag?" & vbCrLf & "        Dim isHtml As Boolean = False" & vbCrLf & "        Dim isNoCr As Boolean = False" & vbCrLf & "        If content.StartsWith(""htmlnocr:"", StringComparison.OrdinalIgnoreCase) Then" & vbCrLf & "            isHtml = True" & vbCrLf & "            isNoCr = True" & vbCrLf & "            content = content.Substring(""htmlnocr:"".Length)" & vbCrLf & "        ElseIf content.StartsWith(""html:"", StringComparison.OrdinalIgnoreCase) Then" & vbCrLf & "            isHtml = True" & vbCrLf & "            content = content.Substring(""html:"".Length)" & vbCrLf & "        ElseIf content.StartsWith(""nocr:"", StringComparison.OrdinalIgnoreCase) Then" & vbCrLf & "            isNoCr = True" & vbCrLf & "            content = content.Substring(""nocr:"".Length)" & vbCrLf & "        End If" & vbCrLf & "        ' Nur am ersten ""|"" trennen" & vbCrLf & "        Dim parts = content.Split(New Char() {""|""c}, 2)" & vbCrLf & "        Dim pathPh = parts(0).Trim()" & vbCrLf & "        Dim remainder = If(parts.Length > 1, parts(1), String.Empty)" & vbCrLf & "        ' Separator-Override (z.B. ""/"") oder Mapping-Definition (enthält ""="")" & vbCrLf & "        Dim sep As String = vbCrLf" & vbCrLf & "        Dim mappings As Dictionary(Of String, String) = Nothing" & vbCrLf & "        If Not String.IsNullOrEmpty(remainder) Then" & vbCrLf & "            If remainder.Contains(""=""c) Then" & vbCrLf & "                mappings = ParseMappings(remainder)" & vbCrLf & "            Else" & vbCrLf & "                sep = remainder.Replace(""\\n"", vbCrLf)" & vbCrLf & "            End If" & vbCrLf & "        End If" & vbCrLf & "        Dim replacement = RenderTokens(jObj, pathPh, sep, isHtml, isNoCr, mappings)" & vbCrLf & "        result = result.Replace(fullPh, replacement)" & vbCrLf & "    Next" & vbCrLf & "    Return result" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "''' <summary>" & vbCrLf & "''' Wandelt ausgewählte Tokens in einen String um, wendet Mapping, HTML→Markdown und No-CR an." & vbCrLf & "''' </summary>" & vbCrLf & "Private Function RenderTokens(" & vbCrLf & "    jObj As JObject," & vbCrLf & "    path As String," & vbCrLf & "    sep As String," & vbCrLf & "    isHtml As Boolean," & vbCrLf & "    isNoCr As Boolean," & vbCrLf & "    mappings As Dictionary(Of String, String)" & vbCrLf & ") As String" & vbCrLf & "    Try" & vbCrLf & "        If Not path.StartsWith(""$"") AndAlso Not path.StartsWith(""@"") Then" & vbCrLf & "            path = ""$."" & path" & vbCrLf & "        End If" & vbCrLf & "        Dim tokens = jObj.SelectTokens(path)" & vbCrLf & "        Dim list As New List(Of String)" & vbCrLf & "        For Each t In tokens" & vbCrLf & "            Dim raw = t.ToString()" & vbCrLf & "            ' Mapping anwenden, falls definiert" & vbCrLf & "            If mappings IsNot Nothing AndAlso mappings.ContainsKey(raw) Then raw = mappings(raw)" & vbCrLf & "            ' HTML→Markdown, falls gewünscht" & vbCrLf & "            If isHtml Then raw = HtmlToMarkdownSimple(raw)" & vbCrLf & "            ' No-CR: alle Zeilenumbrüche durch Leerzeichen" & vbCrLf & "            'If isNoCr Then raw = Regex.Replace(raw, ""[\\r\\n]+"", "" "").Trim()" & vbCrLf & "            If isNoCr Then" & vbCrLf & "                ' 1) Turn all line-breaks into single spaces" & vbCrLf & "                raw = Regex.Replace(raw, ""[\\r\\n]+"", "" "")" & vbCrLf & "                ' 2) Collapse any run of whitespace into one space" & vbCrLf & "                raw = Regex.Replace(raw, ""\\s{2,}"", "" "")" & vbCrLf & "                ' 3) Remove common Unicode bullet characters only" & vbCrLf & "                raw = Regex.Replace(raw, ""[\\u2022\\u2023\\u25E6]"", String.Empty)" & vbCrLf & "                ' 4) Trim leading/trailing spaces" & vbCrLf & "                raw = raw.Trim()" & vbCrLf & "            End If" & vbCrLf & "            list.Add(raw)" & vbCrLf & "        Next" & vbCrLf & "        Return If(list.Count = 0, """", String.Join(sep, list))" & vbCrLf & "    Catch ex As System.Exception" & vbCrLf & "        Return """"" & vbCrLf & "    End Try" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "''' <summary>" & vbCrLf & "''' Parst Mapping-Definitionen der Form ""key1=Text1;key2=Text2;…""" & vbCrLf & "''' </summary>" & vbCrLf & "Private Function ParseMappings(defs As String) As Dictionary(Of String, String)" & vbCrLf & "    Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)" & vbCrLf & "    For Each pair In defs.Split("";""c)" & vbCrLf & "        Dim kv = pair.Split(New Char() {""=""c}, 2)" & vbCrLf & "        If kv.Length = 2 Then dict(kv(0).Trim()) = kv(1).Trim()" & vbCrLf & "    Next" & vbCrLf & "    Return dict" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "''' <summary>" & vbCrLf & "''' Einfacher HTML→Markdown-Konverter (inkl. SPAN → *italic*)" & vbCrLf & "''' </summary>" & vbCrLf & "Public Function HtmlToMarkdownSimple(html As String) As String" & vbCrLf & "    Dim s = WebUtility.HtmlDecode(html)" & vbCrLf & "    ' Absätze → zwei Zeilenumbrüche            " & vbCrLf & "    s = Regex.Replace(s, ""</?p\\s*/?>"", vbCrLf & vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "    ' Zeilenumbruch-Tags" & vbCrLf & "    s = Regex.Replace(s, ""<br\\s*/?>"", vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "    ' Fett/strong → **text**" & vbCrLf & "    s = Regex.Replace(s, ""<strong>(.*?)</strong>"", ""**$1**"", RegexOptions.IgnoreCase)" & vbCrLf & "    ' Kursiv/em → *text*" & vbCrLf & "    s = Regex.Replace(s, ""<em>(.*?)</em>"", ""*$1*"", RegexOptions.IgnoreCase)" & vbCrLf & "    ' SPAN-Tags → *text*" & vbCrLf & "    s = Regex.Replace(s, ""<span\\b[^>]*>(.*?)</span>"", ""*$1*"", RegexOptions.IgnoreCase)" & vbCrLf & "    ' Listenpunkte <li> → ""- text""" & vbCrLf & "    s = Regex.Replace(s, ""<li>(.*?)</li>"", ""- $1"" & vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "    ' Fußnoten-Tags <fn>…</fn> → <sup>…</sup>" & vbCrLf & "    s = Regex.Replace(s, ""<fn>(.*?)</fn>"", ""<sup>$1</sup>"", RegexOptions.IgnoreCase)" & vbCrLf & "    ' Alle übrigen Tags entfernen" & vbCrLf & "    s = Regex.Replace(s, ""<(?!/?sup\\b)[^>]+>"", String.Empty, RegexOptions.IgnoreCase)" & vbCrLf & "    's = Regex.Replace(s, ""<[^>]+>"", String.Empty)" & vbCrLf & "    ' Mehrfache Zeilenumbrüche aufräumen" & vbCrLf & "    s = Regex.Replace(s, ""("" & vbCrLf & ""){3,}"", vbCrLf & vbCrLf)" & vbCrLf & "    Return s.Trim()" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "Private Sub NormalizeSources(jObj As JObject)" & vbCrLf & "    Dim srcToken = jObj.SelectToken(""sources"")" & vbCrLf & "    If srcToken IsNot Nothing AndAlso srcToken.Type = JTokenType.Array Then" & vbCrLf & "        Dim newArray As New JArray()" & vbCrLf & "        For Each item In CType(srcToken, JArray)" & vbCrLf & "            If item.Type = JTokenType.Array AndAlso item.Count >= 3 Then" & vbCrLf & "                Dim objStr = item(2).ToString()" & vbCrLf & "                Try" & vbCrLf & "                    Dim o = JObject.Parse(objStr)" & vbCrLf & "                    newArray.Add(o)" & vbCrLf & "                Catch ex As System.Exception" & vbCrLf & "                    ' Ungültiges JSON überspringen" & vbCrLf & "                End Try" & vbCrLf & "            ElseIf item.Type = JTokenType.Object Then" & vbCrLf & "                newArray.Add(item)" & vbCrLf & "            End If" & vbCrLf & "        Next" & vbCrLf & "        jObj(""sources"") = newArray" & vbCrLf & "    End If" & vbCrLf & "End Sub" & vbCrLf & "" & vbCrLf & "End Module"

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

    Public Shared Property INI_TokenCount As String
        Get
            Return _context.INI_TokenCount
        End Get
        Set(value As String)
            _context.INI_TokenCount = value
        End Set
    End Property

    Public Shared Property INI_TokenCount_2 As String
        Get
            Return _context.INI_TokenCount_2
        End Get
        Set(value As String)
            _context.INI_TokenCount_2 = value
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

    Public Shared Property SP_Add_Slides As String
        Get
            Return _context.SP_Add_Slides
        End Get
        Set(value As String)
            _context.SP_Add_Slides = value
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
    Public Shared Property SP_MergePrompt2 As String
        Get
            Return _context.SP_MergePrompt2
        End Get
        Set(value As String)
            _context.SP_MergePrompt2 = value
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
            If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> System.Drawing.Size.Empty Then
                chatForm.StartPosition = FormStartPosition.Manual
                chatForm.Location = My.Settings.FormLocation
                chatForm.Size = My.Settings.FormSize
            Else
                ' Default to center screen if no settings are available
                chatForm.StartPosition = FormStartPosition.Manual
                Dim screenBounds As System.Drawing.Rectangle = Screen.PrimaryScreen.WorkingArea
                chatForm.Location = New System.Drawing.Point((screenBounds.Width - chatForm.Width) \ 2, (screenBounds.Height - chatForm.Height) \ 2)
                chatForm.Size = New System.Drawing.Size(650, 500) ' Set default size if needed
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
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, False, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText1)
    End Sub
    Public Async Sub InLanguage2()

        If INILoadFail() Then Return
        TranslateLanguage = INI_Language2
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, False, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText1)
    End Sub
    Public Async Sub InOther()
        If INILoadFail() Then Return
        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, False, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText1)
        End If
    End Sub

    Public Async Sub Correct()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Correct), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText2)
    End Sub
    Public Async Sub Improve()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Improve), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText2)
    End Sub
    Public Async Sub Friendly()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Friendly), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText2)
    End Sub
    Public Async Sub Convincing()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Convincing), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText2)
    End Sub
    Public Async Sub NoFillers()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_NoFillers), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText2)
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


        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Anonymize), True, INI_KeepFormat2, INI_KeepParaFormatInline, DoReplace, DoMarkup, MarkupMethod, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not DoReplace)
    End Sub
    Public Async Sub Explain()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Explain), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, True, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText2)
    End Sub
    Public Async Sub SuggestTitles()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_SuggestTitles), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, True, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText2)
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
        Dim Selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If Selection.Type = WdSelectionType.wdSelectionIP Then
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
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Improve), True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText2)
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

        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_SwitchParty), True, INI_KeepFormat2, INI_KeepParaFormatInline, DoReplace, DoMarkup, MarkupMethod, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not DoReplace)

    End Sub
    Public Async Sub Summarize()
        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim Selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If Selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        Dim Textlength As Integer = GetSelectedTextLength()

        Dim UserInput As String
        SummaryLength = 0

        Do
            UserInput = SLib.ShowCustomInputBox("Enter the number of words your summary shall have (the selected text has " & Textlength & " words; the proposal " & SummaryPercent & "%):", $"{AN} Summarizer", True, CStr(System.Math.Round(SummaryPercent * Textlength / 100 / 5) * 5)).Trim()

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
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

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

        Dim FilePath As String = ""
        Dim FromFile As String = ""
        SelectedText = ""

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
        If selection.Type = WdSelectionType.wdSelectionIP Then
            Dim answer As Integer = ShowCustomYesNoBox("You have not selected any text. Do you instead want to create audio from a document file or add audio to a powerpoint with speaker notes?", "Yes", "No")
            If answer <> 1 Then Return

            DragDropFormLabel = "Document files (.txt, .docx, .pdf) or Powerpoint (.pptx)."
            DragDropFormFilter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.pptx||" &
                             "Text Files (*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                             "Rich Text Files (*.rtf)|*.rtf|" &
                             "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                             "PDF Files (*.pdf)|*.pdf" &
                             "Powerpoint Files (*.pptx)|*.pptx"

            FilePath = GetFileName()
            DragDropFormLabel = ""
            DragDropFormFilter = ""
            If String.IsNullOrWhiteSpace(FilePath) Then
                ShowCustomMessageBox("No file has been selected - will abort.")
                Return
            End If

            Dim ext As String = IO.Path.GetExtension(FilePath).ToLowerInvariant()

            Select Case ext
                Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm"
                    FromFile = ReadTextFile(FilePath, True)
                Case ".rtf"
                    FromFile = ReadRtfAsText(FilePath, True)
                Case ".doc", ".docx"
                    FromFile = ReadWordDocument(FilePath, True)
                Case ".pdf"
                    FromFile = ReadPdfAsText(FilePath, True)
                Case ".pptx"
                    FromFile = "pptx"
                Case Else
                    FromFile = "Error: File type not supported."
            End Select
            If FromFile.StartsWith("Error:") Then
                ShowCustomMessageBox(FromFile)
                Return
            End If
            If String.IsNullOrWhiteSpace(FromFile) Then
                ShowCustomMessageBox("The file you provided did not contain any text - will abort.")
                Return
            End If
            If FromFile <> "pptx" Then

                Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                newDoc.Activate()
                Dim currentSelection As Word.Selection = newDoc.Application.Selection
                currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                currentSelection.TypeText(FromFile)
                newDoc.Content.Select()
                SelectedText = newDoc.Application.Selection.Text.Trim()

                answer = ShowCustomYesNoBox("The content of your document has been inserted into a new document. Continue with the audio generation?", "Yes", "No")
                If answer <> 1 Then Return

            End If
        Else
            SelectedText = selection.Text.Trim()
        End If
        If SelectedText.Contains("H: ") And SelectedText.Contains("G: ") Then
            ReadPodcast(SelectedText)
        Else
            If selection.Text.Trim().StartsWith("{") Then
                Dim selectedoutputpath As String = (If(String.IsNullOrEmpty(My.Settings.TTSOutputPath), System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile), My.Settings.TTSOutputPath))
                selectedoutputpath = ShowCustomInputBox("Where should the audio generated from your JSON TTS file be saved to?", $"{AN} Create Audiobook", True, selectedoutputpath)
                If String.IsNullOrWhiteSpace(selectedoutputpath) Then
                    ' Use default path (Desktop) with default filename
                    selectedoutputpath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
                ElseIf selectedoutputpath.EndsWith("\") OrElse selectedoutputpath.EndsWith("/") Then
                    ' If only a folder is given, append default filename
                    selectedoutputpath = System.IO.Path.Combine(selectedoutputpath, TTSDefaultFile)
                Else
                    Dim dir As String = System.IO.Path.GetDirectoryName(selectedoutputpath)
                    Dim fileName As String = System.IO.Path.GetFileName(selectedoutputpath)

                    ' If no directory is found, assume Desktop as the base
                    If String.IsNullOrWhiteSpace(dir) Then
                        selectedoutputpath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName)
                        dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    End If

                    ' If no filename is given, use the default filename
                    If String.IsNullOrWhiteSpace(fileName) Then
                        selectedoutputpath = System.IO.Path.Combine(dir, TTSDefaultFile)
                    End If

                    ' Ensure the filename has ".mp3" extension
                    If Not fileName.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) Then
                        selectedoutputpath = System.IO.Path.Combine(dir, fileName & ".mp3")
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
                        If FromFile = "pptx" Then
                            GenerateAndPlayAudioFromSpeakerNotes(FilePath, selectedLanguage, selectedVoices(0).Replace(" (male)", "").Replace(" (female)", ""), If(Voices = 2, selectedVoices(1).Replace(" (male)", "").Replace(" (female)", ""), ""))

                            'Dim TokenErrorResponse As String = ValidatePptx(FilePath)

                            'If TokenErrorResponse = "" Then
                            'ShowCustomMessageBox($"Your slide deck at '{FilePath}' has been amended. Check it out.")
                            'Else
                            'ShowCustomMessageBox($"Your slide deck at '{FilePath}' has been amended, but the file may show certain problems and may require a repair (internal error: {TokenErrorResponse}).")
                            'End If

                        Else
                            GenerateAndPlayAudioFromSelectionParagraphs(outputPath, selectedLanguage, selectedVoices(0).Replace(" (male)", "").Replace(" (female)", ""), If(Voices = 2, selectedVoices(1).Replace(" (male)", "").Replace(" (female)", ""), ""))
                        End If
                    End If
                End Using
            End If
        End If
    End Sub

    Public Shared LastFreestyleModelConfig As ModelConfig
    Public Shared LastFreestyleWasAM As Boolean = False
    Public Shared LastFreestylePrompt As String = ""

    Public Async Sub FreeStyleNM()
        If INILoadFail() Then Return
        FreeStyle(False)

        My.Settings.LastFreestyleModelConfig = Nothing
        My.Settings.LastFreestyleWasAM = False
        My.Settings.LastFreestylePrompt = My.Settings.LastPrompt
        My.Settings.Save()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

    End Sub


    Public Async Sub FreeStyleAM()
        If INILoadFail() Then Return

        If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then

            If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                originalConfigLoaded = False
                Return
            End If

        End If

        LastFreestyleModelConfig = GetCurrentConfig(_context)

        FreeStyle(True)

        My.Settings.LastFreestyleModelConfig = LastFreestyleModelConfig
        My.Settings.LastFreestyleWasAM = True
        My.Settings.LastFreestylePrompt = My.Settings.LastPrompt
        My.Settings.Save()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

    End Sub

    Public Async Sub FreeStyleRepeat()
        If INILoadFail() Then Return

        Dim LastFreestylePrompt As String = My.Settings.LastFreestylePrompt

        originalConfig = GetCurrentConfig(_context)

        If String.IsNullOrWhiteSpace(LastFreestylePrompt) Then
            ShowCustomMessageBox("No last Freestyle command has been stored.")
            Return
        End If

        If My.Settings.LastFreestyleWasAM Then
            LastFreestyleModelConfig = My.Settings.LastFreestyleModelConfig

            If LastFreestyleModelConfig IsNot Nothing Then
                Dim ErrorFlag As Boolean = True
                ApplyModelConfig(_context, LastFreestyleModelConfig, ErrorFlag)
                If ErrorFlag Then
                    ShowCustomMessageBox("There was an error assigning the last model configuration. Aborting.")
                    Return
                End If
                originalConfigLoaded = True
            End If
        End If

        FreeStyle(My.Settings.LastFreestyleWasAM, My.Settings.LastFreestylePrompt)

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

            Dim iniValues() As String = {HideEscape(INI_Model_Parameter1), HideEscape(INI_Model_Parameter2), HideEscape(INI_Model_Parameter3), HideEscape(INI_Model_Parameter4)}
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

            If Not String.IsNullOrWhiteSpace(SP_QueryPrompt) Then
                parameterDefs.Add(New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Run query assistant", False))
            End If

            OtherPrompt = ""

            Dim runQueryAssistant As Boolean = False

            If parameterDefs.Count > 0 Then
                Dim parameters() As SharedLibrary.SharedLibrary.SharedMethods.InputParameter = parameterDefs.ToArray()
                If ShowCustomVariableInputForm("Please configure your parameters:", "Use '" & INI_Model_2 & "'", parameters) Then

                    Dim loopCount As Integer = parameters.Length

                    ' Wenn es einen SP_QueryPrompt gibt, ist der letzte Parameter unser Boolean-Flag
                    If Not String.IsNullOrWhiteSpace(SP_QueryPrompt) Then
                        Dim lastParam = parameters(parameters.Length - 1)
                        If TypeOf lastParam.Value Is System.Boolean Then
                            runQueryAssistant = CType(lastParam.Value, System.Boolean)
                            loopCount -= 1    ' excl. Flag aus der folgenden Schleife
                        End If
                    End If

                    ' === NEU: Werte auslesen mit Range-Clamping und Mapping ===
                    For i As Integer = 0 To loopCount - 1
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
                                    num = System.Math.Max(range.Item1, System.Math.Min(range.Item2, num))

                                    ' Für Integer/Long als Ganzzahl zurückgeben
                                    If t = "integer" OrElse t = "long" Then
                                        rawValue = CInt(System.Math.Round(num)).ToString()
                                    Else
                                        rawValue = num.ToString()
                                    End If
                                End If
                            End If


                            ' 3) Mapping von Display-Text → interner Code
                            If dispList IsNot Nothing Then
                                Dim idx As Integer = dispList.IndexOf(rawValue)
                                If idx >= 0 Then
                                    paramValue = UnHideEscape(codeList(idx))
                                Else
                                    ' Fallback: unverändert
                                    paramValue = UnHideEscape(rawValue)
                                End If
                                If paramValue.ToLowerInvariant().StartsWith("(keine Auswahl)") OrElse paramValue.ToLowerInvariant().StartsWith("(no selection)") OrElse paramValue.StartsWith("---") Then
                                    paramValue = ""
                                End If
                            Else
                                ' 4) Normaler String-Fall: (all)/(alle)/--- filtern
                                Dim rvLower As String = rawValue.ToLowerInvariant()
                                If rvLower.StartsWith("(keine Auswahl)") OrElse rvLower.StartsWith("(no selection)") OrElse rawValue.StartsWith("---") Then
                                    rawValue = ""
                                End If
                                paramValue = UnHideEscape(rawValue)

                            End If
                        End If

                        ' 5) Sonderfall Prompt
                        If p.Name.ToLowerInvariant().Contains("prompt") Then
                            OtherPrompt = UnHideEscape(paramValue)
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

            If runQueryAssistant Then

                Dim querytext As String = Await LLM(SP_QueryPrompt, "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>", "", "", 0, False)

                querytext = SLib.ShowCustomInputBox("This prompt has been generated based on your selection; modify it as you wish:", $"{AN} Query Assistant", False, querytext.Trim()).Trim()
                If String.IsNullOrWhiteSpace(querytext) OrElse querytext.ToLower() = "esc" Then
                    Return
                End If
                SelectedText = querytext.Trim()

            End If

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

                                                            Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                                                            If NewDocChoice = 1 Then
                                                                Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                                                                Dim currentSelection As Word.Selection = newDoc.Application.Selection
                                                                currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                                InsertTextWithMarkdown(currentSelection, llmresult, True, True)
                                                            ElseIf NewDocChoice = 2 Then
                                                                Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                                InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & llmresult, False)
                                                            Else
                                                                ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                                                                SLib.PutInClipboard(MarkdownToRtfConverter.Convert((llmresult)))
                                                            End If

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
                                    Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                                    If NewDocChoice = 1 Then
                                        Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                                        Dim currentSelection As Word.Selection = newDoc.Application.Selection
                                        currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                        InsertTextWithMarkdown(currentSelection, llmresult, True, True)
                                    ElseIf NewDocChoice = 2 Then
                                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                        InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & llmresult, False)
                                    Else
                                        ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                                        SLib.PutInClipboard(MarkdownToRtfConverter.Convert((llmresult)))
                                    End If
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


    ''' <summary>
    ''' Ersetzt jede Sequenz \\X durch \\uXXXX (doppelter Backslash!).
    ''' Aus \\; wird \\u003B, aus \\< wird \\u003C, usw.
    ''' </summary>
    Public Function HideEscape(ByVal input As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(input, "\\\\(.)",
            Function(m As System.Text.RegularExpressions.Match) As String
                Dim c As Char = m.Groups(1).Value(0)
                Dim hex As String = System.Convert.ToInt32(c).ToString("X4")
                Return "\\u" & hex
            End Function)
    End Function

    ''' <summary>
    ''' Ersetzt jede Sequenz \\uXXXX (doppelter Backslash!) zurück in das jeweilige Zeichen.
    ''' Aus \\u003B wird ;, aus \\u003C wird &lt;, usw.
    ''' </summary>
    Public Function UnHideEscape(ByVal input As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(input, "\\\\u([0-9A-Fa-f]{4})",
            Function(m As System.Text.RegularExpressions.Match) As String
                Dim code As Integer = Integer.Parse(m.Groups(1).Value, System.Globalization.NumberStyles.HexNumber)
                Return System.Convert.ToChar(code).ToString()
            End Function)
    End Function



    Public Async Sub FreeStyle(UseSecondAPI As Boolean, Optional LastPrompt As String = "")
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
            Dim DoNewDoc As Boolean = False
            Dim DoChunks As Boolean = False
            Dim ChunkSize As Integer = 1
            Dim NoFormatAndFieldSaving As Boolean = False
            Dim DoSlides As Boolean = False

            Dim MarkupInstruct As String = $"start With '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"with '{InPlacePrefix}'/'{AddPrefix} for replacing/adding to the selection"
            Dim BubblesInstruct As String = $"with '{BubblesPrefix}' for having your text commented"
            Dim SlidesInstruct As String = $"with '{SlidesPrefix}' for adding to a Powerpoint file"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}', '{NewdocPrefix}' or '{PanePrefix}' for separate output"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim ExtInstruct As String = $"; inlcude '{ExtTrigger}' for text of a file (txt, docx, pdf)"
            Dim TPMarkupInstruct As String = $"; add '{TPMarkupTriggerInstruct}' if revisions [of user] should be pointed out to the LLM"
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}/{SameAsReplaceTrigger}' for overriding formatting defaults"
            Dim AllInstruct As String = $"; add '{AllTrigger}' to select all"
            Dim LibInstruct As String = $"; add '{LibTrigger}' for library search"
            Dim NetInstruct As String = $"; add '{NetTrigger}' for internet search"
            Dim PureInstruct As String = $"; use '{PurePrefix}' for direct prompting"
            Dim ChunkInstruct As String = $"; add '{ChunkTrigger}' for iterating through the text"
            Dim ObjectInstruct As String = $"; add '{ObjectTrigger}'/'{ObjectTrigger2}' for adding a file object"
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
            Dim FileObject As String = ""
            Dim SlideDeck As String = ""

            Dim application As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

            If selection.Type = WdSelectionType.wdSelectionIP Then NoText = True

            Dim AddOnInstruct As String = AllInstruct

            If Not NoText Then
                AddOnInstruct += NoFormatInstruct.Replace("; add", ", ")
                AddOnInstruct += TPMarkupInstruct.Replace("; add", ", ")
                AddOnInstruct += ChunkInstruct.Replace("; add", ", ")
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

            If LastPrompt.Trim() = "" Then
                If Not NoText Then
                    OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {ClipboardInstruct}, {InplaceInstruct}, {BubblesInstruct} or {SlidesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{LastPromptInstruct}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt).Trim()
                Else
                    OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct} or {SlidesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{LastPromptInstruct}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt).Trim()
                End If
            Else
                OtherPrompt = LastPrompt
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
                ShowCustomMessageBox("I am using the " & INI_Model & " model as my primary model with a default timeout of " & (INI_Timeout / 1000) & " seconds (" & Microsoft.VisualBasic.Strings.Format(INI_Timeout / 60000, "0.00") & " minutes)." & If(INI_MaxOutputToken > 0, "The maximum output token length is " & INI_MaxOutputToken & ".", ""))
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


            If OtherPrompt.StartsWith("redinktest", StringComparison.OrdinalIgnoreCase) Then

                Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                Dim filePath As String = System.IO.Path.Combine(desktopPath, "redinktest.txt")
                If File.Exists(filePath) Then
                    Dim testtextorig As String = File.ReadAllText(filePath).Replace("\n", vbCrLf)
                    Dim testtext As String = SLib.ShowCustomWindow("Testfile content:", testtextorig, "", AN, False, True, True, True)
                    If testtext <> "" And testtext <> "Pane" Then
                        If testtext = "Markdown" Then
                            Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                            InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & testtextorig, False)
                        Else
                            SLib.PutInClipboard(testtext)
                        End If
                    ElseIf testtext = "Pane" Then
                        SP_MergePrompt_Cached = SP_MergePrompt
                        ShowPaneAsync(
                                                                            "Test Pane",
                                                                            testtextorig,
                                                                            "",
                                                                            AN,
                                                                            noRTF:=False,
                                                                            insertMarkdown:=True
                                                                            )
                    End If
                    Return
                Else
                    Return
                End If
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

            If OtherPrompt.StartsWith("clearlastprompt", StringComparison.OrdinalIgnoreCase) Then
                My.Settings.LastPrompt = ""
                My.Settings.LastFreestylePrompt = ""
                My.Settings.LastFreestyleModelConfig = Nothing
                My.Settings.LastFreestyleWasAM = False
                My.Settings.Save()
                Dim resultx = Globals.Ribbons.Ribbon1.InitializeAppAsync()
                ShowCustomMessageBox($"The last Freestyle prompt has been cleared.")

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

            If OtherPrompt.IndexOf(ChunkTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ChunkTrigger, "").Trim()
                DoChunks = True
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
            If Not DoInplace Then
                If OtherPrompt.IndexOf(SameAsReplaceTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    OtherPrompt = OtherPrompt.Replace(SameAsReplaceTrigger, "").Trim()
                Else
                    NoFormatAndFieldSaving = True
                End If
            End If

            If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger, "(a file object follows)").Trim()
            ElseIf DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a clipboard object follows)").Trim()
                DoFileObjectClip = True
            Else
                DoFileObject = False
            End If

            ' Regular expression to find text in the format "(markup:..." and extract until ")"
            Dim pattern As String = Regex.Escape(TPMarkupTriggerL) & "(.*?)" & Regex.Escape(TPMarkupTriggerR)
            'Dim pattern As String = $"\{TPMarkupTriggerL}(.*?)\{TPMarkupTriggerR}"
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
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(ClipboardPrefix2, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix2.Length).Trim()
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(NewdocPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(NewdocPrefix.Length).Trim()
                DoClipboard = True
                DoChunks = False
                DoNewDoc = True
            ElseIf OtherPrompt.StartsWith(BubblesPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(BubblesPrefix.Length).Trim()
                DoBubbles = True
            ElseIf OtherPrompt.StartsWith(SlidesPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(SlidesPrefix.Length).Trim()
                DoSlides = True
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(InPlacePrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(InPlacePrefix.Length).Trim()
                DoInplace = True
            ElseIf OtherPrompt.StartsWith(AddPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(AddPrefix.Length).Trim()
                DoInplace = False
            ElseIf OtherPrompt.StartsWith(AddPrefix2, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(AddPrefix2.Length).Trim()
                DoInplace = False
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
                DoChunks = False
            End If


            If OtherPrompt.IndexOf(NetTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NetTrigger, "").Trim()
                DoNet = True
            End If


            If Not String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
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

            If DoSlides Then
                DragDropFormLabel = "A Powerpoint (pptx) file."
                DragDropFormFilter = "Supported Files|*.pptx"
                SlideDeck = GetFileName()
                DragDropFormLabel = ""
                DragDropFormFilter = ""
                If String.IsNullOrWhiteSpace(SlideDeck) Then
                    ShowCustomMessageBox("No Powerpoint file has been selected - will abort. You can try again (use Ctrl-P to re-insert your prompt).")
                    Return
                End If
            End If


            If NoText AndAlso (DoBubbles Or DoChunks) Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Ask the LLM to comment on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Return
                End If
            End If

            If NoText AndAlso DoMarkup Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Do the markup on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Return
                End If
            End If

            If Not DoInplace AndAlso DoMarkup Then
                Dim AppendMarkup As Integer = ShowCustomYesNoBox("You have asked for a markup to be created, but according to the configuration, it will not replace your current selection but added to it at the end. Is this really what you want?", "Yes, add markup ", "No, replace text with markup")
                If AppendMarkup = 0 Then
                    Return
                ElseIf AppendMarkup = 2 Then
                    DoInplace = True
                    NoFormatAndFieldSaving = False
                End If
            End If

            If OtherPrompt.StartsWith(PurePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PurePrefix.Length).Replace("(a file object follows)", "").Replace("(a clipboard object follows)", "").Trim()
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

            If DoChunks Then
                Dim response As String = SLib.ShowCustomInputBox($"How many paragraphs shall be treated at the same time (max. 25)?", "Iterate through the text", True, ChunkSize.ToString()).Trim()
                If Not Integer.TryParse(response, ChunkSize) Then ChunkSize = 0
                If response = "" OrElse response.ToLower() = "esc" OrElse ChunkSize = 0 Then Return
                If ChunkSize > 25 Then ChunkSize = 25
            Else
                ChunkSize = 0
            End If

            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SysPrompt), True, DoKeepFormat, DoKeepParaFormat, DoInplace, DoMarkup, MarkupMethod, DoClipboard, DoBubbles, False, UseSecondAPI, KeepFormatCap, DoTPMarkup, TPMarkupName, False, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc, SlideDeck)

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

        Dim splash As New SLib.SplashScreen($"{AN6} is preparing to tickle{If(INI_RoastMe, " (inofficial version)", "")}...")
        splash.Show()
        splash.Refresh()

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
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
    ' - WithRevisions: Boolean flag to indicate that the input should contain Word revisions.
    ' - ChunkSize: Integer indicating how many paragraphs should be processed at once.
    ' - NoFormatAndFieldSaving: Boolean flag to indicate that no standard formatting/field saving should be applied to the selected text.
    ' - DoNewDoc: Boolean flag to indicate that the output should be placed in a new document.

    ' Global array to store paragraph formatting information
    Structure ParagraphFormatStructure
        Dim Style As Word.Style
        Dim FontName As String
        Dim FontSize As Nullable(Of Integer)
        Dim FontBold As Nullable(Of Integer)  ' 1=True, 0=False, Nothing=keep
        Dim FontItalic As Nullable(Of Integer)
        Dim FontUnderline As Nullable(Of WdUnderline)
        Dim FontColor As Nullable(Of Long)
        'Dim FontBold As Integer
        'Dim FontItalic As Integer
        'Dim FontUnderline As Word.WdUnderline
        'Dim FontColor As Word.WdColor
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

    Private Async Function ProcessSelectedText(SysCommand As String, CheckMaxToken As Boolean, KeepFormat As Boolean, ParaFormatInline As Boolean, InPlace As Boolean, DoMarkup As Boolean, MarkupMethod As Integer, PutInClipboard As Boolean, PutInBubbles As Boolean, SelectionMandatory As Boolean, UseSecondAPI As Boolean, FormattingCap As Integer, Optional DoTPMarkup As Boolean = False, Optional TPMarkupname As String = "", Optional CreatePodcast As Boolean = False, Optional FileObject As String = "", Optional DoPane As Boolean = False, Optional ChunkSize As Integer = 0, Optional NoFormatAndFieldSaving As Boolean = False, Optional DoNewDoc As Boolean = False, Optional SlideDeck As String = "") As Task(Of String)

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If SysCommand = "" Then
            ShowCustomMessageBox("The (system-)prompt for the LLM is missing.")
            Return ""
        End If

        If selection.Type = WdSelectionType.wdSelectionIP And SelectionMandatory Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return ""
        End If

        If selection.Type = WdSelectionType.wdSelectionIP Or selection.Tables.Count = 0 Or PutInClipboard Or PutInBubbles Then

            Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc, SlideDeck)

        Else

            Dim userdialog As Integer = ShowCustomYesNoBox("Your text contains tables. Shall each text section and each cell content be processed separately to avoid the table falling apart? This will take more time." & If(ChunkSize > 0, $" Your '(iterate)' parameter will apply only outside the tables.", "") & If(DoMarkup And MarkupMethod <> 2, " For the markup, the Diff markup will be used instead of the markup method choosen by you.", "") & " If you want to abort, close this window.", "No", "Yes, process each cell individually", $"{AN} Table Processing")

            If userdialog = 0 Then Return ""

            If userdialog = 2 Then

                MarkupMethod = 2

                Dim selRange As Range = selection.Range
                Dim docTables As Tables = selRange.Tables

                Dim isEntirelyWithinTable As Boolean = False
                Dim isWholeTable As Boolean = False
                Dim isPartialTableSelection As Boolean = False

                If selection.Tables.Count = 1 Then
                    Dim tbl As Microsoft.Office.Interop.Word.Table = selRange.Tables(1)
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


                    ' Fully-qualified per your guidelines
                    Dim sel As Word.Selection = application.Selection
                    Dim selCRange As Word.Range = sel.Range

                    ' Loop _only_ the cells the user actually selected
                    For Each cell As Word.Cell In sel.Cells
                        ' Make a working copy of the cell’s range, minus its end‐of‐cell marker
                        Dim cellRange As Word.Range = cell.Range.Duplicate
                        cellRange.End -= 1

                        ' Compute the overlap of selRange & cellRange
                        Dim intersection As Word.Range = selCRange.Duplicate
                        intersection.Start = System.Math.Max(cellRange.Start, selCRange.Start)
                        intersection.End = System.Math.Min(cellRange.End, selCRange.End)

                        ' If there is any overlap, process _only_ that text
                        If intersection.Start < intersection.End Then
                            ' keep UI responsive
                            System.Windows.Forms.Application.DoEvents()

                            ' show exactly what's being processed
                            intersection.Select()

                            ' your async processing call
                            Dim result = Await TrueProcessSelectedText(
                                SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline,
                                InPlace, DoMarkup, MarkupMethod, PutInClipboard,
                                PutInBubbles, SelectionMandatory, UseSecondAPI,
                                FormattingCap, DoTPMarkup, TPMarkupname, False,
                                FileObject, DoPane, 0, NoFormatAndFieldSaving, DoNewDoc)

                            ' throttle so Word doesn’t lock up
                            Await System.Threading.Tasks.Task.Delay(500)
                        End If
                    Next

                Else

                    ' Sort tables by their start positions in the selection
                    Dim tableList As New List(Of Microsoft.Office.Interop.Word.Table)
                    For i As Integer = 1 To docTables.Count
                        tableList.Add(docTables(i))
                    Next
                    tableList.Sort(Function(t1, t2) t1.Range.Start.CompareTo(t2.Range.Start))

                    Dim lastPos As Integer = selRange.Start

                    Dim splash As New SLib.SplashScreen("Processing table(s)... press 'Esc' to abort")
                    splash.Show()
                    splash.Refresh()

                    Dim IsExit As Boolean = False

                    For Each tbl As Microsoft.Office.Interop.Word.Table In tableList

                        System.Windows.Forms.Application.DoEvents()

                        If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                            Exit For
                        End If

                        If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or IsExit Then
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
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, ChunkSize * -1, NoFormatAndFieldSaving, DoNewDoc)
                                    Await System.Threading.Tasks.Task.Delay(500)
                                End If
                            Else

                                Do
                                    textChunk.Start += 1
                                Loop While textChunk.Tables.Count <> 0 And Not textChunk.Start = textChunk.End

                                If textChunk.Tables.Count = 0 AndAlso textChunk.Start < textChunk.End Then
                                    textChunk.Select()
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, ChunkSize * -1, NoFormatAndFieldSaving, DoNewDoc)
                                    Await System.Threading.Tasks.Task.Delay(500)
                                End If

                            End If
                        End If

                        ' Process the table itself (cells)
                        For Each row As Microsoft.Office.Interop.Word.Row In tbl.Rows
                            System.Windows.Forms.Application.DoEvents()

                            If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                                Exit For
                            End If

                            If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or IsExit Then
                                Exit For
                            End If
                            For Each cell As Microsoft.Office.Interop.Word.Cell In row.Cells
                                System.Windows.Forms.Application.DoEvents()

                                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                                    Exit For
                                End If

                                If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or IsExit Then
                                    ' Exit the loop
                                    Exit For
                                End If
                                Dim cellRange As Range = cell.Range
                                cellRange.End -= 1  ' Exclude cell marker
                                If cellRange.Start < cellRange.End Then
                                    cellRange.Select()
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, 0, NoFormatAndFieldSaving, DoNewDoc)
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
                            Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, ChunkSize * -1, NoFormatAndFieldSaving, DoNewDoc)
                        Else
                            Do
                                finalChunk.Start += 1
                            Loop While finalChunk.Tables.Count <> 0 And Not finalChunk.Start = finalChunk.End

                            finalChunk.End = selRange.End

                            If finalChunk.Tables.Count = 0 AndAlso finalChunk.Start < finalChunk.End Then
                                finalChunk.Select()
                                Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, ChunkSize * -1, NoFormatAndFieldSaving, DoNewDoc)
                            End If
                        End If
                    End If

                    splash.Close()
                End If

            ElseIf userdialog = 1 Then

                Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc)

            End If

        End If

        If Not PutInClipboard Then
            selection.Collapse(WdCollapseDirection.wdCollapseEnd)
            selection.MoveStart(WdUnits.wdCharacter, 0)
            selection.MoveEnd(WdUnits.wdCharacter, 0)
        End If

        Return ""

    End Function



    Private Async Function TrueProcessSelectedText(SysCommand As String, CheckMaxToken As Boolean, KeepFormat As Boolean, ParaFormatInline As Boolean, InPlace As Boolean, DoMarkup As Boolean, MarkupMethod As Integer, PutInClipboard As Boolean, PutInBubbles As Boolean, SelectionMandatory As Boolean, UseSecondAPI As Boolean, FormattingCap As Integer, Optional DoTPMarkup As Boolean = False, Optional TPMarkupname As String = "", Optional CreatePodcast As Boolean = False, Optional FileObject As String = "", Optional DoPane As Boolean = False, Optional ChunkSize As Integer = 0, Optional NoFormatAndFieldSaving As Boolean = False, Optional DoNewDoc As Boolean = False, Optional SlideDeck As String = "") As Task(Of String)

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
        Dim currentdoc As Word.Document = selection.Document

        Debug.WriteLine(
                    vbCrLf & "CheckMaxToken=" & CheckMaxToken &
                    vbCrLf & "KeepFormat=" & KeepFormat &
                    vbCrLf & "ParaFormatInline=" & ParaFormatInline &
                    vbCrLf & "InPlace=" & InPlace &
                    vbCrLf & "DoMarkup=" & DoMarkup &
                    vbCrLf & "PutInClipboard=" & PutInClipboard &
                    vbCrLf & "PutInBubbles=" & PutInBubbles &
                    vbCrLf & "SelectionMandatory=" & SelectionMandatory &
                    vbCrLf & "UseSecondAPI=" & UseSecondAPI &
                    vbCrLf & "DoTPMarkup=" & DoTPMarkup &
                    vbCrLf & "CreatePodcast=" & CreatePodcast &
                    vbCrLf & "DoPane=" & DoPane &
                    vbCrLf & "DoNewDoc=" & DoNewDoc &
                    vbCrLf & "Chunksize=" & ChunkSize &
                    vbCrLf & "Fileobject=" & FileObject &
                    vbCrLf & "Slidedeck=" & SlideDeck &
                    vbCrLf & "NoFormatAndFieldSaving=" & NoFormatAndFieldSaving
                )

        Try

            Dim SelectedText As String = ""
            Dim rng As Range
            Dim i As Integer
            Dim NoFormatting As Boolean = False
            Dim NoSelectedText As Boolean = False
            Dim trailingCR As Boolean
            Dim trailingCRcount As Integer = 0
            Dim DoSilent As Boolean = False

            If selection.Type = WdSelectionType.wdSelectionIP And SelectionMandatory Then
                Return ""
            End If

            If selection.Type = WdSelectionType.wdSelectionIP Then NoSelectedText = True

            rng = selection.Range

            Debug.WriteLine($"1Range Start = {rng.Start} Selection Start = {selection.Start}")
            Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
            Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

            ' Added for processing footnotes etc. too

            ' What story (main text, header, footer, footnote, etc.) am I in?
            Dim storyType As Word.WdStoryType = rng.StoryType
            ' Grab the full Range for that story from the document
            Dim storyRange As Word.Range = currentdoc.StoryRanges(storyType)


            If Not NoSelectedText Then

                If rng.Text.Length = 0 Then NoSelectedText = True
                If Not NoSelectedText And (KeepFormat Or ParaFormatInline) And FormattingCap > 0 And rng.Text.Length > FormattingCap Then NoFormatting = True

            End If

            If PutInBubbles Or PutInClipboard Or NoSelectedText Then NoFormatting = True

            If PutInBubbles Then
                DoMarkup = False
                PutInClipboard = False
            End If

            If PutInClipboard Then DoMarkup = False

            If DoTPMarkup Then NoFormatting = True

            If MarkupMethod = 4 And DoMarkup Then NoFormatting = True

            If ChunkSize > 0 Then
                DoSilent = True
                If DoMarkup Then

                    Select Case MarkupMethod
                        Case 1
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using Word compare. Iteration only works using the Regex method. Continue using Regex markup (the character cap will be ignored) or go without markups?", "Yes, Regex markups", "No, no markups")
                            If SilentMarkup = 1 Then
                                MarkupMethod = 4
                            ElseIf SilentMarkup = 2 Then
                                DoMarkup = False
                            Else
                                Return ""
                            End If
                        Case 2
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using Diff compare. Iteration only works using the Regex method. Continue using Diff markup (the character cap will be ignored) or go without markups?", "Yes, Diff markups", "No, no markups")
                            If SilentMarkup = 1 Then
                                MarkupMethod = 2
                            ElseIf SilentMarkup = 2 Then
                                DoMarkup = False
                            Else
                                Return ""
                            End If
                        Case 3
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using DiffW compare. Iteration only works using the Regex method. Continue using Diff markup (the character cap will be ignored) or go without markups?", "Yes, Diff markups", "No, no markups")
                            If SilentMarkup = 1 Then
                                MarkupMethod = 2
                            ElseIf SilentMarkup = 2 Then
                                DoMarkup = False
                            Else
                                Return ""
                            End If
                        Case 4
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using the Regex method. This works, but may tak a very long time (the character cap will be ignored). Continue with Regex markups or go without markups?", "Yes, Regex markups", "No, no markups")
                            If SilentMarkup = 1 Then
                                MarkupMethod = 4
                            ElseIf SilentMarkup = 2 Then
                                DoMarkup = False
                            Else
                                Return ""
                            End If
                    End Select
                End If
            End If

            If ChunkSize < 0 Then
                ChunkSize = ChunkSize * -1
                If DoMarkup Then MarkupMethod = 2  ' Force Diff when getting a negative ChunkSize (e.g., in tables)
                DoSilent = True
            End If

            Dim effectiveChunk As Integer = If(ChunkSize > 0, ChunkSize, Integer.MaxValue)

            Dim totalEndBm As Word.Bookmark


            ' Added for processing footnotes etc. too

            Dim docEnd As Integer = storyRange.End

            If selection.End < docEnd Then
                totalEndBm = currentdoc.Bookmarks.Add(
                    Name:="TotalEnd",
                    Range:=currentdoc.Range(Start:=selection.End, End:=selection.End))
            Else
                Dim endRange As Word.Range = storyRange.Duplicate
                endRange.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                totalEndBm = currentdoc.Bookmarks.Add(
                    Name:="TotalEnd",
                    Range:=endRange)
            End If

            Debug.WriteLine($"2Range Start = {rng.Start} Selection Start = {selection.Start}")
            Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
            Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

            Dim safeRange As Word.Range = selection.Range
            safeRange.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)

            Dim nextStartBm As Word.Bookmark = currentdoc.Bookmarks.Add(
                    Name:="NextStart",
                    Range:=safeRange)

            Do While NoSelectedText OrElse ((currentdoc.Bookmarks.Exists("NextStart") AndAlso currentdoc.Bookmarks.Exists("TotalEnd") AndAlso currentdoc.Bookmarks("NextStart").Range.Start < currentdoc.Bookmarks("TotalEnd").Range.Start - 1))


                Try

                    If Not NoSelectedText Then

                        Dim curStart As Integer = currentdoc.Bookmarks("NextStart").Range.Start
                        Dim totalEnd As Integer = currentdoc.Bookmarks("TotalEnd").Range.Start

                        Do While currentdoc.Range(Start:=curStart, End:=curStart + 1).Text = vbCr
                            curStart += 1
                        Loop

                        ' ---- 2.1  Chunk-Ende bestimmen ----------------------------
                        'docEnd = currentdoc.Content.End  ' was used before code for footnotes etc. has been added.
                        docEnd = storyRange.End
                        Dim restRng As Word.Range = currentdoc.Range(Start:=curStart, End:=totalEnd)
                        Dim paras As Word.Paragraphs = restRng.Paragraphs

                        Dim chunkEnd As Integer

                        If paras.Count <= effectiveChunk Then
                            chunkEnd = totalEnd
                        Else
                            ' Start with the end of the effectiveChunk-th paragraph
                            Dim xxi As Integer = effectiveChunk
                            Dim paraRng As Word.Range = paras(xxi).Range
                            Dim paraText As String = paraRng.Text.Trim()

                            ' Keep extending while paragraph is empty and more paras are available
                            Do While (paraText = "" OrElse paraText = vbCr) AndAlso xxi < paras.Count
                                xxi += 1
                                paraRng = paras(xxi).Range
                                paraText = paraRng.Text.Trim()
                            Loop

                            chunkEnd = paraRng.End
                        End If


                        ' Grenzen sichern, um Range-Fehler zu vermeiden
                        If chunkEnd > docEnd Then chunkEnd = docEnd
                        If chunkEnd <= curStart Then chunkEnd = System.Math.Min(curStart + 1, docEnd)

                        ' ---- 2.2  Selection auf diesen Chunk ----------------------
                        selection.SetRange(Start:=curStart, End:=chunkEnd)
                        rng = selection.Range

                        If rng Is Nothing OrElse rng.Text.Trim() = "" Then Exit Do

                    End If
                Catch ex As System.Exception
                    Exit Do
                End Try

                Debug.WriteLine($"3Range Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                paraCount = 0
                trailingCR = False
                trailingCRcount = 0


                If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoSelectedText AndAlso Not NoFormatAndFieldSaving Then

                    paraCount = rng.Paragraphs.Count

                    ReDim paragraphFormat(paraCount - 1)
                    Array.Clear(paragraphFormat, 0, paragraphFormat.Length)


                    For i = 1 To paraCount
                        Dim para As Word.Paragraph = rng.Paragraphs(i)
                        Dim paraRange As Word.Range = para.Range

                        '---- bodyRange = text without the paragaph mark -------------------
                        Dim bodyRange As Word.Range = paraRange.Duplicate
                        bodyRange.MoveEnd(Word.WdUnits.wdCharacter, -1)

                        Try
                            '---- character-level attributes – store only when uniform -----
                            Dim boldV As Integer? = Nothing
                            Dim italicV As Integer? = Nothing
                            Dim underlineV As Word.WdUnderline? = Nothing
                            Dim colorV As Long? = Nothing

                            If bodyRange.Font.Bold <> Word.WdConstants.wdUndefined Then _
                                boldV = bodyRange.Font.Bold
                            If bodyRange.Font.Italic <> Word.WdConstants.wdUndefined Then _
                                italicV = bodyRange.Font.Italic
                            If bodyRange.Font.Underline <> Word.WdConstants.wdUndefined Then _
                                underlineV = CType(bodyRange.Font.Underline, Word.WdUnderline)
                            If bodyRange.Font.Color <> Word.WdConstants.wdUndefined Then _
                                colorV = bodyRange.Font.Color

                            Dim fname As String = Nothing
                            Dim fsize As Single? = Nothing
                            If bodyRange.Font.Name <> CStr(Word.WdConstants.wdUndefined) Then _
                                fname = bodyRange.Font.Name
                            If bodyRange.Font.Size <> CSng(Word.WdConstants.wdUndefined) Then _
                                fsize = bodyRange.Font.Size

                            '---- assign into the (freshly resized) array ------------------
                            paragraphFormat(i - 1) = New ParagraphFormatStructure With {
                                .Style = para.Style,
                                .FontName = fname,
                                .FontSize = fsize,
                                .FontBold = boldV,
                                .FontItalic = italicV,
                                .FontUnderline = underlineV,
                                .FontColor = colorV,
                                .ListType = bodyRange.ListFormat.ListType,
                                .ListTemplate = If(bodyRange.ListFormat.ListType <>
                                                    Word.WdListType.wdListNoNumbering,
                                                    bodyRange.ListFormat.ListTemplate, Nothing),
                                .ListLevel = If(bodyRange.ListFormat.ListType <>
                                                    Word.WdListType.wdListNoNumbering,
                                                    bodyRange.ListFormat.ListLevelNumber, 0),
                                .ListNumber = If(bodyRange.ListFormat.ListType <>
                                                    Word.WdListType.wdListNoNumbering,
                                                    bodyRange.ListFormat.ListValue, 0),
                                .HasListFormat = bodyRange.ListFormat.ListType <>
                                                    Word.WdListType.wdListNoNumbering,
                                .Alignment = para.Alignment,
                                .LineSpacing = para.LineSpacing,
                                .SpaceBefore = para.SpaceBefore,
                                .SpaceAfter = para.SpaceAfter
                                        }

                        Catch ex As System.Exception
                            'Debug.Print($"Error extracting paragraph {i} {ex.Message}")
                        End Try
                    Next

                End If

                Debug.WriteLine($"4Range Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)


                Dim raw As String = ""

                If (PutInBubbles Or PutInClipboard) AndAlso Not DoTPMarkup AndAlso rng IsNot Nothing Then
                    raw = GetVisibleText(rng)
                End If

                Debug.WriteLine($"4aRange Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")


                If Not NoSelectedText Then

                    If KeepFormat And Not NoFormatting Then
                        SelectedText = SLib.GetRangeHtml(rng)
                    Else
                        If NoFormatting OrElse NoFormatAndFieldSaving Then
                            If DoTPMarkup Then
                                SelectedText = AddMarkupTags(rng, TPMarkupname)
                            Else
                                SelectedText = rng.Text
                                If Not String.IsNullOrWhiteSpace(raw) Then SelectedText = raw
                            End If
                        Else
                            If INI_MarkdownConvert AndAlso Not KeepFormat AndAlso (Not DoMarkup OrElse (MarkupMethod = 3 Or MarkupMethod = 2)) AndAlso InPlace Then  ' AndAlso rng.Text.Length < INI_MarkupDiffCap 

                                Debug.WriteLine($"4bRange Start = {rng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                                SelectedText = GetTextWithSpecialElementsInline(rng, If(NoFormatting, False, ParaFormatInline), True)

                                Debug.WriteLine($"4cRange Start = {rng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                            Else
                                SelectedText = GetTextWithSpecialElementsInline(rng, If(NoFormatting, False, ParaFormatInline), False)
                                'SelectedText = LegacyGetTextWithSpecialElementsInline(rng, If(NoFormatting, False, ParaFormatInline))
                            End If
                        End If
                        trailingCR = (SelectedText.EndsWith(vbCrLf) Or SelectedText.EndsWith(vbLf) Or SelectedText.EndsWith(vbCr))
                        Dim tempText As String = SelectedText

                        Do While tempText.EndsWith(vbCrLf) Or tempText.EndsWith(vbLf) Or tempText.EndsWith(vbCr)
                            If tempText.EndsWith(vbCrLf) Then
                                trailingCRcount += 1
                                tempText = tempText.Substring(0, tempText.Length - vbCrLf.Length)
                            ElseIf tempText.EndsWith(vbLf) Then
                                trailingCRcount += 1
                                tempText = tempText.Substring(0, tempText.Length - vbLf.Length)
                            ElseIf tempText.EndsWith(vbCr) Then
                                trailingCRcount += 1
                                tempText = tempText.Substring(0, tempText.Length - vbCr.Length)
                            End If
                        Loop

                    End If

                    Debug.WriteLine($"4dRange Start = {rng.Start} Selection Start = {selection.Start}")
                    Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")

                    Dim MaxToken As Integer = If(UseSecondAPI, INI_MaxOutputToken_2, INI_MaxOutputToken)
                    Dim EstimatedTokens As Integer = EstimateTokenCount(SelectedText)

                    If CheckMaxToken And MaxToken > 0 AndAlso EstimatedTokens > MaxToken AndAlso (InPlace Or DoMarkup) AndAlso Not DoSilent Then
                        ShowCustomMessageBox("Your selected text Is larger than the maximum output your LLM can supposedly generate. Therefore, the output may be shorter than expected based on maximum tokens supported, which Is " & MaxToken & " tokens. Your input (with formatting information, as the case may be) has an estimated to be " & EstimatedTokens & " tokens). Therefore, check whether the output Is complete.", AN, 15)
                    End If

                    If DoMarkup AndAlso MarkupMethod = 2 AndAlso Len(SelectedText) > INI_MarkupDiffCap AndAlso Not DoSilent Then
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

                    If DoMarkup And MarkupMethod = 4 And Len(SelectedText) > INI_MarkupRegexCap AndAlso Not DoSilent Then
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

                Debug.WriteLine($"5Range Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                Dim SlideInsert As String = ""

                If SlideDeck <> "" Then
                    SlideInsert = GetPresentationJson(SlideDeck)
                    Debug.WriteLine("SlideInsert = " & SlideInsert)
                    If SlideDeck = "" Then
                        Return ""
                    Else
                        SlideInsert = " <SLIDEDECK>" & SlideInsert & "</SLIDEDECK>"
                    End If
                End If

                Dim LLMResult = Await LLM(SysCommand & If(DoTPMarkup, " " & SP_Add_Revisions, "") & " " & If(SlideDeck = "", If(NoFormatting, "", If(KeepFormat, " " & SP_Add_KeepHTMLIntact, " " & SP_Add_KeepInlineIntact)), " " & SP_Add_Slides), If(NoSelectedText, "" & SlideInsert, "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>" & SlideInsert), "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

                OtherPrompt = ""

                LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

                If Not String.IsNullOrEmpty(LLMResult) Then
                    LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                End If

                Debug.WriteLine($"LLMResult 1 = '{LLMResult}'")

                If ParaFormatInline Then LLMResult = CorrectPFORMarkers(LLMResult)

                Debug.WriteLine($"LLMResult 2 = '{LLMResult}'")

                If DoTPMarkup Then LLMResult = RemoveMarkupTags(LLMResult)

                'If (MarkupMethod <> 4 Or Not DoMarkup) And InPlace And Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                'If (MarkupMethod <> 4 Or Not DoMarkup) And InPlace And Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)

                'If (MarkupMethod <> 4 Or Not DoMarkup) And trailingCR And (LLMResult.EndsWith(ControlChars.Cr) Or LLMResult.EndsWith(ControlChars.Lf)) Then LLMResult = LLMResult.Replace(ControlChars.Cr, ControlChars.CrLf).Replace(ControlChars.Lf, ControlChars.CrLf)

                If Not trailingCR AndAlso LLMResult.EndsWith(ControlChars.CrLf) Then LLMResult = LLMResult.TrimEnd(ControlChars.CrLf)
                If Not trailingCR AndAlso LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                If Not trailingCR AndAlso LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)

                If trailingCR Then
                    LLMResult = LLMResult.TrimEnd({ControlChars.Cr, ControlChars.Lf})
                    If trailingCRcount > 1 Then
                        LLMResult &= String.Concat(Enumerable.Repeat(vbCrLf, trailingCRcount - 1))
                    End If
                End If

                Debug.WriteLine($"LLMResult 3 = '{LLMResult}'")
                Debug.WriteLine($"TrailingCR = {trailingCR} Count = {trailingCRcount}")

                Debug.WriteLine($"6Range Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                If Not String.IsNullOrEmpty(LLMResult) Then


                    Debug.WriteLine(
                                vbCrLf & "PutInClipboard=" & PutInClipboard &
                                vbCrLf & "DoSilent=" & DoSilent &
                                vbCrLf & "DoMarkup=" & DoMarkup &
                                vbCrLf & "MarkupMethod=" & MarkupMethod &
                                vbCrLf & "DoPane=" & DoPane &
                                vbCrLf & "DoNewDoc=" & DoNewDoc &
                                vbCrLf & "PutInBubbles=" & PutInBubbles &
                                vbCrLf & "NoSelectedText=" & NoSelectedText &
                                vbCrLf & "ParaFormatInline=" & ParaFormatInline &
                                vbCrLf & "NoFormatting=" & NoFormatting &
                                vbCrLf & "NoFormatAndFieldSaving=" & NoFormatAndFieldSaving &
                                vbCrLf & "KeepFormat=" & KeepFormat &
                                vbCrLf & "Inplace=" & InPlace
                            )

                    Dim ClipPaneText1 As String = "The LLM has provided the following result (you can edit it)"
                    Dim ClipText2 As String = "You can choose whether you want to have the original text put into the clipboard Or your text with any changes you have made (without formatting), Or you can directly insert the original text in your document. If you select Cancel, nothing will be put into the clipboard."
                    Dim PaneText2 As String = "Choose to put your edited Or original text in the clipboard, Or inserted the original with formatting; the pane will close. You can also copy & paste from the pane."

                    If CreatePodcast AndAlso Not DoSilent Then
                        Dim TTSAvailable As Boolean = False

                        DetectTTSEngines()

                        If Not TTS_googleAvailable AndAlso Not TTS_openAIAvailable Then
                            TTSAvailable = False
                        Else
                            TTSAvailable = True
                        End If


                        If TTSAvailable Then
                            Dim FinalText = ShowCustomWindow("The LLM has created the following podcast script for you (you can edit it; you do Not have to manually remove the SSML codes, if you do Not Like them)", LLMResult, "The next step Is the production of an audio file. You can choose whether you want to use the original text or your text with any changes you have made. The text will also be put in the clipboard. If you select Cancel, the original text will only be put into the clipboard.", AN, True)

                            If FinalText = "" Then
                                SLib.PutInClipboard(MarkdownToRtfConverter.Convert(LLMResult))
                            Else
                                FinalText = FinalText.Trim()
                                SLib.PutInClipboard(FinalText)
                                If FinalText.Contains("H: ") AndAlso FinalText.Contains("G: ") Then ReadPodcast(FinalText)
                            End If
                        Else
                            Dim FinalText = ShowCustomWindow("The LLM has created the following podcast script for you (you can edit it; you do Not have to manually remove the SSML codes, if you do Not Like them)", LLMResult, $"The next step Is the production of an audio file. Since you have not configured {AN} for Google, you unfortunately cannot do that here. However, you can choose whether you want the original text Or the text with your changes to put in the clipboard for further use. If you select Cancel, no text will be put in the clipboard.", AN, True)

                            If FinalText <> "" Then
                                SLib.PutInClipboard(MarkdownToRtfConverter.Convert(LLMResult))
                            Else
                                FinalText = FinalText.Trim()
                                SLib.PutInClipboard(FinalText)
                            End If
                        End If
                    ElseIf SlideInsert <> "" Then

                        Dim Jsonstring As String = CleanJsonString(LLMResult)

                        Debug.WriteLine("JsonString=" & Jsonstring)

                        If Not String.IsNullOrEmpty(Jsonstring) Then

                            If ApplyPlanToPresentation(SlideDeck, Jsonstring) Then

                                Dim TokenErrorResponse As String = ValidatePptx(SlideDeck)

                                If TokenErrorResponse = "" Then
                                    ShowCustomMessageBox($"Your slide deck at '{SlideDeck}' has been amended as per the AI's instruction. Check it out.")
                                Else
                                    ShowCustomMessageBox($"Your slide deck at '{SlideDeck}' has been amended as per the AI's instruction, but the file may show certain problems and may require a repair (internal error: {TokenErrorResponse}).")
                                End If
                            End If

                        Else
                            ShowCustomMessageBox($"There was a problem converting the AI response. You may want to retry.")
                        End If


                    ElseIf DoPane AndAlso Not DoSilent Then

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

                    ElseIf DoNewDoc AndAlso Not DoSilent Then

                        Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                        newDoc.Activate()
                        Dim newSelection As Word.Selection = Globals.ThisAddIn.Application.Selection
                        InsertTextWithMarkdown(newSelection, LLMResult, True, True)

                    ElseIf PutInClipboard AndAlso Not DoSilent Then

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

                                                            Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                                                            If NewDocChoice = 1 Then
                                                                Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                                                                Dim currentSelection As Word.Selection = newDoc.Application.Selection
                                                                currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                                InsertTextWithMarkdown(currentSelection, LLMResult, True, True)
                                                            ElseIf NewDocChoice = 2 Then
                                                                Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                                InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & LLMResult, False)
                                                            Else
                                                                ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                                                                SLib.PutInClipboard(MarkdownToRtfConverter.Convert((LLMResult)))
                                                            End If
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

                                    Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                                    If NewDocChoice = 1 Then
                                        Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                                        Dim currentSelection As Word.Selection = newDoc.Application.Selection
                                        currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                        InsertTextWithMarkdown(currentSelection, LLMResult, True, True)
                                    ElseIf NewDocChoice = 2 Then
                                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                        InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, LLMResult, False)
                                    Else
                                        ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                                        SLib.PutInClipboard(MarkdownToRtfConverter.Convert((LLMResult)))
                                    End If
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
                        Dim BubbleCount As Integer = 0
                        Dim MaxBubbles As Integer = responseItems.Count

                        If MaxBubbles = 0 Then
                            If Not DoSilent Then ShowCustomMessageBox($"The bubble command did Not result in any comment(s) by the LLM.")
                        Else

                            Dim splash As New SLib.SplashScreen($"Adding {MaxBubbles} bubble(s) to your text... press 'Esc' to abort")
                            splash.Show()
                            splash.Refresh()

                            For Each item In responseItems

                                splash.UpdateMessage($"Adding {MaxBubbles - BubbleCount} bubble(s) to your text... press 'Esc' to abort")

                                System.Windows.Forms.Application.DoEvents()

                                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exit For

                                If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                                    Exit For
                                End If

                                Dim parts() As String = item.Split({"@@"}, StringSplitOptions.None)
                                If parts.Length = 2 Then

                                    Dim findText As String = parts(0).Trim().Trim("'"c).Trim(""""c)
                                    Dim commentText As String = parts(1).Trim()

                                    Try
                                        If findText.Length <= SearchChunkSize Then
                                            ' Use the built-in Find directly if <= 255 characters
                                            With selection.Find
                                                .ClearFormatting()
                                                .Text = NormalizeTextForSearch(findText, INI_Clean)
                                                .Forward = True
                                                .Wrap = Word.WdFindWrap.wdFindStop
                                                .MatchWildcards = True
                                            End With
                                            Debug.WriteLine($"Searching for: '{findText}' with normalized text: '{selection.Find.Text}'")

                                            If selection.Find.Execute() Then
                                                Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, $"{AN5}: " & commentText)
                                                BubbleCount += 1
                                            Else
                                                notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}: " & commentText & vbCrLf & vbCrLf)
                                            End If

                                        Else
                                            ' Use chunk-by-chunk search for > 255 characters
                                            If FindLongTextInChunks(findText, SearchChunkSize, selection, True) Then
                                                ' If found, selection now covers the entire matched text
                                                Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, $"{AN5}: " & commentText)
                                                BubbleCount += 1
                                            Else
                                                notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}: " & commentText & vbCrLf & vbCrLf)
                                            End If
                                        End If

                                    Catch ex As Exception
                                        notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}: " & commentText & " [Error: " & ex.Message & "]" & vbCrLf & vbCrLf)
                                    End Try

                                Else
                                    If Not String.IsNullOrWhiteSpace(item) Then
                                        wrongformatresponse.Add(item)
                                    End If
                                End If

                                selection.SetRange(originalRange.Start, originalRange.End) ' Restore the original selection
                            Next

                            splash.Close()

                            Dim ErrorList As String = ""

                            If notfoundresponse.Count > 0 Then
                                ErrorList += "The following comments could not be assigned to your text (they were not found, typically because of formatting or markup issues):" & vbCrLf & vbCrLf
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

                                If Not DoSilent Then
                                    ErrorList = ShowCustomWindow($"{BubbleCount} bubble comment(s) applied (Warning: complicated formatting and markups may cause misalignments of the commented portions of the text). The following errors occurred when implementing the 'bubbles' feedback of the LLM:", ErrorList, "The above error list will be included in a final comment at the end of your selection (it will also be included in the clipboard). You can have the original list included, or you can now make changes and have this version used. If you select Cancel, nothing will be put added to the document.", AN, True)
                                End If
                                If ErrorList <> "" And ErrorList.ToLower() <> "esc" Then
                                    SLib.PutInClipboard(ErrorList)
                                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                    Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, $"{AN5}: " & ErrorList)
                                End If

                            Else

                                If Not DoSilent Then ShowCustomMessageBox($"{BubbleCount} bubble comment(s) provided by the LLM applied to to your text (Warning: complicated formatting and markups may cause misalignments of the commented portions of the text)." & If(BubblecutHappened, $"Some of the sections to which the bubble comments relate were too long for selecting. Only the initial part has been selected. This is indicated by '{BubbleCutText}' in the bubble comments, as applicable.", ""))
                            End If
                        End If

                    ElseIf MarkupMethod = 4 Then

                        Dim RegexResult = Await LLM(SP_MarkupRegex, "<ORIGINALTEXT>" & SelectedText & "</ORIGINALTEXT> /n <NEWTEXT>" & LLMResult & "</NEWTEXT>", "", "", 0, False)

                        MarkupSelectedTextWithRegex(RegexResult)

                        ' End Extended Selection Mode
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ElseIf NoSelectedText Then

                        InsertTextWithMarkdown(selection, vbCrLf & LLMResult, trailingCR, True)

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
                                If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then
                                    ApplyParagraphFormat(rng)
                                End If
                                If Not NoFormatAndFieldSaving Then RestoreSpecialTextElements(SaveRng)
                                SaveRng.Document.Fields.Update()
                            Else
                                CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                                If Not NoFormatAndFieldSaving Then RestoreSpecialTextElements(rng)
                                rng.Document.Fields.Update()
                            End If
                        End If

                        ' End Extended Selection Mode
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    Else
                        SelectedText = selection.Text

                        Debug.WriteLine($"7Range Start = {rng.Start} Selection Start = {selection.Start}")
                        Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                        Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

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
                                    If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then
                                        ApplyParagraphFormat(rng)
                                    End If
                                    If Not NoFormatAndFieldSaving Then RestoreSpecialTextElements(SaveRng)
                                    SaveRng.Document.Fields.Update()
                                Else
                                    If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                    CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                                    If Not NoFormatAndFieldSaving Then RestoreSpecialTextElements(rng)
                                    rng.Document.Fields.Update()
                                End If
                            Else
                                InsertTextWithMarkdown(selection, LLMResult, trailingCR)

                                Debug.WriteLine($"8Range Start = {rng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                                rng = selection.Range
                                Dim SaveRng As Range = rng.Duplicate
                                If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then   '  xxxxx 
                                    Debug.WriteLine($"9Range Start = {rng.Start} Selection Start = {selection.Start}")
                                    Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                                    Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                                    ApplyParagraphFormat(rng)
                                End If

                                Debug.WriteLine($"10Range Start = {rng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")

                                Debug.WriteLine($"SaveRange Start = {SaveRng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"SaveRange End = {SaveRng.End} Selection End = {selection.End}")

                                If Not NoFormatAndFieldSaving Then RestoreSpecialTextElements(SaveRng)
                                SaveRng.Document.Fields.Update()
                            End If

                        Else
                            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            selection.TypeText(vbCrLf)
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
                                            InsertTextWithMarkdown(selection, LLMResult, trailingCR, True)
                                            'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                        End If
                                        rng = selection.Range
                                    End If
                                    Dim SaveRng As Range = rng.Duplicate
                                    CompareAndInsert(SelectedText, LLMResult, rng.Duplicate, MarkupMethod = 3, "This is the markup of the text inserted:")
                                    If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then
                                        ApplyParagraphFormat(rng)
                                    End If
                                    If Not NoFormatAndFieldSaving Then RestoreSpecialTextElements(SaveRng)
                                    SaveRng.Document.Fields.Update()
                                Else
                                    If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                    CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                                    If Not NoFormatAndFieldSaving Then RestoreSpecialTextElements(rng)
                                End If
                            Else
                                Dim pattern As String = "\{\{.*?\}\}"
                                If System.Text.RegularExpressions.Regex.IsMatch(LLMResult, pattern) Then
                                    SLib.InsertTextWithBoldMarkers(selection, LLMResult & vbCrLf)
                                    'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                Else
                                    InsertTextWithMarkdown(selection, LLMResult, trailingCR, True)
                                    'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                End If
                                rng = selection.Range
                                Dim SaveRng As Range = rng.Duplicate
                                If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then
                                    ApplyParagraphFormat(rng)
                                End If
                                If Not NoFormatting Then
                                    If Not NoFormatAndFieldSaving Then RestoreSpecialTextElements(SaveRng)
                                    SaveRng.Document.Fields.Update()
                                End If
                            End If

                        End If

                        ' End Extended Selection Mode
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    End If

                Else
                    If Not DoSilent Then ShowCustomMessageBox("The LLM did not return any content to process.")
                End If

                If NoSelectedText Or ChunkSize = 0 Then
                    Exit Do
                End If

                Try
                    If currentdoc.Bookmarks.Exists("NextStart") Then
                        Try
                            currentdoc.Bookmarks("NextStart").Delete()
                        Catch ex As System.Exception
                            '
                        End Try
                    End If

                    nextStartBm = currentdoc.Bookmarks.Add(
                    Name:="NextStart",
                    Range:=currentdoc.Range(Start:=selection.End, End:=selection.End))
                    nextStartBm.Range.Collapse(WdCollapseDirection.wdCollapseEnd)

                    If nextStartBm Is Nothing OrElse Not currentdoc.Bookmarks.Exists("NextStart") Then
                        Exit Do
                    End If

                Catch ex As System.Exception
                    Exit Do
                End Try

            Loop

        Catch ex As System.Exception

#If DEBUG Then
            Debug.WriteLine("Error: " & ex.Message)
            Debug.WriteLine("Stacktrace: " & ex.StackTrace)

            System.Diagnostics.Debugger.Break()
#End If
            MessageBox.Show("Error in TrueProcessSelectedText:  " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            INIloaded = False

        Finally

            Try
                ' Aufräumen aller temporären Bookmarks
                For yi As Integer = currentdoc.Bookmarks.Count To 1 Step -1
                    Dim bm As Word.Bookmark = currentdoc.Bookmarks(yi)
                    If bm.Name = "TotalEnd" OrElse bm.Name = "NextStart" Then bm.Delete()
                Next
            Catch ex As System.Exception
                '
            End Try

        End Try

        Return ""

    End Function


    Public Sub ConvertRangeFormattingToMarkdown()
        Dim app As Word.Application = Globals.ThisAddIn.Application
        Dim sel As Word.Selection = app.Selection
        If sel Is Nothing OrElse sel.Type = Word.WdSelectionType.wdSelectionIP Then Return

        Dim rng As Word.Range = sel.Range
        Dim originalStart As Integer = rng.Start
        If rng.Characters.Count = 0 Then Return

        Dim sb As New StringBuilder()

        ' --- state flags ---------------------------------------------------------
        Dim isBold, isItalic, isUnderline, isSuper, isSub As Boolean

        For i As Integer = 1 To rng.Characters.Count
            Dim charRng As Word.Range = rng.Characters(i)
            Dim ch As String = charRng.Text
            Dim isEOL As Boolean = (ch = vbCr)

            'current states -------------------------------------------------------
            Dim nowBold As Boolean = (charRng.Font.Bold = -1)
            Dim nowItalic As Boolean = (charRng.Font.Italic = -1)
            Dim nowUnderline As Boolean = (charRng.Font.Underline <> Word.WdUnderline.wdUnderlineNone)
            'Dim nowSuper As Boolean = (charRng.Font.Superscript = -1)
            Dim nowSub As Boolean = (charRng.Font.Subscript = -1)

            '--- close tags where state turns OFF ---------------------------------
            If isSub AndAlso (Not nowSub OrElse isEOL) Then sb.Append(")")        'subscript
            'If isSuper AndAlso (Not nowSuper OrElse isEOL) Then sb.Append(")")    'superscript
            If isUnderline AndAlso (Not nowUnderline OrElse isEOL) Then sb.Append("</u>")
            If isItalic AndAlso (Not nowItalic OrElse isEOL) Then sb.Append("*")
            If isBold AndAlso (Not nowBold OrElse isEOL) Then sb.Append("**")

            '--- open tags where state turns ON -----------------------------------
            If Not isBold AndAlso nowBold AndAlso Not isEOL Then sb.Append("**")
            If Not isItalic AndAlso nowItalic AndAlso Not isEOL Then sb.Append("*")
            If Not isUnderline AndAlso nowUnderline AndAlso Not isEOL Then sb.Append("<u>")
            'If Not isSuper AndAlso nowSuper AndAlso Not isEOL Then sb.Append("^(")
            If Not isSub AndAlso nowSub AndAlso Not isEOL Then sb.Append("~(")

            '--- write the character itself ---------------------------------------
            sb.Append(ch)

            '--- update state ------------------------------------------------------
            If isEOL Then
                isBold = isItalic = isUnderline = isSuper = isSub = False
            Else
                isBold = nowBold : isItalic = nowItalic : isUnderline = nowUnderline
                'isSuper = nowSuper
                isSub = nowSub
            End If
        Next

        '--- close any still-open tags --------------------------------------------
        If isSub Then sb.Append(")")
        'If isSuper Then sb.Append(")")
        If isUnderline Then sb.Append("</u>")
        If isItalic Then sb.Append("*")
        If isBold Then sb.Append("**")

        '--- replace text ---------------------------------------------------------
        rng.Text = sb.ToString()

        '--- *only* neutralise caps flags so case is preserved --------------------
        rng.Font.AllCaps = 0        'systematically turn OFF ALL-CAPS
        rng.Font.SmallCaps = 0      'turn OFF small-caps; leave rest untouched

        're-select the new range
        rng.Select()
    End Sub


    Public Sub oldConvertRangeFormattingToMarkdown()
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







    Public Sub ApplyParagraphFormat(ByRef rng As Word.Range)

        Dim maxParaStylesCount As Integer = paragraphFormat.Length
        Dim paraCount As Integer = rng.Paragraphs.Count

        If paraCount = 0 Then Exit Sub

        For i As Integer = 1 To paraCount
            If i - 1 >= maxParaStylesCount Then Exit For

            Dim pf As ParagraphFormatStructure = paragraphFormat(i - 1)
            Dim pRange As Word.Range = rng.Paragraphs(i).Range

            '--- 1. paragraph style ------------------------------------------------
            If pf.Style IsNot Nothing Then
                Try
                    pRange.Style = pf.Style
                Catch ex As System.Exception
                    ' handle / log if necessary
                End Try
            End If

            '--- 2. character-level attributes – use them *only when supplied* -----
            With pRange.Font
                If Not String.IsNullOrEmpty(pf.FontName) Then .Name = pf.FontName
                If pf.FontSize.HasValue Then .Size = pf.FontSize.Value
                If pf.FontBold.HasValue Then .Bold = pf.FontBold.Value
                If pf.FontItalic.HasValue Then .Italic = pf.FontItalic.Value
                If pf.FontUnderline.HasValue Then .Underline = pf.FontUnderline.Value
                If pf.FontColor.HasValue Then .Color = pf.FontColor.Value
            End With                           ' everything that stays Nothing is left untouched

            '--- 3. list formatting -----------------------------------------------
            If pf.HasListFormat AndAlso pf.ListTemplate IsNot Nothing Then
                Try
                    If pRange.ListFormat.ListType <> Word.WdListType.wdListNoNumbering Then
                        pRange.ListFormat.RemoveNumbers()
                    End If

                    pRange.ListFormat.ApplyListTemplateWithLevel(
                        ListTemplate:=pf.ListTemplate,
                        ContinuePreviousList:=pf.ListLevel > 0,
                        ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList,
                        DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                    pRange.ListFormat.ListLevelNumber = pf.ListLevel
                Catch ex As System.Exception
                    ' handle / log if necessary
                End Try
            End If

            '--- 4. paragraph-level attributes ------------------------------------
            With pRange.ParagraphFormat
                .Alignment = pf.Alignment
                .LineSpacing = pf.LineSpacing
                .SpaceBefore = pf.SpaceBefore
                .SpaceAfter = pf.SpaceAfter
            End With
        Next
    End Sub


    Public Sub OldApplyParagraphFormat(ByRef rng As Range)

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


    Public Function GetVisibleText(ByVal src As Range) As String
        Try
            ' 1) Null-/Leerauswahl abfangen
            If src Is Nothing Then
                Return String.Empty
            End If

            ' 2) Rohtext einmal holen für Fast-Path
            Dim raw As String
            Try
                raw = src.Text
                If String.IsNullOrEmpty(raw) Then
                    Return String.Empty
                End If
            Catch
                Return String.Empty
            End Try

            ' 3) Fast-Path: keine Revisionen in dieser Range → sofort zurückgeben
            Dim revCount As Integer
            Try
                revCount = src.Revisions.Count
                If revCount = 0 Then
                    Return raw
                End If
            Catch
                Return raw ' If we can't access revisions count, return raw text
            End Try

            ' Alternative approach: use Word's built-in view settings
            Try
                Dim doc As Microsoft.Office.Interop.Word.Document = src.Document
                Dim origShowRevs As Boolean = doc.ShowRevisions
                Dim origPrintRevs As Boolean = doc.PrintRevisions

                ' Temporarily hide revisions to get clean text
                doc.ShowRevisions = False
                doc.PrintRevisions = False

                ' Get text without revisions
                Dim visibleText As String = src.Text

                ' Restore original settings
                doc.ShowRevisions = origShowRevs
                doc.PrintRevisions = origPrintRevs

                Return visibleText
            Catch ex As Exception
                Debug.WriteLine($"Alternative method failed: {ex.Message}")
                ' Continue with original algorithm if alternative fails
            End Try

            ' Original algorithm with better error handling
            Dim sliceStart As Integer = src.Start
            Dim sliceEnd As Integer = src.End    ' exklusiv

            ' 4) Collect deleted intervals with safer revision handling
            Dim skips As New List(Of (s As Integer, e As Integer))()
            For Each rev As Revision In src.Revisions
                Try
                    ' Skip if we can't safely get revision type
                    Dim revType As WdRevisionType = rev.Type

                    If revType = WdRevisionType.wdRevisionInsert _
                    OrElse revType = WdRevisionType.wdRevisionMovedTo Then
                        Continue For
                    End If

                    Dim revRange As Range = rev.Range
                    Dim fromPos As Integer = system.math.max(revRange.Start, sliceStart)
                    Dim toPos As Integer = system.math.min(revRange.End, sliceEnd)
                    If fromPos < toPos Then
                        skips.Add((fromPos, toPos))
                    End If
                Catch ex As Exception
                    ' Skip problematic revisions
                    Debug.WriteLine($"Error processing revision: {ex.Message}")
                    Continue For
                End Try
            Next

            ' 5) Merge intervals
            Dim merged As List(Of (s As Integer, e As Integer)) = MergeIntervals(skips)

            ' 6) Determine visible segments
            Dim keep As New List(Of (s As Integer, e As Integer))()
            Dim pos As Integer = sliceStart
            For Each iv In merged
                If iv.s > pos Then
                    keep.Add((pos, iv.s))
                End If
                pos = system.math.max(pos, iv.e)
            Next
            If pos < sliceEnd Then
                keep.Add((pos, sliceEnd))
            End If

            ' 7) Read text segment by segment
            Dim sb As New StringBuilder()
            For Each iv In keep
                Try
                    sb.Append(src.Document.Range(iv.s, iv.e).Text)
                Catch ex As Exception
                    Debug.WriteLine($"Error reading segment {iv.s}-{iv.e}: {ex.Message}")
                End Try
            Next

            Return sb.ToString()

        Catch ex As Exception
            Debug.WriteLine($"Exception in GetVisibleText: {ex.Message}{vbCrLf}{ex.StackTrace}")
            System.Diagnostics.Debugger.Break()
            ' Fall back to raw text or empty string in worst case
            Try
                Return If(src IsNot Nothing, src.Text, String.Empty)
            Catch
                Return String.Empty
            End Try
        End Try

    End Function



    Private Function MergeIntervals(ByVal intervals As List(Of (s As Integer, e As Integer))) _
    As List(Of (s As Integer, e As Integer))

        Dim result As New List(Of (s As Integer, e As Integer))()
        If intervals.Count = 0 Then
            Return result
        End If

        intervals.Sort(Function(a, b) a.s.CompareTo(b.s))
        Dim cur = intervals(0)

        For i As Integer = 1 To intervals.Count - 1
            Dim nxt = intervals(i)
            If nxt.s <= cur.e Then
                cur.e = system.math.max(cur.e, nxt.e)
            Else
                result.Add(cur)
                cur = nxt
            End If
        Next

        result.Add(cur)
        Return result
    End Function




    Public Function oldGetVisibleText(ByVal src As Microsoft.Office.Interop.Word.Range) As String
        ' Gracefully handle no selection or null input
        If src Is Nothing Then
            Return String.Empty
        End If

        Dim raw As String = src.Text
        If String.IsNullOrEmpty(raw) Then
            Return String.Empty
        End If

        Dim sliceStart As Integer = src.Start
        Dim sliceEnd As Integer = src.End         ' exclusive
        Dim rawLen As Integer = raw.Length

        ' Phase 1: collect intervals of *deleted* text (revisions that are not insert/move-to)
        Dim skip As New List(Of (s As Integer, e As Integer))
        For Each rev As Microsoft.Office.Interop.Word.Revision In src.Document.Revisions
            ' skip revisions outside our slice
            If rev.Range.End <= sliceStart OrElse rev.Range.Start >= sliceEnd Then Continue For
            ' keep insertions and moves; everything else is invisible
            If rev.Type = Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionInsert _
                OrElse rev.Type = Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionMovedTo Then Continue For

            Dim fromPos As Integer = system.math.max(rev.Range.Start, sliceStart)
            Dim toPos As Integer = system.math.min(rev.Range.End, sliceEnd)
            skip.Add((fromPos, toPos))
        Next

        ' Merge overlapping or adjacent intervals
        If skip.Count > 0 Then
            skip.Sort(Function(a, b) a.s.CompareTo(b.s))
            Dim merged As New List(Of (s As Integer, e As Integer))
            Dim cur = skip(0)
            For i As Integer = 1 To skip.Count - 1
                If skip(i).s <= cur.e Then
                    cur.e = system.math.max(cur.e, skip(i).e)
                Else
                    merged.Add(cur)
                    cur = skip(i)
                End If
            Next
            merged.Add(cur)
            skip = merged
        End If

        ' Phase 2: build visible text buffer
        Dim sb As New StringBuilder()
        Dim relPos As Integer = 0

        For Each iv In skip
            Dim delStartRel As Integer = iv.s - sliceStart
            Dim delEndRel As Integer = iv.e - sliceStart    ' exclusive

            ' Append visible segment before this deletion
            Dim visLen As Integer = delStartRel - relPos
            If visLen > 0 Then
                sb.Append(raw, relPos, visLen)
            End If

            ' Skip over deleted segment
            relPos = system.math.max(relPos, delEndRel)
        Next

        ' Append remaining tail after last deletion
        If relPos < rawLen Then
            sb.Append(raw, relPos, rawLen - relPos)
        End If

        Return sb.ToString()
    End Function





    Public Function FindLongTextInChunks(ByVal findText As String, ByVal chunkSize As Integer, ByRef selection As Word.Selection, Optional Skipdeleted As Boolean = False) As Boolean

        Return WordSearchHelper.FindLongTextAnchoredFast(selection, findText, Skipdeleted)

    End Function

    Function xxx(ByRef selection As Word.Selection, ByVal findText As String, Optional INI_Clean As Boolean = False) As Boolean
        ' This function searches for a long text in chunks of a specified sizeed as needed) 

        Dim chunkSize As Integer = 1000 ' Define the chunk size for breaking the text into smaller parts
        ' Store original selection to restore if needed
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim originalStart As Integer = selection.Start
        Dim originalEnd As Integer = selection.End

        ' Break the long text into chunks of up to chunkSize characters
        Dim chunks As New List(Of String)
        Dim startIndex As Integer = 0
        While startIndex < findText.Length
            Dim length As Integer = system.math.min(chunkSize, findText.Length - startIndex)
            chunks.Add(findText.Substring(startIndex, length))
            startIndex += length
        End While

        ' We'll need to track the final Start/End of the matched text
        Dim overallMatchStart As Integer = -1
        Dim overallMatchEnd As Integer = -1

        ' Move the selection to the beginning of the document (or keep at original if you prefer)

        For i As Integer = 0 To chunks.Count - 1
            Dim currentChunk As String = chunks(i)

            Dim chunk As String = NormalizeTextForSearch(chunks(i), INI_Clean)

            With selection.Find
                .Forward = True
                .Wrap = Word.WdFindWrap.wdFindStop
                .Format = False
                .Text = chunk ' Set the text to search for
                If INI_Clean Then
                    .MatchWildcards = True
                    ' replace every literal space with [ ]@ (one-or-more spaces)
                    ' .Text = chunk.Replace(" ", "[ ]@")
                Else
                    .MatchWildcards = False
                    '.Text = chunk
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

    ''' <summary>
    ''' Erzeugt ein Such­muster, das variable Leerzeichen, Absatz-
    ''' und Zeilen­umbrüche abfängt.
    ''' </summary>
    Public Shared Function NormalizeTextForSearch(txt As String, allowMultiSpaces As Boolean) As String

        Dim pattern As String = txt
        Dim whiteToken As String = ""

        ' Token bestimmen: nur Spaces → "[ ]@",  sonst alles → "([ ]|^13|^l)@"
        If allowMultiSpaces Then
            whiteToken = "[ ]@"
            pattern = pattern.Replace(" ", whiteToken)
        End If

        Return pattern
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
            Dim selection As Microsoft.Office.Interop.Word.Selection = app.Selection

            If selection Is Nothing OrElse selection.Range Is Nothing Then
                MessageBox.Show("Error in MarkupSelectedTextWithRegex: No text selected (anymore). Can't proceed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            'Dim splash As New Slib.Splashscreen("Applying changes... press 'Esc' to abort") 
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

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                        Exited = True
                        Exit For
                    End If

                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
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
        Dim doc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try
            ' 1) Arbeitsbereich festlegen
            Dim workRange As Microsoft.Office.Interop.Word.Range
            If OnlySelection AndAlso doc.Application.Selection IsNot Nothing _
           AndAlso doc.Application.Selection.Range.Text <> "" Then
                workRange = doc.Application.Selection.Range.Duplicate
            Else
                workRange = doc.Content.Duplicate
                OnlySelection = False
            End If

            ' 2) Marker in neuen Text einfügen
            Dim newTextWithMarker As String
            If newText.Length > 2 AndAlso Marker <> "" Then
                newTextWithMarker = newText.Substring(0, newText.Length - 2) & Marker & newText.Substring(newText.Length - 2)
            Else
                newTextWithMarker = newText
            End If

            ' 3) Ursprüngliche Selektion merken
            Dim selectionStart As Integer = doc.Application.Selection.Start
            Dim selectionEnd As Integer = doc.Application.Selection.End

            ' 4) Chunk‑ oder Standard‑Suche
            If oldText.Length > SearchChunkSize Then
                ' --- Long‑Chunk‑Suche ----------------------------------------
                doc.Application.Selection.SetRange(workRange.Start, workRange.End)
                Dim foundAny As Boolean = False

                Do While Globals.ThisAddIn.FindLongTextInChunks(oldText, SearchChunkSize, doc.Application.Selection)
                    If doc.Application.Selection Is Nothing Then Exit Do
                    ' Escape‑Taste prüfen
                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then Exit Do

                    foundAny = True
                    Dim selRange As Microsoft.Office.Interop.Word.Range = doc.Application.Selection.Range

                    ' Auf Löschungen prüfen
                    Dim isDeleted As Boolean = False
                    For Each rev As Microsoft.Office.Interop.Word.Revision In selRange.Revisions
                        If rev.Type = Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionDelete Then
                            isDeleted = True
                            Exit For
                        End If
                    Next

                    ' Ersetzen
                    Dim replaceStart As Integer = selRange.Start
                    Dim replaceEnd As Integer = selRange.End
                    If Not isDeleted Then
                        selRange.Text = newTextWithMarker
                        replaceEnd += newTextWithMarker.Length
                        selectionEnd += newTextWithMarker.Length
                    End If

                    ' Clamping und Range‑Vorschub
                    Dim newStart As Integer = System.Math.Max(0, System.Math.Min(replaceEnd, doc.Content.End))
                    Dim newEnd As Integer = If(OnlySelection,
                                            System.Math.Min(selectionEnd, doc.Content.End),
                                            doc.Content.End)
                    If newStart <= newEnd Then
                        doc.Application.Selection.SetRange(newStart, newEnd)
                    Else
                        Exit Do
                    End If
                Loop

                ' Selektion nur bei Treffern wiederherstellen
                If foundAny Then
                    selectionStart = System.Math.Max(0, System.Math.Min(selectionStart, doc.Content.End))
                    selectionEnd = System.Math.Max(0, System.Math.Min(selectionEnd, doc.Content.End))
                    If selectionStart <= selectionEnd Then
                        doc.Application.Selection.SetRange(selectionStart, selectionEnd)
                        doc.Application.Selection.Select()
                    End If
                Else
                    Debug.WriteLine("Hinweis: Begriff nicht gefunden, Restore übersprungen.")
                End If

            Else
                ' --- Standard‑Find -------------------------------------------
                If String.IsNullOrEmpty(oldText) Then
                    Debug.WriteLine("Hinweis: Suchbegriff leer.")
                Else
                    Dim replacementsMade As Boolean = False
                    Dim initialRangeEnd As Integer = workRange.End

                    Do
                        If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then Exit Do
                        oldText = NormalizeTextForSearch(oldText, ThisAddIn.INI_Clean)

                        With workRange.Find
                            .ClearFormatting()
                            .Text = oldText
                            .MatchWildcards = ThisAddIn.INI_Clean
                            .Forward = True
                            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                            .MatchWholeWord = True

                            If .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone) Then
                                Dim foundRange As Microsoft.Office.Interop.Word.Range = workRange.Duplicate

                                ' Auf Löschungen prüfen
                                Dim isDeleted As Boolean = False
                                For Each rev As Microsoft.Office.Interop.Word.Revision In foundRange.Revisions
                                    If rev.Type = Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionDelete Then
                                        isDeleted = True
                                        Exit For
                                    End If
                                Next

                                Dim prevStart As Integer = workRange.Start
                                If Not isDeleted Then
                                    foundRange.Text = newTextWithMarker
                                    replacementsMade = True
                                    initialRangeEnd += (newTextWithMarker.Length - oldText.Length)
                                End If

                                ' Arbeitsbereich vorschieben
                                workRange.Start = foundRange.End
                                If workRange.Start <= prevStart Then
                                    workRange.Start = prevStart + 1
                                End If
                                workRange.End = System.Math.Min(initialRangeEnd, doc.Content.End)
                            Else
                                Exit Do
                            End If
                        End With
                    Loop

                    If Not replacementsMade Then
                        Debug.WriteLine("Hinweis: Suchbegriff nicht gefunden.")
                    End If
                End If
            End If

        Catch ex As System.Exception
            MsgBox("Error in SearchReplace: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub



    Private Sub oldSearchAndReplace(oldText As String, newText As String, OnlySelection As Boolean, Marker As String)

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


            If Len(oldText) > SearchChunkSize Then

                Dim selectionStart As Integer = doc.Application.Selection.Start
                Dim selectionEnd As Integer = doc.Application.Selection.End
                doc.Application.Selection.SetRange(workRange.Start, workRange.End)
                Dim found As Boolean = False

                ' Loop through the content to find and replace all instances
                Do While Globals.ThisAddIn.FindLongTextInChunks(oldText, SearchChunkSize, doc.Application.Selection) = True

                    If doc.Application.Selection Is Nothing Then Exit Do

                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
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

                        If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                            Exit Do
                        End If

                        oldText = NormalizeTextForSearch(oldText, ThisAddIn.INI_Clean)

                        With workRange.Find
                            .ClearFormatting()
                            .Text = oldText
                            If ThisAddIn.INI_Clean Then
                                .MatchWildcards = True
                                ' turn each " " into "[ ]@" so Word will match 1+ spaces
                                '.Text = oldText.Replace(" ", "[ ]@")
                            Else
                                .MatchWildcards = False
                                '.Text = oldText
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



    Private Shared emojiSet As New HashSet(Of String)()
    Private Shared ReadOnly _emojiPairRegex As New System.Text.RegularExpressions.Regex(
    "[\uD83C-\uDBFF][\uDC00-\uDFFF]",
    System.Text.RegularExpressions.RegexOptions.Compiled Or
    System.Text.RegularExpressions.RegexOptions.CultureInvariant)


    Public Shared Sub InsertTextWithMarkdown(selection As Microsoft.Office.Interop.Word.Selection, Result As String, Optional TrailingCR As Boolean = False, Optional AddTrailingIfNeeded As Boolean = False)

        If selection Is Nothing Then
            MessageBox.Show("Error in InsertTextWithMarkdown: The selection object is null", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Extract the range from the selection
        Dim range As Microsoft.Office.Interop.Word.Range = selection.Range

        Dim LeadingTrailingSpace As Boolean = False

        Debug.WriteLine($"IM1-Range Start = {selection.Start}")
        Debug.WriteLine($"Range End = {selection.End}")
        Debug.WriteLine("TrailingCR = " & TrailingCR)
        Debug.WriteLine(selection.Text)

        If range.Start < range.End AndAlso Not TrailingCR Then

            ' Prüfen, ob vor und hinter range Platz im Dokument ist; erforderlich, weil beim Löschen eines solchen Texts Word automatisch einen Space entfernt
            Dim docStart As Integer = range.Document.Content.Start
            Dim docEnd As Integer = range.Document.Content.End

            If range.Start > docStart AndAlso range.End < docEnd Then
                ' Ein 1‐Zeichen‐Range vor range
                Dim beforerange As Range = range.Document.Range(range.Start - 1, range.Start)

                ' Ein 1‐Zeichen‐Range nach range
                'Dim afterrange As Range = range.Document.Range(range.End - 1, range.End + 1)
                Dim afterrange As Range = range.Document.Range(range.End - 1, range.End)

                'If beforerange.Text = " " AndAlso afterrange.Text = " " Then
                Debug.WriteLine($"Beforetext='{beforerange.Text}'")
                Debug.WriteLine($"Aftertext='{afterrange.Text}'")
                'If afterrange.Text.EndsWith(" "c) OrElse afterrange.Text.StartsWith(" "c) Then
                If afterrange.Text.EndsWith(" "c) OrElse afterrange.Text.StartsWith(" "c) Then
                    LeadingTrailingSpace = True
                Else
                    LeadingTrailingSpace = False
                End If
            Else
                LeadingTrailingSpace = False
            End If
        End If


        'If range.Start < range.End Then
        'If TrailingCR Then
        'range.End = range.End - 1
        'End If

        'range.Delete()

        'End If

        Dim insertionStart As Integer = selection.Range.Start

        'Debug.WriteLine($"IM2-Range Start = {selection.Start}")
        'Debug.WriteLine($"Range End = {selection.End}")
        'Debug.WriteLine("TrailingCR = " & TrailingCR)
        'Debug.WriteLine(selection.Text)

        Dim ResultBack As String = Result
        Try
            Result = System.Text.RegularExpressions.Regex.Unescape(Result)
        Catch
            Debug.WriteLine("Error unescaping Result with: " & Result)
            Result = ResultBack
        End Try

        Dim markdownSource As String = Result

        'emojiSet = New HashSet(Of String)()

        For i As Integer = 0 To Result.Length - 1
            Debug.WriteLine($"Char: '{Result(i)}'  ASCII: {Asc(Result(i))}")
        Next

        Result = Result.Replace(vbLf & " " & vbLf, vbLf & vbLf)

        Dim pattern As String = "((\r\n|\n|\r){2,})"
        Result = Regex.Replace(Result, pattern, Function(m As Match)
                                                    ' Prüfen, ob das Match bis zum Ende des Strings reicht:
                                                    If m.Index + m.Length = Result.Length Then
                                                        ' Am Ende: Rückgabe der Umbrüche wie sie sind
                                                        Return m.Value
                                                    Else
                                                        ' Andernfalls: &nbsp; zwischen die Umbrüche einfügen
                                                        Dim breaks As String = m.Value
                                                        Dim regexBreaks As New Regex("(\r\n|\n|\r)")
                                                        Dim splitBreaks = regexBreaks.Matches(breaks)
                                                        If splitBreaks.Count <= 1 Then Return breaks
                                                        Dim resultx As String = splitBreaks(0).Value
                                                        For i As Integer = 1 To splitBreaks.Count - 1
                                                            resultx &= vbCrLf & "&nbsp;" & vbCrLf & splitBreaks(i).Value
                                                        Next
                                                        Return resultx
                                                    End If
                                                End Function)


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

        Debug.WriteLine("Result=" & Result)

        Dim htmlResult As String = Markdown.ToHtml(Result, markdownPipeline).Trim


        ' ─── alle echten Newlines raus, damit sie nicht als Text umgewandelt werden ───
        htmlResult = htmlResult _
                .Replace(vbCrLf, "") _
                .Replace(vbCr, "") _
                .Replace(vbLf, "")

        'emojiSet = New HashSet(Of String)(
        'System.Text.RegularExpressions.Regex _
        '.Matches(htmlResult, "[\uD83C-\uDBFF][\uDC00-\uDFFF]") _
        '.Cast(Of System.Text.RegularExpressions.Match)() _
        '.Select(Function(m) m.Value)
        '       )

        ' Load the HTML into HtmlDocument
        Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
        Dim fullhtml As String
        htmlDoc.LoadHtml(htmlResult)

        fullHtml = htmlDoc.DocumentNode.OuterHtml
        Debug.WriteLine("HTML1=" & fullhtml)

        'RemoveTrailingParagraph(htmlDoc)

        'fullhtml = htmlDoc.DocumentNode.OuterHtml
        'Debug.WriteLine("HTML3=" & fullhtml)

        SLib.InsertTextWithFormat(fullhtml, range, True, Not TrailingCR)

        ' Nach dem Paste und Sleep …
        range = range.Application.Selection.Range

        ' Wenn das letzte Zeichen ein Absatzmarke (vbCr) ist, lösche es
        'If range.Characters.Last.Text = vbCr Then
        'range.Characters.Last.Delete()
        'End If

        'ParseHtmlNode(htmlDoc.DocumentNode, range)

        'emojiSet = Nothing

        Debug.WriteLine($"IM3-Range Start = {selection.Start}")
        Debug.WriteLine($"Range End = {selection.End}")
        Debug.WriteLine("LeadingTrailingSpace = " & LeadingTrailingSpace)
        Debug.WriteLine("TrailingCR = " & TrailingCR)
        Debug.WriteLine(selection.Text)

        If LeadingTrailingSpace Then
            range.Collapse(WdCollapseDirection.wdCollapseEnd)
            range.InsertAfter(" ")
        End If

        'If TrailingCR AndAlso AddTrailingIfNeeded Then
        'range.Collapse(WdCollapseDirection.wdCollapseEnd)
        'range.InsertParagraphAfter()
        'range.Collapse(WdCollapseDirection.wdCollapseEnd)
        'End If

        Dim InsertionEnd As Integer = range.End

        Dim doc As Microsoft.Office.Interop.Word.Document = selection.Document
        selection.SetRange(insertionStart, InsertionEnd)
        selection.Select()

        Debug.WriteLine($"IM4-Range Start = {selection.Start}")
        Debug.WriteLine($"Range End = {selection.End}")
        Debug.WriteLine("LeadingTrailingSpace = " & LeadingTrailingSpace)
        Debug.WriteLine("TrailingCR = " & TrailingCR)
        Debug.WriteLine(selection.Text)

    End Sub



    ''' <summary>
    ''' Verknüpft zwei Formatierungs‑Delegates, ohne die ursprünglichen zu verlieren.
    ''' </summary>
    Private Shared Function CombineStyle(
    baseAction As Action(Of Microsoft.Office.Interop.Word.Range),
    additional As Action(Of Microsoft.Office.Interop.Word.Range)
) As Action(Of Microsoft.Office.Interop.Word.Range)

        If baseAction Is Nothing Then Return additional
        If additional Is Nothing Then Return baseAction

        Return Sub(rng As Microsoft.Office.Interop.Word.Range)
                   baseAction(rng)
                   additional(rng)
               End Sub
    End Function


    ''' <summary>
    ''' Rendert (ggf. rekursiv) reine Inline‑Knoten mit kumulativer Formatierung.
    ''' </summary>
    Private Shared Sub RenderInline(
    node As HtmlAgilityPack.HtmlNode,
    rng As Microsoft.Office.Interop.Word.Range,
    styleAction As Action(Of Microsoft.Office.Interop.Word.Range),
    inheritedHref As String
)

        ' Kommentare ignorieren
        If node.NodeType = HtmlAgilityPack.HtmlNodeType.Comment Then Return

        ' ------------------------------------------------- Leaf: #text -------------
        'If node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
        'Dim txt As String = HtmlEntity.DeEntitize(node.InnerText)
        'If Not String.IsNullOrWhiteSpace(txt) Then
        'InsertInline(rng, txt, styleAction, inheritedHref)
        'End If
        'Return
        'End If

        ' ─── Newline‑Handling: Nur echte, mit Inhalt versehene Zeilenumbrüche splitten
        If node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
            Dim rawText = HtmlEntity.DeEntitize(node.InnerText)
            Dim hasNewline = rawText.IndexOfAny({vbCr(0), vbLf(0)}) >= 0
            Dim stripped = rawText.Replace(vbCr, "").Replace(vbLf, "")

            ' 1) reine Whitespace‑-only‑Zeilenumbrüche → komplett überspringen
            If hasNewline AndAlso String.IsNullOrWhiteSpace(stripped) Then
                Return
            End If

            ' 2) gemischter Text mit echten Newlines → splitten & umbruchsweise einfügen
            If hasNewline Then
                Dim parts = rawText.Split(
            New String() {vbCrLf, vbCr, vbLf},
            StringSplitOptions.None)
                For i = 0 To parts.Length - 1
                    Dim segment = parts(i)
                    If Not String.IsNullOrWhiteSpace(segment) Then
                        InsertInline(rng, segment, styleAction, inheritedHref)
                    End If
                    If i < parts.Length - 1 Then
                        rng.InsertAfter(vbCr)
                        rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    End If
                Next
                Return
            End If

            ' 3) kein Newline → ganz normal einfügen
            If Not String.IsNullOrWhiteSpace(rawText) Then
                InsertInline(rng, rawText, styleAction, inheritedHref)
            End If
            Return
        End If


        ' ------------------------------------------------- Leaf: <br> --------------
        'If node.Name.Equals("br", StringComparison.OrdinalIgnoreCase) Then
        'rng.Font.Reset()
        'rng.Text = vbCr
        'rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        'Return
        'End If

        If node.Name.Equals("br", StringComparison.OrdinalIgnoreCase) Then
            ' Soft line break in Word (Shift+Enter) statt harter Absatz
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            rng.InsertAfter(ChrW(11))  ' manueller Zeilenumbruch
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            Return
        End If

        ' ------------------------------------------------- Leaf: <img> -------------
        If node.Name.Equals("img", StringComparison.OrdinalIgnoreCase) Then
            Dim src As String = node.GetAttributeValue("src", String.Empty)
            If Not String.IsNullOrWhiteSpace(src) Then
                ' Statt InlineShapes.AddPicture direkt aufzurufen,
                ' ruf den robusten Helper auf:
                InsertImageFromSrc(rng, src)
            End If
            Return
        End If

        ' ------------------------------------------------- Leaf/Semi‑Leaf: <a> -----
        Dim thisHref As String = inheritedHref
        If node.Name.Equals("a", StringComparison.OrdinalIgnoreCase) Then
            thisHref = node.GetAttributeValue("href", String.Empty)

            ' einfache Textlinks direkt ausgeben …
            If node.ChildNodes.All(Function(c) c.NodeType = HtmlAgilityPack.HtmlNodeType.Text) Then
                Dim txtLink As String = HtmlEntity.DeEntitize(node.InnerText)
                InsertInline(rng, txtLink, styleAction, thisHref)
                Return
            End If
            ' … ansonsten werden die Kinder rekursiv mit demselben href gerendert
        End If

        ' ------------------------------------------------- Style‑Weiche ------------
        Select Case node.Name.ToLowerInvariant()
            Case "strong", "b"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Bold = True)

            Case "em", "i"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Italic = True)

            Case "u"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Underline = Word.WdUnderline.wdUnderlineSingle)

            Case "del", "s"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.StrikeThrough = True)

            Case "code"
                styleAction = CombineStyle(styleAction,
                Sub(r)
                    r.Font.Name = "Courier New"
                    r.Font.Size = 10
                    r.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25
                End Sub)

            Case "sub"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Subscript = True)

            Case "sup"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Superscript = True)

            Case "span"
                Dim cls = node.GetAttributeValue("class", String.Empty)
                If cls.Contains("emoji") Then
                    styleAction = CombineStyle(styleAction,
                    Sub(r)
                        r.Font.Name = "Segoe UI Emoji"
                        r.Font.Color = Word.WdColor.wdColorWhite
                        r.Shading.BackgroundPatternColor =
                            System.Drawing.ColorTranslator.ToOle(
                                System.Drawing.Color.FromArgb(0, 112, 192))
                    End Sub)
                End If
                ' sonst keine eigene Formatierung → einfach durchreichen
        End Select

        ' ------------------------------------------------- Rekursion ---------------
        For Each child In node.ChildNodes
            RenderInline(child, rng, styleAction, thisHref)
        Next
    End Sub




    ''' <summary>
    ''' Fügt ein Bild aus src in den Word-Range ein.
    ''' Unterstützt lokale Pfade und Web-URLs, und fängt alle Fehler intern ab.
    ''' </summary>
    Private Shared Sub InsertImageFromSrc(
    rng As Microsoft.Office.Interop.Word.Range,
    src As String
)
        If String.IsNullOrWhiteSpace(src) Then Return

        Dim fileName As String = src
        Dim tempFile As String = String.Empty
        Dim isUrl As Boolean = False

        Try
            ' URL-Erkennung
            Dim uri = New System.Uri(src, UriKind.RelativeOrAbsolute)
            If uri.IsAbsoluteUri AndAlso
           (uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) _
            OrElse uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) Then

                isUrl = True
                tempFile = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                System.IO.Path.GetFileName(uri.LocalPath)
            )
                Using client As New System.Net.WebClient()
                    client.DownloadFile(uri, tempFile)
                End Using
                fileName = tempFile
            End If

            ' Existenz prüfen
            If Not System.IO.File.Exists(fileName) Then
                Throw New System.Exception(
                $"Bilddatei nicht gefunden bzw. Download fehlgeschlagen: '{fileName}'"
            )
            End If

            ' Einfügen
            Dim pic As Microsoft.Office.Interop.Word.InlineShape =
            rng.InlineShapes.AddPicture(
                FileName:=fileName,
                LinkToFile:=False,
                SaveWithDocument:=True
            )
            rng.SetRange(pic.Range.End, pic.Range.End)

        Catch ex As System.Exception
            ' Fehler intern loggen, nicht weiterwerfen
            Debug.WriteLine(
            $"[InsertImageFromSrc] {ex.GetType().FullName}: {ex.Message}"
        )
            rng.InsertAfter("[Image missing]")
        Finally
            ' Temp-Datei entfernen
            If isUrl AndAlso Not String.IsNullOrWhiteSpace(tempFile) Then
                Try
                    System.IO.File.Delete(tempFile)
                Catch ioEx As System.Exception
                    Debug.WriteLine(
                    $"[InsertImageFromSrc] Temp-Datei konnte nicht gelöscht werden: {ioEx.Message}"
                )
                End Try
            End If
        End Try
    End Sub

    Private Shared Sub RemoveTrailingParagraph(htmlDoc As HtmlAgilityPack.HtmlDocument)

        Dim candidates = New String() {"p", "br"}

        For Each TagName In candidates
            Dim lastNode = htmlDoc.DocumentNode.SelectSingleNode("(//" & TagName & ")[last()]")
            If lastNode Is Nothing Then Continue For

            ' liegt wirklich ganz am Schluss?
            Dim cur = lastNode.NextSibling
            Dim canDelete As Boolean = True
            While cur IsNot Nothing
                Select Case cur.NodeType
                    Case HtmlAgilityPack.HtmlNodeType.Comment
                    ' ignorieren
                    Case HtmlAgilityPack.HtmlNodeType.Text
                        If Not String.IsNullOrWhiteSpace(cur.InnerText) Then
                            canDelete = False : Exit While
                        End If
                    Case Else
                        canDelete = False : Exit While
                End Select
                cur = cur.NextSibling
            End While

            If canDelete Then
                If TagName = "p" Then
                    ' Kinder retten, <p> selbst raus
                    For Each child In lastNode.ChildNodes.ToList()
                        lastNode.ParentNode.InsertBefore(child, lastNode)
                    Next
                End If
                lastNode.Remove()
                Exit For            ' nur EIN Abschlusstag anfassen
            End If
        Next
    End Sub




    Private Shared ulLevels As List(Of Integer)
    Private Shared ulStartPos As Integer

    Private Shared Sub ParseHtmlNode(
    node As HtmlAgilityPack.HtmlNode,
    range As Microsoft.Office.Interop.Word.Range,
    Optional currentLevel As Integer = 0)

        ' -------------------------------- Text‑Shortcut ---------------------------
        If Not node.HasChildNodes AndAlso node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
            RenderInline(node, range, Nothing, String.Empty)
            Return
        End If


        If node.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then
            ' (1) Inline‐Inhalt rendern
            RenderInline(node, range, Nothing, String.Empty)

            ' — nächstes echtes Node finden (Whitespace/Comments überspringen) —
            Dim nxt As HtmlAgilityPack.HtmlNode = node.NextSibling
            While nxt IsNot Nothing _
            AndAlso (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                    OrElse (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                            AndAlso String.IsNullOrWhiteSpace(nxt.InnerText)))
                nxt = nxt.NextSibling
            End While

            ' — nur wenn wirklich noch etwas folgt, Absatz(e) einfügen —
            If nxt IsNot Nothing Then
                ' 1× Absat­zumbruch
                range.InsertParagraphAfter()
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' noch ein Leerabsatz, falls nächstes Geschwister ein <p> ist
                If nxt.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then
                    range.InsertParagraphAfter()
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                End If
            End If

            Return
        End If


        If node.Name.Equals("li", StringComparison.OrdinalIgnoreCase) Then

            Dim isFootnoteEntry As Boolean =
                    node.GetAttributeValue("id", String.Empty) _
                        .StartsWith("fn:", StringComparison.OrdinalIgnoreCase)

            If isFootnoteEntry Then
                ' --- (A) Bookmark‑Name und Fußnotenzahl ermitteln ---
                Dim rawId As String = node.GetAttributeValue("id", String.Empty)   ' z.B. "fn:1"
                Dim fnNum As String = rawId.Substring(rawId.IndexOf(":"c) + 1)     ' "1"
                Dim bookmarkName As String = "fn" & fnNum                          ' "fn1" (gültiger Bookmark-Name)

                ' --- (B) Superscript-Zahl einfügen und Bookmark um sie herum anlegen ---
                Dim bmStart As Integer = range.End
                InsertInline(
                                range,
                                fnNum,
                                Sub(r) r.Font.Superscript = True,
                                String.Empty
                            )
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                Dim bmRange As Word.Range = range.Document.Range(bmStart, range.End)
                range.Document.Bookmarks.Add(Name:=bookmarkName, Range:=bmRange)

                Debug.WriteLine($"[ParseHtmlNode] Footnote Bookmark '{bookmarkName}' at Range=({bmStart},{range.End})")

                ' Leerzeichen nach der Zahl
                range.InsertAfter(" ")
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' --- (C) Den gesamten Fußnoten-Text (und Rücksprung-Arrow) rendern ---
                ' Falls das <li> nur ein <p> enthält, entpacken wir es
                If node.ChildNodes.Count = 1 _
                     AndAlso node.FirstChild.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then

                    Dim pNode As HtmlAgilityPack.HtmlNode = node.FirstChild
                    For Each subNode As HtmlAgilityPack.HtmlNode In pNode.ChildNodes
                        ParseHtmlNode(subNode, range, currentLevel)
                    Next

                Else

                    For Each subNode As HtmlAgilityPack.HtmlNode In node.ChildNodes

                        Select Case subNode.Name.ToLowerInvariant()

                            Case "p"
                                ' (1) Inline-Inhalt rendern (inkl. <br> als manueller Break)
                                RenderInline(subNode, range, Nothing, String.Empty)

                                ' (2) echten Absatz abschließen
                                range.InsertParagraphAfter()
                                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                                ' (3) prüfen, ob direkt ein <p> oder <blockquote> folgt (nach Whitespace/Comments)
                                Dim nxt As HtmlAgilityPack.HtmlNode = subNode.NextSibling
                                While nxt IsNot Nothing _
                                          AndAlso (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                                                   OrElse (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                                           AndAlso String.IsNullOrWhiteSpace(nxt.InnerText)))
                                    nxt = nxt.NextSibling
                                End While

                                ' (4) wenn ja, noch eine leere Zeile einfügen
                                If nxt IsNot Nothing _
                                       AndAlso (nxt.Name.Equals("p", StringComparison.OrdinalIgnoreCase) _
                                                OrElse nxt.Name.Equals("blockquote", StringComparison.OrdinalIgnoreCase)) Then

                                    range.InsertParagraphAfter()
                                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                End If

                                Exit Select


                            Case "blockquote"
                                ' (1) Zitat‑Absätze einzeln rendern
                                Dim quoteParas = subNode.SelectNodes("./p")
                                If quoteParas IsNot Nothing Then
                                    For Each pNode As HtmlAgilityPack.HtmlNode In quoteParas

                                        Dim paraStart As Integer = range.Start
                                        RenderInline(pNode, range, Nothing, String.Empty)

                                        range.InsertParagraphAfter()
                                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                                        ' Den soeben eingefügten Absatz einrücken
                                        Dim indentRg As Microsoft.Office.Interop.Word.Range =
                                            range.Document.Range(paraStart, range.End)
                                        indentRg.ListFormat.RemoveNumbers()
                                        indentRg.ParagraphFormat.LeftIndent +=
                                            indentRg.Application.CentimetersToPoints(0.75)


                                    Next
                                Else
                                    ' Fallback: ohne <p>-Wrapper rekursiv parsen
                                    For Each innerNode As HtmlAgilityPack.HtmlNode In subNode.ChildNodes
                                        ParseHtmlNode(innerNode, range, currentLevel)
                                    Next
                                End If

                                Exit Select


                        ' Inline‑Elemente direkt rendern, damit RenderInline den Fett‑Stil anwendet:
                            Case "#text", "strong", "b", "em", "i", "u",
                             "del", "s", "sub", "sup", "code", "span", "img", "br", "a"

                                RenderInline(subNode, range, Nothing, String.Empty)

                        ' verschachtelte Listen wie gehabt überspringen
                            Case "ul", "ol"
                                ' nichts tun, wird unten separat behandelt

                                ' alle anderen Block‑Elemente rekursiv parsen
                            Case Else
                                ParseHtmlNode(subNode, range, currentLevel)

                        End Select

                    Next

                End If

                ' --- (D) Absatz nach jeder Fußnote und Rückkehr ---
                range.InsertParagraphAfter()
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Return
            End If

            If currentLevel > 1 Then
                ' --- (0) CR VOR dem LI, je nach Listen‑Typ ---
                Dim parentName = node.ParentNode.Name.ToLowerInvariant()
                If parentName = "ol" Or parentName = "ul" Then
                    ' OL: nur vor Unterpunkten außer dem ersten
                    Dim sibs = node.ParentNode.SelectNodes("li")
                    If sibs IsNot Nothing Then
                        Dim idx As Integer = 0
                        For i As Integer = 0 To sibs.Count - 1
                            If sibs(i) Is node Then
                                idx = i
                                Exit For
                            End If
                        Next
                        If idx > 0 Then
                            range.InsertAfter(vbCr)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        End If
                    End If
                End If
            End If

            ' --- (1) Level speichern ---
            If ulLevels IsNot Nothing Then
                ulLevels.Add(currentLevel)
            End If

            ' --- (2) P‑Wrapper entfernen, wenn er das einzige direkte Kind ist ---
            If node.ChildNodes.Count = 1 _
       AndAlso node.FirstChild.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then

                Dim pNode As HtmlAgilityPack.HtmlNode = node.FirstChild
                For Each subNode As HtmlAgilityPack.HtmlNode In pNode.ChildNodes
                    ParseHtmlNode(subNode, range, currentLevel)
                Next

            Else

                For Each subNode As HtmlAgilityPack.HtmlNode In node.ChildNodes

                    Select Case subNode.Name.ToLowerInvariant()

                        Case "blockquote"
                            ' (1) Zitat‑Absätze einzeln rendern
                            Dim quoteParas = subNode.SelectNodes("./p")
                            If quoteParas IsNot Nothing Then
                                For Each pNode As HtmlAgilityPack.HtmlNode In quoteParas
                                    Dim paraStart As Integer = range.Start
                                    RenderInline(pNode, range, Nothing, String.Empty)

                                    range.InsertParagraphAfter()
                                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                                    ' Den soeben eingefügten Absatz einrücken
                                    Dim indentRg As Microsoft.Office.Interop.Word.Range =
                                        range.Document.Range(paraStart, range.End)
                                    indentRg.ListFormat.RemoveNumbers()
                                    indentRg.ParagraphFormat.LeftIndent +=
                                        indentRg.Application.CentimetersToPoints(0.75)


                                Next
                            Else
                                ' Fallback: ohne <p>-Wrapper rekursiv parsen
                                For Each innerNode As HtmlAgilityPack.HtmlNode In subNode.ChildNodes
                                    ParseHtmlNode(innerNode, range, currentLevel)
                                Next
                            End If

                            Exit Select

                        ' Inline‑Elemente direkt rendern, damit RenderInline den Fett‑Stil anwendet:
                        Case "#text", "strong", "b", "em", "i", "u",
                             "del", "s", "sub", "sup", "code", "span", "img", "br", "a"

                            RenderInline(subNode, range, Nothing, String.Empty)

                        ' verschachtelte Listen wie gehabt überspringen
                        Case "ul", "ol"
                            ' nichts tun, wird unten separat behandelt

                            ' alle anderen Block‑Elemente rekursiv parsen
                        Case Else
                            ParseHtmlNode(subNode, range, currentLevel)

                    End Select

                Next

            End If

            ' --- (3) Verschachtelte Listen am Ende behandeln ---
            Dim nestedUl As HtmlAgilityPack.HtmlNode = node.SelectSingleNode("ul")
            Dim nestedOl As HtmlAgilityPack.HtmlNode = node.SelectSingleNode("ol")
            If (nestedUl IsNot Nothing OrElse nestedOl IsNot Nothing) Then
                range.InsertAfter(vbCr)
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            End If
            If nestedUl IsNot Nothing Then
                ParseHtmlNode(nestedUl, range, currentLevel + 1)
            ElseIf nestedOl IsNot Nothing Then
                ParseHtmlNode(nestedOl, range, currentLevel + 1)
            End If

            If isFootnoteEntry Then
                range.InsertParagraphAfter()
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            End If

            Return
        End If


        Debug.WriteLine($"[ParseHtmlNode] Enter node=<{node.Name}> Range=({range.Start},{range.End})")

        ' -------------------------------- Haupt, und Kindknoten‑Schleife ---------------------
        For Each childNode As HtmlAgilityPack.HtmlNode In node.ChildNodes


            Debug.WriteLine($"  └─ Child: <{childNode.Name}> Type={childNode.NodeType}")

            Dim nestedLinkNode As HtmlAgilityPack.HtmlNode = Nothing
            If Not childNode.Name.Equals("a", StringComparison.OrdinalIgnoreCase) Then
                nestedLinkNode = childNode.SelectSingleNode(".//a")
            End If
            Dim nestedHref As String = If(nestedLinkNode IsNot Nothing,
                                      nestedLinkNode.GetAttributeValue("href", String.Empty),
                                      String.Empty)

            Select Case childNode.Name.ToLowerInvariant()


                Case "blockquote"
                    ' (1) Absatz vor dem Zitat
                    'range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    'range.InsertParagraphAfter()
                    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                    ' (2) Nur direkte <p>-Kinder verarbeiten
                    Dim quoteParas As HtmlAgilityPack.HtmlNodeCollection =
                        childNode.SelectNodes("./p")

                    If quoteParas IsNot Nothing Then
                        For Each pNode As HtmlAgilityPack.HtmlNode In quoteParas
                            ' Markiere den Anfang des neuen Absatzes
                            Dim paraStart As Integer = range.Start

                            ' (3) Inline-Inhalt des Zitats rendern
                            RenderInline(pNode, range, Nothing, String.Empty)

                            ' (4) Absatz nach jedem Zitat-Absatz
                            range.InsertParagraphAfter()
                            range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                            ' (5) Den soeben eingefügten Absatz einrücken
                            Dim indentRg As Microsoft.Office.Interop.Word.Range =
                                range.Document.Range(paraStart, range.End)
                            indentRg.ListFormat.RemoveNumbers()
                            indentRg.ParagraphFormat.LeftIndent +=
                                indentRg.Application.CentimetersToPoints(0.75)

                        Next
                    Else
                        ' Fallback: kein <p> im Blockquote → normal parsen
                        For Each subNode As HtmlAgilityPack.HtmlNode In childNode.ChildNodes
                            ParseHtmlNode(subNode, range, currentLevel)
                        Next
                    End If

                    Exit Select

                    ' (6) Absatz nach dem gesamten Blockquote
                    range.InsertParagraphAfter()
                    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                    ' (7) zusätzlichen Leerabsatz nur, wenn nach </blockquote> direkt ein <p> folgt
                    Dim nxtBQ As HtmlAgilityPack.HtmlNode = childNode.NextSibling
                    While nxtBQ IsNot Nothing _
                              AndAlso (nxtBQ.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                                       OrElse (nxtBQ.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                               AndAlso String.IsNullOrWhiteSpace(nxtBQ.InnerText)))
                        nxtBQ = nxtBQ.NextSibling
                    End While

                    If nxtBQ IsNot Nothing _
                           AndAlso nxtBQ.Name.Equals("p", System.StringComparison.OrdinalIgnoreCase) Then

                        range.InsertParagraphAfter()
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    End If

                    Exit Select

                Case "p"
                    ' Inline-Inhalt rendern (inkl. <br> als manueller Umbruch)
                    RenderInline(childNode, range, Nothing, String.Empty)

                    ' nächstes echtes Geschwister-Node finden
                    Dim nxt As HtmlAgilityPack.HtmlNode = childNode.NextSibling
                    While nxt IsNot Nothing _
                          AndAlso (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                                   OrElse (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                           AndAlso String.IsNullOrWhiteSpace(nxt.InnerText)))
                        nxt = nxt.NextSibling
                    End While

                    ' nur wenn wirklich etwas folgt, 1× Absatz
                    If nxt IsNot Nothing Then
                        range.InsertParagraphAfter()
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        ' zusätzlich eine Leerzeile, wenn das Geschwister ein weiterer <p> ist
                        If nxt.Name.Equals("p", StringComparison.OrdinalIgnoreCase) OrElse nxt.Name.Equals("blockquote", StringComparison.OrdinalIgnoreCase) Then
                            range.InsertParagraphAfter()
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        End If
                    End If

                    Exit Select



                Case "#text", "strong", "b", "em", "i", "u", "del", "s",
                    "sub", "sup", "code", "span", "img", "br"
                    RenderInline(childNode, range, Nothing, String.Empty)

                Case "div"
                    Dim cls As String = childNode.GetAttributeValue("class", String.Empty)
                    If cls.Equals("footnotes", StringComparison.OrdinalIgnoreCase) Then
                        ' Statt zu überspringen: das OL innerhalb der Fußnoten parsen
                        Dim footOl As HtmlAgilityPack.HtmlNode = childNode.SelectSingleNode("ol")
                        If footOl IsNot Nothing Then
                            ' currentLevel evtl. gleich lassen oder auf 0 setzen,
                            ' je nachdem, wie du die Nummerierung haben willst
                            ParseHtmlNode(footOl, range, currentLevel)
                        End If
                        Exit Select
                    End If

                Case "br"
                    RenderInline(childNode, range, Nothing, String.Empty)

                Case "h1", "h2", "h3", "h4", "h5", "h6"
                    ' 1) Welcher Built-In Heading-Style?

                    Dim style As WdBuiltinStyle = WdBuiltinStyle.wdStyleNormal ' Default for 'p'
                    Select Case childNode.Name.ToLower()
                        Case "h1" : style = WdBuiltinStyle.wdStyleHeading1
                        Case "h2" : style = WdBuiltinStyle.wdStyleHeading2
                        Case "h3" : style = WdBuiltinStyle.wdStyleHeading3
                        Case "h4" : style = WdBuiltinStyle.wdStyleHeading4
                        Case "h5" : style = WdBuiltinStyle.wdStyleHeading5
                        Case "h6" : style = WdBuiltinStyle.wdStyleHeading6
                    End Select

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
                    Dim cls As String = childNode.GetAttributeValue("class", String.Empty)
                    Dim href As String = childNode.GetAttributeValue("href", String.Empty)
                    Dim id As String = childNode.GetAttributeValue("id", String.Empty)

                    ' 1) Inline-Fußnoten-Referenz im Text?
                    If id.StartsWith("fnref:", StringComparison.OrdinalIgnoreCase) _
                       AndAlso href.StartsWith("#fn:", StringComparison.OrdinalIgnoreCase) Then

                        Debug.WriteLine("Setting Bookmark in Case 'a'")

                        ' Fußnotenzahl extrahieren
                        Dim fnNum As String = id.Substring(id.IndexOf(":"c) + 1)  ' z.B. "1"
                        Dim bookmarkName As String = "fn" & fnNum                         ' "fn1"

                        ' Anzeige-Text (die <sup>1</sup>)
                        Dim displayText As String = HtmlEntity.DeEntitize(childNode.InnerText)

                        'Hyperlink auf unser Bookmark anlegen
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        range.Document.Hyperlinks.Add(
                                                Anchor:=range,
                                                Address:="",                ' keine externe URL
                                                SubAddress:=bookmarkName,   ' internes Ziel
                                                TextToDisplay:=displayText
                                                )
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        Exit Select    ' <-- hier beenden wir NUR den Case, nicht die ganze Sub
                    End If

                    ' 2) Rücksprung-Pfeil in der Fußnotenliste
                    If cls.Equals("footnote-back-ref", StringComparison.OrdinalIgnoreCase) Then
                        RenderInline(childNode, range, Nothing, String.Empty)
                        Exit Select
                    End If

                    ' 3) alle anderen Links normal
                    RenderInline(childNode, range, Nothing, String.Empty)
                    Exit Select




                Case "ul"
                    ' a) Task‑List (Checkboxen) wie gehabt …
                    If childNode.GetAttributeValue("class", "").Contains("contains-task-list") Then
                        For Each li As HtmlAgilityPack.HtmlNode In childNode.SelectNodes("li")
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Dim chk = li.SelectSingleNode(".//input[@type='checkbox']")
                            Dim symbol = If(chk IsNot Nothing _
                           AndAlso chk.GetAttributeValue("checked", False), "☑", "☐")
                            range.InsertAfter(symbol & " " &
                              HtmlEntity.DeEntitize(li.InnerText.Trim()) &
                              vbCr)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Next
                        Exit Select
                    End If

                    Dim enteringTopUL = (currentLevel = 0)
                    If enteringTopUL AndAlso ulLevels Is Nothing Then
                        ulLevels = New List(Of Integer)()
                        ulStartPos = range.Start
                        Debug.WriteLine("[ul] Entering top UL – reset ulLevels")
                    End If

                    ' c) Originales Einfügen der LI‑Nodes (recursive)
                    Dim liNodes = childNode.SelectNodes("li")
                    If liNodes IsNot Nothing Then
                        Dim listStart = range.Start
                        For Each liNode As HtmlAgilityPack.HtmlNode In liNodes
                            ParseHtmlNode(liNode, range, currentLevel + 1)
                            range.InsertAfter(vbCr)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Next
                        Dim ulRange As Word.Range = range.Document.Range(listStart, range.End)
                        ulRange.ListFormat.ApplyBulletDefault()
                        ulRange.ListFormat.ListIndent()
                        With ulRange.ParagraphFormat
                            .LeftIndent = .Application.CentimetersToPoints(0.75)
                            .FirstLineIndent = - .Application.CentimetersToPoints(0.75)
                        End With
                        range.SetRange(ulRange.End, ulRange.End)

                        ' d) Am Ende des ersten UL: Array anwenden
                        If enteringTopUL Then
                            Debug.WriteLine($"[ul] At end of top UL – ulLevels.Count = {ulLevels.Count}")
                            Debug.WriteLine($"[ul] Levels array: {String.Join(",", ulLevels)}")

                            Dim paras = ulRange.Paragraphs
                            Dim maxItems = system.math.min(paras.Count, ulLevels.Count)
                            For i = 1 To maxItems
                                Dim p = paras(i)
                                Dim lvl = ulLevels(i - 1)
                                Debug.WriteLine($"[ul] Paragraph {i} initial level={lvl}")

                                ' jede weitere Ebene einmal ListIndent()
                                For stepIndent = 1 To (lvl - 1)
                                    p.Range.ListFormat.ListIndent()
                                Next
                            Next
                            ulLevels = Nothing
                        End If
                    End If
                    If enteringTopUL Then
                        ulLevels = Nothing
                    End If

                    Exit Select


                Case "ol"

                    ' a) Alle <li>-Knoten holen
                    Dim liNodes As HtmlAgilityPack.HtmlNodeCollection = childNode.SelectNodes("li")
                    If liNodes Is Nothing OrElse liNodes.Count = 0 Then
                        Debug.WriteLine("[ol] No <li> nodes found in <ol> – skipping")
                        Exit Select
                    End If

                    ' b) Start-Attribut auslesen (Startnummer)
                    Dim startAttr As Integer = 1
                    Dim startStr As String = childNode.GetAttributeValue("start", String.Empty)
                    Dim tmpInt As Integer
                    If Integer.TryParse(startStr, tmpInt) Then startAttr = tmpInt

                    ' c) Top-Level-OL: ulLevels initialisieren (shared mit UL)
                    Dim enteringTopUL As Boolean = (currentLevel = 0)
                    If enteringTopUL AndAlso ulLevels Is Nothing Then
                        ulLevels = New List(Of Integer)()
                        ulStartPos = range.Start
                        Debug.WriteLine("[ol] Entering top OL – reset ulLevels")
                    End If

                    ' d) Jedes LI rekursiv rendern (die LI-Logik übernimmt CR und Level-Push)
                    Dim listStart As Integer = range.Start
                    For Each liNode As HtmlAgilityPack.HtmlNode In liNodes
                        ParseHtmlNode(liNode, range, currentLevel + 1)
                        range.InsertAfter(vbCr)
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Next

                    ' e) Den kompletten Bereich der OL-Liste erfassen
                    Dim olRange As Word.Range = range.Document.Range(listStart, range.End)

                    ' f) Prüfen, ob olRange überhaupt Absätze enthält
                    If olRange.Paragraphs.Count = 0 Then
                        Debug.WriteLine("[ol] olRange enthält keine Absätze – überspringe Nummerierung")
                        Exit Select
                    End If

                    ' j) Formatierung nur beim obersten OL (enteringTopUL) anwenden
                    If enteringTopUL Then
                        Dim paras As Word.Paragraphs = olRange.Paragraphs

                        ' Remove any previous numbering
                        olRange.ListFormat.RemoveNumbers()

                        ' Use a multi-level list template from ListGalleries
                        Dim multiLevelTemplate As Word.ListTemplate = olRange.Application.ListGalleries(Word.WdListGalleryType.wdOutlineNumberGallery).ListTemplates(1)

                        ' Set custom start value for level 1 if needed
                        If startAttr <> 1 Then
                            multiLevelTemplate.ListLevels(1).StartAt = startAttr
                        End If

                        ' Apply the multi-level template to the range
                        olRange.ListFormat.ApplyListTemplateWithLevel(
                                ListTemplate:=multiLevelTemplate,
                                ContinuePreviousList:=False,
                                ApplyTo:=Word.WdListApplyTo.wdListApplyToSelection,
                                DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior
                            )

                        ' Set the correct level for each paragraph
                        For i As Integer = 1 To paras.Count
                            Dim p As Word.Paragraph = paras(i)
                            Dim lvl As Integer = ulLevels(i - 1)
                            If lvl >= 1 AndAlso lvl <= multiLevelTemplate.ListLevels.Count Then
                                p.Range.ListFormat.ListLevelNumber = lvl
                            End If
                        Next
                        ulLevels = Nothing
                    End If

                    ' j) Cursor hinter die Liste setzen
                    range.SetRange(olRange.End, olRange.End)

                    Exit Select

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

                    RenderInline(childNode, range, Nothing, String.Empty)

                Case "pre"
                    ' Code block
                    Dim codeBlock As Microsoft.Office.Interop.Word.Range = range.Duplicate
                    codeBlock.Text = HtmlEntity.DeEntitize(childNode.InnerText) & vbCr
                    codeBlock.Font.Name = "Courier New"
                    codeBlock.Font.Size = 10
                    codeBlock.ParagraphFormat.LeftIndent += 14.18
                    codeBlock.Collapse(False)
                    range.SetRange(codeBlock.End, codeBlock.End)




                Case "table"
                    '---------- 1) Top-Level-Rows holen ----------------------------
                    Dim topRows As New List(Of HtmlNode)

                    'direkte <tr> plus <thead>/<tbody>-Kinder, aber KEINE rekursiven
                    For Each tr As HtmlNode In childNode.SelectNodes("./tr|./thead/tr|./tbody/tr")
                        topRows.Add(tr)
                    Next
                    If topRows.Count = 0 Then Exit Select
                    '----------------------------------------------------------------

                    '---------- 2) Beste Spaltenzahl ermitteln ----------------------
                    Dim colCount As Integer = 0
                    For Each tr In topRows
                        Dim cells = tr.SelectNodes("th|td")
                        If cells IsNot Nothing AndAlso cells.Count > colCount Then
                            colCount = cells.Count
                        End If
                    Next
                    If colCount = 0 Then Exit Select
                    '----------------------------------------------------------------

                    '---------- 3) Tabelle an Cursor anlegen ------------------------
                    Dim tbl As Microsoft.Office.Interop.Word.Table =
                        range.Document.Tables.Add(range, topRows.Count, colCount)

                    '---------- 4) Zellen befüllen ----------------------------------
                    Dim rIdx As Integer = 1
                    For Each tr In topRows
                        Dim cells = tr.SelectNodes("th|td")
                        Dim cIdx As Integer = 1

                        If cells IsNot Nothing Then
                            For Each cell In cells
                                Dim cellRg As Word.Range = tbl.Cell(rIdx, cIdx).Range
                                'unsichtbares Zellenendzeichen abschneiden
                                cellRg.SetRange(cellRg.Start, cellRg.End - 1)

                                ParseHtmlNode(cell, cellRg, currentLevel)          '← rekursiv, kein Datenverlust

                                'Headerzelle fett
                                If cell.Name.Equals("th", StringComparison.OrdinalIgnoreCase) Then
                                    cellRg.Font.Bold = True
                                End If

                                '---------- 4a) Colspan behandeln ------------------
                                Dim cSpan As Integer = cell.GetAttributeValue("colspan", 1)
                                If cSpan > 1 AndAlso cIdx + cSpan - 1 <= colCount Then
                                    Dim tgtCell = tbl.Cell(rIdx, cIdx + cSpan - 1)
                                    tbl.Cell(rIdx, cIdx).Merge(tgtCell)
                                    cIdx += cSpan                    'gleich weiter hinter dem Merge
                                Else
                                    cIdx += 1
                                End If
                                '----------------------------------------------------
                            Next
                        End If
                        rIdx += 1
                    Next

                    '---------- 5) Cursor hinter Tabelle setzen --------------------
                    range.SetRange(tbl.Range.End, tbl.Range.End)


                Case Else

                    ParseHtmlNode(childNode, range, currentLevel)

            End Select

        Next
    End Sub





    Private Shared Sub InsertInline(
        ByRef mainRg As Range,
        txt As String,
        baseStyle As Action(Of Range),
        Optional href As String = "")

        ' 1) Kein Emoji‑Fall → direkter Einfügen‑Pfad
        If emojiSet Is Nothing OrElse emojiSet.Count = 0 OrElse Not _emojiPairRegex.IsMatch(txt) Then
            TrueInsertInline(mainRg, txt, baseStyle, href)
            Return
        End If

        ' 2) Sonst: nur an den Emoji-Punkten splitten und stylen
        Dim lastPos As Integer = 0
        For Each m As Match In _emojiPairRegex.Matches(txt)
            ' (a) Alles vor dem Emoji
            If m.Index > lastPos Then
                Dim segment As String = txt.Substring(lastPos, m.Index - lastPos)
                TrueInsertInline(mainRg, segment, baseStyle, href)
            End If

            ' (b) Das Emoji selbst (nur, wenn es im Set ist)
            Dim emoji As String = m.Value
            If emojiSet.Contains(emoji) Then
                Dim emojiStyle = CombineStyle(baseStyle,
                    Sub(r As Range) r.Font.Name = "Segoe UI Emoji")
                TrueInsertInline(mainRg, emoji, emojiStyle, href)
            Else
                ' falls doch nicht im Set – als normaler Text
                TrueInsertInline(mainRg, emoji, baseStyle, href)
            End If

            lastPos = m.Index + m.Length
        Next

        ' (c) Rest nach dem letzten Emoji
        If lastPos < txt.Length Then
            Dim tail As String = txt.Substring(lastPos)
            TrueInsertInline(mainRg, tail, baseStyle, href)
        End If
    End Sub




    Private Shared Sub TrueInsertInline(
    ByRef mainRg As Word.Range,
    txt As String,
    styleAction As Action(Of Word.Range),
    Optional href As String = "")

        mainRg.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

        Dim wrk As Word.Range = mainRg.Duplicate
        wrk.Text = txt

        ' **Hier passiert der Reset** – denk daran, vor und nach dem Reset zu loggen
        wrk.Font.Reset()

        If href <> "" Then
            Dim hl = mainRg.Document.Hyperlinks.Add(Anchor:=wrk, Address:=href)
            If styleAction IsNot Nothing Then
                Debug.WriteLine("[InsertInline] → Anwenden styleAction auf Hyperlink‑Range")
                styleAction(hl.Range)
            End If
            hl.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            mainRg.SetRange(hl.Range.End, hl.Range.End)
        Else
            If styleAction IsNot Nothing Then
                Debug.WriteLine("[InsertInline] → Anwenden styleAction auf Text‑Range")
                styleAction(wrk)
            End If
            wrk.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            mainRg.SetRange(wrk.End, wrk.End)
        End If

    End Sub



    ' Structure to store revision information for fast processing
    Private Structure RevInfo
        Public Start As Integer
        Public EndPos As Integer  ' Using EndPos instead of End to avoid keyword conflict
        Public Text As String
        Public Type As WdRevisionType
        Public Author As String
    End Structure

    Public Function AddMarkupTags(ByVal rng As Range, Optional ByVal TPMarkupName As String = Nothing) As String

        Dim splash As New SLib.SplashScreen("Coding markups...  counting")
        splash.Show()
        splash.Refresh()

        ' Quick exit for ranges without revisions
        Dim revCount As Integer = 0
        Try
            revCount = rng.Revisions.Count
        Catch ex As Exception
            splash.Close()
            Return rng.Text
        End Try

        If revCount = 0 Then
            splash.Close()
            Return rng.Text
        End If

        ' Get range boundaries
        Dim rangeStart As Integer = rng.Start
        Dim rangeEnd As Integer = rng.End
        Dim resultBuilder As New StringBuilder(rng.Text.Length * 2)

        ' Create a collection to hold all revision data in memory
        Dim revInfos As New List(Of RevInfo)(revCount)

        ' Collect all revision data in a single pass to minimize COM calls
        For i As Integer = 1 To revCount
            splash.UpdateMessage($"Collecting markups... {revCount - i} left")
            Try

                Dim rev As Revision = rng.Revisions(i)

                Try
                    Dim revRange As Range = rev.Range

                    Dim revStart As Integer = revRange.Start
                    Dim revEnd As Integer = revRange.End

                    ' Only process revisions that overlap with our range
                    If revEnd > rangeStart AndAlso revStart < rangeEnd Then
                        Dim revText As String = revRange.Text
                        Dim revType As WdRevisionType = rev.Type
                        Dim revAuthor As String = rev.Author

                        ' Create a value type to store data efficiently
                        revInfos.Add(New RevInfo() With {
                        .Start = revStart,
                        .EndPos = revEnd,
                        .Text = revText,
                        .Type = revType,
                        .Author = revAuthor
                    })
                    End If
                Catch ex As Exception
                    ' Skip this revision and continue
                    Continue For
                End Try
            Catch ex As Exception
                Debug.WriteLine($"AddMarkupTags: ERROR with revision {i}: {ex.Message}")
                ' Skip and continue with next revision
                Continue For
            End Try

        Next

        ' Sort revisions by start position
        revInfos.Sort(Function(a, b) a.Start.CompareTo(b.Start))

        ' Process document with minimal COM access
        Dim currentPos As Integer = rangeStart
        Dim ii As Integer = 0

        For Each info In revInfos

            splash.UpdateMessage("Coding markups... " & revInfos.Count - ii & " left")
            ii = ii + 1

            ' Add text before this revision
            If info.Start > currentPos Then
                Try
                    Debug.WriteLine($"AddMarkupTags: Getting text before revision: {currentPos} to {info.Start}")
                    Dim beforeText As String = rng.Document.Range(currentPos, info.Start).Text
                    resultBuilder.Append(beforeText)
                Catch ex As Exception
                    Debug.WriteLine($"AddMarkupTags: Error getting text before revision: {ex.Message}")
                    ' If we can't get the text, just continue
                End Try
            End If

            ' Check if we should include markup
            Dim includeMarkup As Boolean = String.IsNullOrEmpty(TPMarkupName) OrElse
            String.Equals(info.Author, TPMarkupName, StringComparison.OrdinalIgnoreCase)

            ' Add revision text with markup
            If includeMarkup Then
                Select Case info.Type
                    Case WdRevisionType.wdRevisionDelete
                        resultBuilder.Append("<del>").Append(info.Text).Append("</del>")
                        Debug.WriteLine($"AddMarkupTags: Added delete markup: {info.Text.Length} chars")
                    Case WdRevisionType.wdRevisionInsert
                        resultBuilder.Append("<ins>").Append(info.Text).Append("</ins>")
                        Debug.WriteLine($"AddMarkupTags: Added insert markup: {info.Text.Length} chars")
                    Case Else
                        resultBuilder.Append(info.Text)
                End Select
            Else
                resultBuilder.Append(info.Text)
            End If

            ' Update position
            currentPos = info.EndPos
        Next

        ' Add any remaining text
        If currentPos < rangeEnd Then
            Try
                Dim tailText As String = rng.Document.Range(currentPos, rangeEnd).Text
                resultBuilder.Append(tailText)
            Catch ex As Exception
                ' If we can't get the remaining text, just return what we have
            End Try
        End If

        splash.Close()

        Return resultBuilder.ToString()
    End Function


    Public Function OldAddMarkupTags(ByVal rng As Range, Optional ByVal TPMarkupName As String = Nothing) As String
        ' Read the entire range text at once to minimize COM calls
        Dim fullText As String = rng.Text
        Dim resultBuilder As New StringBuilder(fullText.Length * 2) ' Pre-allocate with extra space for tags

        ' Get all revisions and sort them
        Dim revList = rng.Revisions _
        .Cast(Of Revision)() _
        .OrderBy(Function(r) r.Range.Start - rng.Start) _
        .ToList()

        ' Track our position in the source text
        Dim currentPos As Integer = 0

        Debug.WriteLine("Revlist Number: " & revList.Count)

        ' Process all revisions
        For Each rev As Revision In revList
            ' Calculate position relative to our range
            Dim relativeStart As Integer = rev.Range.Start - rng.Start
            Dim relativeEnd As Integer = rev.Range.End - rng.Start

            ' Add text before this revision
            If relativeStart > currentPos Then
                resultBuilder.Append(fullText.Substring(currentPos, relativeStart - currentPos))
            End If

            ' Check if we should include markup for this author
            Dim includeMarkup As Boolean = String.IsNullOrEmpty(TPMarkupName) OrElse
            String.Equals(rev.Author, TPMarkupName, StringComparison.OrdinalIgnoreCase)

            ' Get the revision text
            Dim revText As String = rev.Range.Text

            Debug.WriteLine("MarkupText: " & revText)

            ' Append with appropriate tags
            If includeMarkup Then
                Select Case rev.Type
                    Case WdRevisionType.wdRevisionDelete
                        resultBuilder.Append("<del>").Append(revText).Append("</del>")
                    Case WdRevisionType.wdRevisionInsert
                        resultBuilder.Append("<ins>").Append(revText).Append("</ins>")
                    Case Else
                        resultBuilder.Append(revText)
                End Select
            Else
                resultBuilder.Append(revText)
            End If

            ' Update current position
            currentPos = relativeEnd
        Next

        ' Add any remaining text
        If currentPos < fullText.Length Then
            resultBuilder.Append(fullText.Substring(currentPos))
        End If

        Return resultBuilder.ToString()
    End Function


    Public Function xxxAddMarkupTags(ByVal rng As Range, Optional ByVal TPMarkupName As String = Nothing) As String
        Dim resultBuilder As New StringBuilder()

        ' 1. Alle Revisionen in Dokumentreihenfolge sortieren
        Dim revList = rng.Revisions _
        .Cast(Of Revision)() _
        .OrderBy(Function(r) r.Range.Start) _
        .ToList()

        ' 2. Startposition auf Anfang des Bereichs setzen
        Dim currentPos As Integer = rng.Start

        ' 3. Über jede Revision iterieren
        For Each rev As Revision In revList
            ' 3a. Unveränderten Text vor der Revision anhängen
            If rev.Range.Start > currentPos Then
                Dim unchangedRange As Range = rng.Document.Range(currentPos, rev.Range.Start)
                resultBuilder.Append(unchangedRange.Text)
            End If

            ' 3b. Prüfen, ob nach TPMarkupName gefiltert werden soll
            Dim includeMarkup As Boolean = True
            If Not String.IsNullOrEmpty(TPMarkupName) Then
                If Not String.Equals(rev.Author, TPMarkupName, StringComparison.OrdinalIgnoreCase) Then
                    includeMarkup = False
                End If
            End If

            ' 3c. Revisionstext mit Tags oder neutral anhängen
            If includeMarkup Then
                Select Case rev.Type
                    Case WdRevisionType.wdRevisionDelete
                        resultBuilder.Append("<del>").Append(rev.Range.Text).Append("</del>")
                    Case WdRevisionType.wdRevisionInsert
                        resultBuilder.Append("<ins>").Append(rev.Range.Text).Append("</ins>")
                    Case Else
                        resultBuilder.Append(rev.Range.Text)
                End Select
            Else
                resultBuilder.Append(rev.Range.Text)
            End If

            ' 3d. Position hinter der Revision setzen
            currentPos = rev.Range.End
        Next

        ' 4. Restlichen unveränderten Text bis zum Ende des Bereichs anhängen
        If currentPos < rng.End Then
            Dim tailRange As Range = rng.Document.Range(currentPos, rng.End)
            resultBuilder.Append(tailRange.Text)
        End If

        ' Ergebnis zurückgeben
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


    Private Sub oldCompareAndInsert(text1 As String, text2 As String, targetRange As Range, Optional ShowInWindow As Boolean = False, Optional TextforWindow As String = "A text with these changes will be inserted ('Esc' to abort):", Optional paraformatinline As Boolean = False, Optional noformatting As Boolean = True)
        Try

            Dim diffBuilder As New InlineDiffBuilder(New Differ())
            Dim sText As String = String.Empty

            Debug.WriteLine("A Text1 = " & text1)
            Debug.WriteLine("A Text2 = " & text2)

            ' Pre-process the texts to replace line breaks with a unique marker
            text1 = text1.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")
            text2 = text2.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")

            ' Normalize the texts by removing extra spaces
            text1 = text1.Replace("  ", " ").Trim()
            text2 = text2.Replace("  ", " ").Trim()

            Debug.WriteLine("B Text1 = " & text1)
            Debug.WriteLine("B Text2 = " & text2)

            ' Split the texts into words and convert them into a line-by-line format
            ' 3) In Worte splitten (ohne leere Einträge) und zeilenweise darstellen
            Dim words1 As String = String.Join(
              Environment.NewLine,
              text1.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    )
            Dim words2 As String = String.Join(
              Environment.NewLine,
              text2.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    )
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

            Debug.WriteLine("1 = " & sText)

            ' Remove preceding and trailing spaces around placeholders
            sText = sText.Replace("{vbCr}", "{vbCrLf}")
            sText = sText.Replace("{vbLf}", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf} ", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf}", "{vbCrLf}")
            sText = sText.Replace("{vbCrLf} ", "{vbCrLf}")

            Debug.WriteLine("2 = " & sText)

            ' Remove instances of line breaks surrounded by [DEL_START] and [DEL_END]
            sText = sText.Replace("[DEL_START]{vbCrLf}[DEL_END] ", "")
            sText = sText.Replace("[DEL_START]{vbCrLf}{vbCrLf}[DEL_END] ", "")
            sText = sText.Replace("{vbCrLf}[DEL_END] ", "{vbCrLf}[DEL_END]")

            ' Include instances of line breaks surrounded by [INS_START] and [INS_END] without the [INS...] text
            sText = sText.Replace("[INS_START]{vbCrLf}[INS_END] ", "{vbCrLf}")
            sText = sText.Replace("[INS_START]{vbCrLf}{vbCrLf}[INS_END] ", "{vbCrLf}{vbCrLf}")

            ' Entferne alle überflüssigen Leerzeilen-Platzhalter am Ende

            Debug.WriteLine("3 = " & sText)

            sText = sText.Replace(vbCrLf, "").Replace(vbCr, "").Replace(vbLf, "")

            ' Replace placeholders with actual line breaks
            sText = sText.Replace("{vbCrLf}", vbCrLf)

            ' Adjust overlapping tags
            sText = sText.Replace("[DEL_END] [INS_START]", "[DEL_END][INS_START]")
            sText = sText.Replace("[INS_START][INS_END] ", "")
            sText = RemoveInsDelTagsInPlaceholders(sText)

            ' Insert formatted text into the specified range
            If Not ShowInWindow Then
                Debug.WriteLine("Text with tags: " & vbCrLf & "'" & sText & "'" & vbCrLf & vbCrLf)
                InsertMarkupText(sText, targetRange)
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


    Private Sub CompareAndInsert(text1 As String, text2 As String, targetRange As Range, Optional ShowInWindow As Boolean = False, Optional TextforWindow As String = "A text with these changes will be inserted ('Esc' to abort):", Optional paraformatinline As Boolean = False, Optional noformatting As Boolean = True)
        Try

            Dim diffBuilder As New InlineDiffBuilder(New Differ())
            Dim sText As String = String.Empty

            Debug.WriteLine("A Text1 = " & text1)
            Debug.WriteLine("A Text2 = " & text2)

            ' Pre-process the texts to replace line breaks with a unique marker
            text1 = text1.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")
            text2 = text2.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")

            ' Normalize the texts by removing extra spaces
            text1 = text1.Replace("  ", " ").Trim()
            text2 = text2.Replace("  ", " ").Trim()

            Debug.WriteLine("B Text1 = " & text1)
            Debug.WriteLine("B Text2 = " & text2)

            ' Split the texts into words and convert them into a line-by-line format
            ' In Worte splitten (ohne leere Einträge) und zeilenweise darstellen
            '--- 1) pull out all {{…}} fields into a list and replace them with placeholders:
            Dim mergefields As New List(Of String)
            text1 = System.Text.RegularExpressions.Regex.Replace(text1, "\{\{.*?\}\}",
    Function(m)
        mergefields.Add(m.Value)
        Return $"[[MF{mergefields.Count - 1}]]"
    End Function)
            text2 = System.Text.RegularExpressions.Regex.Replace(text2, "\{\{.*?\}\}",
    Function(m)
        mergefields.Add(m.Value)
        Return $"[[MF{mergefields.Count - 1}]]"
    End Function)

            ' Split the texts into words and convert them into a line-by-line format
            ' 3) In Worte splitten (ohne leere Einträge) und zeilenweise darstellen
            Dim words1 As String = String.Join(
              Environment.NewLine,
              text1.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    )
            Dim words2 As String = String.Join(
              Environment.NewLine,
              text2.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    )
            ' Generate word-based diff using DiffPlex
            Dim diffResult As DiffPaneModel = diffBuilder.BuildDiffModel(words1, words2)

            '--- 4) emit tags *per run* rather than per word:
            Dim prevType = ChangeType.Unchanged
            For i = 0 To diffResult.Lines.Count - 1
                Dim line = diffResult.Lines(i)
                Dim nextType = If(i < diffResult.Lines.Count - 1, diffResult.Lines(i + 1).Type, ChangeType.Unchanged)

                ' open tag when entering an Insert or Delete run
                If line.Type = ChangeType.Inserted AndAlso prevType <> ChangeType.Inserted Then
                    sText &= "[INS_START]"
                ElseIf line.Type = ChangeType.Deleted AndAlso prevType <> ChangeType.Deleted Then
                    sText &= "[DEL_START]"
                End If

                ' the word itself
                sText &= line.Text.Trim() & " "

                ' close tag when exiting a run
                If line.Type = ChangeType.Inserted AndAlso nextType <> ChangeType.Inserted Then
                    sText &= "[INS_END] "
                ElseIf line.Type = ChangeType.Deleted AndAlso nextType <> ChangeType.Deleted Then
                    sText &= "[DEL_END] "
                End If

                prevType = line.Type
            Next

            '--- 5) put your merge‑fields back in-place:
            For idx = 0 To mergefields.Count - 1
                sText = sText.Replace($"[[MF{idx}]]", mergefields(idx))
            Next

            Debug.WriteLine("1 = " & sText)

            ' Remove preceding and trailing spaces around placeholders
            sText = sText.Replace("{vbCr}", "{vbCrLf}")
            sText = sText.Replace("{vbLf}", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf} ", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf}", "{vbCrLf}")
            sText = sText.Replace("{vbCrLf} ", "{vbCrLf}")

            Debug.WriteLine("2 = " & sText)

            ' Remove instances of line breaks surrounded by [DEL_START] and [DEL_END]
            sText = sText.Replace("[DEL_START]{vbCrLf}[DEL_END] ", "")
            sText = sText.Replace("[DEL_START]{vbCrLf}{vbCrLf}[DEL_END] ", "")
            sText = sText.Replace("{vbCrLf}[DEL_END] ", "{vbCrLf}[DEL_END]")

            ' Include instances of line breaks surrounded by [INS_START] and [INS_END] without the [INS...] text
            sText = sText.Replace("[INS_START]{vbCrLf}[INS_END] ", "{vbCrLf}")
            sText = sText.Replace("[INS_START]{vbCrLf}{vbCrLf}[INS_END] ", "{vbCrLf}{vbCrLf}")
            sText = sText.Replace("{vbCrLf}[INS_END] ", "{vbCrLf}[INS_END]")

            ' Entferne alle überflüssigen Leerzeilen-Platzhalter am Ende

            Debug.WriteLine("3 = " & sText)

            sText = sText.Replace(vbCrLf, "").Replace(vbCr, "").Replace(vbLf, "")

            ' Replace placeholders with actual line breaks
            sText = sText.Replace("{vbCrLf}", vbCrLf)

            ' Adjust overlapping tags
            sText = sText.Replace("[DEL_END] [INS_START]", "[DEL_END][INS_START]")
            sText = sText.Replace("[INS_START][INS_END] ", "")
            'sText = RemoveInsDelTagsInPlaceholders(sText)

            ' Insert formatted text into the specified range
            If Not ShowInWindow Then
                Debug.WriteLine("Text with tags: " & vbCrLf & "'" & sText & "'" & vbCrLf & vbCrLf)
                InsertMarkupText(sText, targetRange)
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


    Public Function RemoveInsDelTagsInPlaceholders(input As String) As String
        Try
            ' Regex-Pattern:
            ' \{\{                   Literal '{{'
            ' (?<content>.*?)        Inhalt vor dem Tag (Lazy)
            ' \[(?:INS_START|DEL_START)\]   Start-Tag
            ' (?<after>.*?)          Inhalt bis zum Ende der Klammer (Lazy)
            ' \}\}                   Literal '}}'
            ' \[(?:INS_END|DEL_END)\] Direkt folgendes End-Tag
            Dim pattern As String = "\{\{(?<content>.*?)\[(?:INS_START|DEL_START)\](?<after>.*?)\}\}\[(?:INS_END|DEL_END)\]"

            ' MatchEvaluator als Delegate
            Dim evaluator As System.Text.RegularExpressions.MatchEvaluator =
                Function(m As System.Text.RegularExpressions.Match) As String
                    Return "{{" & m.Groups("content").Value & m.Groups("after").Value & "}}"
                End Function

            ' Replace mit korrektem Overload
            Return System.Text.RegularExpressions.Regex.Replace(
                input,
                pattern,
                evaluator,
                System.Text.RegularExpressions.RegexOptions.Singleline
            )
        Catch ex As System.Exception
            ' Gracefully

        End Try
    End Function

    Public Sub InsertMarkupText(ByVal inputText As String, ByVal targetRange As Microsoft.Office.Interop.Word.Range)
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.ActiveDocument

        Dim originalTrack As Boolean = doc.TrackRevisions
        Dim originalUpdate As Boolean = wordApp.ScreenUpdating

        ' Die Positions‑Variablen VOR dem Try deklarieren,
        ' damit sie auch in Finally noch gültig sind:
        Dim docStart As Integer
        Dim startPos As Integer
        Dim endPosNoCR As Integer

        Try
            wordApp.ScreenUpdating = False
            doc.TrackRevisions = False

            '------------------------------------------------------------------
            '  A) Preserve the trailing ¶ so the next paragraph never joins in
            '------------------------------------------------------------------
            docStart = doc.Content.Start
            startPos = targetRange.Start
            endPosNoCR = targetRange.End

            If endPosNoCR > docStart Then
                Dim checkRange As Microsoft.Office.Interop.Word.Range =
                    doc.Range(endPosNoCR - 1, endPosNoCR)
                If checkRange.Text = vbCr Then
                    endPosNoCR -= 1
                End If
            End If

            If endPosNoCR >= startPos Then
                doc.Range(startPos, endPosNoCR).Delete()
            End If

            targetRange.SetRange(startPos, startPos)

            '------------------------------------------------------------------
            '  Merge contiguous INS‑ und DEL‑Tags mit nur Leerzeichen dazwischen
            '------------------------------------------------------------------
            Dim txt As String = inputText
            txt = RemoveMergeFormatFromBraces(txt)

            '--- Strip merge‑fields out of **closed** delete‑runs:
            txt = System.Text.RegularExpressions.Regex.Replace(
                txt,
                "\[DEL_START\]([\s\S]*?)\[DEL_END\]",
                Function(m As System.Text.RegularExpressions.Match) As String
                    Return "[DEL_START]" &
                           System.Text.RegularExpressions.Regex.Replace(
                               m.Groups(1).Value,
                               "\{\{(?:WFLD|WFNT|WENT|PFOR):.*?\}\}",
                               String.Empty
                           ) &
                           "[DEL_END]"
                End Function,
                System.Text.RegularExpressions.RegexOptions.Singleline
            )

            '--- Strip merge‑fields out of **open** delete‑runs (no closing tag),
            '    but only if wirklich kein [DEL_END] folgt:
            txt = System.Text.RegularExpressions.Regex.Replace(
                txt,
                "\[DEL_START\]((?:(?!\[DEL_END\]).)*)$",
                Function(m As System.Text.RegularExpressions.Match) As String
                    Return "[DEL_START]" &
                           System.Text.RegularExpressions.Regex.Replace(
                               m.Groups(1).Value,
                               "\{\{(?:WFLD|WFNT|WENT|PFOR):.*?\}\}",
                               String.Empty
                           )
                End Function,
                System.Text.RegularExpressions.RegexOptions.Singleline
            )

            Debug.WriteLine("Stripped txt1 = " & txt)

            txt = System.Text.RegularExpressions.Regex.Replace(txt, "\[INS_END\](\s*)\[INS_START\]", "$1")
            txt = System.Text.RegularExpressions.Regex.Replace(txt, "\[DEL_END\](\s*)\[DEL_START\]", "$1")

            Debug.WriteLine("Stripped txt2 = " & txt)

            While txt.Length > 0
                System.Windows.Forms.Application.DoEvents()
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exit While

                ' locate next opening tag
                Dim insPos As Integer = txt.IndexOf("[INS_START]", StringComparison.Ordinal)
                Dim delPos As Integer = txt.IndexOf("[DEL_START]", StringComparison.Ordinal)

                Dim nextTagPos As Integer
                Dim tagType As String = Nothing
                If insPos = -1 AndAlso delPos = -1 Then
                    nextTagPos = -1
                ElseIf insPos = -1 OrElse (delPos <> -1 AndAlso delPos < insPos) Then
                    nextTagPos = delPos : tagType = "DEL"
                Else
                    nextTagPos = insPos : tagType = "INS"
                End If

                ' Plain text vor dem nächsten Tag
                If nextTagPos = -1 OrElse nextTagPos > 0 Then
                    Dim plain As String = If(nextTagPos = -1, txt, txt.Substring(0, nextTagPos))
                    If plain.Length > 0 Then
                        doc.TrackRevisions = False
                        targetRange.InsertAfter(plain)
                        targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    End If
                End If
                If nextTagPos = -1 Then Exit While

                If tagType = "INS" Then
                    '==============================================================
                    '  INSERT block
                    '==============================================================
                    txt = txt.Substring(nextTagPos + "[INS_START]".Length)
                    Dim endIns As Integer = txt.IndexOf("[INS_END]", StringComparison.Ordinal)
                    Dim insText As String = If(endIns = -1, txt, txt.Substring(0, endIns))
                    If endIns <> -1 Then txt = txt.Substring(endIns + "[INS_END]".Length)
                    doc.TrackRevisions = True
                    targetRange.InsertAfter(insText)
                    targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    doc.TrackRevisions = False

                Else
                    '==============================================================
                    '  DELETION block
                    '==============================================================
                    txt = txt.Substring(nextTagPos + "[DEL_START]".Length)
                    Dim endDel As Integer = txt.IndexOf("[DEL_END]", StringComparison.Ordinal)
                    Dim delText As String = If(endDel = -1, txt, txt.Substring(0, endDel))
                    If endDel <> -1 Then txt = txt.Substring(endDel + "[DEL_END]".Length)

                    ' absorb following space/CR
                    If txt.StartsWith(" ") Then
                        delText &= " " : txt = txt.Substring(1)
                    ElseIf txt.StartsWith(vbCrLf) Then
                        delText &= vbCrLf : txt = txt.Substring(2)
                    ElseIf txt.StartsWith(vbCr) Then
                        delText &= vbCr : txt = txt.Substring(1)
                    End If

                    ' a) einfügen (silent)
                    doc.TrackRevisions = False
                    targetRange.Text = delText
                    ' b) löschen (mit Tracking)
                    doc.TrackRevisions = True
                    targetRange.Delete()
                    doc.TrackRevisions = False
                    targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                End If
            End While

        Catch ex As System.Exception
            Debug.WriteLine("InsertMarkupText error: " & ex.Message & vbCrLf & inputText)
        Finally

            ' --- Final-View Replace Test (Space-Bereinigung) ---

            ' 2) Final-View aktivieren
            With wordApp.ActiveWindow.View
                .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                .ShowRevisionsAndComments = False
            End With
            ' 3) Replace doppelte Spaces
            ' Temporär Revisionen ausschalten, damit die Ersetzungen nicht als Änderungen protokolliert werden
            doc.TrackRevisions = False

            Dim endPosInserted2 As Integer = targetRange.End
            Dim insertedRange As Microsoft.Office.Interop.Word.Range =
                doc.Range(startPos, endPosInserted2)


            ' Find/Replace für zwei Leerzeichen → ein Leerzeichen
            With insertedRange.Find
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Text = "  "    ' genau zwei Leerzeichen
                .Replacement.Text = " "
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .Format = False
                .MatchWildcards = False
            End With

            ' Solange noch ein Replace stattfindet, wiederholen
            Do
                ' Execute gibt True zurück, wenn etwas ersetzt wurde
            Loop While insertedRange.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

            With wordApp.ActiveWindow.View
                .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                .ShowRevisionsAndComments = True
            End With

            ' Tracking wieder in den Ursprungszustand versetzen
            doc.TrackRevisions = originalTrack

            wordApp.ScreenUpdating = originalUpdate

            ' Range auf die volle eingefügte Länge setzen
            Dim endPosInserted As Integer = targetRange.End
            targetRange.SetRange(startPos, endPosInserted)
            wordApp.Selection.SetRange(targetRange.Start, targetRange.End)
        End Try
    End Sub


    Private Sub oldInsertMarkupText(inputText As String, targetRange As Word.Range)

        Dim wordApp As Word.Application = Globals.ThisAddIn.Application
        Dim doc As Word.Document = wordApp.ActiveDocument

        Dim originalTrack As Boolean = doc.TrackRevisions
        Dim originalUpdate As Boolean = wordApp.ScreenUpdating

        '------------------------------------------------------------------
        '  A) Preserve the trailing ¶ so the next paragraph never joins in
        '------------------------------------------------------------------
        Dim startPos As Integer = targetRange.Start
        Dim endPosNoCR As Integer = targetRange.End
        If doc.Range(endPosNoCR - 1, endPosNoCR).Text = vbCr Then endPosNoCR -= 1

        Try
            wordApp.ScreenUpdating = False
            doc.TrackRevisions = False

            doc.Range(startPos, endPosNoCR).Delete()      ' keep ¶ intact
            targetRange.SetRange(startPos, startPos)      ' collapsed cursor

            '------------------------------------------------------------------
            '  Merge contiguous INS- und DEL-Tags mit nur Leerzeichen dazwischen
            '------------------------------------------------------------------
            Dim txt As String = inputText

            txt = RemoveMergeFormatFromBraces(txt)

            '--- Strip merge‑fields out of **closed** delete‑runs:
            txt = System.Text.RegularExpressions.Regex.Replace(
              txt,
              "\[DEL_START\]([\s\S]*?)\[DEL_END\]",
              Function(m) _
                "[DEL_START]" &
                System.Text.RegularExpressions.Regex.Replace(
                  m.Groups(1).Value,
                  "\{\{(?:WFLD|WFNT|WENT|PFOR):.*?\}\}",
                  String.Empty
                ) &
                "[DEL_END]" _
              , System.Text.RegularExpressions.RegexOptions.Singleline
            )

            '--- 2) Strip merge‑fields out of **open** delete‑runs (no closing tag),
            '       but only if there really is no [DEL_END] after this [DEL_START]:
            txt = System.Text.RegularExpressions.Regex.Replace(
                  txt,
                  "\[DEL_START\]((?:(?!\[DEL_END\]).)*)$",
                  Function(m) _
                    "[DEL_START]" &
                    System.Text.RegularExpressions.Regex.Replace(
                      m.Groups(1).Value,
                      "\{\{(?:WFLD|WFNT|WENT|PFOR):.*?\}\}",
                      String.Empty
                    ) _
                  , System.Text.RegularExpressions.RegexOptions.Singleline
                )

            Debug.WriteLine("Stripped txt1 = " & txt)

            txt = System.Text.RegularExpressions.Regex.Replace(
                  txt, "\[INS_END\](\s*)\[INS_START\]", "$1"
              )
            txt = System.Text.RegularExpressions.Regex.Replace(
                  txt, "\[DEL_END\](\s*)\[DEL_START\]", "$1"
              )

            Debug.WriteLine("Stripped txt2 = " & txt)

            While txt.Length > 0
                System.Windows.Forms.Application.DoEvents()
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exit While

                ' locate next opening tag
                Dim insPos As Integer = txt.IndexOf("[INS_START]", StringComparison.Ordinal)
                Dim delPos As Integer = txt.IndexOf("[DEL_START]", StringComparison.Ordinal)

                Dim nextTagPos As Integer
                Dim tagType As String = Nothing
                If insPos = -1 AndAlso delPos = -1 Then
                    nextTagPos = -1
                ElseIf insPos = -1 OrElse (delPos <> -1 AndAlso delPos < insPos) Then
                    nextTagPos = delPos : tagType = "DEL"
                Else
                    nextTagPos = insPos : tagType = "INS"
                End If

                ' plain text before tag
                If nextTagPos = -1 OrElse nextTagPos > 0 Then
                    Dim plain As String = If(nextTagPos = -1, txt, txt.Substring(0, nextTagPos))
                    If plain.Length > 0 Then
                        doc.TrackRevisions = False
                        targetRange.InsertAfter(plain)
                        targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    End If
                End If
                If nextTagPos = -1 Then Exit While

                '==============================================================
                '  INSERT block
                '==============================================================
                If tagType = "INS" Then
                    txt = txt.Substring(nextTagPos + "[INS_START]".Length)
                    Dim endIns As Integer = txt.IndexOf("[INS_END]", StringComparison.Ordinal)
                    Dim insText As String = If(endIns = -1, txt, txt.Substring(0, endIns))
                    If endIns <> -1 Then txt = txt.Substring(endIns + "[INS_END]".Length) Else txt = ""

                    doc.TrackRevisions = True
                    targetRange.InsertAfter(insText)
                    targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    doc.TrackRevisions = False

                    '==============================================================
                    '  DELETION block  (with padding space)
                    '==============================================================
                Else
                    txt = txt.Substring(nextTagPos + "[DEL_START]".Length)
                    Dim endDel As Integer = txt.IndexOf("[DEL_END]", StringComparison.Ordinal)
                    Dim delText As String = If(endDel = -1, txt, txt.Substring(0, endDel))
                    If endDel <> -1 Then txt = txt.Substring(endDel + "[DEL_END]".Length) Else txt = ""

                    ' absorb space/¶ immediately following the tag
                    If txt.StartsWith(" ") Then
                        delText &= " " : txt = txt.Substring(1)
                    ElseIf txt.StartsWith(vbCrLf) Then
                        delText &= vbCrLf : txt = txt.Substring(2)
                    ElseIf txt.StartsWith(vbCr) Then
                        delText &= vbCr : txt = txt.Substring(1)
                    End If

                    '--- PAD with an extra space so Word won't merge partial-word deletions
                    Dim paddedDel As String = delText & " "

                    ' a) insert silently
                    doc.TrackRevisions = False
                    targetRange.Text = delText

                    ' c) delete with tracking ON
                    doc.TrackRevisions = True
                    targetRange.Delete()

                    doc.TrackRevisions = False
                    targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                End If
            End While

        Catch ex As System.Exception
            Debug.WriteLine("InsertMarkupText error: " & ex.Message & vbCrLf & inputText)

        Finally
            doc.TrackRevisions = originalTrack
            wordApp.ScreenUpdating = originalUpdate

            ' Set targetRange to the full inserted text
            Dim endPosInserted As Integer = targetRange.End
            targetRange.SetRange(startPos, endPosInserted)

            ' Set the selection to targetRange
            wordApp.Selection.SetRange(targetRange.Start, targetRange.End)

        End Try
    End Sub

    ''' <summary>
    ''' Removes any “\* MERGEFORMAT” switch from inside {{…}} fields.
    ''' </summary>
    ''' <param name="input">Your full diff‑markup string.</param>
    ''' <returns>The same string, but with MERGEFORMAT gone from inside all {{…}}.</returns>
    Function RemoveMergeFormatFromBraces(input As String) As String
        ' Process each {{…}} as one chunk
        Return System.Text.RegularExpressions.Regex.Replace(
        input,
        "\{\{(.*?)\}\}",
        Function(m As System.Text.RegularExpressions.Match) As String
            ' m.Groups(1).Value is the interior of the braces
            Dim inner As String = m.Groups(1).Value
            ' remove any \* MERGEFORMAT (case‑insensitive)
            inner = System.Text.RegularExpressions.Regex.Replace(
                inner,
                "\\\*\s*MERGEFORMAT",
                String.Empty,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            )
            ' stitch it back together
            Return "{{" & inner & "}}"
        End Function,
        System.Text.RegularExpressions.RegexOptions.Singleline
    )
    End Function



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


    Private Structure PlaceholderInfo
        Public Offset As Integer     'offset relative to rng.Start (0-based)
        Public Length As Integer     'chars to skip in Word range
        Public Token As String       'replacement text ({{WFNT:…}}, {{PFOR:…}}, …)
    End Structure


    Private Shared ReadOnly PlaceholderComparer As Comparison(Of PlaceholderInfo) =
    Function(a, b)
        If a.Offset <> b.Offset Then
            Return a.Offset.CompareTo(b.Offset)
        End If
        Return a.Length.CompareTo(b.Length)
    End Function


    Public Function GetTextWithSpecialElementsInline(
        ByVal workingrange As Word.Range,
        PreserveParagraphFormatInline As Boolean, DoMarkdown As Boolean) As String

        Dim splash As New Slib.Splashscreen("Extracting text and format ...")
        splash.Show()
        splash.Refresh()

        Dim app As Word.Application = CType(workingrange.Application, Word.Application)
        Dim oldSU As Boolean = app.ScreenUpdating
        app.ScreenUpdating = False

        Try

            '──────────── 0)  Vorbereitung (Range klonen, Settings) ───────────────
            Dim rng As Word.Range = workingrange.Duplicate
            If rng.End < rng.Document.Content.End - 1 Then
                rng.End = rng.End + 1
            End If

            '──────────── Formatierungen vornehmen ────────────────────────────

            Debug.WriteLine($"4-1 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
            Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

            ' 0a) Markdown für Kombinationen & Einzelformate (mit CR-Handling)
            Dim origSel As Word.Range = app.Selection.Range.Duplicate

            ' Annahme: rng ist dein Ursprungsbereich (Word.Range)

            If DoMarkdown Then

                ' 1) Fett + Italic  (Absatz)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Font.Italic = True
                            f.Font.Underline = Word.WdUnderline.wdUnderlineNone
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "***\1***^13",
                        Sub(rep)                          ' nur Bold & Italic abstellen
                            rep.Bold = False
                            rep.Italic = False
                        End Sub)

                ' 2) Fett + Italic  (Inline)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Font.Italic = True
                            f.Font.Underline = Word.WdUnderline.wdUnderlineNone
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "***^&***",
                        Sub(rep)
                            rep.Bold = False
                            rep.Italic = False
                        End Sub)

                Debug.WriteLine($"4-2 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

                ' 3) Nur Fett  (Absatz)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "**\1**^13",
                        Sub(rep)
                            rep.Bold = False
                        End Sub)

                ' 4) Nur Fett  (Inline)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "**^&**",
                        Sub(rep)
                            rep.Bold = False
                        End Sub)

                Debug.WriteLine($"4-3 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")


                ' 5) Nur Italic  (Absatz)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Italic = True
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "*\1*^13",
                        Sub(rep)
                            rep.Italic = False
                        End Sub)

                ' 6) Nur Italic  (Inline)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Italic = True
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "*^&*",
                        Sub(rep)
                            rep.Italic = False
                        End Sub)

                Debug.WriteLine($"4-4 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")


                ' 7) Underline  (Absatz)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Underline = Word.WdUnderline.wdUnderlineSingle
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "<u>\1</u>^13",
                        Sub(rep)
                            rep.Underline = Word.WdUnderline.wdUnderlineNone
                        End Sub)

                ' 8) Underline  (Inline)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Underline = Word.WdUnderline.wdUnderlineSingle
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "<u>^&</u>",
                        Sub(rep)
                            rep.Underline = Word.WdUnderline.wdUnderlineNone
                        End Sub)

                ' 9) Strikethrough  (Absatz)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.StrikeThrough = True
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "~~\1~~^13",
                        Sub(rep)
                            rep.StrikeThrough = False
                        End Sub)

                '10) Strikethrough  (Inline)
                ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.StrikeThrough = True
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "~~^&~~",
                        Sub(rep)
                            rep.StrikeThrough = False
                        End Sub)

            End If

            ' Auswahl wiederherstellen
            'rng = workingrange.Duplicate

            rng.End = rng.End - 1
            rng.Select()

            Debug.WriteLine($"4-5 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
            Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")


            '──────────── Platzhalter vorbereiten ────────────────────────────

            Dim doc As Word.Document = workingrange.Application.ActiveDocument
            ' Einzigartigen Namen wählen, damit nichts kollidiert:
            Dim bmName As String = "__TMP_RNG_" & Guid.NewGuid().ToString("N")

            ' Bookmark anlegen (speichert Start & End)
            doc.Bookmarks.Add(Name:=bmName, Range:=rng)

            ' Umschalten
            With rng.TextRetrievalMode
                .IncludeHiddenText = True
                .IncludeFieldCodes = True
            End With

            ' Bookmark auslesen und rng zurücksetzen
            Dim bmRange As Word.Range = doc.Bookmarks(bmName).Range
            rng.SetRange(bmRange.Start, bmRange.End)

            ' Aufräumen
            doc.Bookmarks(bmName).Delete()

            'rng.TextRetrievalMode.IncludeHiddenText = True
            'rng.TextRetrievalMode.IncludeFieldCodes = True

            Dim placeholders As New List(Of PlaceholderInfo)

            '──────────── Fuß- & Endnoten sammeln ────────────────────────────
            For Each fn As Microsoft.Office.Interop.Word.Footnote In rng.Document.Footnotes
                If fn.Reference.Start >= rng.Start AndAlso fn.Reference.Start < rng.End Then
                    Dim s As Integer = system.math.max(fn.Reference.Start, rng.Start)
                    Dim e As Integer = system.math.min(fn.Reference.End, rng.End)
                    placeholders.Add(New PlaceholderInfo With {
                    .Offset = s - rng.Start,
                    .Length = e - s,
                    .Token = $"{{{{WFNT:{fn.Range.Text}}}}}"
                })
                End If
            Next

            For Each en As Microsoft.Office.Interop.Word.Endnote In rng.Document.Endnotes
                If en.Reference.Start >= rng.Start AndAlso en.Reference.Start < rng.End Then
                    Dim s As Integer = system.math.max(en.Reference.Start, rng.Start)
                    Dim e As Integer = system.math.min(en.Reference.End, rng.End)
                    placeholders.Add(New PlaceholderInfo With {
                    .Offset = s - rng.Start,
                    .Length = e - s,
                    .Token = $"{{{{WENT:{en.Range.Text}}}}}"
                })
                End If
            Next

            '──────────── Felder – GANZES Feld bestimmen ──────────────
            Const WD_FIELD_BEGIN As Integer = 19   'Chr(19)
            Const WD_FIELD_END As Integer = 21   'Chr(21)

            For Each fld As Word.Field In rng.Fields

                Dim codeText As String = fld.Code.Text.Trim()

                '----- A) exakten Feld-Begin ermitteln ----------------------------------
                Dim fldStartAbs As Integer = fld.Code.Start
                Do While fldStartAbs > rng.Start AndAlso
          AscW(rng.Characters(fldStartAbs - rng.Start + 1).Text) <> WD_FIELD_BEGIN
                    fldStartAbs -= 1
                Loop
                If AscW(rng.Characters(fldStartAbs - rng.Start + 1).Text) <> WD_FIELD_BEGIN Then _
        Continue For   'kein gültiger Begin im Range

                '----- B) Feld-Ende (0x15) suchen --------------------------------------
                Dim scanAbs As Integer = fldStartAbs
                Do While scanAbs < rng.End
                    Dim relIdx As Integer = scanAbs - rng.Start + 1
                    If AscW(rng.Characters(relIdx).Text) = WD_FIELD_END Then Exit Do
                    scanAbs += 1
                Loop

                'Fallback, falls das 0x15 außerhalb liegt
                If scanAbs >= rng.End Then
                    scanAbs = fld.Result.End
                    If scanAbs >= rng.End Then Continue For
                End If

                '----- C) Länge & Platzhalter ------------------------------------------
                Dim fldEndAbs As Integer = scanAbs
                Dim fldLength As Integer = fldEndAbs - fldStartAbs + 1   'inkl. 0x15

                placeholders.Add(New PlaceholderInfo With {
        .Offset = fldStartAbs - rng.Start,
        .Length = fldLength,
        .Token = $"{{{{WFLD:{codeText}}}}}"
    })
            Next

            '──────────── Absatz-Platzhalter (optional) ───────────────────────
            If PreserveParagraphFormatInline AndAlso rng.Paragraphs.Count > 0 Then

                Dim paraCount As Integer = rng.Paragraphs.Count
                ReDim paragraphFormat(paraCount - 1)
                Array.Clear(paragraphFormat, 0, paragraphFormat.Length)

                For i As Integer = 1 To paraCount
                    Dim p As Word.Paragraph = rng.Paragraphs(i)

                    'Absatzformate erfassen
                    Dim fmt As New ParagraphFormatStructure With {
                        .Style = p.Style,
                        .FontName = p.Range.Font.Name,
                        .FontSize = p.Range.Font.Size,
                        .FontBold = p.Range.Font.Bold,
                        .FontItalic = p.Range.Font.Italic,
                        .FontUnderline = p.Range.Font.Underline,
                        .FontColor = p.Range.Font.Color,
                        .ListType = p.Range.ListFormat.ListType,
                        .ListTemplate = If(p.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                                           p.Range.ListFormat.ListTemplate, Nothing),
                        .ListLevel = If(p.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                                           p.Range.ListFormat.ListLevelNumber, 0),
                        .ListNumber = If(p.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                                           p.Range.ListFormat.ListValue, 0),
                        .HasListFormat = p.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                        .Alignment = p.Alignment,
                        .LineSpacing = p.LineSpacing,
                        .SpaceBefore = p.SpaceBefore,
                        .SpaceAfter = p.SpaceAfter
                    }

                    paragraphFormat(i - 1) = fmt

                    placeholders.Add(New PlaceholderInfo With {
                    .Offset = p.Range.Start - rng.Start,   'Einfüge-Pos
                    .Length = 0,                           'überspringt nichts
                    .Token = $"{{{{PFOR:{i - 1}}}}}"
                })
                Next
            End If

            '──────────── Platzhalter sortieren (Offset ↑, Length ↑) ──────────

            Debug.WriteLine($"4-6 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
            Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")


            placeholders.Sort(PlaceholderComparer)

            Debug.WriteLine("placeholders: " & String.Join(", ", placeholders.Select(Function(ph) $"[Offset={ph.Offset}, Length={ph.Length}, Token={ph.Token}]")))

            ' ───── Platzhalter einfügen ────────────────────────────────
            Dim fullText As String = rng.Text
            Dim sbInline As New System.Text.StringBuilder(fullText.Length + placeholders.Count * 16)
            Dim lastPos As Integer = 0

            For Each ph As PlaceholderInfo In placeholders
                ' 1) Alles vor dem Platzhalter hinzufügen
                If ph.Offset > lastPos Then
                    sbInline.Append(fullText.Substring(lastPos, ph.Offset - lastPos))
                End If

                ' 2) Den Token einsetzen
                sbInline.Append(ph.Token)

                ' 3) Position nach dem Platzhalter merken
                lastPos = ph.Offset + ph.Length
            Next

            ' 4) Resttext anhängen
            If lastPos < fullText.Length Then
                sbInline.Append(fullText.Substring(lastPos))
            End If

            Debug.WriteLine($"4-7 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
            Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

            ' Ergebnis verwenden
            Return sbInline.ToString()


        Catch ex As System.Exception
            'Fail-Safe: reiner Text
            Debug.WriteLine("Error in GetTextWithSpecialElementsInline: " & ex.Message)
            Return workingrange.Text
        Finally
            app.ScreenUpdating = oldSU
            splash.Close()

        End Try
    End Function


    Private Function LegacyGetTextWithSpecialElementsInline(ByRef workingrange As Word.Range, PreserveParagraphFormatInline As Boolean) As String

        Debug.WriteLine("LegacyGetTextWithSpecialElementsInline called")

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


    Private Sub ReplaceWithinRange(
    ByVal rng As Word.Range,
    ByVal configureFind As Action(Of Word.Find),
    ByVal replacementText As String,
    ByVal tweakReplacement As Action(Of Word.Font))

        Dim doc As Word.Document = rng.Document
        Dim originalStart As Long = rng.Start
        Dim originalEnd As Long = rng.End
        Dim currentPosition As Long = originalStart

        Do
            ' Create a range from current position to the end of the original range
            Dim searchRange As Word.Range = doc.Range(currentPosition, originalEnd)
            Dim f As Word.Find = searchRange.Find

            Debug.WriteLine($"Searchrange = '{searchRange.Text}'")

            f.ClearFormatting()
            f.Replacement.ClearFormatting()

            configureFind(f)
            f.Replacement.Text = replacementText
            tweakReplacement(f.Replacement.Font)

            f.Forward = True
            f.Wrap = Word.WdFindWrap.wdFindStop
            f.Format = True

            ' If no more matches, exit
            If Not f.Execute(Replace:=Word.WdReplace.wdReplaceOne) Then Exit Do

            Debug.WriteLine($"Searchrange = '{searchRange.Text}' (after change)")

            ' After replacement, searchRange now points to the match
            ' Check if this match/replacement went beyond our boundary
            If searchRange.End > originalEnd Then
                Debug.WriteLine("Went too far!")
                doc.Undo()
                Exit Do
            End If

            ' Set the current position to continue from the end of this match
            currentPosition = searchRange.End
            originalEnd = rng.End

        Loop While currentPosition < originalEnd

        ' Update the original range to reflect the final processed area
        rng.SetRange(originalStart, originalEnd)
    End Sub


    Private Sub oldReplaceWithinRange(
        ByVal rng As Word.Range,
        ByVal configureFind As Action(Of Word.Find),
        ByVal replacementText As String,
        ByVal tweakReplacement As Action(Of Word.Font))

        Dim doc As Word.Document = rng.Document
        Dim startPos As Long = rng.Start
        Dim limitPos As Long = rng.End
        Dim cursor As Long = startPos

        Do
            Dim win As Word.Range = doc.Range(Start:=cursor, End:=limitPos)
            Dim f As Word.Find = win.Find
            Debug.WriteLine("Range: " & win.Text)

            f.ClearFormatting()
            f.Replacement.ClearFormatting()

            configureFind(f)                 ' Suchformat & -Text
            f.Replacement.Text = replacementText
            tweakReplacement(f.Replacement.Font)  ' nur gewünschtes Attribut zurücksetzen

            f.Forward = True
            f.Wrap = Word.WdFindWrap.wdFindStop
            f.Format = True

            If Not f.Execute(Replace:=Word.WdReplace.wdReplaceOne) Then Exit Do

            ' falls Ersatz über Limit hinausging → rückgängig & abbrechen
            If win.End > limitPos Then
                Debug.WriteLine("Went too far!")
                doc.Undo() : Exit Do
            End If

            cursor = win.End                 ' weiter hinter dem letzten Treffer
        Loop
    End Sub



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

        Dim splash As New Slib.Splashscreen("Accepting revisions related to formatting... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        ' Loop through all markups in the range
        For Each rev As Word.Revision In sel.Revisions

            System.Windows.Forms.Application.DoEvents()

            If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exit For

            If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
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

                    Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                    If NewDocChoice = 1 Then
                        Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                        Dim currentSelection As Word.Selection = newDoc.Application.Selection
                        currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        InsertTextWithMarkdown(currentSelection, OriginalText, True, True)
                    ElseIf NewDocChoice = 2 Then
                        Dim currentSelection As Word.Selection = Globals.ThisAddIn.Application.Selection
                        currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                        InsertTextWithMarkdown(currentSelection, OriginalText, False)
                    Else
                        ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                        SLib.PutInClipboard(MarkdownToRtfConverter.Convert((OriginalText)))
                    End If
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
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text in your document with which your selection in the pane shall be merged.")
            Return
        End If
        OtherPrompt = SLib.ShowCustomInputBox("If you want, you can amend the prompt that will be used to intelligently merge your selection into your document:", $"{AN} Intelligent Merge", False, SP_MergePrompt_Cached).Trim()
        If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return
        Dim result As String = Await ProcessSelectedText(OtherPrompt & " " & SP_Add_MergePrompt & " <INSERT>" & newtext & "</INSERT> ", True, INI_KeepFormat2, INI_KeepParaFormatInline, INI_ReplaceText2, INI_DoMarkupWord, INI_MarkupMethodWord, False, False, True, False, INI_KeepFormatCap)
    End Sub


    ''' <summary>
    ''' Collects comment text in <c>newtext</c>, selects either the anchor
    ''' text or its paragraph, and continues with your downstream processing.
    ''' </summary>
    Public Async Function BalloonMerge(
        ByVal selectWholeParagraph As Boolean, Silent As Boolean) As System.Threading.Tasks.Task

        Dim app As Word.Application = Globals.ThisAddIn.Application
        Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
        Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument

        Dim activeComment As Microsoft.Office.Interop.Word.Comment = Nothing
        Dim newtext As String = String.Empty

        Try
            '------------ 1) Find the comment the caret belongs to ------------------------
            If sel.StoryType = WdStoryType.wdCommentsStory Then
                For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                    If sel.Range.Start >= c.Range.Start AndAlso
                   sel.Range.End <= c.Range.End Then
                        activeComment = c : Exit For
                    End If
                Next
            Else
                For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                    Dim anchor As Range = c.Scope
                    If sel.Range.End > anchor.Start AndAlso
                   sel.Range.Start < anchor.End Then
                        activeComment = c : Exit For
                    End If
                Next
            End If

            '------------ 2) Quit if we are not in / on a comment -------------------------
            If activeComment Is Nothing Then
                ShowCustomMessageBox(
                "This command only works when the cursor is inside a comment " &
                "balloon or on text that has a comment.")
                Return
            End If

            '------------ 3) Determine what goes into newtext -----------------------------
            Dim selectedText As String = SafeRangeText(sel.Range)

            If sel.StoryType = WdStoryType.wdCommentsStory Then
                ' Inside balloon
                If selectedText.Trim().Length = 0 Then
                    newtext = SafeRangeText(activeComment.Range)      ' whole balloon
                    sel.SetRange(activeComment.Range.Start, activeComment.Range.End)
                Else
                    newtext = selectedText                            ' user selection
                End If
            Else
                ' On anchor in main story
                If selectedText.Trim().Length = 0 Then
                    newtext = SafeRangeText(activeComment.Range)      ' whole balloon
                Else
                    newtext = selectedText                            ' user selection
                End If
            End If

            '------------ 4) Adjust the selection in the main story -----------------------
            Dim anchorRange As Range = activeComment.Scope
            Dim targetRange As Range

            If selectWholeParagraph Then
                If anchorRange.Paragraphs.Count > 0 Then
                    '–– The anchor spans ≥ 1 paragraphs → select them ALL ––
                    Dim firstPara As Range = anchorRange.Paragraphs(1).Range
                    Dim lastPara As Range = anchorRange.Paragraphs(anchorRange.Paragraphs.Count).Range
                    targetRange = doc.Range(firstPara.Start, lastPara.End)
                Else
                    '–– Collapsed anchor (no text selected when comment was made) ––
                    '   Select the paragraph where the anchor is located.
                    targetRange = doc.Range(anchorRange.Start, anchorRange.Start).Paragraphs(1).Range
                End If
            Else
                ' Only the exact anchor text
                targetRange = anchorRange
            End If

            targetRange.Select()

            '------------ 5) Your downstream processing -----------------------------------
            If Not Silent Or String.IsNullOrWhiteSpace(SP_MergePrompt2) Then
                OtherPrompt = SLib.ShowCustomInputBox(
                "If you want, you can amend the prompt that will be used to " &
                "intelligently merge your comment into your document:",
                $"{AN} Intelligent Merge", False, SP_MergePrompt2).Trim()
                If String.IsNullOrEmpty(OtherPrompt) OrElse OtherPrompt = "ESC" Then Return
            Else
                OtherPrompt = SP_MergePrompt2
            End If

            Dim items = {
                New SelectionItem("Word", 1),
                New SelectionItem("Diff", 2),
                New SelectionItem("Diff Window", 3),
                New SelectionItem("Regex", 4),
                New SelectionItem("None", 5)
                }

            Dim DefaultItem As Integer = 5
            If INI_DoMarkupWord Then DefaultItem = INI_MarkupMethodWord
            Dim picked As Integer = SelectValue(items, DefaultItem, "Choose markup method ...")

            If picked < 1 Then Return

            Dim result As String = Await ProcessSelectedText(
            OtherPrompt & " " & SP_Add_MergePrompt & " <INSERT>" &
            newtext & "</INSERT> ",
            True, INI_KeepFormat2, INI_KeepParaFormatInline,
            INI_ReplaceText2, If(picked < 5, True, False), If(picked < 5, picked, INI_MarkupMethodWord),
            False, False, True, False, INI_KeepFormatCap)

        Catch ex As System.Exception
            MessageBox.Show(
            $"Error in BalloonMerge:{Environment.NewLine}{ex.Message}",
            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    ''' <summary>Returns <c>r.Text</c> or an empty string when Word gives back Nothing.</summary>
    Private Function SafeRangeText(r As Word.Range) As String
        If r Is Nothing Then Return String.Empty
        Try
            Dim t As String = r.Text          ' can be Nothing in edge-cases
            If t Is Nothing Then t = String.Empty
            Return t
        Catch
            ' extremely rare: r.Text itself can throw in corrupt docs
            Return String.Empty
        End Try
    End Function


    Public Async Sub IntelligentMergeBalloon(newtext As String)
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
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
                SLib.ShowCustomMessageBox($"Error loading And initializing the NER model ({ex.Message}).")
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
                Dim result As String = ShowCustomWindow("The anonymization returned the following text:", AnonText, $"Beware that this anonymization depends entirely on the keys you provided in your file '{AnonFile}' (for your model '{INI_Model}') or your prompt. Check the result. Choose what to put into the clipboard.", $"{AN} Anonymization", False)

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

        Dim FinalText As String = ShowCustomWindow("The NER anonymization returned the following text:", sb.ToString(), "Beware that this anonymization method is fast, but not of very high precision. Check the result.", AN, False)

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
            Dim userNames As New Microsoft.VisualBasic.Collection
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


            ' ————————————————————————————————————————————————————————————
            ' Prompt für earliest date
            Dim userDateInput As String
            Dim earliestDate As System.DateTime = System.DateTime.MinValue
            Dim earliestDateFiltered As Boolean = False

            userDateInput = ShowCustomInputBox(
                    "Please enter the earliest date (and time, if you wish) to consider (leave empty for no filter):",
                    "Markup Time Span",
                    True,
                    System.DateTime.Now.AddDays(-2).ToString(System.Globalization.CultureInfo.CurrentCulture)
                )
            userDateInput = userDateInput.Trim()

            Dim parsed As System.DateTime
            If String.IsNullOrEmpty(userDateInput) Then
                earliestDateFiltered = False
            ElseIf System.DateTime.TryParse(
                      userDateInput,
                      System.Globalization.CultureInfo.CurrentCulture,
                      System.Globalization.DateTimeStyles.None,
                      parsed
                  ) Then
                earliestDate = parsed
                earliestDateFiltered = True
            Else
                ShowCustomMessageBox("Improper date/time format - exiting.")
                Exit Sub
            End If

            ' ————————————————————————————————————————————————————————————


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
                If (String.IsNullOrEmpty(userInput) OrElse rev.Author.Equals(userInput, StringComparison.OrdinalIgnoreCase)) _
                       AndAlso (Not earliestDateFiltered OrElse rev.Date >= earliestDate) Then
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
                If (String.IsNullOrEmpty(userInput) OrElse comment.Author.Equals(userInput, StringComparison.OrdinalIgnoreCase)) _
                       AndAlso (Not earliestDateFiltered OrElse comment.Date >= earliestDate) Then

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
                timeSpan = System.Math.Floor(timeDiff / 1440).ToString() & " days, " &
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
                ShowCustomMessageBox(outputUserNames & vbCrLf & If(earliestDateFiltered, "Earliest considered: " & earliestDate.ToString("dd/MM/yyyy HH:mm") & vbCrLf, "") & "First markup/comment: " & formattedFirstTimestamp & vbCrLf &
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

        If INILoadFail() Then Return

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

                Dim splash As New Slib.Splashscreen($"Highlighting {parts.Count} hits... Press 'Esc' to abort")
                splash.Show()
                splash.Refresh()

                Dim Aborted As Boolean = False

                Dim trackChangesEnabled As Boolean = doc.TrackRevisions
                Dim originalAuthor As String = doc.Application.UserName

                doc.TrackRevisions = True

                Dim SuccessHits As Integer = 0

                For Each part As String In parts

                    splash.UpdateMessage($"Highlighting {parts.Count - SuccessHits} hits... Press 'Esc' to abort")

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                        Aborted = True
                        Exit For
                    End If

                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                        ' Exit the loop
                        Aborted = True
                        Exit For
                    End If

                    Dim findText As String = part.Trim()
                    If FindLongTextInChunks(findText, SearchChunkSize, selection) And selection IsNot Nothing Then
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
                    Dim errorlist As String = ShowCustomWindow($"{SuccessHits} hit(s) have been highlighted using Context Search. The following hit(s) could not be found:", String.Join(vbCrLf, notFoundParts), "The above error list will be included in a final comment at the end of your last hit (it will also be included in the clipboard). You can have the original list included, or you can now make changes and have this version used. If you select Cancel, nothing will be put added to the document.", AN, True)
                    If errorlist <> "" And errorlist.ToLower() <> "esc" Then
                        SLib.PutInClipboard(errorlist)
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, $"{AN5} could not locate these sections: " & vbCrLf & errorlist)
                    End If
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

                If FindLongTextInChunks(FindText, SearchChunkSize, selection) And selection IsNot Nothing Then
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
                {"Clean", "Clean the LLM response"},
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
                {"Clean", "To remove double-spaces and hidden markers that may have been inserted by the LLM"},
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

        Dim splash As New Slib.Splashscreen("Updating menu following your changes ...")
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
            filePath = System.IO.Path.GetFullPath(filePath)
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
            filePath = System.IO.Path.GetFullPath(filePath)
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



    '--- New Web Integration ---

    Private httpListener As System.Net.HttpListener
    Private listenerTask As System.Threading.Tasks.Task   ' replaces the raw Thread
    Private isShuttingDown As Boolean = False

    '───────────────────────────────────────────────────────────────────────────
    ' Run a Sub on the UI thread and *wait* for it to finish.
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

        Return tcs.Task
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ' Run a Func(Of T) on the UI thread and wait for its return value.
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

        Return tcs.Task
    End Function


    Private Sub StartupHttpListener()
        ' fire-and-forget – no raw Thread needed
        listenerTask = StartHttpListener()      ' captures the returned Task
    End Sub

    Private Sub ShutdownHttpListener()
        isShuttingDown = True
        If httpListener IsNot Nothing AndAlso httpListener.IsListening Then
            httpListener.Stop()
            httpListener.Close()
        End If
    End Sub


    Private Async Function StartHttpListener() As System.Threading.Tasks.Task
        Const prefix As String = "http://127.0.0.1:12334/"   ' ← Word gets its own port
        Dim consecutiveFailures As Integer = 0

        While Not isShuttingDown
            Try
                ' ensure listener exists and is running
                If httpListener Is Nothing Then
                    httpListener = New System.Net.HttpListener()
                    httpListener.Prefixes.Add(prefix)
                    httpListener.Start()
                    Debug.WriteLine("HttpListener started.")
                ElseIf Not httpListener.IsListening Then
                    httpListener.Close()
                    httpListener = Nothing
                    Continue While                      ' next loop restarts it
                End If

                ' wait for one incoming request
                Dim ctx As System.Net.HttpListenerContext =
                Await httpListener.GetContextAsync().ConfigureAwait(False)

                ' handle the request (fire-and-forget)
                Call HandleHttpRequest(ctx) _
                .ContinueWith(
                    Sub(t)
                        If t.IsFaulted AndAlso t.Exception IsNot Nothing Then
                            Debug.WriteLine("HandleHttpRequest error: " &
                                            t.Exception.GetBaseException().Message)
                        End If
                    End Sub,
                    System.Threading.Tasks.TaskScheduler.Default)

                consecutiveFailures = 0                       ' success
            Catch ex As System.ObjectDisposedException
                consecutiveFailures += 1
            Catch ex As System.Exception
                consecutiveFailures += 1
                Debug.WriteLine("Listener error: " & ex.Message)
            End Try

            ' recycle after too many consecutive errors
            If consecutiveFailures >= 10 AndAlso Not isShuttingDown Then
                Debug.WriteLine("Restarting HttpListener after 10 failures.")
                Try
                    If httpListener IsNot Nothing Then httpListener.Close()
                Catch
                End Try
                httpListener = Nothing
                consecutiveFailures = 0
                Await System.Threading.Tasks.Task.Delay(5000).ConfigureAwait(False)
            End If
        End While
    End Function


    Private Async Function HandleHttpRequest(
        ctx As System.Net.HttpListenerContext) _
        As System.Threading.Tasks.Task

        Dim req = ctx.Request
        Dim res = ctx.Response

        '─── CORS pre-flight────────────────────────────────────────────────────
        If req.HttpMethod = "OPTIONS" Then
            res.AddHeader("Access-Control-Allow-Origin", "*")
            res.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
            res.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization")
            res.StatusCode = 204 : res.Close() : Return
        End If

        '─── Read body (if any)─────────────────────────────────────────────────
        Dim body As String = ""
        If req.HasEntityBody Then
            Using rdr As New IO.StreamReader(req.InputStream, System.Text.Encoding.UTF8)
                body = Await rdr.ReadToEndAsync().ConfigureAwait(False)
            End Using
        End If

        '─── Dispatch to our add-in logic───────────────────────────────────────
        Dim responseText As String =
        Await ProcessRequestInAddIn(body, req.RawUrl).ConfigureAwait(False)

        '─── Send response──────────────────────────────────────────────────────
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
    ' MAIN REQUEST DISPATCH (Word – only "redink_sendtoword")
    ' ---------------------------------------------------------------------------
    Private Async Function ProcessRequestInAddIn(
        body As String,
        rawUrl As String) _
        As System.Threading.Tasks.Task(Of String)

        ' guard clause – empty body
        If String.IsNullOrWhiteSpace(body) Then Return ""

        Dim j = Newtonsoft.Json.Linq.JObject.Parse(body)
        Dim cmd = j("Command")?.ToString()
        Dim textBody = j("Text")?.ToString()
        Dim sourceUrl = j("URL")?.ToString()

        Select Case cmd
        '───────────────────────────────────────────────────────────────────
            Case "redink_sendtoword"
                If String.IsNullOrWhiteSpace(textBody) Then Return ""

                ' Everything that touches Word must run on the UI thread
                Await SwitchToUi(Sub()

                                     Dim wdApp As Microsoft.Office.Interop.Word.Application =
                    Globals.ThisAddIn.Application

                                     Dim sel As Microsoft.Office.Interop.Word.Selection = wdApp.Selection

                                     wdApp.ScreenUpdating = False
                                     sel.TypeText(textBody & " (" & sourceUrl & ")")
                                     wdApp.ScreenUpdating = True

                                     ' Release COM objects explicitly (good hygiene)
                                     System.Runtime.InteropServices.Marshal.ReleaseComObject(sel)
                                 End Sub)

                Return ""      ' nothing needs to be sent back in this scenario
        End Select

        Return ""              ' unknown command → no-op
    End Function


    Public Class TranscriptionForm

        Inherits Form

        ' --- P/Invoke for SetThreadExecutionState ---
        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function SetThreadExecutionState(ByVal esFlags As UInteger) As UInteger
        End Function

        ' Constants for sleep prevention
        Private Const ES_CONTINUOUS As UInteger = &H80000000UI
        Private Const ES_SYSTEM_REQUIRED As UInteger = &H1UI
        Private Const ES_DISPLAY_REQUIRED As UInteger = &H2UI ' Optional: keeps the display on too

        Private _iSetTheSleepLock As Boolean = False

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
                    Dim dirName As String = System.IO.Path.GetFileName(dir)
                    If dirName.StartsWith("vosk-model") Then
                        cultureComboBox.Items.Add(dirName)
                        modelsexist = True
                    End If
                Next

                For Each file As String In Directory.GetFiles(modelPath)
                    Dim fileName As String = System.IO.Path.GetFileName(file)
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
            Me.MinimumSize = New System.Drawing.Size(800, 440)

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

            ' Only release the sleep lock IF we were the ones who set it.
            If _iSetTheSleepLock Then
                ' We are responsible, so we release the lock.
                SetThreadExecutionState(ES_CONTINUOUS)
                _iSetTheSleepLock = False ' Reset our flag
                Debug.WriteLine("This form released the sleep lock.")
            Else
                ' We are not responsible, so we do nothing to the execution state.
                Debug.WriteLine("Another component is managing the sleep lock. This form took no action.")
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

            Dim splash As New Slib.Splashscreen($"Loading model...")
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

                        splash = New Slib.Splashscreen($"Transcribing file ...")
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
                    minSpeakers = system.math.max(2, minSpeakers)
                    maxSpeakers = system.math.max(minSpeakers, maxSpeakers)

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

            Dim splash As New Slib.Splashscreen($"Loading model...")
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

            ' Always request that the system stay awake.
            ' The function returns the PREVIOUS state.
            Dim previousState As UInteger = SetThreadExecutionState(ES_CONTINUOUS Or ES_SYSTEM_REQUIRED)

            ' Now, check if the SYSTEM_REQUIRED flag was already set in the previous state.
            ' We use a bitwise AND. If the result is 0, the flag was NOT set before our call.
            If (previousState And ES_SYSTEM_REQUIRED) = 0 Then
                ' The lock was NOT active before. Therefore, *we* are responsible for releasing it later.
                _iSetTheSleepLock = True
                Debug.WriteLine("Sleep lock was not active. This form has now acquired it.")
            Else
                ' The lock was ALREADY active. We are not responsible for releasing it.
                _iSetTheSleepLock = False
                Debug.WriteLine("Sleep lock was already active. This form will not release it.")
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
            Dim maxSample As Single = samples.Max(Function(x) System.Math.Abs(x))
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

        Private Async Function ProcessWhisper(samples As Single()) As System.Threading.Tasks.Task
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

        Public Async Function WhisperTranscribeAudioFile(filepath As String) As System.Threading.Tasks.Task

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
                Dim endPos = system.math.min(offset + sliceSize, pcmData.Length)
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
                Dim len = system.math.min(chunkSize, pcmData.Length - pos)
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


        Public Async Function VoskTranscribeAudioFile(filepath As String) As System.Threading.Tasks.Task
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

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                        Exited = True
                        Exit While
                    End If

                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                        ' Exit the loop
                        Exited = True
                        Exit While
                    End If

                    Dim chunkLength As Integer = system.math.min(chunkSize, pcmData.Length - offset)
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
            Dim norm As Double = System.Math.Sqrt(embedding.Sum(Function(x) x * x))
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
            Return System.Math.Sqrt(sum)
        End Function


        ' Function to compute cosine similarity between two speaker embeddings
        Private Function CosineSimilarity(vec1 As List(Of Double), vec2 As List(Of Double)) As Double
            Dim dotProduct As Double = vec1.Zip(vec2, Function(a, b) a * b).Sum()
            Dim magnitude1 As Double = System.Math.Sqrt(vec1.Sum(Function(a) a * a))
            Dim magnitude2 As Double = System.Math.Sqrt(vec2.Sum(Function(b) b * b))

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

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetThreadExecutionState(ByVal esFlags As UInteger) As UInteger
    End Function

    Private Const ES_CONTINUOUS As UInteger = &H80000000UI
    Private Const ES_SYSTEM_REQUIRED As UInteger = &H1UI

    ' Flag to track if the TTS engine is responsible for the current sleep lock.
    Private Shared _ttsAcquiredTheSleepLock As Boolean = False

    ''' <summary>
    ''' Cooperatively acquires a system sleep lock for TTS operations.
    ''' It checks if a lock is already active before taking responsibility for it.
    ''' </summary>
    Public Shared Sub AcquireTTSSleepLock()
        ' Always request that the system stay awake.
        ' The function returns the PREVIOUS state.
        Dim previousState As UInteger = SetThreadExecutionState(ES_CONTINUOUS Or ES_SYSTEM_REQUIRED)

        ' Check if the SYSTEM_REQUIRED flag was already set in the previous state.
        If (previousState And ES_SYSTEM_REQUIRED) = 0 Then
            ' The lock was NOT active before. Therefore, the TTS engine is now responsible.
            _ttsAcquiredTheSleepLock = True
            Debug.WriteLine("[TTS] Sleep lock was not active. TTS has now acquired it.")
        Else
            ' The lock was ALREADY active. The TTS engine is not responsible for releasing it.
            _ttsAcquiredTheSleepLock = False
            Debug.WriteLine("[TTS] Sleep lock was already active. TTS will not release it.")
        End If
    End Sub

    ''' <summary>
    ''' Cooperatively releases the system sleep lock, but only if the TTS
    ''' engine was the component that originally acquired it.
    ''' </summary>
    Public Shared Sub ReleaseTTSSleepLock()
        ' Only release the sleep lock IF we were the ones who set it.
        If _ttsAcquiredTheSleepLock Then
            ' We are responsible, so we release the lock.
            SetThreadExecutionState(ES_CONTINUOUS)
            _ttsAcquiredTheSleepLock = False ' Reset our flag
            Debug.WriteLine("[TTS] TTS has released the sleep lock.")
        Else
            ' We are not responsible, so we do nothing.
            Debug.WriteLine("[TTS] Another component is managing the sleep lock. TTS took no action.")
        End If
    End Sub

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

        AcquireTTSSleepLock()

        Try

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
            ' If IsNothing(prevExecState) Then ...
            ReleaseTTSSleepLock()
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

            Dim outputFiles As New List(Of String)

            ' ensure a valid output path
            If String.IsNullOrWhiteSpace(filepath) Then
                filepath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
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

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exited = True : Exit For
                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then Exited = True : Exit For

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
                            audioBytes = System.Convert.FromBase64String(respJson("audioContent").ToString())
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
                    Dim tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{AN2}_podcast_temp_{i}.mp3")
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


        Dim splash As New Slib.Splashscreen($"Playing MP3... press 'Esc' to abort")
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
                            If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                                Exit While
                            End If
                            If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
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
            Dim selection As Microsoft.Office.Interop.Word.Selection = app.Selection
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
            For Each para As Microsoft.Office.Interop.Word.Paragraph In selection.Paragraphs
                ' Allow the user to abort by pressing Escape.
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Or (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or ProgressBarModule.CancelOperation Then
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
                    Dim tempParaFile As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{AN2}_temp_para_{paragraphIndex}.mp3")
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
            Dim tempFile As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{AN2}_silence_{CInt(durationSeconds * 1000)}ms.mp3")

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

        Private Async Function GetVoicesByLanguageAsync(languageCode As String) As System.Threading.Tasks.Task(Of List(Of GoogleVoice))
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

        Private Async Function PlaySelectedVoiceAsync(cmbLang As Forms.ComboBox, cmbVoice As Forms.ComboBox) As System.Threading.Tasks.Task
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
                Await System.Threading.Tasks.Task.Run(Sub()
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
                SelectedOutputPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
            ElseIf SelectedOutputPath.EndsWith("\") OrElse SelectedOutputPath.EndsWith("/") Then
                ' If only a folder is given, append default filename
                SelectedOutputPath = System.IO.Path.Combine(SelectedOutputPath, TTSDefaultFile)
            Else
                Dim dir As String = System.IO.Path.GetDirectoryName(SelectedOutputPath)
                Dim fileName As String = System.IO.Path.GetFileName(SelectedOutputPath)

                ' If no directory is found, assume Desktop as the base
                If String.IsNullOrWhiteSpace(dir) Then
                    SelectedOutputPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName)
                    dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                End If

                ' If no filename is given, use the default filename
                If String.IsNullOrWhiteSpace(fileName) Then
                    SelectedOutputPath = System.IO.Path.Combine(dir, TTSDefaultFile)
                End If

                ' Ensure the filename has ".mp3" extension
                If Not fileName.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) Then
                    SelectedOutputPath = System.IO.Path.Combine(dir, fileName & ".mp3")
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
            Dim fileName As String = System.IO.Path.GetFileName(txtOutputPath.Text)

            ' Get the user's Desktop path
            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

            ' Construct new file path
            txtOutputPath.Text = System.IO.Path.Combine(desktopPath, fileName)

        End Sub

    End Class



    ' Code für "Slides"

    ' Liest die bestehende Präsentation aus

    Public Function GetPresentationJson(pptxPath As String) As String
        ' 0) Path check
        If Not System.IO.File.Exists(pptxPath) Then
            ShowCustomMessageBox($"File not found: {pptxPath}")
            Return String.Empty
        End If

        Try
            ' 1) Try to open the presentation
            Using presDoc As DocumentFormat.OpenXml.Packaging.PresentationDocument =
            DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, False)

                Dim presPart As DocumentFormat.OpenXml.Packaging.PresentationPart = presDoc.PresentationPart
                If presPart Is Nothing OrElse presPart.Presentation Is Nothing Then
                    ShowCustomMessageBox("Invalid or corrupted presentation.")
                    Return String.Empty
                End If

                Dim result As New PresentationJson With {
                .Title = presDoc.PackageProperties.Title,
                .Slides = New List(Of SlideJson)(),
                .Layouts = New List(Of LayoutJson)()
            }

                ' --- START OF ADDED CODE ---
                ' Add slide dimensions to the result object for the LLM.
                If presPart.Presentation.SlideSize IsNot Nothing AndAlso
               presPart.Presentation.SlideSize.Cx IsNot Nothing AndAlso
               presPart.Presentation.SlideSize.Cy IsNot Nothing Then

                    result.SlideSize = New SlideSizeJson With {
                    .Width = presPart.Presentation.SlideSize.Cx.Value,
                    .Height = presPart.Presentation.SlideSize.Cy.Value
                }
                End If
                ' --- END OF ADDED CODE ---

                ' 2) Extract slides (This part remains unchanged)
                Dim slideIdList = presPart.Presentation.SlideIdList
                If slideIdList IsNot Nothing Then
                    Dim idx As Integer = 0
                    For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In slideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                        Dim sp As DocumentFormat.OpenXml.Packaging.SlidePart =
                        CType(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)

                        Dim title As String = GetSlideTitle(sp)
                        Dim key As String = If(
                        String.IsNullOrWhiteSpace(title),
                        $"SID-{sid.Id.Value}",
                        $"{SanitizeKey(title)}-{sid.Id.Value}"
                    )

                        Dim layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = sp.SlideLayoutPart
                        Dim layoutName As String = GetLayoutName(layoutPart)
                        Dim masterName As String = If(
                        layoutPart IsNot Nothing,
                        GetMasterName(layoutPart.SlideMasterPart),
                        String.Empty
                    )

                        Dim placeholders As New List(Of String)
                        Dim content As New List(Of String)
                        If sp.Slide IsNot Nothing AndAlso
                       sp.Slide.CommonSlideData IsNot Nothing AndAlso
                       sp.Slide.CommonSlideData.ShapeTree IsNot Nothing Then

                            For Each shp As DocumentFormat.OpenXml.Presentation.Shape In
                            sp.Slide.CommonSlideData.ShapeTree.OfType(Of DocumentFormat.OpenXml.Presentation.Shape)()

                                If shp.TextBody IsNot Nothing Then
                                    content.Add(shp.TextBody.InnerText.Trim())
                                End If

                                Dim nv = shp.NonVisualShapeProperties
                                If nv IsNot Nothing AndAlso nv.ApplicationNonVisualDrawingProperties IsNot Nothing Then
                                    Dim ph = nv.ApplicationNonVisualDrawingProperties.PlaceholderShape
                                    If ph IsNot Nothing AndAlso ph.Type IsNot Nothing Then
                                        placeholders.Add(ph.Type.Value.ToString())
                                    End If
                                End If
                            Next
                        End If

                        result.Slides.Add(New SlideJson With {
                        .SlideKey = key,
                        .SlideId = sid.Id.Value,
                        .Index = idx,
                        .Title = title,
                        .Layout = layoutName,
                        .Master = masterName,
                        .Placeholders = placeholders,
                        .Content = content
                    })
                        idx += 1
                    Next
                End If

                ' 3) Extract layouts (This part remains unchanged)
                For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
                    For Each layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                        Dim name As String = GetLayoutName(layoutPart)
                        Dim layoutUri As String = layoutPart.Uri.ToString()
                        Dim relId As String = sm.GetIdOfPart(layoutPart)

                        result.Layouts.Add(New LayoutJson With {
                        .Name = name,
                        .LayoutId = layoutUri,
                        .LayoutRelId = relId
                    })
                    Next
                Next

                ' 4) Serialize (This part remains unchanged)
                Return System.Text.Json.JsonSerializer.Serialize(
                result,
                New System.Text.Json.JsonSerializerOptions With {.WriteIndented = True}
            )
            End Using

        Catch ex As System.IO.IOException
            ' I/O errors when opening the file
            ShowCustomMessageBox($"Error opening presentation (I/O): {ex.Message}")
            Return String.Empty

        Catch ex As DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
            ' OpenXML package parsing errors
            ShowCustomMessageBox($"Error processing presentation (OpenXML): {ex.Message}")
            Return String.Empty

        Catch ex As System.Exception
            ' All other unexpected errors
            ShowCustomMessageBox($"Unexpected error: {ex.Message}")
            Return String.Empty
        End Try
    End Function


    Public Class SlideJson
        <JsonPropertyName("slideKey")>
        Public Property SlideKey As String

        <JsonPropertyName("slideId")>
        Public Property SlideId As UInteger

        <JsonPropertyName("index")>
        Public Property Index As Integer

        <JsonPropertyName("title")>
        Public Property Title As String

        <JsonPropertyName("layout")>
        Public Property Layout As String

        <JsonPropertyName("master")>
        Public Property Master As String

        <JsonPropertyName("placeholders")>
        Public Property Placeholders As List(Of String)

        <JsonPropertyName("content")>
        Public Property Content As List(Of String)
    End Class

    Public Class LayoutJson
        <JsonPropertyName("name")>
        Public Property Name As String

        <JsonPropertyName("layoutId")>
        Public Property LayoutId As String

        <JsonPropertyName("layoutRelId")>
        Public Property LayoutRelId As String
    End Class

    Public Class SlideSizeJson
        <JsonPropertyName("width")>
        Public Property Width As Long

        <JsonPropertyName("height")>
        Public Property Height As Long
    End Class

    ' [MODIFIED CLASS] The main DTO, now with a property for slide dimensions.
    Public Class PresentationJson
        <JsonPropertyName("title")>
        Public Property Title As String

        ' [NEW PROPERTY]
        <JsonPropertyName("slideSize")>
        Public Property SlideSize As SlideSizeJson

        <JsonPropertyName("slides")>
        Public Property Slides As List(Of SlideJson)

        <JsonPropertyName("layouts")>
        Public Property Layouts As List(Of LayoutJson)
    End Class


    ' --- Hilfsfunktionen ---
    Private Function GetSlideTitle(sp As SlidePart) As String
        If sp.Slide Is Nothing OrElse
           sp.Slide.CommonSlideData Is Nothing OrElse
           sp.Slide.CommonSlideData.ShapeTree Is Nothing Then
            Return String.Empty
        End If

        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In
            sp.Slide.CommonSlideData.ShapeTree.OfType(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim nv = shp.NonVisualShapeProperties
            If nv IsNot Nothing AndAlso nv.ApplicationNonVisualDrawingProperties IsNot Nothing Then
                Dim ph = nv.ApplicationNonVisualDrawingProperties.PlaceholderShape
                If ph IsNot Nothing AndAlso
                   (ph.Type Is Nothing OrElse
                    ph.Type.Value = PlaceholderValues.Title OrElse
                    ph.Type.Value = PlaceholderValues.CenteredTitle) Then
                    Return shp.TextBody.InnerText
                End If
            End If
        Next

        Return String.Empty
    End Function

    Private Function GetLayoutName(layoutPart As SlideLayoutPart) As String
        If layoutPart IsNot Nothing AndAlso
           layoutPart.SlideLayout IsNot Nothing AndAlso
           layoutPart.SlideLayout.CommonSlideData IsNot Nothing Then

            Dim nm = layoutPart.SlideLayout.CommonSlideData.Name
            If Not String.IsNullOrWhiteSpace(nm) Then
                Return nm
            End If
        End If
        Return layoutPart.Uri.ToString()
    End Function

    Private Function GetMasterName(smPart As SlideMasterPart) As String
        If smPart IsNot Nothing AndAlso
           smPart.SlideMaster IsNot Nothing AndAlso
           smPart.SlideMaster.CommonSlideData IsNot Nothing Then

            Dim nm = smPart.SlideMaster.CommonSlideData.Name
            If Not String.IsNullOrWhiteSpace(nm) Then
                Return nm
            End If
        End If
        Return smPart.Uri.ToString()
    End Function

    Private Function SanitizeKey(s As String) As String
        Return New String(
            s.Select(Function(ch) If(Char.IsLetterOrDigit(ch), ch, "-"c)).ToArray()
        )
    End Function


    ' 1) DTOs & Polymorph-Converter (verkürzt)
    Public MustInherit Class ActionBase
        <JsonPropertyName("op")>
        Public Property Op As String
    End Class

    Public Class Anchor
        <JsonPropertyName("mode")>
        Public Property Mode As String
        <JsonPropertyName("by")>
        Public Property By As AnchorBy
    End Class
    Public Class AnchorBy
        <JsonPropertyName("slideKey")>
        Public Property SlideKey As String
    End Class

    Public Class AddSlideAction
        Inherits ActionBase
        <JsonPropertyName("anchor")> Public Property Anchor As Anchor
        <JsonPropertyName("layoutRelId")> Public Property LayoutRelId As String
        <JsonPropertyName("elements")> Public Property Elements As List(Of JsonElement)
    End Class



    Public Function CleanJsonString(raw As String) As String
        If String.IsNullOrEmpty(raw) Then
            Return String.Empty
        End If

        ' Look for object vs. array start
        Dim firstObj = raw.IndexOf("{"c)
        Dim firstArr = raw.IndexOf("["c)
        Dim startIdx As Integer
        Dim openChar As Char
        Dim closeChar As Char

        If firstObj >= 0 AndAlso (firstObj < firstArr OrElse firstArr = -1) Then
            startIdx = firstObj
            openChar = "{"c
            closeChar = "}"c
        ElseIf firstArr >= 0 Then
            startIdx = firstArr
            openChar = "["c
            closeChar = "]"c
        Else
            ' No JSON delimiters found – just return trimmed
            Return raw.Trim()
        End If

        ' Find the last matching closing brace/bracket
        Dim lastIdx = raw.LastIndexOf(closeChar)
        If lastIdx > startIdx Then
            Return raw.Substring(startIdx, lastIdx - startIdx + 1).Trim()
        Else
            ' Malformed or unmatched – fallback to trimming
            Return raw.Trim()
        End If
    End Function





    Public Function ApplyPlanToPresentation(pptxPath As String, planJson As String) As Boolean
        Try
            ' 1) Check if the file exists
            If Not System.IO.File.Exists(pptxPath) Then
                ShowCustomMessageBox($"Your file '{pptxPath}' was no longer found - aborting.")
                Return False
            End If

            ' 2) Configure JSON serializer options
            Dim opts As New System.Text.Json.JsonSerializerOptions With {
            .PropertyNameCaseInsensitive = True
        }
            opts.Converters.Add(New System.Text.Json.Serialization.JsonStringEnumConverter())

            Dim actions As System.Text.Json.JsonElement.ArrayEnumerator
            Try
                actions = System.Text.Json.JsonDocument.Parse(planJson) _
                      .RootElement _
                      .GetProperty("actions") _
                      .EnumerateArray()
            Catch ex As System.Text.Json.JsonException
                ShowCustomMessageBox("The AI has sent an invalid instruction on how to build the slides: " & ex.Message)
                Return False
            Catch ex As KeyNotFoundException
                ShowCustomMessageBox("An internal error occurred when amending your slidedeck (the AI sent instructions missing the required 'actions' array).")
                Return False
            End Try

            Dim errorMessages As New List(Of String)

            ' 3) Open the presentation
            Using presDoc As DocumentFormat.OpenXml.Packaging.PresentationDocument =
              DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, True)

                Dim presPart As DocumentFormat.OpenXml.Packaging.PresentationPart = presDoc.PresentationPart
                If presPart Is Nothing Then
                    ShowCustomMessageBox("A presentation is missing in the file you have provided; you may have to include at least one slide.")
                    Return False
                End If

                ' Ensure SlideIdList exists
                If presPart.Presentation.SlideIdList Is Nothing Then
                    presPart.Presentation.AppendChild(New DocumentFormat.OpenXml.Presentation.SlideIdList())
                    presPart.Presentation.Save()
                End If

                ' 4) Build deck index
                Dim idx As DeckIndex = BuildDeckIndex(presPart)
                Dim currentAnchorKey As String = Nothing

                ' 5) Process actions
                For Each actElem In actions
                    If Not actElem.TryGetProperty("op", Nothing) _
                   OrElse actElem.GetProperty("op").GetString() <> "add_slide" Then
                        Continue For
                    End If

                    Try
                        ' 5.1 Anchor
                        Dim anchorKey = actElem.GetProperty("anchor") _
                                    .GetProperty("by") _
                                    .GetProperty("slideKey") _
                                    .GetString()
                        If anchorKey <> "lastInserted" Then
                            currentAnchorKey = anchorKey
                        End If

                        Dim layoutRelId As String = actElem.GetProperty("layoutRelId").GetString()
                        Dim anchorId As UInteger = 0UI
                        If currentAnchorKey IsNot Nothing AndAlso idx.SlideKeyById.ContainsKey(currentAnchorKey) Then
                            anchorId = idx.SlideKeyById(currentAnchorKey)
                        End If

                        ' 5.2 Clone slide
                        Dim newSp As DocumentFormat.OpenXml.Packaging.SlidePart =
                        CloneTemplateSlide(presPart, layoutRelId)
                        Dim newId As UInteger = InsertAfter(presPart, anchorId, newSp)

                        ' 5.3 Populate elements
                        For Each el In actElem.GetProperty("elements").EnumerateArray()
                            Dim t As String = el.GetProperty("type").GetString()
                            Select Case t
                                Case "title"
                                    SetTitle(newSp, el.GetProperty("text").GetString(), el)
                                Case "shape"
                                    AddShape(presPart, newSp, el)
                                Case "svg_icon"
                                    AddSvgIcon(presPart, newSp, el)
                                Case "text"
                                    If el.TryGetProperty("transform", Nothing) Then
                                        CreateFreestandingTextBox(presPart, newSp, el)
                                    Else
                                        SetText(newSp,
                                            el.GetProperty("placeholder").GetString(),
                                            el.GetProperty("text").GetString(),
                                            el)
                                    End If
                                Case "bullet_text"
                                    If el.TryGetProperty("transform", Nothing) Then
                                        ' freestanding list
                                        CreateFreestandingTextBox(presPart, newSp, el)
                                    Else
                                        ' keep using your original placeholder routine
                                        SetBullets(newSp, el)
                                    End If
                            End Select
                        Next

                        ' 5.4 Speaker notes
                        Dim notesEl As System.Text.Json.JsonElement
                        If actElem.TryGetProperty("notes", notesEl) AndAlso notesEl.ValueKind = JsonValueKind.String Then
                            ' Hier die gleiche Variable: newSlidePart
                            SetSpeakerNotes(newSp, notesEl.GetString())
                        End If

                        RemoveEmptyBodyPlaceholder(newSp)

                        ' Save intermediate
                        presPart.Presentation.Save()

                        ' 5.5 Rebuild index
                        idx = BuildDeckIndex(presPart)
                        currentAnchorKey = GetSlideKey(newSp, newId)

                    Catch ex As KeyNotFoundException
                        Debug.WriteLine("Could not implement instruction: " & ex.Message)
                        errorMessages.Add("Could not implement instruction: " & ex.Message)

                    Catch ex As Exception
                        Debug.WriteLine("Error creating slides: " & ex.Message)
                        errorMessages.Add("Error creating slides: " & ex.Message)
                    End Try
                Next

                ' 6) Fallback: ensure at least empty notes for every slide
                For Each sid In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                    Dim spPart As DocumentFormat.OpenXml.Packaging.SlidePart =
                        CType(presPart.GetPartById(sid.RelationshipId),
                              DocumentFormat.OpenXml.Packaging.SlidePart)

                    If spPart.NotesSlidePart Is Nothing Then
                        ' Leere Notes erzeugen:
                        SetSpeakerNotes(spPart, String.Empty)
                    End If
                Next

                ' 7) Final save
                presPart.Presentation.Save()
            End Using

            If errorMessages IsNot Nothing AndAlso errorMessages.Count > 0 Then
                Dim allErrors As String = String.Join(vbCrLf, errorMessages)
                ShowCustomMessageBox("Several errors occurred during applying the AI's instruction to your slidedeck (it may still have worked partially):" & vbCrLf & vbCrLf & allErrors)
                Return False
            End If

            Return True

        Catch oxEx As DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
            ShowCustomMessageBox("A PowerPoint file error occurred: " & oxEx.Message)
            Return False
        Catch ex As Exception
            ShowCustomMessageBox("An unexpected error occurred when amending your slidedeck: " & ex.Message)
            Return False
        End Try
    End Function



    Function ValidatePptx(path As String) As String

        Dim ErrorString As String = ""

        Using doc As PresentationDocument = PresentationDocument.Open(path, False)

            Dim validator As New OpenXmlValidator()
            Dim errors = validator.Validate(doc)

            If Not errors.Any() Then
                Debug.WriteLine("✔ Keine formalen OpenXML-Fehler gefunden.")
                Return ""
            End If

            For Each err As ValidationErrorInfo In errors
                Debug.WriteLine("----------")
                Debug.WriteLine($"Part : {err.Part.Uri}")
                Debug.WriteLine($"XPath: {err.Path.XPath}")
                Debug.WriteLine($"Info : {err.Description}")
                ErrorString = $"Part: {err.Part.Uri}; XPath: {err.Path.XPath}; Info: {err.Description}"
                ' nach dem ersten Fehler abbrechen – reicht zum Debuggen
                Exit For
            Next

        End Using

        Return ErrorString

    End Function


    ' Overload: nimmt PresentationPart statt Pfad
    Public Function BuildDeckIndex(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart
) As DeckIndex
        Dim idx As New DeckIndex With {
        .SlideKeyById = New Dictionary(Of String, UInteger)(),
        .IndexBySlideId = New Dictionary(Of UInteger, Integer)()
    }
        Dim i As Integer = 0
        For Each sid In presPart.Presentation.SlideIdList.Elements(
                            Of DocumentFormat.OpenXml.Presentation.SlideId)()
            idx.IndexBySlideId(sid.Id.Value) = i
            Dim sp = CType(presPart.GetPartById(sid.RelationshipId),
                       DocumentFormat.OpenXml.Packaging.SlidePart)
            Dim key = GetSlideKey(sp, sid.Id.Value)
            idx.SlideKeyById(key) = sid.Id.Value
            i += 1
        Next
        Return idx
    End Function


    Public Function BuildDeckIndex(pptxPath As String) As DeckIndex
        Using presDoc As DocumentFormat.OpenXml.Packaging.PresentationDocument =
              DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, False)
            Dim presPart = presDoc.PresentationPart
            Dim idx As New DeckIndex With {
              .SlideKeyById = New Dictionary(Of String, UInteger)(),
              .IndexBySlideId = New Dictionary(Of UInteger, Integer)()
            }
            Dim i As Integer = 0
            For Each sid As DocumentFormat.OpenXml.Presentation.SlideId _
                In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                idx.IndexBySlideId(sid.Id.Value) = i
                Dim sp = CType(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                Dim key = GetSlideKey(sp, sid.Id.Value)
                idx.SlideKeyById(key) = sid.Id.Value
                i += 1
            Next
            Return idx
        End Using
    End Function

    Public Class DeckIndex
        Public Property SlideKeyById As Dictionary(Of String, UInteger)
        Public Property IndexBySlideId As Dictionary(Of UInteger, Integer)
    End Class



    Private Function CloneTemplateSlide(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    layoutRelId As String
) As DocumentFormat.OpenXml.Packaging.SlidePart

        ' 1) Suche das passende SlideLayoutPart in den SlideMasterParts
        Dim targetLayout As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = Nothing
        For Each masterPart As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In masterPart.SlideLayoutParts
                If masterPart.GetIdOfPart(layoutPart) = layoutRelId Then
                    targetLayout = layoutPart
                    Exit For
                End If
            Next
            If targetLayout IsNot Nothing Then Exit For
        Next

        If targetLayout Is Nothing Then
            Throw New KeyNotFoundException(
            $"No SlideLayoutPart found for LayoutRelId '{layoutRelId}'."
        )
        End If

        ' 2) Neuen SlidePart erstellen
        Dim newSlidePart As DocumentFormat.OpenXml.Packaging.SlidePart =
        presPart.AddNewPart(Of DocumentFormat.OpenXml.Packaging.SlidePart)()

        ' 3) Neuen Slide aufbauen und CommonSlideData + ColorMapOverride klonen
        Dim newSlide As New DocumentFormat.OpenXml.Presentation.Slide()

        ' CommonSlideData (Platzhalter) klonen
        If targetLayout.SlideLayout.CommonSlideData IsNot Nothing Then
            newSlide.CommonSlideData = CType(
            targetLayout.SlideLayout.CommonSlideData.CloneNode(True),
            DocumentFormat.OpenXml.Presentation.CommonSlideData
        )
        End If

        ' Farbanpassung klonen (optional)
        If targetLayout.SlideLayout.ColorMapOverride IsNot Nothing Then
            newSlide.ColorMapOverride = CType(
            targetLayout.SlideLayout.ColorMapOverride.CloneNode(True),
            DocumentFormat.OpenXml.Presentation.ColorMapOverride
        )
        End If

        PurgeLayoutSampleText(newSlide)

        newSlidePart.Slide = newSlide

        CopyLayoutImagesToSlide(targetLayout, newSlidePart)  ' NEW safe version


        ' 4) Verknüpfe das Layout mit der neuen Slide
        newSlidePart.AddPart(targetLayout)

        ' 5) Speichern und zurückgeben
        newSlidePart.Slide.Save()
        Return newSlidePart
    End Function


    ''' <summary>
    ''' Copies every image part that <paramref name="layoutPart"/> uses
    ''' into <paramref name="slidePart"/> and rewrites the embed IDs
    ''' inside the cloned slide so they point to the copied images.
    ''' </summary>
    Private Sub CopyLayoutImagesToSlide(
        layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart,
        slidePart As DocumentFormat.OpenXml.Packaging.SlidePart)

        ' 1) oldRelId → newRelId
        Dim idMap As New System.Collections.Generic.Dictionary(Of String, String)(
        System.StringComparer.OrdinalIgnoreCase)

        ' 2) clone ONLY ImageParts — everything else stays in the layout
        For Each img In layoutPart.ImageParts
            Dim oldId As String = layoutPart.GetIdOfPart(img)

            ' create a fresh ImagePart in the slide
            Dim newImg = slidePart.AddImagePart(img.ContentType)
            ' copy the binary
            Using src = img.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read),
              dst = newImg.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write)
                src.CopyTo(dst)
            End Using

            idMap(oldId) = slidePart.GetIdOfPart(newImg)
        Next

        ' 3) rewrite every <a:blip embed="…">
        For Each blip In slidePart.Slide.
             Descendants(Of DocumentFormat.OpenXml.Drawing.Blip)()
            Dim oldId = blip.Embed?.Value
            If oldId IsNot Nothing AndAlso idMap.ContainsKey(oldId) Then
                blip.Embed.Value = idMap(oldId)
            End If
        Next
    End Sub



    Private Sub PurgeLayoutSampleText(sld As DocumentFormat.OpenXml.Presentation.Slide)

        ' only Title / CenteredTitle / Body placeholders get wiped
        For Each shp As DocumentFormat.OpenXml.Presentation.Shape _
        In sld.CommonSlideData.ShapeTree.
               Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim ph = shp.NonVisualShapeProperties?.
                 ApplicationNonVisualDrawingProperties?.
                 PlaceholderShape
            If ph Is Nothing Then Continue For

            Dim t As DocumentFormat.OpenXml.Presentation.PlaceholderValues? = Nothing
            If ph.Type IsNot Nothing Then t = ph.Type.Value

            If t Is Nothing _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body Then

                ' wipe existing content
                shp.TextBody?.Remove()

                ' insert minimal, valid skeleton
                shp.Append(New DocumentFormat.OpenXml.Presentation.TextBody(
                New DocumentFormat.OpenXml.Drawing.BodyProperties(),
                New DocumentFormat.OpenXml.Drawing.ListStyle(),
                New DocumentFormat.OpenXml.Drawing.Paragraph(
                    New DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties())))
            End If
        Next
    End Sub




    Private Function InsertAfter(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    anchorSlideId As UInteger,
    newSlidePart As DocumentFormat.OpenXml.Packaging.SlidePart
) As UInteger

        Dim slideList = presPart.Presentation.SlideIdList
        Dim relId = presPart.GetIdOfPart(newSlidePart)

        ' Existierende SlideId-Knoten
        Dim existing = slideList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
        Dim newId As UInteger

        If existing.Any() Then
            newId = existing.Max(Function(s) s.Id.Value) + 1UI
        Else
            newId = 256UI   ' Erstes Slide
        End If

        Dim newSlide = New DocumentFormat.OpenXml.Presentation.SlideId() With {
      .Id = newId,
      .RelationshipId = relId
    }

        ' Wenn anchorSlideId = 0, dann immer ans Ende anhängen
        If anchorSlideId = 0UI Then
            slideList.Append(newSlide)
        Else
            ' Ansonsten gezielt nach dem Anker einfügen
            Dim anchor = existing.FirstOrDefault(Function(s) s.Id.Value = anchorSlideId)
            If anchor Is Nothing Then
                slideList.Append(newSlide)
            Else
                anchor.InsertAfterSelf(newSlide)
            End If
        End If

        Return newId
    End Function

    Private Sub SetText(
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    placeholderName As String,
    text As String,
    el As System.Text.Json.JsonElement
)
        ' 1) Alle Shapes auf der Folie ermitteln
        Dim allShapes = sp.Slide.CommonSlideData.ShapeTree.
                    Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
        Dim targetShape As DocumentFormat.OpenXml.Presentation.Shape = Nothing

        ' 1a) Echte Body-Placeholder-Box finden
        For Each shp In allShapes
            Dim ph = shp.NonVisualShapeProperties?.
                 ApplicationNonVisualDrawingProperties?.
                 PlaceholderShape
            If ph IsNot Nothing AndAlso ph.Type IsNot Nothing AndAlso
           ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body Then
                targetShape = shp
                Exit For
            End If
        Next

        ' 1b) Fallback: JSON-placeholder im Shape-Namen
        If targetShape Is Nothing AndAlso Not String.IsNullOrEmpty(placeholderName) Then
            For Each shp In allShapes
                Dim nv = shp.NonVisualShapeProperties?.
                     NonVisualDrawingProperties
                Dim nm = If(nv?.Name?.Value, "")
                If nm.IndexOf(placeholderName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    targetShape = shp
                    Exit For
                End If
            Next
        End If

        ' 1c) Fallback: erstes Nicht-Title-Shape
        If targetShape Is Nothing Then
            For Each shp In allShapes
                Dim ph = shp.NonVisualShapeProperties?.
                     ApplicationNonVisualDrawingProperties?.
                     PlaceholderShape
                Dim typ = If(ph?.Type IsNot Nothing, ph.Type.Value, Nothing)
                If typ <> DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title AndAlso
               typ <> DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle Then
                    targetShape = shp
                    Exit For
                End If
            Next
        End If

        If targetShape Is Nothing Then Return

        ' 2) Neuer TextBody (ohne ListStyle, damit keine Bullets erscheinen)
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody()
        tb.Append(New DocumentFormat.OpenXml.Drawing.BodyProperties())
        ' kein ListStyle hinzufügen

        ' 3) RunProperties aus el("style")
        Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()
        Dim styleEl As System.Text.Json.JsonElement
        If el.TryGetProperty("style", styleEl) Then
            Dim tmp As System.Text.Json.JsonElement
            If styleEl.TryGetProperty("fontFamily", tmp) Then
                rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = tmp.GetString()})
            End If
            If styleEl.TryGetProperty("fontSize", tmp) Then
                rp.FontSize = CUInt(tmp.GetInt32() * 100)
            End If
            If styleEl.TryGetProperty("bold", tmp) AndAlso tmp.GetBoolean() Then rp.Bold = True
            If styleEl.TryGetProperty("italic", tmp) AndAlso tmp.GetBoolean() Then rp.Italic = True
            If styleEl.TryGetProperty("color", tmp) Then
                Dim hex = tmp.GetString().TrimStart("#"c)
                rp.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(
                New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = hex}
            ))
            End If
        End If

        ' 4) ParagraphProperties mit NoBullet, um Aufzählungszeichen zu unterdrücken
        Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {
        .Indent = 0,       ' hanging indent = 0
        .LeftMargin = 0    ' left margin = 0
                }

        pPr.Append(New DocumentFormat.OpenXml.Drawing.NoBullet())


        ' 5) Run und Paragraph erstellen
        Dim runElem = New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text))
        Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph()
        para.Append(pPr)
        para.Append(runElem)
        tb.Append(para)

        ' 6) TextBody dem Shape zuweisen und speichern
        targetShape.TextBody = tb
        sp.Slide.Save()
    End Sub




    Private Sub SetTitle(
        sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        text As System.String,
        el As System.Text.Json.JsonElement)

        Dim shapes = sp.Slide.CommonSlideData.ShapeTree.
                 Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

        Dim titleShape As DocumentFormat.OpenXml.Presentation.Shape = Nothing

        ' --- 1) explicit Title / CenteredTitle --------------------------------
        titleShape = shapes.FirstOrDefault(Function(shp)
                                               Dim ph = shp.NonVisualShapeProperties? _
                 .ApplicationNonVisualDrawingProperties? _
                 .PlaceholderShape
                                               Return ph IsNot Nothing AndAlso ph.Type IsNot Nothing AndAlso
               (ph.Type.Value =
                DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title OrElse
                ph.Type.Value =
                DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle)
                                           End Function)

        ' --- 2) implicit: placeholder with ph:index = 0 -----------------------
        If titleShape Is Nothing Then
            titleShape = shapes.FirstOrDefault(Function(shp)
                                                   Dim ph = shp.NonVisualShapeProperties? _
                     .ApplicationNonVisualDrawingProperties? _
                     .PlaceholderShape
                                                   Return ph IsNot Nothing AndAlso ph.Index IsNot Nothing AndAlso
                   ph.Index.Value = 0UI
                                               End Function)
        End If

        ' --- 3) last fallback: shape name contains "title" --------------------
        If titleShape Is Nothing Then
            titleShape = shapes.FirstOrDefault(Function(shp)
                                                   Dim nm = shp.NonVisualShapeProperties? _
                     .NonVisualDrawingProperties?.Name?.Value
                                                   Return Not System.String.IsNullOrWhiteSpace(nm) AndAlso
                   nm.IndexOf("title",
                              System.StringComparison.OrdinalIgnoreCase) >= 0
                                               End Function)
        End If

        If titleShape Is Nothing Then Return   ' nothing suitable found

        ' --- 4) build a fresh TextBody ----------------------------------------
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody(
        New DocumentFormat.OpenXml.Drawing.BodyProperties(),
        New DocumentFormat.OpenXml.Drawing.ListStyle())

        tb.Append(BuildParagraph(text, el))   ' your existing helper

        titleShape.TextBody = tb
        sp.Slide.Save()
    End Sub



    Private Sub OldSetTitle(
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    text As String,
    el As System.Text.Json.JsonElement
)
        ' 1) Title-Shape finden
        Dim titleShape As DocumentFormat.OpenXml.Presentation.Shape = Nothing
        For Each shp As DocumentFormat.OpenXml.Presentation.Shape _
        In sp.Slide.CommonSlideData.ShapeTree.
           Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

            ' a) PlaceholderShape sicher auslesen
            Dim ph = shp.NonVisualShapeProperties? _
                     .ApplicationNonVisualDrawingProperties? _
                     .PlaceholderShape

            ' b) Erst prüfen, ob ph.Type nicht Nothing ist
            If ph IsNot Nothing AndAlso
           ph.Type IsNot Nothing AndAlso
           (ph.Type.Value = DocumentFormat.OpenXml.Presentation.
                            PlaceholderValues.Title _
            OrElse ph.Type.Value = DocumentFormat.OpenXml.Presentation.
                            PlaceholderValues.CenteredTitle) Then

                titleShape = shp
                Exit For
            End If
        Next

        ' Wenn kein Title-Placeholder gefunden, abbrechen
        If titleShape Is Nothing Then
            Return
        End If

        ' 2) Neuen TextBody bauen und an den Shape hängen
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody()
        tb.Append(New DocumentFormat.OpenXml.Drawing.BodyProperties())
        tb.Append(New DocumentFormat.OpenXml.Drawing.ListStyle())
        tb.Append(BuildParagraph(text, el))

        titleShape.TextBody = tb
        sp.Slide.Save()
    End Sub



    Private Sub SetBullets(
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement
)
        ' 1) Optionalen Placeholder-Namen aus JSON lesen
        Dim placeholderName As String = Nothing
        Dim tmpEl As System.Text.Json.JsonElement
        If el.TryGetProperty("placeholder", tmpEl) Then
            placeholderName = tmpEl.GetString()
        End If

        ' 2) Alle Shapes auf der Folie durchsuchen
        Dim allShapes = sp.Slide.CommonSlideData.ShapeTree.
                    Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
        Dim bodyShape As DocumentFormat.OpenXml.Presentation.Shape = Nothing

        ' 2a) Echte Body-Placeholder-Box finden
        For Each shp In allShapes
            Dim ph = shp.NonVisualShapeProperties? _
                 .ApplicationNonVisualDrawingProperties? _
                 .PlaceholderShape
            If ph IsNot Nothing AndAlso ph.Type IsNot Nothing AndAlso
           ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body Then
                bodyShape = shp
                Exit For
            End If
        Next

        ' 2b) Fallback: JSON-placeholder im Shape-Namen
        If bodyShape Is Nothing AndAlso Not String.IsNullOrEmpty(placeholderName) Then
            For Each shp In allShapes
                Dim nvProps = shp.NonVisualShapeProperties? _
                          .NonVisualDrawingProperties
                Dim shpName As String = If(nvProps?.Name?.Value, "")
                If shpName.IndexOf(placeholderName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    bodyShape = shp
                    Exit For
                End If
            Next
        End If

        ' 2c) Fallback: erstes Nicht-Title-Shape
        If bodyShape Is Nothing Then
            For Each shp In allShapes
                Dim ph = shp.NonVisualShapeProperties? _
                     .ApplicationNonVisualDrawingProperties? _
                     .PlaceholderShape
                Dim typ = If(ph?.Type IsNot Nothing, ph.Type.Value, Nothing)
                If typ <> DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title AndAlso
               typ <> DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle Then
                    bodyShape = shp
                    Exit For
                End If
            Next
        End If

        ' Abbruch, wenn kein Body-Shape gefunden wurde
        If bodyShape Is Nothing Then
            Return
        End If

        ' 3) Neues TextBody mit ListStyle erzeugen
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody()
        tb.Append(New DocumentFormat.OpenXml.Drawing.BodyProperties())
        tb.Append(New DocumentFormat.OpenXml.Drawing.ListStyle())

        ' 4) Bullets aus JSON lesen und als verschachtelte Paragraphen anfügen
        For Each bElem As System.Text.Json.JsonElement In el.GetProperty("bullets").EnumerateArray()
            ' 4a) Text und Level ermitteln
            Dim text As String
            Dim level As Integer = 0
            If bElem.ValueKind = System.Text.Json.JsonValueKind.Object Then
                If bElem.TryGetProperty("text", tmpEl) Then
                    text = tmpEl.GetString()
                Else
                    Continue For
                End If
                If bElem.TryGetProperty("level", tmpEl) Then
                    level = tmpEl.GetInt32()
                End If
            Else
                text = bElem.GetString()
            End If

            ' 4b) RunProperties aus el.style erzeugen
            Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()
            Dim styleEl As System.Text.Json.JsonElement
            If el.TryGetProperty("style", styleEl) Then
                Dim tmp As System.Text.Json.JsonElement
                If styleEl.TryGetProperty("fontFamily", tmp) Then
                    rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = tmp.GetString()})
                End If
                If styleEl.TryGetProperty("fontSize", tmp) Then
                    rp.FontSize = CUInt(tmp.GetInt32() * 100)
                End If
                If styleEl.TryGetProperty("bold", tmp) AndAlso tmp.GetBoolean() Then rp.Bold = True
                If styleEl.TryGetProperty("italic", tmp) AndAlso tmp.GetBoolean() Then rp.Italic = True


            End If

            ' 4c) ParagraphProperties mit Level setzen
            Dim actualLevel = System.Math.Max(0, System.Math.Min(8, level))
            Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {
                    .Level = CByte(actualLevel)
                }
            ' 4d) Run und Paragraph bauen
            Dim runElem = New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text))
            Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph()
            para.Append(pPr)
            para.Append(runElem)

            tb.Append(para)
        Next

        ' 5) TextBody dem Shape zuweisen und speichern
        bodyShape.TextBody = tb
        sp.Slide.Save()
    End Sub


    ''' <summary>
    ''' Baut einen einzelnen Drawing.Paragraph mit Text und RunProperties.
    ''' </summary>
    Private Function BuildParagraph(
      text As String,
      el As JsonElement
    ) As DocumentFormat.OpenXml.Drawing.Paragraph

        ' 1) RunProperties und Style aus JSON
        Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()
        Dim styleEl As JsonElement
        If el.TryGetProperty("style", styleEl) Then
            Dim tmp As JsonElement
            If styleEl.TryGetProperty("fontFamily", tmp) Then
                rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = tmp.GetString()})
            End If
            If styleEl.TryGetProperty("fontSize", tmp) Then
                rp.FontSize = CUInt(tmp.GetInt32() * 100)
            End If
            If styleEl.TryGetProperty("bold", tmp) AndAlso tmp.GetBoolean() Then rp.Bold = True
            If styleEl.TryGetProperty("italic", tmp) AndAlso tmp.GetBoolean() Then rp.Italic = True
            If styleEl.TryGetProperty("color", tmp) Then
                Dim hex = tmp.GetString().TrimStart("#"c)

            End If
        End If

        ' 2) Run + Paragraph erzeugen
        Dim runElem = New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text))
        Dim para = New DocumentFormat.OpenXml.Drawing.Paragraph()
        para.Append(runElem)
        Return para
    End Function


    Private Function BuildParagraph(text As String, el As JsonElement, Optional pPr As DocumentFormat.OpenXml.Drawing.ParagraphProperties = Nothing) As DocumentFormat.OpenXml.Drawing.Paragraph
        Dim rp = New DocumentFormat.OpenXml.Drawing.RunProperties()
        Dim styleEl As JsonElement
        If el.TryGetProperty("style", styleEl) Then
            Dim tmp As JsonElement
            If styleEl.TryGetProperty("fontFamily", tmp) Then rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = tmp.GetString()})
            If styleEl.TryGetProperty("fontSize", tmp) Then rp.FontSize = CInt(tmp.GetInt32() * 100)
            If styleEl.TryGetProperty("bold", tmp) AndAlso tmp.GetBoolean() Then rp.Bold = True
            If styleEl.TryGetProperty("italic", tmp) AndAlso tmp.GetBoolean() Then rp.Italic = True
            If styleEl.TryGetProperty("color", tmp) Then
                rp.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = tmp.GetString().TrimStart("#"c)}))
            End If
        End If

        Dim run = New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text))
        Dim para = New DocumentFormat.OpenXml.Drawing.Paragraph()
        If pPr IsNot Nothing Then
            para.Append(pPr.CloneNode(True))
        End If
        para.Append(run)
        Return para
    End Function

    Private Function BuildRun(
      text As String,
      el As JsonElement
    ) As DocumentFormat.OpenXml.Drawing.Run

        Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()

        ' KORREKT: Füge LatinFont, SolidFill als Kind-Elemente hinzu
        Dim styleEl As JsonElement
        If el.TryGetProperty("style", styleEl) Then
            Dim tmp As JsonElement
            If styleEl.TryGetProperty("fontFamily", tmp) Then
                rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = tmp.GetString()})
            End If
            If styleEl.TryGetProperty("fontSize", tmp) Then
                rp.FontSize = CUInt(tmp.GetInt32() * 100)
            End If
            If styleEl.TryGetProperty("bold", tmp) AndAlso tmp.GetBoolean() Then rp.Bold = True
            If styleEl.TryGetProperty("italic", tmp) AndAlso tmp.GetBoolean() Then rp.Italic = True

        End If

        Return New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text))
    End Function


    Private Function GetSlideKey(
        sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        slideId As UInteger
      ) As String
        Dim title = GetSlideTitle(sp) ' <- Dein vorhandener Helper!
        If String.IsNullOrWhiteSpace(title) Then
            Return $"SID-{slideId}"
        Else
            Return $"{SanitizeKey(title)}-{slideId}"
        End If
    End Function

    Private Sub SetSpeakerNotes(
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    notesText As String)

        Dim notesPart As DocumentFormat.OpenXml.Packaging.NotesSlidePart = sp.NotesSlidePart
        If notesPart Is Nothing Then
            notesPart = sp.AddNewPart(Of DocumentFormat.OpenXml.Packaging.NotesSlidePart)()
            notesPart.NotesSlide = New DocumentFormat.OpenXml.Presentation.NotesSlide(
            New DocumentFormat.OpenXml.Presentation.CommonSlideData(
                New DocumentFormat.OpenXml.Presentation.ShapeTree(
                    New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = 1UI, .Name = ""},
                        New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
                    New DocumentFormat.OpenXml.Presentation.GroupShapeProperties())),
            New DocumentFormat.OpenXml.Presentation.ColorMapOverride(
                New DocumentFormat.OpenXml.Drawing.MasterColorMapping()))
        End If

        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree =
        notesPart.NotesSlide.CommonSlideData.ShapeTree

        ' ----- nur Shapes/Pics entfernen -----
        For Each n In tree.ChildElements.OfType(Of DocumentFormat.OpenXml.OpenXmlElement)().ToList()
            If TypeOf n Is DocumentFormat.OpenXml.Presentation.Shape _
           OrElse TypeOf n Is DocumentFormat.OpenXml.Presentation.Picture _
           OrElse TypeOf n Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                n.Remove()
            End If
        Next

        ' ----- neues Body-Shape -----
        Dim nvSpPr As New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
        New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = 2UI, .Name = "NotesBody"},
        New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(
            New DocumentFormat.OpenXml.Drawing.ShapeLocks() With {.NoGrouping = True}),
        New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties(
            New DocumentFormat.OpenXml.Presentation.PlaceholderShape() With {
                .Type = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                .Index = 1UI}))
        Dim shapePr As New DocumentFormat.OpenXml.Presentation.ShapeProperties()
        Dim noteShape As New DocumentFormat.OpenXml.Presentation.Shape(nvSpPr, shapePr)

        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody(
        New DocumentFormat.OpenXml.Drawing.BodyProperties(),
        New DocumentFormat.OpenXml.Drawing.ListStyle())
        Dim run As New DocumentFormat.OpenXml.Drawing.Run(
        New DocumentFormat.OpenXml.Drawing.RunProperties(),
        New DocumentFormat.OpenXml.Drawing.Text(notesText))
        Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph(run) With {
        .ParagraphProperties = New DocumentFormat.OpenXml.Drawing.ParagraphProperties()}
        tb.Append(para)
        noteShape.Append(tb)

        ' nach Header einsetzen
        If tree.ChildElements.Count >= 2 Then
            tree.InsertAt(noteShape, 2)
        Else
            tree.Append(noteShape)
        End If

        notesPart.NotesSlide.Save()
    End Sub




    ''' <summary>
    ''' Creates an OpenXML Fill element from a JSON definition.
    ''' CORRECTED: Returns OpenXmlElement to allow for NoFill type.
    ''' </summary>
    Private Function CreateFill(fillJson As JsonElement) As DocumentFormat.OpenXml.OpenXmlElement
        Dim fillType As String = ""
        If fillJson.TryGetProperty("type", Nothing) Then fillType = fillJson.GetProperty("type").GetString()

        Select Case fillType.ToLower()
            Case "solid"
                If fillJson.TryGetProperty("color", Nothing) Then
                    Dim colorHex = fillJson.GetProperty("color").GetString().TrimStart("#"c)
                    Return New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex With {.Val = colorHex})
                End If
        End Select
        Return New DocumentFormat.OpenXml.Drawing.NoFill() ' Fallback
    End Function

    ''' <summary>
    ''' [CORRECTED] Creates an OpenXML Outline element from a JSON definition.
    ''' Fixes the file corruption bug by safely parsing numbers for any computer locale.
    ''' </summary>
    Private Function CreateOutline(outlineJson As JsonElement) As DocumentFormat.OpenXml.Drawing.Outline
        Dim outline As New DocumentFormat.OpenXml.Drawing.Outline()

        Dim widthJson As JsonElement
        If outlineJson.TryGetProperty("width", widthJson) Then
            ' [FIX] This safely parses numbers like "1" or "1.5" regardless of system language.
            Dim widthValue As Double
            If Double.TryParse(widthJson.GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, widthValue) Then
                outline.Width = CInt(widthValue * 12700) ' 1 point = 12700 EMUs
            End If
        End If

        Dim colorJson As JsonElement
        If outlineJson.TryGetProperty("color", colorJson) Then
            outline.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex With {.Val = colorJson.GetString().TrimStart("#"c)}))
        End If

        Dim dashJson As JsonElement
        If outlineJson.TryGetProperty("dashType", dashJson) Then
            outline.Append(New DocumentFormat.OpenXml.Drawing.PresetDash With {.Val = JsonDashNameToEnumValue(dashJson.GetString())})
        End If

        Return outline
    End Function



    ''' <summary>
    ''' [NEW] Converts relative percentage-based coordinates from JSON into absolute EMU coordinates.
    ''' </summary>
    ''' <param name="presPart">The presentation part, to get the master slide dimensions.</param>
    ''' <param name="transformJson">The JSON "transform" object.</param>
    ''' <returns>A fully calculated Transform2D object with absolute EMUs.</returns>
    Private Function ConvertRelativeToAbsoluteTransform(presPart As DocumentFormat.OpenXml.Packaging.PresentationPart, transformJson As System.Text.Json.JsonElement) As DocumentFormat.OpenXml.Drawing.Transform2D
        ' Get the master slide dimensions in EMUs
        Dim slideWidthEmu = presPart.Presentation.SlideSize.Cx.Value
        Dim slideHeightEmu = presPart.Presentation.SlideSize.Cy.Value

        ' Safely parse the relative percentage values from JSON
        Dim relX, relY, relW, relH As Double
        Double.TryParse(transformJson.GetProperty("x").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relX)
        Double.TryParse(transformJson.GetProperty("y").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relY)
        Double.TryParse(transformJson.GetProperty("width").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relW)
        Double.TryParse(transformJson.GetProperty("height").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relH)

        ' Calculate the absolute EMU values
        Dim absX = CLng(slideWidthEmu * relX)
        Dim absY = CLng(slideHeightEmu * relY)
        Dim absCx = CLng(slideWidthEmu * relW)
        Dim absCy = CLng(slideHeightEmu * relH)

        Return New DocumentFormat.OpenXml.Drawing.Transform2D(
        New DocumentFormat.OpenXml.Drawing.Offset With {.X = absX, .Y = absY},
        New DocumentFormat.OpenXml.Drawing.Extents With {.Cx = absCx, .Cy = absCy}
    )
    End Function


    Private Function JsonShapeNameToEnumValue(jsonName As String) As DocumentFormat.OpenXml.Drawing.ShapeTypeValues
        Select Case jsonName.ToLower()
            Case "rectangle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
            Case "oval", "ellipse", "circle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse
            Case "line" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Line
            Case "rightarrow" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RightArrow
            Case "leftarrow" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.LeftArrow
            Case "triangle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Triangle ' Corrected from IsoscelesTriangle
            Case "roundedrectangle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle
            Case "flowchartprocess" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.FlowChartProcess
            Case "flowchartdecision" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.FlowChartDecision
            Case "flowchartterminator" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.FlowChartTerminator
            Case "chevron" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Chevron
            Case "pentagon" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Pentagon
            Case "hexagon" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Hexagon
            Case "plus" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Plus
            Case "blockarc" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.BlockArc
            Case Else : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle ' Fallback
        End Select
    End Function


    Private Function JsonDashNameToEnumValue(jsonName As String) As DocumentFormat.OpenXml.Drawing.PresetLineDashValues
        Select Case jsonName.ToLower()
            Case "solid"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid
            Case "dot", "dotted"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot
            Case "dash", "dashed"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dash
            Case "longdash"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.LargeDash
            Case "dashdot"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.DashDot
            Case "longdashdot"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.LargeDashDot
            Case Else
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid ' Fallback
        End Select
    End Function






    Private Function BuildStyledParagraph(text As String, level As Integer, el As System.Text.Json.JsonElement, isBulleted As Boolean) As DocumentFormat.OpenXml.Drawing.Paragraph
        Dim pPr = New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {.Level = level}
        Dim rp = New DocumentFormat.OpenXml.Drawing.RunProperties()

        Dim styleEl As JsonElement
        If el.TryGetProperty("style", styleEl) Then
            If styleEl.TryGetProperty("font", Nothing) Then rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = styleEl.GetProperty("font").GetString()})
            If styleEl.TryGetProperty("size", Nothing) Then rp.FontSize = CInt(styleEl.GetProperty("size").GetInt32() * 100)
            If styleEl.TryGetProperty("bold", Nothing) AndAlso styleEl.GetProperty("bold").GetBoolean() Then rp.Bold = True
            'If styleEl.TryGetProperty("color", Nothing) Then rp.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = styleEl.GetProperty("color").GetString().TrimStart("#"c)}))
            If styleEl.TryGetProperty("align", Nothing) Then
                Select Case styleEl.GetProperty("align").GetString().ToLower()
                    Case "center" : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center
                    Case "right" : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right
                    Case Else : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left
                End Select
            End If
        End If

        If Not isBulleted Then
            pPr.Append(New DocumentFormat.OpenXml.Drawing.NoBullet())
        End If

        Return New DocumentFormat.OpenXml.Drawing.Paragraph(pPr, New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text)))
    End Function


    Private Sub CreateFreestandingTextBox(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement)

        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree = sp.Slide.CommonSlideData.ShapeTree

        ' 1) Find next available shape ID
        Dim maxId As UInteger = 0
        For Each nonVisPr As DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties _
        In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)()
            If nonVisPr.Id.Value > maxId Then maxId = nonVisPr.Id.Value
        Next
        Dim newId As UInteger = maxId + 1

        ' 2) Locate the transform JSON
        Dim tf As System.Text.Json.JsonElement
        If Not el.TryGetProperty("transform", tf) Then
            If el.TryGetProperty("style", tf) AndAlso tf.TryGetProperty("transform", tf) Then
                ' nested under style
            Else
                Return
            End If
        End If

        ' 3) Compute absolute EMU coordinates
        Dim rawX As Double
        Double.TryParse(tf.GetProperty("x").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, rawX)

        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D
        If rawX > 1 Then
            ' already EMU
            Dim ofs As New DocumentFormat.OpenXml.Drawing.Offset() With {
            .X = CLng(tf.GetProperty("x").GetInt64()),
            .Y = CLng(tf.GetProperty("y").GetInt64())
        }
            Dim ext As New DocumentFormat.OpenXml.Drawing.Extents() With {
            .Cx = CLng(tf.GetProperty("width").GetInt64()),
            .Cy = CLng(tf.GetProperty("height").GetInt64())
        }
            xfrm = New DocumentFormat.OpenXml.Drawing.Transform2D(ofs, ext)
        Else
            ' percent → EMU
            xfrm = ConvertRelativeToAbsoluteTransform(presPart, tf)
        End If

        ' 4) Build the textbox shape
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties() With {.Transform2D = xfrm}
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(
        New DocumentFormat.OpenXml.Drawing.AdjustValueList()
    ) With {.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle})
        spPr.Append(New DocumentFormat.OpenXml.Drawing.NoFill())

        Dim nvDr As New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties() With {.TextBox = True}
        nvDr.AppendChild(New DocumentFormat.OpenXml.Drawing.ShapeLocks())

        Dim nvProps As New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
        New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {
            .Id = newId,
            .Name = "TextBox " & newId
        },
        nvDr,
        New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
    )

        Dim shp As New DocumentFormat.OpenXml.Presentation.Shape(nvProps, spPr)

        ' 5) Populate text or bullets
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody(
        New DocumentFormat.OpenXml.Drawing.BodyProperties(),
        New DocumentFormat.OpenXml.Drawing.ListStyle()
    )

        Select Case el.GetProperty("type").GetString()
            Case "text"
                tb.Append(BuildParagraph(el.GetProperty("text").GetString(), el))
            Case "bullet_text"
                For Each b In el.GetProperty("bullets").EnumerateArray()
                    Dim txt As String = If(
                        b.ValueKind = JsonValueKind.Object,
                        b.GetProperty("text").GetString(),
                        b.GetString())

                    Dim lvl As Integer = 0
                    Dim tmp As System.Text.Json.JsonElement
                    If b.ValueKind = JsonValueKind.Object AndAlso b.TryGetProperty("level", tmp) Then
                        lvl = tmp.GetInt32()
                    End If

                    ' Classic 0.5-cm hanging indent: bullet at 0, text at 0.5 cm
                    Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {
                        .Level = CByte(System.Math.Max(0, System.Math.Min(8, lvl))),
                        .LeftMargin = 457200,     ' 0.5 cm
                        .Indent = -457200,    ' hanging indent equals left margin
                        .Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left}

                    pPr.Append(New DocumentFormat.OpenXml.Drawing.BulletFont() With {.Typeface = "Arial"})
                    pPr.Append(New DocumentFormat.OpenXml.Drawing.CharacterBullet() With {.Char = "•"c})

                    ' Run: *no* extra tab needed – indent handles spacing
                    Dim run = New DocumentFormat.OpenXml.Drawing.Run(
                  New DocumentFormat.OpenXml.Drawing.RunProperties(),
                  New DocumentFormat.OpenXml.Drawing.Text(txt))

                    tb.Append(New DocumentFormat.OpenXml.Drawing.Paragraph(pPr, run))
                Next


        End Select

        shp.Append(tb)
        tree.Append(shp)
        sp.Slide.Save()
    End Sub


    Private Sub AddShape(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement)

        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree = sp.Slide.CommonSlideData.ShapeTree

        ' 1) ID ermitteln
        Dim maxId As UInteger = 0UI
        For Each nvPr As DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties _
        In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)()
            If nvPr.Id.Value > maxId Then maxId = nvPr.Id.Value
        Next
        Dim newId As UInteger = maxId + 1UI

        ' 2) Transform
        Dim transformJson = el.GetProperty("transform")

        ' Raw-Wert prüfen: ≤1 = Prozent, >1 = bereits EMUs
        Dim rawX As Double
        If Not Double.TryParse(transformJson.GetProperty("x").GetRawText(),
                       Globalization.NumberStyles.Any,
                       Globalization.CultureInfo.InvariantCulture,
                       rawX) Then
            rawX = 0.0
        End If

        Dim absoluteTransform As DocumentFormat.OpenXml.Drawing.Transform2D

        If rawX <= 1.0 Then
            ' Prozentwerte → in EMU umrechnen
            absoluteTransform = ConvertRelativeToAbsoluteTransform(presPart, transformJson)
        Else
            ' Direkte EMU-Werte übernehmen
            Dim ofs As New DocumentFormat.OpenXml.Drawing.Offset() With {
        .X = CLng(transformJson.GetProperty("x").GetInt64()),
        .Y = CLng(transformJson.GetProperty("y").GetInt64())
    }
            Dim ext As New DocumentFormat.OpenXml.Drawing.Extents() With {
        .Cx = CLng(transformJson.GetProperty("width").GetInt64()),
        .Cy = CLng(transformJson.GetProperty("height").GetInt64())
    }
            absoluteTransform = New DocumentFormat.OpenXml.Drawing.Transform2D(ofs, ext)
        End If

        ' 3) ShapeProperties
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties() With {.Transform2D = absoluteTransform}
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(
        New DocumentFormat.OpenXml.Drawing.AdjustValueList()
    ) With {.Preset = JsonShapeNameToEnumValue(el.GetProperty("shapeType").GetString())})
        If el.TryGetProperty("fill", Nothing) Then spPr.Append(CreateFill(el.GetProperty("fill")))
        If el.TryGetProperty("outline", Nothing) Then spPr.Append(CreateOutline(el.GetProperty("outline")))

        ' 4) nvSpPr (TextBox nur setzen, wenn Text folgt)
        Dim nvSpDr = New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties()
        If el.TryGetProperty("text", Nothing) Then
            nvSpDr.TextBox = True
            nvSpDr.AppendChild(New DocumentFormat.OpenXml.Drawing.ShapeLocks())
        End If
        Dim nvSpPr = New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
    New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = newId, .Name = $"Shape {newId}"},
    nvSpDr,
    New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
)
        ' 5) Neues Shape
        Dim shp As New DocumentFormat.OpenXml.Presentation.Shape(nvSpPr, spPr)

        ' 6) Optional Text
        If el.TryGetProperty("text", Nothing) Then
            Dim tb = New DocumentFormat.OpenXml.Presentation.TextBody(
            New DocumentFormat.OpenXml.Drawing.BodyProperties(),
            New DocumentFormat.OpenXml.Drawing.ListStyle()
        )
            tb.Append(BuildStyledParagraph(el.GetProperty("text").GetString(), 0, el, False))
            shp.Append(tb)
        End If

        ' 7) Einfügen & Speichern
        tree.Append(shp)
        sp.Slide.Save()
    End Sub

    ''' <summary>
    ''' Inserts an SVG icon from the JSON at the given location.
    ''' Uses a standard <p:pic> with <a:blip>; this is the same recipe
    ''' PowerPoint 2019+ generates and shows on Office 2016 (Oct-2018) too.
    ''' </summary>
    Private Sub AddSvgIcon(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement)

        Dim tree = sp.Slide.CommonSlideData.ShapeTree

        ' 1) unique ID on slide
        Dim newId As UInteger =
    tree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)().
        Select(Function(nv) nv.Id.Value).DefaultIfEmpty(0).Max() + 1UI

        ' 2) build Transform2D (percent → EMU if needed)
        Dim tf = el.GetProperty("transform")
        Dim rawX As Double
        Double.TryParse(tf.GetProperty("x").GetRawText(),
        Globalization.NumberStyles.Any,
        Globalization.CultureInfo.InvariantCulture,
        rawX)

        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D
        If rawX > 1 Then
            xfrm = New DocumentFormat.OpenXml.Drawing.Transform2D(
            New DocumentFormat.OpenXml.Drawing.Offset With {
                .X = CLng(tf.GetProperty("x").GetInt64()),
                .Y = CLng(tf.GetProperty("y").GetInt64())},
            New DocumentFormat.OpenXml.Drawing.Extents With {
                .Cx = CLng(tf.GetProperty("width").GetInt64()),
                .Cy = CLng(tf.GetProperty("height").GetInt64())})
        Else
            ' Assuming ConvertRelativeToAbsoluteTransform exists and returns a Transform2D
            ' xfrm = ConvertRelativeToAbsoluteTransform(presPart, tf) 
            xfrm = ConvertRelativeToAbsoluteTransform(presPart, tf)
        End If

        ' 3) embed SVG file
        Dim svgPart = sp.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Svg)
        Using ms As New IO.MemoryStream(
        System.Text.Encoding.UTF8.GetBytes(el.GetProperty("svg").GetString()))
            svgPart.FeedData(ms)
        End Using
        Dim relId As String = sp.GetIdOfPart(svgPart)

        ' 4) build <p:pic>
        Dim nvPic As New DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties(
    New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {
        .Id = newId, .Name = "Icon " & newId},
    New DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties(
        New DocumentFormat.OpenXml.Drawing.PictureLocks() With {.NoChangeAspect = True}),
    New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties())

        Dim blipFill As New DocumentFormat.OpenXml.Presentation.BlipFill(
    New DocumentFormat.OpenXml.Drawing.Blip() With {
        .Embed = relId,
        .CompressionState =
            DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print},
    New DocumentFormat.OpenXml.Drawing.Stretch(
        New DocumentFormat.OpenXml.Drawing.FillRectangle()))

        ' Define the rectangle geometry
        Dim prstGeom As New DocumentFormat.OpenXml.Drawing.PresetGeometry(
        New DocumentFormat.OpenXml.Drawing.AdjustValueList()
    ) With {.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle}

        ' Create the ShapeProperties object
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties()

        ' Append the transform and geometry as child elements
        spPr.Append(xfrm)
        spPr.Append(prstGeom)

        ' Create the final picture by combining all the parts
        Dim pic As New DocumentFormat.OpenXml.Presentation.Picture(nvPic, blipFill, spPr)

        ' 5) append & save
        tree.Append(pic)
        sp.Slide.Save()
    End Sub



    Private Sub RemoveEmptyBodyPlaceholder(sp As DocumentFormat.OpenXml.Packaging.SlidePart)
        Dim shpToRemove As DocumentFormat.OpenXml.Presentation.Shape = Nothing

        For Each shp In sp.Slide.CommonSlideData.ShapeTree.
                         Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim ph = shp.NonVisualShapeProperties?.
                        ApplicationNonVisualDrawingProperties?.
                        PlaceholderShape
            If ph IsNot Nothing AndAlso ph.Type IsNot Nothing AndAlso
               ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body Then

                ' empty = only one paragraph with no text or whitespace
                Dim empty As Boolean =
                    (shp.TextBody Is Nothing) OrElse
                    Not shp.TextBody.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)().
                        Any(Function(t) Not String.IsNullOrWhiteSpace(t.Text))

                If empty Then shpToRemove = shp
                Exit For
            End If
        Next

        If shpToRemove IsNot Nothing Then
            shpToRemove.Remove()
            sp.Slide.Save()
        End If
    End Sub



    Public Async Function GenerateAndPlayAudioFromSpeakerNotes(
        presentationFilePath As String,
        Optional languageCode As String = "en-US",
        Optional voiceName As String = "en-US-Studio-O",
        Optional voiceNameAlt As String = ""
    ) As System.Threading.Tasks.Task

        Dim ppApp As NetOffice.PowerPointApi.Application = Nothing
        Dim presentation As NetOffice.PowerPointApi.Presentation = Nothing

        Try
            '––– Load/save TTS settings –––
            Dim NoSSML As Boolean = My.Settings.NoSSML
            Dim Pitch As Double = My.Settings.Pitch
            Dim SpeakingRate As Double = My.Settings.Speakingrate
            Dim CleanText As Boolean = False
            Dim CleanTextPrompt As String = My.Settings.CleanTextPrompt
            If String.IsNullOrWhiteSpace(CleanTextPrompt) Then CleanTextPrompt = SP_CleanTextPrompt

            Dim params() As SLib.InputParameter = {
                New SLib.InputParameter("Pitch", Pitch),
                New SLib.InputParameter("Speaking Rate", SpeakingRate),
                New SLib.InputParameter("No SSML", NoSSML),
                New SLib.InputParameter("Clean text", CleanText)
            }
            If Not ShowCustomVariableInputForm("Parameters for audio generation:", "Create Audio (Slides)", params) Then
                Return
            End If

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

            Pitch = CDbl(params(0).Value)
            SpeakingRate = CDbl(params(1).Value)
            NoSSML = CBool(params(2).Value)
            CleanText = CBool(params(3).Value)
            My.Settings.NoSSML = NoSSML
            My.Settings.Pitch = Pitch
            My.Settings.Speakingrate = SpeakingRate
            My.Settings.Save()

            Dim useAlternate As Boolean = (voiceNameAlt <> "")
            Dim currentVoice As String = voiceName
            Dim firstUsed As Boolean = False

            '––– Open PowerPoint –––
            ppApp = New NetOffice.PowerPointApi.Application()
            presentation = ppApp.Presentations.Open(
                presentationFilePath,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse)

            If presentation.Slides.Count > 0 Then

                ShowProgressBarInSeparateThread($"{AN} Audio Generation", "Starting audio generation...")
                ProgressBarModule.CancelOperation = False
                GlobalProgressValue = 0
                GlobalProgressMax = presentation.Slides.Count

                For slideIndex As Integer = 1 To presentation.Slides.Count

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Or (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or ProgressBarModule.CancelOperation Then
                        ShowCustomMessageBox("Audio generation aborted by user.")
                        ProgressBarModule.CancelOperation = True
                        presentation.Close()
                        ppApp.Quit()
                        Return
                    End If

                    Dim slide As NetOffice.PowerPointApi.Slide = presentation.Slides(slideIndex)

                    ' 1) Find the notes-placeholder
                    Dim notesShape As NetOffice.PowerPointApi.Shape = Nothing
                    For i As Integer = 1 To slide.NotesPage.Shapes.Placeholders.Count
                        Dim shp = slide.NotesPage.Shapes.Placeholders(i)
                        If shp.PlaceholderFormat.Type = PpPlaceholderType.ppPlaceholderBody Then
                            notesShape = shp
                            Exit For
                        End If
                    Next
                    If notesShape Is Nothing Then Continue For

                    Dim notesText As String = notesShape.TextFrame.TextRange.Text.Trim()
                    If String.IsNullOrWhiteSpace(notesText) Then Continue For
                    If Not notesText.EndsWith(".") Then notesText &= "."

                    ' switch voice if needed
                    If useAlternate Then
                        If Not firstUsed Then
                            firstUsed = True
                        Else
                            currentVoice = If(currentVoice = voiceName, voiceNameAlt, voiceName)
                        End If
                    End If


                    If CleanText Then
                        ' Remove any unwanted characters from the paragraph text.
                        notesText = Await LLM(CleanTextPrompt, "<TEXTTOPROCESS>" & notesText & "</TEXTTOPROCESS>", "", "", 0, False, True)
                        notesText = notesText.Trim().Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "").Trim()
                        Debug.WriteLine("Cleaned notes = " & notesText & vbCrLf & vbCrLf)
                    End If

                    '––– Get audio bytes from TTS –––
                    Dim audioBytes As Byte() = Await GenerateAudioFromText(
                        notesText,
                        languageCode,
                        currentVoice,
                        NoSSML,
                        Pitch,
                        SpeakingRate,
                        "Slide " & slideIndex.ToString()
                    )
                    If audioBytes Is Nothing OrElse audioBytes.Length = 0 Then
                        Debug.WriteLine("[Debug] Slide " & slideIndex & ": no audio returned.")
                        Continue For
                    End If

                    '––– Save raw bytes as MP3 –––
                    Dim tempFile As String = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        $"ppt_audio_slide_{slideIndex}.mp3"
                    )
                    File.WriteAllBytes(tempFile, audioBytes)

                    Dim audioDurationSeconds As Double
                    Using mp3Reader As New NAudio.Wave.Mp3FileReader(tempFile)
                        audioDurationSeconds = mp3Reader.TotalTime.TotalSeconds
                    End Using

                    ' debug info
                    Dim exists As Boolean = File.Exists(tempFile)
                    Dim size As Long = If(exists, (New FileInfo(tempFile)).Length, -1L)
                    Debug.WriteLine($"[Debug] Slide {slideIndex}: tempFile='{tempFile}', Exists={exists}, Size={size}")

                    Dim beforeCount = slide.Shapes.Count
                    Debug.WriteLine($"[Debug] Slide {slideIndex}: Shapes before insert = {beforeCount}")

                    '––– Insert the MP3 –––
                    Dim mediaShape As NetOffice.PowerPointApi.Shape = Nothing
                    Try
                        mediaShape = slide.Shapes.AddMediaObject2(
                            fileName:=tempFile,
                            linkToFile:=MsoTriState.msoFalse,
                            saveWithDocument:=MsoTriState.msoTrue,
                            left:=10, top:=10,
                            width:=10, height:=10
                        )
                        Debug.WriteLine($"[Debug] AddMediaObject2 succeeded: Id={mediaShape.Id}, Type={mediaShape.Type}")

                        If mediaShape IsNot Nothing Then

                            '––– Ermittlung der Audio-Länge in Sekunden –––   ' Playback on entry + hide while not playing
                            With mediaShape.AnimationSettings.PlaySettings
                                .PlayOnEntry = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                                .HideWhileNotPlaying = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                            End With

                            ' 1 Sekunde Verzögerung vor Start des Audio
                            With mediaShape.AnimationSettings
                                .AdvanceMode = NetOffice.PowerPointApi.Enums.PpAdvanceMode.ppAdvanceOnTime
                                .AdvanceTime = 1     ' Sekunden bis zum Abspielen
                            End With

                            ' Slide automatisch advance nach (1s delay + Audio-Länge + 1s hold)
                            With slide.SlideShowTransition
                                .AdvanceOnTime = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                                .AdvanceTime = audioDurationSeconds + 1.0   ' 1s nach Ende
                                .AdvanceOnClick = NetOffice.OfficeApi.Enums.MsoTriState.msoFalse
                            End With


                        End If

                    Catch comEx As System.Runtime.InteropServices.COMException
                        Debug.WriteLine($"[Error] COMException in AddMediaObject2: {comEx}")
                        Continue For
                    End Try

                    '––– Configure play settings and initially hide upon play –––
                    With mediaShape.AnimationSettings.PlaySettings
                        .PlayOnEntry = MsoTriState.msoTrue
                        .HideWhileNotPlaying = MsoTriState.msoTrue
                    End With

                    Dim afterCount = slide.Shapes.Count
                    Debug.WriteLine($"[Debug] Slide {slideIndex}: Shapes after insert = {afterCount}")

                    ' Update the current progress value and status label.
                    GlobalProgressValue = slideIndex
                    GlobalProgressLabel = $"Slide {slideIndex} of {GlobalProgressMax} (some may be skipped)"

                    If File.Exists(tempFile) Then
                        Try
                            File.Delete(tempFile)
                        Catch ex As System.Exception
                            Debug.WriteLine($"[Warning] Could not delete temp file '{tempFile}': {ex.Message}")
                        End Try
                    End If

                Next

                ' save & clean up

                ProgressBarModule.CancelOperation = True

                Try
                    presentation.Save()
                    Debug.WriteLine("[Debug] Presentation.Save succeeded.")
                Catch comSaveEx As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"[Debug] Presentation.Save failed: {comSaveEx.Message}")
                    ' PowerPoint meldet schreibgeschützt → überschreibe per SaveAs
                    presentation.SaveAs(
                    presentationFilePath,
                    NetOffice.PowerPointApi.Enums.PpSaveAsFileType.ppSaveAsOpenXMLPresentation)
                    Debug.WriteLine("[Debug] Presentation.SaveAs succeeded.")
                End Try

                presentation.Close()
                ppApp.Quit()

                ShowCustomMessageBox("All slides with speaker notes have been amended with audio and auto-play.")

            Else
                ShowCustomMessageBox("No slides found in the presentation.")
                Return
            End If


        Catch ex As Exception
            Debug.WriteLine("[Error] Unexpected error: " & ex.ToString())
            ShowCustomMessageBox($"An unexpected error occurred when adding audio to the the slides ({ex.GetType().Name}): {ex.Message}")
        Finally
            If presentation IsNot Nothing Then presentation.Dispose()
            If ppApp IsNot Nothing Then ppApp.Dispose()
        End Try
    End Function


End Class

