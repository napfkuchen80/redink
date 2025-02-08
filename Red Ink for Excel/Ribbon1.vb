' Red Ink Ribbon Code
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See https://vischer.com/redink for more information.
'
' 8.2.2025

Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Public Async Function RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Correct.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.Correct()
    End Function

    Public Async Function RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Correct2.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.Correct()
    End Function

    Public Async Function RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Shorten.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.Shorten()
    End Function

    Public Async Function RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Primlang.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_PrimLang2.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function RI_SecLang_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_SecLang.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage2()
    End Function
    Public Async Function RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Improve.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.Improve()
    End Function

    Public Async Function RI_FreestyleNM_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleNM.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function RI_FreestyleNM2_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleNM2.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function RI_Anonymize_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Anonymize.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.Anonymize()
    End Function

    Public Sub RI_AdjustHeight_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AdjustHeight.Click
        Globals.ThisAddIn.AdjustHeight()
    End Sub

    Public Sub RI_AdjustLegacyNotes_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AdjustLegacyNotes.Click
        Globals.ThisAddIn.AdjustLegacyNotes()
    End Sub

    Private Async Function RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Translate.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InOther()
    End Function

    Private Async Function RI_TranslateF_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_TranslateF.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InOtherFormulas()
    End Function

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Async Function RI_FreestyleAM_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleAM.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleAM()
    End Function

    Private Async Function RI_SwitchParty_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_SwitchParty.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.SwitchParty()
    End Function

    Private Sub RI_Regex_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Regex.Click
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub


End Class