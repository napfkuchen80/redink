Red Ink is an alternative AI assistant add-in for Microsoft Office on Windows that allows you to access your preferred large language model (LLM) directly within Word, Excel and Outlook, enabling you to use the capabilities of your favorite LLM  in a variety of ways for your daily work:

- You can select a text or cells in your worksheet and have the content translated, corrected or improved.
- You can have summaries created, automatically abbreviate or anonymise text, or ask the AI to comment on your document according to your specifications (using Word comments/bubbles) or to analyse specific revisions to your document (‘What did the other side summarise and change?’).
- There is a context search (a search for ‘liability’, for example, not only finds ‘liability’ but also related passages such as ‘we are not liable’), a chatbot integrated into Word (which can also edit your text directly if you wish) and live transcription and dictation are also possible (with post-processing by the AI).
- You can write free prompts or access a prompt library with complex prompts, you can combine external documents with your own text (including searchable PDFs), you can link a small text database, you can have the AI create formulas and content for cells in Excel (‘Insert the formulas for a linear regression for the following case:...’) ) or have it filled in (‘Complete C10:C20 with suggestions for answers to the questions in B10:B20 based on the use case in C9’) or have a long e-mail chain summarised in Outlook before you reply to it.
- If suggestions for changes are made, you can also receive these as markup – not only in Word, but also in Outlook, although Outlook does not actually support this.
- There is a browser extension so that you can use Red Ink in your Chromium-based browser (Edge, Chrome) (e.g. to search the content of webpages or PDFs using AI) or to translate or correct texts that you are supposed to fill out in a form.
- You can have your texts converted into audiobooks and podcasts, similar to Google's NotebookLM function (this requires a Google Cloud Platform account, though).

Unlike some other offerings, Red Ink allows you to freely choose the AI provider you want to work with. If you don't trust any of these providers or they don't offer you the necessary assurances, you can also configure Red Ink to use a self-hosted LLM. The source code is open, and we have no access to your data. We don't want that either, because we originally developed the tool for ourselves. For the time being, we provide the tool for free. Since API access is also very cheap today, this solution is much less expensive to use for such applications than some other offers.

Red Ink has originally been developed by David Rosenthal, partner at the Swiss law firm VISCHER and head of its data & privacy and AI practice for its own internal purposes, but is now making available Generation 2 of the tool to the public. You will find more information (including a demo video) on https://vischer.com/redink and the executables on https://apps.vischer.com.

In the current beta test phase the software is free to everyone, and the source code is open here for review. The license conditions are described in the license file.

## Additional deployment documentation

- [Outlook Gateway Integration Checklist](Documentation%20and%20more/Outlook_Gateway_Integration.md) – network, authentication, policy, and offline requirements for connecting Outlook clients to an internal Red Ink gateway.
