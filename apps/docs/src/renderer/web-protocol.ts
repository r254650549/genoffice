export type OfficeWebMethod = "getLanguage" | "getSystemLocale" | "onLanguageChanged" | "getTheme" | "onThemeChanged" | "getDocumentTheme" | "onDocumentThemeChanged" | "getAutoSaveDefault" | "onAutoSaveDefaultChanged" | "getAiPanelPrefs" | "setAiPanelPrefs" | "onAiPanelPrefsChanged" | "onChromePressed" | "zoteroCommand" | "onZoteroRequest" | "respondToZotero" | "openDocx" | "openDocxPath" | "confirmDocumentReplace" | "openDocxDecrypt" | "convertAltChunkHtml" | "setDocPassword" | "docPasswordIntentRevision" | "discardDocPasswordIntents" | "consumePendingOpenDocx" | "consumeNewBlankDoc" | "consumeAiDocContent" | "consumeHeadlessExport" | "headlessExportDone" | "createDocument" | "onOpenDocx" | "onRenamedDocx" | "saveDocx" | "writeRecoveryCopy" | "onTeardown" | "respellKick" | "spellDiag" | "armContextMenu" | "claimContextMenu" | "onContextMenuRequest" | "spellAddWord" | "spellIgnoreWord" | "spellReplace" | "spellLanguages" | "spellSetLanguages" | "saveDocxAs" | "saveDocxNew" | "saveDocxTo" | "onMcpCommand" | "reportMcpResult" | "signalMcpReady" | "getRecentFiles" | "pickImage" | "fontMetrics" | "getAiSettings" | "setAiSettings" | "onAiSettingsChanged" | "openAiModelSettings" | "print" | "exportPdf" | "exportHtml" | "printPdfBuffer" | "saveMergedPdf" | "pickExportImagesTarget" | "takeExportPdf" | "writeExportImage" | "saveImageAs" | "onViewImage" | "aiChat" | "aiStream" | "aiStreamCancel" | "aiGskStatus" | "aiGskLogin" | "webSearch" | "imageSearch" | "analyzeMedia" | "fetchImage" | "aiGenerateImage" | "pickAttachments" | "addAttachmentPaths" | "addPastedImage" | "copyImageToClipboard" | "readAttachment" | "readAttachmentImage" | "kbList" | "kbSearch" | "kbRead" | "getPathForFile" | "openNewTab" | "openSourcePath" | "listDocsTabs" | "focusDocsTab" | "onAiStream" | "onMenuCommand" | "onCloseCheck" | "reportCloseCheck" | "onCloseSaveRequest" | "reportCloseSaveResult" | "reportViewMenuState"
export const OFFICE_WEB_CHANNEL = 'kompagent.genoffice.host/v1'
export type OfficeWebRequest = {
  channel: typeof OFFICE_WEB_CHANNEL
  kind: 'request' | 'notify'
  id: string
  method: OfficeWebMethod | 'session' | 'draft' | 'openAgent' | 'ready'
  args: unknown[]
}
export type OfficeWebResponse = {
  channel: typeof OFFICE_WEB_CHANNEL
  kind: 'response' | 'event'
  id?: string
  method?: OfficeWebMethod
  value?: unknown
  error?: string
}
