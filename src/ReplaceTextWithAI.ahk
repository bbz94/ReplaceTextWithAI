#Requires AutoHotkey v2.0

; === Replace selection with Azure OpenAI improved text ===
; Hotkey: Alt+F
!f::ImproveSelection()

; ---- Configuration ----
; Default endpoint/deployment set to your instance. Keep API key in env var:
;   AZURE_OPENAI_API_KEY    API key from Azure OpenAI
; Optional override:
;   AZURE_OPENAI_ENDPOINT   e.g. https://YOUR-RESOURCE-NAME.openai.azure.com/openai/v1
;   AZURE_OPENAI_DEPLOYMENT Your model deployment name

global AOAI_ENDPOINT := EnvGet("AZURE_OPENAI_ENDPOINT")
if (AOAI_ENDPOINT = "")
    AOAI_ENDPOINT := "https://{your-resource-name}.openai.azure.com/openai/v1"

global AOAI_DEPLOYMENT := EnvGet("AZURE_OPENAI_DEPLOYMENT")
if (AOAI_DEPLOYMENT = "")
    AOAI_DEPLOYMENT := "gpt-5-nano" ; Chepes o

global AOAI_API_KEY := EnvGet("AZURE_OPENAI_API_KEY")
if (AOAI_API_KEY = "")
    AOAI_API_KEY := "{your-api-key}" ; ADD YOUR API KEY HERE

; Optional prompt tuning
global SYSTEM_PROMPT := "You are a friendly, colleague-style editor. Keep the writer's voice and personality. Make minimal changes: fix only grammar, spelling, and sentence structure. Do not rewrite or significantly change the tone. Return only the improved text. If the input is in English, respond in English. If the input is in Latvian, respond in Latvian."

; Logging (simple file log)
global LOG_FILE := A_Temp "\\ReplaceTextWithAI.log"

ImproveSelection() {
    if (AOAI_ENDPOINT = "" || AOAI_API_KEY = "" || AOAI_DEPLOYMENT = "") {
        MsgBox "Azure OpenAI settings missing. Set AZURE_OPENAI_API_KEY (and optionally AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_DEPLOYMENT)."
        return
    }

    ; Preserve clipboard
    clipSaved := ClipboardAll()
    A_Clipboard := ""
    SendInput "^c"
    Sleep 50
    if !ClipWait(1) {
        A_Clipboard := clipSaved
        MsgBox "No text selected."
        return
    }

    inputText := A_Clipboard
    if (StrLen(Trim(inputText)) = 0) {
        A_Clipboard := clipSaved
        MsgBox "No text selected."
        return
    }

    improved := AzureOpenAIImproveText(inputText)
    if (improved = "") {
        A_Clipboard := clipSaved
        MsgBox "Azure OpenAI returned no text. Check log for details."
        return
    }

    ; Replace selection with improved text
    A_Clipboard := improved
    Send "^v"
    Sleep 50
    A_Clipboard := clipSaved
}

AzureOpenAIImproveText(text) {
    url := AOAI_ENDPOINT "/chat/completions"

    userPrompt := "Rewrite the following text to be professional, business-oriented, and clear. Keep the meaning. Return only the improved text. Text: " . text
    body := Format(
        '{"model":"{3}","messages":[{"role":"system","content":"{1}"},{"role":"user","content":"{2}"}]}',
        JsonEscape(SYSTEM_PROMPT),
        JsonEscape(userPrompt),
        JsonEscape(AOAI_DEPLOYMENT)
    )

    retries := 3
    delayMs := 500

    Loop retries {
        try {
            http := ComObject("WinHttp.WinHttpRequest.5.1")
            http.Open("POST", url, false)
            http.SetRequestHeader("Content-Type", "application/json")
            http.SetRequestHeader("api-key", AOAI_API_KEY)
            http.Send(body)

            status := http.Status
            response := GetResponseTextUtf8(http)
            Log("Status: " status " Response: " response)

            if (status >= 200 && status < 300) {
                return ExtractContentFromResponse(response)
            }

            ; Retry on transient errors
            if (status = 429 || status = 500 || status = 502 || status = 503 || status = 504) {
                Sleep delayMs
                delayMs *= 2
                continue
            }

            return ""
        } catch as err {
            Log("Exception: " err.Message)
            Sleep delayMs
            delayMs *= 2
        }
    }

    return ""
}

ExtractContentFromResponse(jsonText) {
    ; Extract the assistant content from the first choice (simple regex-based parsing)
    if RegExMatch(jsonText, '"content"\s*:\s*"((?:\\.|[^"\\])*)"', &m) {
        return JsonUnescape(m[1])
    }
    ; Fallback for content arrays with text objects
    if RegExMatch(jsonText, '"text"\s*:\s*"((?:\\.|[^"\\])*)"', &m2) {
        return JsonUnescape(m2[1])
    }
    return ""
}

JsonEscape(s) {
    quote := Chr(34)
    escQuote := "\\" . quote ; backslash + quote
    s := StrReplace(s, "\\", "\\\\")
    s := StrReplace(s, quote, escQuote)
    s := StrReplace(s, "`r", "\\r")
    s := StrReplace(s, "`n", "\\n")
    s := StrReplace(s, "`t", "\\t")
    return s
}

JsonUnescape(s) {
    quote := Chr(34)
    escQuote := "\\" . quote ; backslash + quote
    s := DecodeUnicodeEscapes(s)
    s := StrReplace(s, escQuote, quote)
    s := StrReplace(s, "\\n", "`n")
    s := StrReplace(s, "\\r", "`r")
    s := StrReplace(s, "\\t", "`t")
    s := StrReplace(s, "\\\\", "\\")
    return s
}

DecodeUnicodeEscapes(s) {
    pos := 1
    while RegExMatch(s, "\\\\u([0-9A-Fa-f]{4})", &m, pos) {
        repl := Chr("0x" m[1])
        s := SubStr(s, 1, m.Pos(0) - 1) . repl . SubStr(s, m.Pos(0) + m.Len(0))
        pos := m.Pos(0)
    }
    return s
}

GetResponseTextUtf8(http) {
    try {
        body := http.ResponseBody
        stream := ComObject("ADODB.Stream")
        stream.Type := 1 ; binary
        stream.Open()
        stream.Write(body)
        stream.Position := 0
        stream.Type := 2 ; text
        stream.Charset := "utf-8"
        text := stream.ReadText()
        stream.Close()
        return text
    } catch {
        return http.ResponseText
    }
}

Log(message) {
    timestamp := FormatTime(A_NowUTC, "yyyy-MM-dd HH:mm:ss")
    FileAppend("[" timestamp "] " message "`n", LOG_FILE)
}
