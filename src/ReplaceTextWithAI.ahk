#Requires AutoHotkey v2.0

; === Replace selection with Azure OpenAI improved text ===
; Hotkey: Alt+R
!r:: ImproveSelection()
; Hotkey: Alt+F (friendlier rewrite)
!f:: ImproveSelection(FRIENDLIER_PROMPT)
; Hotkey: Alt+P (improve prompt text)
!p:: ImproveSelection(PROMPT_IMPROVER_PROMPT)
; Hotkey: Alt+U (convert to user story)
!u:: ImproveSelection(USER_STORY_PROMPT)

; ---- Configuration ----
; Default endpoint/deployment set to your instance. Keep API key in env var:
;   AZURE_OPENAI_API_KEY    API key from Azure OpenAI
; Optional override:
;   AZURE_OPENAI_ENDPOINT   e.g. https://YOUR-RESOURCE-NAME.openai.azure.com/openai/v1
;   AZURE_OPENAI_DEPLOYMENT Your model deployment name

global AOAI_ENDPOINT := GetEnvOrDefault("AZURE_OPENAI_ENDPOINT",
    "https://{your-resource-name}-sc.openai.azure.com/openai/v1")
global AOAI_DEPLOYMENT := GetEnvOrDefault("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini")
global AOAI_API_KEY := GetEnvOrDefault("AZURE_OPENAI_API_KEY", "{your-api-key}")

; Optional prompt tuning
global SYSTEM_PROMPT :=
    "You are a friendly, colleague-style editor. Keep the writer's voice and personality. Make minimal changes: fix only grammar, spelling, and sentence structure. Do not rewrite or significantly change the tone. Return only the improved text. If the input is in English, respond in English. If the input is in Latvian, respond in Latvian."
global FRIENDLIER_PROMPT :=
    "You are a friendly, colleague-style editor. Keep the writer's voice and personality. Make minimal changes while making the tone kinder and more collaborative. Avoid harsh or blunt wording. Keep the meaning and intent. Return only the improved text. If the input is in English, respond in English. If the input is in Latvian, respond in Latvian."
global PROMPT_IMPROVER_PROMPT :=
    "You are a prompt editor. You improve AI prompts. Rewrite the input prompt to be clearer, more specific, and more effective, while preserving intent. Use concise, direct language, add helpful constraints if missing, and avoid unnecessary verbosity. Do not ask questions. Do not add requests for clarification. Return only the improved prompt. If the input is in English, respond in English. If the input is in Latvian, respond in Latvian."
global USER_STORY_PROMPT :=
    "You are a business analyst and technical writer. Convert the input text into the exact User Story Template below and auto-populate all required fields. Preserve the original intent. If details are missing, infer sensible defaults from context and keep them realistic. Always output all sections, never leave placeholders like [user] or [goal]. For 'ADR Needed', output only Yes or No based on architectural impact (cross-team impact, significant technical decision, new integration, security/compliance, or major design tradeoff = Yes; otherwise No). Return only the completed template, with no extra commentary. IMPORTANT for Azure DevOps work item rendering: section titles must be output exactly as shown below using Unicode bold words. If input is English, respond in English. If input is in Latvian, respond in Latvian.`n`n𝗪𝗵𝘆`nAs a <user>, I want <goal>, so that <reason>.`n`n𝗪𝗵𝗮𝘁`n<Describe the feature or function that will meet the user's goal>.`n`n𝗛𝗼𝘄`n<Describe the technical details of implementation, including resources, programming languages, frameworks, integrations, and constraints>.`n`n𝗧𝗶𝗺𝗲𝗹𝗶𝗻𝗲`n<Provide an estimated timeline, milestones, and dependencies on teams/projects>.`n`n𝗔𝗗𝗥 𝗡𝗲𝗲𝗱𝗲𝗱 :`n<Yes/No>"
; Logging (simple file log)
global LOG_FILE := A_Temp "\\ReplaceTextWithAI.log"

ImproveSelection(systemPrompt := SYSTEM_PROMPT) {
    if (AOAI_ENDPOINT = "" || AOAI_API_KEY = "" || AOAI_DEPLOYMENT = "") {
        MsgBox "Azure OpenAI settings missing. Set AZURE_OPENAI_API_KEY (and optionally AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_DEPLOYMENT)."
        return
    }

    clipSaved := ""
    inputText := GetSelectedText(&clipSaved)
    if (inputText = "")
        return

    improved := AzureOpenAIImproveText(inputText, systemPrompt)
    if (improved = "") {
        A_Clipboard := clipSaved
        MsgBox "Azure OpenAI returned no text. Check log for details."
        return
    }

    ; Fix literal \n/\r sequences if the model returned escaped newlines
    improved := NormalizeNewlines(improved)

    ; Replace selection with improved text
    SendWithShiftEnter(improved)
    Sleep 50
    A_Clipboard := clipSaved
}

AzureOpenAIImproveText(text, systemPrompt) {
    url := AOAI_ENDPOINT "/chat/completions"
    body := BuildRequestBody(text, systemPrompt)

    retries := 3
    delayMs := 500

    loop retries {
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

BuildRequestBody(text, systemPrompt) {
    userPrompt := text
    return Format(
        '{"model":"{3}","messages":[{"role":"system","content":"{1}"},{"role":"user","content":"{2}"}]}',
        JsonEscape(systemPrompt),
        JsonEscape(userPrompt),
        JsonEscape(AOAI_DEPLOYMENT)
    )
}

GetEnvOrDefault(name, defaultValue) {
    value := EnvGet(name)
    return (value = "") ? defaultValue : value
}

GetSelectedText(&clipSaved) {
    ; Preserve clipboard
    clipSaved := ClipboardAll()
    A_Clipboard := ""
    SendInput "^c"
    Sleep 50
    if !ClipWait(1) {
        A_Clipboard := clipSaved
        MsgBox "No text selected."
        return ""
    }

    inputText := A_Clipboard
    if (StrLen(Trim(inputText)) = 0) {
        A_Clipboard := clipSaved
        MsgBox "No text selected."
        return ""
    }

    return inputText
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
    ; Handle double-escaped and normal escaped control characters
    s := StrReplace(s, "\\n", "`n")
    s := StrReplace(s, "\\r", "`r")
    s := StrReplace(s, "\\t", "`t")
    s := StrReplace(s, "\n", "`n")
    s := StrReplace(s, "\r", "`r")
    s := StrReplace(s, "\t", "`t")
    s := StrReplace(s, "\\\\", "\\")
    return s
}

NormalizeNewlines(s) {
    ; Normalize non-breaking spaces
    s := StrReplace(s, Chr(0xA0), " ")
    ; Iterate to handle chained escaping like "\\\\n" -> "\\n" -> newline
    loop 3 {
        old := s
        ; Handle double-escaped and normal escaped control characters
        s := StrReplace(s, "\\r\\n", "`r`n")
        s := StrReplace(s, "\\n", "`n")
        s := StrReplace(s, "\\r", "`r")
        s := StrReplace(s, "\r\n", "`r`n")
        s := StrReplace(s, "\n", "`n")
        s := StrReplace(s, "\r", "`r")
        ; Handle spaced escaped sequences like " \n " or " \r\n "
        s := RegExReplace(s, "[ \t\xA0]*\\\\r\\\\n[ \t\xA0]*", "`r`n")
        s := RegExReplace(s, "[ \t\xA0]*\\\\n[ \t\xA0]*", "`r`n")
        ; Handle one-or-more backslashes + optional spaces + n/r (with or without surrounding spaces)
        s := RegExReplace(s, "[ \t\xA0]*\\\\+[ \t\xA0]*n[ \t\xA0]*", "`r`n")
        s := RegExReplace(s, "[ \t\xA0]*\\\\+[ \t\xA0]*r[ \t\xA0]*", "`r")
        if (s = old)
            break
    }
    ; Ensure Teams-friendly CRLF newlines
    marker := Chr(1)
    s := StrReplace(s, "`r`n", marker)
    s := StrReplace(s, "`r", "`r`n")
    s := StrReplace(s, "`n", "`r`n")
    s := StrReplace(s, marker, "`r`n")
    return s
}

SendWithShiftEnter(text) {
    ; Convert literal backslash-n sequences to real newlines
    text := RegExReplace(text, "[ \t\xA0]*\\\\+[ \t\xA0]*r[ \t\xA0]*\\\\+[ \t\xA0]*n[ \t\xA0]*", "`r`n")
    text := RegExReplace(text, "[ \t\xA0]*\\\\+[ \t\xA0]*n[ \t\xA0]*", "`r`n")
    text := RegExReplace(text, "[ \t\xA0]*\\\\+[ \t\xA0]*r[ \t\xA0]*", "`r`n")
    ; Normalize to CRLF and split into lines
    text := NormalizeNewlines(text)
    lines := StrSplit(text, "`r`n")
    ; Remove empty/whitespace-only lines to avoid extra blank lines
    filtered := []
    for _, line in lines {
        if (StrLen(Trim(line)) > 0)
            filtered.Push(line)
    }
    lines := filtered
    if (lines.Length = 0)
        return
    ; Send first line, then Shift+Enter between lines
    SendText(lines[1])
    if (lines.Length > 1) {
        loop lines.Length - 1 {
            Send "+{Enter}"
            SendText(lines[A_Index + 1])
        }
    }
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
