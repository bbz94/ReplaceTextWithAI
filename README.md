# ReplaceTextWithAI

AutoHotkey v2 script that rewrites selected text using Azure OpenAI and replaces the selection in-place. Supports Latvian and English output based on the input language.

## Features
- Hotkey: Alt + R - Improve selected text (grammar/spelling/structure, minimal tone changes)
- Hotkey: Alt + F - Rewrite selected text in a friendlier, more collaborative tone
- Hotkey: Alt + P - Improve selected prompt text for clarity and effectiveness
- Hotkey: Alt + U - Convert selected text into a structured User Story template
- Uses Azure OpenAI chat completions
- Language-aware output (Latvian/English)
- Preserves clipboard after replacement
- Simple retry/backoff for transient errors

## Quick Start
1. Install AutoHotkey v2.
2. Set environment variables:
   - AZURE_OPENAI_API_KEY
   - (optional) AZURE_OPENAI_ENDPOINT
   - (optional) AZURE_OPENAI_DEPLOYMENT
3. Run src/ReplaceTextWithAI.ahk.
4. Select text and use one of these hotkeys:
    - Alt + R: standard text improvement
    - Alt + F: friendlier rewrite
    - Alt + P: prompt improvement
    - Alt + U: user story conversion

## Configuration
Defaults are set in src/ReplaceTextWithAI.ahk:
- Endpoint: https://{your-resource-name}.openai.azure.com/openai/v1
- Deployment: gpt-4o-mini

Override with environment variables if needed.

## Logs
Log file: %TEMP%\ReplaceTextWithAI.log

## Folder Structure
- src/ - main AutoHotkey script
- icons/ - assets

## Compile to EXE (AutoHotkey v2)

1. Install AutoHotkey v2 (includes Ahk2Exe).
    1. `choco install autohotkey.install`
2. Open Ahk2Exe from the AutoHotkey installation folder.
3. Source: `src/ReplaceTextWithAI.ahk`
4. Destination: choose your output path, for example `dist/ReplaceTextWithAI.exe`.
5. Base file: use the AutoHotkey v2 base executable.
6. Icon: choose an `.ico` from the `icons/` folder.
7. Click Convert.