# ReplaceTextWithAI

AutoHotkey v2 script that rewrites selected text using Azure OpenAI and replaces the selection in-place. Supports Latvian and English output based on the input language.

## Features
- Hotkey: Alt + F
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
4. Select text and press Alt + F.

## Configuration
Defaults are set in src/ReplaceTextWithAI.ahk:
- Endpoint: https://{your-resource-name}.openai.azure.com/openai/v1
- Deployment: gpt-5-nano

Override with environment variables if needed.

## Logs
Log file: %TEMP%\ReplaceTextWithAI.log

## Folder Structure
- src/ — main AutoHotkey script
- icons/ — assets

## Compile to EXE (AutoHotkey v2)

1. Install AutoHotkey v2 (includes Ahk2Exe).
    1. `choco install autohotkey.install`
2. Open Ahk2Exe from the AutoHotkey installation folder.
3. Source: `src/ReplaceTextWithAI.ahk`
4. Destination: choose your output path, for example `dist/ReplaceTextWithAI.exe`.
5. Base file: use the AutoHotkey v2 base executable.
6. Click Convert.

Tip: If you add a custom icon, select it in Ahk2Exe during conversion.

## GitHub Pages
See docs/index.html for a simple installation guide page.