# MyAI — Personal AI workspace

A Streamlit personal assistant using NVIDIA's OpenAI-compatible chat completions API.

## Deploy
1. Push `app.py`, `requirements.txt`, and `README.md` to your GitHub repository.
2. Deploy the repository on Streamlit Community Cloud with `app.py` as the entry point.
3. Add this to Streamlit app Settings → Secrets:

```toml
NVIDIA_API_KEY = "nvapi-your-existing-key"
NVIDIA_MODEL = "nvidia/nemotron-3.5-super-vl-120b-a12b"
```

4. Reboot the app after saving Secrets.

## In-chat capabilities
- File attachment is built into the message composer (not a separate uploader panel).
- Supports PDF, DOCX, PPTX, XLSX, XLS, CSV, TSV, text, common code files, and PNG/JPG/WEBP images.
- Uploaded text and images remain available for follow-up questions within the active chat.
- Voice recording is built into the composer; optional transcription uses SpeechRecognition's Google Web Speech recognizer.
- Browser-native text-to-speech for assistant responses.
- NVIDIA streamed responses, chat history during the active Streamlit session, and TXT export.

## Notes and honest limitations
- Use Streamlit Secrets; never commit API keys to GitHub. The API key is used server-side.
- File content/images are sent to NVIDIA when the user sends a message. Do not upload sensitive documents unless comfortable with that.
- PDF extraction reads selectable text; scanned PDFs need OCR, which is not included.
- XLS/XLSX and PPTX are parsed locally; large files are capped to keep requests manageable.
- Browser speech synthesis depends on the device/browser. Voice recording/transcription depends on microphone permission, network, and SpeechRecognition service availability.
- Chat history is session-based, not permanent cross-device storage. Streamlit Community Cloud may sleep on inactivity.
- This is an advanced personal assistant UI, not a full ChatGPT clone: it does not include hosted file search, code execution sandbox, real-time duplex voice, persistent cloud memory, or guaranteed unlimited free inference.
- NVIDIA API model availability, quotas, and pricing can change; check the NVIDIA model catalogue for the exact model ID and current limits.


## Companion conversation preference
The sidebar has a **Conversation preferences** setting. The user can confirm they are 18+ to enable adult-themed consensual fictional conversation instructions. When off, the assistant can still be affectionate and romantic but stays non-explicit. The setting cannot override restrictions imposed by NVIDIA's model/API or applicable rules.
