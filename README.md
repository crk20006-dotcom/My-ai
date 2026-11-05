# MyAI — NVIDIA Nemotron chat

A small, mobile-friendly chat app that connects directly to NVIDIA's hosted NIM chat-completions API. It follows NVIDIA's OpenAI-compatible `/v1/chat/completions` integration pattern and uses streamed responses.

## What it does

- Streaming chat replies.
- Coding help, explanations and debugging conversations.
- Reads text from PDF, DOCX, TXT, Markdown, CSV, XLSX, JSON and common source-code files.
- Sends PNG, JPG, JPEG and WEBP images to the Nemotron vision-language endpoint.
- Keeps the current chat in the active browser session and supports TXT export.
- Stores the NVIDIA key in Streamlit Secrets, not in GitHub source code.

## You do not need to install Python on your computer

You can deploy this project using a browser:

1. Download and extract the ZIP.
2. Create a GitHub repository named `MyAI` at https://github.com/new.
3. Upload `app.py`, `requirements.txt`, `README.md`, and `.gitignore` from the extracted folder. You may also upload the `.streamlit` folder only if it contains `secrets.toml.example`; never upload a real `secrets.toml`.
4. Open https://share.streamlit.io/ and sign in with GitHub.
5. Choose **Create app**, select your `MyAI` repository, branch `main`, and main file `app.py`.
6. Before/after deployment, open the app's **Settings -> Secrets** and add:

   ```toml
   NVIDIA_API_KEY = "nvapi-your-real-key"
   NVIDIA_MODEL = "nvidia/nemotron-3.5-super-vl-120b-a12b"
   ```

7. Save the settings. Streamlit restarts the app. Open the app URL to chat.

## NVIDIA endpoint

- Base URL: `https://integrate.api.nvidia.com/v1`
- Default model: `nvidia/nemotron-3.5-super-vl-120b-a12b`
- Endpoint style: OpenAI-compatible `chat.completions.create(..., stream=True)`.

Important model note (checked 5 November 2026): NVIDIA’s model catalog currently shows a “Deprecation in 12d” label for the Super VL model, even though its page still lists the free endpoint as available. Keep the model ID in Streamlit Secrets so it can be changed without editing Python. NVIDIA currently also lists `nvidia/nemotron-3.5-lightning-30b-a3b` with a free endpoint; that is a text-only alternative, so image analysis would not carry over.

## Important limits and privacy

- NVIDIA's hosted free endpoints are subject to NVIDIA's API Trial Terms, time/credit limits, rate limits, model availability and service changes. They are intended for prototyping/evaluation and are not an unlimited production service.
- Streamlit Community Cloud free apps can sleep after inactivity and do not guarantee 24/7 uptime.
- Text extracted from uploaded documents and attached image data are sent to NVIDIA as part of your prompt. Do not upload confidential material unless you are comfortable sending it to NVIDIA.
- Text-based PDFs can be read. Scanned/image-only PDFs may not have extractable text; upload the page as an image instead.
- The chat is kept in the active Streamlit session, not a permanent cross-device history database.
- MyAI displays generated code but does not execute it. Review generated code before running it.
- This starter is a chat/document assistant; it does not automatically browse the web or operate your PC.
