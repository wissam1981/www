# OpenAI Word Assistant

Task pane add-in for Microsoft Word that sends the current selection to the OpenAI API and inserts the response back into the document.

## Prerequisites

- [Node.js](https://nodejs.org/) 18 or later.
- [Yeoman Office Add-in generator](https://www.npmjs.com/package/generator-office) (for reference) and the [Office Add-in CLI](https://www.npmjs.com/package/@microsoft/office-addin-cli) if you plan to use additional tooling.
- Word on the web (Office.com) or Word desktop (Microsoft 365 subscription) with sideloading enabled.
- An OpenAI API key created in the [OpenAI API platform](https://platform.openai.com/) with an active billing account. A ChatGPT login is not sufficient.

## Install dependencies

```bash
npm install
```

## Run the development server

```bash
npm start
```

The script starts an HTTPS static server on <https://localhost:3000>. The first run installs Microsoft development certificates if they are missing.

## Sideload into Word

### Word on the web

1. Run `npm start` in this folder.
2. Open <https://www.office.com/launch/word> and create a blank document.
3. Choose **Insert** → **Office Add-ins** → **Upload My Add-in**.
4. Browse to `manifest.xml` in this project and upload it. The task pane appears on the right.

### Word on Windows or Mac

1. Run `npm start` in this folder.
2. In another terminal, run:
   ```bash
   npm run sideload
   ```
   This uses the Office Add-in debugging tools to register the manifest.
3. Open Word. The **OpenAI Assistant** button appears on the **Home** tab.
4. Stop debugging with `npm run stop` when you are finished.

## Using the add-in

1. Open the **OpenAI Word Assistant** task pane.
2. Enter your OpenAI API key in **API Settings** and select **Save key**.
3. Choose an operation (Rewrite, Summarize, Translate, Fix Grammar), optionally set tone and language, and adjust the prompt or model.
4. Select text in Word and choose **Run on Selection** to insert the response below the selection or **Replace Selection** to overwrite it. Use **Insert at Cursor** when no text is selected.
5. The status area shows request progress, usage tokens, and any errors. The telemetry list keeps the five most recent operations.

## Security notes

- The API key is stored with `OfficeRuntime.storage` when available. On hosts that do not expose it, the key falls back to `localStorage`.
- For production, move API calls behind a secure backend proxy and avoid storing long-lived secrets in the client.

## Project structure

- `manifest.xml` – Word manifest pointing to the local dev server.
- `public/taskpane.html` – Task pane UI.
- `public/taskpane.js` – Office.js logic and OpenAI integration.
- `dev-server.js` – Minimal HTTPS static server.
- `assets/` – Placeholder icons.

## Troubleshooting

- If Word cannot load the add-in due to certificate errors, run `npx office-addin-dev-certs install` and restart the dev server.
- HTTP 401 indicates an invalid or missing API key. HTTP 429 indicates rate limiting from OpenAI.
