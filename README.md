# GMath

![GMath cover](assets/readme/gmath-en.png)

[简体中文](READMEZH.md)

`GMath` is a Microsoft Word add-in for writing editable math equations on macOS.
It provides a task pane for visual equation editing, LaTeX input, quick symbol
templates, image-to-formula recognition, and insertion as native Word equations.

The inserted result is not an image. It remains editable in Word.

This project is currently focused on Microsoft Word for Mac. Windows sideloading
is not packaged in this repository yet.

## Highlights

- Edit equations visually in a Word task pane.
- Type LaTeX and keep it synchronized with the visual editor.
- Insert fractions, roots, scripts, braces, matrices, sums, integrals, limits,
  accents, vectors, and common symbols from the quick-pick palette.
- Insert inline, display, and right-numbered Word equations.
- Convert formula screenshots to LaTeX through an OpenAI-compatible vision API.
- Use the hosted GitHub Pages task pane for normal Word usage, with an optional
  local proxy for APIs that block browser CORS requests.

## Requirements

- macOS
- Microsoft Word for Mac
- Node.js and npm

## Installation

Install dependencies from the repository root:

```bash
npm install
```

Sideload the add-in into Word for Mac:

```bash
npm run sideload:mac
```

Quit Word completely with `Cmd+Q`, then open Word again. GMath should appear on
the **Home** tab.

The default sideload manifest points Word to:

```text
https://ge-shun.github.io/GMath/src/taskpane.html
```

Run `npm run sideload:mac` again after updating the repository if Word keeps an
old cached manifest or task pane.

## Usage

1. Open Word.
2. Choose **Home** > **GMath** > **Insert Equation**.
3. Build the equation in the task pane using visual editing, quick symbols, or
   LaTeX source.
4. Choose **Inline**, **Display**, or **Numbered**.
5. Click **Insert into Word**.

## Image To Formula

Open **API Settings** in the task pane and enter:

- An OpenAI-compatible chat completions endpoint, such as
  `https://api.openai.com/v1/chat/completions`
- An API key
- A vision-capable model name

The endpoint and key are stored in the task pane's local browser storage. During
recognition, the selected image and API key are sent to the provider you
configured.

For providers that reject browser CORS requests, start the temporary proxy:

```bash
npm run dev-certs
npm run serve
```

Keep that terminal running while using image recognition. The proxy endpoint is:

```text
https://localhost:3000/api/ai/chat/completions
```

GMath does not install a persistent background service.

## Development

To point Word at the local task pane instead of GitHub Pages:

```bash
npm run dev-certs
npm run serve
```

In another terminal:

```bash
npm run sideload:mac:local
```

Restart Word completely. The local task pane URL is:

```text
https://localhost:3000/src/taskpane.html
```

To switch back to the hosted task pane:

```bash
npm run sideload:mac
```

## Troubleshooting

If GMath does not appear in Word, rerun `npm run sideload:mac` and restart Word
with `Cmd+Q`.

If Word still shows an old UI after an update, restart Word completely. Office
can cache add-in manifests and task pane assets.

If image recognition fails, verify that the API endpoint is OpenAI-compatible,
the model supports image input, and `npm run serve` is running when a CORS proxy
is required.

If an equation inserts incorrectly, open the debug section at the bottom of the
task pane and inspect the generated MathML / OMML. Some advanced structures may
not be supported yet.

## Remove the Add-in

Delete the sideloaded manifest:

```bash
rm ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

Then restart Word.

## Third-Party Software

GMath uses MathLive for visual equation editing. MathLive is licensed under the
MIT License.
