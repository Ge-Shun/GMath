# GMath

![GMath cover](assets/readme/gmath-en.png)

[简体中文](READMEZH.md)

`GMath` is a Microsoft Word add-in for writing editable math equations on macOS.
It provides a task pane for visual equation editing, LaTeX input, quick symbol
templates, image-to-formula recognition, and insertion as native Word equations.

The inserted result is not an image. It remains editable in Word.

This project is currently validated primarily on Microsoft Word for Mac. The
manifest requires WordApi 1.1, so other Word platforms may install it, but
Windows, web, and iPad aren't yet part of the supported test matrix.

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

The endpoint and model are stored in local browser storage. The API key is kept
only for the current task-pane session by default; it is persisted only when you
select **Remember the key on this device**. During recognition, the image and key
are sent to the provider you configured.

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

The local proxy accepts only the hosted GMath page and local task pane, requires
public HTTPS targets, and blocks private/loopback addresses. Set the optional,
comma-separated `GMATH_AI_HOSTS` to restrict target hosts further. For a trusted
local model endpoint, explicitly set `GMATH_ALLOW_PRIVATE_AI=1`; an HTTP endpoint
also requires `GMATH_ALLOW_INSECURE_AI=1`.

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

## Testing

```bash
npm test
npm run check
npm run validate
```

The suite covers authored conversion fixtures, real MathLive serialization,
selection/lossy-conversion guards, and the local proxy's origin, token, target,
and static-path boundaries.

## Remove the Add-in

Delete the sideloaded manifest:

```bash
rm ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

Then restart Word.

## Third-Party Software

GMath uses MathLive, Fraunces, and JetBrains Mono. See `THIRD_PARTY_NOTICES.md`
and the license files shipped under `src/vendor` for licenses and sources.
