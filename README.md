# GMath

GMath is a Microsoft Word add-in for writing and inserting editable math equations.

It lets you compose an equation visually or with LaTeX, then insert it into Word as a native Word equation. The inserted equation is not an image, so it stays editable inside the document.

中文说明见 [READMEZH.md](READMEZH.md).

## Features

- Visual equation editing in a Word task pane
- LaTeX input with live syncing
- Quick buttons for fractions, square roots, powers, subscripts, sums, integrals, limits, and matrices
- Inline or display-style equation insertion
- Native editable Word equations

## Requirements

- Microsoft Word
- Node.js and npm
- Internet access the first time the editor loads, because the current build loads MathLive from a CDN
- A trusted localhost certificate for local side-loading

## Run Locally

From the project folder:

```bash
npm install
npm run dev-certs
npm run serve
```

Keep the server running. It serves the add-in at:

```text
https://localhost:3000/src/taskpane.html
```

In a second terminal, side-load the add-in into Word on macOS:

```bash
npm run sideload:mac
```

Restart Word completely. You should see **GMath** on the **Home** tab.

## Use GMath

1. Open Word.
2. Go to **Home** > **GMath** > **Insert Equation**.
3. Type an equation visually, use the symbol buttons, or enter LaTeX directly.
4. Choose whether the equation should be inserted as a display equation.
5. Click **Insert into Word**.

The equation will be inserted as a native Word equation and can be edited later in Word.

## Troubleshooting

If the add-in does not appear, run `npm run sideload:mac` again and restart Word.

If the task pane is blank, make sure `npm run serve` is still running and that the localhost certificate has been installed with `npm run dev-certs`.

If Word still shows an old icon or old interface, restart Word completely. Office can cache add-in manifests and images.

If an equation is inserted incorrectly, open the debug section in the task pane and check the generated MathML and OMML. Some advanced MathML structures may not be supported yet.

## Third-Party Software

GMath uses MathLive for visual equation editing. MathLive is licensed under the MIT License.

## Remove the Add-in

On macOS, remove the side-loaded manifest:

```bash
rm ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

Then restart Word.
