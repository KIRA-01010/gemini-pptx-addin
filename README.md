# Gemini Slides — PowerPoint Add-in

A PowerPoint task pane add-in (the official extension mechanism for Office
apps) that calls the **Gemini API** to generate a slide deck from a topic or
outline, then builds real slides directly in the open presentation — similar
in spirit to SlidesAI for Google Slides.

## How it works

1. You type a topic/outline, pick a slide count and tone, and paste a Gemini
   API key (stored only in your browser's local storage).
2. The add-in sends one prompt to `models/{model}:generateContent`, asking
   Gemini to return structured JSON (`{title, bullets, notes}` per slide) via
   the API's `responseMimeType: "application/json"` mode.
3. It then uses the **PowerPoint JavaScript API** (`PowerPoint.run`) to add a
   slide per item and draw a title text box + bullet text box on each one, in
   your currently open presentation — no download/upload round trip.

## Fastest path to a working add-in (~5 min, free, no cert hassle)

PowerPoint requires add-in files to be served over HTTPS from a real domain —
there's no "just open a file" mode, on any platform. **GitHub Pages** gives
you free HTTPS with zero certificate setup, which makes it the least-friction
option:

1. Push this folder to a new GitHub repo (or use the GitHub web UI's
   "upload files" — no git CLI needed).
2. In that repo: **Settings → Pages → Deploy from branch → main → / (root)**.
   GitHub gives you a URL like `https://yourname.github.io/gemini-pptx-addin`.
3. Run the included script to point the manifest at that URL:
   ```bash
   ./set-domain.sh yourname.github.io/gemini-pptx-addin
   ```
   (Re-push `manifest.xml` after this, or just edit it in the GitHub web UI —
   it's the only file that changes.)
4. In PowerPoint: **Insert → Add-ins → Upload My Add-in** → select the
   updated `manifest.xml`.

That's it — steps 1-2-4 are unavoidable no matter which add-in you use
(SlidesAI's Google Slides installer just hides the same kind of registration
behind Google's marketplace). Steps 1-3 are one-time; after that, everyone
who sideloads the manifest gets the "Generate Slides" ribbon button.

## Project layout

```
manifest.xml                  Add-in manifest (registers the ribbon button)
src/taskpane/taskpane.html    Task pane UI
src/taskpane/taskpane.css     Styles
src/taskpane/taskpane.js      Gemini call + PowerPoint slide-building logic
src/commands/                 Required function-file stub for the manifest
assets/                       Ribbon icons (16/32/80px)
```

There's no build step — it's plain HTML/CSS/JS, loaded straight in the task
pane's embedded browser, so you can host it on literally any static file host.

## Alternative: run it locally first (for testing changes)

If you want to test edits before deploying, Office Add-ins still need HTTPS
even on localhost:

```bash
npm install -g http-server office-addin-dev-certs
npx office-addin-dev-certs install          # trust a local dev certificate
http-server . -S -C ~/.office-addin-dev-certs/localhost.crt \
                 -K ~/.office-addin-dev-certs/localhost.key -p 3000
./set-domain.sh localhost:3000
```

### Sideload into PowerPoint

- **PowerPoint on the web / Microsoft 365**: Insert → Add-ins → *Upload My
  Add-in* → select `manifest.xml`.
- **PowerPoint desktop (Windows/Mac)**: Insert → My Add-ins → *Upload My
  Add-in* (or place `manifest.xml` in the shared network/sideload folder your
  org uses for LOB add-ins).

A **Generate Slides** button appears on the Home ribbon; clicking it opens
the task pane described above.

## Deploying for real use

For anything beyond your own testing, host the files somewhere permanent
(GitHub Pages, Azure Static Web Apps, S3+CloudFront, Netlify, etc.), point
every URL in `manifest.xml` at that domain, and either:

- distribute the manifest directly to users/org admins to sideload, or
- publish to your org's Microsoft 365 admin center as a private/LOB add-in, or
- submit to **AppSource** for public distribution.

## API key handling

The key is stored in the task pane's `localStorage` and sent only in
requests straight from the user's browser to
`generativelanguage.googleapis.com` — it never passes through any server of
yours. For a team deployment, you may prefer to proxy the Gemini call through
your own backend so you can use a shared, server-side key instead of asking
each user for their own.

## Known limitations

- **Speaker notes**: the PowerPoint JavaScript API does not currently expose
  the real speaker-notes pane (this is a known, still-open gap in Office.js).
  As a stand-in, when "Add speaker notes" is checked, the add-in drops a
  small italic note box near the bottom of the slide instead of true notes.
- **Layout**: slides are built from plain text boxes positioned for a
  standard 13.33"×7.5" widescreen slide, not your theme's actual title/content
  placeholders (the JS API doesn't yet support inserting into inherited
  layout placeholders), so generated slides won't inherit your master
  slide's placeholder styling automatically — only manually placed text/box
  formatting. You can apply a theme afterward via Design → Themes, and text
  in the boxes will pick up your theme's default fonts/colors.
- **Model names**: the model dropdown lists a couple of current Gemini
  models, but Google updates these regularly — use "Custom model name…" if
  you want to point at something newer.

## Extending it

Ideas if you want to take this further:
- Generate a matching image per slide with Imagen and insert it via
  `shapes.addPicture` (base64).
- Let users edit the generated JSON outline before it's inserted, instead of
  going straight from prompt to slides.
- Add an "Improve this slide" action scoped to the currently selected slide.
