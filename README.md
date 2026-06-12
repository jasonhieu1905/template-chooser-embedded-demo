# Template Chooser Embedded Outlook Demo

This is a plain HTML, JavaScript, and CSS Outlook add-in. It has no build step and no dependency on the officeatwork Angular application.

## Deploy

1. Upload every file in this directory to the root of an HTTPS static website.
2. Replace `https://tc-embedded-demo.example.com` in `manifests/template-chooser/template-chooser-embedded-demo.outlook.manifest.local.xml` with the website origin.
3. Sideload the updated manifest in Outlook.
4. Compose a message or appointment and open **TC Embedded Demo** from the ribbon.

Opening `index.html` directly in a browser loads the Office.js library, but it does not create an Office host context. In that case `Office.context.ui` is unavailable by design. The page must be opened from the sideloaded add-in command in Outlook.

The website must serve these paths without authentication:

- `/index.html`
- `/app.js`
- `/styles.css`
- `/dialog-redirect.html`
- `/function-file.html`

Do not configure `Cross-Origin-Opener-Policy: same-origin`. It prevents Office from returning cross-domain dialog messages in Outlook on the web and new Outlook.

## Flow

1. `index.html` runs as the customer Outlook task pane.
2. It opens same-origin `dialog-redirect.html` through the Office Dialog API.
3. The redirect page navigates to Template Chooser on the officeatwork domain.
4. Template Chooser creates the document and sends it to the task pane with `Office.context.ui.messageParent`.
5. The task pane exposes the returned document as a download link.
