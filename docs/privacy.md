# Privacy

Budgets, forecasts, salaries and board packs are some of the most sensitive files a company has. Sheet Sentry never sends them anywhere.

| Behaviour | How it is implemented |
| --- | --- |
| No upload | Workbooks are read with the browser's `File` API and unzipped and compared in the tab. There is no server endpoint. The production configuration sends `Content-Security-Policy: connect-src 'none'`, which blocks `fetch`, XHR, WebSocket, EventSource and beacons. |
| No third-party code | `default-src 'self'`; the only runtime dependency, fflate, is bundled into the app. No analytics, fonts or CDNs. |
| No storage | Nothing is written to cookies, `localStorage`, `sessionStorage` or IndexedDB. Reloading the tab discards everything. |
| Nothing executes | Macros are detected by the presence of `xl/vbaProject.bin` and never run. Formulas are read as text and never evaluated. Cell contents are inserted as text nodes, never as HTML. |
| Local exports | Markdown, CSV and JSON files are created in the tab and saved through the browser's downloads. |
| Command line | Reads the two files you name and writes only the `--json`/`--csv` files you ask for. No network access. |

Loading the application contacts the host for its own assets. External GitHub and Buy Me a Coffee links open their respective services when you choose them. `connect-src` blocks scripted data connections; it does not prevent ordinary navigation to an external link. README badges are GitHub documentation assets, not dependencies of the running application.

## What remains your responsibility

- **Exports** contain sheet names, cell addresses, formulas and values from your workbooks. Share them as you would the workbook.
- **Your device** — browser extensions with page access, screen recording and security software that injects scripts into pages are outside the app's control.
- **Hosting** — keep the headers from [`vercel.json`](../vercel.json) when self-hosting (see [deployment](deployment.md)). The Vite development server does not send them; the app still never sends workbook data anywhere.

Report a way to make the page send data, or to execute content from a workbook, privately as described in [SECURITY.md](../SECURITY.md).
