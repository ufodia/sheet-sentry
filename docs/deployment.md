# Deployment

Sheet Sentry is a static site: `npm run build` writes everything to `dist/`.

Production: **[sheet-sentry.vercel.app](https://sheet-sentry.vercel.app/)**. The production origin is `https://sheet-sentry.vercel.app/`.

## Build

```bash
npm ci
npm run check   # strict typecheck, unit tests and production build into dist/
```

`postbuild` runs `scripts/prepare-hosting.mjs`, which only touches `dist/`. On a Vercel production build (`VERCEL_ENV=production` with `VERCEL_PROJECT_PRODUCTION_URL`) it adds a canonical URL, makes the social image absolute and writes `robots.txt` and `sitemap.xml`. Other builds are marked `noindex`.

## Security headers

| Header | Value |
| --- | --- |
| `Content-Security-Policy` | `default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self'; connect-src 'none'; worker-src 'self' blob:; object-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'` |
| `X-Content-Type-Options` | `nosniff` |
| `Referrer-Policy` | `no-referrer` |
| `Permissions-Policy` | `camera=(), microphone=(), geolocation=(), payment=(), usb=()` |
| `Cross-Origin-Opener-Policy` | `same-origin` |

## Vercel

[`vercel.json`](../vercel.json) sets the framework, `npm ci`, `npm run build`, the `dist` output directory and the headers above. Node.js 24 comes from `engines`.

To recover hosting, restore the source from `https://github.com/ufodia/sheet-sentry.git` on `main`, install Node.js 24, run `npm ci` and `npm run check`, and reconnect the authorized Vercel account to this repository. Select `main` as the production branch and retain the repository's build and header settings. Reassign `sheet-sentry.vercel.app` to the restored project and verify that `VERCEL_PROJECT_PRODUCTION_URL` resolves to that production hostname. Deploy the checked commit, then confirm its source SHA, canonical URL, security headers and synthetic demo at the public address. Keep account tokens and `.vercel/` metadata outside Git.

## Other hosts

Publish `dist/` and add the headers. For hosts that read `_headers`, create `public/_headers` before building:

```
/*
  Content-Security-Policy: default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self'; connect-src 'none'; worker-src 'self' blob:; object-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'
  X-Content-Type-Options: nosniff
  Referrer-Policy: no-referrer
  Permissions-Policy: camera=(), microphone=(), geolocation=(), payment=(), usb=()
  Cross-Origin-Opener-Policy: same-origin
```

## Verify a deployment

1. Check the response headers of the page.
2. Run the demo; the Network panel should show only same-origin files.
3. `PLAYWRIGHT_BASE_URL=https://your-deployment.example npx playwright test`
