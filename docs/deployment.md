# Deployment

Sheet Sentry is a static site: `npm run build` writes everything to `dist/`.

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
