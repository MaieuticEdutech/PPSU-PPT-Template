# Deploying PPSU PPT Designer — Render (backend) + Vercel (frontend)

Everything code-side is already prepared in this repo:

- `ppsu1/Dockerfile` — Python 3.11 + LibreOffice + Calibri-metric fonts; runs
  gunicorn (1 worker, `gthread` class, 8 threads, 15-min request timeout).
  Threads matter: with the default `sync` worker class gunicorn silently
  ignores `--threads`, so ONE slow request blocks every other endpoint
  (including the health check) until Render's edge proxy gives up — that
  looks like the whole service is down, not just a slow response.
- `render.yaml` — Render Blueprint: free Docker web service named
  `ppsu-ppt-designer`, built from `ppsu1/`. Sets `FRONTEND_URL` (locks CORS
  to the Vercel origin) and `DEFAULT_TEMPLATE_PATH` (see the OOM row below)
  as real values, so both apply automatically — no dashboard step needed.
- `ppsu1/frontend/index.html` — auto-detects where it is served from:
  on Vercel it calls `HOSTED_BACKEND` (https://ppsu-ppt-designer.onrender.com),
  on localhost / the LAN it keeps calling the local backend as before. A
  502/503/504 (the hosting proxy giving up on a slow backend) or a dropped
  connection both fall back to polling `/progress` + `/result` instead of
  showing the user a raw error — a build that outlives Render's ~100s edge
  timeout still finishes and still reaches the browser.
- The pre-computed caches (template catalog, 302 design thumbnails, icon PNGs)
  are **committed** — the 512 MB free instance can never rebuild them
  (the catalog build alone peaks at ~840 MB of RAM).
- `/result/<token>` endpoint: the finished payload for a session, so the page
  can recover a build the proxy dropped instead of re-uploading.

## Manual steps (one-time, ~15 minutes)

### 1. Push the repo to GitHub
```
cd C:\Users\GA01\Desktop\Template_Designer-main
git remote add origin https://github.com/<your-account>/<repo-name>.git
git push -u origin main
```
(Create the empty repo on github.com first — private is fine, both Render and
Vercel can read private repos once you connect your GitHub account.)

### 2. Backend on Render
1. Sign up / log in at https://render.com with the GitHub account.
2. **New + → Blueprint** → select this repo → Render reads `render.yaml` and
   pre-fills `FRONTEND_URL` and `DEFAULT_TEMPLATE_PATH` from it — nothing to
   type.
3. Apply. First build takes ~10 min (the image installs LibreOffice).
4. Note the service URL. If `ppsu-ppt-designer` was taken and Render gave a
   different URL, edit the `HOSTED_BACKEND` line near the top of the
   `<script>` block in `ppsu1/frontend/index.html`, commit, push.
5. Check it: open `https://<service>.onrender.com/` — should answer
   `{"service": "ppt-designer", "status": "ok"}`.

### 3. Frontend on Vercel
1. Sign up / log in at https://vercel.com with the GitHub account.
2. **Add New → Project** → import this repo.
3. Set **Root Directory** to `ppsu1/frontend`. Framework preset **Other**,
   no build command, no output directory changes.
4. Deploy, note the production URL (e.g. `https://<project>.vercel.app`).
5. If it doesn't match the `FRONTEND_URL` already set in `render.yaml`
   (`https://ppsu-ppt-template.vercel.app`), update that value, commit, push.

### 4. Smoke test
Open the Vercel URL, upload `PPSU RAW PPT/RAW PPT1.pptx`, Generate,
wait, preview, download.

## Free-tier realities (measured)

| Fact | Consequence |
|---|---|
| 512 MB RAM, 0.1 vCPU | A build that takes ~40 s locally takes much longer here. |
| The 50 MB / 302-slide default template OOMs the instance | **Confirmed 2026-09-01**: parsing that much XML plus spawning LibreOffice to render previews exceeds 512 MB — the build died ~19s in, right as LibreOffice launched, and the service briefly 503'd while Render restarted the container. `render.yaml` now defaults to the small 0.46 MB / 22-design template (`DEFAULT_TEMPLATE_PATH`) instead — noticeably less design variety, but it fits. Restoring the full 285-design library needs either an instance upgrade or caching the parsed template across requests instead of re-parsing it on every build (not yet done). |
| `/build` still shows a proxy error in the browser console on a genuinely slow build | Render's edge times out around 100s regardless of app config; the frontend now recovers via `/progress` + `/result` instead of surfacing it, but the console entry is cosmetic noise, not a bug to chase. |
| Spins down after 15 min idle | First request after a pause takes ~1 min to answer. |
| No persistent disk | Sessions (uploads/previews) vanish on restart — users just re-upload. The bundled template's caches are safe because they ship in the image. |
| Cold catalog build peaks ~840 MB | Only happens if a user **uploads their own large template** — that request will be killed (OOM) on the free tier. Small uploaded templates are fine (~seconds). |
