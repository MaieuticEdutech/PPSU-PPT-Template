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
- `render.yaml` sets `DISABLE_PREVIEW_RENDER=1`: on the free instance, live
  slide previews and "Change design" thumbnails are off. Not a bug — see the
  table below. Downloads work fully; the frontend already shows a dedicated
  "Visual preview isn't available on this server" panel for this state.

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
| **LibreOffice itself doesn't fit alongside the app** | **Confirmed 2026-09-02**, by isolating each cost: `engine.build()` alone (python-pptx, no rendering) peaked at ~61 MB even for a 23-slide deck — nowhere near the ceiling. Yet a single, solo `/build` request (no concurrency, small template) still silently killed the container within ~2 minutes of calling the renderer — Render's log shows gunicorn's master restarting with zero warning, no traceback, no shutdown message: a kernel OOM-kill signature, not a Python error. A single headless LibreOffice conversion carries large, mostly fixed memory overhead of its own that this instance cannot absorb, regardless of template size or how many renders run at once. `DISABLE_PREVIEW_RENDER=1` skips rendering entirely: `/build` still produces a real, downloadable `.pptx`; only the PNG previews and "Change design" thumbnails are unavailable, a state the app already handles gracefully. Remove the flag once the instance is upgraded off the free plan. |
| The 50 MB / 302-slide default template also doesn't fit | Separately from the above: its COLD catalog build alone peaks at ~840 MB (confirmed 2026-09-01). `render.yaml` defaults to the small 0.46 MB / 22-design template (`DEFAULT_TEMPLATE_PATH`) instead — noticeably less design variety, but it fits comfortably. |
| `/build` may still show a proxy error in the browser console on a genuinely slow build | Render's edge times out around 100s regardless of app config; the frontend recovers via `/progress` + `/result` instead of surfacing it, but the console entry is cosmetic noise, not a bug to chase. |
| Concurrent LibreOffice launches compound the memory problem | Not the root cause (a *single* build already doesn't fit), but still fixed as defense in depth: `render.py`'s `_LO_LOCK` serializes every render so a second `/build`, the background thumbnail warm-up, or a second user can never launch `soffice` twice at once. |
| Spins down after 15 min idle | First request after a pause takes ~1 min to answer. |
| No persistent disk | Sessions (uploads/previews) vanish on restart — users just re-upload. The bundled template's caches are safe because they ship in the image. |
| Cold catalog build peaks ~840 MB | Also happens if a user **uploads their own large template** — that request will be killed (OOM) on the free tier. Small uploaded templates are fine (~seconds). |
