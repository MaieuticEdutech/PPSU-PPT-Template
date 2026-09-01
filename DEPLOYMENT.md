# Deploying PPSU PPT Designer — Render (backend) + Vercel (frontend)

Everything code-side is already prepared in this repo:

- `ppsu1/Dockerfile` — Python 3.11 + LibreOffice + Calibri-metric fonts; runs
  gunicorn (1 worker, 8 threads, 15-min request timeout).
- `render.yaml` — Render Blueprint: free Docker web service named
  `ppsu-ppt-designer`, built from `ppsu1/`.
- `ppsu1/frontend/index.html` — auto-detects where it is served from:
  on Vercel it calls `HOSTED_BACKEND` (https://ppsu-ppt-designer.onrender.com),
  on localhost / the LAN it keeps calling the local backend as before.
- The pre-computed caches (template catalog, 302 design thumbnails, icon PNGs)
  are **committed** — the 512 MB free instance can never rebuild them
  (the catalog build alone peaks at ~840 MB of RAM).
- `/result/<token>` endpoint + frontend fallback: if the slow `/build` request
  is dropped by a proxy, the page polls `/progress` and picks the finished
  result up afterwards.

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
2. **New + → Blueprint** → select this repo → Render reads `render.yaml`.
3. When prompted for `FRONTEND_URL`, leave it empty for now (empty = allow all
   origins) — you'll set it after step 3.
4. Apply. First build takes ~10 min (the image installs LibreOffice).
5. Note the service URL. If `ppsu-ppt-designer` was taken and Render gave a
   different URL, edit the `HOSTED_BACKEND` line near the top of the
   `<script>` block in `ppsu1/frontend/index.html`, commit, push.
6. Check it: open `https://<service>.onrender.com/` — should answer
   `{"service": "ppt-designer", "status": "ok"}`.

### 3. Frontend on Vercel
1. Sign up / log in at https://vercel.com with the GitHub account.
2. **Add New → Project** → import this repo.
3. Set **Root Directory** to `ppsu1/frontend`. Framework preset **Other**,
   no build command, no output directory changes.
4. Deploy, note the production URL (e.g. `https://<project>.vercel.app`).

### 4. Lock CORS (recommended)
In the Render dashboard → the service → **Environment** → set
`FRONTEND_URL` = the Vercel production URL (no trailing slash) → save
(the service redeploys). Until you do this the API accepts any origin.

### 5. Smoke test
Open the Vercel URL, upload `PPSU RAW PPT/RAW PPT1.pptx`, Generate,
wait, preview, download.

## Free-tier realities (measured)

| Fact | Consequence |
|---|---|
| 512 MB RAM, 0.1 vCPU | A build that takes ~40 s locally will take **several minutes**; the warm build peaks at ~460 MB — it fits, but with little headroom. |
| Spins down after 15 min idle | First request after a pause takes ~1 min to answer. |
| No persistent disk | Sessions (uploads/previews) vanish on restart — users just re-upload. The bundled template's caches are safe because they ship in the image. |
| Cold catalog build peaks ~840 MB | Only happens if a user **uploads their own large template** — that request will be killed (OOM) on the free tier. Small uploaded templates are fine (~seconds). |

If the 50 MB default template ever OOMs in practice, set the
`DEFAULT_TEMPLATE_PATH` env var on Render to
`/app/template/PPSU TEMPLATE.pptx` (the 0.46 MB, 22-design template) —
no code change needed.
