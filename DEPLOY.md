# Deploying PPT Brand Converter

The Railway project and service already exist and are linked below — this
gets you live in about a minute. Everything else (Dockerfile, brand assets,
the full pipeline) is already in this folder, ready to go.

- Project: **ppt-brand-converter** (`6c6ab928-72ac-4c0d-803c-42a361cbeefa`)
- Service: **web** (`b9a874a4-be0e-458d-8948-9591e591418f`)

## Fastest path: Railway CLI, no GitHub needed

```bash
npm install -g @railway/cli
railway login                # opens a browser to authenticate — your existing Railway account
cd ppt-brand-converter/server
railway link -p 6c6ab928-72ac-4c0d-803c-42a361cbeefa -s web
railway variables --set ANTHROPIC_API_KEY=sk-ant-...   # your real key
railway up
```

`railway up` builds the Dockerfile (Node + Python/python-pptx + LibreOffice)
and deploys it to the `web` service directly — no GitHub repo required.

Once it finishes, get the public URL with:

```bash
railway domain
```

(or run it from the Railway dashboard → the `web` service → Settings →
Networking → **Generate Domain**, if you'd rather click than type).

## Alternative: deploy from a GitHub repo

If you'd rather have Railway auto-deploy on every push:

1. Push this folder to a new GitHub repo (e.g. `your-org/ppt-brand-converter`).
2. Tell me the repo name and I'll attach it to the existing `web` service
   (`connect-service-source`) and trigger the first build — no need to
   recreate the project.
3. Set `ANTHROPIC_API_KEY` the same way, via `railway variables --set` or
   the dashboard's Variables tab.

## After it's live

Open the app, upload any `.pptx`, click **Convert to Brand**. First run on
a large deck (100+ slides) will take a few minutes — most of that time is
the Claude planning calls (batches of 12 source slides each) plus the
LibreOffice visual QA pass.

## What to watch for on the first real conversions

This is v1. The component library and system prompt encode everything
proven out across this conversation, and the render/QA pipeline is tested
end-to-end — but the planning step (Claude choosing a layout per slide) is
now running unsupervised for the first time, where previously I was
reviewing every slide by eye before delivery. Worth spot-checking the
first few real conversions against decks you know well, the same way any
new automated step earns trust before you stop checking it.
