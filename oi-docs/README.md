# OI Document Generator

Paste any supplier quote → get a Purchase Order + Tax Invoice as Word documents.

## Deploy to Vercel in 3 steps

### 1. Put this folder on GitHub

- Go to [github.com/new](https://github.com/new)
- Create a new **private** repository called `oi-doc-generator`
- Upload these files (drag the whole folder in, or use git):
  ```
  oi-docs/
  ├── api/
  │   ├── parse.js
  │   └── generate.js
  ├── public/
  │   └── index.html
  ├── package.json
  ├── vercel.json
  └── README.md
  ```

### 2. Deploy on Vercel

- Go to [vercel.com](https://vercel.com) → **Add New Project**
- Import your `oi-doc-generator` GitHub repo
- Click **Deploy** (no build settings needed)

### 3. Add your API key

- In Vercel dashboard → your project → **Settings → Environment Variables**
- Add: `ANTHROPIC_API_KEY` = `sk-ant-...your key...`
- Go to **Deployments** → click the three dots on your latest deploy → **Redeploy**

That's it. Your tool is live at `https://oi-doc-generator.vercel.app`

---

## How to use

1. Open your Vercel URL
2. Paste any quote text (from PDF, email, or just describe the deal)
3. Click **Generate Documents**
4. Download your Purchase Order and Invoice as `.docx` Word files

## Local development

```bash
npm install -g vercel
npm install
vercel dev
```

Then open http://localhost:3000
