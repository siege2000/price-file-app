# Publishing a GitHub Release

## One-time Setup — Authenticate `gh` CLI

### Step 1 — Create a Personal Access Token

1. Go to GitHub → **Settings** → **Developer settings** → **Personal access tokens** → **Fine-grained tokens**
2. Click **Generate new token**
3. Set:
   - **Token name:** e.g. `gh-cli-price-file-app`
   - **Expiration:** your preference (90 days is a good default)
   - **Repository access:** Only select repositories → `price-file-app`
   - **Permissions → Repository permissions → Contents:** `Read and write`
4. Click **Generate token** and copy it immediately (it won't be shown again)

### Step 2 — Log in with the token

Open a terminal in the project folder and run:

```bash
gh auth login --with-token
```

Paste your token when prompted. Then verify:

```bash
gh auth status
```

---

## Building a New Release

### 1 — Update the version number

Edit `price_file_app.spec` and `version_info.txt` — change the version following the pattern `YYYY.MM.N`:
- `YYYY` = year
- `MM` = month (zero-padded)
- `N` = release number for that month (1, 2, 3 …)

Example: `2026.04.02` = second release in April 2026.

In **`price_file_app.spec`**:
```python
APP_VERSION = '2026.04.02'
```

In **`version_info.txt`** update both `filevers`/`prodvers` tuples and all four `StringStruct` version fields:
```python
filevers=(2026, 4, 2, 0),
prodvers=(2026, 4, 2, 0),
...
StringStruct(u'FileVersion',    u'2026.04.02'),
StringStruct(u'ProductVersion', u'2026.04.02'),
```

### 2 — Rebuild the exe

```bash
venv\Scripts\activate
pyinstaller -y price_file_app.spec
```

Output: `dist\PriceFileApp\PriceFileApp.exe`

### 3 — Zip the output

```powershell
Compress-Archive -Path dist\PriceFileApp\* -DestinationPath dist\PriceFileApp-2026.04.02.zip -Force
```

### 4 — Commit and push

```bash
git add price_file_app.spec version_info.txt
git commit -m "Bump version to 2026.04.02"
git push
```

### 5 — Create the GitHub release

```bash
gh release create v2026.04.02 dist/PriceFileApp-2026.04.02.zip \
  --title "Price File Builder 2026.04.02" \
  --notes "Release notes here." \
  --draft
```

Remove `--draft` to publish immediately, or go to GitHub → Releases → edit the draft and publish manually.

---

## Notes

- The `dist/` and `build/` folders are in `.gitignore` — the zip is never committed to git, only uploaded to the release.
- `gh` CLI docs: https://cli.github.com/manual/gh_release_create
