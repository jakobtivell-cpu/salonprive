# Personalbil – SharePoint setup

Two pieces: **the web part** (rendering engine, deployed once) and **personalbil-data.xlsx** (the data the client edits). The web part reads the Excel live on every page load. "Publishing a new promotion" = saving the Excel.

## 1. Scaffold the project (Node 22 LTS)

```bash
npm install -g @microsoft/generator-sharepoint@latest yo
mkdir kia-personalbil && cd kia-personalbil
yo @microsoft/sharepoint
#  Solution name:   kia-personalbil
#  Component type:  WebPart
#  Web part name:   Personalbil      <-- exact name; generates PersonalbilWebPart.ts
#  Framework:       No framework
```

Add the Excel reader dependency:

```bash
npm install xlsx@0.18.5
```

Replace the generated `src/webparts/personalbil/PersonalbilWebPart.ts` with the one provided.

## 2. Build the package

```bash
# SPFx 1.22 (Heft toolchain):
npm run build && npm run package-solution
# SPFx <= 1.21.1 (gulp):
# gulp bundle --ship && gulp package-solution --ship
```

The `.sppkg` lands in `sharepoint/solution/`.

## 3. Deploy

Upload the `.sppkg` to the tenant (or site) **App Catalog** → trust/deploy. The "Personalbil" web part now appears in the web part toolbox.

## 4. Upload the data file

Put `personalbil-data.xlsx` in a document library on the same site (e.g. **Shared Documents** / Delade dokument). Everyone who should see the calculator needs *read* access to it (default library permissions are fine).

To get the exact path: open the file's library, note the part after the site URL. The web part wants it **relative to the site**, e.g. `Shared Documents/personalbil-data.xlsx`.

## 5. Add + configure on the News post

Edit the News post → **+** → **Personalbil**. In the web part's edit pane (pencil) set:

- **Sökväg till datafilen** — e.g. `Shared Documents/personalbil-data.xlsx`
- **Rubrik** — e.g. `Personalbilskalkyl`
- **Hero-bild URL** *(optional)* — e.g. `https://egauximages.blob.core.windows.net/public/Appsida-bild-ev3.jpg`

Publish. Done.

## Editing afterwards (the client's workflow)

Open `personalbil-data.xlsx` in the library, edit cells, save. Changes show on next page load. No redeploy.

- **Modeller** — models + förmånsvärde/mån (what employees see)
- **Parametrar** — Biladministration, Brytpunkt statlig skatt, Standard hyresandel
- **Bilar** — fleet list (paste the dealer order export; clean `Status` column OR raw `Logistikstatus`/`Orderstatus` both work)
- **FAQ** — questions/answers (Svar accepts `<strong>` and `<a href="https://...">`)

## Notes

- Margin/profitability data is **not** in this file and never reaches employees — only förmånsvärde is exposed.
- No iframe, no external DB, no Power Automate required. Add a draft→live gate later only if the client wants approval before changes go live.
- Swedish characters in the file path are handled (`GetFileByServerRelativePath`). Keep the filename ASCII to be safe.
