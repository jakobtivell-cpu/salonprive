import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as XLSX from 'xlsx';

export interface IPersonalbilWebPartProps {
  dataFilePath: string;   // server-relative-to-site, e.g. "Shared Documents/personalbil-data.xlsx"
  heroImageUrl: string;
  heroTitle: string;
}

interface IModel { id: string; nm: string; dr: string; rn: string; fv: number; img: string; }
interface ICar { reg: string; desc: string; ext: string; int: string; status: string; }
interface IFaq { q: string; a: string; }
interface IData { models: IModel[]; cars: ICar[]; faq: IFaq[]; bilAdm: number; bp: number; defRp: number; }

const DEFAULT_FAQ: IFaq[] = [
  { q: 'Vad ingår i min månadskostnad?', a: 'Försäkring, däck och däckförvaring, service samt administration ingår. Drivmedel/laddning och eventuella tillval tillkommer.' },
  { q: 'Hur fungerar förmånsbeskattningen?', a: 'Förmånsvärdet delas enligt företagets policy: en del betalas via löneavdrag, en del beskattas som förmån. Du ser din exakta månadskostnad i kalkylen.' }
];

const SWATCH: { [k: string]: string } = {
  'shale grey': '#6f7377', 'snow white pearl': '#f3f3f0', 'aventurine green': '#3b4a3f',
  'wolf gray m': '#7a7d80', 'deluxe white pearl m': '#f3f3f0', 'morning haze s': '#cfd2d4',
  'penta metal m': '#9aa0a4', 'dark grey': '#3a3d40', 'medium grey': '#7a7d80'
};

export default class PersonalbilWebPart extends BaseClientSideWebPart<IPersonalbilWebPartProps> {
  private _data: IData | null = null;
  private _err = '';
  private _tab: 'emp' | 'cars' = 'emp';
  private _ov = false; // user overrode statlig skatt

  public render(): void {
    if (!this._data && !this._err) {
      this.domElement.innerHTML = `<div class="egx-pbk"><div class="egx-load">Laddar kalkyl…</div></div>`;
      this._load().then(() => this.render()).catch(() => this.render());
      return;
    }
    if (this._err) {
      this.domElement.innerHTML =
        `<div class="egx-pbk"><div class="egx-err"><strong>Kunde inte ladda datafilen.</strong><br>${this._esc(this._err)}` +
        `<br><span class="egx-err-hint">Kontrollera sökvägen i webbdelens inställningar (för närvarande: ${this._esc(this.properties.dataFilePath || '–')}).</span></div></div>`;
      return;
    }
    this._build();
  }

  /* ---------------- data load ---------------- */
  private async _load(): Promise<void> {
    try {
      const site = this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
      const rel = (this.properties.dataFilePath || 'Shared Documents/personalbil-data.xlsx').replace(/^\//, '');
      const path = `${site}/${rel}`;
      const api = `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/GetFileByServerRelativePath(decodedUrl='${path.replace(/'/g, "''")}')/$value`;
      const resp: SPHttpClientResponse = await this.context.spHttpClient.get(api, SPHttpClient.configurations.v1);
      if (!resp.ok) { this._err = `HTTP ${resp.status} – filen hittades inte eller saknar behörighet.`; return; }
      const buf = await resp.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
      this._data = this._parse(wb);
    } catch (e) {
      this._err = (e && (e as Error).message) || 'Okänt fel.';
    }
  }

  private _parse(wb: XLSX.WorkBook): IData {
    const sheet = (name: string): any[] => {
      const ws = wb.Sheets[name];
      return ws ? XLSX.utils.sheet_to_json(ws, { defval: '' }) : [];
    };
    const get = (row: any, ...names: string[]): string => {
      for (const k in row) {
        const kk = String(k).trim().toLowerCase();
        if (names.some(n => n.toLowerCase() === kk)) return String(row[k]);
      }
      return '';
    };

    // Modeller
    const models: IModel[] = sheet('Modeller').map(r => ({
      id: get(r, 'id').trim() || ('m' + Math.random().toString(36).slice(2, 7)),
      nm: get(r, 'Modell', 'Namn').trim(),
      dr: get(r, 'Drivlinje').trim(),
      rn: get(r, 'Räckvidd').trim(),
      fv: parseFloat(get(r, 'Förmånsvärde').replace(/\s/g, '').replace(',', '.')) || 0,
      img: get(r, 'Bild-URL', 'Bild').trim()
    })).filter(m => m.nm);

    // Parametrar
    let bilAdm = 275, bp = 53591, defRp = 50;
    sheet('Parametrar').forEach(r => {
      const key = get(r, 'Parameter').toLowerCase();
      const val = parseFloat(get(r, 'Värde', 'Value').replace(/\s/g, '').replace(',', '.')) || 0;
      if (key.indexOf('biladmin') >= 0) bilAdm = val;
      else if (key.indexOf('brytpunkt') >= 0) bp = val;
      else if (key.indexOf('hyresandel') >= 0) defRp = val;
    });

    // Bilar (clean list OR raw order export)
    const cars: ICar[] = sheet('Bilar').map(r => {
      let status = get(r, 'Status').trim();
      if (!status) status = this._statusFrom(get(r, 'Logistikstatus'), get(r, 'Orderstatus'));
      return {
        reg: get(r, 'Reg. nr', 'Regnr', 'Reg.nr').trim(),
        desc: get(r, 'Beskrivning').trim(),
        ext: this._cleanColor(get(r, 'Ext')),
        int: this._cleanColor(get(r, 'Int')),
        status
      };
    }).filter(c => c.desc);

    // FAQ
    let faq: IFaq[] = sheet('FAQ').map(r => ({ q: get(r, 'Fråga', 'Question').trim(), a: get(r, 'Svar', 'Answer').trim() })).filter(f => f.q);
    if (!faq.length) faq = DEFAULT_FAQ;

    return { models, cars, faq, bilAdm, bp, defRp };
  }

  private _statusFrom(log: string, order: string): string {
    const l = (log || '').toLowerCase();
    if (l.indexOf('compound') >= 0) return 'På plats';
    if (l.indexOf('shipping') >= 0) return 'På väg';
    if (l.indexOf('production') >= 0) return 'I produktion';
    const o = (order || '').trim();
    return o ? o.charAt(0).toUpperCase() + o.slice(1).toLowerCase() : 'Okänd';
  }
  private _cleanColor(s: string): string {
    if (!s) return '';
    s = String(s).replace(/^DFS:\s*/i, '').replace(/\([^)]*\)\s*/g, '').trim();
    return s.toLowerCase().replace(/\b\w/g, ch => ch.toUpperCase());
  }

  /* ---------------- UI ---------------- */
  private _build(): void {
    const d = this._data!;
    const opts = d.models.map(m => `<option value="${m.id}">${this._esc(m.nm)}</option>`).join('');
    const hero = this.properties.heroImageUrl
      ? `<div class="egx-hero"><img src="${this._esc(this.properties.heroImageUrl)}" alt=""><div class="egx-hero-ov"></div>
           <div class="egx-hero-tx"><h1>${this._esc(this.properties.heroTitle || 'Personalbilskalkyl')}</h1>
           <p>Beräkna din månadskostnad. Välj modell och se vad förmånen kostar dig.</p></div></div>`
      : `<h1 class="egx-h1">${this._esc(this.properties.heroTitle || 'Personalbilskalkyl')}</h1>`;

    this.domElement.innerHTML = `
<div class="egx-pbk">
  ${this._css()}
  ${hero}
  <div class="egx-subnav">
    <a data-tab="emp" class="active">Medarbetarens vy</a>
    <a data-tab="cars">Tillgängliga bilar</a>
  </div>

  <div data-panel="emp" class="egx-panel active">
    <div class="egx-emp-top">
      <div class="egx-form"><div class="egx-card"><div class="egx-grid">
        <div class="egx-g"><label>Månadslön brutto</label><input type="number" data-f="sa" value="45000"><span class="egx-u">kr</span></div>
        <div class="egx-g"><label>Kommunalskatt</label><input type="number" data-f="tx" value="32.50" step="0.01"><span class="egx-u">%</span></div>
        <div class="egx-g"><label>Statlig skatt</label><select data-f="st"><option value="0">Nej</option><option value="1">Ja (+20%)</option></select></div>
        <div class="egx-g"><label>Avdragstyp</label><select data-f="at"><option value="netto" selected>Nettolöneavdrag</option><option value="brutto">Bruttolöneavdrag</option></select></div>
        <div class="egx-g"><label>Hyresandel</label><input type="number" data-f="rp" value="${d.defRp}" step="5"><span class="egx-u">% av förmånsvärde</span></div>
        <div class="egx-g"><label>Välj modell</label><select data-f="model">${opts}</select></div>
      </div></div></div>
      <div class="egx-img"><img data-el="ci" src="" alt=""></div>
    </div>
    <div data-el="er"></div>
    <div class="egx-ft">Förmånsvärden enligt Skatteverkets fastställda värden. Verifiera alltid med Skatteverkets e-tjänst.</div>
  </div>

  <div data-panel="cars" class="egx-panel">
    <p class="egx-hint">Aktuell fordonslista för personalbilsprogrammet. Boka önskad bil via HR/programansvarig.</p>
    <div class="egx-cars-tb">
      <div class="egx-g"><label>Sök reg.nr / modell</label><input data-c="q" type="text" placeholder="t.ex. EV3 eller ABC123"></div>
      <div class="egx-g"><label>Status</label><select data-c="status"><option value="">Alla</option><option>På plats</option><option>På väg</option><option>I produktion</option></select></div>
      <span class="egx-count" data-el="count"></span>
    </div>
    <div class="egx-card egx-scroll"><table class="egx-cars">
      <thead><tr><th>Reg. nr</th><th>Beskrivning</th><th>Ext</th><th>Int</th><th>Status</th></tr></thead>
      <tbody data-el="cars"></tbody>
    </table></div>
  </div>

  <div class="egx-faqwrap">
    <div class="egx-faq-lbl">Vanliga frågor</div>
    <h2 class="egx-faq-h">Allt du behöver veta om din personalbil</h2>
    <div class="egx-faq">${this._faqHtml(d.faq)}</div>
  </div>
</div>`;

    this._wire();
    this._sync();
    this._recalc();
    this._renderCars();
  }

  private _wire(): void {
    const root = this.domElement;
    root.querySelectorAll('.egx-subnav a').forEach(a =>
      a.addEventListener('click', () => {
        this._tab = (a.getAttribute('data-tab') as 'emp' | 'cars');
        root.querySelectorAll('.egx-subnav a').forEach(x => x.classList.remove('active'));
        a.classList.add('active');
        root.querySelectorAll('.egx-panel').forEach(p =>
          p.classList.toggle('active', p.getAttribute('data-panel') === this._tab));
      }));

    const on = (sel: string, ev: string, cb: () => void) => {
      const el = root.querySelector(sel); if (el) el.addEventListener(ev, cb);
    };
    on('[data-f="sa"]', 'input', () => { this._sync(); this._recalc(); });
    on('[data-f="tx"]', 'input', () => this._recalc());
    on('[data-f="st"]', 'change', () => { this._ov = true; this._recalc(); });
    on('[data-f="at"]', 'change', () => this._recalc());
    on('[data-f="rp"]', 'input', () => this._recalc());
    on('[data-f="model"]', 'change', () => this._recalc());
    on('[data-c="q"]', 'input', () => this._renderCars());
    on('[data-c="status"]', 'change', () => this._renderCars());

    root.querySelectorAll('.egx-fq').forEach(b =>
      b.addEventListener('click', () => b.parentElement!.classList.toggle('open')));
  }

  private _sync(): void {
    if (this._ov) return;
    const sel = this.domElement.querySelector('[data-f="st"]') as HTMLSelectElement;
    if (sel) sel.value = this._val('sa') > this._data!.bp ? '1' : '0';
  }

  private _recalc(): void {
    const d = this._data!;
    const m = d.models.filter(x => x.id === this._str('model'))[0] || d.models[0];
    if (!m) return;
    const fv = m.fv || 0, fY = fv * 12;
    const rp = this._val('rp'), hy = (rp / 100) * fv;
    const sa = this._val('sa'), ct = this._val('tx') / 100;
    const st = (parseInt(this._str('st'), 10) || 0) * 0.2, mt = ct + st;
    const at = this._str('at'), bilAdm = d.bilAdm;

    let fvT: number, fs: number, hn: number, tot: number;
    if (at === 'netto') { fvT = Math.max(0, fv - hy); fs = fvT * mt; hn = hy; tot = hn + fs + bilAdm; }
    else { fvT = fv; fs = fv * mt; hn = hy * (1 - mt); tot = hn + fs + bilAdm; }

    const img = this.domElement.querySelector('[data-el="ci"]') as HTMLImageElement;
    img.src = m.img || ''; img.alt = m.nm;

    const policyNote = (at === 'netto' && rp === 50)
      ? `<tr><td colspan="2" class="egx-policy">Enligt företagets policy fördelas förmånsvärdet 50/50: hälften betalas som nettolöneavdrag, hälften beskattas som förmån.</td></tr>` : '';

    let rows: string;
    if (at === 'netto') {
      rows =
        `<tr class="egx-sec"><td colspan="2">Nettolöneavdrag</td></tr>${policyNote}` +
        `<tr><td>Marginalskatt</td><td>${(mt * 100).toFixed(1)}%</td></tr>` +
        `<tr><td>Nettolöneavdrag (${rp}% av förmånsvärde)</td><td>${this._kr(hy)}</td></tr>` +
        `<tr><td>Reducerat förmånsvärde</td><td>${this._kr(fvT)}</td></tr>` +
        `<tr><td>Förmånsskatt</td><td>${this._kr(fs)}</td></tr>` +
        `<tr><td>Biladministration <span class="egx-sm">(inkl. moms)</span></td><td>${this._kr(bilAdm)}</td></tr>` +
        `<tr class="egx-tot"><td>Total månadskostnad (netto)</td><td class="egx-green">${this._kr(tot)}</td></tr>`;
    } else {
      rows =
        `<tr class="egx-sec"><td colspan="2">Bruttolöneavdrag</td></tr>` +
        `<tr><td>Marginalskatt</td><td>${(mt * 100).toFixed(1)}%</td></tr>` +
        `<tr><td>Bruttolöneavdrag (${rp}% av förmånsvärde)</td><td>${this._kr(hy)}</td></tr>` +
        `<tr><td>Skatteeffekt på avdraget</td><td class="egx-muted">−${this._kr(hy * mt)}</td></tr>` +
        `<tr><td>Netto-effekt av hyra</td><td>${this._kr(hn)}</td></tr>` +
        `<tr><td>Förmånsskatt</td><td>${this._kr(fs)}</td></tr>` +
        `<tr><td>Biladministration <span class="egx-sm">(inkl. moms)</span></td><td>${this._kr(bilAdm)}</td></tr>` +
        `<tr class="egx-tot"><td>Total månadskostnad (netto)</td><td class="egx-green">${this._kr(tot)}</td></tr>`;
    }

    const lon = at === 'brutto'
      ? `<tr><td>Bruttolön före</td><td>${this._kr(sa)}</td></tr>` +
        `<tr><td>Bruttolön efter avdrag (hyra)</td><td>${this._kr(sa - hy)}</td></tr>` +
        `<tr><td>Beskattningsbar inkomst (lön + förmånsvärde)</td><td>${this._kr(sa - hy + fv)}</td></tr>`
      : `<tr><td>Bruttolön (oförändrad)</td><td>${this._kr(sa)}</td></tr>` +
        `<tr><td>Beskattningsbar inkomst (lön + reducerat förmånsvärde)</td><td>${this._kr(sa + fvT)}</td></tr>`;

    (this.domElement.querySelector('[data-el="er"]') as HTMLElement).innerHTML =
      `<div class="egx-card egx-scroll"><div class="egx-head"><h3>${this._esc(m.nm)}</h3>` +
      `<span class="egx-badge">${this._esc(m.dr)}${m.rn ? ' · ' + this._esc(m.rn) : ''}</span></div>` +
      `<table class="egx-bd"><tbody>` +
      `<tr class="egx-sec"><td colspan="2">Förmånsvärde</td></tr>` +
      `<tr class="egx-hl"><td>Förmånsvärde per månad</td><td class="egx-green">${this._kr(fv)}</td></tr>` +
      `<tr><td>Förmånsvärde per år</td><td>${this._kr(fY)}</td></tr>` +
      rows +
      `<tr class="egx-sec"><td colspan="2">Löneeffekt</td></tr>${lon}` +
      `</tbody></table></div>`;
  }

  private _renderCars(): void {
    const d = this._data!;
    const q = ((this.domElement.querySelector('[data-c="q"]') as HTMLInputElement)?.value || '').toLowerCase();
    const sf = (this.domElement.querySelector('[data-c="status"]') as HTMLSelectElement)?.value || '';
    const rows = d.cars.filter(c => {
      if (sf && c.status !== sf) return false;
      if (q && (c.reg + ' ' + c.desc + ' ' + c.ext + ' ' + c.int).toLowerCase().indexOf(q) === -1) return false;
      return true;
    });
    const pill = (s: string) => s === 'På plats' ? 'egx-pill-g' : (s === 'I produktion' ? 'egx-pill-x' : 'egx-pill-a');
    const sw = (n: string) => { const c = SWATCH[(n || '').toLowerCase()]; return c ? `<span class="egx-sw" style="background:${c}"></span>` : ''; };
    let h = rows.map(c =>
      `<tr><td class="egx-reg${c.reg ? '' : ' none'}">${this._esc(c.reg) || '–'}</td>` +
      `<td>${this._esc(c.desc)}</td><td>${sw(c.ext)}${this._esc(c.ext)}</td>` +
      `<td>${sw(c.int)}${this._esc(c.int)}</td>` +
      `<td><span class="egx-pill ${pill(c.status)}">${this._esc(c.status)}</span></td></tr>`).join('');
    if (!rows.length) h = `<tr><td colspan="5" class="egx-empty">Inga bilar matchar filtret.</td></tr>`;
    (this.domElement.querySelector('[data-el="cars"]') as HTMLElement).innerHTML = h;
    const avail = rows.filter(c => c.status === 'På plats').length;
    (this.domElement.querySelector('[data-el="count"]') as HTMLElement).textContent = `${rows.length} bilar · ${avail} på plats`;
  }

  private _faqHtml(faq: IFaq[]): string {
    const chev = '<svg class="egx-fic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 9l6 6 6-6"/></svg>';
    return faq.map((f, i) =>
      `<div class="egx-fi"><button class="egx-fq"><span><span class="egx-n">${String(i + 1).padStart(2, '0')}</span>${this._esc(f.q)}</span>${chev}</button>` +
      `<div class="egx-fa"><div><p>${this._sanitize(f.a)}</p></div></div></div>`).join('');
  }

  /* ---------------- utils ---------------- */
  private _val(f: string): number { const el = this.domElement.querySelector(`[data-f="${f}"]`) as HTMLInputElement; return parseFloat(el && el.value) || 0; }
  private _str(f: string): string { const el = this.domElement.querySelector(`[data-f="${f}"]`) as HTMLInputElement; return (el && el.value) || ''; }
  private _kr(x: number): string { return Math.round(x).toLocaleString('sv-SE') + ' kr'; }
  private _esc(s: string): string { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  // allow only <strong>/<b>/<em>/<a href> from FAQ answers; strip everything else (incl. scripts)
  private _sanitize(s: string): string {
    let out = String(s || '').replace(/<(?!\/?(strong|b|em|a)\b)[^>]*>/gi, '');
    out = out.replace(/<a\b([^>]*)>/gi, (m, attrs) => {
      const href = (attrs.match(/href\s*=\s*"([^"]*)"/i) || [])[1] || '#';
      const safe = /^(https?:|mailto:|tel:|#)/i.test(href) ? href : '#';
      return `<a href="${safe}" target="_blank" rel="noopener">`;
    });
    return out;
  }

  private _css(): string {
    return `<style>
.egx-pbk{--green:#00c853;--red:#bb162b;--bd:#e0e0e0;--tx2:#666;--bg2:#f5f5f5;font-family:'KiaSignature',-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;color:#111;max-width:1100px;margin:0 auto}
.egx-pbk *{box-sizing:border-box}
.egx-load,.egx-err{padding:32px;text-align:center;color:var(--tx2);font-size:14px}
.egx-err{text-align:left;border:1px solid var(--red);border-radius:12px;color:#111}.egx-err-hint{color:var(--tx2);font-size:12px}
.egx-h1{font-size:22px;font-weight:700;margin:0 0 16px}
.egx-hero{position:relative;height:240px;overflow:hidden;border-radius:12px;display:flex;align-items:flex-end;padding:0 28px 24px;margin-bottom:18px}
.egx-hero img{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;object-position:center 40%}
.egx-hero-ov{position:absolute;inset:0;background:linear-gradient(180deg,rgba(0,0,0,.15),rgba(0,0,0,.7))}
.egx-hero-tx{position:relative;z-index:2;color:#fff}.egx-hero-tx h1{font-size:26px;font-weight:700;margin:0}.egx-hero-tx p{font-size:13px;color:#d4d8dc;margin:6px 0 0;max-width:440px}
.egx-subnav{display:flex;gap:28px;border-bottom:.5px solid var(--bd);margin-bottom:24px;font-size:14px}
.egx-subnav a{color:var(--tx2);padding-bottom:8px;cursor:pointer;font-weight:500;border-bottom:2px solid transparent}
.egx-subnav a.active{color:#111;border-bottom-color:#111}
.egx-panel{display:none}.egx-panel.active{display:block}
.egx-emp-top{display:flex;gap:28px;align-items:flex-start;margin-bottom:22px;flex-wrap:wrap}
.egx-form{flex:1;min-width:280px}.egx-img{width:300px;max-width:100%}.egx-img img{width:100%;height:auto;display:block}
.egx-card{background:#fff;border:.5px solid var(--bd);border-radius:12px;overflow:hidden}
.egx-scroll{overflow-x:auto}
.egx-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:16px;padding:20px}
.egx-g{display:flex;flex-direction:column;gap:5px}
.egx-g label{font-size:11px;font-weight:500;color:var(--tx2)}
.egx-g input,.egx-g select{padding:9px 11px;background:var(--bg2);border:1px solid var(--bd);border-radius:6px;font-family:'SF Mono','Cascadia Code',Consolas,monospace;font-size:13px;outline:none}
.egx-g select{font-family:inherit;cursor:pointer}.egx-u{font-size:10px;color:#999}
.egx-head{padding:14px 20px;border-bottom:.5px solid var(--bd);display:flex;justify-content:space-between;align-items:center}
.egx-head h3{font-size:15px;font-weight:500;margin:0}.egx-badge{font-size:11px;color:var(--green);font-weight:500}
table.egx-bd,table.egx-cars{width:100%;border-collapse:collapse}
.egx-bd td{padding:9px 20px;font-size:13px;border-bottom:.5px solid rgba(0,0,0,.04)}
.egx-bd td:last-child{text-align:right;font-family:'SF Mono',Consolas,monospace;font-weight:500}
.egx-sec td{font-weight:600;background:var(--bg2);text-transform:uppercase;letter-spacing:.5px;font-size:11px;padding-top:13px!important}
.egx-tot td{font-weight:700;font-size:14px;border-top:2px solid var(--bd);padding-top:13px!important}
.egx-hl td{background:#f0fff4;font-weight:600}.egx-hl td:last-child{font-weight:700}
.egx-green{color:var(--green)!important}.egx-muted{color:var(--tx2)}.egx-sm{font-size:10px;color:var(--tx2)}
.egx-policy{background:#f0fff4;font-size:11px!important;color:var(--tx2);font-style:italic;text-align:left!important;font-family:inherit!important;font-weight:400!important}
.egx-ft{margin-top:16px;font-size:11px;color:var(--tx2);line-height:1.6}
.egx-hint{font-size:12px;color:var(--tx2);margin:0 0 14px;line-height:1.6}
.egx-cars-tb{display:flex;flex-wrap:wrap;gap:14px;align-items:flex-end;margin-bottom:16px}
.egx-count{font-size:13px;color:var(--tx2);margin-left:auto}
.egx-cars{min-width:640px;font-size:13px}
.egx-cars th{text-align:left;padding:12px 16px;background:var(--bg2);border-bottom:1px solid var(--bd);font-size:10px;font-weight:600;color:var(--tx2);text-transform:uppercase;letter-spacing:.6px}
.egx-cars td{padding:12px 16px;border-bottom:.5px solid rgba(0,0,0,.05)}
.egx-cars tr:hover td{background:#fafafa}
.egx-reg{font-family:'SF Mono',Consolas,monospace;font-weight:600;letter-spacing:.5px}.egx-reg.none{color:var(--tx2);font-weight:400}
.egx-sw{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:7px;vertical-align:-1px;border:1px solid rgba(0,0,0,.15)}
.egx-pill{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600}
.egx-pill-g{background:#e6f7ee;color:#1a7f4b}.egx-pill-a{background:#fdf2e3;color:#a25a00}.egx-pill-x{background:#eef0f2;color:#5b6675}
.egx-empty{text-align:center;color:var(--tx2);padding:28px}
.egx-faqwrap{margin-top:36px;padding-top:28px;border-top:.5px solid var(--bd)}
.egx-faq-lbl{font-size:11px;letter-spacing:1px;text-transform:uppercase;color:var(--tx2);font-weight:700}
.egx-faq-h{font-size:18px;font-weight:700;margin:4px 0 0}
.egx-faq{margin-top:18px;border:.5px solid var(--bd);border-radius:12px;overflow:hidden;background:#fff}
.egx-fi{border-bottom:.5px solid var(--bd)}.egx-fi:last-child{border-bottom:none}
.egx-fq{width:100%;text-align:left;background:none;border:none;padding:16px 20px;font-family:inherit;font-size:14px;font-weight:600;color:#111;cursor:pointer;display:flex;justify-content:space-between;gap:14px;align-items:center}
.egx-fq:hover{background:var(--bg2)}.egx-n{color:var(--tx2);font-family:'SF Mono',Consolas,monospace;font-size:12px;margin-right:11px;font-weight:400}
.egx-fic{width:18px;height:18px;flex-shrink:0;transition:transform .25s;color:var(--tx2)}
.egx-fi.open .egx-fic{transform:rotate(180deg)}
.egx-fa{display:grid;grid-template-rows:0fr;transition:grid-template-rows .28s}.egx-fi.open .egx-fa{grid-template-rows:1fr}
.egx-fa>div{overflow:hidden}.egx-fa p{padding:0 20px 18px 51px;font-size:13px;line-height:1.7;color:var(--tx2)}
.egx-fa a{color:#111;font-weight:500;text-decoration:underline;text-underline-offset:2px}
@media(max-width:640px){.egx-emp-top{flex-direction:column}.egx-img{width:100%;max-width:300px}.egx-count{margin-left:0;width:100%}.egx-bd td{padding:9px 14px}}
</style>`;
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Personalbil – inställningar' },
        groups: [{
          groupName: 'Datakälla',
          groupFields: [
            PropertyPaneTextField('dataFilePath', {
              label: 'Sökväg till datafilen (relativt webbplatsen)',
              description: 'T.ex. Shared Documents/personalbil-data.xlsx'
            }),
            PropertyPaneTextField('heroTitle', { label: 'Rubrik' }),
            PropertyPaneTextField('heroImageUrl', { label: 'Hero-bild URL (valfritt)' })
          ]
        }]
      }]
    };
  }

  protected onPropertyPaneFieldChanged(): void {
    this._data = null; this._err = ''; this.render();
  }
}
