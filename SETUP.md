# Vista Portál — Setup instrukce

## Co dostaneš
- `index.html` — celý frontend, jde na GitHub Pages
- `Code.gs` — backend (Google Apps Script, zdarma, bez serveru)
- Data se ukládají do Google Sheets (máš je v Google Drive)
- Soubory/dokumenty jdou do složky "Vista Portal Dokumenty" v Google Drive

---

## KROK 1 — Google Sheets + Apps Script (10 minut)

1. Jdi na [sheets.google.com](https://sheets.google.com) a vytvoř **nový prázdný spreadsheet**
   - Pojmenuj ho třeba "Vista Portal Data"

2. V menu klikni **Rozšíření > Apps Script**

3. Smaž vše co tam je a **zkopíruj celý obsah souboru `Code.gs`**

4. Uložit (Ctrl+S), pojmenuj projekt "Vista Portal"

5. Spusť funkci `setup`:
   - V horním dropdownu vyber funkci `setup`
   - Klikni ▶ Spustit
   - Autorizuj přístup (klikni Pokročilé > přejít na Vista Portal)
   - Tím se vytvoří listy: Zaměstnanci, Dokumenty, Potvrzení

6. Nasaď jako Web App:
   - Klikni **Nasadit** (vpravo nahoře) > **Nové nasazení**
   - Typ: **Webová aplikace**
   - Spustit jako: **Já (tvůj účet)**
   - Kdo má přístup: **Kdokoli**
   - Klikni **Nasadit**
   - **Zkopíruj URL** — vypadá nějak takto:
     `https://script.google.com/macros/s/XXXXXXXXXXXX/exec`

---

## KROK 2 — Nastav URL v index.html

Otevři `index.html` a na začátku najdi:

```javascript
const CONFIG = {
  SCRIPT_URL: 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE',
  ADMIN_PIN: 'vista2024admin',
  ...
};
```

Nahraď `YOUR_GOOGLE_APPS_SCRIPT_URL_HERE` svojí URL z kroku 6.

Změň také `ADMIN_PIN` na vlastní heslo.

---

## KROK 3 — GitHub Pages (5 minut)

1. Vytvoř nové GitHub repo: [github.com/new](https://github.com/new)
   - Název: `vista-portal`
   - Viditelnost: **Public** (nebo Private s GitHub Pro)

2. Nahraj `index.html` do repa

3. Jdi do **Settings > Pages**
   - Source: **Deploy from a branch**
   - Branch: `main` / `/(root)`
   - Uložit

4. Za 1-2 minuty bude web dostupný na:
   `https://davidsiwy.github.io/vista-portal/`

---

## Přihlašování

### Zaměstnanec
- Vybere své jméno ze seznamu
- Zadá PIN (přidáváš v admin panelu)

### Admin
- Klikne "Přihlásit jako admin" (malý link dole)
- Zadá admin PIN (nastaven v CONFIG.ADMIN_PIN)

---

## Admin panel — co vidíš

- **Statistiky**: počet dokumentů, zaměstnanců, potvrzení
- **Nahrát dokument**: drag & drop nebo kliknutí, soubor jde do Google Drive
- **Přehled potvrzení**: tabulka kdo/co/kdy potvrdil
- **Zaměstnanci**: přidat/odstranit, nastavit PIN

---

## Kde jsou data?

Všechno v Google Sheets:
- List **Zaměstnanci**: jméno, pozice, PIN
- List **Dokumenty**: název, kategorie, URL souboru
- List **Potvrzení**: kdo, co, kdy potvrdil

Soubory (PDF atd.) jsou v Google Drive ve složce **"Vista Portal Dokumenty"**.

---

## Změna admin PINu

V `index.html` v CONFIG:
```javascript
ADMIN_PIN: 'tvoje_heslo_zde'
```

---

## Potřebuješ pomoct?

Dej vědět, udělám:
- Email notifikaci při každém potvrzení (přes Gmail + Apps Script)
- Export potvrzení do PDF
- Přidat povinné dokumenty (musí potvrdit do X dnů)
- Vlastní doménu (portal.vistaresort.cz)
