# Batch Order Tracker — Netlify deploy

Frontend i `public/`, backend i `netlify/functions/`, datalagring via Netlify Blobs.

## Krav

- Node.js 18+ lokalt (kun for å initialisere passord-hash)
- En Netlify-konto (gratis)

## 1. Generer passord-hasher

Kjør dette én gang lokalt. Passordene lagres **aldri** i klartekst — kun sha256-hash i miljøvariabler.

```bash
node -e "console.log(require('crypto').createHash('sha256').update(process.argv[1]).digest('hex'))" "ditt-bruker-passord"
node -e "console.log(require('crypto').createHash('sha256').update(process.argv[1]).digest('hex'))" "ditt-admin-passord"
```

Du får ut to 64-tegns hex-strenger. Lagre dem trygt.

Generer også et JWT-secret:

```bash
node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
```

## 2. Deploy til Netlify

### Alternativ A — drag and drop (raskest)

1. Logg inn på app.netlify.com
2. Klikk "Add new site" → "Deploy manually"
3. Dra hele denne mappen inn

### Alternativ B — via Git (anbefalt)

1. Push denne mappen til et GitHub/GitLab-repo
2. I Netlify: "Add new site" → "Import from Git"
3. Pek på repoet. Netlify oppdager `netlify.toml` automatisk.

### Alternativ C — Netlify CLI

```bash
npm install -g netlify-cli
netlify login
netlify init
netlify deploy --prod
```

## 3. Sett miljøvariabler i Netlify

Site settings → Environment variables → Add a variable:

| Navn | Verdi | Bruk |
|---|---|---|
| `APP_PASSWORD_HASH` | hex fra steg 1 | Brukerpålogging |
| `ADMIN_PASSWORD_HASH` | hex fra steg 1 | Sletting |
| `JWT_SECRET` | hex fra steg 1 | Token-signering |
| `INGEST_TOKEN` | 64 hex-tegn fra `crypto.randomBytes(32)` | Power Automate (valgfri) |

`INGEST_TOKEN` trengs bare hvis du skal sette opp Power Automate til å synkronisere batch orders. Se `POWER-AUTOMATE.md` for fullstendig oppsett.

Etter at variablene er satt, klikk **Redeploy** på siste deploy.

## 4. Aktiver Netlify Blobs

Blobs er innebygd og krever ingen separat aktivering — det fungerer rett ut av boksen for hver Netlify-site. Første gang en Function skriver, opprettes lagringen automatisk.

## 5. Sjekk at det virker

Gå til `https://<din-site>.netlify.app`. Du skal se login-skjermen.
Logg inn med brukerpassordet. Last opp en BO.json. Start en registrering. Logg ut og inn igjen — registreringene skal være der.

For å teste admin-sletting: gå inn på en registrering, trykk SLETT, og oppgi admin-passordet.

## Endepunkter

| Path | Metode | Rolle | Beskrivelse |
|---|---|---|---|
| `/api/login` | POST | offentlig | `{ password, asAdmin? }` → `{ token, role }` |
| `/api/me` | GET | user/admin | Sjekker tokenet — returnerer `{ role }` |
| `/api/registrations` | GET | user/admin | Henter alle registreringer |
| `/api/registrations` | POST | user/admin | Oppretter eller oppdaterer |
| `/api/registrations/:id` | DELETE | **admin** | Sletter |
| `/api/active` | GET | user/admin | Alle pågående sesjoner |
| `/api/active` | POST | user/admin | Start eller oppdater sesjon |
| `/api/active?bo=...&line=...` | DELETE | user/admin | Forkast sesjon |
| `/api/batch-orders` | GET | user/admin | Henter siste lagrede batch orders |
| `/api/batch-orders` | POST | x-api-key | Erstatt hele lista (for Power Automate) |

## Sikkerhetsmodell

- Vanlig bruker (én passord) får alt unntatt sletting
- Admin-passord trengs kun for DELETE på registreringer
- JWT-token utløper etter 30 dager — brukeren må logge inn igjen
- Passord lagres som sha256-hash i miljøvariabler, aldri i klartekst
- Sammenligning gjøres i konstant tid (motvirker timing-angrep)
- 400ms forsinkelse ved feil passord (bremser brute-force)

## Endre passord senere

Generer ny hash (steg 1), oppdater miljøvariabelen i Netlify, redeploy. Eksisterende tokens fortsetter å virke til de utløper (det er en avveining — om du må kicke alle ut umiddelbart, bytt også `JWT_SECRET`).

## Datafiler i Netlify Blobs

To stores:
- `registrations` — alle BO_work_info
- `active-sessions` — pågående sesjoner

Hvert store har én nøkkel `all` som inneholder hele lista som JSON-array. Dette er enkelt og holder for 15 brukere og noen tusen registreringer. Hvis volumet vokser, bytt til én nøkkel per registrering.

Du kan inspisere data via Netlify CLI:

```bash
netlify blobs:list registrations
netlify blobs:get registrations all
```
