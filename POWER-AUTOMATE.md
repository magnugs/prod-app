# Power Automate → Batch Tracker (automatisk synkronisering)

Denne guiden setter opp en Power Automate-flow som henter batch orders fra D365 F&O hver time og sender dem til Netlify-backenden via `/api/batch-orders`.

## 1. Generer et API-token

I terminal eller nettleserkonsoll:

```javascript
// Nettleserkonsoll:
console.log([...crypto.getRandomValues(new Uint8Array(32))].map(x => x.toString(16).padStart(2,"0")).join(""))
```

```bash
# Terminal:
node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
```

Du får 64 hex-tegn. Lagre dette som `INGEST_TOKEN` i Netlify (Site settings → Environment variables → Add). Redeploy etter.

**Dette tokenet identifiserer Power Automate.** Det er ikke et passord — det er en API-key som gir tilgang KUN til å skrive batch orders, ingenting annet. Behandle det likevel som hemmelig.

## 2. Opprett Power Automate-flow

### Trigger: Recurrence

- **Frequency:** Hour
- **Interval:** 1

(Du kan også sette en tidsperiode hvis du bare vil at den skal kjøre i arbeidstid — f.eks. mellom 06:00 og 22:00.)

### Steg 1: Hent batch orders fra D365 F&O

Bruk **"Get items"**-action fra Dataverse/Common Data Service-konnektoren — eller HTTP-action mot OData-endepunktet direkte hvis dere har OAuth satt opp.

Direkte mot OData:

- **Method:** GET
- **URI:** `https://<din-d365-tenant>.dynamics.com/data/BatchOrderHeaders?$top=10000`
- **Authentication:** Active Directory OAuth (D365 service principal)
- **Headers:**
  - `Accept: application/json`
  - `OData-Version: 4.0`

Hvis dere har mange batch orders, filtrer på relevante (f.eks. ikke fullførte):

```
?$filter=BatchOrderStatus ne Microsoft.Dynamics.DataEntities.BatchOrderStatus'Completed'&$top=10000
```

### Steg 2: Send til Netlify

Legg til en **HTTP**-action:

- **Method:** POST
- **URI:** `https://<din-site>.netlify.app/api/batch-orders`
- **Headers:**
  - `Content-Type: application/json`
  - `x-api-key: <ditt-INGEST_TOKEN>`
  - `x-source: PowerAutomate` (valgfri — vises som "synket fra" i appen)
- **Body:** Bruk output fra steg 1 direkte.

Hvis OData-responsen er `{ "value": [...] }`, kan du sende hele objektet — backenden støtter det. Hvis du har transformert det til en ren array, sender du arrayen.

### Steg 3 (anbefalt): Feilhåndtering

Legg til en **"Configure run after"** på HTTP-actionen som varsler deg ved feil:

- Klikk på Steg 2 → "..." → Configure run after
- Sjekk "has failed", "is skipped", "has timed out"
- Legg til en handling som sender deg en e-post eller Teams-melding

Vanlige feilmeldinger:

| Status | Betydning | Hva du gjør |
|---|---|---|
| 401 Ugyldig API-key | `INGEST_TOKEN` matcher ikke | Sjekk miljøvariabel i Netlify, redeploy |
| 400 Tom liste | OData returnerte ingen batch orders | Sjekk filteret ditt |
| 400 Forventet objekter med BatchOrderNumber | Feil format på body | Sjekk at du sender selve listen |
| 500 INGEST_TOKEN ikke konfigurert | Miljøvariabel mangler | Legg til i Netlify |

## 3. Test første kjøring

Klikk **"Test"** i Power Automate-flowen → "Manually". Du bør se grønt på begge stegene innen noen sekunder.

Gå deretter til Batch Tracker-appen og trykk **OPPDATER**. Du skal se batch orders dukke opp, og headeren skal vise "synket [tidspunkt]".

## 4. Test fra terminal (uten Power Automate)

For å teste backenden direkte:

```bash
curl -X POST https://<din-site>.netlify.app/api/batch-orders \
  -H "Content-Type: application/json" \
  -H "x-api-key: <ditt-INGEST_TOKEN>" \
  -d '{"value":[{"BatchOrderNumber":"TEST-001","BatchOrderName":"Test","ItemNumber":"X","ScheduledDate":"2026-01-01T00:00:00Z","ScheduledQuantity":1,"BatchOrderStatus":"Created","ProductionSiteId":"BBS","PlannedReceiptWarehouseId":"Inbound"}]}'
```

Skal returnere `{"ok":true,"updatedAt":"...","count":1,"source":"api"}`.

## Datakilde i appen

Når Power Automate er aktivt, har dere to kilder til batch orders:

1. **Automatisk:** Power Automate skriver → backend → "OPPDATER"-knappen i appen henter
2. **Manuell:** "FIL"-knappen lar deg laste opp en JSON-fil hvis du trenger å teste noe spesifikt

Ved oppstart prøver appen alltid backenden først. Manuell upload-panel vises bare hvis backenden er tom.

## Hva skjer ved feil i Power Automate?

Hvis en kjøring feiler, beholdes forrige vellykkede last i Netlify Blobs. Det vil si: appen viser fortsatt batch orders fra den siste vellykkede synkroniseringen. Sist-synket-tidspunktet i headeren forteller deg hvor gamle dataene er.

## Begrensninger

- **Erstatter hele lista hver kjøring.** Hvis Power Automate sender 5000 batch orders kl. 13:00 og kl. 14:00 er det bare 4900 (en ble slettet i D365), reflekterer appen automatisk det.
- **Ingen historikk.** Vi lagrer kun siste versjon. Hvis dere trenger historikk over hvordan en batch order endret seg over tid, må vi utvide datalaget.
- **Maks Netlify Blob-størrelse:** 5 MB per nøkkel. Med 10 000 batch orders á ~1 KB hver havner vi rundt 10 MB → over grensen. Hvis dere har mer enn ~3000 batch orders, må vi splitte på flere nøkler. Si fra hvis det blir aktuelt.
