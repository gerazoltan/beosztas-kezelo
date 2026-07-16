# Beosztáskezelő

Magyar nyelvű, teljesen kliensoldali webalkalmazás havi Excel-beosztások ellenőrzésére, ICS-fájlba exportálására és opcionális Google Naptárba írására.

## Fő funkciók

- `.xlsx` fájl kiválasztása vagy behúzása;
- magyar hónapnevű munkalapok, hónap/év fejléc, `Név` oszlop és két fizikai oszlopból álló napi cellacsoportok felismerése;
- egy kiválasztott dolgozó sorának feldolgozása, név kódba égetése nélkül;
- kék és zöld 12, 17–7, 5–7, napi KMR és hónapváltó párok értelmezése;
- érvénytelen nap, kettős napi cella, párosítatlan vagy ismeretlen jelölés biztonságos kizárása;
- részletes, kijelölhető ellenőrző nézet és mobil kártyanézet;
- Europe/Budapest időzónás ICS-export;
- opcionális, duplikációt ellenőrző Google Naptár-integráció.

## Adatvédelem

A feltöltött fájlt a böngésző helyben olvassa. Nincs szerver, adatbázis vagy feltöltési végpont; a fájl és az eredmény nem kerül `localStorage`-ba vagy `IndexedDB`-be. Új fájl vagy oldalfrissítés törli a memóriabeli állapotot. A Google OAuth-token is csak memóriában él, és kijelentkezéskor törlődik.

A `local-samples/`, `*.xlsx` és `*.xls` git által ignorált. Valódi minta nem része a nyilvános repositorynak.

## Technológia

React, Vite, szigorú TypeScript, ExcelJS, JSZip, Vitest, Testing Library, Playwright, ESLint és Prettier. Az ExcelJS a fő feldolgozó; egy vékony OOXML-réteg csak a theme színek és a diagnosztikai style ID kiolvasására szolgál.

Részletesebb leírás: [docs/ARCHITEKTURA.md](docs/ARCHITEKTURA.md).

## Helyi indítás

Node.js 22 vagy újabb ajánlott.

```bash
npm ci
npm run dev
```

A Vite által kiírt helyi URL-t kell megnyitni. A production előnézet:

```bash
npm run build
npm run preview
```

## Ellenőrzések

```bash
npm run lint
npm run typecheck
npm test
npx playwright install chromium
npm run test:e2e
npm run build
```

A tesztfixture-ök futás közben, anonimizált adatokból készülnek. A `tests/localSample.regression.test.ts` helyben a `local-samples/` első `.xlsx` fájlját is ellenőrzi, de a fájl hiányában CI-ben automatikusan kihagyódik.

## Támogatott Excel-szerkezet

- magyar hónapnév a munkalap nevében és az évvel együtt a fejlécben;
- egyetlen `Név` fejlécet és 1–31 napértékeket tartalmazó fejlécsor;
- egymás utáni napfejlécekből meghatározható fizikai cellacsoportok (a referenciaformátumban két oszlop/nap);
- merge-elt vagy különálló napi cellák, bal vagy jobb oldali értékkel;
- `Összesen` vagy `Összesen:` zárósor;
- közvetlen ARGB vagy theme + tint/shade színek.

A régi bináris `.xls` formátumot az ExcelJS nem támogatja; ezt előbb `.xlsx` formátumba kell menteni. Titkosított/jelszavas fájl nem dolgozható fel. Feltételes formázással előállított, közvetlen cellastílusban és jelmagyarázatban nem azonosítható szín bizonytalanként kizárható.

## Google OAuth konfiguráció

Google nélkül az ICS-export teljesen működik. A közvetlen integrációhoz:

1. Google Cloud projektben engedélyezd a **Google Calendar API**-t.
2. Állítsd be az OAuth consent screent és hozz létre **Web application** típusú OAuth client ID-t. Client secret nem kell és nem kerülhet a repositoryba.
3. Helyi fejlesztéshez add az authorized JavaScript origins listához a tényleges Vite origint, például `http://localhost:5173`.
4. GitHub Pageshez add hozzá: `https://<felhasználó>.github.io`. A Google Identity Services token flow popupot használ; külön redirect URI-t ez az alkalmazás nem használ.
5. Helyben készíts `.env.local` fájlt:

   ```text
   VITE_GOOGLE_CLIENT_ID=1234567890-example.apps.googleusercontent.com
   ```

6. GitHubon a repository **Settings → Secrets and variables → Actions → Variables** részében hozz létre `VITE_GOOGLE_CLIENT_ID` változót.

Az alkalmazás csak a Calendar eseményírási és naptárlista-olvasási scope-okat kéri. Pontos summary/kezdés/befejezés egyezésnél nem készít duplikátumot. Meglévő eseményt nem módosít és nem töröl.

## GitHub Pages telepítés

A Vite base útvonal `/beosztas-kezelo/`. A `.github/workflows/deploy-pages.yml` minden `main` pushnál futtatja a lintet, typechecket, unit/integrációs és Chromium E2E teszteket, majd a production buildet. Csak sikeres `verify` job után telepít.

GitHubon a **Settings → Pages → Build and deployment → Source** értéke legyen **GitHub Actions**. A workflow kézzel is indítható az Actions felületen.

## Hibaelhárítás és korlátozások

- „Nincs havi munkalap”: ellenőrizd a magyar hónapnevet, a hónap–év fejlécet, a `Név` fejlécet és a napértékeket.
- Többször szereplő név: a felület kötelezően kéri a konkrét sorszám kézi kiválasztását.
- Ismeretlen 12: a jelmagyarázat, a feloldott szín és a dőlt állapot alapján sem volt biztonságosan besorolható; szándékosan nem exportálódik.
- Első/utolsó napi 7, 17 vagy 5: a szomszédos havi lap és ugyanazon dolgozó szükséges a párosításhoz.
- Google nincs konfigurálva: állítsd be a környezeti változót, vagy használd az ICS-exportot.
- Google 403: ellenőrizd a Calendar API-t, az origint, a consent screen tesztfelhasználóit és a kiválasztott naptár írási jogát.
- A Google valódi API-hívásai csak kézzel konfigurált Cloud projekt mellett próbálhatók; az automatizált tesztek mockolt API-válaszokat használnak.
