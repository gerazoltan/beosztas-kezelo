# Fejlesztői architektúra

## Adatfolyam és adatvédelem

A `File` objektum `ArrayBuffer` tartalma közvetlenül a böngésző memóriájába kerül. A kód nem használ feltöltési végpontot, `localStorage`-ot vagy `IndexedDB`-t. Új fájl választásakor az előző `WorkbookSession`, a kiválasztott dolgozó és az eredmények kikerülnek a React állapotából. A munkafüzet tartalma és az OAuth-token nem kerül naplózásra.

## Parser

- `src/excel/workbookParser.ts`: felismeri a magyar hónapnevű munkalapokat, ellenőrzi a lap fejlécében a hónapot és az évet, megkeresi az egyetlen `Név` fejlécet és a naptári napok oszlopait. A `Munka*` lapokat kihagyja.
- `src/excel/ooxml.ts`: vékony, csak olvasó ZIP/XML réteg. A workbook relationshipjeiből diagnosztikai `styleId`-térképet, a theme XML-ből színpalettát készít.
- `src/excel/cellValues.ts`: az ExcelJS által megjelenített szöveget olvassa ki, és kezeli a `null` képleteredményhez hasonló hibás belső értékeket.
- `src/excel/dayEntries.ts`: kizárólag a kiválasztott dolgozó sorát olvassa. A deklarált worksheet dimension helyett a felismert napfejlécek határozzák meg a napi mátrixot, így az oldalsó KMR-blokk kívül marad.

Egy nap `startColumn..endColumn` fizikai cellacsoport. A következő nap fejlécének oszlopa zárja le az előző csoportot; az utolsó nap szélessége az előző csoportok jellemző szélességéből származik. Merge esetén a master cella egyedi címe deduplikálja a két fizikai cellát. Nem merge-elt csoportban nulla, egy vagy több nem üres cella rendre üres, biztos vagy kettős bejegyzést jelent.

## Színfeloldás és 12-es szolgálat

`src/excel/colors.ts` először a közvetlen ARGB-t használja. Theme indexnél az OOXML-paletta színére alkalmazza az Excel tint/shade transzformációját. Solid pattern kitöltésnél kizárólag az `fgColor` a látható háttér. A 12-es besorolása:

1. egyezés a havi jelmagyarázatból felismert stílusmintával;
2. zöld: dőlt és feloldott zöld betűszín;
3. kék: nem dőlt és feloldott kék kitöltés;
4. egyébként bizonytalan, nem exportálható.

A `styleId` soha nem döntési feltétel, csak diagnosztikai adat.

## Műszakpárosítás

`src/services/shifts.ts` tiszta függvénye a kiválasztott hónap napjain halad:

- kék/zöld `12` és napi `KMR`: azonnali esemény;
- `17` vagy `5`: csak a valóban következő naptári napi `7` esetén esemény, a `7` elfogyasztott záró jelölés;
- első napi `7`: az előző havi lap utolsó napjával párosítható;
- utolsó napi `17`/`5`: a következő havi lap első napjával párosítható;
- távollét: kizárt sor;
- kettős, érvénytelen, ismeretlen vagy párosítatlan adat: nem exportálható hiba/bizonytalanság.

## ICS

`src/services/ics.ts` CRLF sorvégű iCalendar 2.0 fájlt készít. A `VTIMEZONE` CET/CEST szabályokat tartalmaz, az események `TZID=Europe/Budapest` helyi időpontok. Az UID az esemény tartalmából stabilan származik. A szöveg escape-elt, a sorhajtás 75 UTF-8 bájton történik.

## Google Naptár

`src/services/googleOAuth.ts` a Google Identity Services token kliensét tölti be. A token kizárólag memóriában él, kijelentkezéskor törlődik és visszavonásra kerül. `src/services/googleCalendar.ts`:

- csak tulajdonosi/írói naptárakat listáz;
- pontos summary + helyi start + helyi end egyezésre ellenőriz duplikációt;
- lekéri a színpalettát és RGB-távolsággal sötétzöldet választ;
- szekvenciálisan ír, így részleges hibánál eseményenkénti eredmény marad;
- meglévő eseményt nem módosít és nem töröl.

## Tesztrétegek

- `tests/`: Vitest unit, ExcelJS-sel futás közben előállított anonimizált munkafüzetek és Testing Library UI-tesztek.
- `tests/localSample.regression.test.ts`: kizárólag helyi, ignorált minta; hiányában automatikusan skip.
- `e2e/`: Playwright desktop és mobil smoke teszt; a fixture memóriában készül, bináris személyes adat nem kerül a repositoryba.
