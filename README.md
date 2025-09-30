# ewiser_public

## A munkafolyamat két fő szakaszból áll:
- Fejlesztői feladatok: Az adatok manuális letöltése az Ewiser felületéről és azok feldolgozása egy Python script segítségével.
- Ügyintézői feladatok: A feldolgozott adatok alapján a hibák kezelése és adminisztrációja a Google Sheets felületén.

## 1. Fejlesztői lépések (Adatkinyerés és -feldolgozás)
Ezt a folyamatot minden reggel el kell végezni a naprakész hibajelentések érdekében.

# Adatok letöltése
Két különböző adatkészlet letöltése szükséges:
# A. Piacos erőművek inverter hibái (siId=424)
1. Nyissa meg az Ewiser vezérlőpultját, majd navigáljon az Inverter hibák menüponthoz.
2. A böngésző fejlesztői eszköztárát (F12 vagy jobb klikk -> Inspect) megnyitva válassza a Network (Hálózat) fület.
3. Keresse meg és töltse le a latest-error-log?siId=424 hálózati kérésre érkező választ (response).
4. Mentse el a JSON fájlt a json_files mappába, a következő névadási konvencióval: éééé-hh-nn.json (pl. 2025-09-30.json).

# B. Kiemelt KÁT erőművek hibái (sig_id=119)
1. Hajtsa végre ugyanazt a letöltési folyamatot a sig_id=119 azonosítóval is, hogy a specifikus, monitorozni kívánt erőművek adatai is rendelkezésre álljanak.
2. Mentse el ezt a JSON fájlt a market_119 mappába, szintén éééé-hh-nn.json néven.

# Adatok feldolgozása
A letöltések után futtassa a load_data_to_gsheet.py Python scriptet. Ez a script feldolgozza a JSON fájlokat, alkalmazza a szűrési logikát és feltölti a releváns adatokat a cél Google Sheets dokumentumba.

## 2. Ügyintézői lépések (Hibakezelés)
# Hibák frissítése és szűrése
A Google Sheets dokumentumban kattintson a Hibák frissítése gombra. A megjelenő lista az alábbi feltételek alapján szűrt hibákat tartalmazza:
- A referencia inverter állapota nem OK.
- A teljesítménykülönbség aránya 10% vagy annál nagyobb (>= 0.1).
- A hibás inverterek száma több, mint nulla.

Fontos: A rendszer automatikusan figyelmen kívül hagy bizonyos, előre definiált erőműveket, miközben a kiemelt KÁT erőműveket (sig_id=119) mindig feldolgozza.

     

(Further Improvements) import the data into the spreadsheet (RAW sheet)
   - using python code to upload data (create the connection, requires: GCP project and enabling Google Sheets API + Google Drive API and service account or OAuth 2.0 Client IDs setup)
   - connect the spreadshet directly with the webpage (e.g.: https://hasdata.com/blog/google-sheets-web-scraping)
