# AutoPräsi – Prozessdokumentation

Automatische Erstellung der Gottesdienst-Präsentation (PowerPoint) aus dem GoDi-Plan (Excel).

---

## Gesamtablauf

```
Start (main.py)
  |
  v
1. Nächsten Sonntag berechnen
  |
  v
2. Dropbox starten & synchronisieren (2 Min. warten)
  |
  v
3. GoDi-Plan Excel aus Dropbox lesen
  |
  v
4. Lied-Index aus Brüderrecords aufbauen
  |
  v
5. 7 Lieder im Index suchen & zuordnen
  |
  v
6. Hintergrundbild suchen
  |
  v
7. Präsentation zusammenbauen
  |     a) Folienplan erstellen
  |     b) Template + Lieder zusammenkopieren
  |     c) Hintergrundbild einsetzen
  |     d) Texte ersetzen (Thema, Datum, Bibeltexte)
  |
  v
8. Speichern auf Desktop + Kopie in Dropbox
  |
  v
9. Status an Supabase-Dashboard melden
```

---

## Schritt 1: Nächsten Sonntag berechnen

**Datei:** `main.py` → `next_sunday()`

- Berechnet automatisch den nächsten Sonntag ab heute
- Alternativ: Datum als Argument übergeben (`python main.py 2026-03-08`)

---

## Schritt 2: Dropbox synchronisieren

**Datei:** `excel_reader.py` → `_trigger_dropbox_sync()` + `_ensure_dropbox_sync()`

Damit immer die aktuelle Version der Excel-Datei verwendet wird:

1. **Prüfen** ob Dropbox läuft (`pgrep -x Dropbox`)
2. **Starten** falls nicht aktiv (`open -a Dropbox`), bis zu 15 Sek. warten
3. **Sync anstoßen** – Dropbox-Ordner zugreifen
4. **2 Minuten warten** damit Dropbox alle Änderungen synchronisieren kann
5. **Prüfen** ob die Datei aus dem Dropbox-Ordner stammt (kein lokaler Fallback)
6. **Zeitstempel loggen** – wann die Datei zuletzt geändert wurde

**Fehler wenn:** Dropbox nicht startet oder die Datei nicht aus Dropbox kommt → Abbruch mit RuntimeError.

---

## Schritt 3: GoDi-Plan Excel lesen

**Datei:** `excel_reader.py` → `read_godi_plan()`

### Datei finden
- Sucht `GoDi-Plan*.xlsx` in `~/Dropbox/Gemeinde/`
- Ignoriert Lock-Dateien (`~$...`)
- Öffnet jede gefundene Datei und prüft ob das Sheet `So {Tag}.{Monat}` existiert (z.B. "So 8.3")

### Daten auslesen (aus dem passenden Sheet)

**Seit Q4/2026 (8.10.2026) über Beschriftungen statt fester Zeilen**, weil sich die Zeilen
zwischen den Quartalen verschieben:
- **Layout:** Sonntags-Sheets haben die Inhalte in Spalte D, die schmaleren **Samstags-Sheets
  in Spalte C** (erkannt an Thema/GoDi-Leitung bzw. "Predigttext:"). Die API meldet
  `layout: "sonntag" | "samstag"`.
- **Thema / Lesung / Predigttext (Sa):** Zeile über Spalte-A-Beschriftung ("Thema:",
  "Lesungstext …", "Predigttext:").
- **Begrüßungsvers:** Inhalt der Zeile "Begrüßung" (Spalte B), sonst erstes Zitat ("…") vor
  den Abkündigungen. (Vorher fest Zeile 16 → in Q4 kam der "Liedbeitrag" als Vers.)
- **Lieder:** grüne Zelle (inkl. Sa-Grün `99CC00`) ODER Spalte B = Gemeindelied/
  Lobpreisstrophe/Kinderlied – jeweils nur, wenn der Text wie ein Lied aussieht.
  "Liedbeitrag" zählt nicht.
- **Abendmahl:** Zelle, die mit "Abendmahl" beginnt, vor den Abkündigungen (Q4: Spalte B
  Zeile 33 statt "Predigt 2", oder Spalte F bei Verkündigung).
- **Abkündigungen/Einladung:** ab der Zeile "Abkündigungen" bis "Ausblick"/"Danke an"/
  Schlusslied; Datumsspalte je Zeile erkannt (So: B, Sa: C), daneben Uhrzeit, Veranstaltung,
  (So) Hinweis in F. Termin ohne Datum (zweiter am selben Tag) übernimmt das Datum davor.
- **Header ohne Jahr** ("Gottesdienst am Samstag 03.10.") → Jahr aus dem Dateinamen.
- Die festen Zeilen unten gelten nur noch als (validierter) Fallback im Sonntags-Layout.

Im Präsentations-Reiter werden Abschnitte ohne Inhalt in der Excel (leere Lied-Slots,
keine Predigt 2/Lesung/Begrüßungsvers, Kinderstunde am Samstag) automatisch abgewählt –
sonst blieb z.B. die Platzhalter-Folie "Lied5" in der Präsentation.

| Daten (Fallback) | Zeile | Spalte |
|-------|-------|--------|
| Header ("Gottesdienst am...") | 1 | A |
| Thema | 3 | D |
| Begrüßungsvers | 16 | D |
| Lied 1 (Gemeindelied) | 19 | D |
| Lied 2 (Lobpreisstrophe) | 21 | D |
| Lied 3 (Kinderlied) | 24 | D |
| Predigt 1 Bibelstelle | 26 | D |
| Predigt 1 Thema | 26 | E |
| Lied 4 (Gemeindelied) | 27 | D |
| Lied 5 (Gemeindelied) | 32 | D |
| Predigt 2 Bibelstelle | 33 | D |
| Predigt 2 Thema | 33 | E |
| Lied 6 (Lobpreisstrophe) | 34 | D |
| Lesung Bibelstelle | 7 | D |
| Abendmahl-Indikator | 38 | B |
| Lied 7 (Schlusslied) | 47 | D |
| Abkündigungen | 37–45 | B–E |

### Lied-Einträge parsen
Jeder Lied-Eintrag wird aus dem Rohtext analysiert:
- **Kategorie** erkennen: "Lobpreisstrophe:", "Kinderlied:", "Sonstige Lieder"
- **Buch + Nummer** extrahieren: z.B. "FJ1 235" → Buch=FJ1, Nummer=235
- **Titel** extrahieren: z.B. "Jesus, dir nach"
- **Titelwörter** für Suche: erste 3 Wörter des Titels

Beispiel: `"Lobpreisstrophe: FJ1 239 - Immer mehr"` →
- Kategorie: Lobpreisstrophe
- Buch: FJ1
- Nummer: 239
- Titel: Immer mehr

---

## Schritt 4: Lied-Index aufbauen

**Datei:** `song_finder.py` → `build_song_index()`

- Durchsucht alle Unterordner in `~/Dropbox/Gemeinde/Brüderrecords°/`
- Indexiert alle `.pptx`-Dateien (ca. 727 Dateien in 23 Ordnern)
- Ergebnis: Dictionary `{Ordnername: [(Dateiname, voller_Pfad), ...]}`

### Lied-Ordner (Brüderrecords)
```
FJ1, FJ2, FJ3, FJ4, FJ5, FJ6     (Feiert Jesus, Band 1–6)
GLS, SUG, IWDD, SGIDH 2           (Verschiedene Liederbücher)
Loben, Lobpreisstrophen            (Lobpreis)
Kinderlieder                       (Kinderlieder)
Sonstige Lieder                    (Einzellieder)
Einfach spitze                     (Kinderlieder-Sammlung)
```

---

## Schritt 5: Lieder suchen

**Datei:** `song_finder.py` → `find_song()`

Für jedes der 7 Lieder:

### Suchordner bestimmen
- Lobpreisstrophe → Ordner "Lobpreisstrophen"
- Kinderlied → Ordner "Kinderlieder"
- Sonstige Lieder → Ordner "Sonstige Lieder"
- Buch "FJ1" → Ordner "FJ1"
- Buch "FJ" (ohne Nummer) → alle FJ1–FJ6 durchsuchen

### Buch normalisieren
- "FJ 3", "FJ3" und "fj 3" treffen denselben Ordner (Leerzeichen werden ignoriert)
- Ohne diese Normalisierung landete "FJ 3 87 - Ich bin bei dir" früher im
  Fallback und bekam den ersten 87er aus irgendeinem Ordner (FJ1: "Lord you have my heart")

### Suchstrategie 1: Nach Nummer + Titel
- Dateiname beginnt mit der Liednummer ("235 - …", "235- …", "235 -…", "235-…")
- Lobpreisstrophen: Buch + Nummer → "FJ1 239 - Immer mehr.pptx" (Buch-Präfix optional)
- **In den primären Ordnern** ist die Nummer verbindlich. Gibt es mehrere Dateien
  mit derselben Nummer (z.B. "87 - Lord you have my heart.pptx" und
  "87 - … - deutsch.pptx"), entscheidet der Titel aus dem Excel
- **In allen anderen Ordnern** zählt ein Nummern-Treffer nur, wenn auch der Titel
  passt (jedes Liederbuch hat eine Nr. 87)
- Rangfolge bei mehreren Kandidaten: Titel passt → primärer Ordner → Ordner-Priorität
  (Lobpreisstrophen vor FJ1) → vollständigerer Titel → wenigste überzählige Wörter

### Suchstrategie 2: Nach Titelwörtern
- Wenn keine Nummer vorhanden (Kinderlieder, Sonstige) oder Strategie 1 nichts liefert
- Alle 3 ersten Titelwörter müssen als ganze Wörter im Dateinamen vorkommen
  (ab 4 Zeichen reicht der Wortanfang: "herz" trifft "Herzen"; "ich" trifft NICHT "nicht")
- Umlaute werden auf NFC normalisiert (Dateinamen aus Dropbox/macOS kommen teils
  zerlegt an – sonst findet "Höre mich" nie "Höre mich.pptx")

### Prüf-Hinweise (`match_warning`)
Die API liefert zu jedem Treffer einen Hinweis, der im Frontend als gelber Punkt +
Text angezeigt wird, wenn der Treffer nicht eindeutig ist:
- "Titel passt nicht zum Dateinamen" – kein Titelwort passt (z.B. andere Sprachfassung)
- "Mehrere passende Dateien (auch: …)" – gleich gute Kandidaten im selben Ordner
- "In anderem Ordner gefunden (…)" – Treffer außerhalb des Buch-/Kategorie-Ordners

### Tests
```bash
cd autopraesi_python && venv/bin/python -m unittest -v
```
`tests/test_song_finder.py` sichert die Suchlogik gegen einen synthetischen Index ab,
`tests/test_text_fit.py` die Schriftanpassung gegen das Template,
`tests/test_excel_reader.py` das Einlesen (Q4-Sonntag, Samstags-Layout, Abendmahl) gegen
im Speicher nachgebaute Arbeitsmappen.

---

## Schritt 6: Hintergrundbild suchen

**Datei:** `main.py` → `_find_image()`

- Sucht in `/Gemeinde/` (api.py `_find_image`) tolerant: optional "Bild ", Tag/Monat ein-
  oder zweistellig, ein oder zwei Punkte, `.jpg/.jpeg/.png` – z.B. `Bild 08.03.jpg`,
  `8.3.jpg`, `11.10..png`. Mehrere Treffer → das zuletzt geänderte gewinnt.
- Beim Einsetzen wird das Bild als JPEG ins Folienformat 4:3 gebracht
  (`_fit_image_to_slide`): nahe 4:3 → mittig zugeschnitten; Hochformat o.ä. → vollständig
  auf unscharfem Hintergrund aus demselben Bild (nicht mehr verzerrt).
- Optional – wenn nicht vorhanden, wird das Template-Bild beibehalten

---

## Schritt 7: Präsentation zusammenbauen

**Dateien:** `presentation_builder.py` + `slide_copier.py`

### Phase 7a: Folienplan erstellen
Das Template hat 37 Folien. An 7 Positionen stehen Platzhalter-Folien für Lieder:

| Slot | Template-Folie | Beschreibung |
|------|---------------|--------------|
| song1 | 5 | Gemeindelied 1 |
| song2 | 10 | Lobpreisstrophe |
| song3 | 15 | Kinderlied |
| song4 | 22 | Gemeindelied |
| song5 | 26 | Gemeindelied |
| song6 | 32 | Lobpreisstrophe |
| song7 | 36 | Schlusslied |

Für jede Position wird entschieden:
- Lied gefunden → Lied-PPTX einfügen (kann mehrere Folien haben)
- Nicht gefunden → Platzhalter-Folie beibehalten

### Phase 7b: Template + Lieder kopieren
- Template wird als Basis kopiert
- Alle bestehenden Folien werden entfernt
- Folien werden gemäß Plan neu aufgebaut:
  - Template-Folien: Shapes, Hintergründe und Bilder werden 1:1 kopiert
  - Lied-Folien: Aus den Song-PPTX-Dateien eingefügt, mit eigenem Song-Master
  - Seitenzahl ("1 / 3") und Quellenangabe ("Text: …") werden auf die Foliennummer- bzw.
    Fußzeilen-Platzhalter des Templates umgehängt, auch wenn die Lied-Datei sie auf anderen
    Platzhaltern hat (z.B. "Sonstige Lieder/Gottes guter Segen" – sonst überlagerte die
    Quellenangabe den Liedtext).

### Phase 7c: Hintergrundbild einsetzen
- Sucht das Template-Platzhalter-Bild auf allen Folien
- Ersetzt es durch das Sonntagsbild (nur kleinere Bilder werden ersetzt → keine Song-Logos)

### Phase 7d: Texte ersetzen
**Einfache Platzhalter:**
- "Gottes Liebe erkennen" → Thema des Gottesdienstes
- "Gottesdienst am 26.11.2023," → aktuelles Datum
- "Letzter Sonntag des Kirchenjahres" → Kirchenkalender-Name
- "Gottesdienst am XY.XY.XY" → aktuelles Datum (auf Zwischenfolien)
- "Thema" → Thema (auf Zwischenfolien)
- "Name des Sonntags" → Kirchenkalender-Name
- "Begrüßungsvers" → Vers aus Excel
- "Predigt zu Predigtext1/2" → Bibelstelle + Thema

**Schrift an die Folie anpassen:**
Begrüßungsvers, Predigt-Thema und Predigt-Bibelstelle sind im Template
`spAutoFit`-Boxen (Box wächst mit dem Text, Schrift bleibt gleich). Lange Texte
liefen deshalb unten aus der Folie. Da python-pptx Text nicht messen kann, wird
der Zeilenumbruch näherungsweise simuliert (`_estimate_line_count`, kalibriert auf
Baskerville Old Face) und die Schrift in 2pt-Schritten verkleinert, bis der Block
in den verfügbaren Bereich passt (`_fit_shape_to_slide`), mindestens 24pt
(Bibelstellen-Zeile 20pt). Beispiel: 215-Zeichen-Vers → 54pt → 42pt.

**Bibeltexte (automatisch geladen):**

**Datei:** `bible_fetcher.py`

1. Bibelstelle parsen: "Lukas 9, 57-62" → Buch=LUK, Kapitel=9, Verse=57–62
2. HTML laden von `die-bibel.de/bibel/LU84/{Buch}.{Kapitel}` (Luther 1984)
3. Verse aus HTML extrahieren (CSS-Klasse `LU84.{Buch}.{Kapitel}.{Vers}`)
4. Verse auf Folien aufteilen (max. 350 Zeichen pro Folie)
5. Bei langen Texten: Folie duplizieren und Seitenzahlen aktualisieren (z.B. "1 / 3")
6. Versnummern werden hochgestellt dargestellt

Drei Bibeltexte werden geladen:
- **Lesung** (Row 7) → Platzhalter "Lesungstext"
- **Predigt 1** (Row 26) → Platzhalter "Predigttext1"
- **Predigt 2** (Row 33) → Platzhalter "Predigttext2"

---

## Schritt 8: Speichern

**Datei:** `presentation_builder.py`

- **Desktop:** `~/Desktop/Autopräsi/So {Tag}.{Monat}_ungeprüft.pptx`
- **Dropbox:** `~/Dropbox/Gemeinde/So {Tag}.{Monat}_ungeprüft.pptx` (Kopie via `shutil.copy2`)

Der Dateiname enthält "_ungeprüft" als Hinweis, dass die Präsentation noch manuell kontrolliert werden soll.

---

## Schritt 9: Status melden

**Datei:** `status_reporter.py`

Sendet den Ergebnis-Status an eine Supabase-Datenbank (Tabelle `runs`):
- **success** – alles gefunden und erstellt
- **partial** – Präsentation erstellt, aber Lieder fehlen
- **error** – Abbruch wegen Fehler

Gemeldete Daten:
- Sonntags-Datum, Thema, Kirchenkalender
- Status aller 7 Lieder (gefunden/nicht gefunden)
- Bibelstellen, Abkündigungen
- Anzahl Folien, Laufzeit
- Ob Hintergrundbild gefunden wurde

Das Dashboard (`dashboard/index.html`) zeigt die letzten 20 Läufe an und aktualisiert sich alle 60 Sekunden.

---

## Ordnerstruktur

```
~/Desktop/Autopräsi/
  ├── autopraesi_python/
  │   ├── main.py                  ← Einstiegspunkt
  │   ├── config.py                ← Pfade & Konfiguration
  │   ├── excel_reader.py          ← Excel lesen + Dropbox-Sync
  │   ├── song_finder.py           ← Lieder suchen
  │   ├── presentation_builder.py  ← Präsentation zusammenbauen
  │   ├── slide_copier.py          ← Folien kopieren (low-level)
  │   ├── bible_fetcher.py         ← Bibeltexte von die-bibel.de
  │   ├── status_reporter.py       ← Status an Supabase melden
  │   ├── dashboard/
  │   │   └── index.html           ← Web-Dashboard
  │   ├── logs/                    ← Log-Dateien
  │   └── venv/                    ← Python-Umgebung
  └── v_1_4_3_ab2026/
      └── Vorlage_Godi_Standard_Sonntag_v1_4_3.pptx  ← Template

~/Dropbox/Gemeinde/
  ├── GoDi-Plan 2026_1.xlsx        ← Eingabe (Excel)
  ├── Brüderrecords°/              ← Lied-Bibliothek (23 Ordner, ~727 Dateien)
  ├── Bild DD.MM..jpg              ← Hintergrundbilder
  └── So X.Y_ungeprüft.pptx        ← Ausgabe (Kopie)
```

---

## Abhängigkeiten

| Paket | Zweck |
|-------|-------|
| python-pptx | PowerPoint lesen/schreiben |
| openpyxl | Excel lesen |
| requests | HTTP-Requests (Bibeltexte, Supabase) |
| lxml | XML-Verarbeitung (Folien-Interna) |
| beautifulsoup4 | HTML parsen (Bibeltexte) |

---

## Ausführung

```bash
cd ~/Desktop/Autopräsi/autopraesi_python
source venv/bin/activate
python main.py                    # Nächsten Sonntag
python main.py 2026-03-08         # Bestimmtes Datum
```

---

## Fehlerfälle

| Fehler | Ursache | Reaktion |
|--------|---------|----------|
| RuntimeError: Dropbox läuft nicht | Dropbox-App nicht installiert/startbar | Abbruch |
| GoDi-Plan nicht gefunden | Kein Sheet für den Sonntag in der Excel | Abbruch |
| Lied nicht gefunden | Dateiname in Brüderrecords passt nicht | Warnung, Platzhalter bleibt |
| Bibeltext nicht ladbar | die-bibel.de nicht erreichbar | Platzhalter-Text bleibt |
| Dropbox-Kopie fehlgeschlagen | Dropbox-Ordner nicht vorhanden | Warnung, Desktop-Version existiert |
