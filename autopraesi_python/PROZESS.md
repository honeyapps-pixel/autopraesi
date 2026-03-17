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
| Daten | Zeile | Spalte |
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

### Suchstrategie 1: Nach Nummer
- Dateiname beginnt mit der Liednummer
- z.B. Nummer "235" → findet "235 - Jesus dir nach.pptx"
- Lobpreisstrophen: Buch + Nummer → "FJ1 239 - Immer mehr.pptx"

### Suchstrategie 2: Nach Titelwörtern
- Wenn keine Nummer vorhanden (Kinderlieder, Sonstige)
- Alle 3 ersten Titelwörter müssen im Dateinamen vorkommen
- z.B. "Gottes große Liebe" → findet "Gottes große Liebe.pptx"

### Fallback
- Wenn in den primären Ordnern nichts gefunden wird → alle Ordner durchsuchen

---

## Schritt 6: Hintergrundbild suchen

**Datei:** `main.py` → `_find_image()`

- Sucht in `~/Dropbox/Gemeinde/` nach `Bild DD.MM..jpg`
- z.B. für 08.03.2026 → `Bild 08.03..jpg`
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
