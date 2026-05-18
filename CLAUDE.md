# CLAUDE.md – Arbeitsanweisungen für Claude in diesem Repo

> **Sprache:** Deutsch für alles – Chat, Commit Messages, Doku, Code-Kommentare, UI-Texte.

## 1. Was ist das hier?

Repo der KI-Use-Case-Sammlung von **AI at AI Factory Austria** (Ann-Sophie Brenk, Venture Studio).
Vier zusammenhängende Bausteine:

1. **Use-Case-Inventar** – `KI_UseCases_Uebersicht.xlsx` ist die Single Source of Truth.
2. **Dashboard** – `KI_UseCases_Dashboard.html` wird **aus der Excel** gerendert. Zielgruppe: Management/Leitung und externe Partner/Kunden.
3. **KI-Inbox-Processor (Skill)** – `ki-inbox-processor.skill` verarbeitet neue Einträge aus der Inbox: recherchiert, schlägt KI-Lösung vor, klassifiziert (Abteilung + Use Case Art), prüft auf Duplikate, übernimmt nur Neues.
4. **Stakeholder-Reporting** – Dashboard ist primärer Reporting-Kanal.

**Projektstadium:** Prototyp / Exploration. Strukturen dürfen sich ändern, Geschwindigkeit > Konsistenz.

## 2. Excel-Schema (Stand: erfasst beim Repo-Setup)

### Sheet `Use Cases` – die Haupttabelle
| Spalte | Inhalt |
| --- | --- |
| `Nr.` | Fortlaufende Nummer |
| `Problem` | Identifiziertes Problem im Geschäftsalltag |
| `KI-Lösung` | Konkreter KI-basierter Lösungsvorschlag |
| `Abteilung` | Eine der 11 definierten Abteilungen (z. B. HR & People, Operations & Produktion, Abteilungsübergreifend) |
| `Use Case Art` | Eine der 10 definierten Arten (z. B. Prozessautomatisierung, Datenanalyse & Reporting) |
| `Quelle Workshop` | Workshop-Bezug (z. B. "Mai 2026") |

### Sheet `Cluster-Übersicht` – Taxonomie
`Cluster-Typ` (Abteilung/Use Case Art) · `Cluster` · `Beschreibung`. Wenn der Skill klassifiziert, **muss** er auf die hier definierten Cluster zurückgreifen, nicht neue erfinden – außer es gibt einen klaren Grund, dann **erst nachfragen**.

### Sheet `Auswertung`
Aggregierte Zählung pro Cluster. Wird beim Übernehmen neuer Use Cases mit aktualisiert.

### Sheet `Inbox`
Roher Sammelplatz für neue Use Cases. Spalten: `Identifiziertes Problem` · `KI-Lösung (Idee/Stichworte)` · `Workshop / Quelle` · `Notizen`. Aus diesem Sheet wandern Einträge nach Prüfung in `Use Cases`.

## 3. Workflow: Neuer Use Case (Inbox → Haupttabelle)

Wenn Ann-Sophie sagt "es sind neue Einträge in der Inbox" oder eine aktualisierte Excel ablegt:

1. **Inbox lesen** – jeden neuen Eintrag durchgehen.
2. **Recherchieren** – auf Basis Problem + Notizen eine fundierte KI-Lösung formulieren (Wirkprinzip, kein Marketing). Wenn externe Recherche/APIs nötig wären → **erst fragen** (siehe No-Gos).
3. **Duplikat-Check** – mit allen 56+ Einträgen in `Use Cases` abgleichen. Bei Treffer oder sehr ähnlichem Fall:
   - **Nicht** automatisch übernehmen.
   - In der Antwort an Ann-Sophie melden: "Eintrag X ähnelt Use Case Nr. Y – übernehmen, zusammenführen oder verwerfen?"
4. **Klassifizieren** – `Abteilung` und `Use Case Art` aus der Taxonomie (`Cluster-Übersicht`) wählen.
5. **In `Use Cases` ergänzen** – nächste freie `Nr.`, sauber formulierte Spalten.
6. **Excel-Datei nie destruktiv überschreiben** – siehe No-Gos. Änderungen über eine Kopie/Backup oder explizit autorisierten Schreibpfad.
7. **Dashboard sofort neu rendern** – sobald neue Use Cases in die Haupttabelle übernommen wurden.
8. **Bericht an Ann-Sophie** – kurze Zusammenfassung: was übernommen, was als Duplikat geflaggt, wie klassifiziert.

## 4. Dashboard

- `KI_UseCases_Dashboard.html` ist self-contained (CSS + Daten inline). Wird beim Rebuild komplett überschrieben.
- Stil aktuell: dunkelblau (`#1f3a5f`), weiße Karten, Akzent `#4a90e2`. **AI-Austria-Corporate-Look bei Überarbeitung mitdenken** (Branding/Farben/Logo konsistent halten).
- Filter: nach `Abteilung` und `Use Case Art`, plus Volltextsuche.

## 5. Tech-Stack

- **Python 3** für die Render-Pipeline (`rebuild_dashboard.py` ist im `ki-inbox-processor.skill` enthalten – Skill-File ist ein ZIP).
- `openpyxl` zum Lesen der Excel.
- Skill-Backend (Claude API vs. Claude Code Skill) **noch offen** – Architektur so halten, dass beides möglich bleibt.

## 6. Wie wir zusammenarbeiten

- **Erst fragen, dann ändern.** Prototyp-Phase: bei nicht-trivialen Änderungen Plan zeigen oder Rückfrage stellen.
- **Mehrwellen-Klärung ist ok.** Lieber 2–3 gezielte Fragerunden als raten.
- **Vorschläge willkommen** – z. B. Verbesserungen an der Excel-Struktur, neue Spalten, bessere Cluster.
- **Commits:** freier deutscher Stil, kurze klare Subjects ohne Prefix-Schema (z. B. `Inbox-Workflow: Duplikat-Check ergänzt`).
- **Vertraulichkeit:** locker – die Excel enthält keine streng vertraulichen Daten. Trotzdem: keine Use-Case-Details an externe Dienste schicken ohne Absprache.

## 7. Harte No-Gos

1. **Nie ohne Rückfrage Use Cases löschen.** Excel ist Source of Truth – Deletes immer explizit bestätigen lassen.
2. **Keine externen Daten/APIs ohne Absprache.** Web-Recherche, externe LLM-Aufrufe, Tool-Use über das hier Definierte hinaus → vorher fragen.
3. **Keine destruktiven Git-Ops** ohne ausdrückliche Anweisung: `git push --force`, `git reset --hard`, `git checkout .`, Branch-Löschungen.
4. **Excel-Datei nie automatisch überschreiben.** Schreibvorgänge auf `KI_UseCases_Uebersicht.xlsx` immer mit Backup-Kopie oder über einen sauberen, dokumentierten Schreibpfad – nie blind das Original ersetzen.

## 8. Offene Punkte / TODO

- Skill-Backend entscheiden (Claude API vs. Claude Code Skill).
- Render-Pipeline aus `.skill` herausziehen oder so lassen?
- AI-Austria-Branding für Dashboard finalisieren (Logo, Farbpalette verifizieren).
- Auto-Trigger Excel → Dashboard (lokaler Watcher? GitHub Action? Manueller Befehl?) noch nicht festgelegt.
