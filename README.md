# AI Austria – Use Cases (Ann-Sophie)

Strukturierte Sammlung der KI-Use-Cases aus den Discovery-Workshops von **AI at AI Factory Austria**, plus ein gerendertes Dashboard für Management und Partner.

## Inhalt

| Datei | Zweck |
| --- | --- |
| `KI_UseCases_Uebersicht.xlsx` | **Single Source of Truth** – alle Use Cases, Cluster, Auswertung und die Inbox für neue Einträge |
| `KI_UseCases_Dashboard.html` | Aus der Excel generiertes, filterbares Dashboard (Abteilung × Use Case Art) |
| `ki-inbox-processor.skill` | Claude-Skill, der neue Inbox-Einträge prüft, klassifiziert und in die Haupttabelle übernimmt |
| `CLAUDE.md` | Arbeitsanweisungen für Claude in diesem Repo (Ziele, Workflow, No-Gos) |

## Struktur der Excel

- **Use Cases** – Nr., Problem, KI-Lösung, Abteilung, Use Case Art, Quelle Workshop
- **Cluster-Übersicht** – Definition der Abteilungen und Use-Case-Arten
- **Auswertung** – Anzahl Use Cases je Cluster
- **Inbox** – Sammelplatz für neue, noch unverarbeitete Use Cases

## Workflow

1. Neue Use-Case-Idee → in Sheet **Inbox** eintragen (oder Claude dazu sagen)
2. Claude (über den Skill) prüft Duplikate, formuliert sauber, ordnet **Abteilung** + **Use Case Art** zu
3. Nur **neue** Use Cases werden in **Use Cases** ergänzt; Duplikate werden gemeldet
4. Dashboard wird neu gerendert

## Status

Prototyp / Exploration. Strukturen dürfen sich noch ändern.
