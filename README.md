# Excel-Funktionen – Lernpfad

Interaktiver Lernpfad „Excel-Funktionen" für Kaufleute für Büromanagement (BBZ Rendsburg-Eckernförde),
gebaut mit dem [Lernpfad-Baukasten](https://github.com/dclausen01/Lernpfad-baukasten).

- **Webversion:** wird automatisch per GitHub Pages veröffentlicht (siehe Workflow
  [`pages.yml`](.github/workflows/pages.yml)); Link erscheint nach dem ersten erfolgreichen Lauf unter
  „Settings → Pages" bzw. im Tab „Actions".
- **Details zu Auftrag, Struktur und offenen Punkten:** [uebergabe.md](uebergabe.md)
- **Planung, offene Arbeitspakete und Ideensammlung:** [planung.md](planung.md)
- **Ordner:**
  - `module/` – Modultexte (Markdown, 23 Module)
  - `material/` – Übungsdateien (`.xlsx`)
  - `lehrkraft/loesungen/` – Musterlösungen (nicht für Web/Moodle)
  - `portfolio/` – Kann-Liste (Lernportfolio) mit Verknüpfung zu den Modulen (Prototyp)
  - `werkzeuge/` – Generatoren der Übungsdateien (siehe `werkzeuge/README.md`)
  - `ausgabe/web/` – gebaute Webversion (**nie von Hand ändern**, wird vom Baukasten erzeugt). Aktuell **nicht enthalten**: Der alte Stand liegt im Branch `lernpfad-v1` des alten Repos; die neue Webversion entsteht mit dem nächsten Build
  - `kurs.yaml`, `medien.yaml`, `quellen.yaml` – Kursbeschreibung, Medienlinks, Quellenbelege

Alle 23 Module stehen aktuell auf `review: entwurf` (siehe `uebergabe.md`, Abschnitt 6 „Offene Punkte").
