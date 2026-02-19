# HDsEMG Protokoll JSON-Format

## Übersicht

Das HDsEMG Versuchsreihe Protokoll wird automatisch in zwei Formaten gespeichert:
1. **Textformat** (`.txt`) - Menschenlesbar, für Ausdrucke und schnelle Übersichten
2. **JSON-Format** (`.json`) - Maschinenlesbar, für Datenanalyse und Weiterverarbeitung

Beide Dateien werden automatisch nach Abschluss der Session im Ordner `protokolle/` gespeichert.

## Dateiname-Konvention

```
{PID}_{TIMESTAMP}_protokoll.txt
{PID}_{TIMESTAMP}_protokoll.json
```

Beispiel:
```
PID1_20251216_151504_protokoll.txt
PID1_20251216_151504_protokoll.json
```

## JSON-Struktur

### Root-Level

```json
{
  "protocol_version": "1.0",
  "session": { ... },
  "metadata": { ... },
  "declaration": { ... },
  "output_directory": "...",
  "steps": [ ... ]
}
```

### Session-Informationen

```json
"session": {
  "started_at": "2025-12-16T15:15:04",
  "ended_at": "2025-12-16T17:24:24",
  "duration_seconds": 7760.0,
  "duration_formatted": "02:09:20",
  "timestamp": "20251216_151504"
}
```

### Metadaten

Enthält alle Session-weiten Metadaten wie PID, Messdatum, Session-Typ, etc.

```json
"metadata": {
  "pid": "1",
  "mess_tag": "16.12.2025",
  "session_type": "Day 2 - Intervention",
  "randomization": "CON first",
  "doms_score": "",
  "notes_general": "..."
}
```

### Declaration

Informationen über die verwendete Konfigurationsdatei:

```json
"declaration": {
  "title": "HDsEMG Intervention Session",
  "description": "Day 2 Intervention Protocol",
  "declaration_file": "/path/to/config/intervention.json"
}
```

### Steps (Schritte)

Array mit allen dokumentierten Schritten:

```json
"steps": [
  {
    "step_number": 1,
    "step_id": "arrival_setup",
    "title": "Ankunft & Setup",
    "description": "Vorbereitung und Einrichtung",
    "started_at": "2025-12-16T15:15:04",
    "completed_at": "2025-12-16T15:43:25",
    "duration_seconds": 1701.0,
    "duration_formatted": "00:28:21",
    "expected_duration_seconds": 1350,
    "expected_duration_formatted": "00:22:30",
    "fields": { ... },
    "repeated_measurements": { ... },
    "notes": "...",
    "otbiolab_files": { ... }
  }
]
```

### Normale Felder

Einfache Felder mit einem einzelnen Wert:

```json
"fields": {
  "body_height": {
    "label": "Körpergröße (cm)",
    "value": 180.0,
    "type": "float",
    "otbiolab_files": []
  },
  "dominant_leg": {
    "label": "Dominantes Bein",
    "value": "rechts",
    "type": "choice",
    "otbiolab_files": []
  }
}
```

### Wiederholbare Messungen

Felder, die mehrfach durchgeführt werden können:

```json
"repeated_measurements": {
  "trapezoid_tracking": {
    "label": "Trapezoid Tracking",
    "attempts": [
      {
        "attempt_number": 1,
        "fields": {
          "tracking_quality": {
            "label": "Tracking-Qualität",
            "value": "Sehr gut",
            "type": "choice"
          },
          "notes": {
            "label": "Notizen",
            "value": "",
            "type": "multiline"
          }
        },
        "otbiolab_file": "/path/to/output/PID1_20251216_151504_bl1_trap1.otb4"
      },
      {
        "attempt_number": 2,
        "fields": { ... },
        "otbiolab_file": "/path/to/output/PID1_20251216_151504_bl1_trap2.otb4"
      }
    ]
  }
}
```

### OTBioLab-Dateien

Zwei Ebenen von OTBioLab-Dateien:

```json
"otbiolab_files": {
  "step_level": [
    "/path/to/file1.otb4",
    "/path/to/file2.otb4"
  ],
  "field_level": {
    "field_id": [
      "/path/to/field_file1.otb4",
      "/path/to/field_file2.otb4"
    ]
  }
}
```

## Verwendung des JSON-Formats

### Python-Beispiel

```python
import json
from pathlib import Path

# JSON-Protokoll laden
with open("PID1_20251216_151504_protokoll.json", "r", encoding="utf-8") as f:
    protocol = json.load(f)

# Zugriff auf Session-Daten
print(f"Session Start: {protocol['session']['started_at']}")
print(f"Dauer: {protocol['session']['duration_formatted']}")

# Iteration über alle Schritte
for step in protocol['steps']:
    print(f"Schritt {step['step_number']}: {step['title']}")
    print(f"  Dauer: {step['duration_formatted']}")

    # Zugriff auf normale Felder
    for field_id, field_data in step['fields'].items():
        print(f"  {field_data['label']}: {field_data['value']}")

    # Zugriff auf wiederholbare Messungen
    for measurement_id, measurement_data in step['repeated_measurements'].items():
        print(f"  {measurement_data['label']}:")
        for attempt in measurement_data['attempts']:
            print(f"    Versuch {attempt['attempt_number']}:")
            if attempt['otbiolab_file']:
                print(f"      Datei: {attempt['otbiolab_file']}")
            for field_id, field_data in attempt['fields'].items():
                print(f"      {field_data['label']}: {field_data['value']}")
```

### R-Beispiel

```r
library(jsonlite)

# JSON-Protokoll laden
protocol <- fromJSON("PID1_20251216_151504_protokoll.json")

# Session-Informationen
cat("Session Start:", protocol$session$started_at, "\n")
cat("Dauer:", protocol$session$duration_formatted, "\n")

# Schritte als Data Frame
steps_df <- as.data.frame(protocol$steps)
print(steps_df[, c("step_number", "title", "duration_formatted")])

# Wiederholbare Messungen extrahieren
for (step_idx in seq_along(protocol$steps)) {
  step <- protocol$steps[[step_idx]]
  if (length(step$repeated_measurements) > 0) {
    for (measurement_name in names(step$repeated_measurements)) {
      measurement <- step$repeated_measurements[[measurement_name]]
      cat(sprintf("Schritt %d - %s:\n", step$step_number, measurement$label))
      for (attempt_idx in seq_along(measurement$attempts)) {
        attempt <- measurement$attempts[[attempt_idx]]
        cat(sprintf("  Versuch %d: %s\n",
                    attempt$attempt_number,
                    attempt$otbiolab_file))
      }
    }
  }
}
```

## Vorteile des JSON-Formats

1. **Maschinell lesbar**: Einfache Weiterverarbeitung in Python, R, MATLAB, etc.
2. **Strukturiert**: Klare Hierarchie und Datentypen
3. **Vollständig**: Alle Daten aus der Session sind enthalten
4. **Versioniert**: `protocol_version` ermöglicht zukünftige Erweiterungen
5. **Standardformat**: JSON wird von fast allen Programmiersprachen unterstützt

## Änderungshistorie

- **Version 1.0** (2025-01-13): Initiale Version mit vollständiger Protokoll-Struktur
