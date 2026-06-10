# Azure Function von Grund auf – am Beispiel ManagePermissions (TypeScript)

Diese Lernunterlage zeigt, **wie man eine Azure Function ohne KI-Unterstützung von Hand
aufbaut** – so, wie man es für ein SPFx-Web-Part längst kann. Als durchgehendes Beispiel
bauen wir die `ManagePermissions`-Lösung nach, diesmal
aber in **TypeScript/Node.js** mit **PnPjs** statt in C#.

Vom leeren Ordner bis zur laufenden Cloud-Function: Voraussetzungen zuerst, dann Schritt
für Schritt.

## Inhalt

| Datei | Zweck |
|---|---|
| [`anleitung.md`](anleitung.md) | **Die Anleitung zum Mitmachen.** Voraussetzungen + alle Schritte mit komplettem Code. Folge ihr von oben nach unten. |
| [`build-slides.py`](build-slides.py) | Reproduzierbarer Generator für die Vortragsfolien (python-pptx). |
| `slides.pptx` | Das generierte Foliendeck (aus `build-slides.py`). |

## Folien neu bauen

```powershell
pip install python-pptx          # einmalig
python build-slides.py           # erzeugt slides.pptx
python build-slides.py -o slides.new.pptx   # falls slides.pptx in PowerPoint geöffnet/gesperrt ist
```

## Verhältnis zur C#-Lösung

Die produktive `ManagePermissions`-Function im Schwester-Repo ist in **C# / .NET 8 isolated**
geschrieben. Diese Unterlage ist die **TypeScript-Variante als Lehrstück** – gleiche
Architektur, gleiche Sicherheits- und Berechtigungslogik, andere Sprache und andere
Werkzeuge. Was sich nicht ändert: die Muster (Identität, Token, least-privilege-Zugriff)
gelten unabhängig von der Sprache.
