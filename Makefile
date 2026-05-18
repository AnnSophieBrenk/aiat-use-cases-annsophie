.PHONY: update render setup help

EXCEL    := KI_UseCases_Uebersicht.xlsx
DASHBOARD := KI_UseCases_Dashboard.html
URL      := https://annsophiebrenk.github.io/aiat-use-cases-annsophie/

help:
	@echo ""
	@echo "  Verfügbare Befehle:"
	@echo ""
	@echo "    make update   Dashboard rendern, committen und pushen → Live-Seite aktualisieren"
	@echo "    make render   Nur das Dashboard lokal rendern (kein Git)"
	@echo "    make setup    Einmalig: Python-Abhängigkeit installieren (openpyxl)"
	@echo ""

setup:
	@python3 -m pip install --user --quiet openpyxl
	@echo "Setup fertig."

render:
	@python3 scripts/rebuild_dashboard.py $(EXCEL)

update: render
	@git add $(EXCEL) $(DASHBOARD)
	@if git diff --cached --quiet; then \
		echo "Keine Änderungen – nichts zu committen."; \
	else \
		git commit -m "Dashboard aktualisiert ($$(date +%Y-%m-%d))"; \
		git push; \
		echo ""; \
		echo "→ Live in ~30s unter: $(URL)"; \
	fi
