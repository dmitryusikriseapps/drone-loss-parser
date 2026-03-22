SCRIPT = parse_drone_losses.py
DIST   = dist

.PHONY: run build-mac build-win clean

## Run the script locally (poetry venv)
run:
	poetry run python $(SCRIPT)

## Build standalone binary for macOS
build-mac:
	poetry run pyinstaller --onefile --console --name parse_drone_losses $(SCRIPT)
	@echo "\nBuild complete: $(DIST)/parse_drone_losses"

## Build standalone .exe for Windows (run this on a Windows machine)
build-win:
	poetry run pyinstaller --onefile --console --name parse_drone_losses $(SCRIPT)
	@echo "\nBuild complete: $(DIST)/parse_drone_losses.exe"

clean:
	rm -rf build dist *.spec