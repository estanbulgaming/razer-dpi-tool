# Razer DPI Tool

Simple tool to change DPI on Razer mice without Synapse.

## Features

- Change DPI without Razer Synapse installed
- Preset DPI buttons (400, 800, 1200, 1600, 3200)
- Custom DPI input (100-25600)
- Supports 90+ Razer mice

## Usage

1. Download `RazerDPI-Python.exe` from [Releases](https://github.com/estanbulgaming/razer-dpi-tool/releases)
2. Run the exe
3. Click a DPI preset or enter custom value

## Supported Mice

DeathAdder, Viper, Basilisk, Naga, Mamba, Cobra, Orochi, Lancehead, Abyssus and more.

## Build

```bash
pip install hidapi pyinstaller
pyinstaller --onefile --windowed --name RazerDPI razer_dpi.py
```
