# Connections – Revit Electrical Connection Tool

A Revit add-in for quickly connecting electrical elements (security cameras, fire alarm devices, data outlets, power equipment, etc.) to distribution panels.

## Features

- **Panel Selection** – searchable ComboBox listing all electrical equipment panels in the active model
- **Existing Connections Count** – displays the number of circuits already connected to the selected panel
- **Connection Limit** – optional per-panel limit to prevent over-provisioning (0 = unlimited)
- **Max Cable Length** – set a maximum allowed cable length in metres; elements that exceed it are highlighted orange and a warning popup is shown (0 = off)
- **Clear Warnings** – button that clears all orange highlights from the active view (appears only when highlights are present)
- **Connect Individually** – creates a separate circuit for each selected element
- **Connect Combined** – creates a single shared circuit for all selected elements
- **Auto System Type Detection** – automatically detects the correct `ElectricalSystemType` (Power, Security, Data, Fire Alarm, etc.) from each element's connector
- **Circuit Parameter Writing** – optionally writes a parameter name/value pair to each created circuit
- **Session Counter** – tracks how many connections have been made in the current session, with a clear button
- **Dark / Light Theme** – toggle between dark and light UI themes
- **State Persistence** – remembers panel selection, parameter values, connection mode, limits, window position, and theme across sessions

## Cable Length Warning

When **Max Cable Length (m)** is set to a value greater than 0:

- After each circuit is committed, the plugin reads the circuit's `Length` parameter (the value Revit shows in electrical schedules)
- If the length exceeds the limit, the connected elements are highlighted **orange** in the active view
- A styled warning popup appears listing each offending element and its actual length
- Use the **Clear Warnings** button to remove all orange highlights

## Supported Revit Versions

| Target Framework | Revit Version |
|---|---|
| .NET Framework 4.8 | Revit 2024 |
| .NET 8.0 (Windows) | Revit 2026 |

## Build

```
dotnet build Connections.sln
```

The build produces:
- `bin/Debug/net48/Connections.dll` – for Revit 2024
- `bin/Debug/net8.0-windows/Connections.dll` – for Revit 2026

## Installation

Copy the built `Connections.dll` to your Revit add-ins folder:

- **Revit 2024**: `%AppData%\Autodesk\Revit\Addins\2024\`
- **Revit 2026**: `%AppData%\Autodesk\Revit\Addins\2026\`

The add-in registers itself via the `ricaun.Revit.UI` AppLoader attribute and creates a **Connections** button under the **RK Tools** ribbon tab.

## Usage

1. Click the **Connections** button in the RK Tools ribbon tab
2. Select a **Panel** from the dropdown
3. (Optional) Set a **Connection Limit** and/or **Max Cable Length (m)**
4. (Optional) Enter a **Circuit Parameter Name** and **Value** to write to created circuits
5. Choose **Connect Individually** or **Connect Combined**
6. Click **Select Elements & Connect**, then pick elements in the Revit view
7. Results and any cable length warnings are displayed in the output area
8. If warnings are shown, use **Clear Warnings** to remove orange highlights from the view

## Dependencies

- [ricaun.Revit.UI](https://github.com/ricaun-io/ricaun.Revit.UI) – Ribbon UI helpers and AppLoader
- [MaterialDesignInXAML](https://github.com/MaterialDesignInXAML/MaterialDesignInXamlToolkit) – WPF theme and icons
- [Newtonsoft.Json](https://www.nuget.org/packages/Newtonsoft.Json) – Settings serialization
- [Costura.Fody](https://github.com/Fody/Costura) – Assembly embedding

## License

This project is proprietary software by RK Tools.
