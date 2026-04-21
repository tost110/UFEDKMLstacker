# UFEDKMLstacker
UFEDKMLstacker is a Python script designed to merge and visualize KML files exported from Cellebrite UFED [^1] using Plotly [^2]. This tool allows the stacking of multiple KML files to demonstrate overlapping locations of different data sources, each represented by a distinct color. By integrating this functionality, UFEDKMLstacker aids in forensic analysis and visualization, offering valuable insights into spatial relationships between different entities. Inspired by [UFEDMapper](https://github.com/ot2i7ba/UFEDMapper) and [UFEDKMLmerge](https://github.com/ot2i7ba/UFEDKMLmerge), this script has been customized to meet specific forensic needs, enabling professionals to streamline their workflow. Sharing this tool aims to benefit others in the forensic community.

> [!NOTE]
> This script is specifically tailored for KML files exported from Cellebrite UFED. Its compatibility or performance with KML files from other sources has not been tested.

> [!WARNING]
Please note that this script is currently under development, and I cannot provide a 100% guarantee that it operates in a forensically sound manner. It is tailored to meet specific needs at this stage. Use it with caution, especially in environments where forensic integrity is critical.

## Table of Contents
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
   - [Usage](#usage)
   - [PyInstaller](#pyinstaller)
   - [Releases](#releases)
- [Example](#example)
- [Changes](#changes)
- [License](#license)

# Features
- Synchronous processing for enhanced compatibility and reliability when handling multiple KML files.
- Merging of up to 10 KML files into a single visualization, with a limit to ensure performance and manageability.
- Unique color coding for each KML file to facilitate easy differentiation in the merged visualization.
- Comprehensive logging: Logs activities in detail to a file, with warnings and errors also displayed in the console.
- Metadata extraction: Includes file creation and modification times, along with SHA-256 hash for integrity verification.
- Saves results and maps in multiple formats, including KML, Excel, CSV, and HTML for thorough analysis and documentation.
- Improved user interaction and error handling, with clear prompts and robust validation to prevent and manage errors effectively.

# Requirements
- Python 3.7 or higher
pandas>=2.0.0,<3.0.0
lxml>=4.9.0,<6.0.0
arrow>=1.0.0,<2.0.0
folium>=0.14.0,<1.0.0


# Installation
1. **Clone the repository**

   ```sh
   git clone https://github.com/ot2i7ba/UFEDKMLstacker.git
   cd UFEDKMLstacker
   ```

2. **To install the required dependencies, use the following command**
   ```sh
   pip install -r requirements.txt
   ```

## Usage
1. **Place your KML files in the same directory as the script**

2. **Run the script**
   ```sh
   python UFEDKMLstacker.py
   ```

## Follow the Prompts
- **KML File Selection**<br>The script will search the directory for KML files and present a numbered list for selection.
- **Remark for Each File**<br>Enter a remark for each selected file to identify it in the visualization.
- **Color Assignment**<br>Colors are automatically assigned to each file for differentiation in the visualization.

## Interactive Map
The resulting map visualizes the data points from each KML file in distinct colors. The user can interact with the map, hover over points to see details, and filter based on the provided remarks.

## PyInstaller
To compile the UFEDKMLmerge script into a standalone executable, you can use PyInstaller. Follow the steps below:

1. Install PyInstaller (if not already installed):
   ```bash
   pip install pyinstaller
   ```

2. Compile the script using the following command:
   ```bash
   pyinstaller --onefile --name UFEDKMLstacker --icon=ufedkmlstacker.ico UFEDKMLstacker.py
   ```

- `--onefile`: Create a single executable file.
- `--name UFEDKMLmerge`: Name the executable UFEDMapper.
- `--icon=ufedkmlstacker.ico`: Use ufedkmlmerge.ico as the icon for the executable.

**Running the executable**: After compilation, you can run the executable found in the dist directory.

## Releases
A compiled and 7zip-packed version of UFEDKMLstacker for Windows is available as a release. You can download it from the **[Releases](https://github.com/ot2i7ba/UFEDKMLstacker/releases)** section on GitHub. This version includes all necessary dependencies and can be run without requiring Python to be installed on your system.

Please note that console logging (log_to_console=False) is set to False in compiled version. This means that log outputs will be written to a file rather than displayed in the console.

> [!IMPORTANT]
> An internet connection is required to display the maps in the generated HTML files. This is because the maps are rendered using Plotly, which relies on online resources to load the map tiles and other visualization components. Rest assured, no information from your local system is sent to Plotly during this process. The map data is sourced from OpenStreetMap [^3], ensuring compliance with GDPR regulations.

# Example

## The script lists available KML files and prompts for selection**
   ```sh
   Available KML files:
   1. Example1.kml
   2. Example2.kml
   3. Example3.kml
   e. Exit
   
   Enter file numbers to merge (e.g., 1, 2, 5) or 'e' to exit:
   ```

## After selecting files, the script prompts for remarks
   ```sh
   Enter a remark for the file Example1.kml: Alpha
   Enter a remark for the file Example2.kml: Bravo
   Enter a remark for the file Example3.kml: Charlie

   Processing Example1.kml ... done!
   Processing Example2.kml ... done!
   Processing Example3.kml ... done!
   ```

## The interactive map is then generated and saved
   ```sh
   Statistics saved in C:\YOUR\PATH\TO\UFEDKMLstacker\KML_Statistics.xlsx.
   Statistics saved in C:\YOUR\PATH\TO\UFEDKMLstacker\KML_Statistics.csv.

   Process completed successfully. Details are available in the log, Excel file, and CSV file.
   ```

___

# Changes

# UFEDKMLstacker v0.0.4

> Optimised continuation of [v0.0.3](https://github.com/ot2i7ba/UFEDKMLstacker) by [ot2i7ba](https://github.com/ot2i7ba).

---

## What's New

### 🌐 Language Selection (English / German)

At startup the tool now asks for a preferred language:

```
Language / Sprache ([EN] = English / DE = Deutsch):
```

All console output, prompts, map layer names, popups and the speed-band legend in the interactive map respond to this choice. The default is English when pressing Enter.

---

### 🗺️ Interactive Map (Folium / Leaflet)

The map output was completely rewritten using **Folium (Leaflet.js)**, replacing the previous Plotly implementation which did not support reliable zoom in saved HTML files.

- Mouse-wheel zoom, pinch-zoom, +/− buttons
- Fullscreen button and mini-map overview
- Scale bar
- Three selectable base maps: **CartoDB Positron** (default), **CartoDB Dark Matter**, **OpenStreetMap (mirror)** — all tile-server friendly, no 403 errors
- All layers individually toggleable via the layer control
- Auto-fit viewport to all points on load

---

### 🚗 Speed Analysis

Automatic speed calculation between consecutive location points using GPS timestamps and the **Haversine formula**:

- Distance (km), duration (s) and speed (km/h) per segment
- Colour-coded polylines as a toggleable map layer:

  | Band | Default colour |
  |---|---|
  | < 30 km/h | Grey |
  | 30–50 km/h | Green |
  | 50–70 km/h | Blue |
  | 70–100 km/h | Yellow |
  | 100–130 km/h | Orange |
  | 130–150 km/h | Dark orange |
  | > 150 km/h | Red |

- Segments with time gaps > 4 hours are skipped automatically
- Click popup per segment: from/to point, timestamps, distance, duration, speed and band label
- **Speed-band legend** rendered as a draggable overlay on the map

#### Configurable Band Colours

Before processing begins the tool optionally asks whether to adjust band colours. Colours can be entered as names (`red`, `blue`, `orange`, …) or as hex codes (`#FF4500`). Settings are saved to `speed_colors.json` and loaded automatically on the next run.

---

### 📊 Extended Statistics Export

| File | Contents |
|---|---|
| `KML_Statistics.xlsx` | Summary sheet + new **Speed Segments** and **Speed Summary** sheets |
| `KML_Statistics.csv` | New columns: `speed_segments`, `max_speed_kmh` |
| `KML_Speed_Segments.csv` | **New** — all calculated speed segments sorted by speed |
| `speed_colors.json` | **New** — persisted colour configuration for speed bands |

---

## Bug Fixes (vs. v0.0.3)

| # | Area | Fix |
|---|---|---|
| 1 | KML colour format | `define_styles()` used `ff` + `#RRGGBB`; correct KML format is `AABBGGRR`. Fixed via `_hex_to_kml_color()`. |
| 2 | Double SHA-256 | Hash was computed in `extract_file_metadata()` and again in `verify_file_integrity()`. Now computed once and passed through. |
| 3 | Duplicate `color_name_map` | Defined globally and again locally in `save_statistics_to_excel()` with differing keys. Merged into single `HEX_TO_NAME`. |
| 4 | `create_interactive_map()` crashes | Missing `<name>` / `<description>` tags in Placemarks caused `AttributeError`. All element accesses now None-safe. |

---

## Code Quality Improvements (vs. v0.0.3)

- `pathlib.Path` used consistently for all path operations
- `COLOR_PALETTE` as the single source of truth for colour assignment
- `MAX_REMARK_LENGTH` and `MAX_SELECTION_LENGTH` constants for input validation
- `validate_selection()` deduplicates repeated file numbers
- SHA-256 hash read buffer increased from 4 KB to 64 KB
- Full type annotations (`typing`) on all functions
- `json` imported at module level (was imported locally in three places)

---

## New Dependency

```
folium>=0.14.0,<1.0.0
```

Install / update with:

```bash
pip install -r requirements.txt
```

---

## Notes

- **Tile servers and the interactive map require an internet connection.** All other functionality works fully offline.
- Speed calculations are approximations based on GPS coordinates and timestamps extracted from UFED exports. Results must be verified against source data.
> [!IMPORTANT]
- Still testing because 0.0.4 is created via vibe coding.

---

## Changes in 0.0.3

- **Global Color Mapping Definition**<br>The color mapping (color_name_map) has been moved to a global scope, making it accessible throughout the script without needing to redefine it locally within functions.
- **Enhanced Logging Configuration**<br>The configure_logging() function now includes an optional log_to_console parameter, allowing users to control whether log messages are shown in the console. By default, this parameter is set to True, showing only ERROR level messages in the console while logging all levels to the log file.
- **Extended Timestamp Parsing**<br>Additional regular expression patterns have been added to the parse_timestamp() function to handle a wider range of timestamp formats, improving the robustness and versatility of timestamp parsing.
- **Corrected Statistics Calculation**<br>The calculation of statistics has been corrected to ensure all values are counted accurately. This includes proper counting of total points, points with timestamps, and points without timestamps, enhancing the reliability of the reported data.
- **New Statistic Metric**<br>A new metric, points_without_timestamps, has been added to the statistics. This metric tracks the number of geopoints that do not have associated timestamps, providing more detailed insights into the data.
- **Spinner Animation During Processing**<br>A spinner animation has been introduced to provide visual feedback during the processing of KML files. The spinner shows the current file being processed and stops once processing is complete, improving user experience by indicating ongoing activity.

## Changes in 0.0.2

- **Synchronous Processing**<br>The script has been updated from asynchronous to synchronous processing to enhance compatibility and reliability. The use of asyncio and aiofiles has been removed.
- **File Selection Limitation**<br>A new feature limits the selection to a maximum of 10 KML files at a time. This constraint ensures that the script processes only a manageable number of files simultaneously.
- **Enhanced Error Handling**<br>Improved error handling mechanisms have been introduced to catch and log specific exceptions such as FileNotFoundError, ValueError, and XMLSyntaxError more effectively.
- **Input Validation**<br>Additional validation steps have been implemented to ensure that user inputs are correct, avoiding invalid characters or excessively long entries.
- **Extended Timestamp Processing**<br>The timestamp parsing has been extended to support more formats, ensuring that various timestamp formats can be recognized and parsed correctly.
- **Saving Statistics**<br>Separate Excel and CSV files (KML_Statistic.xlsx and KML_Statistic.csv) have been introduced to store detailed statistics about the processed KML files. This adds an extra layer of control and documentation, complementing the existing Merged_Colored files.
- **Logging Configuration**<br>Logging settings have been optimized to provide a logical separation between console and file logging. Console logging now only displays messages at the WARNING level and above, while detailed debug information is still recorded in the log file.

## Changes in 0.0.1
- Initial release.

___

> [!CAUTION]
> This project is based on the original repository by [ot2i7ba](https://github.com/ot2i7ba).

# License
This project is licensed under the **[MIT license](https://github.com/ot2i7ba/UFEDKMLstacker/blob/main/LICENSE)**, providing users with flexibility and freedom to use and modify the software according to their needs.

# Contributing
Contributions are welcome! Please fork the repository and submit a pull request for review.

# Disclaimer
This project is provided without warranties. Users are advised to review the accompanying license for more information on the terms of use and limitations of liability.

# Conclusion
This script has been tailored to fit my specific professional needs, and while it may seem like a small tool, it has a significant impact on my workflow. This script is designed to aid forensic professionals in visualizing and analyzing location data extracted from multiple sources. By automating the merging and visualization process, UFEDKMLstacker enhances efficiency, allowing users to focus on critical analysis. Greetings to my dear colleagues who avoid scripts like the plague and think that consoles and Bash are some sort of dark magic – the [compiled](https://github.com/ot2i7ba/UFEDKMLstacker/releases) version will spare you the console kung-fu and hopefully be a helpful tool for you as well. 😉

[^1]: [Cellebrite UFED](https://cellebrite.com/) (Universal Forensic Extraction Device) is a forensic tool to extract and analyze data from mobile devices.
[^2]: Thanks to the [Plotly](https://plotly.com/python/) team for their excellent visualization library, which made creating interactive maps a breeze.
[^3]: [OpenStreetMap](https://www.openstreetmap.org/) is a collaborative mapping project that provides freely accessible map data.


