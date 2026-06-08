# 3D ROI Editor

## Overview

Mech-Eye 3D ROI Editor is a PyQt6-based desktop application designed for configuring and managing ROI (Region of Interest) parameters used in Mech-Eye 3D vision systems. The tool provides an easy-to-use interface for editing ROI JSON files, synchronizing data with Excel and PLC systems, managing part dimensions, and automating ROI position updates in industrial robotics environments.

The application is designed for industrial automation environments where robot vision systems, PLCs, and part dimension databases must work together. It allows users to load ROI configurations, modify ROI parameters, synchronize data with Excel files, and automatically update ROI positions from PLC-generated data.

---

## Features

### JSON ROI Editing

* Load and edit ROI configuration files.
* Modify:

  * ROI Half Lengths (X, Y, Z)
  * ROI Center Pose (X, Y, Z, QX, QY, QZ, QW)
* Save changes directly to the JSON file.

### Excel Integration

Supports two different Excel modes:

#### PART Mode

* Import part dimension data from Excel.
* Display:

  * Length
  * Width
  * Height
* Automatically populate part information.
* Add new parts directly into the Excel database.

#### PLC Mode

* Read ROI Center Pose values directly from a PLC-generated Excel file.
* Automatically update ROI parameters in real time.
* Synchronize updated values back into the JSON configuration.

### Offset Calculation

* Calculate ROI offsets based on selected part dimensions.
* Maintain offset history.
* Automatically store calculated offsets in the JSON file.

### Manual ROI Adjustment

* Manually set ROI Center Pose values.
* Increment or decrement numerical values using built-in controls.

### Reset Function

* Reset ROI Center Pose to default values.
* Password-protected safety mechanism.

### Real-Time Monitoring

* Automatic Excel file monitoring using timers.
* Automatic JSON updates when PLC values change.

---

## Application Workflow

### 1. Load a JSON File

Open an ROI configuration JSON file containing:

```json
{
    "roi3d_in_camera_coord": {
        "half_lengths": [],
        "roi_center_pose": []
    },
    "roi3d_in_robot_coord": {
        "half_lengths": [],
        "roi_center_pose": []
    }
}
```

### 2. Load PART Excel File

Import a part database Excel file containing:

| ID | Part Code | Length | Width | Height |
| -- | --------- | ------ | ----- | ------ |

Select a part from the dropdown menu to view its dimensions.

### 3. Calculate Offset

Use the OFFSET button to:

* Read the selected part width.
* Calculate a new ROI Y coordinate.
* Store the result in the offset history.
* Save automatically to the JSON file.

### 4. PLC Synchronization (Optional)

Load a PLC Excel file containing:

| X | Y | Z | QX | QY | QZ | QW |
| - | - | - | -- | -- | -- | -- |

Switch to PLC Mode.

The application will:

* Continuously monitor the Excel file.
* Read pose values from row 2.
* Update ROI Center Pose automatically.
* Write changes back to the JSON file.

---

## Use Cases

### Robot Vision Systems

Adjust camera ROI regions according to part dimensions.

### Industrial Automation

Synchronize robot positioning data with PLC-generated information.

### Production Lines

Manage multiple part definitions and automatically update inspection regions.

### Quality Control Systems

Configure and maintain 3D inspection areas for machine vision applications.

---

## Technologies Used

* Python 3
* PyQt6
* OpenPyXL
* JSON
* Excel (.xlsx)

---

## Installation

Install required dependencies:

```bash
pip install PyQt6 openpyxl
```

Run the application:

```bash
python 3D_ROI_JSON.py
```

---

## Supported File Formats

### JSON

Used for ROI configuration storage.

### Excel (.xlsx)

#### PART Excel Format

| ID | Part Code | Length | Width | Height |
| -- | --------- | ------ | ----- | ------ |

#### PLC Excel Format

| X | Y | Z | QX | QY | QZ | QW |
| - | - | - | -- | -- | -- | -- |

---

## Notes

* The JSON file must be loaded before Excel files.
* PART and PLC modes can be switched using the integrated toggle switch.
* PLC mode automatically updates ROI values in real time.
* Offset calculations are saved automatically.
* Excel files should not be opened in another application while being updated.

---

## Author

Developed for industrial robotics, machine vision, and automation workflows requiring dynamic ROI configuration management.
