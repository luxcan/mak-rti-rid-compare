# RID Comparison App

The RID Comparison App is a standalone console application designed to compile and compare multiple [MAK](https://www.mak.com/) RTI rid files (`.mtl`) into a single Excel report. This tool reads the configuration data from these files and generates a comprehensive Excel spreadsheet that lists the configuration names along with their corresponding values from each file.

## Prerequisites

- **Operating System**: Windows 10 or later.
- **.NET Runtime**: The app is self-contained, so you do not need to install .NET separately.
- **Excel Viewer**: To view the generated Excel file, you will need Microsoft Excel or another compatible viewer.

## Setup

### 1. Download the App
   - Download the standalone app package [here](https://github.com/luxcan/mak-rti-rid-compare/blob/main/MAKRtiRidCompare.zip).

### 2. Extract the Package
   - Extract the contents of the `RidComparisonApp.zip` file to a directory of your choice (e.g., `C:\RidComparisonApp`).

### 3. Prepare the RID Files
   - Create a folder named `rids` within the extracted directory (e.g., `C:\RidComparisonApp\rids`).
   - Place all your `.mtl` files that you want to compare into the `rids` folder. Ensure the files are named appropriately for easy identification.

## Running the App

### 1. Run the Application
   - Navigate to the directory where you extracted the app (e.g., `C:\RidComparisonApp`).
   - Double-click on `RidComparisonApp.exe` to run the application.

### 2. App Execution
   - The app will process all `.mtl` files in the `rids` directory, compare their configuration values, and generate an Excel file named `RidData.xlsx` in the same directory.

## Viewing the Results

### 1. Locate the Excel File
   - After the app has finished running, you will find the `RidData.xlsx` file in the main application directory (e.g., `C:\RidComparisonApp\RidData.xlsx`).

### 2. Open the Excel File
   - Open the `RidData.xlsx` file using Microsoft Excel or any compatible Excel viewer to review the comparison results.

## Example Directory Structure
RidComparisonApp (folder)
├── RidComparisonApp.exe
├── rids
│ ├── file1.mtl
│ ├── file2.mtl
│ └── file3.mtl
└── RidData.xlsx
