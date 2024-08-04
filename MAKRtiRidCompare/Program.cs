using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

var workingDir = Directory.GetCurrentDirectory();
var ridDir = Path.Combine(workingDir, "rids");
if (!Directory.Exists(ridDir)) {
    Console.WriteLine($"Rids folder not exists ({ridDir}).");
    return;
}

var ridFiles = Directory.GetFiles(ridDir, "*.mtl");
if (ridFiles.Length == 0) {
    // Exit application, don't waste process.
    Console.WriteLine("The 'rids' folder is empty.");
    return;
}

// Sort the files in alphabetical order
Array.Sort(ridFiles, StringComparer.OrdinalIgnoreCase);

var configData = new Dictionary<string, Dictionary<string, string>>();
var splitCondition = new[] { ' ', '\t', '(', ')' };

foreach (var file in ridFiles) {
    var fileName = Path.GetFileName(file);

    foreach (var line in File.ReadLines(file)) {
        // Ignore commented lines
        if (line.Trim().StartsWith(";;")) {
            continue;
        }

        // Process lines with the format (setqb <key> <value>)
        if (line.Trim().StartsWith("(setqb")) {
            var parts = line.Split(splitCondition, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 3) {
                var key = parts[1];
                var value = parts[2];

                if (!configData.TryGetValue(key, out var fileValues)) {
                    fileValues = new Dictionary<string, string>();
                    configData[key] = fileValues;
                }

                fileValues[fileName] = value;
            }
        }
    }
}

// Create an Excel file using NPOI
IWorkbook workbook = new XSSFWorkbook();
ISheet sheet = workbook.CreateSheet("Config Data");

// Write the header row
IRow headerRow = sheet.CreateRow(0);
headerRow.CreateCell(0).SetCellValue("Configuration Name");
for (int colIndex = 0; colIndex < ridFiles.Length; colIndex++) {
    headerRow.CreateCell(colIndex + 1).SetCellValue(Path.GetFileName(ridFiles[colIndex]));
}

// Write the data rows
int rowIndex = 1;
foreach (var key in configData.Keys) {
    IRow row = sheet.CreateRow(rowIndex);
    row.CreateCell(0).SetCellValue(key);

    for (int colIndex = 0; colIndex < ridFiles.Length; colIndex++) {
        var fileName = Path.GetFileName(ridFiles[colIndex]);
        row.CreateCell(colIndex + 1).SetCellValue(configData[key].TryGetValue(fileName, out var value) ? value : string.Empty);
    }

    rowIndex++;
}

// Auto-size the columns
for (int colIndex = 0; colIndex <= ridFiles.Length; colIndex++) {
    sheet.AutoSizeColumn(colIndex);
}

// Freeze the first column
sheet.CreateFreezePane(1, 0);

// Save the Excel file
var excelFilePath = Path.Combine(workingDir, "RidData.xlsx");
try {
    using (var fileStream = new FileStream(excelFilePath, FileMode.Create, FileAccess.Write)) {
        workbook.Write(fileStream);
    }
} catch (Exception ex) {
    Console.WriteLine("Unable to save the excel file (" + workingDir + ").");
    Console.WriteLine(ex.ToString());
    return;
}

Console.WriteLine("Excel file created successfully.");