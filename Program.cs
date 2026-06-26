using System.Text;
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

foreach (var file in ridFiles) {
    var fileName = Path.GetFileName(file);
    var text = File.ReadAllText(file);

    foreach (var statement in ExtractStatements(text)) {
        var tokens = Tokenize(statement);
        if (tokens.Count == 0) {
            continue;
        }

        string key;
        string value;

        if (tokens[0] == "setqb") {
            // (setqb <key> <value>) -> key is the config name, value its setting.
            if (tokens.Count < 3) {
                continue;
            }
            key = tokens[1];
            value = tokens[2];
        } else {
            // Any other directive, e.g. (RTI-addUpdateRate "high" 10.0).
            // Key on the directive plus its first argument so that repeated
            // directives stay distinct; the remaining arguments form the value.
            if (tokens.Count < 2) {
                continue;
            }
            key = tokens[0] + " " + tokens[1];
            value = string.Join(" ", tokens.Skip(2));
        }

        if (!configData.TryGetValue(key, out var fileValues)) {
            fileValues = new Dictionary<string, string>();
            configData[key] = fileValues;
        }

        fileValues[fileName] = value;
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

// Pink fill used to highlight configs whose values differ between files
var pinkStyle = (XSSFCellStyle)workbook.CreateCellStyle();
pinkStyle.SetFillForegroundColor(new XSSFColor(new byte[] { 255, 192, 203 }));
pinkStyle.FillPattern = FillPattern.SolidForeground;

// Write the data rows
int rowIndex = 1;
foreach (var key in configData.Keys) {
    IRow row = sheet.CreateRow(rowIndex);
    row.CreateCell(0).SetCellValue(key);

    // Collect this config's value from each file (empty when not defined).
    var rowValues = new string[ridFiles.Length];
    for (int colIndex = 0; colIndex < ridFiles.Length; colIndex++) {
        var fileName = Path.GetFileName(ridFiles[colIndex]);
        rowValues[colIndex] = configData[key].TryGetValue(fileName, out var value) ? value : string.Empty;
    }

    // A row "differs" when its values are not all identical across the files.
    // A missing value (empty string) also counts as a difference.
    bool hasDifference = rowValues.Distinct().Count() > 1;

    for (int colIndex = 0; colIndex < ridFiles.Length; colIndex++) {
        var cell = row.CreateCell(colIndex + 1);
        cell.SetCellValue(rowValues[colIndex]);
        if (hasDifference) {
            cell.CellStyle = pinkStyle;
        }
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

// --- Parsing helpers ---

// Splits RID file text into top-level parenthesised statements. Statements that
// span multiple physical lines are joined, and ;; comments are stripped.
static List<string> ExtractStatements(string text) {
    var statements = new List<string>();
    var sb = new StringBuilder();
    int depth = 0;
    bool inString = false;

    for (int i = 0; i < text.Length; i++) {
        char c = text[i];

        if (inString) {
            sb.Append(c);
            if (c == '"') {
                inString = false;
            }
            continue;
        }

        // ;; starts a comment that runs to the end of the line.
        if (c == ';' && i + 1 < text.Length && text[i + 1] == ';') {
            while (i < text.Length && text[i] != '\n') {
                i++;
            }
            i--; // let the loop's increment land on the newline (a token separator)
            continue;
        }

        if (c == '"') {
            inString = true;
            if (depth > 0) {
                sb.Append(c);
            }
            continue;
        }

        if (c == '(') {
            depth++;
            if (depth == 1) {
                sb.Clear(); // start a fresh statement, dropping the outer paren
            } else {
                sb.Append(c);
            }
            continue;
        }

        if (c == ')') {
            if (depth == 1) {
                statements.Add(sb.ToString());
                sb.Clear();
                depth = 0;
            } else if (depth > 1) {
                sb.Append(c);
                depth--;
            }
            // A stray ')' at depth 0 is ignored.
            continue;
        }

        if (depth > 0) {
            sb.Append(c);
        }
    }

    return statements;
}

// Splits a statement into tokens, keeping a "quoted string" (with its quotes)
// as a single token.
static List<string> Tokenize(string statement) {
    var tokens = new List<string>();
    var sb = new StringBuilder();
    bool inString = false;

    foreach (char c in statement) {
        if (inString) {
            sb.Append(c);
            if (c == '"') {
                tokens.Add(sb.ToString());
                sb.Clear();
                inString = false;
            }
        } else if (c == '"') {
            if (sb.Length > 0) {
                tokens.Add(sb.ToString());
                sb.Clear();
            }
            inString = true;
            sb.Append(c);
        } else if (char.IsWhiteSpace(c)) {
            if (sb.Length > 0) {
                tokens.Add(sb.ToString());
                sb.Clear();
            }
        } else {
            sb.Append(c);
        }
    }

    if (sb.Length > 0) {
        tokens.Add(sb.ToString());
    }

    return tokens;
}
