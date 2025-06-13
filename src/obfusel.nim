import os, strutils, random, tables, sequtils
import xl  # Pure Nim Excel library for reading and writing XLSX files

randomize()

type
  ObfuscationOptions = object
    preserveHeaders: bool
    preserveFormulas: bool
    preserveNumbers: bool
    shuffleRows: bool
    shuffleColumns: bool
    stringReplacement: StringReplacementType
    numberReplacement: NumberReplacementType
    outputFormat: OutputFormat

  StringReplacementType = enum
    srRandom,          # Replace with random strings
    srConsistent,      # Use consistent replacements for the same values
    srNone             # Don't replace strings

  NumberReplacementType = enum
    nrJitter,          # Add random noise to numbers
    nrRandom,          # Replace with random numbers
    nrConsistent,      # Use consistent replacements for the same values
    nrNone             # Don't replace numbers

  OutputFormat = enum
    ofXlsx,            # Output as XLSX files
    ofCsv              # Output as CSV files (legacy)

  CellValueKind = enum
    cvString, cvNumber, cvBool, cvEmpty

  CellValue = object
    case kind: CellValueKind
    of cvString: strVal: string
    of cvNumber: numVal: float
    of cvBool: boolVal: bool
    of cvEmpty: discard

# Generate a random string of a specified length
proc randomString(length: int): string =
  result = ""
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
  for i in 1..length:
    result.add(chars[rand(chars.len - 1)])

# Add jitter to a number value (add or subtract up to maxPercent percent)
proc jitterNumber(value: float, maxPercent: float = 15.0): float =
  let jitterMultiplier = 1.0 + (rand(maxPercent * 2) - maxPercent) / 100.0
  return value * jitterMultiplier

# Convert string to CellValue
proc toCellValue(s: string): CellValue =
  if s.strip() == "":
    return CellValue(kind: cvEmpty)
  
  # Try to parse as number
  try:
    let numVal = parseFloat(s)
    return CellValue(kind: cvNumber, numVal: numVal)
  except ValueError:
    discard
  
  # Try to parse as boolean
  let lowerS = s.toLowerAscii()
  if lowerS in ["true", "false", "yes", "no", "1", "0"]:
    let boolVal = lowerS in ["true", "yes", "1"]
    return CellValue(kind: cvBool, boolVal: boolVal)
  
  # Default to string
  return CellValue(kind: cvString, strVal: s)

# Convert CellValue back to string
proc toString(cv: CellValue): string =
  case cv.kind:
    of cvString: return cv.strVal
    of cvNumber: return $cv.numVal
    of cvBool: return $cv.boolVal
    of cvEmpty: return ""

# Print help information
proc printHelp() =
  echo "Excel Obfuscator - A tool to obfuscate Excel files"
  echo "Usage: obfusel [options] input_file output_file_or_directory"
  echo ""
  echo "By default, creates a single XLSX file with obfuscated data."
  echo "With --csv option, creates CSV files for each sheet in a directory."
  echo ""
  echo "Options:"
  echo "  --help, -h                 Show this help message"
  echo "  --csv                      Output as CSV files (one per sheet) instead of XLSX"
  echo "  --preserve-headers, -ph    Preserve header values in the first row"
  echo "  --preserve-formulas, -pf   Preserve formulas (don't replace them with values)"
  echo "  --preserve-numbers, -pn    Keep numeric values unchanged"
  echo "  --shuffle-rows, -sr        Shuffle the order of rows (except headers if preserved)"
  echo "  --shuffle-columns, -sc     Shuffle the order of columns"
  echo "  --string-replacement=TYPE  How to replace string values:"
  echo "                             random (default): replace with random strings"
  echo "                             consistent: same original value gets same replacement"
  echo "                             none: don't replace strings"
  echo "  --number-replacement=TYPE  How to replace numeric values:"
  echo "                             jitter (default): add random noise to numbers"
  echo "                             random: replace with completely random numbers"
  echo "                             consistent: same original value gets same replacement"
  echo "                             none: don't replace numbers"
  echo ""
  echo "Examples:"
  echo "  obfusel --preserve-headers input.xlsx output.xlsx"
  echo "  obfusel --csv --preserve-headers input.xlsx ./obfuscated_csvs/"
  echo "  obfusel --string-replacement=consistent input.xlsx output.xlsx"

# Parse command-line arguments
proc parseArgs(): tuple[options: ObfuscationOptions, inputFile, outputPath: string] =
  var 
    options: ObfuscationOptions
    inputFile = ""
    outputPath = ""
  
  # Set defaults
  options.preserveHeaders = false
  options.preserveFormulas = false
  options.preserveNumbers = false
  options.shuffleRows = false
  options.shuffleColumns = false
  options.stringReplacement = srRandom
  options.numberReplacement = nrJitter
  options.outputFormat = ofXlsx
  
  var i = 1
  while i <= paramCount():
    let param = paramStr(i)
    
    if param in ["--help", "-h"]:
      printHelp()
      quit(0)
    elif param == "--csv":
      options.outputFormat = ofCsv
    elif param in ["--preserve-headers", "-ph"]:
      options.preserveHeaders = true
    elif param in ["--preserve-formulas", "-pf"]:
      options.preserveFormulas = true
    elif param in ["--preserve-numbers", "-pn"]:
      options.preserveNumbers = true
      options.numberReplacement = nrNone
    elif param in ["--shuffle-rows", "-sr"]:
      options.shuffleRows = true
    elif param in ["--shuffle-columns", "-sc"]:
      options.shuffleColumns = true
    elif param.startsWith("--string-replacement="):
      let value = param.split('=')[1].toLowerAscii()
      case value:
        of "random": options.stringReplacement = srRandom
        of "consistent": options.stringReplacement = srConsistent
        of "none": options.stringReplacement = srNone
        else:
          echo "Invalid string replacement type: ", value
          printHelp()
          quit(1)
    elif param.startsWith("--number-replacement="):
      let value = param.split('=')[1].toLowerAscii()
      case value:
        of "jitter": options.numberReplacement = nrJitter
        of "random": options.numberReplacement = nrRandom
        of "consistent": options.numberReplacement = nrConsistent
        of "none": options.numberReplacement = nrNone
        else:
          echo "Invalid number replacement type: ", value
          printHelp()
          quit(1)
    elif inputFile == "":
      inputFile = param
    elif outputPath == "":
      outputPath = param
    else:
      echo "Too many arguments provided."
      printHelp()
      quit(1)
    
    i += 1
  
  if inputFile == "" or outputPath == "":
    echo "Input file and output path must be specified."
    printHelp()
    quit(1)
  
  return (options, inputFile, outputPath)

# Check if a cell has content (not empty)
proc hasContent(cell: XlCell): bool =
  # Check if cell has a value by trying to access it
  try:
    let val = cell.value
    return val != "" and val != "0" # Basic check for non-empty content
  except:
    return false

# Process a single sheet and return its data
proc processSheet(sheet: XlSheet, sheetName: string): seq[seq[CellValue]] =
  echo "Reading sheet: ", sheetName
  var sheetData = newSeq[seq[CellValue]]()
  
  # Find the actual used range by scanning for content
  var maxRow = 0
  var maxCol = 0
  
  # Scan for data - check up to 1000 rows and 100 columns
  for row in 1..1000:
    var hasDataInRow = false
    for col in 1..100:
      let cell = sheet.cell(row, col)
      if cell.hasContent:
        hasDataInRow = true
        maxRow = max(maxRow, row)
        maxCol = max(maxCol, col)
    
    # If we haven't found data in the last 10 rows, assume we're done
    if not hasDataInRow and row > maxRow + 10:
      break
  
  if maxRow == 0:
    echo "Sheet '", sheetName, "' appears to be empty"
    return sheetData
  
  echo "Sheet dimensions: rows 1-", maxRow, ", cols 1-", maxCol
  
  # Read the actual data
  for row in 1..maxRow:
    var rowData = newSeq[CellValue]()
    for col in 1..maxCol:
      let cell = sheet.cell(row, col)
      var cellValue: CellValue
      
      if not cell.hasContent:
        cellValue = CellValue(kind: cvEmpty)
      else:
        let cellVal = cell.value
        # Check if it's a number
        if cell.isNumber:
          cellValue = CellValue(kind: cvNumber, numVal: cell.number)
        else:
          # Try to parse as boolean
          let lowerVal = cellVal.toLowerAscii().strip()
          if lowerVal in ["true", "false", "yes", "no"]:
            let boolVal = lowerVal in ["true", "yes"]
            cellValue = CellValue(kind: cvBool, boolVal: boolVal)
          else:
            # Default to string
            cellValue = CellValue(kind: cvString, strVal: cellVal)
      
      rowData.add(cellValue)
    sheetData.add(rowData)
  
  echo "Sheet '", sheetName, "' loaded with ", sheetData.len, " rows and ", 
       if sheetData.len > 0: sheetData[0].len else: 0, " columns"
  
  return sheetData

# Get all sheet names from the workbook using the built-in iterator
proc getAllSheetNames(workbook: XlWorkbook): seq[string] =
  var sheetNames: seq[string] = @[]
  
  # Use the sheetNames iterator from xl library
  for sheetName in workbook.sheetNames:
    sheetNames.add(sheetName)
    echo "Found sheet: ", sheetName
  
  echo "Total sheets found: ", sheetNames.len
  return sheetNames

# Read Excel file using xl library
proc readExcelFile(filename: string): Table[string, seq[seq[CellValue]]] =
  result = initTable[string, seq[seq[CellValue]]]()
  
  try:
    echo "Loading workbook: ", filename
    let workbook = xl.load(filename)
    
    # Get all sheet names using the built-in iterator
    let sheetNames = getAllSheetNames(workbook)
    
    if sheetNames.len == 0:
      echo "No sheets found in the workbook."
      return
    
    # Process each discovered sheet
    for sheetName in sheetNames:
      try:
        let sheet = workbook.sheet(sheetName)
        if sheet != nil:
          let sheetData = processSheet(sheet, sheetName)
          if sheetData.len > 0:  # Only add sheets with data
            result[sheetName] = sheetData
          else:
            echo "Skipping empty sheet: ", sheetName
        else:
          echo "Warning: Could not access sheet: ", sheetName
      except Exception as e:
        echo "Error processing sheet '", sheetName, "': ", e.msg
        continue
    
    echo "Successfully loaded ", result.len, " non-empty sheets"
    
  except Exception as e:
    echo "Error reading Excel file: ", e.msg
    echo "Make sure the file exists and is a valid Excel file (.xlsx)"
    quit(1)

# Write data to XLSX file using xl library
proc writeToXlsx(data: Table[string, seq[seq[CellValue]]], filename: string) =
  try:
    let workbook = newWorkbook()
    var sheetCount = 0
    var worksheets: seq[XlSheet] = @[]
    
    for sheetName, sheetData in data.pairs:
      echo "Writing sheet: ", sheetName
      
      let worksheet = if sheetCount == 0:
        # For the first sheet, create a new one with the desired name
        # instead of trying to use the default active sheet
        let newSheet = workbook.add(sheetName)
        newSheet
      else:
        # Add additional sheets
        try:
          workbook.add(sheetName)
        except:
          # If adding sheet fails, try with a modified name
          let modifiedName = sheetName & "_" & $sheetCount
          echo "Failed to add sheet '", sheetName, "', trying '", modifiedName, "'"
          workbook.add(modifiedName)
      
      worksheets.add(worksheet)
      sheetCount += 1
      
      # Write data row by row with bounds checking
      for rowIdx, row in sheetData.pairs:
        if rowIdx >= 1000000:  # Excel row limit safety check
          echo "Warning: Skipping rows beyond Excel limit (1M rows)"
          break
          
        for colIdx, cellValue in row.pairs:
          if colIdx >= 16384:  # Excel column limit safety check  
            echo "Warning: Skipping columns beyond Excel limit (16K cols)"
            break
            
          # Skip empty cells to avoid potential issues
          if cellValue.kind == cvEmpty:
            continue
            
          try:
            let cell = worksheet.cell(rowIdx + 1, colIdx + 1)  # xl uses 1-based indexing
            
            case cellValue.kind:
              of cvString:
                if cellValue.strVal != "":
                  cell.value = cellValue.strVal
              of cvNumber:
                cell.number = cellValue.numVal
              of cvBool:
                cell.value = $cellValue.boolVal
              of cvEmpty:
                discard # Already handled above
          except Exception as cellError:
            echo "Warning: Failed to write cell (", rowIdx + 1, ",", colIdx + 1, "): ", cellError.msg
            continue
    
    workbook.save(filename)
    echo "XLSX file written successfully: ", filename
  except Exception as e:
    echo "Error writing XLSX file: ", e.msg
    echo "Error details: ", e.getStackTrace()
    quit(1)

# Write data to CSV files (legacy function)
proc writeToCsv(data: Table[string, seq[seq[CellValue]]], outputDir: string) =
  if not dirExists(outputDir):
    echo "Output directory not found, creating: ", outputDir
    createDir(outputDir)
  
  for sheetName, sheetData in data.pairs:
    let outputFilePath = outputDir / (sheetName & ".csv")
    let csvFile = open(outputFilePath, fmWrite)
    defer: csvFile.close()

    for row in sheetData:
      let line = row.map(proc(cv: CellValue): string = 
        let s = cv.toString()
        if s.contains(',') or s.contains('"') or s.contains('\n'):
          return "\"" & s.replace("\"", "\"\"") & "\""
        else:
          return s
      ).join(",")
      csvFile.writeLine(line)
    
    echo "CSV file written: ", outputFilePath

# Obfuscate cell value
proc obfuscateCell(cellValue: CellValue, options: ObfuscationOptions, 
                  stringReplacements, numberReplacements: var Table[string, string]): CellValue =
  case cellValue.kind:
    of cvEmpty:
      return cellValue
    of cvString:
      if options.stringReplacement == srNone:
        return cellValue
      
      case options.stringReplacement:
        of srRandom:
          let newLength = max(3, cellValue.strVal.len)
          return CellValue(kind: cvString, strVal: randomString(newLength))
        of srConsistent:
          if not stringReplacements.hasKey(cellValue.strVal):
            stringReplacements[cellValue.strVal] = randomString(max(3, cellValue.strVal.len))
          return CellValue(kind: cvString, strVal: stringReplacements[cellValue.strVal])
        of srNone:
          return cellValue
    
    of cvNumber:
      if options.preserveNumbers or options.numberReplacement == nrNone:
        return cellValue
      
      case options.numberReplacement:
        of nrJitter:
          return CellValue(kind: cvNumber, numVal: jitterNumber(cellValue.numVal))
        of nrRandom:
          return CellValue(kind: cvNumber, numVal: rand(cellValue.numVal * 2.0))
        of nrConsistent:
          let key = $cellValue.numVal
          if not numberReplacements.hasKey(key):
            numberReplacements[key] = $(rand(1000.0))
          return CellValue(kind: cvNumber, numVal: parseFloat(numberReplacements[key]))
        of nrNone:
          return cellValue
    
    of cvBool:
      return cellValue  # Keep booleans unchanged

# Main obfuscation function
proc obfuscateExcel(inputFile, outputPath: string, options: ObfuscationOptions) =
  echo "Processing file: ", inputFile
  
  # Read the Excel file
  var data = readExcelFile(inputFile)
  
  if data.len == 0:
    echo "No sheets found in the Excel file."
    quit(1)
  
  # Process each sheet
  for sheetName, sheetData in data.mpairs:
    echo "-----------------------------------"
    echo "Processing sheet: ", sheetName
    
    if sheetData.len == 0:
      echo "Sheet '", sheetName, "' is empty, skipping."
      continue
    
    # Replacement tables for consistent replacements (reset for each sheet)
    var 
      stringReplacements = initTable[string, string]()
      numberReplacements = initTable[string, string]()
    
    # Shuffle rows if requested (preserve header if needed)
    if options.shuffleRows:
      let startIdx = if options.preserveHeaders and sheetData.len > 0: 1 else: 0
      if sheetData.len > startIdx:
        var rowsToShuffle = sheetData[startIdx..^1]
        shuffle(rowsToShuffle)
        for i in startIdx..<sheetData.len:
          sheetData[i] = rowsToShuffle[i - startIdx]
    
    # Shuffle columns if requested
    if options.shuffleColumns and sheetData.len > 0 and sheetData[0].len > 0:
      let numCols = sheetData[0].len
      var colIndices = toSeq(0..<numCols)
      shuffle(colIndices)
      
      var newData = newSeq[seq[CellValue]](sheetData.len)
      for rowIdx in 0..<sheetData.len:
        let originalRow = sheetData[rowIdx]
        newData[rowIdx] = newSeq[CellValue](numCols)
        for colIdx in 0..<numCols:
          if colIdx < originalRow.len and colIndices[colIdx] < originalRow.len:
            newData[rowIdx][colIdx] = originalRow[colIndices[colIdx]]
          else:
            newData[rowIdx][colIdx] = CellValue(kind: cvEmpty)
      sheetData = newData
    
    # Obfuscate the data
    for rowIdx in 0..<sheetData.len:
      let isHeaderRow = (rowIdx == 0 and options.preserveHeaders)
      
      if not isHeaderRow:
        for colIdx in 0..<sheetData[rowIdx].len:
          sheetData[rowIdx][colIdx] = obfuscateCell(
            sheetData[rowIdx][colIdx], 
            options, 
            stringReplacements, 
            numberReplacements
          )
    
    echo "Sheet '", sheetName, "' processed successfully"
  
  # Write output
  case options.outputFormat:
    of ofXlsx:
      writeToXlsx(data, outputPath)
    of ofCsv:
      writeToCsv(data, outputPath)
  
  echo "-----------------------------------"
  echo "Obfuscation completed successfully!"

proc main() =
  let args = parseArgs()
  obfuscateExcel(args.inputFile, args.outputPath, args.options)

when isMainModule:
  main()
