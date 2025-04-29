import os, strutils, random, tables, sequtils
import xlsx

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

  StringReplacementType = enum
    srRandom,          # Replace with random strings
    srConsistent,      # Use consistent replacements for the same values
    srNone             # Don't replace strings

  NumberReplacementType = enum
    nrJitter,          # Add random noise to numbers
    nrRandom,          # Replace with random numbers
    nrConsistent,      # Use consistent replacements for the same values
    nrNone             # Don't replace numbers

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

# Print help information
proc printHelp() =
  echo "Excel Obfuscator - A tool to obfuscate Excel files"
  echo "Usage: excel_obfuscator [options] input_file output_file"
  echo ""
  echo "Options:"
  echo "  --help, -h                 Show this help message"
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
  echo "Example:"
  echo "  excel_obfuscator --preserve-headers --string-replacement=consistent input.xlsx output.xlsx"

# Parse command-line arguments
proc parseArgs(): tuple[options: ObfuscationOptions, inputFile, outputFile: string] =
  var 
    options: ObfuscationOptions
    inputFile = ""
    outputFile = ""
  
  # Set defaults
  options.preserveHeaders = false
  options.preserveFormulas = false
  options.preserveNumbers = false
  options.shuffleRows = false
  options.shuffleColumns = false
  options.stringReplacement = srRandom
  options.numberReplacement = nrJitter
  
  var i = 1
  while i <= paramCount():
    let param = paramStr(i)
    
    if param in ["--help", "-h"]:
      printHelp()
      quit(0)
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
    elif outputFile == "":
      outputFile = param
    else:
      echo "Too many arguments provided."
      printHelp()
      quit(1)
    
    i += 1
  
  if inputFile == "" or outputFile == "":
    echo "Input and output files must be specified."
    printHelp()
    quit(1)
  
  return (options, inputFile, outputFile)

# Obfuscate Excel file
proc obfuscateExcel(inputFile, outputFile: string, options: ObfuscationOptions) =
  echo "Processing file: ", inputFile
  
  # Load the Excel workbook
  var wb = Workbook()
  discard wb.load(inputFile)
  
  # Create a new workbook for the obfuscated data
  var obfuscatedWb = Workbook()
  
  # Used for consistent replacement
  var 
    stringReplacements = initTable[string, string]()
    numberReplacements = initTable[float, float]()
  
  # Process each sheet
  for sheetName in wb.getSheetNames():
    echo "Processing sheet: ", sheetName
    var sheet = wb.getSheet(sheetName)
    
    # Get all non-empty cells
    var 
      allRows: seq[int] = @[]
      allCols: seq[int] = @[]
      maxRow = 0
      maxCol = 0
    
    for row, col, cell in sheet.cells():
      maxRow = max(maxRow, row)
      maxCol = max(maxCol, col)
      if row notin allRows:
        allRows.add(row)
      if col notin allCols:
        allCols.add(col)
    
    # Sort the arrays
    allRows.sort()
    allCols.sort()
    
    # Prepare shuffling if required
    var 
      rowMap = newSeq[int](maxRow + 1)
      colMap = newSeq[int](maxCol + 1)
    
    for i in 0..maxRow:
      rowMap[i] = i
    
    for i in 0..maxCol:
      colMap[i] = i
    
    # Shuffle rows if requested (but keep header if needed)
    if options.shuffleRows:
      var startIdx = if options.preserveHeaders: 1 else: 0
      if startIdx < allRows.len:
        var shuffleRows = allRows[startIdx..^1]
        shuffle(shuffleRows)
        for i in startIdx..<allRows.len:
          rowMap[allRows[i]] = shuffleRows[i - startIdx]
    
    # Shuffle columns if requested
    if options.shuffleColumns:
      var shuffleCols = allCols
      shuffle(shuffleCols)
      for i in 0..<allCols.len:
        colMap[allCols[i]] = shuffleCols[i]
    
    # Create a new sheet in the obfuscated workbook
    var obfuscatedSheet = obfuscatedWb.createSheet(sheetName)
    
    # Copy and obfuscate the data
    for row, col, cell in sheet.cells():
      let 
        newRow = rowMap[row]
        newCol = colMap[col]
      
      # Determine if this is a header cell that should be preserved
      let isHeader = (row == 0 and options.preserveHeaders)
      
      # Handle different cell types
      if cell.kind == CellType.formula and options.preserveFormulas:
        # Copy formula unchanged
        obfuscatedSheet.setCellFormula(newRow, newCol, cell.formula)
      else:
        case cell.kind:
          of CellType.empty:
            # Skip empty cells
            discard
          of CellType.boolean:
            # For boolean values, randomly flip or keep
            if not isHeader:
              let newValue = if rand(1) == 0: not cell.boolVal else: cell.boolVal
              obfuscatedSheet.setCellBool(newRow, newCol, newValue)
            else:
              obfuscatedSheet.setCellBool(newRow, newCol, cell.boolVal)
          
          of CellType.numeric:
            # For numeric values
            if isHeader or options.numberReplacement == nrNone:
              # Preserve original value
              obfuscatedSheet.setCellFloat(newRow, newCol, cell.numVal)
            else:
              case options.numberReplacement:
                of nrJitter:
                  # Add jitter to the value
                  obfuscatedSheet.setCellFloat(newRow, newCol, jitterNumber(cell.numVal))
                of nrRandom:
                  # Replace with random number (between 0 and twice the original)
                  let newValue = rand(cell.numVal * 2.0)
                  obfuscatedSheet.setCellFloat(newRow, newCol, newValue)
                of nrConsistent:
                  # Use consistent replacements
                  if not numberReplacements.hasKey(cell.numVal):
                    # First time seeing this value, create a replacement
                    numberReplacements[cell.numVal] = rand(1000.0)
                  obfuscatedSheet.setCellFloat(newRow, newCol, numberReplacements[cell.numVal])
                of nrNone:
                  # This case is handled above
                  discard
          
          of CellType.string:
            # For string values
            if isHeader or options.stringReplacement == srNone:
              # Preserve original value
              obfuscatedSheet.setCellString(newRow, newCol, cell.strVal)
            else:
              case options.stringReplacement:
                of srRandom:
                  # Replace with random string of similar length
                  let newLength = max(3, cell.strVal.len)
                  obfuscatedSheet.setCellString(newRow, newCol, randomString(newLength))
                of srConsistent:
                  # Use consistent replacements
                  if not stringReplacements.hasKey(cell.strVal):
                    # First time seeing this value, create a replacement
                    stringReplacements[cell.strVal] = randomString(max(3, cell.strVal.len))
                  obfuscatedSheet.setCellString(newRow, newCol, stringReplacements[cell.strVal])
                of srNone:
                  # This case is handled above
                  discard
          
          of CellType.date:
            # For date values, add random days (±30 days)
            if isHeader:
              obfuscatedSheet.setCellFloat(newRow, newCol, cell.numVal)
            else:
              let jitter = rand(60) - 30 # Random between -30 and +30
              obfuscatedSheet.setCellFloat(newRow, newCol, cell.numVal + jitter.float)
              # Format as date
              let dateFormat = obfuscatedWb.addFormat()
              dateFormat.setNumberFormat("yyyy-mm-dd")
              obfuscatedSheet.setCellFormat(newRow, newCol, dateFormat)
          
          of CellType.formula:
            # Formula cells are handled above (with the preserveFormulas check)
            # This branch handles the case where we don't preserve formulas
            # In that case, we obfuscate the formula's result value
            let formulaResult = cell.numVal
            if not isHeader and options.numberReplacement != nrNone:
              case options.numberReplacement:
                of nrJitter:
                  obfuscatedSheet.setCellFloat(newRow, newCol, jitterNumber(formulaResult))
                of nrRandom:
                  obfuscatedSheet.setCellFloat(newRow, newCol, rand(formulaResult * 2.0))
                of nrConsistent:
                  if not numberReplacements.hasKey(formulaResult):
                    numberReplacements[formulaResult] = rand(1000.0)
                  obfuscatedSheet.setCellFloat(newRow, newCol, numberReplacements[formulaResult])
                of nrNone:
                  # This case is handled by the outer condition
                  discard
            else:
              obfuscatedSheet.setCellFloat(newRow, newCol, formulaResult)
    
    # Copy column widths (assuming the sheet provides this info)
    for col in 0..maxCol:
      let newCol = colMap[col]
      try:
        let width = sheet.getColWidth(col)
        obfuscatedSheet.setColWidth(newCol, width)
      except:
        # If getColWidth is not available, just skip it
        discard
  
  # Save the obfuscated workbook
  obfuscatedWb.save(outputFile)
  echo "Obfuscated Excel file saved to: ", outputFile

proc main() =
  let args = parseArgs()
  obfuscateExcel(args.inputFile, args.outputFile, args.options)
  echo "Done!"

when isMainModule:
  main()