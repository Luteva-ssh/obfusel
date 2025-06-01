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
  echo "  excel_obfuscator --preserve-headers --string-replacement=consistent input.xlsx output.csv"

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
  
  # Load the Excel workbook using parseExcel
  let sheetTable = parseExcel(inputFile)
  
  # Global replacement tables for consistent replacements
  var 
    stringReplacements = initTable[string, string]()
    numberReplacements = initTable[string, string]()
  
  # Determine output format based on file extension
  let outputExt = splitFile(outputFile).ext.toLowerAscii()
  let isCSVOutput = outputExt == ".csv"
  
  if isCSVOutput:
    # For CSV output, we can only process one sheet
    # Use the first sheet or let user specify
    var targetSheetName = ""
    for sheetName in sheetTable.data.keys:
      targetSheetName = sheetName
      break
    
    if targetSheetName == "":
      echo "No sheets found in the Excel file"
      return
      
    echo "Processing sheet: ", targetSheetName, " (CSV output)"
    let sheet = sheetTable[targetSheetName]
    
    # Convert to sequence of sequences for easier manipulation
    var data = sheet.toSeq()
    
    # Shuffle rows if requested (preserve header if needed)
    if options.shuffleRows and data.len > 0:
      let startIdx = if options.preserveHeaders: 1 else: 0
      if data.len > startIdx:
        var rowsToShuffle = data[startIdx..^1]
        shuffle(rowsToShuffle)
        for i in startIdx..<data.len:
          data[i] = rowsToShuffle[i - startIdx]
    
    # Shuffle columns if requested
    if options.shuffleColumns and data.len > 0 and data[0].len > 0:
      let numCols = data[0].len
      var colIndices = toSeq(0..<numCols)
      shuffle(colIndices)
      
      for rowIdx in 0..<data.len:
        let originalRow = data[rowIdx]
        for colIdx in 0..<numCols:
          data[rowIdx][colIdx] = originalRow[colIndices[colIdx]]
    
    # Obfuscate the data
    for rowIdx in 0..<data.len:
      let isHeaderRow = (rowIdx == 0 and options.preserveHeaders)
      
      for colIdx in 0..<data[rowIdx].len:
        if not isHeaderRow:
          let cellValue = data[rowIdx][colIdx].strip()
          
          if cellValue != "":
            # Try to parse as number
            try:
              let numValue = parseFloat(cellValue)
              # It's a number
              if not options.preserveNumbers:
                case options.numberReplacement:
                  of nrJitter:
                    data[rowIdx][colIdx] = $jitterNumber(numValue)
                  of nrRandom:
                    data[rowIdx][colIdx] = $(rand(numValue * 2.0))
                  of nrConsistent:
                    if not numberReplacements.hasKey(cellValue):
                      numberReplacements[cellValue] = $(rand(1000.0))
                    data[rowIdx][colIdx] = numberReplacements[cellValue]
                  of nrNone:
                    discard # Keep original
            except ValueError:
              # It's a string
              if options.stringReplacement != srNone:
                case options.stringReplacement:
                  of srRandom:
                    let newLength = max(3, cellValue.len)
                    data[rowIdx][colIdx] = randomString(newLength)
                  of srConsistent:
                    if not stringReplacements.hasKey(cellValue):
                      stringReplacements[cellValue] = randomString(max(3, cellValue.len))
                    data[rowIdx][colIdx] = stringReplacements[cellValue]
                  of srNone:
                    discard # Keep original
    
    # Write to CSV
    let csvFile = open(outputFile, fmWrite)
    for row in data:
      csvFile.writeLine(row.join(","))
    csvFile.close()
    
    echo "Obfuscated data saved to CSV: ", outputFile
  else:
    echo "Error: This tool can only output to CSV format (.csv extension required)"
    echo "The xlsx library used doesn't support writing Excel files."
    quit(1)

proc main() =
  let args = parseArgs()
  obfuscateExcel(args.inputFile, args.outputFile, args.options)
  echo "Done!"

when isMainModule:
  main()