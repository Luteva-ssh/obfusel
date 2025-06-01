# Package

version       = "0.1.0"
author        = "Janni Adamski"
description   = "An obfuscator for excel sheets. If you are not allowed to transfer data to an AI system, this can be an easy solution :)."
license       = "MIT"
srcDir        = "src"
bin           = @["obfusel"]
binDir        = "bin"


# Dependencies
requires "xlsx"
requires "nim >= 1.6.0"
