import config
import processForms
import processBranch
import processSpreadsheet
import processDatabase
import riteOfPassage
import processNEFT
from sys import exit

command = riteOfPassage.main()
cmd = config.initVarCmd()
var = config.initVarCommon()

if command == cmd["form"]:
    processForms.main(var)
    exit(0)

if command == cmd["db"]:
    processDatabase.main(var)
    exit(0)

if command == cmd["ifsc"]:
    processBranch.main(var)
    exit(0)

if command == cmd["excel"]:
    processSpreadsheet.main(var)
    exit(0)

if command == cmd["bank"]:
    processNEFT.main(var)
    exit(0)
