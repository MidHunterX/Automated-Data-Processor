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

if command == cmd["form"]:
    processForms.main()
    exit(0)

if command == cmd["db"]:
    processDatabase.main()
    exit(0)

if command == cmd["ifsc"]:
    processBranch.main()
    exit(0)

if command == cmd["excel"]:
    processSpreadsheet.main()
    exit(0)

if command == cmd["bank"]:
    processNEFT.main()
    exit(0)
