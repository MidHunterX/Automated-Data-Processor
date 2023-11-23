import config
import sys          # Command line arguments and exit


def main():
    """
    Returns: User Command if it passes the rite of passage
    """
    user_command = riteOfPassage()
    return user_command


def getUserCmd():
    """
    Returns: A single argument from commandline
    """
    max_args = 1
    arg = sys.argv[1:]  # 0th arg is filename.py
    if len(arg) == max_args:
        command = str(sys.argv[1])
    elif len(arg) > max_args:
        sys.exit("Error: Too much arguments")
    elif len(arg) < 1:
        sys.exit("Error: No argument")
    return command


def riteOfPassage():
    """
    Returns: User Command if it passes the rite of passage
    """

    cmd = config.initVarCmd()
    command = getUserCmd()
    if command in cmd.values():
        return command
    else:
        print("Unrecognized Command 😕")
        print("Available CMD:\n > " + "\n > ".join(map(str, cmd.values())))
        sys.exit()
