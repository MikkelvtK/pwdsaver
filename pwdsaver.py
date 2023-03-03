import openpyxl
import pyperclip
import argparse
import io
import time
import sys

PERSONAS = "personas"
DESCRIPTION = "Copies user personas to clipboard. Default behavior \
    is to copy all data found. Provide the application an xlsx file \
    with a 'personas' sheet. Column 0 must contain the usernames and \
    column 1 the passwords"


class ValidateExcel(argparse.Action):
    """
    A custom validation action for an argument of the argparse.ArgumentParser.

    Checks whether the user provided a correct .xlsx file.
    """

    def __call__(self, parser, namespace, values, option_string=None):
        if not values.endswith(".xlsx"):
            parser.error("Please enter a valid .xlsx file. Got: {}".format(values))
        setattr(namespace, self.dest, values)


def update_progress(n: int, m: int):
    """Writes the progress of the script to the terminal for feedback to the user.

    Args:
        n (int): nth iteration
        m (int): m total amount of iterations
    """    

    progress = int((n) / m * 100)
    sys.stdout.write("\r [{0}] {1}%".format('#'*(progress//10), progress))


def get_args() -> argparse.Namespace:
    """A function to parse command line arguments and provide helpfull messages
    and options to the user.

    Returns:
        argparse.Namespace: The parsed arguments
    """    
    parser = argparse.ArgumentParser(description=DESCRIPTION)
    parser.add_argument("--user", 
        type=str, 
        nargs='?', 
        const='arg_was_not_given', 
        help="Get a specific user persona"
    )
    parser.add_argument("file", 
        help="Input file. Only accepts .xlsx files for now.", 
        action=ValidateExcel, 
        type=str
    )

    return parser.parse_args()


def retrieve_personas(path: str) -> dict[str, str]:
    """A function that will retrieve both persona usernames and passwords from 
    an excel sheet and return it in a list.

    Args:
        path (str): Path to excel file to read

    Returns:
        dict[str, str]: Dictionary with usernames as keys and passwords as values.
    """   

    with open(path, "rb") as f:
        in_mem_file = io.BytesIO(f.read())

    wb = openpyxl.load_workbook(in_mem_file)
    ws = wb[PERSONAS]
    return {row[0]: row[1] for row in ws.iter_rows(values_only=True)}


def copy(item: str):
    """Copies an item to the clipboard and sleeps for 0.5 seconds for the OS
    to properly process the item.

    Args:
        item (str): String to be saved to the clipboard
    """    
    
    pyperclip.copy(item)
    time.sleep(0.5)


def copy_all(personas: dict[str, str]):
    """A function to copy the usernames and passwords of the personas
    to the clipboard. 

    Args:
        personas (dict[str, str]): Dictionary with usernames as keys and passwords 
        as values.
    """ 

    amount = len(personas)
    count = 0

    for user, pwd in personas.items():
        copy(user)
        copy(pwd)

        update_progress(count + 1, amount)
        count += 1     


def copy_single_user(personas: dict[str, str], user: str):
    """Copies a single user persona to the clipboard.

    Args:
        personas (dict[str, str]): Dictionary with usernames as keys and passwords 
        as values.
        user (str): The user to copy to the clipboard
    """   

    for persona, pwd in personas.items():
        if user == persona:
            copy(persona)
            copy(pwd)
            print("Succesfully copied {} to clipboard".format(persona))


def main():
    args = get_args()
    personas = retrieve_personas(args.file)

    if args.user and args.user not in personas:
        print("Application error: Could not find {}".format(args.user))
        sys.exit(1)

    print("Loading personas to clipboard...")

    if args.user:
        copy_single_user(personas, args.user)
    else:
        copy_all(personas)


if __name__ == "__main__":
    main()