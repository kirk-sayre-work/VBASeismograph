#!/usr/bin/env python

# This file is subject to the terms and conditions defined in file 'LICENSE.txt', which is part of this source code package.

import argparse
import sys
import subprocess
import os

###########################################################################
def _get_pcode_ids(pcode):
    """
    Pull out the function and variable names from the given p-code
    disassembly.

    pcode - (str) The p-code disassembly.

    return - (set) The set of function and variable names.
    """

    # Look at each line of the disassembly.
    ids = set()
    in_id_section = False
    skip = False
    instructions = None
    for line in pcode.split("\n"):

        # Should we skip this line?
        if (skip):
            skip = False
            continue
        
        # Is this the start of the ID section?
        if (line == "Identifiers:"):
            in_id_section = True
            # Skip the next blank line.
            skip = True
            continue

        # Is this the start of the instruction disassembly?
        if (line.startswith("Line #") and (instructions is None)):

            # Start saving instructions.
            instructions = ""
            continue
        
        # Is this an instruction?
        if (instructions is not None):
            instructions += line + "\n"
            continue
        
        # Are we saving IDs?
        if (in_id_section):

            # Is this an ID line?
            if (":" in line):
                curr_id = line[line.index(":") + 1:].strip()
                ids.add(curr_id)
                continue

            # Done with ID section?
            else:
                in_id_section = False

    # Now filter out the IDs that don't appear in the p-code
    # instructions.
    tmp = set()
    for curr_id in ids:
        if (curr_id in instructions):
            tmp.add(curr_id)
    ids = tmp

    # Return the function names and variables.
    return ids

###########################################################################
def _missing_ids(vba, pcode_ids, verbose=False):
    """
    See if there are any function names or variables that appear in
    the p-code that do not appear in the decompressed VBA source code.

    vba - (str) The decompressed VBA source code.

    pcode_ids - (set) The IDs defined in the p-code.
    
    return - (boolean) True if there are function names or variables
    that appear in the p-code that do not appear in the decompressed
    VBA source code, False if they match.
    """

    # Check each ID.
    missing = False
    for curr_id in pcode_ids:
        if (curr_id not in vba):
            if (verbose):
                print "P-code ID '" + str(curr_id) + "' is missing."
            missing = True
    return missing

###########################################################################
def _get_pcode_strs(pcode):
    """
    Pull out string literals from the given p-code disassembly.

    pcode - (str) The p-code disassembly.

    return - (set) The set of literal strings.
    """

    # Look at each line of the disassembly.
    strs = set()
    for line in pcode.split("\n"):

        # Is this a string literal instruction?
        line = line.strip()
        if (line.startswith("LitStr ")):
            curr_str = line[line.index('"') + 1:-1]
            strs.add(curr_str)
            
    # Return the string literals.
    return strs

###########################################################################
def _missing_strs(vba, pcode_strs, verbose=False):
    """
    See if there are any string literals appear in the p-code that do
    not appear in the decompressed VBA source code.

    vba - (str) The decompressed VBA source code.

    pcode_strs - (set) The string literals defined in the p-code.
    
    return - (boolean) True if there are string lierals that appear in
    the p-code that do not appear in the decompressed VBA source code,
    False if they match.
    """

    # Check each string.
    missing = False
    for curr_str in pcode_strs:
        if ((('"' + curr_str + '"') not in vba) and
            (("'" + curr_str + "'") not in vba)):
            if (verbose):
                print "P-code string '" + str(curr_str) + "' is missing."
            missing = True
    return missing

###########################################################################
def detect_stomping_via_pcode(filename, verbose=False):
    """
    Detect VBA stomping by comparing variables, function names, and
    static strings in the Office doc p-code to the same items in the
    decompressed VBA source code.

    filename - (str) The name of the Office file to check for VBA
    stomping.

    verbose - (boolean) If True print out detailed debugging
    information.

    return - (boolean) True if the given Office doc has stomped VBA
    source code, False if not.

    raises - ValueError, if running sigtool or pcodedmp.py fails.
    """

    # Get the p-code disassembly.
    pcode = None
    try:
        pcode = subprocess.check_output(["python", os.environ["PCODEDMP_DIR"] + "/pcodedmp.py", filename])
    except Exception as e:
        raise ValueError("Running pcodedmp.py on " + filename + \
                         " failed. " + str(e))
    if (verbose):
        print "----------------------------------------------"
        print pcode
    
    # Get the decompressed VBA source code.
    vba = None
    try:
        vba = subprocess.check_output(["sigtool", "--vba", filename])
    except Exception as e:
        raise ValueError("Running sigtool on " + filename + \
                         " failed. " + str(e))
    if (verbose):
        print "----------------------------------------------"
        print vba
        print "----------------------------------------------"
    
    # Get the variable and function names from the p-code.
    pcode_ids = _get_pcode_ids(pcode)

    # Check to see if all the function names and variables from the
    # p-code appear in the decompressed VBA source code.
    stomped = False
    if (_missing_ids(vba, pcode_ids, verbose)):
        stomped = True

    # Get the string literals from the p-code.
    pcode_strs = _get_pcode_strs(pcode)
    
    # Check to see if all the string literals from the p-code appear
    # in the decompressed VBA source code.
    if (_missing_strs(vba, pcode_strs, verbose)):
        stomped = True
    
    # Return whether the VBA source code was stomped.
    return stomped

###########################################################################
def is_vba_stomped(filename, verbose=False):
    """
    Check to see if the given Office doc file has had its VBA source
    code stomped.

    filename - (str) The name of the Office file to check for VBA
    stomping.

    verbose - (boolean) If True print out detailed debugging
    information.

    return - (boolean) True if the given Office doc has stomped VBA
    source code, False if not.

    raises - ValueError, if running sigtool or pcodedmp.py fails.
    """

    # TODO: For now just detect with 1 method.
    return detect_stomping_via_pcode(filename, verbose)
    
###########################################################################
## Main Program
###########################################################################

if __name__ == '__main__':

    # Check to see if prerequisites are installed.

    # Check pcodedmp.py
    if ("PCODEDMP_DIR" not in os.environ):
        print "ERROR: PCODEDMP_DIR environment variable not set. " + \
            "This is the install directory for pcodedmp.py (https://github.com/bontchev/pcodedmp)."
        sys.exit(1)
    try:
        subprocess.check_output(["python", os.environ["PCODEDMP_DIR"] + "/pcodedmp.py", "-h"])
    except Exception as e:
        print "ERROR: It looks like pcodedmp is not installed. " + str(e) + "\n"
        print "To install pcodedmp do the following:\n"
        print "git clone https://github.com/bontchev/pcodedmp.git\n"
        print "You will also need to set the PCODEDMP_DIR environment " + \
            "variable to the pcodedmp install directory."
        sys.exit(1)

    # Check ClamAV sigtool.
    try:
        subprocess.check_output(["sigtool", "-h"])
    except Exception as e:
        print "ERROR: It looks like ClamAV sigtool is not installed. " + str(e) + "\n"
        print "To install sigtool do the following:\n"
        print "sudo apt-get install clamav"
        sys.exit(1)

    # Get the arguments.
    help_msg = "Check to see if a given Office doc file has had its" + \
               " VBA source code stomped."
    parser = argparse.ArgumentParser(description=help_msg)
    parser.add_argument('-v', "--verbose",
                        help="Print debug information.",
                        action='store_true',
                        required=False)
    parser.add_argument("doc",
                        help="The Office doc to check.")
    args = parser.parse_args()
        
    # Check for VBA stomping.
    if (is_vba_stomped(args.doc, args.verbose)):
        print "WARNING: File " + args.doc + " is VBA stomped."
    else:
        print "File " + args.doc + " is NOT VBA stomped."
