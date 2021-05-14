#!/usr/bin/env python

# This file is subject to the terms and conditions defined in file 'LICENSE.txt', which is part of this source code package.

# Check to see if an Office doc has had its compressed VBA source code stomped.

import argparse
import sys
import subprocess
import os
import zipfile
import time
import os.path
import shutil
import tempfile
import re
import traceback

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
        if skip:
            skip = False
            continue

        # Is this the start of the ID section?
        if line == "Identifiers:":
            in_id_section = True
            # Skip the next blank line.
            skip = True
            continue

        # Is this the start of the instruction disassembly?
        if line.startswith("Line #") and (instructions is None):

            # Start saving instructions.
            instructions = ""
            continue

        # Is this an instruction?
        if instructions is not None:
            instructions += line + "\n"
            continue

        # Are we saving IDs?
        if in_id_section:

            # Is this an ID line?
            if ":" in line:
                curr_id = line[line.index(":") + 1 :].strip()
                ids.add(curr_id)
                continue

            # Done with ID section?
            else:
                in_id_section = False

    # These IDs seem to appear in the p-code and not necessarily in
    # the VBA source code. Filter them out.
    common_ids = set(
        [
            "Word",
            "VBA",
            "Win16",
            "Win32",
            "Win64",
            "Mac",
            "VBA6",
            "VBA7",
            "Project1",
            "stdole",
            "VBAProject",
            "Excel",
            "Project",
            "ThisDocument",
            "_Evaluate",
            "Normal",
            "Office",
            "Add",
            "MSForms",
            "UserForm",
            "Document",
        ]
    )

    # Now filter out the IDs that don't appear in the p-code
    # instructions.
    tmp = set()
    for curr_id in ids:

        # Skip IDs that are obviously not used or are common.
        if ((instructions is not None) and (curr_id not in instructions)) or (curr_id in common_ids) or (curr_id.startswith("_B_var_")):
            continue

        # Make sure the ID string is not embedded in some other
        # string.
        if instructions is not None:
            pat = "." + curr_id + "."
            strs = re.findall(pat, instructions)
            keep = False
            for curr_str in strs:
                if (not curr_str[0].isalnum()) and (not curr_str[len(curr_str) - 1].isalnum()):
                    keep = True
                    break
            if not keep:
                continue

        # It looks like some IDs in the p-code have underscores added
        # to the prefix or suffix of the name. Strip those off so we
        # can properly match the IDs in the VBA.
        while curr_id.startswith("_"):
            curr_id = curr_id[1:]
        while curr_id.endswith("_"):
            curr_id = curr_id[:-1]

        # This is a valid ID. Save it.
        tmp.add(curr_id)

    # Use the filtered IDs.
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

    return - (float) % missing items.
    """

    # Check each ID.
    num_missing = 0.0
    for curr_id in pcode_ids:
        if curr_id not in vba:
            if verbose:
                print("P-code ID '{0}' is missing.".format(curr_id))
            num_missing += 1
    if len(pcode_ids) == 0:
        return 0.0
    return num_missing / len(pcode_ids)


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
        if line.startswith("LitStr "):
            curr_str = line[line.index('"') + 1 : -1]
            strs.add(curr_str.replace('"', '""'))

    # Return the string literals.
    return strs


###########################################################################
def _missing_strs(vba, pcode_strs, verbose=False):
    """
    See if there are any string literals appear in the p-code that do
    not appear in the decompressed VBA source code.

    vba - (str) The decompressed VBA source code.

    pcode_strs - (set) The string literals defined in the p-code.

    return - (float) % missing items.
    """

    # Check each string.
    num_missing = 0.0
    for curr_str in pcode_strs:
        if (('"' + curr_str + '"') not in vba) and (("'" + curr_str + "'") not in vba):
            if verbose:
                print("P-code string '" + str(curr_str) + "' is missing.")
            num_missing += 1
    if len(pcode_strs) == 0:
        return 0.0
    return num_missing / len(pcode_strs)


###########################################################################
def _get_pcode_comments(pcode):
    """
    Pull out comments from the given p-code disassembly.

    pcode - (str) The p-code disassembly.

    return - (set) The set of comments.
    """

    # Look at each line of the disassembly.
    comments = set()
    for line in pcode.split("\n"):

        # Is this a comment instruction?
        line = line.strip()
        if line.startswith("QuoteRem "):
            curr_str = line[line.index('"') + 1 : -1]
            if curr_str.endswith("_"):
                curr_str = curr_str[-1]
            comments.add(curr_str)

    # Return the comments.
    return comments


###########################################################################
def _missing_comments(vba, pcode_comments, verbose=False):
    """
    See if there are any comments appear in the p-code that do not
    appear in the decompressed VBA source code.

    vba - (str) The decompressed VBA source code.

    pcode_comments - (set) The comments defined in the p-code.

    return - (float) % missing items.
    """

    # Check each comment.
    num_missing = 0.0
    for curr_str in pcode_comments:

        # Try the easy case (1 line comments) first.
        if curr_str in vba:
            continue

        # Need to handle multiline comments. We do this by allowing
        # newlines to appear where spaces appear.
        pat = curr_str
        pat = pat.replace("'", " ")
        pat = pat.replace("\\", "\\\\")
        pat = pat.replace("(", "\\(")
        pat = pat.replace(")", "\\)")
        pat = pat.replace("[", "\\[")
        pat = pat.replace("]", "\\]")
        pat = pat.replace("}", "\\}")
        pat = pat.replace("}", "\\}")
        pat = pat.replace("+", "\\+")
        pat = pat.replace("^", "\\^")
        pat = pat.replace(".", "\\.")
        pat = pat.replace("$", "\\$")
        pat = pat.replace("?", "\\?")
        pat = pat.replace("*", "\\*")
        pat = pat.replace(",", "\\,")
        pat = pat.replace("|", "\\|")
        tmp = ""
        for i in pat.split(" "):
            if len(i) == 0:
                continue
            if len(tmp) > 0:
                tmp += "[\\s\\r\\n']{1,50}"
            tmp += i
        pat = tmp
        if re.search(pat, vba, re.MULTILINE) is None:
            if verbose:
                print("P-code comment '{0}' is missing".format(curr_str))
            num_missing += 1
    if len(pcode_comments) == 0:
        return 0.0
    return num_missing / len(pcode_comments)


###########################################################################
def _unzip_office_doc(filename):
    """
    Extract the vbaProject.bin macro file from Office 2007+ files if
    needed.

    filename - (str) The Office doc file name.

    return - (str) The name of the macro file to analyze. This will be
    a temporary copy of vbaProject.bin for Office 2007+ files and will
    be the original file name if not given a Office 2007+ file.
    """

    # Is this an Office 2007+ file?
    try:
        file_type = subprocess.check_output(["file", filename])
        if (b"2007+" not in file_type) or (b"Microsoft" not in file_type):
            return filename
    except Exception as e:

        raise ValueError("Running file on {0} failed. {1}".format(filename, e))

    # If we get here we have a Office 2007+ file. Unzip it.
    out_dir = None
    try:
        zip_ref = zipfile.ZipFile(filename, "r")
        millis = int(round(time.time() * 1000))
        out_dir = filename
        if os.sep in out_dir:
            out_dir = out_dir[out_dir.rindex(os.sep) + 1 :]
        out_dir = "{0}{1}{2}_{3}".format(tempfile.gettempdir(), os.sep, out_dir, millis)
        zip_ref.extractall(out_dir)
        zip_ref.close()
    except Exception as e:
        raise ValueError("Zip extraction of {0} failed. {1}".format(filename, e))

    # Look for word/vbaProject.bin in the unzipped directory.
    # TODO: Handle renamed/extra vbaProject.bin files.
    unzipped_filename = os.path.join(out_dir, "word", "vbaProject.bin")
    if os.path.isfile(unzipped_filename):
        return unzipped_filename
    else:
        raise ValueError("{0} not found after unzipping Office 2007+ file {1}".format(unzipped_filename, filename))


###########################################################################
def _cleanup_office_doc(orig_filename, filename):
    """
    Delete the temporary directory where an Office 2007+ file was
    unzipped if needed.

    orig_filename - The original Office doc file name.

    filename - The file that was analyzed. This could be
    vbaProject.bin from the unzipped original file.
    """

    # Is there anything to clean up?
    if orig_filename == filename:
        return

    # Delete the temporary directory of unzipped files.
    subdir = "{0}word{1}".format(os.sep, os.sep)
    if subdir in filename:
        tmp_dir = filename[: filename.index(subdir)]
        shutil.rmtree(tmp_dir)


###########################################################################
def detect_stomping_via_pcode(filename, verbose=False, sensitivity="medium"):
    """
    Detect VBA stomping by comparing variables, function names, and
    static strings in the Office doc p-code to the same items in the
    decompressed VBA source code.

    filename - (str) The name of the Office file to check for VBA
    stomping.

    verbose - (boolean) If True print out detailed debugging
    information.

    sensitivity - (string) The sensitivity of the VBA stomping check to
    perform ("low", "medium", or "high").

    return - (boolean) True if the given Office doc has stomped VBA
    source code, False if not.

    raises - ValueError, if running sigtool or pcodedmp.py fails.
    """

    # Extract the VBA macro file from Office 2007+ files if needed.
    orig_filename = filename
    filename = _unzip_office_doc(filename)

    # Get the p-code disassembly.
    pcode = None
    try:
        pcode = subprocess.check_output(["python3", os.path.join(os.environ["PCODEDMP_DIR"], "pcodedmp.py"), filename])
        pcode = pcode.decode("utf-8")
    except Exception as e:
        raise ValueError("Running pcodedmp.python {0} failed. {1}".format(orig_filename, e))
    if verbose:
        print("----------------------------------------------")
        print(pcode)

    # Get the decompressed VBA source code.
    vba = None
    try:
        vba = subprocess.check_output(["olevba", filename])
    except Exception as e:
        raise ValueError("Running olevba on {0} failed. {1}".format(orig_filename, e))
    vba = vba.decode("utf-8")
    vba = vba.replace(chr(0x0D), "")
    vba = vba.replace("_\n", "\n")
    if verbose:
        print("----------------------------------------------")
        print(vba)
        print("----------------------------------------------")

    # Figure out the threshold of missing items to trigger VBA stomping.
    if sensitivity == "low":
        threshold = 0.5
    elif sensitivity == "medium":
        threshold = 0.3
    elif sensitivity == "high":
        threshold = 0.1

    # Get the variable and function names from the p-code.
    pcode_ids = _get_pcode_ids(pcode)

    # Check to see if all the function names and variables from the
    # p-code appear in the decompressed VBA source code.
    stomped = False
    pct_missing_ids = _missing_ids(vba, pcode_ids, verbose)
    if pct_missing_ids > threshold:
        stomped = True

    # Get the string literals from the p-code.
    pcode_strs = _get_pcode_strs(pcode)

    # Check to see if all the string literals from the p-code appear
    # in the decompressed VBA source code.
    pct_missing_strs = _missing_strs(vba, pcode_strs, verbose)
    if pct_missing_strs > threshold:
        stomped = True

    # Get the comments from the p-code.
    pcode_comments = _get_pcode_comments(pcode)

    # Check to see if all the comments from the p-code appear
    # in the decompressed VBA source code.
    pct_missing_comments = _missing_comments(vba, pcode_comments, verbose)
    if pct_missing_comments > threshold:
        stomped = True

    # Clean up extracted 2007+ macro file if needed.
    _cleanup_office_doc(orig_filename, filename)

    # Print more info if needed.
    if verbose:
        print("\n---------------------------------")
        print("% Missing IDs:\t\t\t {0}".format(pct_missing_ids))
        print("% Missing Strings:\t\t {0}".format(pct_missing_strs))
        print("% Missing Comments:\t\t {0}".format(pct_missing_comments))
        print("---------------------------------\n")

    # Return whether the VBA source code was stomped.
    return stomped


###########################################################################
def is_vba_stomped(filename, verbose=False, sensitivity="medium"):
    """
    Check to see if the given Office doc file has had its VBA source
    code stomped.

    filename - (str) The name of the Office file to check for VBA
    stomping.

    verbose - (boolean) If True print out detailed debugging
    information.

    sensitivity - (string) The sensitivity of the VBA stomping check to
    perform ("low", "medium", or "high").

    return - (boolean) True if the given Office doc has stomped VBA
    source code, False if not.

    raises - ValueError, if running olevba or pcodedmp.py fails.
    """

    # TODO: For now just detect with 1 method.
    return detect_stomping_via_pcode(filename, verbose, sensitivity)


###########################################################################
## Main Program
###########################################################################

if __name__ == "__main__":

    # Check to see if prerequisites are installed.

    # Check pcodedmp.py
    if "PCODEDMP_DIR" not in os.environ:
        print(
            "ERROR: PCODEDMP_DIR environment variable not set. This is the install directory for pcodedmp.py (https://github.com/bontchev/pcodedmp)."
        )
        sys.exit(1)
    try:
        subprocess.check_output(["python3", os.path.join(os.environ["PCODEDMP_DIR"], "pcodedmp.py"), "-h"])
    except Exception as e:
        print("ERROR: It looks like pcodedmp is not installed. {0}\n".format(e))
        print("To install pcodedmp do the following:\n")
        print("git clone https://github.com/bontchev/pcodedmp.git\n")
        print("You will also need to set the PCODEDMP_DIR environment " + "variable to the pcodedmp install directory.")
        sys.exit(1)

    # Check olevba.
    try:
        subprocess.check_output(["olevba", "-h"])
    except Exception as e:
        print("ERROR: It looks like olevba is not installed. {0} \n".format(e))
        print("To install olevba do the following:\n")
        print("pip install oletools")
        sys.exit(1)

    # Get the arguments.
    help_msg = "Check to see if a given Office doc file has had its VBA source code stomped."
    parser = argparse.ArgumentParser(description=help_msg)
    parser.add_argument(
        "-s",
        "--sensitivity",
        dest="sensitivity",
        action="store",
        default="medium",
        help="Sensitivity of check (low, medium, high) (default=%default)",
    )
    parser.add_argument("-v", "--verbose", help="Print debug information.", action="store_true", required=False)
    parser.add_argument("doc", help="The Office doc to check.")
    args = parser.parse_args()

    # Check for VBA stomping.
    try:
        if is_vba_stomped(args.doc, args.verbose, args.sensitivity):
            print("WARNING: File {0} is VBA stomped.".format(args.doc))
        else:
            print("File {0} is NOT VBA stomped.".format(args.doc))
        sys.exit(0)
    except ValueError as e:
        print("ERROR: VBA stomping check of {0}".format(args.doc, e))
        # print("{}".format(traceback.format_exc()))

        sys.exit(1)
