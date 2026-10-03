"""Opens the File Converter app window (start it with "File Converter.cmd").

Any files passed on the command line are added to the list when it opens.
"""

import sys

from file_converter.gui import main

main(sys.argv[1:])
