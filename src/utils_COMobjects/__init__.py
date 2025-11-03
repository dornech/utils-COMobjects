# COM object utilities

"""
Package with various utilities to support COM object development with Python based on PyWin32

Set of submodules contains:

- base module with abstract base class definition for Python COM object classes
- submodule with a logger mixin class for Python COM object classes
- submodule with various check routines for own-defined Python COM object classes
- submodule with helpful decorators for own-defined Python COM object classes
- submodule for TypeLib generation (used within module to check Python COM object classes)
"""


# ruff and mypy per file settings
#
# empty lines
# ruff: noqa: E302, E303
# naming conventions
# ruff: noqa: N801, N802, N803, N806, N812, N813, N815, N816, N818, N999

# fmt: off



# version determination - latest import requirement for hatch-vcs-footgun-example
from utils_COMobjects.version import __version__

import utils_COM_classes as COMclass
import utils_COM_logging as COMlogging
import utils_COM_checkreg as UtilsCOMcheckreg
import utils_COM_decorators as UtilsCOMdecorators
# import utils_COM_typelib as UtilsCOMTypeLib -> imported and used in UtilsCOMcheckreg
