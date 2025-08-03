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



from importlib.metadata import PackageNotFoundError, version

# try:
#     __version__ = version('utils-COMobjects')
# except PackageNotFoundError:  # pragma: no cover
#     __version__ = 'unknown'
# finally:
#     del version, PackageNotFoundError

# up-to-date version tag for modules installed in editable mode inspired by
# https://github.com/maresb/hatch-vcs-footgun-example/blob/main/hatch_vcs_footgun_example/__init__.py
# Define the variable '__version__':
try:

    # own developed alternative variant to hatch-vcs-footgun overcoming problem of ignored setuptools_scm settings
    # from hatch-based pyproject.toml libraries
    from hatch.cli import hatch
    from click.testing import CliRunner
    # determine version via hatch
    __version__ = CliRunner().invoke(hatch, ["version"]).output.strip()

except (ImportError, LookupError):
    # As a fallback, use the version that is hard-coded in the file.
    try:
        from ._version import __version__  # noqa: F401
    except ModuleNotFoundError:
        # The user is probably trying to run this without having installed the
        # package, so complain.
        raise RuntimeError(
            f"Package {__package__} is not correctly installed. Please install it with pip."
        )



import sys
import os.path

# switch os-path -> pathlib
sys.path.insert(1, os.path.dirname(os.path.realpath(__file__)))
# sys.path.insert(1, str(pathlib.Path(__file__).resolve().parent))

import utils_COM_classes as COMclass
import utils_COM_logging as COMlogging
import utils_COM_checkreg as UtilsCOMcheckreg
import utils_COM_decorators as UtilsCOMdecorators
# import utils_COM_typelib as UtilsCOMTypeLib -> imported and used in UtilsCOMcheckreg
