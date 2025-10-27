# COM object utilities
# COM mixin class for logging - simulate stdout, stderr for COM objects


# mixinCOMclass_logger(object)
# - mixin class for logging
# - simulate stdout, stderr for COM objects
# - Note: simple redirection does not work!
#
# alternative: normal logging with logging.exception
# https://stackoverflow.com/questions/1508467/log-exception-with-traceback-in-python


"""
Module provides a mixin class to enable logging for Python-based COM object.
"""


# ruff and mypy per file settings
#
# empty lines
# ruff: noqa: E302, E303
# naming conventions
# ruff: noqa: N801, N802, N803, N806, N812, N813, N815, N816, N818, N999
# boolean-type arguments
# ruff: noqa: FBT001, FBT002
# others
# ruff: noqa: B019, E115, E501, DTZ005, PLW1514, SIM102, SIM115
#
# disable mypy errors
# mypy: disable-error-code = "attr-defined, union-attr"

# fmt: off



from typing import ClassVar

import inspect
import functools

import os.path
# import pathlib

import tempfile
import traceback
from datetime import datetime

import utils_mystuff as Utils



class mixinCOMclass_logger:
    """
    mixin class for logging

    - simulate stdout, stderr for COM objects (Note: simple redirection does not work!)
    - logging via logging object and logging.exception
    """

    # public methods to be registered:
    _public_methods_addon: ClassVar[list[str]] = []

    @functools.lru_cache(5)
    def _basename_stdXX_log(self, prefix: str = "", postfix: str = ""):

        # loop to determine correct prefix derived from COM object class module
        # static mechanism does not work in relevant call situations
        # a) run from IDE, debug in IDE c) COM call
        if prefix != "":
            prefix += "_"
            prefix.replace("__", "_")
        if postfix != "":
            postfix += "_"
            postfix.replace("__", "_")
        basefilenameprefix = ""
        basefilenameprefix = prefix + os.path.basename(inspect.getsourcefile(self.__class__)).split(".")[0] + "_" + postfix  # type: ignore[type-var]
        # basefilenameprefix = prefix + pathlib.Path(inspect.getsourcefile(self.__class__)).name.split(".")[0] + "_" + postfix  # type: ignore[type-var]
        return basefilenameprefix

    @staticmethod
    def _timestamp_stdXX():
        return f"\n\n{datetime.now()}--------------------\n\n"

    # private variables to COM object for stdout, stderr simulation
    _stdoutCOM = None
    _stdoutCOM_wroteTS = False
    _stderrCOM = None
    _stderrCOM_wroteTS = False

    def _open_stdoutCOM(self):

        if self._stdoutCOM is None:
            stdoutfilename = tempfile.gettempdir() + "\\" + self._basename_stdXX_log() + "stdoutCOM.txt"
            if os.path.exists(stdoutfilename):
            # if pathlib.Path(stdoutfilename).exists():
                self._stdoutCOM = open(stdoutfilename, "a")
            else:
                self._stdoutCOM = open(stdoutfilename, "w+")

    def _write2stdoutCOM(
        self, text: str, force_open: bool = True, force_close: bool = False, write_timestamp: bool = True,
        firstonly: bool = True
    ):

        opened = False
        if self._stdoutCOM is None and force_open:
            self._open_stdoutCOM()
            opened = True
        if self._stdoutCOM is not None:
            if write_timestamp:
                if not firstonly or not self._stdoutCOM_wroteTS:
                    self._stdoutCOM.write(self._timestamp_stdXX())
                    self._stdoutCOM_wroteTS = True
            if text is not None:
                self._stdoutCOM.write(text)
        if opened or force_close:
            self._close_stdoutCOM()

    def _close_stdoutCOM(self):

        if self._stdoutCOM is not None:
            self._stdoutCOM.close()
            self._stdoutCOM = None


    def _open_stderrCOM(self):

        if self._stderrCOM is None:
            stderrfilename = tempfile.gettempdir() + "\\" + self._basename_stdXX_log() + "stderrCOM.txt"
            if os.path.exists(stderrfilename):
            # if pathlib.Path(stderrfilename).exists():
                self._stderrCOM = open(stderrfilename, "a")
            else:
                self._stderrCOM = open(stderrfilename, "w+")

    def _write2stderrCOM(
        self, text: str, force_open: bool = True, force_close: bool = False, write_timestamp: bool = True,
        firstonly: bool = True
    ):

        opened = False
        if self._stderrCOM is None and force_open:
            self._open_stderrCOM()
            opened = True
        if self._stderrCOM is not None:
            if write_timestamp:
                if not firstonly or not self._stderrCOM_wroteTS:
                    self._stderrCOM.write(self._timestamp_stdXX())
                    self._stderrCOM_wroteTS = True
            if text is not None:
                self._stderrCOM.write(text)
        if opened or force_close:
            self._close_stderrCOM()

    def _traceback2stderrCOM(self, exc: Exception, write_timestamp: bool = True, firstonly: bool = True):

        self._open_stderrCOM()
        if not firstonly or not self._stderrCOM_wroteTS:
            self._stderrCOM.write(self._timestamp_stdXX())
            self._stderrCOM_wroteTS = True
        traceback.print_exception(type(exc), exc, exc.__traceback__, file=self._stderrCOM)

    def _close_stderrCOM(self):
        if self._stderrCOM is not None:
            self._stderrCOM.close()
            self._stderrCOM = None


    # logger via logging - used for call and exception logging

    # private variables to COM object for logging
    _logger = None

    def _initCOMlogger(self):

        if self._logger is None:
            loggername = self._basename_stdXX_log(prefix="Log")
            if loggername[-1] == "_":
                loggername = loggername[:-1]
            self._logger = Utils.initLogger(loggername=loggername, filename=tempfile.gettempdir() + "\\" + loggername + ".txt")

    def _logMessage(self, text: str):

        if self._logger is None:
            self._initCOMlogger()
        self._logger.info(text)

    def _logException(self, exception: Exception):

        if self._logger is None:
            self._initCOMlogger()
        self._logger.exception(exception)

    def _shutdownCOMlogger(self):

        if self._logger is not None:
            for handler in self._logger.handlers:
                handler.close()
