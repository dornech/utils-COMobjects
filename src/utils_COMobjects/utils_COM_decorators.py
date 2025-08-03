# COM object utilities
# decorator for COM class calls

# decorators
# - decorator for supporting named parameters when calling from outside
# - call logger (main purpose is to analyse changes applied by decorator for named parameters support


"""
Module providing decorators to support Python COM object programming.
Decorators provided allow a passing of parameters by name as well as a call logging.
"""


# ruff and mypy per file settings
#
# empty lines
# ruff: noqa: E302, E303
# naming conventions
# ruff: noqa: N801, N802, N803, N806, N812, N813, N815, N816, N818, N999
# others
# ruff: noqa: B009, PLR1702, PLR2004, SIM102

# fmt: off




from typing import Any

import inspect
import functools

from win32com.server.exception import COMException
import winerror



# decorator for calling methods with named parameters
# 1. convert tuples (param, value) to entry in kwargs dictionary
# 2. convert string parameter param:=value to entry in kwargs dictionary
# note: **kwargs is empty because not passed on by COM caller
def calltypewrapper(COMcall):
    """
    decorator to enable a more comfortable call of COM object methods.

    The decorator provides two modes:

    1. convert tuples (param, value) to entry in kwargs dictionary
    2. convert string parameter param:=value to entry in kwargs dictionary
    """

    @functools.wraps(COMcall)
    def wrapper_calltypewrapper(self, *args, **kwargs):

        def adddictfull(kwargdict, key: str, value: Any, classname: str, methodname: str, ):
            if key in kwargdict:
                raise COMException(
                    description=f"Duplicate parameter '{key}' calling '{methodname}' in '{classname}'.",
                    scode=winerror.E_FAIL,
                    source=f"call of '{methodname}' in '{classname}'."
                )
            else:
                kwargdict[key] = value

        # set caller information for COM exception for duplicate parameter error
        # watch out for issue with keyword arguments as described here:
        # https://stackoverflow.com/questions/24755463/functools-partial-wants-to-use-a-positional-argument-as-a-keyword-argument
        adddict = functools.partial(adddictfull, classname=self.__class__.__name__, methodname=COMcall.__name__)

        # process args
        args_new = []
        kwargs_new = {}
        COMcall_signature = inspect.signature(COMcall)
        for arg in args:
            if isinstance(arg, str):
                argsplit = arg.split(":=")
                if len(argsplit) == 2:
                    keyword = argsplit[0]
                    value = argsplit[1]
                    if keyword in COMcall_signature.parameters:
                        if COMcall_signature.parameters[keyword].annotation != inspect.Parameter.empty:
                            adddict(kwargs_new, keyword, COMcall_signature.parameters[keyword].annotation(value))
                        else:
                            try:
                                adddict(kwargs_new, keyword, float(value))
                            except BaseException:
                                adddict(kwargs_new, keyword, value)
                    else:
                        args_new.append(arg)
                else:
                    args_new.append(arg)
            elif isinstance(arg, tuple):
                if len(arg) == 2 and arg[0] in COMcall_signature.parameters:
                    adddict(kwargs_new, arg[0], arg[1])
                else:
                    args_new.append(arg)
            else:
                args_new.append(arg)

        return COMcall(self, *args_new, **kwargs_new)

    return wrapper_calltypewrapper


# decorator for logging - only if registered in debug mode or logging via class attribute "logcalls" activated
def logcall(COMcall):
    """
    logcall - decorator to activate logging of calls of an COM object method.

    Call logging is activated by

    - registering COM object in debug mode
    - set logcalls attribute in COM class
    """

    def calllogger(func, self, *args, **kwargs):

        args_repr = [repr(a) for a in args]
        kwargs_repr = [f"{k}={v!r}" for k, v in kwargs.items()]
        signature = ", ".join(args_repr + kwargs_repr)
        print(f"Calling {self.__class__.__name__}.{func.__name__}({signature})")
        if self._logger is not None:
            self._logger.info(f"Calling {self.__class__.__name__}.{func.__name__}({signature})")

    @functools.wraps(COMcall)
    def wrapper_logcall(self, *args, **kwargs):

        logged = False
        if hasattr(self, "_checkDebug"):
            if COMcall.__name__ != "_checkDebug":    # safeguard against endless recursive calling
                if self._checkDebug():
                    # log because registered in debug mode
                    calllogger(COMcall, self, *args, **kwargs)
                    logged = True
        if not logged and hasattr(self, "logcalls"):
            if getattr(self, "logcalls"):
                # log because log flag set
                calllogger(COMcall, self, *args, **kwargs)
                logged = True
        return COMcall(self, *args, **kwargs)

    return wrapper_logcall
