# COM object utilities
# COM class - abstract classes for generating a COM object from Python


# NOTE:
# when executed as registered COM object, this module must be in the correct relative path to any
# registered Python COM object using this module (if not contained in a standard library path)
# Generic example:
# - path to registered COM object script:   xxx\SampleObject.py
# - path to abstract base class script:     xxx\utils_COMobjects\utils_COM_classes.py


# baseCOMclass()
# - abstract class for DesignatedWrapPolicy (or EventHandlerPolicy based hereupon)
# - enforces mandatory attributes to be set in derived class

# typelibCOMclass()
# - abstract class based on typelibCOMclass derived from baseCOMclass
# - enforces additional mandatory attributes for typelib generation to be set in derived class

# NOTE: NEVER copy GUIDs f. e. for class IDs !
# Use "print(pythoncom.CreateGuid())" to make a new one.

# further support modules:
# - utils_COM_logging.py     / COMlogging          : mixin class for logging
# - utils_COM_checkreg.py    / UtilsCOMcheckreg    : consistency check for COM objects
# - utils_COM_decorators.py  / UtilsCOMdecorators  : decorators
# - utils_COM_typeLib.py     / UtilsCOMTyprLib     : typelib support (experimental)


"""
Module provides to abstract base classes for writing COM object modules in Python.
"""


# ruff and mypy per file settings
#
# empty lines
# ruff: noqa: E302, E303
# naming conventions
# ruff: noqa: N801, N802, N803, N806, N812, N813, N815, N816, N818, N999
# others
# ruff: noqa: PLW3201, RUF012
#
# disable mypy errors
# - "'object' has no attribute 'xyz' [attr-defined]" when accessing attributes of
#   dynamically bound wrapped COM object
# - "Only instance methods can be decorated with @property"
# mypy: disable-error-code = "attr-defined, misc"

# fmt: off



from abc import ABC, abstractmethod

import winreg
import pythoncom
# import win32com.server.register
# from win32com.server.exception import COMException



# base class without typelib for DesignatedWrapPolicy or EventHandlerPolicy based hereupon
class baseCOMclass(ABC):
    """
    baseCOMclass - abstract base class for COM objects in Python

    baseCOMclass serves as abstract class for DesignatedWrapPolicy (or
    EventHandlerPolicy based hereupon) COM classes.

    The abstract base class enforces mandatory COM object specific
    attributes to be set in a derived class by implementing them as
    abstract class method  marked as property and raising a
    NonImplementedError if not defined in a class based hereupon.

    The abstract base class presets mandatory COM attributes that are not
    specific for an implementation i.e. not mandatory to be changed.

    For additional optional attributes see comments in the code.

    Please note comments in code also.
    """

    # mandatory attributes from abstract class for COM object registration
    # class specific change mandatory

    @property
    @classmethod
    @abstractmethod
    def _reg_clsid_(cls):
        return NotImplementedError

    @property
    @classmethod
    @abstractmethod
    def _reg_progid_(cls):
        return NotImplementedError

    # mandatory attributes from abstract class for COM object registration
    # class specific change optional (however required f. e. for _public_methods_)

    _reg_desc_ = "Python COM Server"
    _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER

    # "DesignatedWrapPolicy" is mandatory for using _public_methods_, _public_attrs_ and _readonly_attrs_
    # in IDispatch mapping and for support of typelib generation
    _reg_policy_spec_ = "DesignatedWrapPolicy"

    # public methods
    # _public_methods_ = []
    @property
    @classmethod
    @abstractmethod
    def _public_methods_(cls):
        """ _public_methods_- property """
        return NotImplementedError

    # For class attributes be aware of Python handling:
    # - if attribute is an IMMUTABLE object (i.e. an integer), the class attribute is not changed
    #   for all class instances but changing the class attribute implicitly turns the class attribute
    #   into an instance attribute.
    # - if attribute is an MUTABLE object (i.e. a list), changing the class attribute with its
    #   object methods changes the class attribute for all class instances.
    # However, for COM objects this means that class attributes cannot be changed from outside without
    # implementing public methods in the COM object class itself to change a mutable class attribute
    # object (i.e. passed parameters are always immutable objects).

    _public_attrs_: list[str] = ["logcalls"]
    _readonly_attrs_: list[str] = []

    # optional attributes for COM object registration
    # _reg_verprogid_               - suggested value _reg_progid_ + ".1"
    # _reg_class_spec_              - value is module name containing class + class name,
    #                                 i. e. = __name__ + ".<COM class name in Python>"
    #                                 NOTE: path is needed, latest win32com.server.register module generates
    #                                 attribute automatically
    # _reg_threading_               - Both -> default
    # _reg_options_                 - no suggestion
    # _reg_remove_keys_             - no suggestion
    # _reg_debug_dispatcher_spec_   - no suggestion

    # flag for call logging (see respective decorator)
    logcalls: bool = True

    #  method to report registration mode for COM caller with other bitness
    @classmethod
    def _checkDebug(cls) -> bool:

        regpath = fr"\CLSID\{cls._reg_clsid_}\Debugging"  # for HKEY_CLASSES_ROOT
        # regpath = fr"SOFTWARE\Classes\CLSID\{cls._reg_clsid_}\Debugging"   # for HKEY_LOCAL_MACHINE
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, regpath) as regkey:
                debug = winreg.QueryValue(regkey, None)
                return (debug == "1")
        except OSError:
            return False

    # For COM object methods please note that all strings which come from COM will
    # actually be unicode objects rather than string objects according to
    # http://timgolden.me.uk/pywin32-docs/html/com/win32com/HTML/QuickStartServerCom.html
    # This may require special precautions.

    # For method implementation note difference between high-level call interface and
    # low-level call-interface supported by the Python comtypes module. While the
    # high-level call interface is the "normal" Python call interface with limited control
    # over the HRESULT returned, the low-level call allows setting the HRESULT:
    # - high-level : def MyMethod(self, <param 1>, ... <param N>): return <result>
    # - low-level  : def MyMethod(self, this, <param 1>, ... <param N>, presult): presult[0] = <result>: return HRESULT
    # for further reference refer to:
    # https://pythonhosted.org/comtypes/server.html


# base class with typelib
class typelibCOMclass(baseCOMclass):
    """
    typelibCOMclass - abstract base class for COM objects with typelib in Python

    typelibCOMclass is a class derived from baseCOMclass and as such an
    abstract base class as well.

    The abstract base class enforces additional mandatory attributes for
    typelib generation to be set in derived class.

    The abstract base class presets mandatory COM attributes that are not
    specific for an implementation i.e. not mandatory to be changed.

    Please note comments in code also.
    """

    # NOTE limitations for classes with typelib generation:
    # - type hints / annotations are required for classes building COM objects with _typelib_guid_
    # - only one return value is currently allowed in typelib

    # mandatory attributes from abstract class for typelib generation
    # class specific change mandatory
    # _typelib_name_ is an own-defined attribute for the name of the typelib
    # _typelib_interfaceID_ is an own-defined attribute for identifying interface of a class

    @property
    @classmethod
    @abstractmethod
    def _reg_typelib_filename_(cls):
        return NotImplementedError

    @property
    @classmethod
    @abstractmethod
    def _typelib_name_(cls):
        return NotImplementedError

    @property
    @classmethod
    @abstractmethod
    def _typelib_guid_(cls):
        return NotImplementedError

    @property
    @classmethod
    @abstractmethod
    def _com_interfaces_(cls):
        return NotImplementedError

    @property
    @classmethod
    @abstractmethod
    def _typelib_interfaceID_(cls):
        return NotImplementedError

    # mandatory attributes from abstract class for typelib registration
    # class specific change optional

    _typelib_version_ = 1, 0
    # _typelib_lcid_ = LOCALE_USER_DEFAULT
