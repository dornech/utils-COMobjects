# COM object utilities
# check and register utilities for COM class

"""
Module provides comprehensive checks for Python classes for registration as COM object
and registration control.
Checks include COM classes for DesignatedWrapPolicy (or EventHandlerPolicy based
hereupon) without as well as with typelib attributes.
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
# ruff: noqa: B007, B904, E501, E721, PLR1702, PLR1714, PLR5501, PLR6201, SIM102
#
# disable mypy errors
# - mypy error "'object' has no attribute 'xyz' [attr-defined]" when accessing attributes of
#   dynamically bound wrapped COM object
# mypy: disable-error-code = attr-defined

# fmt: off



from typing import Any
from collections.abc import Callable

import types
import inspect

import sys
import os.path
# import pathlib

import win32com.server
import winreg
from uuid import UUID

import utils_COM_classes as COMclass
import utils_COM_typelib as UtilsCOMTypeLib



# exception class
class ErrorCOMregistration(Exception):
    pass



# check COM registration readiness


def getmodulefile() -> str:
    """
    getmodulefile - get filename of module

    Returns:
        str: filename of calling module
    """

    if sys.argv[0] != "-c":
        script = sys.argv[0]
    else:
        stack = inspect.stack()
        script = stack[len(inspect.stack()) - 2].filename
    return os.path.basename(script).split(".")[0]
    # return pathlib.Path(script).name.split(".")[0]

def is_valid_uuid(uuid_to_test: str, version=4) -> bool:
    """
    is_valid_uuid - check uuid is valid

    alternative: use is_uuid from validator-collection

    Args:
        uuid_to_test (str): UUID to test.
        version (int): UUID-verison. Default is 4.

    Returns:
        bool: uuid is valid or not
    """

    try:
        uuid_obj = UUID(uuid_to_test, version=version)
    except ValueError:
        return False
    return (str(uuid_obj).upper() == uuid_to_test.upper()) or (str(uuid_obj).upper() == uuid_to_test[1:-1].upper())

def checkRegKey(regroot: int, regpath: str) -> bool:
    """
    checkRegKey - check registry key

    Args:
        regroot (int): registry root
        regpath (str): registry path to check

    Returns:
        bool: chec k result if registry path exists under root (or not)
    """

    try:
        with winreg.OpenKey(regroot, regpath) as regkey:
            winreg.CloseKey(regkey)
            return True
    except OSError:
        return False

def checkCOMattrib(
    cls: type[COMclass.baseCOMclass],
    attrib: str,
    checkfunction: Callable | None = None,
    optional: bool = False
) -> bool:
    """
    checkCOMattrib - check single COM class registration attribute

    Args:
        cls (COMclass.baseCOMclass): class to be checked
        attrib (str): attribute to be checked
        checkfunction (Callable): specific check function
        optional (bool): positive check result optional (i. e. if True only message is generated)

    Returns:
        bool: check result
    """

    if hasattr(cls, attrib):
        attribvalue = getattr(cls, attrib)
        if not isinstance(attribvalue, property):
            if checkfunction is not None:
                if checkfunction(attribvalue):
                    return True
                else:
                    print(f"attribute {cls.__name__}.{attrib} does not have a valid value.")
                    return False
            else:
                return True
        else:
            print(f"attribute {cls.__name__}.{attrib} is not initialized.")
            return False
    else:
        print(f"{cls.__name__} does not have attribute {attrib} required for COM registration.")
        return False or optional


def checkAttribsCOM(cls: type[COMclass.baseCOMclass], checkpubattrib: bool = False) -> bool:
    """
    checkAttribsCOM - check COM registration readiness, basic COM object attributes

    Args:
        cls (type[COMclass.baseCOMclass]): Python COM object class to be checked
        checkpubattrib (bool, optional): flag if _public_attrs_ is to be checked

    Returns:
        bool: check result
    """

    # check class_spec
    def is_correct_class_spec(cls, class_spec):
        return class_spec == getmodulefile() + "." + cls.__name__

    check = True

    # check attributes relevant for COM registration
    check = check and checkCOMattrib(cls, "_reg_clsid_", is_valid_uuid)
    check = check and checkCOMattrib(cls, "_reg_progid_")
    # value for _reg_class_spec_ is module name containing class + class name, path is needed as well
    # NOTE: latest win32com.server.register module generates attribute automatically
    # is_correct_class_spec_for_cls = functools.partial(is_correct_class_spec, cls)
    # check = check and checkCOMattrib(cls, "_reg_class_spec_", is_correct_class_spec_for_cls, optional=True)
    check = check and checkCOMattrib(cls, "_reg_desc_")
    check = check and checkCOMattrib(cls, "_reg_clsctx_")

    # check registration of methods in _public_methods_
    # 1. complete registration of public methods
    # 2. no registration of 'private' methods not intended for public use (i. e. preceeded by "_")
    # NOTE: linked to DesignatedWrapPolicy or EventHandlerPolicy because only these policies support the standard
    # mapping for IDispatch
    if cls._reg_policy_spec_ == "DesignatedWrapPolicy" or cls._reg_policy_spec_ == "EventHandlerPolicy":

        # check functions/methods if registration for access form outside in line with Python naming conventions
        # method_members = inspect.getmembers(cls, inspect.isfunction) + inspect.getmembers(cls, inspect.ismethod)
        method_members = inspect.getmembers(cls, lambda cls_member: inspect.isfunction(cls_member) or inspect.ismethod(cls_member))
        for member in method_members:
            if member[0][0:2] != "__":
                # check on types.xxx potentially double-checking after correction of predicate in getmembers
                if type(member[1]) in (types.FunctionType, types.MethodType):
                    if member[0] not in cls._public_methods_:
                        if member[0][0:1] == "_":
                            print(f"method {cls.__name__}.{member[0]} marked as private (i. e. not registered in {cls.__name__}._public_methods_ and not accessible from outside).")
                        else:
                            print(f"method {cls.__name__}.{member[0]} not registered in {cls.__name__}._public_methods_. Please check and re-register.")
                            check = False
                    elif member[0][0:1] == "_":
                        print(f"method {cls.__name__}.{member[0]} marked as private and registered in {cls.__name__}._public_methods_. Please check and re-register.")
                        check = False

        # check entries in _public_methods_ if function/method exists
        method_members_dict = dict(method_members)
        for member in cls._public_methods_:
            if member not in method_members_dict:
                print(f"method {cls.__name__}.{member} registered in {cls.__name__}._public_methods_ but does not exist. Please check and re-register.")
                check = False
            elif type(method_members_dict[member]) != types.FunctionType and type(method_members_dict[member]) != types.MethodType:
                print(f"method {cls.__name__}.{member} registered in {cls.__name__}._public_methods_ is not a valid method. Please check and re-register.")
                check = False

        # check _public_attrs_
        if checkpubattrib:
            # tempinstance = cls()
            # invalid_attrib = [attrib for attrib in cls._public_attrs_ if not hasattr(tempinstance, attrib)]
            invalid_attrib = [attrib for attrib in cls._public_attrs_ if not hasattr(cls, attrib)]
            if len(invalid_attrib) > 0:
                print(f"attributes {invalid_attrib} are not valid attributes of {cls.__name__}. Please check and re-register.")
                check = False

        # check  _readonly_attrs_
        invalid_readonly = [attrib for attrib in cls._readonly_attrs_ if attrib not in cls._public_attrs_]
        if len(invalid_readonly) > 0:
            print(f"read-only attribute(s) {invalid_readonly} not registered in {cls.__name__}._public_attributes_. Please check and re-register.")
            check = False

    else:

        print(f"{cls.__name__} does not have DesignatedWrapPolicy or EventHandlerPolicy assigned to support PyWin32 standard mappings.")
        check = False

    return check

def check_attribs_COM(cls: type[COMclass.baseCOMclass], checkpubattrib: bool = False) -> bool:
    """
    check_attribs_COM - check COM registration readiness, basic COM object attributes

    Args:
        cls (type[COMclass.baseCOMclass]): Python COM object class to be checked
        checkpubattrib (bool, optional): flag if _public_attrs_ is to be checked

    Returns:
        bool: check result
    """
    return checkAttribsCOM(cls, checkpubattrib)


def checkAttribsTypeLib(
    cls: type[COMclass.baseCOMclass] | type[COMclass.typelibCOMclass],
    clsmodule: types.ModuleType | None = None
) -> bool:
    """
    checkAttribsTypeLib - check COM registration readiness, typelib registration attributes

    Args:
        cls (Union[type[COMclass.baseCOMclass], type[COMclass.typelibCOMclass]]): Python COM object class to be checked
        clsmodule (types.ModuleType, optional): module object containing class definition. Defaults to None.

    Returns:
        bool: check result
    """

    def print_clslist(cls, clslist: list[tuple[str, Any]]):

        print(f"List of classes (including typelib file and typelib UUID) in conflict with {cls.__name__}:")
        for clslistid, clsconflicting in clslist:
            print(f"{clsconflicting.__name__} \t {clsconflicting._reg_typelib_filename_} \t {clsconflicting._typelib_guid_}")

    check = True

    # check attributes relevant for COM registration
    check = check and checkCOMattrib(cls, "_typelib_guid_", is_valid_uuid)
    if checkCOMattrib(cls, "_com_interfaces_"):
        if len(cls._com_interfaces_) > 1:
            print(f"Relationship between interface and Python class is not 1:1 for {cls.__name__}.")
            check = False
    else:
        check = False
    check = check and checkCOMattrib(cls, "_typelib_interfaceID_", is_valid_uuid)
    if checkCOMattrib(cls, "_reg_typelib_filename_"):
        if (cls._reg_typelib_filename_ != getmodulefile() + ".tlb") and (cls._reg_typelib_filename_ != cls.__name__ + ".tlb") and (cls._reg_typelib_filename_ != ""):
            print(f"Typelib filename registered in {cls.__name__} is '{cls._reg_typelib_filename_}'. Must be module name or class name (default) plus extension '.tlb'.")
            check = False
    else:
        check = False
    if checkCOMattrib(cls, "_typelib_name_"):
        if cls._typelib_name_ != getmodulefile() and cls._typelib_name_ != cls.__name__ and cls._typelib_name_ == "":
            print(f"Typelib name registered in {cls.__name__} is '{cls._typelib_name_}'. Should be module name or class name (default) plus extension '.tlb'.")
            # check = False
    else:
        check = False
    if checkCOMattrib(cls, "_reg_policy_spec_"):
        if not (cls._reg_policy_spec_ == "DesignatedWrapPolicy" or cls._reg_policy_spec_ == "EventHandlerPolicy"):
            print(f"Policy {cls._reg_policy_spec_} registered for {cls.__name__} not suitable for typelib generation.")
            check = False
    else:
        check = False

    # find module object for inspect.getmembers(cls) - not directly accessible otherwise
    # assumption: checker is called from module containing COM class to be checked
    if clsmodule is None:
        callerframeinfo = inspect.stack()[1]
        clsmodule = inspect.getmodule(callerframeinfo.frame)

    # check same typelib for all classes assigned to same typelib file
    if hasattr(cls, "_reg_typelib_filename_") and hasattr(cls, "_typelib_name_"):
        # instance property means not initialized (see ABC class definition with property decorator)
        # if not isinstance(getattr(cls, "_reg_typelib_filename_"), property) and not isinstance(getattr(cls, "_typelib_name_"), property):

        try:
            clslist_wrongname = inspect.getmembers(
                clsmodule,
                lambda cls_member:
                    (getattr(cls_member, "__name__", "") != cls.__name__) and
                    (getattr(cls_member, "_reg_typelib_filename_", "") == cls._reg_typelib_filename_) and
                    (getattr(cls_member, "_typelib_name_", "") != cls._typelib_name_) and
                    inspect.isclass(cls_member)
            )
        except BaseException:
            err_msg = f"COM registration could not be done. Error checking for same typelib name for classes assigned to same typelib {cls._reg_typelib_filename_} have same typelib name."
            raise ErrorCOMregistration(err_msg)
        else:
            if len(clslist_wrongname) > 0:
                print(f"COM registration could not be done. Not all classes assigned to typelib {cls._reg_typelib_filename_} have same typelib name.")
                print_clslist(cls, clslist_wrongname)
                check = False

    if hasattr(cls, "_reg_typelib_filename_") and hasattr(cls, "_typelib_guid_"):
        # instance property means not initialized (see ABC class definition with property decorator)
        # if not isinstance(getattr(cls, "_reg_typelib_filename_"), property) and not isinstance(getattr(cls, "_typelib_guid_"), property):

        # check same typelib GUIDs for all classes assigned to same typelib file
        try:
            clslist_wronguuid = inspect.getmembers(
                clsmodule,
                lambda cls_member:
                    (getattr(cls_member, "__name__", "") != cls.__name__) and
                    (getattr(cls_member, "_reg_typelib_filename_", "") == cls._reg_typelib_filename_) and
                    (getattr(cls_member, "_typelib_guid_", "") != cls._typelib_guid_) and
                    inspect.isclass(cls_member)
            )
        except BaseException:
            err_msg = f"COM registration could not be done. Error checking for same typelib GUID for classes assigned to same typelib {cls._reg_typelib_filename_} have same uuid."
            raise ErrorCOMregistration(err_msg)
        else:
            if len(clslist_wronguuid) > 0:
                print(f"COM registration could not be done. Not all classes assigned to typelib {cls._reg_typelib_filename_} have same uuid.")
                print_clslist(cls, clslist_wronguuid)
                check = False

        # check typelib GUID is unique
        # -> assumption that checked by typelib generator and/or registry not assumed valid
        try:
            clslist_duplicateuuid = inspect.getmembers(
                clsmodule,
                lambda cls_member:
                    (getattr(cls_member, "__name__", "") != cls.__name__) and
                    (getattr(cls_member, "_reg_typelib_filename_", "") != cls._reg_typelib_filename_) and
                    (getattr(cls_member, "_typelib_guid_", "") == cls._typelib_guid_) and
                    inspect.isclass(cls_member)
            )
        except BaseException:
            err_msg = "COM registration could not be done. Error checking for same typelib GUID for different typelibs."
            raise ErrorCOMregistration(err_msg)
        else:
            if len(clslist_duplicateuuid) > 0:
                print(f"COM registration could not be done. GUID {cls._typelib_guid_} assigned to {cls._reg_typelib_filename_} not unique.")
                print_clslist(cls, clslist_duplicateuuid)
                check = False

    # check interface GUID is unique
    # -> assumption that checked by typelib generator and/or registry not assumed valid
    if hasattr(cls, "_typelib_interfaceID_"):
        # instance property means not initialized (see ABC class definition with property decorator)
        # if not isinstance(getattr(cls, "_typelib_interfaceID_"), property):

        try:
            clslist_duplicateIID = inspect.getmembers(
                clsmodule,
                lambda cls_member:
                    (getattr(cls_member, "__name__", "") != cls.__name__) and
                    (getattr(cls_member, "_typelib_interfaceID_", "") == cls._typelib_interfaceID_) and
                    inspect.isclass(cls_member)
            )
        except BaseException:
            err_msg = "COM registration could not be done. Error checking if Typelib interface GUIDs are unique."
            raise ErrorCOMregistration(err_msg)
        else:
            if len(clslist_duplicateIID) > 0:
                print(f"COM registration could not be done. GUID {cls._typelib_guid_} assigned to {cls._reg_typelib_filename_} not unique.")
                print_clslist(cls, clslist_duplicateIID)
                check = False

    return check

def check_attribs_typelib(
    cls: type[COMclass.baseCOMclass] | type[COMclass.typelibCOMclass],
    clsmodule: types.ModuleType | None = None
) -> bool:
    """
    check_attribs_typelib - check COM registration readiness, typelib registration attributes

    Args:
        cls (Union[type[COMclass.baseCOMclass], type[COMclass.typelibCOMclass]]): Python COM object class to be checked
        clsmodule (types.ModuleType, optional): module object containing class definition. Defaults to None.

    Returns:
        bool: check result
    """
    return checkAttribsTypeLib(cls, clsmodule)


def processCOMregistration(cls: type[COMclass.baseCOMclass], gentypelib: bool = False, testmode: bool = False) -> None:
    """
    processCOMregistration - check and register Python COM object class as COM object

    To register class add following call in object module:
    if __name__ == '__main__':
        <import name of this module>.processCOMregistration(<classname>)

    Args:
        cls (type[COMclass.baseCOMclass]): Python COM object class to be registered
        gentypelib (bool, optional): activate typelib generation. Defaults to False.
        testmode (bool, optional): Test only. Defaults to False.
    """

    def errorhandling(msg: str, testmode: bool):
        if testmode:
            print(msg)
        else:
            raise ErrorCOMregistration(msg)

    # check if class with CLSID registered
    def checkRegistryCLSID(clsid) -> bool:

        regpath = fr'\CLSID\{clsid}'  # for HKEY_CLASSES_ROOT
        # regpath = fr'SOFTWARE\Classes\CLSID\{clsid}'   # for HKEY_LOCAL_MACHINE
        return checkRegKey(winreg.HKEY_CLASSES_ROOT, regpath)

    # check if typelib with TlbID registered
    def checkRegistryTypelibID(tlbid):

        regpath = fr'\TypeLib\{tlbid}'  # for HKEY_CLASSES_ROOT
        # regpath = fr'SOFTWARE\Classes\TypeLib\{tlbid}'   # for HKEY_LOCAL_MACHINE
        return checkRegKey(winreg.HKEY_CLASSES_ROOT, regpath)

    # check if parameter is class (not class instance!)
    modeprefix = ""
    modepostfix = ""
    if hasattr(cls, "__name__"):
        if "--unregister" not in sys.argv:
            if "--debug" in sys.argv:
                modepostfix = " (debug-mode)"
            # check COM registrability
            print(f"Check settings for COM object registration of {cls.__name__} ...")
            # if not unregister run and basic COM attributes OK, do typelib processing
            # (unregistering typelibs happens automatically)
            if checkAttribsCOM(cls):
                # check if class is typelib generation relevant
                if COMclass.typelibCOMclass in cls.__bases__:
                    if gentypelib:
                        # check TypeLib attributes
                        print(f"Check settings for TypeLib generation and registration of {cls.__name__} ...")
                        # determine calling module (assumption: contains class definitions for cross-checks)
                        callerframeinfo = inspect.stack()[1]
                        clsmodule = inspect.getmodule(callerframeinfo.frame)
                        if checkAttribsTypeLib(cls, clsmodule):
                            if checkRegistryTypelibID(cls._typelib_guid_) and "--unregister" not in sys.argv:
                                err_msg = f"Typelib with GUID {cls._typelib_guid_} already registered."
                                raise ErrorCOMregistration(err_msg)
                            else:
                                # add call for IDL creation
                                if UtilsCOMTypeLib.generateIDL(cls, clsmodule):
                                    # compile TypeLib
                                    UtilsCOMTypeLib.compileTypeLib(cls)
                                    # register TypeLib
                                    if not testmode:
                                        UtilsCOMTypeLib.registerTypeLib(cls)
                                    else:
                                        print("Update of registry for typelib registration not activated / surpressed via parameter.")
                                else:
                                    print(f"IDL for typelib generation for {cls.__name__} generated with errors. Please check.")
                        else:
                            errorhandling(f"{cls.__name__} does not have all necessary and valid attributes set for creating and registering typelib.", testmode)
                    else:
                        print(f"{cls.__name__} has attributes for typelib generation but generation not activated.")
                if checkRegistryCLSID(cls._reg_clsid_):
                    errorhandling(f"Class with GUID {cls._reg_clsid_} already registered.", testmode)
                    return
            else:
                errorhandling(f"{cls.__name__} does not have necessary and valid attributes for registering as COM object.", testmode)
        else:
            modeprefix = "un"
        # analyse call stack if called via Python interpreter using switch -c
        # required to identify correct python module
        if sys.argv[0] == "-c":
            stack = inspect.stack()
            sys.argv[0] = stack[len(inspect.stack()) - 2].filename
        # register class as COM object
        print(f"process {modeprefix}registration{modepostfix} of COM server {cls._reg_progid_} ...")
        # Utils.logCLIargs()
        if not testmode:
            win32com.server.register.UseCommandLine(cls)
        else:
            print("Update of registry for COM object registration not activated / surpressed via parameter.")
    else:
        errorhandling(f"COM registration requested for instance not object class. Try registration for {cls.__class__.__name__}.", testmode)
    print()

def process_COM_registration(cls: type[COMclass.baseCOMclass], gentypelib: bool = False, testmode: bool = False) -> None:
    """
    process_COM_registration - check and register Python COM object class as COM object

    Args:
        cls (type[COMclass.baseCOMclass]): Python COM object class to be registered
        gentypelib (bool, optional): activate typelib generation. Defaults to False.
        testmode (bool, optional): Test only. Defaults to False.
    """
    processCOMregistration(cls, gentypelib, testmode)


def printCOMpublicmethods(cls: COMclass.baseCOMclass | COMclass.typelibCOMclass | Any) -> None:
    """
    printCOMpublicmethods - print public methods of Python COM object class

    Args:
        cls (Union[COMclass.baseCOMclass, COMclass.typelibCOMclass, object]): Python COM object class
    """

    method_members = inspect.getmembers(cls, inspect.ismethod(cls))  # type: ignore
    for member in method_members:
        if member[0][0:2] != "__":
            if type(member[1]) == types.FunctionType or type(member[1]) == types.MethodType:
                if member[0] in cls._public_methods_:
                    print(member[0], type(member[1]), type(cls.__dict__[member[0]]))
                    print(f"    Signature: {inspect.signature(member[1])}")
                    print(f"    {inspect.getfullargspec(member[1])}")

def print_COM_publicmethods(cls: COMclass.baseCOMclass | COMclass.typelibCOMclass | object) -> None:
    """
    print_COM_publicmethods - print public methods of Python COM object class

    Args:
        cls (Union[COMclass.baseCOMclass, COMclass.typelibCOMclass, object]): Python COM object class
    """
    printCOMpublicmethods(cls)
