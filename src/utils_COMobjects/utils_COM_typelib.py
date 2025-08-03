# COM object utilities
# IDL utils (experimental)


# IDL generator
# - generate IDL, idea and source for template
#   https://exceldevelopmentplatform.blogspot.com/2019/08/vba-new-python-com-classes.html
# -  additional features and improvements see below
# - IDL / typelib is generated for all classes assigned to same typelib file
#   -> control via _reg_typelib_filename_
# - align order of generation according to
#   https://raw.githubusercontent.com/mhammond/pywin32/master/com/win32com/test/pippo.idl
#   https://docs.microsoft.com/en-us/windows/win32/com/anatomy-of-an-idl-file
# - integrated enumerations with Enum class, IDL see http://bytepointer.com/resources/hludzinski_understanding_idl.htm
# - intrinsic limitations:
#   type hints / annotations are required for classes building COM objects with _typelib_guid_
#   support only for "simple" types
#   only one return value is currently allowed in typelib
# - low-level calls for IDL generation not yet supported (currently only high-level calls are supported)


"""
Module provides experimental typelib support.
"""


# ruff and mypy per file settings
#
# empty lines
# ruff: noqa: E302, E303
# naming conventions
# ruff: noqa: N801, N802, N803, N806, N812, N813, N815, N816, N818, N999
# others
# ruff: noqa: B007, B009, E501, F841, PLR1702, PLR0914, PLW1514, Q003, S605, SIM102, SIM108, UP031
#
# disable mypy errors
# mypy: disable-error-code = "arg-type, truthy-function, var-annotated"

# fmt: off



# TODO Dev:
# - in/out parameters - multiple return parameters ?
#   https://social.msdn.microsoft.com/Forums/en-US/72494228-664d-48e9-b9e7-11378465a09a/return-multiple-values-from-functions-in-com-interfaces?forum=vcgeneral
#   Problems with VBA and multiple parameters
#   https://mail.python.org/pipermail/python-win32/2006-October/thread.html#5093
#   http://python.6.x6.nabble.com/Re-Input-only-vs-In-Out-Function-Parameters-td1950757.html#none
#   -> not needed for interfacing to VBA anyway?
# - support for "complex" types and SAFEARRAYs needed ?
#   https://flylib.com/books/en/4.346.1.32/1/

# TODO Test:
# - comprehensive test especially typelib generation stuff



from typing import Any

import types
import inspect

import sys
import os
# from os import system
# import pathlib
from distutils.dep_util import newer

import pythoncom



# typelib generation -  typelib per class or common typelib for all classes in py-file with same typelib assignment
def generateIDL(cls, clsmodule: types.ModuleType) -> bool:

    def setinterface(classname, cominterfaces: list[str]) -> str:
        if len(cominterfaces) == 0:
            return "I" + str(classname)
        else:
            return cominterfaces[0]

    def PythonTypeToIDLType(annotations, argname: str) -> str:
        argtypeIDL = None
        try:
            if argname in annotations:
                pythonargtype = annotations[argname]
                argtypePy = pythonargtype.__name__
                argtypeIDL = {
                    'bool' : "VARIANT_BOOL",
                    'str'  : "BSTR*",
                    'float': "double",
                    'int'  : "long",
                    ''     : "VARIANT*"
                }[argtypePy]
            else:
                argtypeIDL = "VARIANT*"
        except Exception as err:
            #  print("Error: " + str(err) + "\n")
            argtypeIDL = "VARIANT*"
        return argtypeIDL

    # list of all classes assigned to same typelib
    clslist = get_typelib_classes(cls, clsmodule)
    if len(clslist) > 1:
        print(f"Typelib generation not only for {cls.__name__} but all classes defined in {cls.__module__} assigned to {cls._reg_typelib_filename_}")
    # determine filename
    idlfile = get_filename(cls, ".idl")
    tlbfile = get_filename(cls, ".tlb")

    okIDL = True

    # update IDL if python module file changed
    if newer(sys.modules[cls.__module__].__file__, idlfile):

        # import IDL from other files
        idl = "// Generated .IDL file (by Python code)\n//\n" + f"// typelib filename: {tlbfile} \n\n"
        idl += "import \"oaidl.idl\";\nimport \"ocidl.idl\";\nimport \"unknwn.idl\";\n\n"

        for clslistmemberid, clslistmember in clslist:

            # determine interface name
            interface = setinterface(clslistmember.__name__, clslistmember._com_interfaces_)

            # interface header
            idl += "[\n"
            idl += "\tobject,\n"
            idl += "\tuuid(" + clslistmember._typelib_interfaceID_[1:-1] + "),\n"
            idl += "\tdual,\n"
            idl += "\thelpstring(\"" + interface + " Interface\"),\n"
            idl += "\tpointer_default(unique)\n"
            idl += "]\n"
            idl += "interface " + interface + " : IDispatch\n"

            # counter for interface items
            dispid = 1  # start from 1 as zero equates to default member
            idl += "{\n"

            # interface methods
            if hasattr(clslistmember, '_public_methods_'):
                clsmember = inspect.getmembers(clslistmember, inspect.isfunction or inspect.ismethod)
                # for method in clslistmember._public_methods_:
                for method in clsmember:
                    if method[0] in clslistmember._public_methods_:
                        idl += "\tid(" + str(dispid) + "), helpstring(\"method " + method[0] + "\")]\n"
                        idl += "\tHRESULT " + method[0] + "(\n"
                        fullargspec = inspect.getfullargspec(method[1])
                        arga = len(fullargspec.annotations)
                        if arga > 0:
                            argc = len(fullargspec.args)
                            if argc > 0:
                                if isinstance(cls.__dict__[method[0]], staticmethod):
                                    argidxstart = 0
                                else:
                                    argidxstart = 1
                                for argidx in range(argidxstart, argc):
                                    arg = fullargspec.args[argidx]
                                    argtype = PythonTypeToIDLType(fullargspec.annotations, arg)
                                    idl += "\t\t[in] " + argtype + " " + arg + ",\n"
                            argtype = PythonTypeToIDLType(fullargspec.annotations, "return")
                            # idl += "\t\t[out, retval] " + argtype + "* retval\n"
                            idl += "\t\t[out, retval] " + argtype
                            if argtype[-1] != "*":
                                idl += "*"
                            idl += " retval\n"
                        else:
                            print(f"Type information not available for {clslistmember.__name__}.{method[0]}. Please check IDL-File.")
                            okIDL = False
                        dispid += 1
                        idl += "\t);\n"

            # interface properties
            if hasattr(clslistmember, '_public_attrs_'):
                for attr in clslistmember._public_attrs_:
                    idl += "\t[id(" + str(dispid) + "), propget, helpstring(\"property getter for " + attr + "\")]\n"
                    idl += "\tHRESULT " + attr + "(\n"
                    idl += "\t\t[out, retval] VARIANT *getval\n"
                    idl += "\t);\n"
                    if hasattr(clslistmember, '_readonly_attrs_'):
                        if attr not in clslistmember._readonly_attrs_:
                            idl += "\tid(" + str(dispid) + "), propput, helpstring(\"property setter for " + attr + "\")]\n"
                            idl += "\tHRESULT " + attr + "(\n"
                            idl += "\t\t[in] VARIANT setval\n"
                            idl += "\t);\n"
                    dispid += 1

            idl += "};\n\n"

        # library header
        idl += "[\n"
        idl += "\tuuid(" + cls._typelib_guid_[1:-1] + "),\n"
        idl += "\tversion(" + str(cls._typelib_version_[0]) + "." + str(cls._typelib_version_[1]) + "),\n"
        idl += "\thelpstring(\"" + cls._reg_typelib_filename_ + " - Python generated Type Library for COM registration\")\n"
        idl += "]\n"
        typelibname = getattr(cls, "_typelib_name_", "")
        if typelibname == "":
            typelibname = cls.__name__
        idl += "library " + typelibname + "\n"
        # library imports
        idl += "{\n"
        idl += "\t// TLib : OLE Automation : {00020430-0000-0000-C000-000000000046}\n"
        idl += "\timportlib(\"stdole32.tlb\");\n"
        idl += "\timportlib(\"stdole2.tlb\");\n\n"

        # library enums
        enumclslist = get_enum_classes(clsmodule)
        for clslistmemberid, clslistmember in enumclslist:
            idl += "\ttypedef enum {\n"
            firstkey = True
            for enumkey in clslistmember.__members__:
                if firstkey:
                    firstkey = False
                else:
                    idl += ",\n"
                if isinstance(clslistmember[enumkey].value, str):
                    idl += "\t\t" + enumkey + " = \"" + clslistmember[enumkey].value + "\""
                else:
                    idl += "\t\t" + enumkey + " = " + str(clslistmember[enumkey].value) + ""
            idl += "\n\t} " + clslistmember.__name__ + ";\n"
        if len(enumclslist) > 0:
            idl += "\n"

        for clslistmemberid, clslistmember in clslist:

            # library class
            interface = setinterface(clslistmember.__name__, clslistmember._com_interfaces_)   # determine interface name again because new loop
            idl += "\t[\n"
            idl += "\t\tuuid(" + clslistmember._reg_clsid_[1:-1] + "),\n"
            idl += "\t\thelpstring(\"coclass COM object for Python class " + clslistmember.__name__ + "\"),\n"
            idl += "\t]\n"
            idl += "\tcoclass " + clslistmember.__name__ + " : IDispatch\n"
            idl += "\t{\n"
            idl += "\t\t[default] interface " + interface + ";\n"
            idl += "\t}\n"

        idl += "}"

        if not okIDL:
            idlfile = idlfile.replace(".idl", "_to-be-checked.idl")
        with open(idlfile, "w+") as f:
            f.write(idl)

    else:
        okIDL = False

    return okIDL

# typelib compilation (i. e. call Microsoft IDL compiler midl.exe)
def compileTypeLib(cls):
    idl = get_filename(cls)
    tlb = os.path.splitext(idl)[0] + '.tlb'
    # tlb = pathlib.Path(idl).stem + '.tlb'
    if newer(idl, tlb):
        print("Compiling %s" % (idl,))
        rc = os.system('midl "%s"' % (idl,))
        if rc:
            err_msg = f"Compiling typelib for {cls.__name__} with MIDL failed!"
            raise RuntimeError(err_msg)
        # can't prevent MIDL from generating the stubs, just nuke them
        tlbdir = os.path.dirname(inspect.getfile(cls))
        # tlbdir = pathlib.Path(inspect.getfile(cls)).parent
        for helpfile in f"dlldata.c {cls.__name__ + '_i.c'} {cls.__name__ + '_p.c'} {cls.__name__ + '.h'}".split():
            os.remove(os.path.join(tlbdir, helpfile))
            # pathlib.Path(tlbdir).joinpath(helpfile).unlink(missing_ok=True)



# TypeLib registration / unregistration
# NOTE_: if _reg_typelib_filename_ is set, registration is done automatically as part of COM object
#        registration and unregistration (see Python code in win32com.server.register)

# register typelib
def registerTypeLib(cls):

    def registerTypeLibfile(tlbfile: str):
        tlbfile = os.path.abspath(tlbfile)
        # tlbfile = pathlib.Path(tlbfile).absolute()
        typelib = pythoncom.LoadTypeLib(tlbfile)
        pythoncom.RegisterTypeLib(typelib, tlbfile)

    unregister_typelib(cls)
    tlbfile = getattr(cls, "_reg_typelib_filename_", "")
    if tlbfile == "":
        tlbfile = get_filename(cls, ".tbl")
    if tlbfile != "":
        registerTypeLibfile(tlbfile)

# unregister typelib - copy from win32com.server.register
def unregister_typelib(cls):
    tlb_guid = getattr(cls, "_typelib_guid_")
    major, minor = getattr(cls, "_typelib_version_", (1, 0))
    lcid = getattr(cls, "_typelib_lcid_", 0)
    try:
        pythoncom.UnRegisterTypeLib(tlb_guid, major, minor, lcid)
        print('Unregistered type library for class {cls.__name__}.')
    except pythoncom.com_error:
        raise



# utilities for typelib generation

# get classes with same typelib assignment via _reg_typelib_filename_
def get_typelib_classes(cls, clsmodule: types.ModuleType) -> list[Any]:
    tlbfile = cls._reg_typelib_filename_
    if tlbfile == "":
        return [cls]
    else:
        return inspect.getmembers(clsmodule, lambda clsmember: getattr(clsmember, "_reg_typelib_filename_", "") == tlbfile and inspect.isclass)

# filename for typelib IDL belonging to class(es)
def get_filename(cls, ext: str = ".idl"):
    if cls._reg_typelib_filename_ == "" or cls._reg_typelib_filename_ == cls.__name__ + ".tlb" or cls._reg_typelib_filename_ == cls.__name__:
        return os.path.join(os.path.dirname(inspect.getfile(cls)), cls.__name__ + ext)
        # return pathlib.Path(inspect.getfile(cls)).parent.joinpath(cls.__name__ + ext)
    else:
        return os.path.join(os.path.dirname(inspect.getfile(cls)), os.path.splitext(cls._reg_typelib_filename_)[0] + ext)
        # return pathlib.Path(inspect.getfile(cls)).parent.joinpath(pathlib.Path(cls._reg_typelib_filename_).stem + ext)

# get enum classes in module
def get_enum_classes(clsmodule: types.ModuleType) -> list[Any]:
    return inspect.getmembers(clsmodule, lambda clsmember: getattr(getattr(clsmember, "__base__", ""), "__name__", "") == "Enum" and inspect.isclass)
