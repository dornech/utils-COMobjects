# COM utilities for Python

## Short description

Package with various utilities to support COM object development with Python based on PyWin32

## Package content

The package  contains:

- abstract base classes for COM objects
- consistency checker and registration tool for a Python COM object
- decorators for COM object method calls
- mixin class for logging in COM objects
- typelib generator (experimental)

For example see demo script:
```{python}
# COM object utilities
# test and demo program



# from traceback_with_variables import activate_by_import

from typing import *
from enum import Enum

import utils_COMobjects as UtilsCOMobjects
# import UtilsCOMobjects



class PyCOMtest_baseClass1(UtilsCOMobjects.COMclass.baseCOMclass):

    _reg_clsid_ = "{C284AD0B-E2C3-45D2-8D55-A4327FB4AD5C}"
    _reg_progid_ = "PyCOMtest.baseCOMclass1"
    _public_methods_ = ["dummy", "dummy_not_existing"]
    _public_attrs_ = [
        "publicattr_class",
        "publicattr_class_readonly",
        "publicattr_instance",
        "publicattr_instance_readonly",
        "publicattr_class_error",
    ]
    _readonly_attrs_ = [
        "publicattr_class_readonly",
        "publicattr_instance_readonly",
        "publicattr_readonly_error",
    ]

    # public attributes for testing
    publicattr_class: int = 10
    publicattr_class_readonly: int = 11

    # init to set instance _readonly_attrs_
    def __init__(self):
        self.publicattr_instance: int = 20
        self.publicattr_instance_readonly: int = 21
        pass

    # dummy method
    def dummy(self) -> any:
        return "test"

    # dummy method not registered
    def dummy2(self) -> any:
        return "test"

    # dummy method marked as private (i.e. not callable)
    def _dummy(self) -> any:
        return "test"


class PyCOMtest_typelibDummy(UtilsCOMobjects.COMclass.typelibCOMclass):

    _reg_clsid_ = "{5E6C5628-170A-4E6E-B16A-406773B72433}"
    _reg_progid_ = "PyCOMtest.typelibDummy"
    _public_methods_ = ["dummy", "teststatic"]
    _public_attrs_ = []
    _readonly_attrs_ = []

    # dummy method
    def checkDebug(self) -> any:
        return self._checkDebug

    # dummy method
    def dummy(self) -> any:
        return "test"

    # test static
    @staticmethod
    def teststatic() -> str:
        return "statictest"

class PyCOMtest_typelibSuccess(UtilsCOMobjects.COMclass.typelibCOMclass):

    _reg_clsid_ = "{76140A16-1E6A-46CC-B5D7-9934FF7BB836}"
    _reg_progid_ = "PyCOMtest.typelibSuccess"
    _public_methods_ = ["dummy", "teststatic", "teststatic2", "test"]
    _public_attrs_ = ["dummyattr", "dummyattrstr"]
    _readonly_attrs_ = ["dummyattrstr"]

    _reg_typelib_filename_ = "PyCOMtest_typelibSuccess.tlb"
    _typelib_name_ = "typelibSuccess"
    _typelib_guid_ = "{F699E296-3997-4742-99E1-47BE4CFE3C83}"
    _com_interfaces_ = ["ItypelibClassSuccess"]
    _typelib_interfaceID_ = "{86D65A77-3B9D-4BAB-9173-ABE87320987C}"

    dummyattr = "test"
    dummyattrstr = "test"

    # dummy method
    def dummy(self):
        return "test"

    # test static
    @staticmethod
    def teststatic(inint: int) -> str:
        return "statictest"

    # test static2
    @staticmethod
    def teststatic2() -> Any:
        return "statictest2"

    # test method
    def test(self, inint: int, inout: int, instr: str, intypeless) -> Any:
        return inout, "test"


class PyCOMtest_typelibClass1(UtilsCOMobjects.COMclass.typelibCOMclass):

    _reg_clsid_ = "{46120ADB-B37E-48DC-B14B-D41D465F8EEA}"
    _reg_progid_ = "PyCOMtest.typelibClass1"
    _public_methods_ = ["dummy", "test"]
    _public_attrs_ = []
    _readonly_attrs_ = []

    _reg_typelib_filename_ = "testUtilsCOMobjects.tlb"
    _typelib_name_ = "testUtilsCOMobjects"
    _typelib_guid_ = "{0EA11420-6434-421D-ACA6-FD03AD9DFD94}"
    _com_interfaces_ = ["ItypelibClass1"]
    _typelib_interfaceID_ = "{CFB8F4E0-FE79-4CC8-B3BA-A3DE56E13EC5}"

    # dummy method
    def dummy(self) -> any:
        return "test"

    # test method
    def test(self, inint: int, inout: int, instr: str, intypeless) -> Any:
        return inout, "test"


class PyCOMtest_typelibClass2(UtilsCOMobjects.COMclass.typelibCOMclass):

    _reg_clsid_ = "{30D6B85B-5E4B-41B2-A984-7A365FE5E5D5}"
    _reg_progid_ = "PyCOMtest.typelibClass2"
    _public_methods_ = ["dummy"]
    _public_attrs_ = []
    _readonly_attrs_ = []

    _reg_typelib_filename_ = "testUtilsCOMobjects.tlb"
    _typelib_name_ = "testUtilsCOMobjects"
    _typelib_guid_ = "{F17C46CB-549F-4BF6-8460-68A926F622C8}"   # other GUID
    _com_interfaces_ = ["ItypelibClass2"]
    _typelib_interfaceID_ = "{D2FECFAD-321E-413D-A210-027738446639}"

    # dummy method
    def dummy(self) -> str:
        return "test"


class PyCOMtest_typelibClass3(UtilsCOMobjects.COMclass.typelibCOMclass):

    _reg_clsid_ = "{9965F0F5-1509-4EB4-971D-2E75DCACBA4F}"
    _reg_progid_ = "PyCOMtest.typelibClass3"
    _public_methods_ = ["dummy"]
    _public_attrs_ = []
    _readonly_attrs_ = []

    _reg_typelib_filename_ = "testUtilsCOMobjects_OTHER.tlb"
    _typelib_name_ = "testUtilsCOMobjects"
    _typelib_guid_ = "{0EA11420-6434-421D-ACA6-FD03AD9DFD94}"
    _com_interfaces_ = ["ItypelibClass3a", "ItypelibClass3b"]
    _typelib_interfaceID_ = "{32C00ACF-343B-4C46-A84A-A160CDDCFAFA}"


    # dummy method
    def dummy(self) -> str:
        return "test"


# required for testing
class DummyClass():
    pass


# enum class
class testenum(Enum):
    ENUMVAL1 = 1
    ENUMVAL2 = 2
    ENUMVAL3 = 3



if __name__ == "__main__":

    test = PyCOMtest_baseClass1()

    UtilsCOMobjects.UtilsCOMcheckreg.checkAttribsTypeLib(PyCOMtest_typelibDummy)
    print()
    UtilsCOMobjects.UtilsCOMcheckreg.checkAttribsTypeLib(PyCOMtest_typelibSuccess)
    print()
    UtilsCOMobjects.UtilsCOMcheckreg.checkAttribsTypeLib(PyCOMtest_typelibClass1)
    print()
    UtilsCOMobjects.UtilsCOMcheckreg.checkAttribsTypeLib(PyCOMtest_typelibClass2)
    print()
    UtilsCOMobjects.UtilsCOMcheckreg.checkAttribsTypeLib(PyCOMtest_typelibClass3)
    print()

    UtilsCOMobjects.UtilsCOMcheckreg.processCOMregistration(test, gentypelib=True, testmode=True)
    UtilsCOMobjects.UtilsCOMcheckreg.processCOMregistration(PyCOMtest_baseClass1, gentypelib=True, testmode=True)
    UtilsCOMobjects.UtilsCOMcheckreg.processCOMregistration(PyCOMtest_typelibDummy, gentypelib=True, testmode=True)
    UtilsCOMobjects.UtilsCOMcheckreg.processCOMregistration(PyCOMtest_typelibSuccess, gentypelib=True, testmode=True)
    UtilsCOMobjects.UtilsCOMcheckreg.processCOMregistration(PyCOMtest_typelibClass1, gentypelib=True, testmode=True)
    UtilsCOMobjects.UtilsCOMcheckreg.processCOMregistration(PyCOMtest_typelibClass2, gentypelib=True, testmode=True)
    UtilsCOMobjects.UtilsCOMcheckreg.processCOMregistration(PyCOMtest_typelibClass3, gentypelib=True, testmode=True)
```

## Navigation

Documentation for specific `MAJOR.MINOR` versions can be chosen by using the dropdown on the top of every page.
The `dev` version reflects changes that have not yet been released. Shortcuts can be used for navigation, i.e.
<kbd>,</kbd>/<kbd>p</kbd> and <kbd>.</kbd>/<kbd>n</kbd> for previous and next page, respectively, as well as
<kbd>/</kbd>/<kbd>s</kbd> for searching.
