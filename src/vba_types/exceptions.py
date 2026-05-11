from typing import TypeVar


T = TypeVar('T', bound='VBARuntimeError')


class VBARuntimeError(Exception):
    """Base class that pulls metadata from class attributes."""
    code: int
    message: str

    def __init__(self: T) -> None:
        # Standard VBA error format: "Run-time error 'X': Message"
        super().__init__(f"Run-time error '{self.code}': {self.message}")


class ReturnWithoutGosubError(VBARuntimeError):
    code = 3
    message = "Return without GoSub"

class InvalidProcedureCallError(VBARuntimeError):
    code = 5
    message = "Invalid procedure call"

class OverflowError(VBARuntimeError):
    code = 6
    message = "Overflow"

class OutOfMemoryError(VBARuntimeError):
    code = 7
    message = "Out of memory"

class SubscriptOutOfRangeError(VBARuntimeError):
    code = 9
    message = "Subscript out of range"

class ArrayLockedError(VBARuntimeError):
    code = 10
    message = "This array is fixed or temporarily locked"

class DivisionByZeroError(VBARuntimeError):
    code = 11
    message = "Division by zero"

class TypeMismatchError(VBARuntimeError):
    code = 13
    message = "Type mismatch"

class OutOfStringSpaceError(VBARuntimeError):
    code = 14
    message = "Out of string space"

class ExpressionTooComplexError(VBARuntimeError):
    code = 16
    message = "Expression too complex"

class CantPerformOperationError(VBARuntimeError):
    code = 17
    message = "Can't perform requested operation"

class UserInterruptError(VBARuntimeError):
    code = 18
    message = "User interrupt occurred"

class ResumeWithoutError(VBARuntimeError):
    code = 20
    message = "Resume without error"

class OutOfStackSpaceError(VBARuntimeError):
    code = 28
    message = "Out of stack space"

class ProcedureNotDefinedError(VBARuntimeError):
    code = 35
    message = "Sub, Function, or Property not defined"

class ErrorLoadingDllError(VBARuntimeError):
    code = 48
    message = "Error in loading DLL"

class BadDllCallingConventionError(VBARuntimeError):
    code = 49
    message = "Bad DLL calling convention"

class InternalError(VBARuntimeError):
    code = 51
    message = "Internal error"

class BadFileNameOrNumberError(VBARuntimeError):
    code = 52
    message = "Bad file name or number"

class FileNotFoundError(VBARuntimeError):
    code = 53
    message = "File not found"

class BadFileModeError(VBARuntimeError):
    code = 54
    message = "Bad file mode"

class FileAlreadyOpenError(VBARuntimeError):
    code = 55
    message = "File already open"

class DeviceIOError(VBARuntimeError):
    code = 57
    message = "Device I/O error"

class FileAlreadyExistsError(VBARuntimeError):
    code = 58
    message = "File already exists"

class BadRecordLengthError(VBARuntimeError):
    code = 59
    message = "Bad record length"

class DiskFullError(VBARuntimeError):
    code = 61
    message = "Disk full"

class InputPastEndOfFileError(VBARuntimeError):
    code = 62
    message = "Input past end of file"

class BadRecordNumberError(VBARuntimeError):
    code = 63
    message = "Bad record number"

class TooManyFilesError(VBARuntimeError):
    code = 67
    message = "Too many files"

class DeviceUnavailableError(VBARuntimeError):
    code = 68
    message = "Device unavailable"

class PermissionDeniedError(VBARuntimeError):
    code = 70
    message = "Permission denied"

class DiskNotReadyError(VBARuntimeError):
    code = 71
    message = "Disk not ready"

class PathFileAccessError(VBARuntimeError):
    code = 75
    message = "Path/File access error"

class PathNotFoundError(VBARuntimeError):
    code = 76
    message = "Path not found"

class ObjectVariableNotSetError(VBARuntimeError):
    code = 91
    message = "Object variable or With block variable not set"

class ForLoopNotInitializedError(VBARuntimeError):
    code = 92
    message = "For loop not initialized"

class InvalidPatternStringError(VBARuntimeError):
    code = 93
    message = "Invalid pattern string"

class InvalidUseOfNullError(VBARuntimeError):
    code = 94
    message = "Invalid use of Null"

# --- Object and Automation Errors (Codes 424 - 1004) ---

class ObjectRequiredError(VBARuntimeError):
    code = 424
    message = "Object required"

class ActiveXComponentCantCreateObjectError(VBARuntimeError):
    code = 429
    message = "ActiveX component can't create object"

class ClassDoesNotSupportAutomationError(VBARuntimeError):
    code = 430
    message = "Class doesn't support Automation"

class ObjectDoesntSupportPropertyOrMethodError(VBARuntimeError):
    code = 438
    message = "Object doesn't support this property or method"

class AutomationError(VBARuntimeError):
    code = 440
    message = "Automation error"

class ObjectDoesNotSupportThisActionError(VBARuntimeError):
    code = 445
    message = "Object doesn't support this action"

class NamedArgumentNotFoundError(VBARuntimeError):
    code = 448
    message = "Named argument not found"

class ArgumentNotOptionalError(VBARuntimeError):
    code = 449
    message = "Argument not optional"

class ApplicationDefinedError(VBARuntimeError):
    code = 1004
    message = "Application-defined or object-defined error"
