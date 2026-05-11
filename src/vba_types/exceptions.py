class VBARuntimeError(Exception):
    """Base class for all trappable VBA runtime errors."""
    def __init__(self, code, message):
        self.code = code
        self.message = message
        super().__init__(f"Run-time error '{self.code}': {self.message}")


class ReturnWithoutGosubError(VBARuntimeError):
    def __init__(self): super().__init__(3, "Return without GoSub")


class InvalidProcedureCallError(VBARuntimeError):
    def __init__(self): super().__init__(5, "Invalid procedure call")


class OverflowError(VBARuntimeError):
    def __init__(self): super().__init__(6, "Overflow")


class OutOfMemoryError(VBARuntimeError):
    def __init__(self): super().__init__(7, "Out of memory")


class SubscriptOutOfRangeError(VBARuntimeError):
    def __init__(self): super().__init__(9, "Subscript out of range")


class ArrayLockedError(VBARuntimeError):
    def __init__(self): super().__init__(10, "This array is fixed or temporarily locked")


class DivisionByZeroError(VBARuntimeError):
    def __init__(self): super().__init__(11, "Division by zero")


class TypeMismatchError(VBARuntimeError):
    def __init__(self): super().__init__(13, "Type mismatch")


class OutOfStringSpaceError(VBARuntimeError):
    def __init__(self): super().__init__(14, "Out of string space")


class ExpressionTooComplexError(VBARuntimeError):
    def __init__(self): super().__init__(16, "Expression too complex")


class CantPerformOperationError(VBARuntimeError):
    def __init__(self): super().__init__(17, "Can't perform requested operation")


class UserInterruptError(VBARuntimeError):
    def __init__(self): super().__init__(18, "User interrupt occurred")


class ResumeWithoutError(VBARuntimeError):
    def __init__(self): super().__init__(20, "Resume without error")


class OutOfStackSpaceError(VBARuntimeError):
    def __init__(self): super().__init__(28, "Out of stack space")


class ProcedureNotDefinedError(VBARuntimeError):
    def __init__(self): super().__init__(35, "Sub, Function, or Property not defined")


class ErrorLoadingDllError(VBARuntimeError):
    def __init__(self): super().__init__(48, "Error in loading DLL")


class BadDllCallingConventionError(VBARuntimeError):
    def __init__(self): super().__init__(49, "Bad DLL calling convention")


class InternalError(VBARuntimeError):
    def __init__(self): super().__init__(51, "Internal error")


class BadFileNameOrNumberError(VBARuntimeError):
    def __init__(self): super().__init__(52, "Bad file name or number")


class FileNotFoundError(VBARuntimeError):
    def __init__(self): super().__init__(53, "File not found")


class BadFileModeError(VBARuntimeError):
    def __init__(self): super().__init__(54, "Bad file mode")


class FileAlreadyOpenError(VBARuntimeError):
    def __init__(self): super().__init__(55, "File already open")


class DeviceIOError(VBARuntimeError):
    def __init__(self): super().__init__(57, "Device I/O error")


class FileAlreadyExistsError(VBARuntimeError):
    def __init__(self): super().__init__(58, "File already exists")


class BadRecordLengthError(VBARuntimeError):
    def __init__(self): super().__init__(59, "Bad record length")


class DiskFullError(VBARuntimeError):
    def __init__(self): super().__init__(61, "Disk full")


class InputPastEndOfFileError(VBARuntimeError):
    def __init__(self): super().__init__(62, "Input past end of file")


class BadRecordNumberError(VBARuntimeError):
    def __init__(self): super().__init__(63, "Bad record number")


class TooManyFilesError(VBARuntimeError):
    def __init__(self): super().__init__(67, "Too many files")


class DeviceUnavailableError(VBARuntimeError):
    def __init__(self): super().__init__(68, "Device unavailable")


class PermissionDeniedError(VBARuntimeError):
    def __init__(self): super().__init__(70, "Permission denied")


class DiskNotReadyError(VBARuntimeError):
    def __init__(self): super().__init__(71, "Disk not ready")


class PathFileAccessError(VBARuntimeError):
    def __init__(self): super().__init__(75, "Path/File access error")


class PathNotFoundError(VBARuntimeError):
    def __init__(self): super().__init__(76, "Path not found")


class ObjectVariableNotSetError(VBARuntimeError):
    def __init__(self): super().__init__(91, "Object variable or With block variable not set")


class ForLoopNotInitializedError(VBARuntimeError):
    def __init__(self): super().__init__(92, "For loop not initialized")


class InvalidPatternStringError(VBARuntimeError):
    def __init__(self): super().__init__(93, "Invalid pattern string")


class InvalidUseOfNullError(VBARuntimeError):
    def __init__(self): super().__init__(94, "Invalid use of Null")


class ActiveXComponentCantCreateObjectError(VBARuntimeError):
    def __init__(self): super().__init__(429, "ActiveX component can't create object")


class ClassDoesNotSupportAutomationError(VBARuntimeError):
    def __init__(self): super().__init__(430, "Class doesn't support Automation")


class ObjectDoesntSupportPropertyOrMethodError(VBARuntimeError):
    def __init__(self): super().__init__(438, "Object doesn't support this property or method")


class AutomationError(VBARuntimeError):
    def __init__(self): super().__init__(440, "Automation error")


class ObjectDoesNotSupportThisActionError(VBARuntimeError):
    def __init__(self): super().__init__(445, "Object doesn't support this action")


class NamedArgumentNotFoundError(VBARuntimeError):
    def __init__(self): super().__init__(448, "Named argument not found")


class ArgumentNotOptionalError(VBARuntimeError):
    def __init__(self): super().__init__(449, "Argument not optional")


class ApplicationDefinedError(VBARuntimeError):
    def __init__(self): super().__init__(1004, "Application-defined or object-defined error")
