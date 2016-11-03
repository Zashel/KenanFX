# -*- coding: mbcs -*-
typelib_path = 'msado25.tlb'
_lcid = 0 # change this if required
from ctypes import *
import comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0
from comtypes import GUID
from ctypes import HRESULT
from comtypes import BSTR
from comtypes import dispid
from comtypes import DISPMETHOD, DISPPROPERTY, helpstring
from comtypes import helpstring
from comtypes import COMMETHOD
from comtypes import IUnknown
from comtypes.automation import VARIANT
from comtypes.automation import IDispatch
from ctypes.wintypes import VARIANT_BOOL
from comtypes import CoClass



# values for enumeration 'SearchDirectionEnum'
adSearchForward = 1
adSearchBackward = -1
SearchDirectionEnum = c_int # enum

# values for enumeration 'LockTypeEnum'
adLockUnspecified = -1
adLockReadOnly = 1
adLockPessimistic = 2
adLockOptimistic = 3
adLockBatchOptimistic = 4
LockTypeEnum = c_int # enum

# values for enumeration 'AffectEnum'
adAffectCurrent = 1
adAffectGroup = 2
adAffectAll = 3
adAffectAllChapters = 4
AffectEnum = c_int # enum

# values for enumeration 'StringFormatEnum'
adClipString = 2
StringFormatEnum = c_int # enum

# values for enumeration 'ConnectOptionEnum'
adConnectUnspecified = -1
adAsyncConnect = 16
ConnectOptionEnum = c_int # enum

# values for enumeration 'GetRowsOptionEnum'
adGetRowsRest = -1
GetRowsOptionEnum = c_int # enum

# values for enumeration 'ParameterAttributesEnum'
adParamSigned = 16
adParamNullable = 64
adParamLong = 128
ParameterAttributesEnum = c_int # enum

# values for enumeration 'BookmarkEnum'
adBookmarkCurrent = 0
adBookmarkFirst = 1
adBookmarkLast = 2
BookmarkEnum = c_int # enum

# values for enumeration 'MarshalOptionsEnum'
adMarshalAll = 0
adMarshalModifiedOnly = 1
MarshalOptionsEnum = c_int # enum

# values for enumeration 'ObjectStateEnum'
adStateClosed = 0
adStateOpen = 1
adStateConnecting = 2
adStateExecuting = 4
adStateFetching = 8
ObjectStateEnum = c_int # enum

# values for enumeration 'DataTypeEnum'
adEmpty = 0
adTinyInt = 16
adSmallInt = 2
adInteger = 3
adBigInt = 20
adUnsignedTinyInt = 17
adUnsignedSmallInt = 18
adUnsignedInt = 19
adUnsignedBigInt = 21
adSingle = 4
adDouble = 5
adCurrency = 6
adDecimal = 14
adNumeric = 131
adBoolean = 11
adError = 10
adUserDefined = 132
adVariant = 12
adIDispatch = 9
adIUnknown = 13
adGUID = 72
adDate = 7
adDBDate = 133
adDBTime = 134
adDBTimeStamp = 135
adBSTR = 8
adChar = 129
adVarChar = 200
adLongVarChar = 201
adWChar = 130
adVarWChar = 202
adLongVarWChar = 203
adBinary = 128
adVarBinary = 204
adLongVarBinary = 205
adChapter = 136
adFileTime = 64
adPropVariant = 138
adVarNumeric = 139
adArray = 8192
DataTypeEnum = c_int # enum

# values for enumeration 'ResyncEnum'
adResyncUnderlyingValues = 1
adResyncAllValues = 2
ResyncEnum = c_int # enum

# values for enumeration 'CursorLocationEnum'
adUseNone = 1
adUseServer = 2
adUseClient = 3
adUseClientBatch = 3
CursorLocationEnum = c_int # enum
class ConnectionEvents(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000400-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = []
    _methods_ = []
class Error(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000500-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']

# values for enumeration 'EventStatusEnum'
adStatusOK = 1
adStatusErrorsOccurred = 2
adStatusCantDeny = 3
adStatusCancel = 4
adStatusUnwantedEvent = 5
EventStatusEnum = c_int # enum
class _ADO(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000534-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class Connection15(_ADO):
    _case_insensitive_ = True
    _iid_ = GUID('{00000515-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'oleautomation']
class _Connection(Connection15):
    _case_insensitive_ = True
    _iid_ = GUID('{00000550-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'oleautomation']

# values for enumeration 'CursorTypeEnum'
adOpenUnspecified = -1
adOpenForwardOnly = 0
adOpenKeyset = 1
adOpenDynamic = 2
adOpenStatic = 3
CursorTypeEnum = c_int # enum
class Command15(_ADO):
    _case_insensitive_ = True
    _iid_ = GUID('{00000508-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'nonextensible', 'oleautomation']
class _Command(Command15):
    _case_insensitive_ = True
    _iid_ = GUID('{0000054E-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class Recordset15(_ADO):
    _case_insensitive_ = True
    _iid_ = GUID('{0000050E-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'nonextensible', 'oleautomation']
class Recordset20(Recordset15):
    _case_insensitive_ = True
    _iid_ = GUID('{0000054F-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'nonextensible', 'oleautomation']
class Recordset21(Recordset20):
    _case_insensitive_ = True
    _iid_ = GUID('{00000555-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'nonextensible', 'oleautomation']
class _Recordset(Recordset21):
    _case_insensitive_ = True
    _iid_ = GUID('{00000556-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ConnectionEvents._disp_methods_ = [
    DISPMETHOD([dispid(0)], HRESULT, 'InfoMessage',
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Connection), 'pConnection' )),
    DISPMETHOD([dispid(1)], HRESULT, 'BeginTransComplete',
               ( ['in'], c_int, 'TransactionLevel' ),
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Connection), 'pConnection' )),
    DISPMETHOD([dispid(3)], HRESULT, 'CommitTransComplete',
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Connection), 'pConnection' )),
    DISPMETHOD([dispid(2)], HRESULT, 'RollbackTransComplete',
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Connection), 'pConnection' )),
    DISPMETHOD([dispid(4)], HRESULT, 'WillExecute',
               ( ['in', 'out'], POINTER(BSTR), 'Source' ),
               ( ['in', 'out'], POINTER(CursorTypeEnum), 'CursorType' ),
               ( ['in', 'out'], POINTER(LockTypeEnum), 'LockType' ),
               ( ['in', 'out'], POINTER(c_int), 'Options' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Command), 'pCommand' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' ),
               ( ['in'], POINTER(_Connection), 'pConnection' )),
    DISPMETHOD([dispid(5)], HRESULT, 'ExecuteComplete',
               ( ['in'], c_int, 'RecordsAffected' ),
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Command), 'pCommand' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' ),
               ( ['in'], POINTER(_Connection), 'pConnection' )),
    DISPMETHOD([dispid(6)], HRESULT, 'WillConnect',
               ( ['in', 'out'], POINTER(BSTR), 'ConnectionString' ),
               ( ['in', 'out'], POINTER(BSTR), 'UserID' ),
               ( ['in', 'out'], POINTER(BSTR), 'Password' ),
               ( ['in', 'out'], POINTER(c_int), 'Options' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Connection), 'pConnection' )),
    DISPMETHOD([dispid(7)], HRESULT, 'ConnectComplete',
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Connection), 'pConnection' )),
    DISPMETHOD([dispid(8)], HRESULT, 'Disconnect',
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Connection), 'pConnection' )),
]

# values for enumeration 'FieldAttributeEnum'
adFldUnspecified = -1
adFldMayDefer = 2
adFldUpdatable = 4
adFldUnknownUpdatable = 8
adFldFixed = 16
adFldIsNullable = 32
adFldMayBeNull = 64
adFldLong = 128
adFldRowID = 256
adFldRowVersion = 512
adFldCacheDeferred = 4096
adFldIsChapter = 8192
adFldNegativeScale = 16384
adFldKeyColumn = 32768
adFldIsRowURL = 65536
adFldIsDefaultStream = 131072
adFldIsCollection = 262144
FieldAttributeEnum = c_int # enum
class IRecFields(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{00000563-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'oleautomation']
IRecFields._methods_ = [
    COMMETHOD([dispid(1)], HRESULT, 'ADOCheck'),
]
################################################################
## code template for IRecFields implementation
##class IRecFields_Impl(object):
##    def ADOCheck(self):
##        '-no docstring-'
##        #return 
##


# values for enumeration 'PersistFormatEnum'
adPersistADTG = 0
adPersistXML = 1
PersistFormatEnum = c_int # enum
class _Collection(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000512-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class Fields15(_Collection):
    _case_insensitive_ = True
    _iid_ = GUID('{00000506-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'nonextensible', 'oleautomation']
class Fields20(Fields15):
    _case_insensitive_ = True
    _iid_ = GUID('{0000054D-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'nonextensible', 'oleautomation']
_Collection._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Count',
              ( ['retval', 'out'], POINTER(c_int), 'c' )),
    COMMETHOD([dispid(-4), 'restricted'], HRESULT, '_NewEnum',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppvObject' )),
    COMMETHOD([dispid(2)], HRESULT, 'Refresh'),
]
################################################################
## code template for _Collection implementation
##class _Collection_Impl(object):
##    def _NewEnum(self):
##        '-no docstring-'
##        #return ppvObject
##
##    @property
##    def Count(self):
##        '-no docstring-'
##        #return c
##
##    def Refresh(self):
##        '-no docstring-'
##        #return 
##

class Field20(_ADO):
    _case_insensitive_ = True
    _iid_ = GUID('{0000054C-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'nonextensible', 'oleautomation']
class Field(Field20):
    _case_insensitive_ = True
    _iid_ = GUID('{00000569-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
Fields15._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Item',
              ( ['in'], VARIANT, 'Index' ),
              ( ['retval', 'out'], POINTER(POINTER(Field)), 'ppvObject' )),
]
################################################################
## code template for Fields15 implementation
##class Fields15_Impl(object):
##    @property
##    def Item(self, Index):
##        '-no docstring-'
##        #return ppvObject
##

Fields20._methods_ = [
    COMMETHOD([dispid(1610874880), 'hidden'], HRESULT, '_Append',
              ( ['in'], BSTR, 'Name' ),
              ( ['in'], DataTypeEnum, 'Type' ),
              ( ['in', 'optional'], c_int, 'DefinedSize', 0 ),
              ( ['in', 'optional'], FieldAttributeEnum, 'Attrib', -1 )),
    COMMETHOD([dispid(4)], HRESULT, 'Delete',
              ( ['in'], VARIANT, 'Index' )),
]
################################################################
## code template for Fields20 implementation
##class Fields20_Impl(object):
##    def _Append(self, Name, Type, DefinedSize, Attrib):
##        '-no docstring-'
##        #return 
##
##    def Delete(self, Index):
##        '-no docstring-'
##        #return 
##


# values for enumeration 'FilterGroupEnum'
adFilterNone = 0
adFilterPendingRecords = 1
adFilterAffectedRecords = 2
adFilterFetchedRecords = 3
adFilterPredicate = 4
adFilterConflictingRecords = 5
FilterGroupEnum = c_int # enum
class _Record(_ADO):
    _case_insensitive_ = True
    _iid_ = GUID('{00000562-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'oleautomation']
class Properties(_Collection):
    _case_insensitive_ = True
    _iid_ = GUID('{00000504-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
_ADO._methods_ = [
    COMMETHOD([dispid(500), 'propget'], HRESULT, 'Properties',
              ( ['retval', 'out'], POINTER(POINTER(Properties)), 'ppvObject' )),
]
################################################################
## code template for _ADO implementation
##class _ADO_Impl(object):
##    @property
##    def Properties(self):
##        '-no docstring-'
##        #return ppvObject
##


# values for enumeration 'ConnectModeEnum'
adModeUnknown = 0
adModeRead = 1
adModeWrite = 2
adModeReadWrite = 3
adModeShareDenyRead = 4
adModeShareDenyWrite = 8
adModeShareExclusive = 12
adModeShareDenyNone = 16
adModeRecursive = 4194304
ConnectModeEnum = c_int # enum

# values for enumeration 'MoveRecordOptionsEnum'
adMoveUnspecified = -1
adMoveOverWrite = 1
adMoveDontUpdateLinks = 2
adMoveAllowEmulation = 4
MoveRecordOptionsEnum = c_int # enum

# values for enumeration 'CopyRecordOptionsEnum'
adCopyUnspecified = -1
adCopyOverWrite = 1
adCopyAllowEmulation = 4
adCopyNonRecursive = 2
CopyRecordOptionsEnum = c_int # enum

# values for enumeration 'RecordCreateOptionsEnum'
adCreateCollection = 8192
adCreateStructDoc = -2147483648
adCreateNonCollection = 0
adOpenIfExists = 33554432
adCreateOverwrite = 67108864
adFailIfNotExists = -1
RecordCreateOptionsEnum = c_int # enum

# values for enumeration 'RecordOpenOptionsEnum'
adOpenRecordUnspecified = -1
adOpenSource = 8388608
adOpenAsync = 4096
adDelayFetchStream = 16384
adDelayFetchFields = 32768
RecordOpenOptionsEnum = c_int # enum
class Fields(Fields20):
    _case_insensitive_ = True
    _iid_ = GUID('{00000564-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']

# values for enumeration 'RecordTypeEnum'
adSimpleRecord = 0
adCollectionRecord = 1
adStructDoc = 2
RecordTypeEnum = c_int # enum
_Record._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ActiveConnection',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(1), 'propput'], HRESULT, 'ActiveConnection',
              ( ['in'], BSTR, 'pvar' )),
    COMMETHOD([dispid(1), 'propputref'], HRESULT, 'ActiveConnection',
              ( ['in'], POINTER(_Connection), 'pvar' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'State',
              ( ['retval', 'out'], POINTER(ObjectStateEnum), 'pState' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'Source',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(3), 'propput'], HRESULT, 'Source',
              ( ['in'], BSTR, 'pvar' )),
    COMMETHOD([dispid(3), 'propputref'], HRESULT, 'Source',
              ( ['in'], POINTER(IDispatch), 'pvar' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'Mode',
              ( ['retval', 'out'], POINTER(ConnectModeEnum), 'pMode' )),
    COMMETHOD([dispid(4), 'propput'], HRESULT, 'Mode',
              ( ['in'], ConnectModeEnum, 'pMode' )),
    COMMETHOD([dispid(5), 'propget'], HRESULT, 'ParentURL',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstrParentURL' )),
    COMMETHOD([dispid(6)], HRESULT, 'MoveRecord',
              ( ['in', 'optional'], BSTR, 'Source', '' ),
              ( ['in', 'optional'], BSTR, 'Destination', '' ),
              ( ['in', 'optional'], BSTR, 'UserName', '' ),
              ( ['in', 'optional'], BSTR, 'Password', '' ),
              ( ['in', 'optional'], MoveRecordOptionsEnum, 'Options', -1 ),
              ( ['in', 'optional'], VARIANT_BOOL, 'Async', False ),
              ( ['retval', 'out'], POINTER(BSTR), 'pbstrNewURL' )),
    COMMETHOD([dispid(7)], HRESULT, 'CopyRecord',
              ( ['in', 'optional'], BSTR, 'Source', '' ),
              ( ['in', 'optional'], BSTR, 'Destination', '' ),
              ( ['in', 'optional'], BSTR, 'UserName', '' ),
              ( ['in', 'optional'], BSTR, 'Password', '' ),
              ( ['in', 'optional'], CopyRecordOptionsEnum, 'Options', -1 ),
              ( ['in', 'optional'], VARIANT_BOOL, 'Async', False ),
              ( ['retval', 'out'], POINTER(BSTR), 'pbstrNewURL' )),
    COMMETHOD([dispid(8)], HRESULT, 'DeleteRecord',
              ( ['in', 'optional'], BSTR, 'Source', '' ),
              ( ['in', 'optional'], VARIANT_BOOL, 'Async', False )),
    COMMETHOD([dispid(9)], HRESULT, 'Open',
              ( ['in', 'optional'], VARIANT, 'Source' ),
              ( ['in', 'optional'], VARIANT, 'ActiveConnection' ),
              ( ['in', 'optional'], ConnectModeEnum, 'Mode', 0 ),
              ( ['in', 'optional'], RecordCreateOptionsEnum, 'CreateOptions', -1 ),
              ( ['in', 'optional'], RecordOpenOptionsEnum, 'Options', -1 ),
              ( ['in', 'optional'], BSTR, 'UserName', '' ),
              ( ['in', 'optional'], BSTR, 'Password', '' )),
    COMMETHOD([dispid(10)], HRESULT, 'Close'),
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Fields',
              ( ['retval', 'out'], POINTER(POINTER(Fields)), 'ppFlds' )),
    COMMETHOD([dispid(11), 'propget'], HRESULT, 'RecordType',
              ( ['retval', 'out'], POINTER(RecordTypeEnum), 'ptype' )),
    COMMETHOD([dispid(12)], HRESULT, 'GetChildren',
              ( ['retval', 'out'], POINTER(POINTER(_Recordset)), 'pprset' )),
    COMMETHOD([dispid(13)], HRESULT, 'Cancel'),
]
################################################################
## code template for _Record implementation
##class _Record_Impl(object):
##    def Close(self):
##        '-no docstring-'
##        #return 
##
##    @property
##    def State(self):
##        '-no docstring-'
##        #return pState
##
##    def CopyRecord(self, Source, Destination, UserName, Password, Options, Async):
##        '-no docstring-'
##        #return pbstrNewURL
##
##    @property
##    def RecordType(self):
##        '-no docstring-'
##        #return ptype
##
##    def MoveRecord(self, Source, Destination, UserName, Password, Options, Async):
##        '-no docstring-'
##        #return pbstrNewURL
##
##    def Open(self, Source, ActiveConnection, Mode, CreateOptions, Options, UserName, Password):
##        '-no docstring-'
##        #return 
##
##    def Cancel(self):
##        '-no docstring-'
##        #return 
##
##    @property
##    def ParentURL(self):
##        '-no docstring-'
##        #return pbstrParentURL
##
##    def _get(self):
##        '-no docstring-'
##        #return pMode
##    def _set(self, pMode):
##        '-no docstring-'
##    Mode = property(_get, _set, doc = _set.__doc__)
##
##    def ActiveConnection(self, pvar):
##        '-no docstring-'
##        #return 
##
##    def Source(self, pvar):
##        '-no docstring-'
##        #return 
##
##    def DeleteRecord(self, Source, Async):
##        '-no docstring-'
##        #return 
##
##    def GetChildren(self):
##        '-no docstring-'
##        #return pprset
##
##    @property
##    def Fields(self):
##        '-no docstring-'
##        #return ppFlds
##


# values for enumeration 'ADCPROP_UPDATECRITERIA_ENUM'
adCriteriaKey = 0
adCriteriaAllCols = 1
adCriteriaUpdCols = 2
adCriteriaTimeStamp = 3
ADCPROP_UPDATECRITERIA_ENUM = c_int # enum
class ConnectionEventsVt(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{00000402-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden']
ConnectionEventsVt._methods_ = [
    COMMETHOD([], HRESULT, 'InfoMessage',
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Connection), 'pConnection' )),
    COMMETHOD([], HRESULT, 'BeginTransComplete',
              ( ['in'], c_int, 'TransactionLevel' ),
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Connection), 'pConnection' )),
    COMMETHOD([], HRESULT, 'CommitTransComplete',
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Connection), 'pConnection' )),
    COMMETHOD([], HRESULT, 'RollbackTransComplete',
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Connection), 'pConnection' )),
    COMMETHOD([], HRESULT, 'WillExecute',
              ( ['in', 'out'], POINTER(BSTR), 'Source' ),
              ( ['in', 'out'], POINTER(CursorTypeEnum), 'CursorType' ),
              ( ['in', 'out'], POINTER(LockTypeEnum), 'LockType' ),
              ( ['in', 'out'], POINTER(c_int), 'Options' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Command), 'pCommand' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' ),
              ( ['in'], POINTER(_Connection), 'pConnection' )),
    COMMETHOD([], HRESULT, 'ExecuteComplete',
              ( ['in'], c_int, 'RecordsAffected' ),
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Command), 'pCommand' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' ),
              ( ['in'], POINTER(_Connection), 'pConnection' )),
    COMMETHOD([], HRESULT, 'WillConnect',
              ( ['in', 'out'], POINTER(BSTR), 'ConnectionString' ),
              ( ['in', 'out'], POINTER(BSTR), 'UserID' ),
              ( ['in', 'out'], POINTER(BSTR), 'Password' ),
              ( ['in', 'out'], POINTER(c_int), 'Options' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Connection), 'pConnection' )),
    COMMETHOD([], HRESULT, 'ConnectComplete',
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Connection), 'pConnection' )),
    COMMETHOD([], HRESULT, 'Disconnect',
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Connection), 'pConnection' )),
]
################################################################
## code template for ConnectionEventsVt implementation
##class ConnectionEventsVt_Impl(object):
##    def CommitTransComplete(self, pError, pConnection):
##        '-no docstring-'
##        #return adStatus
##
##    def RollbackTransComplete(self, pError, pConnection):
##        '-no docstring-'
##        #return adStatus
##
##    def Disconnect(self, pConnection):
##        '-no docstring-'
##        #return adStatus
##
##    def WillExecute(self, pCommand, pRecordset, pConnection):
##        '-no docstring-'
##        #return Source, CursorType, LockType, Options, adStatus
##
##    def ExecuteComplete(self, RecordsAffected, pError, pCommand, pRecordset, pConnection):
##        '-no docstring-'
##        #return adStatus
##
##    def ConnectComplete(self, pError, pConnection):
##        '-no docstring-'
##        #return adStatus
##
##    def InfoMessage(self, pError, pConnection):
##        '-no docstring-'
##        #return adStatus
##
##    def BeginTransComplete(self, TransactionLevel, pError, pConnection):
##        '-no docstring-'
##        #return adStatus
##
##    def WillConnect(self, pConnection):
##        '-no docstring-'
##        #return ConnectionString, UserID, Password, Options, adStatus
##


# values for enumeration 'ErrorValueEnum'
adErrProviderFailed = 3000
adErrInvalidArgument = 3001
adErrOpeningFile = 3002
adErrReadFile = 3003
adErrWriteFile = 3004
adErrNoCurrentRecord = 3021
adErrIllegalOperation = 3219
adErrCantChangeProvider = 3220
adErrInTransaction = 3246
adErrFeatureNotAvailable = 3251
adErrItemNotFound = 3265
adErrObjectInCollection = 3367
adErrObjectNotSet = 3420
adErrDataConversion = 3421
adErrObjectClosed = 3704
adErrObjectOpen = 3705
adErrProviderNotFound = 3706
adErrBoundToCommand = 3707
adErrInvalidParamInfo = 3708
adErrInvalidConnection = 3709
adErrNotReentrant = 3710
adErrStillExecuting = 3711
adErrOperationCancelled = 3712
adErrStillConnecting = 3713
adErrInvalidTransaction = 3714
adErrNotExecuting = 3715
adErrUnsafeOperation = 3716
adwrnSecurityDialog = 3717
adwrnSecurityDialogHeader = 3718
adErrIntegrityViolation = 3719
adErrPermissionDenied = 3720
adErrDataOverflow = 3721
adErrSchemaViolation = 3722
adErrSignMismatch = 3723
adErrCantConvertvalue = 3724
adErrCantCreate = 3725
adErrColumnNotOnThisRow = 3726
adErrURLDoesNotExist = 3727
adErrTreePermissionDenied = 3728
adErrInvalidURL = 3729
adErrResourceLocked = 3730
adErrResourceExists = 3731
adErrCannotComplete = 3732
adErrVolumeNotFound = 3733
adErrOutOfSpace = 3734
adErrResourceOutOfScope = 3735
adErrUnavailable = 3736
adErrURLNamedRowDoesNotExist = 3737
adErrDelResOutOfScope = 3738
adErrPropInvalidColumn = 3739
adErrPropInvalidOption = 3740
adErrPropInvalidValue = 3741
adErrPropConflicting = 3742
adErrPropNotAllSettable = 3743
adErrPropNotSet = 3744
adErrPropNotSettable = 3745
adErrPropNotSupported = 3746
adErrCatalogNotSet = 3747
adErrCantChangeConnection = 3748
adErrFieldsUpdateFailed = 3749
adErrDenyNotSupported = 3750
adErrDenyTypeNotSupported = 3751
ErrorValueEnum = c_int # enum
class Errors(_Collection):
    _case_insensitive_ = True
    _iid_ = GUID('{00000501-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']

# values for enumeration 'IsolationLevelEnum'
adXactUnspecified = -1
adXactChaos = 16
adXactReadUncommitted = 256
adXactBrowse = 256
adXactCursorStability = 4096
adXactReadCommitted = 4096
adXactRepeatableRead = 65536
adXactSerializable = 1048576
adXactIsolated = 1048576
IsolationLevelEnum = c_int # enum

# values for enumeration 'SchemaEnum'
adSchemaProviderSpecific = -1
adSchemaAsserts = 0
adSchemaCatalogs = 1
adSchemaCharacterSets = 2
adSchemaCollations = 3
adSchemaColumns = 4
adSchemaCheckConstraints = 5
adSchemaConstraintColumnUsage = 6
adSchemaConstraintTableUsage = 7
adSchemaKeyColumnUsage = 8
adSchemaReferentialContraints = 9
adSchemaReferentialConstraints = 9
adSchemaTableConstraints = 10
adSchemaColumnsDomainUsage = 11
adSchemaIndexes = 12
adSchemaColumnPrivileges = 13
adSchemaTablePrivileges = 14
adSchemaUsagePrivileges = 15
adSchemaProcedures = 16
adSchemaSchemata = 17
adSchemaSQLLanguages = 18
adSchemaStatistics = 19
adSchemaTables = 20
adSchemaTranslations = 21
adSchemaProviderTypes = 22
adSchemaViews = 23
adSchemaViewColumnUsage = 24
adSchemaViewTableUsage = 25
adSchemaProcedureParameters = 26
adSchemaForeignKeys = 27
adSchemaPrimaryKeys = 28
adSchemaProcedureColumns = 29
adSchemaDBInfoKeywords = 30
adSchemaDBInfoLiterals = 31
adSchemaCubes = 32
adSchemaDimensions = 33
adSchemaHierarchies = 34
adSchemaLevels = 35
adSchemaMeasures = 36
adSchemaProperties = 37
adSchemaMembers = 38
adSchemaTrustees = 39
SchemaEnum = c_int # enum
Connection15._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'ConnectionString',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(0), 'propput'], HRESULT, 'ConnectionString',
              ( ['in'], BSTR, 'pbstr' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'CommandTimeout',
              ( ['retval', 'out'], POINTER(c_int), 'plTimeout' )),
    COMMETHOD([dispid(2), 'propput'], HRESULT, 'CommandTimeout',
              ( ['in'], c_int, 'plTimeout' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'ConnectionTimeout',
              ( ['retval', 'out'], POINTER(c_int), 'plTimeout' )),
    COMMETHOD([dispid(3), 'propput'], HRESULT, 'ConnectionTimeout',
              ( ['in'], c_int, 'plTimeout' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'Version',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(5)], HRESULT, 'Close'),
    COMMETHOD([dispid(6)], HRESULT, 'Execute',
              ( ['in'], BSTR, 'CommandText' ),
              ( ['out', 'optional'], POINTER(VARIANT), 'RecordsAffected' ),
              ( ['in', 'optional'], c_int, 'Options', -1 ),
              ( ['retval', 'out'], POINTER(POINTER(_Recordset)), 'ppiRset' )),
    COMMETHOD([dispid(7)], HRESULT, 'BeginTrans',
              ( ['retval', 'out'], POINTER(c_int), 'TransactionLevel' )),
    COMMETHOD([dispid(8)], HRESULT, 'CommitTrans'),
    COMMETHOD([dispid(9)], HRESULT, 'RollbackTrans'),
    COMMETHOD([dispid(10)], HRESULT, 'Open',
              ( ['in', 'optional'], BSTR, 'ConnectionString', '' ),
              ( ['in', 'optional'], BSTR, 'UserID', '' ),
              ( ['in', 'optional'], BSTR, 'Password', '' ),
              ( ['in', 'optional'], c_int, 'Options', -1 )),
    COMMETHOD([dispid(11), 'propget'], HRESULT, 'Errors',
              ( ['retval', 'out'], POINTER(POINTER(Errors)), 'ppvObject' )),
    COMMETHOD([dispid(12), 'propget'], HRESULT, 'DefaultDatabase',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(12), 'propput'], HRESULT, 'DefaultDatabase',
              ( ['in'], BSTR, 'pbstr' )),
    COMMETHOD([dispid(13), 'propget'], HRESULT, 'IsolationLevel',
              ( ['retval', 'out'], POINTER(IsolationLevelEnum), 'Level' )),
    COMMETHOD([dispid(13), 'propput'], HRESULT, 'IsolationLevel',
              ( ['in'], IsolationLevelEnum, 'Level' )),
    COMMETHOD([dispid(14), 'propget'], HRESULT, 'Attributes',
              ( ['retval', 'out'], POINTER(c_int), 'plAttr' )),
    COMMETHOD([dispid(14), 'propput'], HRESULT, 'Attributes',
              ( ['in'], c_int, 'plAttr' )),
    COMMETHOD([dispid(15), 'propget'], HRESULT, 'CursorLocation',
              ( ['retval', 'out'], POINTER(CursorLocationEnum), 'plCursorLoc' )),
    COMMETHOD([dispid(15), 'propput'], HRESULT, 'CursorLocation',
              ( ['in'], CursorLocationEnum, 'plCursorLoc' )),
    COMMETHOD([dispid(16), 'propget'], HRESULT, 'Mode',
              ( ['retval', 'out'], POINTER(ConnectModeEnum), 'plMode' )),
    COMMETHOD([dispid(16), 'propput'], HRESULT, 'Mode',
              ( ['in'], ConnectModeEnum, 'plMode' )),
    COMMETHOD([dispid(17), 'propget'], HRESULT, 'Provider',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(17), 'propput'], HRESULT, 'Provider',
              ( ['in'], BSTR, 'pbstr' )),
    COMMETHOD([dispid(18), 'propget'], HRESULT, 'State',
              ( ['retval', 'out'], POINTER(c_int), 'plObjState' )),
    COMMETHOD([dispid(19)], HRESULT, 'OpenSchema',
              ( ['in'], SchemaEnum, 'Schema' ),
              ( ['in', 'optional'], VARIANT, 'Restrictions' ),
              ( ['in', 'optional'], VARIANT, 'SchemaID' ),
              ( ['retval', 'out'], POINTER(POINTER(_Recordset)), 'pprset' )),
]
################################################################
## code template for Connection15 implementation
##class Connection15_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return plTimeout
##    def _set(self, plTimeout):
##        '-no docstring-'
##    ConnectionTimeout = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstr
##    def _set(self, pbstr):
##        '-no docstring-'
##    ConnectionString = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def State(self):
##        '-no docstring-'
##        #return plObjState
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstr
##    def _set(self, pbstr):
##        '-no docstring-'
##    DefaultDatabase = property(_get, _set, doc = _set.__doc__)
##
##    def OpenSchema(self, Schema, Restrictions, SchemaID):
##        '-no docstring-'
##        #return pprset
##
##    def Close(self):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return plAttr
##    def _set(self, plAttr):
##        '-no docstring-'
##    Attributes = property(_get, _set, doc = _set.__doc__)
##
##    def Open(self, ConnectionString, UserID, Password, Options):
##        '-no docstring-'
##        #return 
##
##    def RollbackTrans(self):
##        '-no docstring-'
##        #return 
##
##    def CommitTrans(self):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return plMode
##    def _set(self, plMode):
##        '-no docstring-'
##    Mode = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def Errors(self):
##        '-no docstring-'
##        #return ppvObject
##
##    def _get(self):
##        '-no docstring-'
##        #return plCursorLoc
##    def _set(self, plCursorLoc):
##        '-no docstring-'
##    CursorLocation = property(_get, _set, doc = _set.__doc__)
##
##    def BeginTrans(self):
##        '-no docstring-'
##        #return TransactionLevel
##
##    def Execute(self, CommandText, Options):
##        '-no docstring-'
##        #return RecordsAffected, ppiRset
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstr
##    def _set(self, pbstr):
##        '-no docstring-'
##    Provider = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return plTimeout
##    def _set(self, plTimeout):
##        '-no docstring-'
##    CommandTimeout = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return Level
##    def _set(self, Level):
##        '-no docstring-'
##    IsolationLevel = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def Version(self):
##        '-no docstring-'
##        #return pbstr
##

_Connection._methods_ = [
    COMMETHOD([dispid(21)], HRESULT, 'Cancel'),
]
################################################################
## code template for _Connection implementation
##class _Connection_Impl(object):
##    def Cancel(self):
##        '-no docstring-'
##        #return 
##


# values for enumeration 'ConnectPromptEnum'
adPromptAlways = 1
adPromptComplete = 2
adPromptCompleteRequired = 3
adPromptNever = 4
ConnectPromptEnum = c_int # enum
Field20._methods_ = [
    COMMETHOD([dispid(1109), 'propget'], HRESULT, 'ActualSize',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1114), 'propget'], HRESULT, 'Attributes',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1103), 'propget'], HRESULT, 'DefinedSize',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1100), 'propget'], HRESULT, 'Name',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(1102), 'propget'], HRESULT, 'Type',
              ( ['retval', 'out'], POINTER(DataTypeEnum), 'pDataType' )),
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Value',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(0), 'propput'], HRESULT, 'Value',
              ( ['in'], VARIANT, 'pvar' )),
    COMMETHOD([dispid(1112), 'propget'], HRESULT, 'Precision',
              ( ['retval', 'out'], POINTER(c_ubyte), 'pbPrecision' )),
    COMMETHOD([dispid(1113), 'propget'], HRESULT, 'NumericScale',
              ( ['retval', 'out'], POINTER(c_ubyte), 'pbNumericScale' )),
    COMMETHOD([dispid(1107)], HRESULT, 'AppendChunk',
              ( ['in'], VARIANT, 'Data' )),
    COMMETHOD([dispid(1108)], HRESULT, 'GetChunk',
              ( ['in'], c_int, 'Length' ),
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(1104), 'propget'], HRESULT, 'OriginalValue',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(1105), 'propget'], HRESULT, 'UnderlyingValue',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(1115), 'propget'], HRESULT, 'DataFormat',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppiDF' )),
    COMMETHOD([dispid(1115), 'propputref'], HRESULT, 'DataFormat',
              ( ['in'], POINTER(IUnknown), 'ppiDF' )),
    COMMETHOD([dispid(1112), 'propput'], HRESULT, 'Precision',
              ( ['in'], c_ubyte, 'pbPrecision' )),
    COMMETHOD([dispid(1113), 'propput'], HRESULT, 'NumericScale',
              ( ['in'], c_ubyte, 'pbNumericScale' )),
    COMMETHOD([dispid(1102), 'propput'], HRESULT, 'Type',
              ( ['in'], DataTypeEnum, 'pDataType' )),
    COMMETHOD([dispid(1103), 'propput'], HRESULT, 'DefinedSize',
              ( ['in'], c_int, 'pl' )),
    COMMETHOD([dispid(1114), 'propput'], HRESULT, 'Attributes',
              ( ['in'], c_int, 'pl' )),
]
################################################################
## code template for Field20 implementation
##class Field20_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return pl
##    def _set(self, pl):
##        '-no docstring-'
##    Attributes = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pDataType
##    def _set(self, pDataType):
##        '-no docstring-'
##    Type = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def OriginalValue(self):
##        '-no docstring-'
##        #return pvar
##
##    def _get(self):
##        '-no docstring-'
##        #return pl
##    def _set(self, pl):
##        '-no docstring-'
##    DefinedSize = property(_get, _set, doc = _set.__doc__)
##
##    def AppendChunk(self, Data):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pbNumericScale
##    def _set(self, pbNumericScale):
##        '-no docstring-'
##    NumericScale = property(_get, _set, doc = _set.__doc__)
##
##    def DataFormat(self, ppiDF):
##        '-no docstring-'
##        #return 
##
##    @property
##    def UnderlyingValue(self):
##        '-no docstring-'
##        #return pvar
##
##    @property
##    def ActualSize(self):
##        '-no docstring-'
##        #return pl
##
##    def GetChunk(self, Length):
##        '-no docstring-'
##        #return pvar
##
##    def _get(self):
##        '-no docstring-'
##        #return pbPrecision
##    def _set(self, pbPrecision):
##        '-no docstring-'
##    Precision = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def Name(self):
##        '-no docstring-'
##        #return pbstr
##
##    def _get(self):
##        '-no docstring-'
##        #return pvar
##    def _set(self, pvar):
##        '-no docstring-'
##    Value = property(_get, _set, doc = _set.__doc__)
##

class _DynaCollection(_Collection):
    _case_insensitive_ = True
    _iid_ = GUID('{00000513-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
_DynaCollection._methods_ = [
    COMMETHOD([dispid(1610809344)], HRESULT, 'Append',
              ( ['in'], POINTER(IDispatch), 'Object' )),
    COMMETHOD([dispid(1610809345)], HRESULT, 'Delete',
              ( ['in'], VARIANT, 'Index' )),
]
################################################################
## code template for _DynaCollection implementation
##class _DynaCollection_Impl(object):
##    def Append(self, Object):
##        '-no docstring-'
##        #return 
##
##    def Delete(self, Index):
##        '-no docstring-'
##        #return 
##


# values for enumeration 'PropertyAttributesEnum'
adPropNotSupported = 0
adPropRequired = 1
adPropOptional = 2
adPropRead = 512
adPropWrite = 1024
PropertyAttributesEnum = c_int # enum

# values for enumeration 'CursorOptionEnum'
adHoldRecords = 256
adMovePrevious = 512
adAddNew = 16778240
adDelete = 16779264
adUpdate = 16809984
adBookmark = 8192
adApproxPosition = 16384
adUpdateBatch = 65536
adResync = 131072
adNotify = 262144
adFind = 524288
adSeek = 4194304
adIndex = 8388608
CursorOptionEnum = c_int # enum
class Stream(CoClass):
    _reg_clsid_ = GUID('{00000566-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['licensed']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{00000205-0000-0010-8000-00AA006D2EA4}', 2, 5)
class _Stream(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000565-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'oleautomation']
Stream._com_interfaces_ = [_Stream]

class ADOStreamConstruction(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000568-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['restricted']
ADOStreamConstruction._methods_ = [
    COMMETHOD(['propget'], HRESULT, 'Stream',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppStm' )),
    COMMETHOD(['propput'], HRESULT, 'Stream',
              ( ['in'], POINTER(IUnknown), 'ppStm' )),
]
################################################################
## code template for ADOStreamConstruction implementation
##class ADOStreamConstruction_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return ppStm
##    def _set(self, ppStm):
##        '-no docstring-'
##    Stream = property(_get, _set, doc = _set.__doc__)
##

SearchDirection = SearchDirectionEnum

# values for enumeration 'XactAttributeEnum'
adXactCommitRetaining = 131072
adXactAbortRetaining = 262144
adXactAsyncPhaseOne = 524288
adXactSyncPhaseOne = 1048576
XactAttributeEnum = c_int # enum

# values for enumeration 'EditModeEnum'
adEditNone = 0
adEditInProgress = 1
adEditAdd = 2
adEditDelete = 4
EditModeEnum = c_int # enum
Field._methods_ = [
    COMMETHOD([dispid(1116), 'propget'], HRESULT, 'Status',
              ( ['retval', 'out'], POINTER(c_int), 'pFStatus' )),
]
################################################################
## code template for Field implementation
##class Field_Impl(object):
##    @property
##    def Status(self):
##        '-no docstring-'
##        #return pFStatus
##


# values for enumeration 'RecordStatusEnum'
adRecOK = 0
adRecNew = 1
adRecModified = 2
adRecDeleted = 4
adRecUnmodified = 8
adRecInvalid = 16
adRecMultipleChanges = 64
adRecPendingChanges = 128
adRecCanceled = 256
adRecCantRelease = 1024
adRecConcurrencyViolation = 2048
adRecIntegrityViolation = 4096
adRecMaxChangesExceeded = 8192
adRecObjectOpen = 16384
adRecOutOfMemory = 32768
adRecPermissionDenied = 65536
adRecSchemaViolation = 131072
adRecDBDeleted = 262144
RecordStatusEnum = c_int # enum
class ADORecordsetConstruction(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000283-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['restricted']
ADORecordsetConstruction._methods_ = [
    COMMETHOD(['propget'], HRESULT, 'Rowset',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppRowset' )),
    COMMETHOD(['propput'], HRESULT, 'Rowset',
              ( ['in'], POINTER(IUnknown), 'ppRowset' )),
    COMMETHOD(['propget'], HRESULT, 'Chapter',
              ( ['retval', 'out'], POINTER(c_int), 'plChapter' )),
    COMMETHOD(['propput'], HRESULT, 'Chapter',
              ( ['in'], c_int, 'plChapter' )),
    COMMETHOD(['propget'], HRESULT, 'RowPosition',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppRowPos' )),
    COMMETHOD(['propput'], HRESULT, 'RowPosition',
              ( ['in'], POINTER(IUnknown), 'ppRowPos' )),
]
################################################################
## code template for ADORecordsetConstruction implementation
##class ADORecordsetConstruction_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return ppRowset
##    def _set(self, ppRowset):
##        '-no docstring-'
##    Rowset = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return plChapter
##    def _set(self, plChapter):
##        '-no docstring-'
##    Chapter = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return ppRowPos
##    def _set(self, ppRowPos):
##        '-no docstring-'
##    RowPosition = property(_get, _set, doc = _set.__doc__)
##

class ADOCommandConstruction(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{00000517-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['restricted']
ADOCommandConstruction._methods_ = [
    COMMETHOD(['propget'], HRESULT, 'OLEDBCommand',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppOLEDBCommand' )),
    COMMETHOD(['propput'], HRESULT, 'OLEDBCommand',
              ( ['in'], POINTER(IUnknown), 'ppOLEDBCommand' )),
]
################################################################
## code template for ADOCommandConstruction implementation
##class ADOCommandConstruction_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return ppOLEDBCommand
##    def _set(self, ppOLEDBCommand):
##        '-no docstring-'
##    OLEDBCommand = property(_get, _set, doc = _set.__doc__)
##

class Record(CoClass):
    _reg_clsid_ = GUID('{00000560-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['licensed']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{00000205-0000-0010-8000-00AA006D2EA4}', 2, 5)
Record._com_interfaces_ = [_Record]


# values for enumeration 'PositionEnum'
adPosUnknown = -1
adPosBOF = -2
adPosEOF = -3
PositionEnum = c_int # enum
Recordset15._methods_ = [
    COMMETHOD([dispid(1000), 'propget'], HRESULT, 'AbsolutePosition',
              ( ['retval', 'out'], POINTER(PositionEnum), 'pl' )),
    COMMETHOD([dispid(1000), 'propput'], HRESULT, 'AbsolutePosition',
              ( ['in'], PositionEnum, 'pl' )),
    COMMETHOD([dispid(1001), 'propputref'], HRESULT, 'ActiveConnection',
              ( ['in'], POINTER(IDispatch), 'pvar' )),
    COMMETHOD([dispid(1001), 'propput'], HRESULT, 'ActiveConnection',
              ( ['in'], VARIANT, 'pvar' )),
    COMMETHOD([dispid(1001), 'propget'], HRESULT, 'ActiveConnection',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(1002), 'propget'], HRESULT, 'BOF',
              ( ['retval', 'out'], POINTER(VARIANT_BOOL), 'pb' )),
    COMMETHOD([dispid(1003), 'propget'], HRESULT, 'Bookmark',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvBookmark' )),
    COMMETHOD([dispid(1003), 'propput'], HRESULT, 'Bookmark',
              ( ['in'], VARIANT, 'pvBookmark' )),
    COMMETHOD([dispid(1004), 'propget'], HRESULT, 'CacheSize',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1004), 'propput'], HRESULT, 'CacheSize',
              ( ['in'], c_int, 'pl' )),
    COMMETHOD([dispid(1005), 'propget'], HRESULT, 'CursorType',
              ( ['retval', 'out'], POINTER(CursorTypeEnum), 'plCursorType' )),
    COMMETHOD([dispid(1005), 'propput'], HRESULT, 'CursorType',
              ( ['in'], CursorTypeEnum, 'plCursorType' )),
    COMMETHOD([dispid(1006), 'propget'], HRESULT, 'EOF',
              ( ['retval', 'out'], POINTER(VARIANT_BOOL), 'pb' )),
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Fields',
              ( ['retval', 'out'], POINTER(POINTER(Fields)), 'ppvObject' )),
    COMMETHOD([dispid(1008), 'propget'], HRESULT, 'LockType',
              ( ['retval', 'out'], POINTER(LockTypeEnum), 'plLockType' )),
    COMMETHOD([dispid(1008), 'propput'], HRESULT, 'LockType',
              ( ['in'], LockTypeEnum, 'plLockType' )),
    COMMETHOD([dispid(1009), 'propget'], HRESULT, 'MaxRecords',
              ( ['retval', 'out'], POINTER(c_int), 'plMaxRecords' )),
    COMMETHOD([dispid(1009), 'propput'], HRESULT, 'MaxRecords',
              ( ['in'], c_int, 'plMaxRecords' )),
    COMMETHOD([dispid(1010), 'propget'], HRESULT, 'RecordCount',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1011), 'propputref'], HRESULT, 'Source',
              ( ['in'], POINTER(IDispatch), 'pvSource' )),
    COMMETHOD([dispid(1011), 'propput'], HRESULT, 'Source',
              ( ['in'], BSTR, 'pvSource' )),
    COMMETHOD([dispid(1011), 'propget'], HRESULT, 'Source',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvSource' )),
    COMMETHOD([dispid(1012)], HRESULT, 'AddNew',
              ( ['in', 'optional'], VARIANT, 'FieldList' ),
              ( ['in', 'optional'], VARIANT, 'Values' )),
    COMMETHOD([dispid(1013)], HRESULT, 'CancelUpdate'),
    COMMETHOD([dispid(1014)], HRESULT, 'Close'),
    COMMETHOD([dispid(1015)], HRESULT, 'Delete',
              ( ['in', 'optional'], AffectEnum, 'AffectRecords', 1 )),
    COMMETHOD([dispid(1016)], HRESULT, 'GetRows',
              ( ['in', 'optional'], c_int, 'Rows', -1 ),
              ( ['in', 'optional'], VARIANT, 'Start' ),
              ( ['in', 'optional'], VARIANT, 'Fields' ),
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(1017)], HRESULT, 'Move',
              ( ['in'], c_int, 'NumRecords' ),
              ( ['in', 'optional'], VARIANT, 'Start' )),
    COMMETHOD([dispid(1018)], HRESULT, 'MoveNext'),
    COMMETHOD([dispid(1019)], HRESULT, 'MovePrevious'),
    COMMETHOD([dispid(1020)], HRESULT, 'MoveFirst'),
    COMMETHOD([dispid(1021)], HRESULT, 'MoveLast'),
    COMMETHOD([dispid(1022)], HRESULT, 'Open',
              ( ['in', 'optional'], VARIANT, 'Source' ),
              ( ['in', 'optional'], VARIANT, 'ActiveConnection' ),
              ( ['in', 'optional'], CursorTypeEnum, 'CursorType', -1 ),
              ( ['in', 'optional'], LockTypeEnum, 'LockType', -1 ),
              ( ['in', 'optional'], c_int, 'Options', -1 )),
    COMMETHOD([dispid(1023)], HRESULT, 'Requery',
              ( ['in', 'optional'], c_int, 'Options', -1 )),
    COMMETHOD([dispid(1610809378), 'hidden'], HRESULT, '_xResync',
              ( ['in', 'optional'], AffectEnum, 'AffectRecords', 3 )),
    COMMETHOD([dispid(1025)], HRESULT, 'Update',
              ( ['in', 'optional'], VARIANT, 'Fields' ),
              ( ['in', 'optional'], VARIANT, 'Values' )),
    COMMETHOD([dispid(1047), 'propget'], HRESULT, 'AbsolutePage',
              ( ['retval', 'out'], POINTER(PositionEnum), 'pl' )),
    COMMETHOD([dispid(1047), 'propput'], HRESULT, 'AbsolutePage',
              ( ['in'], PositionEnum, 'pl' )),
    COMMETHOD([dispid(1026), 'propget'], HRESULT, 'EditMode',
              ( ['retval', 'out'], POINTER(EditModeEnum), 'pl' )),
    COMMETHOD([dispid(1030), 'propget'], HRESULT, 'Filter',
              ( ['retval', 'out'], POINTER(VARIANT), 'Criteria' )),
    COMMETHOD([dispid(1030), 'propput'], HRESULT, 'Filter',
              ( ['in'], VARIANT, 'Criteria' )),
    COMMETHOD([dispid(1050), 'propget'], HRESULT, 'PageCount',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1048), 'propget'], HRESULT, 'PageSize',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1048), 'propput'], HRESULT, 'PageSize',
              ( ['in'], c_int, 'pl' )),
    COMMETHOD([dispid(1031), 'propget'], HRESULT, 'Sort',
              ( ['retval', 'out'], POINTER(BSTR), 'Criteria' )),
    COMMETHOD([dispid(1031), 'propput'], HRESULT, 'Sort',
              ( ['in'], BSTR, 'Criteria' )),
    COMMETHOD([dispid(1029), 'propget'], HRESULT, 'Status',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1054), 'propget'], HRESULT, 'State',
              ( ['retval', 'out'], POINTER(c_int), 'plObjState' )),
    COMMETHOD([dispid(1610809392), 'hidden'], HRESULT, '_xClone',
              ( ['retval', 'out'], POINTER(POINTER(_Recordset)), 'ppvObject' )),
    COMMETHOD([dispid(1035)], HRESULT, 'UpdateBatch',
              ( ['in', 'optional'], AffectEnum, 'AffectRecords', 3 )),
    COMMETHOD([dispid(1049)], HRESULT, 'CancelBatch',
              ( ['in', 'optional'], AffectEnum, 'AffectRecords', 3 )),
    COMMETHOD([dispid(1051), 'propget'], HRESULT, 'CursorLocation',
              ( ['retval', 'out'], POINTER(CursorLocationEnum), 'plCursorLoc' )),
    COMMETHOD([dispid(1051), 'propput'], HRESULT, 'CursorLocation',
              ( ['in'], CursorLocationEnum, 'plCursorLoc' )),
    COMMETHOD([dispid(1052)], HRESULT, 'NextRecordset',
              ( ['out', 'optional'], POINTER(VARIANT), 'RecordsAffected' ),
              ( ['retval', 'out'], POINTER(POINTER(_Recordset)), 'ppiRs' )),
    COMMETHOD([dispid(1036)], HRESULT, 'Supports',
              ( ['in'], CursorOptionEnum, 'CursorOptions' ),
              ( ['retval', 'out'], POINTER(VARIANT_BOOL), 'pb' )),
    COMMETHOD([dispid(-8), 'hidden', 'propget'], HRESULT, 'Collect',
              ( ['in'], VARIANT, 'Index' ),
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(-8), 'hidden', 'propput'], HRESULT, 'Collect',
              ( ['in'], VARIANT, 'Index' ),
              ( ['in'], VARIANT, 'pvar' )),
    COMMETHOD([dispid(1053), 'propget'], HRESULT, 'MarshalOptions',
              ( ['retval', 'out'], POINTER(MarshalOptionsEnum), 'peMarshal' )),
    COMMETHOD([dispid(1053), 'propput'], HRESULT, 'MarshalOptions',
              ( ['in'], MarshalOptionsEnum, 'peMarshal' )),
    COMMETHOD([dispid(1058)], HRESULT, 'Find',
              ( ['in'], BSTR, 'Criteria' ),
              ( ['in', 'optional'], c_int, 'SkipRecords', 0 ),
              ( ['in', 'optional'], SearchDirectionEnum, 'SearchDirection', 1 ),
              ( ['in', 'optional'], VARIANT, 'Start' )),
]
################################################################
## code template for Recordset15 implementation
##class Recordset15_Impl(object):
##    def Close(self):
##        '-no docstring-'
##        #return 
##
##    @property
##    def State(self):
##        '-no docstring-'
##        #return plObjState
##
##    @property
##    def Status(self):
##        '-no docstring-'
##        #return pl
##
##    def AddNew(self, FieldList, Values):
##        '-no docstring-'
##        #return 
##
##    def CancelBatch(self, AffectRecords):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return Criteria
##    def _set(self, Criteria):
##        '-no docstring-'
##    Filter = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return plCursorType
##    def _set(self, plCursorType):
##        '-no docstring-'
##    CursorType = property(_get, _set, doc = _set.__doc__)
##
##    def GetRows(self, Rows, Start, Fields):
##        '-no docstring-'
##        #return pvar
##
##    def _xResync(self, AffectRecords):
##        '-no docstring-'
##        #return 
##
##    def _get(self, pvar):
##        '-no docstring-'
##        #return 
##    def _set(self, pvar):
##        '-no docstring-'
##    ActiveConnection = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return Criteria
##    def _set(self, Criteria):
##        '-no docstring-'
##    Sort = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self, pvSource):
##        '-no docstring-'
##        #return 
##    def _set(self, pvSource):
##        '-no docstring-'
##    Source = property(_get, _set, doc = _set.__doc__)
##
##    def Move(self, NumRecords, Start):
##        '-no docstring-'
##        #return 
##
##    def Requery(self, Options):
##        '-no docstring-'
##        #return 
##
##    def Update(self, Fields, Values):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pl
##    def _set(self, pl):
##        '-no docstring-'
##    CacheSize = property(_get, _set, doc = _set.__doc__)
##
##    def NextRecordset(self):
##        '-no docstring-'
##        #return RecordsAffected, ppiRs
##
##    def Supports(self, CursorOptions):
##        '-no docstring-'
##        #return pb
##
##    @property
##    def EditMode(self):
##        '-no docstring-'
##        #return pl
##
##    def _get(self):
##        '-no docstring-'
##        #return pl
##    def _set(self, pl):
##        '-no docstring-'
##    AbsolutePosition = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pl
##    def _set(self, pl):
##        '-no docstring-'
##    PageSize = property(_get, _set, doc = _set.__doc__)
##
##    def MoveNext(self):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return plMaxRecords
##    def _set(self, plMaxRecords):
##        '-no docstring-'
##    MaxRecords = property(_get, _set, doc = _set.__doc__)
##
##    def Delete(self, AffectRecords):
##        '-no docstring-'
##        #return 
##
##    @property
##    def BOF(self):
##        '-no docstring-'
##        #return pb
##
##    def _xClone(self):
##        '-no docstring-'
##        #return ppvObject
##
##    def MoveLast(self):
##        '-no docstring-'
##        #return 
##
##    def MovePrevious(self):
##        '-no docstring-'
##        #return 
##
##    @property
##    def PageCount(self):
##        '-no docstring-'
##        #return pl
##
##    def _get(self):
##        '-no docstring-'
##        #return plCursorLoc
##    def _set(self, plCursorLoc):
##        '-no docstring-'
##    CursorLocation = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def EOF(self):
##        '-no docstring-'
##        #return pb
##
##    def Find(self, Criteria, SkipRecords, SearchDirection, Start):
##        '-no docstring-'
##        #return 
##
##    def UpdateBatch(self, AffectRecords):
##        '-no docstring-'
##        #return 
##
##    def MoveFirst(self):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return plLockType
##    def _set(self, plLockType):
##        '-no docstring-'
##    LockType = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pl
##    def _set(self, pl):
##        '-no docstring-'
##    AbsolutePage = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pvBookmark
##    def _set(self, pvBookmark):
##        '-no docstring-'
##    Bookmark = property(_get, _set, doc = _set.__doc__)
##
##    def CancelUpdate(self):
##        '-no docstring-'
##        #return 
##
##    def _get(self, Index):
##        '-no docstring-'
##        #return pvar
##    def _set(self, Index, pvar):
##        '-no docstring-'
##    Collect = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def RecordCount(self):
##        '-no docstring-'
##        #return pl
##
##    def _get(self):
##        '-no docstring-'
##        #return peMarshal
##    def _set(self, peMarshal):
##        '-no docstring-'
##    MarshalOptions = property(_get, _set, doc = _set.__doc__)
##
##    def Open(self, Source, ActiveConnection, CursorType, LockType, Options):
##        '-no docstring-'
##        #return 
##
##    @property
##    def Fields(self):
##        '-no docstring-'
##        #return ppvObject
##


# values for enumeration 'CompareEnum'
adCompareLessThan = 0
adCompareEqual = 1
adCompareGreaterThan = 2
adCompareNotEqual = 3
adCompareNotComparable = 4
CompareEnum = c_int # enum
Recordset20._methods_ = [
    COMMETHOD([dispid(1055)], HRESULT, 'Cancel'),
    COMMETHOD([dispid(1056), 'propget'], HRESULT, 'DataSource',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppunkDataSource' )),
    COMMETHOD([dispid(1056), 'propputref'], HRESULT, 'DataSource',
              ( ['in'], POINTER(IUnknown), 'ppunkDataSource' )),
    COMMETHOD([dispid(1610874883), 'hidden'], HRESULT, '_xSave',
              ( ['in', 'optional'], BSTR, 'FileName', '' ),
              ( ['in', 'optional'], PersistFormatEnum, 'PersistFormat', 0 )),
    COMMETHOD([dispid(1061), 'propget'], HRESULT, 'ActiveCommand',
              ( ['retval', 'out'], POINTER(POINTER(IDispatch)), 'ppCmd' )),
    COMMETHOD([dispid(1063), 'propput'], HRESULT, 'StayInSync',
              ( ['in'], VARIANT_BOOL, 'pbStayInSync' )),
    COMMETHOD([dispid(1063), 'propget'], HRESULT, 'StayInSync',
              ( ['retval', 'out'], POINTER(VARIANT_BOOL), 'pbStayInSync' )),
    COMMETHOD([dispid(1062)], HRESULT, 'GetString',
              ( ['in', 'optional'], StringFormatEnum, 'StringFormat', 2 ),
              ( ['in', 'optional'], c_int, 'NumRows', -1 ),
              ( ['in', 'optional'], BSTR, 'ColumnDelimeter', '' ),
              ( ['in', 'optional'], BSTR, 'RowDelimeter', '' ),
              ( ['in', 'optional'], BSTR, 'NullExpr', '' ),
              ( ['retval', 'out'], POINTER(BSTR), 'pRetString' )),
    COMMETHOD([dispid(1064), 'propget'], HRESULT, 'DataMember',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstrDataMember' )),
    COMMETHOD([dispid(1064), 'propput'], HRESULT, 'DataMember',
              ( ['in'], BSTR, 'pbstrDataMember' )),
    COMMETHOD([dispid(1065)], HRESULT, 'CompareBookmarks',
              ( ['in'], VARIANT, 'Bookmark1' ),
              ( ['in'], VARIANT, 'Bookmark2' ),
              ( ['retval', 'out'], POINTER(CompareEnum), 'pCompare' )),
    COMMETHOD([dispid(1034)], HRESULT, 'Clone',
              ( ['in', 'optional'], LockTypeEnum, 'LockType', -1 ),
              ( ['retval', 'out'], POINTER(POINTER(_Recordset)), 'ppvObject' )),
    COMMETHOD([dispid(1024)], HRESULT, 'Resync',
              ( ['in', 'optional'], AffectEnum, 'AffectRecords', 3 ),
              ( ['in', 'optional'], ResyncEnum, 'ResyncValues', 2 )),
]
################################################################
## code template for Recordset20 implementation
##class Recordset20_Impl(object):
##    def Cancel(self):
##        '-no docstring-'
##        #return 
##
##    def Clone(self, LockType):
##        '-no docstring-'
##        #return ppvObject
##
##    def _xSave(self, FileName, PersistFormat):
##        '-no docstring-'
##        #return 
##
##    def Resync(self, AffectRecords, ResyncValues):
##        '-no docstring-'
##        #return 
##
##    def GetString(self, StringFormat, NumRows, ColumnDelimeter, RowDelimeter, NullExpr):
##        '-no docstring-'
##        #return pRetString
##
##    def DataSource(self, ppunkDataSource):
##        '-no docstring-'
##        #return 
##
##    @property
##    def ActiveCommand(self):
##        '-no docstring-'
##        #return ppCmd
##
##    def CompareBookmarks(self, Bookmark1, Bookmark2):
##        '-no docstring-'
##        #return pCompare
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstrDataMember
##    def _set(self, pbstrDataMember):
##        '-no docstring-'
##    DataMember = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pbStayInSync
##    def _set(self, pbStayInSync):
##        '-no docstring-'
##    StayInSync = property(_get, _set, doc = _set.__doc__)
##


# values for enumeration 'SeekEnum'
adSeekFirstEQ = 1
adSeekLastEQ = 2
adSeekAfterEQ = 4
adSeekAfter = 8
adSeekBeforeEQ = 16
adSeekBefore = 32
SeekEnum = c_int # enum
Recordset21._methods_ = [
    COMMETHOD([dispid(1066)], HRESULT, 'Seek',
              ( ['in'], VARIANT, 'KeyValues' ),
              ( ['in', 'optional'], SeekEnum, 'SeekOption', 1 )),
    COMMETHOD([dispid(1067), 'propput'], HRESULT, 'Index',
              ( ['in'], BSTR, 'pbstrIndex' )),
    COMMETHOD([dispid(1067), 'propget'], HRESULT, 'Index',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstrIndex' )),
]
################################################################
## code template for Recordset21 implementation
##class Recordset21_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return pbstrIndex
##    def _set(self, pbstrIndex):
##        '-no docstring-'
##    Index = property(_get, _set, doc = _set.__doc__)
##
##    def Seek(self, KeyValues, SeekOption):
##        '-no docstring-'
##        #return 
##

class RecordsetEventsVt(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{00000403-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden']

# values for enumeration 'EventReasonEnum'
adRsnAddNew = 1
adRsnDelete = 2
adRsnUpdate = 3
adRsnUndoUpdate = 4
adRsnUndoAddNew = 5
adRsnUndoDelete = 6
adRsnRequery = 7
adRsnResynch = 8
adRsnClose = 9
adRsnMove = 10
adRsnFirstChange = 11
adRsnMoveFirst = 12
adRsnMoveNext = 13
adRsnMovePrevious = 14
adRsnMoveLast = 15
EventReasonEnum = c_int # enum
RecordsetEventsVt._methods_ = [
    COMMETHOD([], HRESULT, 'WillChangeField',
              ( ['in'], c_int, 'cFields' ),
              ( ['in'], VARIANT, 'Fields' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'FieldChangeComplete',
              ( ['in'], c_int, 'cFields' ),
              ( ['in'], VARIANT, 'Fields' ),
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'WillChangeRecord',
              ( ['in'], EventReasonEnum, 'adReason' ),
              ( ['in'], c_int, 'cRecords' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'RecordChangeComplete',
              ( ['in'], EventReasonEnum, 'adReason' ),
              ( ['in'], c_int, 'cRecords' ),
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'WillChangeRecordset',
              ( ['in'], EventReasonEnum, 'adReason' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'RecordsetChangeComplete',
              ( ['in'], EventReasonEnum, 'adReason' ),
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'WillMove',
              ( ['in'], EventReasonEnum, 'adReason' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'MoveComplete',
              ( ['in'], EventReasonEnum, 'adReason' ),
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'EndOfRecordset',
              ( ['in', 'out'], POINTER(VARIANT_BOOL), 'fMoreData' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'FetchProgress',
              ( ['in'], c_int, 'Progress' ),
              ( ['in'], c_int, 'MaxProgress' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    COMMETHOD([], HRESULT, 'FetchComplete',
              ( ['in'], POINTER(Error), 'pError' ),
              ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
              ( ['in'], POINTER(_Recordset), 'pRecordset' )),
]
################################################################
## code template for RecordsetEventsVt implementation
##class RecordsetEventsVt_Impl(object):
##    def RecordChangeComplete(self, adReason, cRecords, pError, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def WillChangeRecord(self, adReason, cRecords, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def WillChangeRecordset(self, adReason, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def RecordsetChangeComplete(self, adReason, pError, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def FieldChangeComplete(self, cFields, Fields, pError, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def WillChangeField(self, cFields, Fields, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def FetchComplete(self, pError, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def WillMove(self, adReason, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def FetchProgress(self, Progress, MaxProgress, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def MoveComplete(self, adReason, pError, pRecordset):
##        '-no docstring-'
##        #return adStatus
##
##    def EndOfRecordset(self, pRecordset):
##        '-no docstring-'
##        #return fMoreData, adStatus
##

class Library(object):
    'Microsoft ActiveX Data Objects 2.5 Library'
    name = 'ADODB'
    _reg_typelib_ = ('{00000205-0000-0010-8000-00AA006D2EA4}', 2, 5)

class Recordset(CoClass):
    _reg_clsid_ = GUID('{00000535-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['licensed']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{00000205-0000-0010-8000-00AA006D2EA4}', 2, 5)
class RecordsetEvents(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000266-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = []
    _methods_ = []
Recordset._com_interfaces_ = [_Recordset]
Recordset._outgoing_interfaces_ = [RecordsetEvents]


# values for enumeration 'StreamTypeEnum'
adTypeBinary = 1
adTypeText = 2
StreamTypeEnum = c_int # enum

# values for enumeration 'LineSeparatorEnum'
adLF = 10
adCR = 13
adCRLF = -1
LineSeparatorEnum = c_int # enum

# values for enumeration 'StreamOpenOptionsEnum'
adOpenStreamUnspecified = -1
adOpenStreamAsync = 1
adOpenStreamFromRecord = 4
StreamOpenOptionsEnum = c_int # enum

# values for enumeration 'SaveOptionsEnum'
adSaveCreateNotExist = 1
adSaveCreateOverWrite = 2
SaveOptionsEnum = c_int # enum

# values for enumeration 'StreamWriteEnum'
adWriteChar = 0
adWriteLine = 1
stWriteChar = 0
stWriteLine = 1
StreamWriteEnum = c_int # enum
_Stream._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Size',
              ( ['retval', 'out'], POINTER(c_int), 'pSize' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'EOS',
              ( ['retval', 'out'], POINTER(VARIANT_BOOL), 'pEOS' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'Position',
              ( ['retval', 'out'], POINTER(c_int), 'pPos' )),
    COMMETHOD([dispid(3), 'propput'], HRESULT, 'Position',
              ( ['in'], c_int, 'pPos' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'Type',
              ( ['retval', 'out'], POINTER(StreamTypeEnum), 'ptype' )),
    COMMETHOD([dispid(4), 'propput'], HRESULT, 'Type',
              ( ['in'], StreamTypeEnum, 'ptype' )),
    COMMETHOD([dispid(5), 'propget'], HRESULT, 'LineSeparator',
              ( ['retval', 'out'], POINTER(LineSeparatorEnum), 'pLS' )),
    COMMETHOD([dispid(5), 'propput'], HRESULT, 'LineSeparator',
              ( ['in'], LineSeparatorEnum, 'pLS' )),
    COMMETHOD([dispid(6), 'propget'], HRESULT, 'State',
              ( ['retval', 'out'], POINTER(ObjectStateEnum), 'pState' )),
    COMMETHOD([dispid(7), 'propget'], HRESULT, 'Mode',
              ( ['retval', 'out'], POINTER(ConnectModeEnum), 'pMode' )),
    COMMETHOD([dispid(7), 'propput'], HRESULT, 'Mode',
              ( ['in'], ConnectModeEnum, 'pMode' )),
    COMMETHOD([dispid(8), 'propget'], HRESULT, 'Charset',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstrCharset' )),
    COMMETHOD([dispid(8), 'propput'], HRESULT, 'Charset',
              ( ['in'], BSTR, 'pbstrCharset' )),
    COMMETHOD([dispid(9)], HRESULT, 'Read',
              ( ['in', 'optional'], c_int, 'NumBytes', -1 ),
              ( ['retval', 'out'], POINTER(VARIANT), 'pval' )),
    COMMETHOD([dispid(10)], HRESULT, 'Open',
              ( ['in', 'optional'], VARIANT, 'Source' ),
              ( ['in', 'optional'], ConnectModeEnum, 'Mode', 0 ),
              ( ['in', 'optional'], StreamOpenOptionsEnum, 'Options', -1 ),
              ( ['in', 'optional'], BSTR, 'UserName', '' ),
              ( ['in', 'optional'], BSTR, 'Password', '' )),
    COMMETHOD([dispid(11)], HRESULT, 'Close'),
    COMMETHOD([dispid(12)], HRESULT, 'SkipLine'),
    COMMETHOD([dispid(13)], HRESULT, 'Write',
              ( ['in'], VARIANT, 'Buffer' )),
    COMMETHOD([dispid(14)], HRESULT, 'SetEOS'),
    COMMETHOD([dispid(15)], HRESULT, 'CopyTo',
              ( ['in'], POINTER(_Stream), 'DestStream' ),
              ( ['in', 'optional'], c_int, 'CharNumber', -1 )),
    COMMETHOD([dispid(16)], HRESULT, 'Flush'),
    COMMETHOD([dispid(17)], HRESULT, 'SaveToFile',
              ( ['in'], BSTR, 'FileName' ),
              ( ['in', 'optional'], SaveOptionsEnum, 'Options', 1 )),
    COMMETHOD([dispid(18)], HRESULT, 'LoadFromFile',
              ( ['in'], BSTR, 'FileName' )),
    COMMETHOD([dispid(19)], HRESULT, 'ReadText',
              ( ['in', 'optional'], c_int, 'NumChars', -1 ),
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(20)], HRESULT, 'WriteText',
              ( ['in'], BSTR, 'Data' ),
              ( ['in', 'optional'], StreamWriteEnum, 'Options', 0 )),
    COMMETHOD([dispid(21)], HRESULT, 'Cancel'),
]
################################################################
## code template for _Stream implementation
##class _Stream_Impl(object):
##    def ReadText(self, NumChars):
##        '-no docstring-'
##        #return pbstr
##
##    def Close(self):
##        '-no docstring-'
##        #return 
##
##    @property
##    def State(self):
##        '-no docstring-'
##        #return pState
##
##    def Cancel(self):
##        '-no docstring-'
##        #return 
##
##    @property
##    def Size(self):
##        '-no docstring-'
##        #return pSize
##
##    def _get(self):
##        '-no docstring-'
##        #return ptype
##    def _set(self, ptype):
##        '-no docstring-'
##    Type = property(_get, _set, doc = _set.__doc__)
##
##    def Open(self, Source, Mode, Options, UserName, Password):
##        '-no docstring-'
##        #return 
##
##    def CopyTo(self, DestStream, CharNumber):
##        '-no docstring-'
##        #return 
##
##    def SaveToFile(self, FileName, Options):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pLS
##    def _set(self, pLS):
##        '-no docstring-'
##    LineSeparator = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pMode
##    def _set(self, pMode):
##        '-no docstring-'
##    Mode = property(_get, _set, doc = _set.__doc__)
##
##    def WriteText(self, Data, Options):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstrCharset
##    def _set(self, pbstrCharset):
##        '-no docstring-'
##    Charset = property(_get, _set, doc = _set.__doc__)
##
##    def LoadFromFile(self, FileName):
##        '-no docstring-'
##        #return 
##
##    def Flush(self):
##        '-no docstring-'
##        #return 
##
##    @property
##    def EOS(self):
##        '-no docstring-'
##        #return pEOS
##
##    def _get(self):
##        '-no docstring-'
##        #return pPos
##    def _set(self, pPos):
##        '-no docstring-'
##    Position = property(_get, _set, doc = _set.__doc__)
##
##    def Write(self, Buffer):
##        '-no docstring-'
##        #return 
##
##    def SetEOS(self):
##        '-no docstring-'
##        #return 
##
##    def Read(self, NumBytes):
##        '-no docstring-'
##        #return pval
##
##    def SkipLine(self):
##        '-no docstring-'
##        #return 
##

class ADOConnectionConstruction15(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{00000516-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['restricted']
class ADOConnectionConstruction(ADOConnectionConstruction15):
    _case_insensitive_ = True
    _iid_ = GUID('{00000551-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['restricted']
ADOConnectionConstruction15._methods_ = [
    COMMETHOD(['propget'], HRESULT, 'DSO',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppDSO' )),
    COMMETHOD(['propget'], HRESULT, 'Session',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppSession' )),
    COMMETHOD([], HRESULT, 'WrapDSOandSession',
              ( ['in'], POINTER(IUnknown), 'pDSO' ),
              ( ['in'], POINTER(IUnknown), 'pSession' )),
]
################################################################
## code template for ADOConnectionConstruction15 implementation
##class ADOConnectionConstruction15_Impl(object):
##    @property
##    def DSO(self):
##        '-no docstring-'
##        #return ppDSO
##
##    def WrapDSOandSession(self, pDSO, pSession):
##        '-no docstring-'
##        #return 
##
##    @property
##    def Session(self):
##        '-no docstring-'
##        #return ppSession
##

ADOConnectionConstruction._methods_ = [
]
################################################################
## code template for ADOConnectionConstruction implementation
##class ADOConnectionConstruction_Impl(object):

class Field15(_ADO):
    _case_insensitive_ = True
    _iid_ = GUID('{00000505-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['hidden', 'dual', 'nonextensible', 'oleautomation']
Field15._methods_ = [
    COMMETHOD([dispid(1109), 'propget'], HRESULT, 'ActualSize',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1114), 'propget'], HRESULT, 'Attributes',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1103), 'propget'], HRESULT, 'DefinedSize',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(1100), 'propget'], HRESULT, 'Name',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(1102), 'propget'], HRESULT, 'Type',
              ( ['retval', 'out'], POINTER(DataTypeEnum), 'pDataType' )),
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Value',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(0), 'propput'], HRESULT, 'Value',
              ( ['in'], VARIANT, 'pvar' )),
    COMMETHOD([dispid(1112), 'propget'], HRESULT, 'Precision',
              ( ['retval', 'out'], POINTER(c_ubyte), 'pbPrecision' )),
    COMMETHOD([dispid(1113), 'propget'], HRESULT, 'NumericScale',
              ( ['retval', 'out'], POINTER(c_ubyte), 'pbNumericScale' )),
    COMMETHOD([dispid(1107)], HRESULT, 'AppendChunk',
              ( ['in'], VARIANT, 'Data' )),
    COMMETHOD([dispid(1108)], HRESULT, 'GetChunk',
              ( ['in'], c_int, 'Length' ),
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(1104), 'propget'], HRESULT, 'OriginalValue',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(1105), 'propget'], HRESULT, 'UnderlyingValue',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
]
################################################################
## code template for Field15 implementation
##class Field15_Impl(object):
##    @property
##    def Attributes(self):
##        '-no docstring-'
##        #return pl
##
##    @property
##    def Type(self):
##        '-no docstring-'
##        #return pDataType
##
##    @property
##    def OriginalValue(self):
##        '-no docstring-'
##        #return pvar
##
##    @property
##    def DefinedSize(self):
##        '-no docstring-'
##        #return pl
##
##    def AppendChunk(self, Data):
##        '-no docstring-'
##        #return 
##
##    @property
##    def NumericScale(self):
##        '-no docstring-'
##        #return pbNumericScale
##
##    @property
##    def UnderlyingValue(self):
##        '-no docstring-'
##        #return pvar
##
##    @property
##    def ActualSize(self):
##        '-no docstring-'
##        #return pl
##
##    def GetChunk(self, Length):
##        '-no docstring-'
##        #return pvar
##
##    @property
##    def Precision(self):
##        '-no docstring-'
##        #return pbPrecision
##
##    @property
##    def Name(self):
##        '-no docstring-'
##        #return pbstr
##
##    def _get(self):
##        '-no docstring-'
##        #return pvar
##    def _set(self, pvar):
##        '-no docstring-'
##    Value = property(_get, _set, doc = _set.__doc__)
##


# values for enumeration 'ADCPROP_UPDATERESYNC_ENUM'
adResyncNone = 0
adResyncAutoIncrement = 1
adResyncConflicts = 2
adResyncUpdates = 4
adResyncInserts = 8
adResyncAll = 15
ADCPROP_UPDATERESYNC_ENUM = c_int # enum

# values for enumeration 'FieldStatusEnum'
adFieldOK = 0
adFieldCantConvertValue = 2
adFieldIsNull = 3
adFieldTruncated = 4
adFieldSignMismatch = 5
adFieldDataOverflow = 6
adFieldCantCreate = 7
adFieldUnavailable = 8
adFieldPermissionDenied = 9
adFieldIntegrityViolation = 10
adFieldSchemaViolation = 11
adFieldBadStatus = 12
adFieldDefault = 13
adFieldIgnore = 15
adFieldDoesNotExist = 16
adFieldInvalidURL = 17
adFieldResourceLocked = 18
adFieldResourceExists = 19
adFieldCannotComplete = 20
adFieldVolumeNotFound = 21
adFieldOutOfSpace = 22
adFieldCannotDeleteSource = 23
adFieldReadOnly = 24
adFieldResourceOutOfScope = 25
adFieldAlreadyExists = 26
adFieldPendingInsert = 65536
adFieldPendingDelete = 131072
adFieldPendingChange = 262144
adFieldPendingUnknown = 524288
adFieldPendingUnknownDelete = 1048576
FieldStatusEnum = c_int # enum

# values for enumeration 'ADCPROP_ASYNCTHREADPRIORITY_ENUM'
adPriorityLowest = 1
adPriorityBelowNormal = 2
adPriorityNormal = 3
adPriorityAboveNormal = 4
adPriorityHighest = 5
ADCPROP_ASYNCTHREADPRIORITY_ENUM = c_int # enum

# values for enumeration 'CommandTypeEnum'
adCmdUnspecified = -1
adCmdUnknown = 8
adCmdText = 1
adCmdTable = 2
adCmdStoredProc = 4
adCmdFile = 256
adCmdTableDirect = 512
CommandTypeEnum = c_int # enum

# values for enumeration 'ExecuteOptionEnum'
adOptionUnspecified = -1
adAsyncExecute = 16
adAsyncFetch = 32
adAsyncFetchNonBlocking = 64
adExecuteNoRecords = 128
ExecuteOptionEnum = c_int # enum

# values for enumeration 'ParameterDirectionEnum'
adParamUnknown = 0
adParamInput = 1
adParamOutput = 2
adParamInputOutput = 3
adParamReturnValue = 4
ParameterDirectionEnum = c_int # enum
class _Parameter(_ADO):
    _case_insensitive_ = True
    _iid_ = GUID('{0000050C-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class Parameters(_DynaCollection):
    _case_insensitive_ = True
    _iid_ = GUID('{0000050D-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
Command15._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ActiveConnection',
              ( ['retval', 'out'], POINTER(POINTER(_Connection)), 'ppvObject' )),
    COMMETHOD([dispid(1), 'propputref'], HRESULT, 'ActiveConnection',
              ( ['in'], POINTER(_Connection), 'ppvObject' )),
    COMMETHOD([dispid(1), 'propput'], HRESULT, 'ActiveConnection',
              ( ['in'], VARIANT, 'ppvObject' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'CommandText',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(2), 'propput'], HRESULT, 'CommandText',
              ( ['in'], BSTR, 'pbstr' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'CommandTimeout',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(3), 'propput'], HRESULT, 'CommandTimeout',
              ( ['in'], c_int, 'pl' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'Prepared',
              ( ['retval', 'out'], POINTER(VARIANT_BOOL), 'pfPrepared' )),
    COMMETHOD([dispid(4), 'propput'], HRESULT, 'Prepared',
              ( ['in'], VARIANT_BOOL, 'pfPrepared' )),
    COMMETHOD([dispid(5)], HRESULT, 'Execute',
              ( ['out', 'optional'], POINTER(VARIANT), 'RecordsAffected' ),
              ( ['in', 'optional'], POINTER(VARIANT), 'Parameters' ),
              ( ['in', 'optional'], c_int, 'Options', -1 ),
              ( ['retval', 'out'], POINTER(POINTER(_Recordset)), 'ppiRs' )),
    COMMETHOD([dispid(6)], HRESULT, 'CreateParameter',
              ( ['in', 'optional'], BSTR, 'Name', '' ),
              ( ['in', 'optional'], DataTypeEnum, 'Type', 0 ),
              ( ['in', 'optional'], ParameterDirectionEnum, 'Direction', 1 ),
              ( ['in', 'optional'], c_int, 'Size', 0 ),
              ( ['in', 'optional'], VARIANT, 'Value' ),
              ( ['retval', 'out'], POINTER(POINTER(_Parameter)), 'ppiprm' )),
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Parameters',
              ( ['retval', 'out'], POINTER(POINTER(Parameters)), 'ppvObject' )),
    COMMETHOD([dispid(7), 'propput'], HRESULT, 'CommandType',
              ( ['in'], CommandTypeEnum, 'plCmdType' )),
    COMMETHOD([dispid(7), 'propget'], HRESULT, 'CommandType',
              ( ['retval', 'out'], POINTER(CommandTypeEnum), 'plCmdType' )),
    COMMETHOD([dispid(8), 'propget'], HRESULT, 'Name',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstrName' )),
    COMMETHOD([dispid(8), 'propput'], HRESULT, 'Name',
              ( ['in'], BSTR, 'pbstrName' )),
]
################################################################
## code template for Command15 implementation
##class Command15_Impl(object):
##    def _set(self, ppvObject):
##        '-no docstring-'
##    ActiveConnection = property(fset = _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return plCmdType
##    def _set(self, plCmdType):
##        '-no docstring-'
##    CommandType = property(_get, _set, doc = _set.__doc__)
##
##    def CreateParameter(self, Name, Type, Direction, Size, Value):
##        '-no docstring-'
##        #return ppiprm
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstr
##    def _set(self, pbstr):
##        '-no docstring-'
##    CommandText = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def Parameters(self):
##        '-no docstring-'
##        #return ppvObject
##
##    def Execute(self, Parameters, Options):
##        '-no docstring-'
##        #return RecordsAffected, ppiRs
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstrName
##    def _set(self, pbstrName):
##        '-no docstring-'
##    Name = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pl
##    def _set(self, pl):
##        '-no docstring-'
##    CommandTimeout = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pfPrepared
##    def _set(self, pfPrepared):
##        '-no docstring-'
##    Prepared = property(_get, _set, doc = _set.__doc__)
##

_Command._methods_ = [
    COMMETHOD([dispid(9), 'propget'], HRESULT, 'State',
              ( ['retval', 'out'], POINTER(c_int), 'plObjState' )),
    COMMETHOD([dispid(10)], HRESULT, 'Cancel'),
]
################################################################
## code template for _Command implementation
##class _Command_Impl(object):
##    @property
##    def State(self):
##        '-no docstring-'
##        #return plObjState
##
##    def Cancel(self):
##        '-no docstring-'
##        #return 
##


# values for enumeration 'ADCPROP_AUTORECALC_ENUM'
adRecalcUpFront = 0
adRecalcAlways = 1
ADCPROP_AUTORECALC_ENUM = c_int # enum
class Connection(CoClass):
    _reg_clsid_ = GUID('{00000514-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['licensed']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{00000205-0000-0010-8000-00AA006D2EA4}', 2, 5)
Connection._com_interfaces_ = [_Connection]
Connection._outgoing_interfaces_ = [ConnectionEvents]

class Parameter(CoClass):
    _reg_clsid_ = GUID('{0000050B-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['licensed']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{00000205-0000-0010-8000-00AA006D2EA4}', 2, 5)
Parameter._com_interfaces_ = [_Parameter]


# values for enumeration 'StreamReadEnum'
adReadAll = -1
adReadLine = -2
StreamReadEnum = c_int # enum
Parameters._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Item',
              ( ['in'], VARIANT, 'Index' ),
              ( ['retval', 'out'], POINTER(POINTER(_Parameter)), 'ppvObject' )),
]
################################################################
## code template for Parameters implementation
##class Parameters_Impl(object):
##    @property
##    def Item(self, Index):
##        '-no docstring-'
##        #return ppvObject
##

Fields._methods_ = [
    COMMETHOD([dispid(3)], HRESULT, 'Append',
              ( ['in'], BSTR, 'Name' ),
              ( ['in'], DataTypeEnum, 'Type' ),
              ( ['in', 'optional'], c_int, 'DefinedSize', 0 ),
              ( ['in', 'optional'], FieldAttributeEnum, 'Attrib', -1 ),
              ( ['in', 'optional'], VARIANT, 'FieldValue' )),
    COMMETHOD([dispid(5)], HRESULT, 'Update'),
    COMMETHOD([dispid(6)], HRESULT, 'Resync',
              ( ['in', 'optional'], ResyncEnum, 'ResyncValues', 2 )),
    COMMETHOD([dispid(7)], HRESULT, 'CancelUpdate'),
]
################################################################
## code template for Fields implementation
##class Fields_Impl(object):
##    def Append(self, Name, Type, DefinedSize, Attrib, FieldValue):
##        '-no docstring-'
##        #return 
##
##    def Resync(self, ResyncValues):
##        '-no docstring-'
##        #return 
##
##    def CancelUpdate(self):
##        '-no docstring-'
##        #return 
##
##    def Update(self):
##        '-no docstring-'
##        #return 
##


# values for enumeration 'FieldEnum'
adDefaultStream = -1
adRecordURL = -2
FieldEnum = c_int # enum
_Recordset._methods_ = [
    COMMETHOD([dispid(1057)], HRESULT, 'Save',
              ( ['in', 'optional'], VARIANT, 'Destination' ),
              ( ['in', 'optional'], PersistFormatEnum, 'PersistFormat', 0 )),
]
################################################################
## code template for _Recordset implementation
##class _Recordset_Impl(object):
##    def Save(self, Destination, PersistFormat):
##        '-no docstring-'
##        #return 
##

class ADORecordConstruction(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000567-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['restricted']
ADORecordConstruction._methods_ = [
    COMMETHOD(['propget'], HRESULT, 'Row',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppRow' )),
    COMMETHOD(['propput'], HRESULT, 'Row',
              ( ['in'], POINTER(IUnknown), 'ppRow' )),
    COMMETHOD(['propput'], HRESULT, 'ParentRow',
              ( ['in'], POINTER(IUnknown), 'rhs' )),
]
################################################################
## code template for ADORecordConstruction implementation
##class ADORecordConstruction_Impl(object):
##    def _set(self, rhs):
##        '-no docstring-'
##    ParentRow = property(fset = _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return ppRow
##    def _set(self, ppRow):
##        '-no docstring-'
##    Row = property(_get, _set, doc = _set.__doc__)
##

_Parameter._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Name',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(1), 'propput'], HRESULT, 'Name',
              ( ['in'], BSTR, 'pbstr' )),
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Value',
              ( ['retval', 'out'], POINTER(VARIANT), 'pvar' )),
    COMMETHOD([dispid(0), 'propput'], HRESULT, 'Value',
              ( ['in'], VARIANT, 'pvar' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'Type',
              ( ['retval', 'out'], POINTER(DataTypeEnum), 'psDataType' )),
    COMMETHOD([dispid(2), 'propput'], HRESULT, 'Type',
              ( ['in'], DataTypeEnum, 'psDataType' )),
    COMMETHOD([dispid(3), 'propput'], HRESULT, 'Direction',
              ( ['in'], ParameterDirectionEnum, 'plParmDirection' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'Direction',
              ( ['retval', 'out'], POINTER(ParameterDirectionEnum), 'plParmDirection' )),
    COMMETHOD([dispid(4), 'propput'], HRESULT, 'Precision',
              ( ['in'], c_ubyte, 'pbPrecision' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'Precision',
              ( ['retval', 'out'], POINTER(c_ubyte), 'pbPrecision' )),
    COMMETHOD([dispid(5), 'propput'], HRESULT, 'NumericScale',
              ( ['in'], c_ubyte, 'pbScale' )),
    COMMETHOD([dispid(5), 'propget'], HRESULT, 'NumericScale',
              ( ['retval', 'out'], POINTER(c_ubyte), 'pbScale' )),
    COMMETHOD([dispid(6), 'propput'], HRESULT, 'Size',
              ( ['in'], c_int, 'pl' )),
    COMMETHOD([dispid(6), 'propget'], HRESULT, 'Size',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(7)], HRESULT, 'AppendChunk',
              ( ['in'], VARIANT, 'Val' )),
    COMMETHOD([dispid(8), 'propget'], HRESULT, 'Attributes',
              ( ['retval', 'out'], POINTER(c_int), 'plParmAttribs' )),
    COMMETHOD([dispid(8), 'propput'], HRESULT, 'Attributes',
              ( ['in'], c_int, 'plParmAttribs' )),
]
################################################################
## code template for _Parameter implementation
##class _Parameter_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return plParmAttribs
##    def _set(self, plParmAttribs):
##        '-no docstring-'
##    Attributes = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pl
##    def _set(self, pl):
##        '-no docstring-'
##    Size = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return psDataType
##    def _set(self, psDataType):
##        '-no docstring-'
##    Type = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pbPrecision
##    def _set(self, pbPrecision):
##        '-no docstring-'
##    Precision = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return plParmDirection
##    def _set(self, plParmDirection):
##        '-no docstring-'
##    Direction = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstr
##    def _set(self, pbstr):
##        '-no docstring-'
##    Name = property(_get, _set, doc = _set.__doc__)
##
##    def AppendChunk(self, Val):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pvar
##    def _set(self, pvar):
##        '-no docstring-'
##    Value = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pbScale
##    def _set(self, pbScale):
##        '-no docstring-'
##    NumericScale = property(_get, _set, doc = _set.__doc__)
##

RecordsetEvents._disp_methods_ = [
    DISPMETHOD([dispid(9)], HRESULT, 'WillChangeField',
               ( ['in'], c_int, 'cFields' ),
               ( ['in'], VARIANT, 'Fields' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(10)], HRESULT, 'FieldChangeComplete',
               ( ['in'], c_int, 'cFields' ),
               ( ['in'], VARIANT, 'Fields' ),
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(11)], HRESULT, 'WillChangeRecord',
               ( ['in'], EventReasonEnum, 'adReason' ),
               ( ['in'], c_int, 'cRecords' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(12)], HRESULT, 'RecordChangeComplete',
               ( ['in'], EventReasonEnum, 'adReason' ),
               ( ['in'], c_int, 'cRecords' ),
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(13)], HRESULT, 'WillChangeRecordset',
               ( ['in'], EventReasonEnum, 'adReason' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(14)], HRESULT, 'RecordsetChangeComplete',
               ( ['in'], EventReasonEnum, 'adReason' ),
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(15)], HRESULT, 'WillMove',
               ( ['in'], EventReasonEnum, 'adReason' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(16)], HRESULT, 'MoveComplete',
               ( ['in'], EventReasonEnum, 'adReason' ),
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(17)], HRESULT, 'EndOfRecordset',
               ( ['in', 'out'], POINTER(VARIANT_BOOL), 'fMoreData' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(18)], HRESULT, 'FetchProgress',
               ( ['in'], c_int, 'Progress' ),
               ( ['in'], c_int, 'MaxProgress' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
    DISPMETHOD([dispid(19)], HRESULT, 'FetchComplete',
               ( ['in'], POINTER(Error), 'pError' ),
               ( ['in', 'out'], POINTER(EventStatusEnum), 'adStatus' ),
               ( ['in'], POINTER(_Recordset), 'pRecordset' )),
]
Errors._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Item',
              ( ['in'], VARIANT, 'Index' ),
              ( ['retval', 'out'], POINTER(POINTER(Error)), 'ppvObject' )),
    COMMETHOD([dispid(1610809345)], HRESULT, 'Clear'),
]
################################################################
## code template for Errors implementation
##class Errors_Impl(object):
##    @property
##    def Item(self, Index):
##        '-no docstring-'
##        #return ppvObject
##
##    def Clear(self):
##        '-no docstring-'
##        #return 
##

class Property(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{00000503-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
Properties._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Item',
              ( ['in'], VARIANT, 'Index' ),
              ( ['retval', 'out'], POINTER(POINTER(Property)), 'ppvObject' )),
]
################################################################
## code template for Properties implementation
##class Properties_Impl(object):
##    @property
##    def Item(self, Index):
##        '-no docstring-'
##        #return ppvObject
##

Error._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Number',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'Source',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Description',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'HelpFile',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'HelpContext',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
    COMMETHOD([dispid(5), 'propget'], HRESULT, 'SQLState',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(6), 'propget'], HRESULT, 'NativeError',
              ( ['retval', 'out'], POINTER(c_int), 'pl' )),
]
################################################################
## code template for Error implementation
##class Error_Impl(object):
##    @property
##    def HelpContext(self):
##        '-no docstring-'
##        #return pl
##
##    @property
##    def Number(self):
##        '-no docstring-'
##        #return pl
##
##    @property
##    def Source(self):
##        '-no docstring-'
##        #return pbstr
##
##    @property
##    def HelpFile(self):
##        '-no docstring-'
##        #return pbstr
##
##    @property
##    def NativeError(self):
##        '-no docstring-'
##        #return pl
##
##    @property
##    def Description(self):
##        '-no docstring-'
##        #return pbstr
##
##    @property
##    def SQLState(self):
##        '-no docstring-'
##        #return pbstr
##

Property._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Value',
              ( ['retval', 'out'], POINTER(VARIANT), 'pval' )),
    COMMETHOD([dispid(0), 'propput'], HRESULT, 'Value',
              ( ['in'], VARIANT, 'pval' )),
    COMMETHOD([dispid(1610743810), 'propget'], HRESULT, 'Name',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(1610743811), 'propget'], HRESULT, 'Type',
              ( ['retval', 'out'], POINTER(DataTypeEnum), 'ptype' )),
    COMMETHOD([dispid(1610743812), 'propget'], HRESULT, 'Attributes',
              ( ['retval', 'out'], POINTER(c_int), 'plAttributes' )),
    COMMETHOD([dispid(1610743812), 'propput'], HRESULT, 'Attributes',
              ( ['in'], c_int, 'plAttributes' )),
]
################################################################
## code template for Property implementation
##class Property_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return plAttributes
##    def _set(self, plAttributes):
##        '-no docstring-'
##    Attributes = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def Name(self):
##        '-no docstring-'
##        #return pbstr
##
##    @property
##    def Type(self):
##        '-no docstring-'
##        #return ptype
##
##    def _get(self):
##        '-no docstring-'
##        #return pval
##    def _set(self, pval):
##        '-no docstring-'
##    Value = property(_get, _set, doc = _set.__doc__)
##

class Command(CoClass):
    _reg_clsid_ = GUID('{00000507-0000-0010-8000-00AA006D2EA4}')
    _idlflags_ = ['licensed']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{00000205-0000-0010-8000-00AA006D2EA4}', 2, 5)
Command._com_interfaces_ = [_Command]

__all__ = [ 'adPropNotSupported', 'adRecNew', 'adFieldResourceExists',
           'adFieldPendingInsert', 'adFldIsNullable', 'adWChar',
           'adSchemaCatalogs', 'adErrObjectClosed', 'adBoolean',
           'adFieldDefault', 'Errors', 'adRsnRequery', 'Recordset',
           'adFldUnknownUpdatable', 'adRsnMove', '_Connection',
           'adRecMultipleChanges', 'adMarshalModifiedOnly', 'adSeek',
           'adOpenDynamic', 'adPosEOF', 'adRsnMoveFirst',
           'adErrInvalidTransaction', 'adParamOutput',
           'adLongVarWChar', 'adFldNegativeScale',
           'adErrObjectNotSet', 'adResyncConflicts',
           'adCompareLessThan', 'adRecPermissionDenied',
           'stWriteLine', 'Stream', 'CopyRecordOptionsEnum',
           'adSchemaViews', 'adPriorityBelowNormal',
           'adSchemaSQLLanguages', 'adUnsignedTinyInt', 'adRecordURL',
           'adErrPropInvalidOption', 'adAffectGroup',
           'adErrSignMismatch', '_Stream', 'adRecObjectOpen',
           'adRsnFirstChange', 'adCurrency',
           'adSchemaReferentialContraints', 'adCmdUnspecified',
           'StringFormatEnum', 'EventStatusEnum', 'adBigInt',
           'adSchemaProperties', 'adChapter', 'adPropOptional',
           'adPropWrite', 'adParamSigned', 'adAddNew', 'adEditNone',
           'adFldCacheDeferred', 'adFldUpdatable',
           'adErrStillExecuting', 'adErrDenyTypeNotSupported',
           'adVariant', 'adRecalcUpFront', 'adErrInvalidURL',
           'adSchemaProcedures', 'adPromptAlways', 'adInteger',
           'ADOConnectionConstruction', 'adXactAbortRetaining',
           'ObjectStateEnum', 'adXactCursorStability',
           'adStateExecuting', 'adAsyncFetch', 'adPropVariant',
           'Record', 'adCompareNotEqual', 'adErrCantChangeProvider',
           'adCompareGreaterThan', 'adXactUnspecified',
           'adErrUnavailable', 'adDate', 'adSchemaDBInfoLiterals',
           'StreamWriteEnum', 'adPromptNever',
           'adRecIntegrityViolation', 'ConnectOptionEnum',
           'adSchemaReferentialConstraints', 'PersistFormatEnum',
           'adCriteriaUpdCols', 'adFldMayDefer', 'SearchDirection',
           'adFldLong', 'adMovePrevious', 'adUseServer',
           'RecordsetEventsVt', 'adLongVarChar', 'StreamTypeEnum',
           'Parameters', 'adSeekBeforeEQ', 'adDelayFetchStream',
           'adUseNone', 'adSchemaKeyColumnUsage', 'StreamReadEnum',
           'adSchemaMembers', 'adXactSerializable', 'adUserDefined',
           'SaveOptionsEnum', 'adFldUnspecified',
           'adFieldUnavailable', 'ADCPROP_ASYNCTHREADPRIORITY_ENUM',
           'adErrResourceExists', 'adSchemaLevels',
           'adSchemaProcedureColumns', 'adOpenIfExists',
           'adParamReturnValue', 'adError', 'adSeekFirstEQ',
           'adXactAsyncPhaseOne', 'adSchemaStatistics', 'adReadLine',
           'ErrorValueEnum', 'ConnectPromptEnum',
           'adXactRepeatableRead', 'adOpenRecordUnspecified',
           'adOpenStreamAsync', 'adReadAll', 'adCriteriaTimeStamp',
           'adRsnMoveNext', 'adErrObjectOpen',
           'adXactReadUncommitted', 'adErrURLDoesNotExist',
           'adResyncUnderlyingValues', 'adErrTreePermissionDenied',
           '_Recordset', 'adWriteChar', 'adAffectCurrent',
           'adFieldPendingDelete', 'adCmdFile', 'adFldIsRowURL',
           'ADORecordConstruction', 'adErrNotExecuting',
           'CompareEnum', 'adSeekBefore', 'EventReasonEnum',
           'adFilterPredicate', 'adDelete', 'adTypeBinary',
           'adPosUnknown', 'adStatusCancel', 'adVarWChar',
           'adBookmarkFirst', 'adModeUnknown', 'adAffectAllChapters',
           'adUseClient', 'adRecDBDeleted', 'adCriteriaAllCols',
           'SchemaEnum', 'adSchemaMeasures', 'adEditDelete',
           'adSearchBackward', 'adFilterNone', 'adStateConnecting',
           'adErrPropInvalidColumn', 'adXactReadCommitted',
           'adSchemaDBInfoKeywords', 'adFieldDoesNotExist',
           'adRsnMovePrevious', '_Record', 'adRsnUndoAddNew',
           'FieldAttributeEnum', 'adFldRowVersion', 'Field',
           '_Command', 'adMoveAllowEmulation', 'RecordStatusEnum',
           'adXactSyncPhaseOne', 'adFieldIsNull',
           'ParameterAttributesEnum', 'adBookmark', 'adPersistXML',
           'EditModeEnum', 'adCmdUnknown', 'adRecPendingChanges',
           'adDelayFetchFields', 'adSchemaColumns', 'adAffectAll',
           'adXactBrowse', 'adRecalcAlways',
           'ADORecordsetConstruction', 'adSchemaPrimaryKeys',
           'adFieldCannotComplete', 'adSchemaProcedureParameters',
           'adErrInvalidConnection', 'adParamNullable', 'Properties',
           '_Parameter', 'adDBDate', 'adOpenKeyset', 'adCompareEqual',
           'adCopyAllowEmulation', 'adSchemaSchemata',
           'adFieldResourceOutOfScope', 'adClipString',
           'adMarshalAll', 'adBSTR', 'adFieldIgnore', 'adVarChar',
           'adErrFieldsUpdateFailed', 'RecordTypeEnum', 'adDBTime',
           'adFilterAffectedRecords', 'adRsnUndoUpdate', 'adRsnClose',
           'adRecSchemaViolation', 'adResyncAllValues',
           'adMoveDontUpdateLinks', 'adErrUnsafeOperation',
           'adErrPropNotAllSettable', 'adModeShareDenyWrite',
           'adParamInputOutput', 'adSchemaViewTableUsage',
           'adModeRecursive', 'RecordCreateOptionsEnum',
           'adRecCanceled', 'adErrURLNamedRowDoesNotExist',
           'adModeShareDenyNone', 'adResyncUpdates',
           'adMoveUnspecified', 'adFldIsChapter',
           'adSchemaProviderSpecific', 'IRecFields',
           'adFieldCannotDeleteSource', 'adSchemaTableConstraints',
           'adFieldPendingUnknown', 'adSeekLastEQ',
           'adSchemaDimensions', 'ADCPROP_UPDATECRITERIA_ENUM',
           'adSchemaForeignKeys', 'adSchemaConstraintTableUsage',
           'adRsnMoveLast', 'adMoveOverWrite', 'adErrSchemaViolation',
           'adEditInProgress', 'adOpenStatic', 'adLockReadOnly',
           'ADCPROP_UPDATERESYNC_ENUM', 'adBookmarkCurrent',
           'Parameter', 'adSchemaCubes', 'adSchemaTranslations',
           'adSchemaHierarchies', 'adErrCannotComplete', 'adRecOK',
           'adCR', 'adEditAdd', '_ADO', 'adFieldCantConvertValue',
           'adOpenAsync', 'adUpdateBatch', 'adFind',
           'adFieldSignMismatch', 'ConnectionEventsVt',
           'IsolationLevelEnum', 'MarshalOptionsEnum',
           'adCopyUnspecified', 'adLF', 'adSchemaAsserts',
           'adSchemaTables', 'adFieldAlreadyExists', 'adStatusOK',
           'Field20', 'adFieldIntegrityViolation',
           'ExecuteOptionEnum', 'RecordOpenOptionsEnum',
           'adOpenStreamUnspecified', 'Fields',
           'adErrPropNotSupported', 'adFilterFetchedRecords',
           'adPriorityNormal', 'adEmpty', 'adFieldTruncated',
           'CursorTypeEnum', 'adIndex', 'adSimpleRecord',
           'XactAttributeEnum', 'PropertyAttributesEnum',
           'adFieldBadStatus', 'adParamUnknown',
           'adErrVolumeNotFound', 'adModeShareExclusive',
           'adConnectUnspecified', 'adModeShareDenyRead',
           'adRecCantRelease', 'adFldIsCollection',
           'adUseClientBatch', 'adRecMaxChangesExceeded',
           'adGetRowsRest', 'adNumeric', 'Connection',
           'adFldIsDefaultStream', 'ConnectModeEnum',
           'StreamOpenOptionsEnum', 'ConnectionEvents',
           'adFldMayBeNull', 'adFailIfNotExists',
           'adFieldPermissionDenied', 'adErrCantConvertvalue',
           'adStateFetching', 'Fields15', 'adRecConcurrencyViolation',
           'adErrPermissionDenied', 'adSeekAfter', 'CommandTypeEnum',
           'adUpdate', 'LockTypeEnum', 'adErrPropNotSettable',
           'adVarNumeric', 'adCmdStoredProc', 'adFieldPendingChange',
           'Command', 'adFieldResourceLocked',
           'adPromptCompleteRequired', 'GetRowsOptionEnum',
           'adRsnUndoDelete', 'adSmallInt', 'adErrDenyNotSupported',
           'adSchemaCharacterSets', 'adTinyInt',
           'adSaveCreateOverWrite', 'adSchemaColumnsDomainUsage',
           'adwrnSecurityDialogHeader', 'adStructDoc',
           'adPromptComplete', 'Connection15', 'adCollectionRecord',
           'ADOConnectionConstruction15', 'adErrOutOfSpace',
           'adAsyncExecute', 'Command15', '_Collection',
           'adStatusUnwantedEvent', 'adRsnResynch', 'adFieldReadOnly',
           'adErrCantChangeConnection', 'adCmdTableDirect',
           'adPropRead', 'adSchemaCollations',
           'adCreateNonCollection', 'adErrWriteFile', 'adArray',
           'adErrInTransaction', 'adResyncAll',
           'adAsyncFetchNonBlocking', 'Recordset20', 'Recordset21',
           'CursorLocationEnum', 'adOpenStreamFromRecord',
           'adErrPropNotSet', 'adFilterConflictingRecords',
           'adSchemaConstraintColumnUsage', 'adResync',
           'adSearchForward', 'adErrInvalidArgument', 'adCriteriaKey',
           'adSchemaUsagePrivileges', 'adSaveCreateNotExist',
           'adLockUnspecified', 'SearchDirectionEnum',
           'adResyncAutoIncrement', 'FilterGroupEnum',
           'adPriorityLowest', 'adHoldRecords', 'PositionEnum',
           'adRecModified', 'adModeWrite', 'adErrDelResOutOfScope',
           'adModeReadWrite', 'FieldEnum', 'adChar', 'BookmarkEnum',
           'adErrResourceLocked', 'adLongVarBinary',
           'adFieldPendingUnknownDelete', 'adFieldOK',
           'adErrNoCurrentRecord', 'adStateOpen',
           'adLockBatchOptimistic', 'adStatusErrorsOccurred',
           'adApproxPosition', 'adSchemaTablePrivileges',
           'adLockOptimistic', 'adParamInput',
           'adErrFeatureNotAvailable', 'adTypeText',
           'ADOCommandConstruction', 'adUnsignedInt', 'adSingle',
           'adRecInvalid', 'ParameterDirectionEnum',
           'adErrIllegalOperation', 'adCmdTable', 'LineSeparatorEnum',
           'adRsnUpdate', 'adIUnknown', 'adErrProviderNotFound',
           'adXactChaos', 'CursorOptionEnum', 'adOpenUnspecified',
           'adBookmarkLast', 'RecordsetEvents', 'adRsnAddNew',
           'adErrCantCreate', 'adFldFixed', 'adErrOpeningFile',
           'adNotify', 'adErrPropConflicting', 'adDecimal',
           'adErrDataOverflow', 'adErrDataConversion', 'Fields20',
           'adDBTimeStamp', 'adSchemaProviderTypes', 'adDouble',
           'Error', 'adFieldOutOfSpace', 'ResyncEnum', 'adCRLF',
           'ADOStreamConstruction', 'adBinary', 'adExecuteNoRecords',
           'adFieldSchemaViolation', 'adErrNotReentrant', 'adGUID',
           'adFldKeyColumn', '_DynaCollection', 'AffectEnum',
           'adPosBOF', 'Field15', 'adCreateOverwrite', 'adWriteLine',
           'adUnsignedSmallInt', 'adErrObjectInCollection',
           'adFilterPendingRecords', 'FieldStatusEnum',
           'adRecOutOfMemory', 'adFieldInvalidURL',
           'adCopyNonRecursive', 'adErrPropInvalidValue',
           'adDefaultStream', 'adErrInvalidParamInfo',
           'adStateClosed', 'DataTypeEnum', 'ADCPROP_AUTORECALC_ENUM',
           'adSchemaColumnPrivileges', 'Recordset15',
           'adPropRequired', 'adModeRead', 'adRecUnmodified',
           'adXactCommitRetaining', 'SeekEnum', 'adOpenForwardOnly',
           'adwrnSecurityDialog', 'adXactIsolated', 'adParamLong',
           'adErrStillConnecting', 'adRecDeleted', 'adSchemaTrustees',
           'adFileTime', 'adPersistADTG', 'adFldRowID',
           'adErrCatalogNotSet', 'adPriorityAboveNormal',
           'adResyncInserts', 'Property', 'adStatusCantDeny',
           'adErrIntegrityViolation', 'adOptionUnspecified',
           'adCmdText', 'adSchemaIndexes', 'adErrProviderFailed',
           'adSchemaCheckConstraints', 'adIDispatch',
           'adCreateStructDoc', 'adResyncNone', 'adPriorityHighest',
           'stWriteChar', 'adSchemaViewColumnUsage',
           'adErrOperationCancelled', 'adSeekAfterEQ',
           'adErrItemNotFound', 'adCreateCollection',
           'adLockPessimistic', 'adErrBoundToCommand',
           'adFieldDataOverflow', 'adCompareNotComparable',
           'adCopyOverWrite', 'adRsnDelete', 'adFieldCantCreate',
           'MoveRecordOptionsEnum', 'adVarBinary', 'adAsyncConnect',
           'adErrReadFile', 'adErrResourceOutOfScope',
           'adUnsignedBigInt', 'adErrColumnNotOnThisRow',
           'adOpenSource', 'adFieldVolumeNotFound']
from comtypes import _check_version; _check_version('')
