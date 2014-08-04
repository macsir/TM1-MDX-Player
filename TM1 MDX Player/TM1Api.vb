Module TM1Api
    'All documented functions
    'Calls first, Properties second and Error Strings last

    'Globals***************************************************************
    Public hUser As Long
    Public pGeneral As Long
    Public voDatabase As Long

    'Blob Functions**********************************************************
    Declare Function TM1BlobClose Lib "tm1api.dll" (ByVal hPool As Long, ByVal hBlob As Long) As Long
    Declare Function TM1BlobCreate Lib "tm1api.dll" (ByVal hPool As Long, ByVal hServer As Long, ByVal sName As Long) As Long
    Declare Function TM1BlobGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hBlob As Long, ByVal x As Long, ByVal n As Long, ByVal buf As String) As Long
    Declare Function TM1BlobOpen Lib "tm1api.dll" (ByVal hPool As Long, ByVal hBlob As Long) As Long
    Declare Function TM1BlobPut Lib "tm1api.dll" (ByVal hPool As Long, ByVal hBlob As Long, ByVal x As Long, ByVal n As Long, ByVal buf As String) As Long

    'Chore functions********************************************************
    Declare Function TM1ChoreExecute Lib "tm1api.dll" (ByVal hPool As Long, ByVal hChore As Long) As Long
    'Client functions********************************************************
    Declare Function TM1ClientAdd Lib "tm1api.dll" (ByVal hPool As Long, ByVal hServer As Long, ByVal sClientName As Long) As Long
    Declare Function TM1ClientGroupAssign Lib "tm1api.dll" (ByVal hPool As Long, ByVal hClient As Long, ByVal hGroup As Long) As Long
    Declare Function TM1ClientGroupIsAssigned Lib "tm1api.dll" (ByVal hPool As Long, ByVal hClient As Long, ByVal hGroup As Long) As Long
    Declare Function TM1ClientGroupRemove Lib "tm1api.dll" (ByVal hPool As Long, ByVal hClient As Long, ByVal hGroup As Long) As Long
    Declare Function TM1ClientPasswordAssign Lib "tm1api.dll" (ByVal hPool As Long, ByVal hClient As Long, ByVal sPassword As Long) As Long

    'Cube functions**********************************************************
    Declare Function TM1CubeCellDrillStringGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hCube As Long, ByVal hArrayOfElements As Long) As Long
    Declare Function TM1CubeCellValueGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hCube As Long, ByVal hArrayOfElements As Long) As Long
    Declare Function TM1CubeCellValueSet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hCube As Long, ByVal hArrayOfElements As Long, ByVal hValue As Long) As Long
    Declare Function TM1CubeCreate Lib "tm1api.dll" (ByVal hPool As Long, ByVal hServer As Long, ByVal hArrayOfDimensions As Long) As Long
    Declare Function TM1CubePerspectiveCreate Lib "tm1api.dll" (ByVal hPool As Long, ByVal hCube As Long, ByVal hArrayOfElementTitles As Long) As Long
    Declare Function TM1CubePerspectiveDestroy Lib "tm1api.dll" (ByVal hPool As Long, ByVal hCube As Long, ByVal hArrayOfElementTitles As Long) As Long

    'Dimension functions*****************************************************
    Declare Function TM1DimensionCheck Lib "tm1api.dll" (ByVal hPool As Long, ByVal hDimension As Long) As Long
    Declare Function TM1DimensionCreateEmpty Lib "tm1api.dll" (ByVal hPool As Long, ByVal hServer As Long) As Long
    Declare Function TM1DimensionElementComponentAdd Lib "tm1api.dll" (ByVal hPool As Long, ByVal hElement As Long, ByVal hComponent As Long, ByVal rWeight As Long) As Long
    Declare Function TM1DimensionElementComponentDelete Lib "tm1api.dll" (ByVal hPool As Long, ByVal hCElement As Long, ByVal hElement As Long) As Long
    Declare Function TM1DimensionElementComponentWeightGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hCElement As Long, ByVal hElement As Long) As Long
    Declare Function TM1DimensionElementDelete Lib "tm1api.dll" (ByVal hPool As Long, ByVal hElement As Long) As Long
    Declare Function TM1DimensionElementInsert Lib "tm1api.dll" (ByVal hPool As Long, ByVal hDimension As Long, ByVal hElementBefore As Long, ByVal sName As Long, ByVal itype As Long) As Long
    Declare Function TM1DimensionUpdate Lib "tm1api.dll" (ByVal hPool As Long, ByVal hOldDimension As Long, ByVal hNewDimension As Long) As Long

    'Group function*********************************************************
    Declare Function TM1GroupAdd Lib "tm1api.dll" (ByVal hPool As Long, ByVal hServer As Long, ByVal sGroupName As Long) As Long

    'Object functions********************************************************
    'Object Attribute
    Declare Function TM1ObjectAttributeDelete Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal hAttribute As Long) As Long
    Declare Function TM1ObjectAttributeInsert Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal hAttributeBefore As Long, ByVal sName As Long, ByVal sType As Long) As Long
    Declare Function TM1ObjectAttributeValueGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal hAttribute As Long) As Long
    Declare Function TM1ObjectAttributeValueSet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal hAttribute As Long, ByVal hValue As Long) As Long
    'Object
    Declare Function TM1ObjectCopy Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSrcObject As Long, ByVal hDstObject As Long) As Long
    Declare Function TM1ObjectDelete Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long
    Declare Function TM1ObjectDestroy Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long
    Declare Function TM1ObjectDuplicate Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long
    'Object File
    Declare Function TM1ObjectFileDelete Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long
    Declare Function TM1ObjectFileLoad Lib "tm1api.dll" (ByVal hPool As Long, ByVal hServer As Long, ByVal hParent As Long, ByVal iObjectType As Long, ByVal sObjectName As Long) As Long
    Declare Function TM1ObjectFileSave Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long
    'Object List
    Declare Function TM1ObjectListCountGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal iPropertyList As Long) As Long
    Declare Function TM1ObjectListHandleByIndexGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal iPropertyList As Long, ByVal iIndex As Long) As Long
    Declare Function TM1ObjectListHandleByNameGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal iPropertyList As Long, ByVal sName As Long) As Long
    'Object Private
    Declare Function TM1ObjectPrivateDelete Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long
    Declare Function TM1ObjectPrivateListCountGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal iPropertyList As Long) As Long
    Declare Function TM1ObjectPrivateListHandleByIndexGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal iPropertyList As Long, ByVal iIndex As Long) As Long
    Declare Function TM1ObjectPrivateListHandleByNameGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal iPropertyList As Long, ByVal sName As Long) As Long
    Declare Function TM1ObjectPrivatePublish Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal sName As Long) As Long
    Declare Function TM1ObjectPrivateRegister Lib "tm1api.dll" (ByVal hPool As Long, ByVal hParent As Long, ByVal hObject As Long, ByVal sName As Long) As Long
    'Object Property
    Declare Function TM1ObjectPropertyGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal vProperty As Long) As Long
    Declare Function TM1ObjectPropertySet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal Property_P As Long, ByVal ValRec_V As Long) As Long
    Declare Function TM1ObjectRegister Lib "tm1api.dll" (ByVal hPool As Long, ByVal hParent As Long, ByVal hObject As Long, ByVal sName As Long) As Long
    'Object Security
    Declare Function TM1ObjectSecurityLock Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long
    Declare Function TM1ObjectSecurityRelease Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long
    Declare Function TM1ObjectSecurityReserve Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long
    Declare Function TM1ObjectSecurityRightGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal hGroup As Long) As Long
    Declare Function TM1ObjectSecurityRightSet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long, ByVal hGroup As Long, ByVal iRight As Long) As Long
    Declare Function TM1ObjectSecurityUnLock Lib "tm1api.dll" (ByVal hPool As Long, ByVal hObject As Long) As Long

    'Process functions**************************************************************
    Declare Function TM1ProcessExecute Lib "tm1api.dll" (ByVal hPool As Long, ByVal hProcess As Long, ByVal hParametersArray As Long) As Long
    Declare Function TM1ProcessExecuteEx Lib "tm1api.dll" (ByVal hPool As Long, ByVal hProcess As Long, ByVal hParametersArray As Long) As Long

    'Rule functions**************************************************************
    Declare Function TM1RuleAttach Lib "tm1api.dll" (ByVal hPool As Long, ByVal hRule As Long) As Long
    Declare Function TM1RuleCheck Lib "tm1api.dll" (ByVal hPool As Long, ByVal hRule As Long) As Long
    Declare Function TM1RuleCreateEmpty Lib "tm1api.dll" (ByVal hPool As Long, ByVal hCube As Long, ByVal hType As Long) As Long
    Declare Function TM1RuleDetach Lib "tm1api.dll" (ByVal hPool As Long, ByVal hRule As Long) As Long
    Declare Function TM1RuleLineGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hRule As Long, ByVal iPosition As Long) As Long
    Declare Function TM1RuleLineInsert Lib "tm1api.dll" (ByVal hPool As Long, ByVal hRule As Long, ByVal iPosition As Long, ByVal sLine As Long) As Long

    'Server functions*************************************************************
    Declare Function TM1ServerLogClose Lib "tm1api.dll" (ByVal hPool As Long, ByVal hLog As Long) As Long
    Declare Function TM1ServerLogNext Lib "tm1api.dll" (ByVal hPool As Long, ByVal hLog As Long) As Long
    Declare Function TM1ServerLogOpen Lib "tm1api.dll" (ByVal hPool As Long, ByVal hServer As Long, ByVal sStartTime As Long, ByVal sCubeFilter As Long, ByVal sUserFilter As Long, ByVal sFlagFilter As Long) As Long
    Declare Function TM1ServerPasswordChange Lib "tm1api.dll" (ByVal hPool As Long, ByVal hServer As Long, ByVal sNewPassword As Long) As Long

    'Subset functions************************************************************
    Declare Function TM1SubsetAll Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long) As Long
    Declare Function TM1SubsetCreateEmpty Lib "tm1api.dll" (ByVal hPool As Long, ByVal hDim As Long) As Long
    'Subset Element
    Declare Function TM1SubsetElementDisplay Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long, ByVal iElement As Long) As Long
    Declare Function TM1SubsetElementDisplayEll Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long) As Integer
    Declare Function TM1SubsetElementDisplayLevel Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long) As Long
    Declare Function TM1SubsetElementDisplayLine Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long, ByVal index As Long) As Integer
    Declare Function TM1SubsetElementDisplayMinus Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long) As Integer
    Declare Function TM1SubsetElementDisplayPlus Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long) As Integer
    Declare Function TM1SubsetElementDisplaySelection Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long) As Long
    Declare Function TM1SubsetElementDisplayTee Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long) As Integer
    Declare Function TM1SubsetElementDisplayWeight Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long) As Double
    'Subset Insert
    Declare Function TM1SubsetInsertElement Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long, ByVal hElement As Long, ByVal iPosition As Long) As Long
    Declare Function TM1SubsetInsertSubset Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubsetA As Long, ByVal hSubsetB As Long, ByVal iPosition As Long) As Long
    'Subset Select
    Declare Function TM1SubsetSelectByAttribute Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long, ByVal hAlias As Long, ByVal sValueToMatch As Long, ByVal bSelection As Long) As Long
    Declare Function TM1SubsetSelectByIndex Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long, ByVal iPosition As Long, ByVal bSelection As Long) As Long
    Declare Function TM1SubsetSelectByLevel Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long, ByVal iLevel As Long, ByVal bSelection As Long) As Long
    Declare Function TM1SubsetSelectByPattern Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long, ByVal sPattern As Long, ByVal hElement As Long) As Long
    Declare Function TM1SubsetSelectionDelete Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long) As Long
    Declare Function TM1SubsetSelectionInsertChildren Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long) As Long
    Declare Function TM1SubsetSelectionInsertParents Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long) As Long
    Declare Function TM1SubsetSelectionKeep Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long) As Long
    Declare Function TM1SubsetSelectNone Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long) As Long

    Declare Function TM1SubsetSort Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long, ByVal bSortDown As Long) As Long
    Declare Function TM1SubsetSortByIndex Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long, ByVal bSortDown As Long) As Long
    Declare Function TM1SubsetSortByHierarchy Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubset As Long) As Long
    Declare Function TM1SubsetSubtract Lib "tm1api.dll" (ByVal hPool As Long, ByVal hSubsetA As Long, ByVal hSubsetB As Long) As Long
    Declare Function TM1SubsetUpdate Lib "tm1api.dll" (ByVal hPool As Long, ByVal hOldSubset As Long, ByVal hNewSubset As Long) As Long

    'System functions***************************************************************
    Declare Sub TM1SystemAdminHostSet Lib "tm1api.dll" (ByVal hUser As Long, ByVal AdminHosts As String)
    Declare Sub TM1SystemClose Lib "tm1api.dll" (ByVal hUser As Long)
    Declare Function TM1SystemOpen Lib "tm1api.dll" () As Long
    Declare Sub TM1SystemProgressHookSet Lib "tm1api.dll" (ByVal hUser As Long, ByVal pHook As Long)
    Declare Sub TM1SystemServerClientName_VB Lib "tm1api.dll" (ByVal hUser As Long, ByVal index As Long, ByVal Res As String, ByVal max As Integer)
    Declare Function TM1SystemServerConnect Lib "tm1api.dll" (ByVal hPool As Long, ByVal sServer As Long, ByVal sClient As Long, ByVal sPassword As Long) As Long
    Declare Function TM1SystemServerDisconnect Lib "tm1api.dll" (ByVal hPool As Long, ByVal hServer As Long) As Long
    Declare Function TM1SystemServerHandle Lib "tm1api.dll" (ByVal hUser As Long, ByVal name As String) As Long
    Declare Sub TM1SystemServerName_VB Lib "tm1api.dll" (ByVal hUser As Long, ByVal index As Long, ByVal Res As String, ByVal max As Integer)
    Declare Function TM1SystemServerNof Lib "tm1api.dll" (ByVal hUser As Long) As Integer
    Declare Sub TM1SystemServerReload Lib "tm1api.dll" (ByVal hUser As Long)
    Declare Function TM1SystemServerStart Lib "tm1api.dll" (ByVal hUser As Long, ByVal szName As String, ByVal szDataDirectory As String, ByVal szAdminHost As String, ByVal szProtocol As String, ByVal iPortNumber As Integer) As Integer
    Declare Function TM1SystemServerStop Lib "tm1api.dll" (ByVal hUser As Long, ByVal szName As String, ByVal bSave As Integer) As Integer

    'Value functions****************************************************************
    'Value Array
    Declare Function TM1ValArray Lib "tm1api.dll" (ByVal hPool As Long, ByRef sArray() As Long, ByVal MaxSize As Long) As Long
    Declare Function TM1ValArrayGet Lib "tm1api.dll" (ByVal hUser As Long, ByVal vArray As Long, ByVal index As Long) As Long
    Declare Function TM1ValArrayMaxSize Lib "tm1api.dll" (ByVal hUser As Long, ByVal vArray As Long) As Long
    Declare Sub TM1ValArraySet Lib "tm1api.dll" (ByVal vArray As Long, ByVal val As Long, ByVal index As Long)
    Declare Sub TM1ValArraySetSize Lib "tm1api.dll" (ByVal vArray As Long, ByVal Size As Long)
    'Value Bool
    Declare Function TM1ValBool Lib "tm1api.dll" (ByVal hPool As Long, ByVal InitBool As Integer) As Long
    Declare Function TM1ValBoolGet Lib "tm1api.dll" (ByVal hUser As Long, ByVal vBool As Long) As Integer
    Declare Sub TM1ValBoolSet Lib "tm1api.dll" (ByVal vBool As Long, ByVal Bool As Long)
    'Value Error
    Declare Function TM1ValErrorCode Lib "tm1api.dll" (ByVal hUser As Long, ByVal vError As Long) As Long
    Declare Sub TM1ValErrorString_VB Lib "tm1api.dll" (ByVal hUser As Long, ByVal vValue As Long, ByVal Res As String, ByVal max As Integer)
    'Value Index
    Declare Function TM1ValIndex Lib "tm1api.dll" (ByVal hPool As Long, ByVal InitIndex As Long) As Long
    Declare Function TM1ValIndexGet Lib "tm1api.dll" (ByVal hUser As Long, ByVal vIndex As Long) As Long
    Declare Sub TM1ValIndexSet Lib "tm1api.dll" (ByVal vIndex As Long, ByVal index As Long)
    'Value
    Declare Function TM1ValIsUndefined Lib "tm1api.dll" (ByVal hUser As Long, ByVal Value As Long) As Long
    Declare Function TM1ValIsUpdatable Lib "tm1api.dll" (ByVal hUser As Long, ByVal Value As Long) As Integer
    'Value Object
    Declare Function TM1ValObject Lib "tm1api.dll" (ByVal hPool As Long, ByRef InitObject As Long) As Long
    Declare Function TM1ValObjectCanRead Lib "tm1api.dll" (ByVal hUser As Long, ByVal vObject As Long) As Integer
    Declare Function TM1ValObjectCanWrite Lib "tm1api.dll" (ByVal hUser As Long, ByVal vObject As Long) As Integer
    Declare Sub TM1ValObjectGet Lib "tm1api.dll" (ByVal hUser As Long, ByVal vObject As Long, ByVal pObject As String)
    Declare Sub TM1ValObjectSet Lib "tm1api.dll" (ByVal vObject As Long, ByVal pObject As String)
    Declare Function TM1ValObjectType Lib "tm1api.dll" (ByVal hUser As Long, ByVal vObject As Long) As Long
    'Value Pool
    Declare Function TM1ValPoolCount Lib "tm1api.dll" (ByVal hPool As Long) As Long
    Declare Function TM1ValPoolCreate Lib "tm1api.dll" (ByVal hUser As Long) As Long
    Declare Sub TM1ValPoolDestroy Lib "tm1api.dll" (ByVal hPool As Long)
    Declare Function TM1ValPoolGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal index As Long) As Long
    Declare Function TM1ValPoolMemory Lib "tm1api.dll" (ByVal hPool As Long) As Long
    'Value Real
    Declare Function TM1ValReal Lib "tm1api.dll" (ByVal hPool As Long, ByVal InitReal As Double) As Long
    Declare Function TM1ValRealGet Lib "tm1api.dll" (ByVal hUser As Long, ByVal vReal As Long) As Double
    Declare Sub TM1ValRealSet Lib "tm1api.dll" (ByVal vReal As Long, ByVal Real As Double)
    'Value String
    Declare Function TM1ValString Lib "tm1api.dll" (ByVal hPool As Long, ByVal InitString As String, ByVal MaxSize As Long) As Long
    Declare Function TM1ValStringMaxSize Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long) As Long
    Declare Sub TM1ValStringGet_VB Lib "tm1api.dll" (ByVal hUser As Long, ByVal vString As Long, ByVal Res As String, ByVal max As Integer)
    Declare Sub TM1ValStringSet Lib "tm1api.dll" (ByVal vString As Long, ByVal sString As String)
    'Value Type
    Declare Function TM1ValType Lib "tm1api.dll" (ByVal hUser As Long, ByVal Value As Long) As Integer

    'View functions***********************************************************
    Declare Function TM1ViewArrayColumnsNof Lib "tm1api.dll" (ByVal hPool As Long, ByVal hView As Long) As Long
    Declare Function TM1ViewArrayConstruct Lib "tm1api.dll" (ByVal hPool As Long, ByVal hView As Long) As Long
    Declare Function TM1ViewArrayDestroy Lib "tm1api.dll" (ByVal hPool As Long, ByVal hView As Long) As Long
    Declare Function TM1ViewArrayRowsNof Lib "tm1api.dll" (ByVal hPool As Long, ByVal hView As Long) As Long
    Declare Function TM1ViewArrayValueGet Lib "tm1api.dll" (ByVal hPool As Long, ByVal hView As Long, ByVal iColumn As Long, ByVal iRow As Long) As Long
    Declare Function TM1ViewCreate Lib "tm1api.dll" (ByVal hPool As Long, ByVal hCube As Long, ByVal hTitleSubsetArray As Long, ByVal hColumnSubsetArray As Long, ByVal hRowSubsetArray As Long) As Long




    '************************************************************************
    'API properties**********************************************************
    Declare Function TM1AttributeType Lib "tm1api.dll" () As Long

    Declare Function TM1BlobSize Lib "tm1api.dll" () As Long


    'Client properties*******************************************************
    Declare Function TM1ClientPassword Lib "tm1api.dll" () As Long
    Declare Function TM1ClientStatus Lib "tm1api.dll" () As Long

    'Cube properties*********************************************************
    Declare Function TM1CubeCellValueUndefined Lib "tm1api.dll" () As Long
    Declare Function TM1CubeDimensions Lib "tm1api.dll" () As Long
    Declare Function TM1CubeLogChanges Lib "tm1api.dll" () As Long
    Declare Function TM1CubeMeasuresDimension Lib "tm1api.dll" () As Long
    Declare Function TM1CubePerspectivesMaxMemory Lib "tm1api.dll" () As Long
    Declare Function TM1CubePerspectivesMinTime Lib "tm1api.dll" () As Long
    Declare Function TM1CubeRule Lib "tm1api.dll" () As Long
    Declare Function TM1CubeTimeDimension Lib "tm1api.dll" () As Long
    Declare Function TM1CubeViews Lib "tm1api.dll" () As Long

    'Dimension properties****************************************************
    Declare Function TM1DimensionCubesUsing Lib "tm1api.dll" () As Long
    Declare Function TM1DimensionElements Lib "tm1api.dll" () As Long
    Declare Function TM1DimensionNofLevels Lib "tm1api.dll" () As Long
    Declare Function TM1DimensionSubsets Lib "tm1api.dll" () As Long
    Declare Function TM1DimensionTopElement Lib "tm1api.dll" () As Long
    Declare Function TM1DimensionWidth Lib "tm1api.dll" () As Long

    'Element properties******************************************************
    Declare Function TM1ElementComponents Lib "tm1api.dll" () As Long
    Declare Function TM1ElementIndex Lib "tm1api.dll" () As Long
    Declare Function TM1ElementLevel Lib "tm1api.dll" () As Long
    Declare Function TM1ElementParents Lib "tm1api.dll" () As Long
    Declare Function TM1ElementType Lib "tm1api.dll" () As Long

    'Object properties******************************************************
    Declare Function TM1ObjectAttributes Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectChangedSinceLoaded Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectLastTimeUpdated Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectMemoryUsed Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectName Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectNull Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectParent Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectPrivate Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectPublic Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectRegistration Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectSecurityOwner Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectSecurityStatus Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectType Lib "tm1api.dll" () As Long
    Declare Function TM1ObjectUnregistered Lib "tm1api.dll" () As Long

    'Progress properties*******************************************************
    Declare Function TM1ProgressActionCalculatingSubsetAll Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionCalculatingSubsetHierarchy Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionCalculatingView Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionDeletingSelection Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionDuplicatingSubset Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionInsertingSubset Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionKeepingSelection Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionLoadingCube Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionLoadingDimension Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionLoadingSubset Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionRunningQuery Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionSavingSubset Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionSelectingSubsetElements Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressActionSortingSubset Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressMessageClosing Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressMessageOpening Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressMessageRunning Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressTypeCounter Lib "tm1api.dll" () As Long
    Declare Function TM1ProgressTypePercent Lib "tm1api.dll" () As Long

    'Rule properties*********************************************************
    Declare Function TM1RuleErrorLine Lib "tm1api.dll" () As Long
    Declare Function TM1RuleErrorString Lib "tm1api.dll" () As Long
    Declare Function TM1RuleNofLines Lib "tm1api.dll" () As Long

    'Security properties*****************************************************
    Declare Function TM1SecurityRightAdmin Lib "tm1api.dll" () As Long
    Declare Function TM1SecurityRightLock Lib "tm1api.dll" () As Long
    Declare Function TM1SecurityRightNone Lib "tm1api.dll" () As Long
    Declare Function TM1SecurityRightRead Lib "tm1api.dll" () As Long
    Declare Function TM1SecurityRightReserve Lib "tm1api.dll" () As Long
    Declare Function TM1SecurityRightWrite Lib "tm1api.dll" () As Long

    'Server properties*******************************************************
    Declare Function TM1ServerBlobs Lib "tm1api.dll" () As Long
    Declare Function TM1ServerClients Lib "tm1api.dll" () As Long
    Declare Function TM1ServerCubes Lib "tm1api.dll" () As Long
    Declare Function TM1ServerDimensions Lib "tm1api.dll" () As Long
    Declare Function TM1ServerDirectories Lib "tm1api.dll" () As Long
    Declare Function TM1ServerGroups Lib "tm1api.dll" () As Long
    Declare Function TM1ServerLogDirectory Lib "tm1api.dll" () As Long
    Declare Function TM1ServerNetworkAddress Lib "tm1api.dll" () As Long
    Declare Function TM1ServerChores Lib "tm1api.dll" () As Long
    Declare Function TM1ServerProcesses Lib "tm1api.dll" () As Long

    'Subset properties******************************************************
    Declare Function TM1SubsetAlias Lib "tm1api.dll" () As Long
    Declare Function TM1SubsetElements Lib "tm1api.dll" () As Long

    'System properties******************************************************
    Declare Function TM1SystemVersionGet Lib "tm1api.dll" () As Integer

    'Type functions*****************************************************************
    Declare Function TM1TypeAttribute Lib "tm1api.dll" () As Long
    Declare Function TM1TypeAttributeAlias Lib "tm1api.dll" () As Long
    Declare Function TM1TypeAttributeNumeric Lib "tm1api.dll" () As Long
    Declare Function TM1TypeAttributeString Lib "tm1api.dll" () As Long

    Declare Function TM1TypeBlob Lib "tm1api.dll" () As Long
    Declare Function TM1TypeClient Lib "tm1api.dll" () As Long
    Declare Function TM1TypeCube Lib "tm1api.dll" () As Long
    Declare Function TM1TypeDimension Lib "tm1api.dll" () As Long

    Declare Function TM1TypeElement Lib "tm1api.dll" () As Long
    Declare Function TM1TypeElementConsolidated Lib "tm1api.dll" () As Long
    Declare Function TM1TypeElementSimple Lib "tm1api.dll" () As Long
    Declare Function TM1TypeElementString Lib "tm1api.dll" () As Long

    Declare Function TM1TypeGroup Lib "tm1api.dll" () As Long
    Declare Function TM1TypeRule Lib "tm1api.dll" () As Long
    Declare Function TM1TypeRuleCalculation Lib "tm1api.dll" () As Long
    Declare Function TM1TypeRuleDrill Lib "tm1api.dll" () As Long
    Declare Function TM1TypeServer Lib "tm1api.dll" () As Long
    Declare Function TM1TypeSubset Lib "tm1api.dll" () As Long
    Declare Function TM1TypeView Lib "tm1api.dll" () As Long

    'Value properties********************************************************
    Declare Function TM1ValTypeArray Lib "tm1api.dll" () As Long
    Declare Function TM1ValTypeBool Lib "tm1api.dll" () As Long
    Declare Function TM1ValTypeError Lib "tm1api.dll" () As Long
    Declare Function TM1ValTypeIndex Lib "tm1api.dll" () As Long
    Declare Function TM1ValTypeObject Lib "tm1api.dll" () As Long
    Declare Function TM1ValTypeReal Lib "tm1api.dll" () As Long
    Declare Function TM1ValTypeString Lib "tm1api.dll" () As Long

    'View properties*********************************************************
    'View Array
    Declare Function TM1ViewArrayCellFormatString Lib "tm1api.dll" () As Long
    Declare Function TM1ViewArrayCellFormattedValue Lib "tm1api.dll" () As Long
    Declare Function TM1ViewArrayCellOrdinal Lib "tm1api.dll" () As Long
    Declare Function TM1ViewArrayCellValue Lib "tm1api.dll" () As Long
    Declare Function TM1ViewArrayMemberDescription Lib "tm1api.dll" () As Long
    Declare Function TM1ViewArrayMemberName Lib "tm1api.dll" () As Long
    Declare Function TM1ViewArrayMemberType Lib "tm1api.dll" () As Long
    'View
    Declare Function TM1ViewColumnSubsets Lib "tm1api.dll" () As Long
    'View Extract
    Declare Function TM1ViewExtractComparisonEQ_A Lib "tm1api.dll" () As Long
    Declare Function TM1ViewExtractComparisonGE_A Lib "tm1api.dll" () As Long
    Declare Function TM1ViewExtractComparisonGE_A_LE_B Lib "tm1api.dll" () As Long
    Declare Function TM1ViewExtractComparisonGT_A Lib "tm1api.dll" () As Long
    Declare Function TM1ViewExtractComparisonGT_A_LT_B Lib "tm1api.dll" () As Long
    Declare Function TM1ViewExtractComparisonLE_A Lib "tm1api.dll" () As Long
    Declare Function TM1ViewExtractComparisonLT_A Lib "tm1api.dll" () As Long
    Declare Function TM1ViewExtractComparisonNE_A Lib "tm1api.dll" () As Long
    Declare Function TM1ViewExtractComparisonNone Lib "tm1api.dll" () As Long
    'View
    Declare Function TM1ViewFormat Lib "tm1api.dll" () As Long
    Declare Function TM1ViewPreConstruct Lib "tm1api.dll" () As Long
    Declare Function TM1ViewRowSubsets Lib "tm1api.dll" () As Long
    Declare Function TM1ViewShowAutomatically Lib "tm1api.dll" () As Long
    Declare Function TM1ViewSuppressZeroes Lib "tm1api.dll" () As Long
    'View Title
    Declare Function TM1ViewTitleElements Lib "tm1api.dll" () As Long
    Declare Function TM1ViewTitleSubsets Lib "tm1api.dll" () As Long



    '*************************************************************************
    'TM1 Errors***************************************************************
    'Error Blob
    Declare Function TM1ErrorBlobCloseFailed Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorBlobCreateFailed Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorBlobGetFailed Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorBlobNotOpen Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorBlobOpenFailed Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorBlobPutFailed Lib "tm1api.dll" () As Long
    'Error Client
    Declare Function TM1ErrorClientAlreadyExists Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorClientPasswordNotDefined Lib "tm1api.dll" () As Long
    'Error Cube
    Declare Function TM1ErrorCubeCellValueTypeMismatch Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubeCreationFailed Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubeDimensionInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubeKeyInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubeMeasuresAndTimeDimension Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubeNotEnoughDimensions Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubeNoTimeDimension Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubeNumberOfKeysInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubePerspectiveAllSimpleElements Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubePerspectiveCreationFailed Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorCubeTooManyDimensions Lib "tm1api.dll" () As Long
    'Error Dimension
    Declare Function TM1ErrorDimensionCouldNotBeCompiled Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionElementAlreadyExists Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionElementComponentAlreadyExists Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionElementComponentDoesNotExist Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionElementComponentNotNumeric Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionElementDoesNotExist Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionElementNotConsolidated Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionHasCircularReferences Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionHasNoElements Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionIsBeingUsedByCube Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionNotChecked Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionNotRegistered Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorDimensionUpdateFailedInvalidHierarchies Lib "tm1api.dll" () As Long
    'Error Group
    Declare Function TM1ErrorGroupAlreadyExists Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorGroupMaximunNumberExceeded Lib "tm1api.dll" () As Long
    'Error Object
    Declare Function TM1ErrorObjectAttributeInvalidType Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectAttributeNotDefined Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectDeleted Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectDuplicationFailed Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectFileInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectFileNotFound Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectFunctionDoesNotApply Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectHandleInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectHasNoParent Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectIncompatibleTypes Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectIndexInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectIsRegistered Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectIsUnregistered Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectListIsEmpty Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectNameExists Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectNameInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectNameIsBlank Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectNotFound Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectNotLoaded Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectPropertyIsList Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectPropertyNotDefined Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectPropertyNotList Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectRegistrationFailed Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectSecurityIsLocked Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectSecurityNoAdminRights Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectSecurityNoLockRights Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectSecurityNoReadRights Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectSecurityNoReserveRights Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorObjectSecurityNoWriteRights Lib "tm1api.dll" () As Long
    'Error Rule
    Declare Function TM1ErrorRuleCubeHasRuleAttached Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorRuleIsAttached Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorRuleIsNotChecked Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorRuleLineNotFound Lib "tm1api.dll" () As Long
    'Error Subset
    Declare Function TM1ErrorSubsetIsBeingUsedByView Lib "tm1api.dll" () As Long
    'Error System
    Declare Function TM1ErrorSystemFunctionObsolete Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorSystemOutOfMemory Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorSystemParameterTypeInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorSystemServerClientAlreadyConnected Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorSystemServerClientNotConnected Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorSystemServerClientNotFound Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorSystemServerClientPasswordInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorSystemServerNotFound Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorSystemUserHandleInvalid Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorSystemValueInvalid Lib "tm1api.dll" () As Long
    'Error Update
    Declare Function TM1ErrorUpdateNonLeafCellValueFailed Lib "tm1api.dll" () As Long
    'Error View
    Declare Function TM1ErrorViewExpressionEmpty Lib "tm1api.dll" () As Long
    Declare Function TM1ErrorViewHasPrivateSubsets Lib "tm1api.dll" () As Long

End Module
