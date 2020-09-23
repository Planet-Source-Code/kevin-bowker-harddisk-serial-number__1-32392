VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHDSerNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'   Wrapper for the GetVolumeInformation API call.  Call the GetRegistrationNumber function, which will return
'   a Long value of the Volume Serial Number, or a 0 if the function fails.  Afterwards, call any or none of the
'   additional properties for their values.
'
'   This code received from William Bailey.
'
'   Class properties added by Kevin Bowker. GetRegistrationNumber function modified
'   by Kevin Bowker to accept DriveLetter parameter and support class properties.
'
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Const FS_CASE_IS_PRESERVED = &H2
Private Const FS_CASE_SENSITIVE = &H1
Private Const FS_UNICODE_STORED_ON_DISK = &H4
Private Const FS_PERSISTENT_ACLS = &H8
Private Const FS_FILE_COMPRESSION = &H10
Private Const FS_VOL_IS_COMPRESSED = &H8000
'   Variables To Support Local Properties - Not required/needed for this API call
Private mvarCASESENSITIVE As Boolean
Private mvarCOMPLENGTH As Long
Private mvarFILECOMPRESSED As Boolean
Private mvarFILESYSTEM As String
Private mvarPERSISTENT As Boolean
Private mvarPRESERVED As Boolean
Private mvarUNICODE As Boolean
Private mvarVOLCOMPRESSED As Boolean
Private mvarVOLNAME As String
Private mvarVOLSERNUM As Long
Public Event NoVolumeInformationAvailable()
Public Event ImproperDriveLetterFormat()
Friend Property Get AccessControlSupported() As Boolean: AccessControlSupported = mvarPERSISTENT: End Property
Friend Property Get CasePreserved() As Boolean: CasePreserved = mvarPRESERVED: End Property
Friend Property Get CaseSensitive() As Boolean: CaseSensitive = mvarCASESENSITIVE: End Property
Friend Property Get ComponentLength() As Long: ComponentLength = mvarCOMPLENGTH: End Property
Friend Property Get FileCompression() As Boolean: FileCompression = mvarFILECOMPRESSED: End Property
Friend Property Get FileSystem() As String: FileSystem = mvarFILESYSTEM: End Property
Friend Property Get Unicode() As Boolean: Unicode = mvarUNICODE: End Property
Friend Property Get VolumeCompression() As Boolean: VolumeCompression = mvarVOLCOMPRESSED: End Property
Friend Property Get VolumeName() As String: VolumeName = mvarVOLNAME: End Property
Friend Property Get VolumeSerialNumber() As Long: VolumeSerialNumber = Abs(mvarVOLSERNUM): End Property
Friend Function GetRegistrationNumber(Optional ByVal DriveLetter As String = "C:\") As Long
Dim volBuf As String, sysName As String
Dim serialNum As Long, sysFlags As Long, compLength As Long
Dim RetValue As Long
    If UCase(Format(DriveLetter, "#:\")) <> UCase(DriveLetter) Then
'        RaiseEvent ImproperDriveLetterFormat
        DriveLetter = Left$(DriveLetter, 1) & ":\"
    End If
    volBuf = String$(256, 0)
    sysName = String$(256, 0)
    RetValue = GetVolumeInformation(DriveLetter, volBuf, 255, serialNum, ComponentLength, sysFlags, sysName, 255)
    If RetValue = 0 Then
        RaiseEvent NoVolumeInformationAvailable
        GetRegistrationNumber = 0
    Else
        GetRegistrationNumber = Abs(serialNum)
        mvarVOLNAME = Left$(volBuf, InStr(1, volBuf, vbNullChar) - 1)
        mvarVOLSERNUM = serialNum
        mvarCOMPLENGTH = compLength
        mvarFILESYSTEM = Left$(sysName, InStr(1, sysName, vbNullChar) - 1)
        If sysFlags And FS_CASE_IS_PRESERVED Then mvarPRESERVED = True   '        "Case is preserved"
        If sysFlags And FS_CASE_SENSITIVE Then mvarCASESENSITIVE = True '           "File system is case sensitive"
        If sysFlags And FS_UNICODE_STORED_ON_DISK Then mvarUNICODE = True   '  "Filenames are stored as Unicode"
        If sysFlags And FS_PERSISTENT_ACLS Then mvarPERSISTENT = True   '             "Access control is supported."
        If sysFlags And FS_FILE_COMPRESSION Then mvarFILECOMPRESSED = True  '   "Individual files may be compressed"
        If sysFlags And FS_VOL_IS_COMPRESSED Then mvarVOLCOMPRESSED = True '"Entire volume is compressed"
    End If
End Function

