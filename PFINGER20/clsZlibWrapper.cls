VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZlibWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CZErrors
    Z_OK = 0
    Z_STREAM_END = 1
    Z_NEED_DICT = 2
    Z_ERRNO = -1
    Z_STREAM_ERROR = -2
    Z_DATA_ERROR = -3
    Z_MEM_ERROR = -4
    Z_BUF_ERROR = -5
    Z_VERSION_ERROR = -6
End Enum

Public Enum CompressionLevels
    Z_NO_COMPRESSION = 0
    Z_BEST_SPEED = 1
    Z_BEST_COMPRESSION = 9
    Z_DEFAULT_COMPRESSION = -1
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Friend Function CompressByteArray(TheData() As Byte, ByVal CompressionLevel As CompressionLevels) As Long
    Dim result As Long
    Dim BufferSize As Long
    Dim TempBuffer() As Byte
    
    'Allocate memory for byte array
    BufferSize = UBound(TheData) + 1
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)
    
    'Compress byte array (data)
    result = compress2(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1, CompressionLevel)
    
    'Truncate to compressed size
    ReDim Preserve TheData(BufferSize - 1)
    CopyMemory TheData(0), TempBuffer(0), BufferSize
    
    'Cleanup
    Erase TempBuffer
    
    'Return error code (if any)
    CompressByteArray = result
    
End Function

Friend Function DecompressByteArray(TheData() As Byte, ByVal OrigSize As Long) As Long
    Dim result As Long
    
    'Allocate memory for buffers
    Dim BufferSize As Long
    Dim TempBuffer() As Byte
    
    BufferSize = OrigSize
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)
    
    'Decompress data
    result = uncompress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)
    
    'Truncate buffer to compressed size
    ReDim Preserve TheData(BufferSize - 1)
    CopyMemory TheData(0), TempBuffer(0), BufferSize
    
    'Return error code (if any)
    DecompressByteArray = result
End Function
