Attribute VB_Name = "mdlPictToArray"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : SaveImage
' Purpose   : Saves a StdPicture object in a byte array.
'---------------------------------------------------------------------------------------
'referencia a "edanmo's OLE Interfaces and functions v 1.81"
Public Function SaveImage( _
   ByVal image As StdPicture) As Byte()
Dim abData() As Byte
Dim oPersist As IPersistStream
Dim oStream As IStream
Dim lSize As Long
Dim tStat As STATSTG
   ' Get the image IPersistStream interface
   Set oPersist = image
   
   ' Create a stream on global memory
   Set oStream = CreateStreamOnHGlobal(0, True)
  
   ' Save the picture in the stream
   oPersist.Save oStream, True
      
   ' Get the stream info
   oStream.Stat tStat, STATFLAG_NONAME
      
   ' Get the stream size
   lSize = tStat.cbSize * 10000
   
   ' Initialize the array
   ReDim abData(0 To lSize - 1)
   
   ' Move the stream position to
   ' the start of the stream
   oStream.Seek 0, STREAM_SEEK_SET
   
   ' Read all the stream in the array
   oStream.Read abData(0), lSize
   
   ' Return the array
   SaveImage = abData
   
   ' Release the stream object
   Set oStream = Nothing

End Function

'---------------------------------------------------------------------------------------
' Procedure : LoadImage
' Purpose   : Creates a StdPicture object from a byte array.
'---------------------------------------------------------------------------------------
'
Public Function LoadImage( _
   ImageBytes() As Byte) As StdPicture
Dim oPersist As IPersistStream
Dim oStream As IStream
Dim lSize As Long
   ' Calculate the array size
   lSize = UBound(ImageBytes) - LBound(ImageBytes) + 1
   
   ' Create a stream object
   ' in global memory
   Set oStream = CreateStreamOnHGlobal(0, True)
   
   ' Write the header to the stream
   oStream.Write &H746C&, 4&
   
   ' Write the array size
   oStream.Write lSize, 4&
   
   ' Write the image data
   oStream.Write ImageBytes(LBound(ImageBytes)), lSize
   
   ' Move the stream position to
   ' the start of the stream
   oStream.Seek 0, STREAM_SEEK_SET
      
   ' Create a new empty picture object
   Set LoadImage = New StdPicture
   
   ' Get the IPersistStream interface
   ' of the picture object
   Set oPersist = LoadImage
   
   ' Load the picture from the stream
   oPersist.Load oStream
      
   ' Release the streamobject
   Set oStream = Nothing
End Function
