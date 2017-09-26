Const ForReading = 1
Const ForWriting = 2

 
Set sb = DotNetFactory.CreateInstance( "System.Text.StringBuilder" )
Set fso = CreateObject( "Scripting.FileSystemObject" )
Set textFile = fso.OpenTextFile( "C:\a\test.log", ForReading )
 
Do Until textFile.AtEndOfStream
   rLine = textFile.ReadLine
   If Left( rLine, 7 ) = "Failure" Then
      sb.AppendLine rLine
   End If
Loop
 
textFile.Close
Set textFile = fso.OpenTextFile( "C:\a\test.log", ForWriting )
textFile.Write( sb.ToString )
textFile.Close