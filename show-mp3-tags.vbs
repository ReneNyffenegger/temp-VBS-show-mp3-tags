Dim f2(25,3) 'folder 2 f2(x,0) is track #, f2(x,1) is song, f2(x,2) is artist, f2(x,3) is initial CD no

Dim f3(25,3) 'folder 3 

Dim id3 'As New CddbID3Tag

Dim id2 'As New CddbID3Tag

 

 Dim FS3: Set FS3 = CreateObject("Scripting.FileSystemObject")

 Dim FS2: Set FS2 = CreateObject("Scripting.FileSystemObject")

 

 

 Set id3 = CreateObject("CDDBControl.CddbID3Tag") ' for f3

 Set id2 = CreateObject("CDDBControl.CddbID3Tag") ' for f3

 

 

 

 

 'Enumerate folder files

  Dim File3

  Dim s3 'song number in album 3 

  s3 = 0

 

 

  For Each File3 In FS3.GetFolder("C:\Documents and Settings\Harry Parkinson\Desktop\3t\").Files

    'Select only mp3 files In the folder

 

    If LCase(Right(File3.Name, 4)) = ".mp3" Then

      'Load id3 data from the file

      id3.LoadFromFile File3.Path, False

 

      'Change Artist In the id3 data

      'id3.LeadArtist = Artist

  

  s3 = mid(Left(file3.name, 5), 3,2)

 

  f3(s3,0) = s3

  f3(s3,1) = id3.Title

      f3(s3,2) = id3.LeadArtist

      f3(s3,3) = mid(Left(file3.name, 5), 1,1)

  
  

      'Save modified id3 data To the file

      'id3.SaveToFile File.Path

  

    End If

 

  Next

    'id3, File3,files,FS3

 

 

  Dim File2

  Dim s2 'song number in album 3

  s2 = 0

  For Each File2 In FS2.GetFolder("C:\Documents and Settings\Harry Parkinson\Desktop\2t\").Files

    'Select only mp3 files In the folder

 

    If LCase(left(File2.Name, 1)) = "3" Then

      'Load id3 data from the file

      id2.LoadFromFile File2.Path, False

 

      'Change Artist In the id3 data

      'id3.LeadArtist = Artist

  

  s2 = mid(Left(File2.name, 5), 3,2)

 

  f2(s2,0) = s2

  f2(s2,1) = id2.Title

      f2(s2,2) = id2.LeadArtist

  f2(s2,3) = mid(Left(file2.name, 5), 1,1)

  

      'Save modified id3 data To the file

      'id3.SaveToFile File.Path

  

    End If

  Next

 

 document.write("<center> origionaly </center>")

 

  s3 = 1

  document.write( "<br> incorrect disk number " & f3(1,3) & " files located in folder 3")

 

  Do while s3 < 23

 

   document.write("<br> song " & s3 & " song title: " & f3(s3,1) & " - Song Artist: " & f3(s3,2) )

   s3 = s3 + 1

 

  Loop

 

 s2 = 1

   document.write( "<br> incorrect disk number " & f2(1,3) & " files located in folder 2")

 

 Do while s2 < 23

 

  document.write("<br> song " & f2(s2,0) & " song title: " & f2(s2,1) & " - Song Artist: " & f2(s2,2) & " - disk " & f2(s2,3))

  s2 = s2 + 1

 

 Loop

 

 document.write("<center> Changed to </center>")

 

 

 document.write( "<br> folder number 2")

 

 

  For Each File2 In FS2.GetFolder("C:\Documents and Settings\Harry Parkinson\Desktop\2t\").Files

    'Select only mp3 files In the folder

 

    If LCase(Right(File2.Name, 4)) = ".mp3" Then

      'Load id3 data from the file

      id2.LoadFromFile File2.Path, False

 

      s2 = mid(Left(file2.name, 5), 3,2)

  

      document.write("<br> song " & s2 & " song title: " & f3(s2,1) & " - Song Artist: " & f3(s2,2))

      id2.LeadArtist = f3(s2,2)

      id2.Title = f3(s2,1)

  'Save modified id3 data To the file

      id2.SaveToFile File2.Path

  
 

    End If

  Next

 

 document.write( "<br> folder number 3")

 

 

  For Each File3 In FS3.GetFolder("C:\Documents and Settings\Harry Parkinson\Desktop\3t\").Files

    'Select only mp3 files In the folder

 

    If LCase(Right(File3.Name, 4)) = ".mp3" Then

      'Load id3 data from the file

      id3.LoadFromFile File3.Path, False

 

      s3 = mid(Left(file3.name, 5), 3,2)

  

      document.write("<br> song " & s3 & " song title: " & f2(s3,1) & " - Song Artist: " & f2(s3,2))

 

  id3.LeadArtist = f2(s3,2)

      id3.Title = f2(s3,1)

      'Save modified id3 data To the file

      id3.SaveToFile File3.Path
 

    End If

  Next

 

 document.write("<center> FUCKN CHECK THE SHIT <br> By Harry Parkinson. <br> RAAKO </center>")

 

 