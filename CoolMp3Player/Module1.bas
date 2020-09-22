Attribute VB_Name = "Module1"

Option Explicit

Type mp3Tag
        tagID As String * 3
        title As String * 30
        artist As String * 30
        album As String * 30
        year As String * 4
        comment As String * 30
        genre As String * 1
End Type

Type genreListEntry
    genre As String
End Type

Public genreList(126) As genreListEntry

Sub populateGenreList()



    genreList(0).genre = "Blues"
    genreList(1).genre = "Classic Rock"
    genreList(10).genre = "New Age"
    genreList(100).genre = "Humour"
    genreList(101).genre = "Speech"
    genreList(102).genre = "Chanson"
    genreList(103).genre = "Opera"
    genreList(104).genre = "Chamber Music"
    genreList(105).genre = "Sonata"
    genreList(106).genre = "Symphony"
    genreList(107).genre = "Booty Brass"
    genreList(108).genre = "Primus"
    genreList(109).genre = "Porn Groove"
    genreList(11).genre = "Oldies"
    genreList(110).genre = "Satire"
    genreList(111).genre = "Slow Jam"
    genreList(112).genre = "Club"
    genreList(113).genre = "Tango"
    genreList(114).genre = "Samba"
    genreList(115).genre = "Folklore"
    genreList(116).genre = "Ballad"
    genreList(117).genre = "Poweer Ballad"
    genreList(118).genre = "Rhytmic Soul"
    genreList(119).genre = "Freestyle"
    genreList(12).genre = "Other"
    genreList(120).genre = "Duet"
    genreList(121).genre = "Punk Rock"
    genreList(122).genre = "Drum Solo"
    genreList(123).genre = "A Capela"
    genreList(124).genre = "Euro-House"
    genreList(125).genre = "Dance Hall"
    genreList(13).genre = "Pop"
    genreList(14).genre = "R&B"
    genreList(15).genre = "Rap"
    genreList(16).genre = "Reggae"
    genreList(17).genre = "Rock"
    genreList(18).genre = "Techno"
    genreList(19).genre = "Industrial"
    genreList(2).genre = "Country"
    genreList(20).genre = "Alternative"
    genreList(21).genre = "Ska"
    genreList(22).genre = "Death Metal"
    genreList(23).genre = "Pranks"
    genreList(24).genre = "Soundtrack"
    genreList(25).genre = "Euro-Techno"
    genreList(26).genre = "Ambient"
    genreList(27).genre = "Trip-Hop"
    genreList(28).genre = "Vocal"
    genreList(29).genre = "Jazz+Funk"
    genreList(3).genre = "Dance"
    genreList(30).genre = "Fusion"
    genreList(31).genre = "Trance"
    genreList(32).genre = "Classical"
    genreList(33).genre = "Instrumental"
    genreList(34).genre = "Acid"
    genreList(35).genre = "House"
    genreList(36).genre = "Game"
    genreList(37).genre = "Sound Clip"
    genreList(38).genre = "Gospel"
    genreList(39).genre = "Noise"
    genreList(4).genre = "Disco"
    genreList(40).genre = "AlternRock"
    genreList(41).genre = "Bass"
    genreList(42).genre = "Soul"
    genreList(43).genre = "Punk"
    genreList(44).genre = "Space"
    genreList(45).genre = "Meditative"
    genreList(46).genre = "Instrumental Pop"
    genreList(47).genre = "InstrumentalRock"
    genreList(48).genre = "Ethnic"
    genreList(49).genre = "Gothic"
    genreList(5).genre = "Funk"
    genreList(50).genre = "Darkwave"
    genreList(51).genre = "Techno-Industrial"
    genreList(52).genre = "Electronic"
    genreList(53).genre = "Pop-Folk"
    genreList(54).genre = "Eurodance"
    genreList(55).genre = "Dream"
    genreList(56).genre = "Southern Rock"
    genreList(57).genre = "Comedy"
    genreList(58).genre = "Cult"
    genreList(59).genre = "Gangsta"
    genreList(6).genre = "Grunge"
    genreList(60).genre = "Top 40"
    genreList(61).genre = "Christian Rap"
    genreList(62).genre = "Pop/Funk"
    genreList(63).genre = "Jungle"
    genreList(64).genre = "Native American"
    genreList(65).genre = "Cabaret"
    genreList(66).genre = "New Wave"
    genreList(67).genre = "Psychadelic"
    genreList(68).genre = "Rave"
    genreList(69).genre = "Showtunes"
    genreList(7).genre = "Hip-Hop"
    genreList(70).genre = "Trailer"
    genreList(71).genre = "Lo-Fi"
    genreList(72).genre = "Tribal"
    genreList(73).genre = "Acid Punk"
    genreList(74).genre = "Acid Jazz"
    genreList(75).genre = "Polka"
    genreList(76).genre = "Retro"
    genreList(77).genre = "Musical"
    genreList(78).genre = "Rock&Roll"
    genreList(79).genre = "Hard Rock"
    genreList(8).genre = "Jazz"
    genreList(80).genre = "Folk"
    genreList(81).genre = "Folk-Rock"
    genreList(82).genre = "National Folk"
    genreList(83).genre = "Swing"
    genreList(84).genre = "Fast Fusion"
    genreList(85).genre = "Bebob"
    genreList(86).genre = "Latin"
    genreList(87).genre = "Revival"
    genreList(88).genre = "Celtic"
    genreList(89).genre = "Bluegrass"
    genreList(9).genre = "Metal"
    genreList(90).genre = "Avantgarde"
    genreList(91).genre = "Gothic Rock"
    genreList(92).genre = "Progressive Rock"
    genreList(93).genre = "Psychedelic Rock"
    genreList(94).genre = "Symphonic Rock"
    genreList(95).genre = "Slow Rock"
    genreList(96).genre = "Big Band"
    genreList(97).genre = "Chorus"
    genreList(98).genre = "Easy Listening"
    genreList(99).genre = "Acoustic"

    
    
    
    
End Sub


Function genreSearch(genreTag As String) As String

    Dim intGenreNumber As Integer
    
    If genreTag <> "" Then ' if genreTag is a nullstring we skip the Asc function
        intGenreNumber = Asc(genreTag) ' or a run-time error occurs
    Else
        intGenreNumber = 255
    End If
    
    If intGenreNumber > 125 Then ' legal genre tags run from 0 to 125
        genreSearch = "Unknown Genre Entry" ' anything else we will call unknown
    Else
        genreSearch = genreList(intGenreNumber).genre ' return the string desc of genre
    End If
    
End Function

Function getMp3Tag(strMP3FileName As String, tag As mp3Tag) As Boolean
On Error GoTo GetProblem

    Dim endOfFile As Long
    Dim iFileHandle As Integer
    
    iFileHandle = FreeFile ' Find out what the next free file number is

    Open strMP3FileName For Binary As #iFileHandle ' open MP3
    
    endOfFile = LOF(1) \ 1 'find out length of the file
    
    Get #iFileHandle, (endOfFile - 127), tag ' read in the last 128 characters of the mp3
    
    Close #iFileHandle ' close the MP3
    
    If tag.tagID = "TAG" Then 'If the tagid field contains the word "TAG" then it is a
        getMp3Tag = True 'valid MP3 Tag
    Else
        getMp3Tag = False
    End If
    
GetProblem:
Close #iFileHandle
End Function

Function putMp3Tag(strMP3FileName As String, tag As mp3Tag) As Boolean

    Dim endOfFile As Long
    Dim iFileHandle As Integer
    Dim newTag As Boolean
    
    iFileHandle = FreeFile 'find out the next free file number
    
  
    If tag.tagID <> "TAG" Then
        tag.tagID = "TAG" 'the first three bytes of the tag must be "TAG" for it to be considered valid
        newTag = True 'we are adding a new tag to the MP3
    Else
        newTag = False
    End If
    
 
    Open strMP3FileName For Binary As #iFileHandle ' open MP3
    
    endOfFile = LOF(1) \ 1 'find out length of the file
    
    If newTag Then
        Put #iFileHandle, (endOfFile + 1), tag 'add tag to end of mp3
    Else
        Put #iFileHandle, (endOfFile - 127), tag ' write tag in the last 128 characters of the mp3
    End If
    
    Close #iFileHandle ' close the MP3
    
    putMp3Tag = True
    
End Function

Function getGenreTagCode(strGenre As String) As String

Dim i As Integer
Dim genreMatchFound As Boolean

genreMatchFound = False

For i = 0 To 125 ' walk the list until a match is found and return the proper numberic value for the
    If strGenre = genreList(i).genre Then 'genre
        getGenreTagCode = Chr$(i)
        genreMatchFound = True
        Exit For ' match found, let's not waste time looking thru the rest of the list
    End If
Next i

If Not genreMatchFound Then
    getGenreTagCode = Chr$(12)
End If

End Function

