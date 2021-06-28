Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Net.Mime.MediaTypeNames
Imports Image = System.Drawing.Image

Module Program

  Private Structure NameCommandLineStuct
    Dim Name As String
    Dim Value As String
  End Structure
  Private CommandLineArgs As New List(Of NameCommandLineStuct)

  Dim pathOrig As String
  Dim pathDest As String
  Dim Height As Integer
  Dim Width As Integer

  Sub Main(args As String())

    Try

      If ParseCommandLine() = False Then Console.WriteLine("usage /ImgSource=" & Chr(34) & "path of the source image" & Chr(34) & " /ImgDest=" & Chr(34) & "path of destination image" & Chr(34) & " /Height=300 /Width=300")

      Dim imgSource As Image = Image.FromFile(pathOrig)
      Dim imgDest As Image
      Dim memStream As MemoryStream = New MemoryStream()

      'imgDest = ResizeImage(imgSource, New Size(width:=Width, height:=Height))
      imgDest = ResizeImage.ResizeImage(imgSource, Width, Height)
      'imgDest.Save(memStream, ImageFormat.Jpeg)
      imgDest.Save(pathDest, ImageFormat.Gif)

      Console.WriteLine("OK")

    Catch ex As Exception
      Console.WriteLine(ex.ToString)
    End Try

  End Sub

  Private Function ParseCommandLine() As Boolean

    Try
      'step one, Do we have a command line?

      Dim clArgs() As String = Environment.GetCommandLineArgs()
      Dim Command As String = ""

      For Index As Integer = 1 To clArgs.Count - 1
        Command = Command & clArgs(Index) & " "
      Next

      If String.IsNullOrEmpty(Command) Then
        'give up if we don't
        Return False
      End If

      'does the command line have at least one named parameter?
      If Not Command.Contains("/") Then
        'give up if we don't
        Return False
      End If
      'Split the command line on our slashes.  
      Dim Params As String() = Split(Command, "/")

      'Iterate through the parameters passed
      For Each arg As String In Params
        'only process if the argument is not empty
        If Not String.IsNullOrEmpty(arg) Then
          'and contains an equal 
          If arg.Contains("=") Then

            Dim tmp As NameCommandLineStuct
            'find the equal sign
            Dim idx As Integer = arg.IndexOf("=")
            'if the equal isn't at the end of the string
            If idx < arg.Length - 1 Then
              'parse the name value pair
              tmp.Name = arg.Substring(0, idx).Trim()
              tmp.Value = arg.Substring(idx + 1).Trim()
              Select Case tmp.Name.ToLower
                Case "imgsource"
                  pathOrig = tmp.Value
                Case "imgdest"
                  pathDest = tmp.Value
                Case "height"
                  Height = Val(tmp.Value)
                Case "width"
                  Width = Val(tmp.Value)
                Case Else
                  Return False
              End Select
              'add it to the list.
              CommandLineArgs.Add(tmp)
            End If
          End If
        End If

      Next
      Return True
    Catch ex As Exception

      Console.WriteLine(ex.Message)
      Return False

    End Try

  End Function

  '#Region " ResizeImage "
  'Public Function ResizeImage(ByVal image As Image, ByVal size As Size, Optional ByVal preserveAspectRatio As Boolean = True) As Image
  ' Dim newWidth As Integer
  'Dim newHeight As Integer
  'If preserveAspectRatio Then
  'Dim originalWidth As Integer = image.Width
  'Dim originalHeight As Integer = image.Height
  'Dim percentWidth As Single = CSng(size.Width) / CSng(originalWidth)
  'Dim percentHeight As Single = CSng(size.Height) / CSng(originalHeight)
  'Dim percent As Single = If(percentHeight < percentWidth,
  '   percentHeight, percentWidth)
  '  newWidth = CInt(originalWidth * percent)
  ' newHeight = CInt(originalHeight * percent)
  'Else
  ' newWidth = size.Width
  'newHeight = size.Height
  'End If
  'Dim newImage As Image = New Bitmap(newWidth, newHeight)
  'Using graphicsHandle As Graphics = Graphics.FromImage(newImage)
  '   graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic
  '  graphicsHandle.DrawImage(image, 0, 0, newWidth, newHeight)
  'End Using
  'Return newImage
  'End Function

  '#End Region

End Module
