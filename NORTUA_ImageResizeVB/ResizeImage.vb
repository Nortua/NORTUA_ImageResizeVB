Imports System.Drawing

Public Class ResizeImage

#Region " ResizeImage "
  Public Overloads Shared Function ResizeImage(SourceImage As Drawing.Image, TargetWidth As Int32, TargetHeight As Int32) As Drawing.Bitmap
    Dim bmSource = New Drawing.Bitmap(SourceImage)

    Return ResizeImage(bmSource, TargetWidth, TargetHeight)
  End Function

  Public Overloads Shared Function ResizeImage(bmSource As Drawing.Bitmap, TargetWidth As Int32, TargetHeight As Int32) As Drawing.Bitmap
    Dim bmDest As New Drawing.Bitmap(TargetWidth, TargetHeight, Drawing.Imaging.PixelFormat.Format32bppArgb)

    Dim nSourceAspectRatio = bmSource.Width / bmSource.Height
    Dim nDestAspectRatio = bmDest.Width / bmDest.Height

    Dim NewX = 0
    Dim NewY = 0
    Dim NewWidth = bmDest.Width
    Dim NewHeight = bmDest.Height

    If nDestAspectRatio = nSourceAspectRatio Then
      'same ratio
    ElseIf nDestAspectRatio > nSourceAspectRatio Then
      'Source is taller
      NewWidth = Convert.ToInt32(Math.Floor(nSourceAspectRatio * NewHeight))
      NewX = Convert.ToInt32(Math.Floor((bmDest.Width - NewWidth) / 2))
    Else
      'Source is wider
      NewHeight = Convert.ToInt32(Math.Floor((1 / nSourceAspectRatio) * NewWidth))
      NewY = Convert.ToInt32(Math.Floor((bmDest.Height - NewHeight) / 2))
    End If

    Using grDest = Drawing.Graphics.FromImage(bmDest)
      With grDest
        .CompositingQuality = Drawing.Drawing2D.CompositingQuality.HighQuality
        .InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
        .PixelOffsetMode = Drawing.Drawing2D.PixelOffsetMode.HighQuality
        .SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
        .CompositingMode = Drawing.Drawing2D.CompositingMode.SourceOver
        .FillRectangle(New SolidBrush(Color.White), 0, 0, TargetWidth, TargetHeight)
        .DrawImage(bmSource, NewX, NewY, NewWidth, NewHeight)
      End With
    End Using

    Return bmDest
  End Function
#End Region

End Class
