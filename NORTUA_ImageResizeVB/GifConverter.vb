
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices

Public Class GifConverter

#Region " Public Methods "
  ' ==========================================================================
  ' Public Methods
  ' ==========================================================================

  ' --------------------------------------------------------------------------
  ' Description: 
  '   Converts the specified GIF image into a transparent GIF with the
  '   specified color as transparent.
  '
  Public Shared Function GetTransparentGif(ByVal Image As Bitmap, ByVal TransparentColor As Color) As Image
    Dim w As Integer = Image.Width
    Dim h As Integer = Image.Height
    Dim bmp As Bitmap = New Bitmap(w, h, PixelFormat.Format8bppIndexed)
    Dim originalPalette As ColorPalette = Image.Palette
    Dim palette As ColorPalette = bmp.Palette

    ' Copy all of the entries from the old palette, removing any
    ' transparency.
    For i As Integer = 0 To originalPalette.Entries.Length - 1
      Dim baseColor As Color = originalPalette.Entries(i)
      Dim alpha As Integer =
          CInt(IIf(baseColor.Equals(TransparentColor), 0, 255))

      ' Set the palette entry, setting the transparent color based on
      ' whether it is the same as the transparent color argument.
      palette.Entries(i) = Color.FromArgb(alpha, baseColor)
    Next

    ' Re-insert the palette.
    bmp.Palette = palette

    ' Transfer the bitmap data from the source image to the new image.
    TransferBytes(Image, bmp)

    ' Return the new image.
    Return bmp
  End Function
#End Region

#Region " Private Methods "
  ' ==========================================================================
  ' Private Methods
  ' ==========================================================================

  ' --------------------------------------------------------------------------
  ' Description: 
  '   Lock the bits of the specified image and return the BitmapData
  '   structure for the bits locked.
  '
  Private Shared Function LockBits(
      ByVal Image As Bitmap,
      ByVal LockMode As ImageLockMode
  ) As BitmapData
    Return Image.LockBits(
        New Rectangle(0, 0, Image.Width, Image.Height),
        LockMode,
        Image.PixelFormat
    )
  End Function

  ' --------------------------------------------------------------------------
  ' Description: 
  '   Transfer the bytes of the source image to the destination image.
  '
  Private Shared Sub TransferBytes(
      ByVal SrcImg As Bitmap,
      ByVal DstImg As Bitmap
  )
    Dim srcData As BitmapData = LockBits(SrcImg, ImageLockMode.ReadOnly)
    Dim dstData As BitmapData = LockBits(DstImg, ImageLockMode.WriteOnly)

    Try
      For y As Integer = 0 To SrcImg.Height - 1
        For x As Integer = 0 To SrcImg.Width - 1
          ' Calculate the offsets of the next byte to read/write.
          Dim srcOffset As Integer = srcData.Stride * y + x
          Dim dstOffset As Integer = dstData.Stride * y + x

          ' Read the next byte from the source bitmap.
          Dim b As Byte = Marshal.ReadByte(srcData.Scan0, srcOffset)

          ' Write the byte to the destination bitmap.
          Marshal.WriteByte(dstData.Scan0, dstOffset, b)
        Next x
      Next y
    Finally
      ' Cleanup (unlock the bits of the bitmaps).
      SrcImg.UnlockBits(srcData)
      DstImg.UnlockBits(dstData)
    End Try
  End Sub
#End Region

End Class
