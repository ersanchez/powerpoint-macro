Attribute VB_Name = "single-slide2png"
Option Explicit

' This exports the active slide as an image file
' of a width and height as designated below
' with a filename: Slide_NUMBER.FILEEXTENSION

' Set output image filetype (options: PNG, BMP, JPG)
Const OutputImageFiletype As String = "PNG"

' Set output image width
Const OutputImageWidth As Long = 1280

' Set output image height
Const OutputImageHeight As Long = 720

' Set output file directory (**must** end with a "\")
' TODO: grab the path using VBA
Const OutputFolder As String = "C:\Users\ers\Pictures"

' Set prefix for output image filename (customize on per project basis)
Const OutputImagePrefix As String = "SlideID-"

' Set macro name as seen in the menu
Sub Slide2PNG()

Dim OutputSlide As Slide

Set OutputSlide = Application.ActiveWindow.View.Slide

OutputSlide.Export OutputFolder & OutputImagePrefix & Format(OutputSlide.SlideIndex, "0000") _
& "." & OutputImageFiletype, _
OutputImageFiletype, OutputImageWidth, OutputImageHeight

End Sub
    




