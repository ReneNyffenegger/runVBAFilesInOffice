'  \lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -word backgroundImage -c main %CD%
'
'  http://renenyffenegger.blogspot.ch/2014/10/creating-psychedelic-images-with-word.html

sub main(path)

    dim background_image as shape

  ' Load image
    set background_image = activeDocument.shapes.addPicture( _
      fileName    := path & "\background.png",  _
      linkToFile  :=  false)

  ' Place Image's top left corner to page's top left corner:
    background_image.relativeVerticalPosition   = wdRelativeVerticalPositionPage
    background_image.top                        = 0

    background_image.relativeHorizontalPosition = wdRelativeHorizontalPositionPage
    background_image.left                       = 0

  ' Use entire page (size of paper is assumed to be A4)
    background_image.width  = inchesToPoints( 8.267)
    background_image.height = inchesToPoints(11.69 )

  ' Allow to write text above the image
    background_image.zOrder msoSendBehindText

  ' Close document without being asked
    activeDocument.saved = true

end sub

