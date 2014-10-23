'  \lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -word backgroundImage -c main %CD%

sub main(path)

    dim background_image as shape

    set background_image = activeDocument.shapes.addPicture( _
      fileName    := path & "\background.png",  _ 
      linkToFile  :=  false)

    background_image.relativeVerticalPosition   = wdRelativeVerticalPositionPage
    background_image.top                        = 0

    background_image.relativeHorizontalPosition = wdRelativeHorizontalPositionPage
    background_image.left                       = 0

    background_image.width  = inchesToPoints( 8.267)
    background_image.height = inchesToPoints(11.69 )

    background_image.zOrder msoSendBehindText

    activeDocument.saved = true

end sub

