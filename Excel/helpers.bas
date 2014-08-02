public sub alignShapeWithCells(shape_ as shape, address_from as string, address_to as string)

     shape_.left    = range(address_from).left
     shape_.top     = range(address_from).top
     shape_.width   = range(address_to  ).left - shape_.left
     shape_.height  = range(address_to  ).top  - shape_.top

end sub
