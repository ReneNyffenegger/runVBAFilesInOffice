public sub alignShapeWithCells(shape_ as shape, address_from as string, address_to as string)
'
'    This sub is used in
'      - https://github.com/ReneNyffenegger/kaggle/blob/master/titanic/oracle/analyze_data/survival_rate_dep_on_age_sex_class.bas
'      - https://github.com/ReneNyffenegger/kaggle/blob/master/titanic/oracle/analyze_data/correlation_age_fare.bas
'

     shape_.left    = range(address_from).left
     shape_.top     = range(address_from).top
     shape_.width   = range(address_to  ).left - shape_.left
     shape_.height  = range(address_to  ).top  - shape_.top

end sub
