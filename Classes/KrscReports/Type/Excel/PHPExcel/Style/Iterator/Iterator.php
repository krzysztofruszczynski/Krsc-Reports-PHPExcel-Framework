<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator implements KrscReports_Type_Excel_PHPExcel_Style_Iterator_IIterator
{
    
    protected $_oCollection;
    
    public function __construct( KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection $oCollection )
    {
        $this->_oCollection = $oCollection;
    }
    
    public function getStyleArray() 
    {
        
    }

    public function hasNextStyle() 
    {
        
    }
}
