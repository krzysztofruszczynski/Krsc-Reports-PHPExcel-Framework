<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection implements KrscReports_Type_Excel_PHPExcel_Style_Iterator_IContainer
{
    public function createIterator() 
    {
        return new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator( $this );
    }
    
    public function addStyleElement( $oElement )
    {
        
    }
    
    public function getElement( $iPosition )
    {
        
    }
    
    public function getElementCount()
    {
        
    }
}
