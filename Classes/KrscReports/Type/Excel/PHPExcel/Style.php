<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style
{
    /**
     * @var KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator
     */
    protected $_oStyleIterator;
    
    public function setStyleCollection( KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection $oStyleCollection )
    {
        $this->_oStyleIterator = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator( $oStyleCollection );
    }
    
    public function getStyleArray() 
    {
        $this->_oStyleIterator->resetIterator();
        $aOutput = array();
        
        while( $this->_oStyleIterator->hasNextElement() )
        {
            $oStyleElement = $this->_oStyleIterator->getStyleElement();
            $aOutput[$oStyleElement->getArrayKey()] = $oStyleElement->getStyleArray();
        }
        
        return $aOutput;
    }
}
