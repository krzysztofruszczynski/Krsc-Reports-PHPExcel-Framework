<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection implements KrscReports_Type_Excel_PHPExcel_Style_Iterator_IContainer
{
    /**
     * @var Array 
     */
    protected $_aCollectionElements = array();
    
    public function createIterator() 
    {
        return new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator( $this );
    }
    
    public function addStyleElement( KrscReports_Type_Excel_PHPExcel_Style_Default $oElement )
    {
        $this->_aCollectionElements[] = $oElement;
    }
    
    /**
     * 
     * @param type $iPosition
     * @return KrscReports_Type_Excel_PHPExcel_Style_Default
     */
    public function getElement( $iPosition )
    {        
        return $this->_aCollectionElements[$iPosition];
    }
    
    public function getElementCount()
    {
        return count( $this->_aCollectionElements );
    }
}
