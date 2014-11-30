<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator implements KrscReports_Type_Excel_PHPExcel_Style_Iterator_IIterator
{
    protected $_iCurrentPosition;
    
    protected $_oCollection;
    
    public function __construct( KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection $oCollection )
    {
        $this->_oCollection = $oCollection;
        $this->_iCurrentPosition = 0;
    }
    
    /**
     * 
     * @return KrscReports_Type_Excel_PHPExcel_Style_Default
     */
    public function getStyleElement() 
    {
        return $this->_oCollection->getElement( $this->_iCurrentPosition++ );
    }

    /**
     * 
     * @return Boolean
     */
    public function hasNextElement() 
    {      
        return $this->_oCollection->getElementCount() > $this->_iCurrentPosition;
    }
    
    public function resetIterator()
    {
        $this->_iCurrentPosition = 0;
    }
}
