<?php
/**
 * Class for creating collection of elements responsible for style of cells (Iterator design pattern).
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection implements KrscReports_Type_Excel_PHPExcel_Style_Iterator_IContainer
{
    /**
     * @var array subsequent elements which are part of collection 
     */
    protected $_aCollectionElements = array();
    
    /**
     * Method creating iterator for collection associated with current object.
     * @return KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator iterator for executed object
     */
    public function createIterator() 
    {
        return new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator( $this );
    }
    
    /**
     * Method for adding next element responsible for style of cell.
     * @param KrscReports_Type_Excel_PHPExcel_Style_Default $oElement style element to be added to current collection
     * @return KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection object on which method was executed
     */
    public function addStyleElement( KrscReports_Type_Excel_PHPExcel_Style_Default $oElement )
    {
        $this->_aCollectionElements[] = $oElement;
        return $this;
    }
    
    /**
     * Method returning style element with given position.
     * @param integer $iPosition position of element to return (starts from 0)
     * @return KrscReports_Type_Excel_PHPExcel_Style_Default requested element responsible for style of cell
     */
    public function getElement( $iPosition )
    {        
        return $this->_aCollectionElements[$iPosition];
    }
    
    /**
     * Method returning number of style elements associated with collection in current object.
     * @return integer number of style elements inside collection
     */
    public function getElementCount()
    {
        return count( $this->_aCollectionElements );
    }
}
