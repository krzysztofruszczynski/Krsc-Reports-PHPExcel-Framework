<?php
/**
 * Interface for classes with iterator on collection of elements responsible for style of cell (Iterator design pattern).
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
interface KrscReports_Type_Excel_PHPExcel_Style_Iterator_IIterator
{
    /**
     * Method checking, whether there are more elements in collection.
     * @return boolean true if there are more elements in collection, false otherwise
     */
    public function hasNextElement();
    
    /**
     * Method returning current element of collection (each time executed it will give next element).
     * @return KrscReports_Type_Excel_PHPExcel_Style_Default element responsible for style of cell
     */
    public function getStyleElement();
    
    /**
     * Method setting pointer to first element of collection.
     * @return KrscReports_Type_Excel_PHPExcel_Style_Iterator_IIterator object on which method was executed
     */
    public function resetIterator();
}
