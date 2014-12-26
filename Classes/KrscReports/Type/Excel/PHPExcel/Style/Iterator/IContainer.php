<?php
/**
 * Interface for classes creating iterator with collection of elements responsible for style of cells (Iterator design pattern).
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
interface KrscReports_Type_Excel_PHPExcel_Style_Iterator_IContainer
{
    /**
     * Method creating iterator for collection associated with current object.
     * @return KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator iterator for executed object
     */
    public function createIterator();
}