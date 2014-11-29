<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
interface KrscReports_Type_Excel_PHPExcel_Style_Iterator_IIterator
{
    public function hasNextStyle();
    
    public function getStyleArray();
}
