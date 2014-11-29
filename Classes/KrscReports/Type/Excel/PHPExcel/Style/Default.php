<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class KrscReports_Type_Excel_PHPExcel_Style_Default
{
    /**
     * key for element, which is created by this object
     */
    const ARRAY_KEY = '';    
    
    abstract public function getStyleArray();
    
    public function getArrayKey()
    {
        return static::ARRAY_KEY;
    }
}
