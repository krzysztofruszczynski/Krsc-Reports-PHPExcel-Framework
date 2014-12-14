<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Fill_ExampleFill extends KrscReports_Type_Excel_PHPExcel_Style_Fill
{
    protected $_sFillColor = PHPExcel_Style_Color::COLOR_YELLOW;
    
    public function setFillColor( $sFillColor )
    {
        $this->_sFillColor = $sFillColor;
    }
    
    protected function _getType() 
    {
        return PHPExcel_Style_Fill::FILL_SOLID;
    }
    
    protected function _getColor()
    {
        return self::_getColorArray( $this->_sFillColor );
    }
}
