<?php
/**
 * Class for creating full fillment for cells (by default yellow, but another color can be set).
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Fill_ExampleFill extends KrscReports_Type_Excel_PHPExcel_Style_Fill
{
    /**
     * @var string color of fillment in rgb (by default yellow)
     */
    protected $_sFillColor = PHPExcel_Style_Color::COLOR_YELLOW;
    
    /**
     * Method setting color of fillment.
     * @param string $sFillColor new color of fillment in rgb
     * @return KrscReports_Type_Excel_PHPExcel_Style_Fill_ExampleFill object on which method was executed
     */
    public function setFillColor( $sFillColor )
    {
        $this->_sFillColor = $sFillColor;
        return $this;
    }
    
    /**
     * Method returning type of fillment.
     * @return string type of fillment (solid)
     */
    protected function _getType() 
    {
        return PHPExcel_Style_Fill::FILL_SOLID;
    }
    
    /**
     * Method returning style array with color of fillment (default or set by user).
     * @return array style array with color of fillment 
     */
    protected function _getColor()
    {
        return self::_getColorArray( $this->_sFillColor );
    }
}
