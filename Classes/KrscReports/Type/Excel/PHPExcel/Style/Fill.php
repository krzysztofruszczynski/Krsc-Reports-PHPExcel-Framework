<?php
/**
 * Main class responsible for style of fillment for cell. Every class generating
 * properties associated with style of fillment should inherit from that class.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Fill extends KrscReports_Type_Excel_PHPExcel_Style_Default
{
    /**
     * key for element, which is created by this object
     */
    const ARRAY_KEY = 'fill';
    
    /**
     * Method returning type of fillment.
     * @return string returns empty array in that implementation
     */
    protected function _getType()
    {
        return '';
    }
    
    /**
     * Method returning rotation.
     * @return double returns null in that implementation
     */
    protected function _getRotation()
    {
        return NULL;
    }
    
    /**
     * Method returning start color for fillment.
     * @return array returns empty array in that implementation
     */
    protected function _getStartColor()
    {
        return array();
    }
    
    /**
     * Method returning end color for fillment.
     * @return array returns empty array in that implementation
     */
    protected function _getEndColor()
    {
        return array();
    }
    
    /**
     * Method returning color for fillment.
     * @return array returns empty array in that implementation
     */
    protected function _getColor()
    {
        return array();
    }
    
    /**
     * Method returning style array for fillment.
     * @return array generated style array for fillment
     */
    public function getStyleArray() 
    {
        $aOutput = array();
        
        $aOutput = $this->_attachToArray( $aOutput, 'type', $this->_getType() );
        $aOutput = $this->_attachToArray( $aOutput, 'rotation', $this->_getRotation() );
        $aOutput = $this->_attachToArray( $aOutput, 'startcolor', $this->_getStartColor() );
        $aOutput = $this->_attachToArray( $aOutput, 'endcolor', $this->_getEndColor() );
        $aOutput = $this->_attachToArray( $aOutput, 'color', $this->_getColor() );
        
        return $aOutput;
    }
}