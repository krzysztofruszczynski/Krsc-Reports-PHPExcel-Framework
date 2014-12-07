<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Fill extends KrscReports_Type_Excel_PHPExcel_Style_Default
{
    /**
     * key for element, which is created by this object
     */
    const ARRAY_KEY = 'fill';
    
    protected function _getType()
    {
        return array();
    }
    
    protected function _getRotation()
    {
        return array();
    }
    
    protected function _getStartColor()
    {
        return array();
    }
    
    protected function _getEndColor()
    {
        return array();
    }
    
    protected function _getColor()
    {
        return array();
    }
    
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