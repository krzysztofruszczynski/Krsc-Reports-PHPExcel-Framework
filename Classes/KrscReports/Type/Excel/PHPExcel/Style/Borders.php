<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Borders extends KrscReports_Type_Excel_PHPExcel_Style_Default
{
    /**
     * key for element, which is created by this object
     */
    const ARRAY_KEY = 'borders';  
    
    const KEY_STYLE = 'style';
    
    const KEY_COLOR = 'color';
    
    protected function _getAllBorders()
    {
        return array();
    }
    
    protected function _getLeft()
    {
        return array();
    }
    
    protected function _getRight()
    {
        return array();
    }
    
    protected function _getTop()
    {
        return array();
    }
    
    protected function _getBottom()
    {
        return array();
    }
    
    protected function _getDiagonal()
    {
        return array();
    }
    
    protected function _getDiagonalDirection()
    {
        return array();
    }
    
    /**
     * @return Array
     */
    public function getStyleArray() 
    {
        $aOutput = array();
        
        $aOutput = $this->_attachToArray( $aOutput, 'allborders', $this->_getAllBorders() );
        $aOutput = $this->_attachToArray( $aOutput, 'left', $this->_getLeft() );
        $aOutput = $this->_attachToArray( $aOutput, 'right', $this->_getRight() );
        $aOutput = $this->_attachToArray( $aOutput, 'bottom', $this->_getBottom() );
        $aOutput = $this->_attachToArray( $aOutput, 'top', $this->_getTop() );
        $aOutput = $this->_attachToArray( $aOutput, 'diagonal', $this->_getDiagonal() );
        $aOutput = $this->_attachToArray( $aOutput, 'diagonaldirection', $this->_getDiagonalDirection() );
        
        return $aOutput;
    }
}
