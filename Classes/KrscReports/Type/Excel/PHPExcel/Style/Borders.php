<?php
/**
 * Main class responsible for style of borders for cell. Every class generating
 * properties associated with style of borders should inherit from that class.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Borders extends KrscReports_Type_Excel_PHPExcel_Style_Default
{
    /**
     * key for element, which is created by this object
     */
    const ARRAY_KEY = 'borders';  
    
    /**
     * key for subarray with style of border
     */
    const KEY_STYLE = 'style';
    
    /**
     * key for subarray with color
     */
    const KEY_COLOR = 'color';
    
    /**
     * Method for getting style array for all borders.
     * @return array returns empty array in that implementation
     */
    protected function _getAllBorders()
    {
        return array();
    }
    
    /**
     * Method for getting style array for left border.
     * @return array returns empty array in that implementation
     */
    protected function _getLeft()
    {
        return array();
    }
    
    /**
     * Method for getting style array for right border.
     * @return array returns empty array in that implementation
     */
    protected function _getRight()
    {
        return array();
    }
    
    /**
     * Method for getting style array for top border.
     * @return array returns empty array in that implementation
     */
    protected function _getTop()
    {
        return array();
    }
    
    /**
     * Method for getting style array for bottom border.
     * @return array returns empty array in that implementation
     */
    protected function _getBottom()
    {
        return array();
    }
    
    /**
     * Method for getting diagonal property.
     * @return array returns empty array in that implementation
     */
    protected function _getDiagonal()
    {
        return array();
    }
    
    /**
     * Method for getting diagonal direction property.
     * @return array returns empty array in that implementation
     */
    protected function _getDiagonalDirection()
    {
        return array();
    }
    
    /**
     * Method returning style array for borders.
     * @return array generated style array for borders
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
