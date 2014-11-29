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
    
    const KEY_COLOR_RGB = 'rgb';
    
    protected function _getAllBorders()
    {
        
    }
    
    
    protected function _getLeft()
    {
        
    }
    
    protected function _getTop()
    {
        
    }
    
    protected function _getBottom()
    {
        
    }
    
    protected function _getDiagonal()
    {
        
    }
    
    protected function _getDiagonalDirection()
    {
        
    }
    
    /**
     * @return Array
     */
    public function getStyleArray() 
    {
        $aOutput = array();
        $aOutput['allborders'] = $this->_getAllBorders();
        $aOutput['left'] = $this->_getLeft();
        $aOutput['top'] = $this->_getTop();
        $aOutput['right'] = $this->_getRight();
        $aOutput['bottom'] = $this->_getBottom();
        $aOutput['diagonal'] = $this->_getDiagonal();
        $aOutput['diagonaldirection'] = $this->_getDiagonalDirection();
        return $aOutput;
    }
}
