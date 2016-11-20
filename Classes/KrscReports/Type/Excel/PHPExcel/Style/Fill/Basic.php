<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2016 Krzysztof Ruszczyński
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA
 *
 * @category KrscReports
 * @package KrscReports_Type
 * @copyright Copyright (c) 2016 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.3, 2016-11-20
 */

/**
 * Class for creating full fillment for cells (by default yellow, but another color can be set).
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic extends KrscReports_Type_Excel_PHPExcel_Style_Fill
{
    /**
     * @var string type of fullfilment (default: solid)
     */
    protected $_sType = PHPExcel_Style_Fill::FILL_SOLID;
    
    /**
     * @var double value for rotation (angle)
     */
    protected $_dRotation = 0;
    
    /**
     * @var string start color of fillment in rgb (for gradient; default: not set)
     */
    protected $_sStartColor;
    
    /**
     * @var string end color of fillment in rgb (for gradient; default: not set)
     */
    protected $_sEndColor;
    
    /**
     * @var string color of fillment in rgb (default: light green)
     */
    protected $_sColor = self::COLOR_LIGHT_GREEN;
    
    /**
     * Method setting type of fillment.
     * @param string $sType new type of fullfilment
     * @return KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic object on which method was executed
     */
    public function setType( $sType )
    {
        $this->_sType = $sType;
        return $this;
    }
    
    /**
     * Method setting fill rotation.
     * @param double $dRotation value for fill rotation (angle)
     * @return KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic object on which method was executed
     */
    public function setRotation( $dRotation )
    {
        $this->_dRotation = $dRotation;
        return $this;
    }
    
    /**
     * Method setting start color for fillment.
     * @param string $sStartColor start color of fillment in rgb (for gradient; default: not set)
     * @return KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic object on which method was executed
     */
    public function setStartColor( $sStartColor )
    {
        $this->_sStartColor = $this->_oLogicColor->filterRgb( $sStartColor );
        return $this;
    }
    
    /**
     * Method setting end color for fillment.
     * @param string $sEndColor end color of fillment in rgb (for gradient; default: not set)
     * @return KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic object on which method was executed
     */
    public function setEndColor( $sEndColor )
    {
        $this->_sEndColor = $this->_oLogicColor->filterRgb( $sEndColor );
        return $this;
    }
    
    /**
     * Method setting color for fillment.
     * @param string $sColor new color of fillment in rgb
     * @return KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic object on which method was executed
     */
    public function setColor( $sColor )
    {
        $this->_sColor = $this->_oLogicColor->filterRgb( $sColor );
        return $this;
    }
    
    /**
     * Method returning type of fillment.
     * @return string type of fillment (solid)
     */
    protected function _getType() 
    {
        return $this->_sType;
    }
    
    /**
     * Method returning fill rotation.
     * @return double angle previously set
     */
    protected function _getRotation()
    {
        return $this->_dRotation;
    }
    
    /**
     * Method returning start color for fillment.
     * @return array style array with end color for gradient
     */
    protected function _getStartColor()
    {
        return self::_getColorArray( $this->_sStartColor );
    }
    
    /**
     * Method returning end color for fillment.
     * @return array style array with end color for gradient
     */
    protected function _getEndColor()
    {
        return self::_getColorArray( $this->_sEndColor );
    }
    
    /**
     * Method returning style array with color of fillment (default or set by user).
     * @return array style array with color of fillment 
     */
    protected function _getColor()
    {
        return self::_getColorArray( $this->_sColor );
    }
}
