<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2017 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2017 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.6, 2017-02-21
 */

/**
 * Main class responsible for style of font for cell. Every class generating
 * properties associated with style of font should inherit from that class.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Font extends KrscReports_Type_Excel_PHPExcel_Style_Default
{
    /**
     * key for element, which is created by this object
     */
    const ARRAY_KEY = 'font';  
    
    /**
     * Method for getting name of font ( Arial, Calibri, etc. ).
     * @return string returns empty string in that implementation
     */
    protected function _getName()
    {
        return '';
    }
    
    /**
     * Method for getting flag, if font in cell is bolded.
     * @return boolean returns false in that implementation
     */
    protected function _getBold()
    {
        return false;
    }
    
    /**
     * Method for getting flag, if font in cell is italic.
     * @return boolean returns false in that implementation
     */
    protected function _getItalic()
    {
        return false;
    }
    
    /**
     * Method for getting flag for super script.
     * @return boolean returns false in that implementation
     */
    protected function _getSuperScript()
    {
        return false;
    }
    
    /**
     * Method for getting flag for sub script.
     * @return boolean returns false in that implementation
     */
    protected function _getSubScript()
    {
        return false;
    }
    
    /**
     * Method for getting flag for underline.
     * @return boolean returns false in that implementation
     */
    protected function _getUnderline()
    {
        return false;
    }
    
    /**
     * Method for getting flag for strike.
     * @return boolean returns false in that implementation
     */
    protected function _getStrike()
    {
        return false;
    }
    
    /**
     * Method for getting font size.
     * @return integer returns 11 in that implementation
     */
    protected function _getSize()
    {
        return 11;
    }
   
    /**
     * Method for getting color of font.
     * @return array returns array with data for black color
     */
    protected function _getColor()
    {
        return self::_getColorArray( PHPExcel_Style_Color::COLOR_BLACK );
    }
    
    /**
     * Method returning style array for alignment.
     * @return array generated style array for alignment
     */
    public function getStyleArray() 
    {
        $aOutput = array();
        
        $aOutput = $this->_attachToArray( $aOutput, 'name', $this->_getName() );
        $aOutput = $this->_attachToArray( $aOutput, 'bold', $this->_getBold() );
        $aOutput = $this->_attachToArray( $aOutput, 'italic', $this->_getItalic() );
        $aOutput = $this->_attachToArray( $aOutput, 'superScript', $this->_getSuperScript() );
        $aOutput = $this->_attachToArray( $aOutput, 'subScript', $this->_getSubScript() );
        $aOutput = $this->_attachToArray( $aOutput, 'underline', $this->_getUnderline() );
        $aOutput = $this->_attachToArray( $aOutput, 'strike', $this->_getStrike() );
        $aOutput = $this->_attachToArray( $aOutput, 'size', $this->_getSize() );
        $aOutput = $this->_attachToArray( $aOutput, 'color', $this->_getColor() );
        
        return $aOutput;
    }
}


