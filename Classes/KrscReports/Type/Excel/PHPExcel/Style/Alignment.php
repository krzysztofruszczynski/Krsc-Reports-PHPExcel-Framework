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
 * @version 1.0.6, 2017-02-05
 */

/**
 * Main class responsible for style of alignment for cell. Every class generating
 * properties associated with style of alignment should inherit from that class.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Alignment extends KrscReports_Type_Excel_PHPExcel_Style_Default
{
    /**
     * key for element, which is created by this object
     */
    const ARRAY_KEY = 'alignment';  
    
    /**
     * Method for getting horizontal style.
     * @return string returns empty string in that implementation
     */
    protected function _getHorizontal()
    {
        return '';
    }
    
    /**
     * Method for getting vertical style.
     * @return string returns empty string in that implementation
     */
    protected function _getVertical()
    {
        return '';
    }
    
    /**
     * Method for getting value for rotation.
     * @return integer returns zero in that implementation
     */
    protected function _getRotation()
    {
        return 0;
    }
    
    /**
     * Method for getting flag for wrap.
     * @return boolean returns false in that implementation
     */
    protected function _getWrap()
    {
        return false;
    }
    
    /**
     * Method for getting flag for shrinkToFit.
     * @return boolean returns false in that implementation
     */
    protected function _getShrinkToFit()
    {
        return false;
    }
    
    /**
     * Method for getting diagonal property.
     * @return integer returns empty array in that implementation
     */
    protected function _getIndent()
    {
        return 0;
    }
   
    /**
     * Method returning style array for alignment.
     * @return array generated style array for alignment
     */
    public function getStyleArray() 
    {
        $aOutput = array();
        
        $aOutput = $this->_attachToArray( $aOutput, 'horizontal', $this->_getHorizontal() );
        $aOutput = $this->_attachToArray( $aOutput, 'vertical', $this->_getVertical() );
        $aOutput = $this->_attachToArray( $aOutput, 'rotation', $this->_getRotation() );
        $aOutput = $this->_attachToArray( $aOutput, 'wrap', $this->_getWrap() );
        $aOutput = $this->_attachToArray( $aOutput, 'shrinkToFit', $this->_getShrinkToFit() );
        $aOutput = $this->_attachToArray( $aOutput, 'indent', $this->_getIndent() );
        
        return $aOutput;
    }
}


