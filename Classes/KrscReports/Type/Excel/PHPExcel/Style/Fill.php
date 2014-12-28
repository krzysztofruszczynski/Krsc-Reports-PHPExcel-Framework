<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2014 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2014 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.0, 2014-12-28
 */

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