<?php
use KrscReports\Type\Excel\PhpSpreadsheet\StyleKeysTranslatorTrait;

/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2018 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2018 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 2.0.0, 2018-09-28
 */

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
    use StyleKeysTranslatorTrait;

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
        
        $aOutput = $this->_attachToArray( $aOutput, $this->getTranslatedStyleKey('allborders'), $this->_getAllBorders() );
        $aOutput = $this->_attachToArray( $aOutput, 'left', $this->_getLeft() );
        $aOutput = $this->_attachToArray( $aOutput, 'right', $this->_getRight() );
        $aOutput = $this->_attachToArray( $aOutput, 'bottom', $this->_getBottom() );
        $aOutput = $this->_attachToArray( $aOutput, 'top', $this->_getTop() );
        $aOutput = $this->_attachToArray( $aOutput, 'diagonal', $this->_getDiagonal() );
        $aOutput = $this->_attachToArray( $aOutput, $this->getTranslatedStyleKey('diagonaldirection'), $this->_getDiagonalDirection() );

        return $aOutput;
    }
}
