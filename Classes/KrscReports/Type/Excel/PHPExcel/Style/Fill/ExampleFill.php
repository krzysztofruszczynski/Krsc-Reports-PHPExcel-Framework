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
 * @version 1.0.1, 2014-12-29
 */

/**
 * Class for creating full fillment for cells (by default yellow, but another color can be set).
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Fill_ExampleFill extends KrscReports_Type_Excel_PHPExcel_Style_Fill
{
    /**
     * @var string color of fillment in rgb (light green)
     */
    protected $_sFillColor = 'CCCC00';
    
    /**
     * Method setting color of fillment.
     * @param string $sFillColor new color of fillment in rgb
     * @return KrscReports_Type_Excel_PHPExcel_Style_Fill_ExampleFill object on which method was executed
     */
    public function setFillColor( $sFillColor )
    {
        $this->_sFillColor = $sFillColor;
        return $this;
    }
    
    /**
     * Method returning type of fillment.
     * @return string type of fillment (solid)
     */
    protected function _getType() 
    {
        return PHPExcel_Style_Fill::FILL_SOLID;
    }
    
    /**
     * Method returning style array with color of fillment (default or set by user).
     * @return array style array with color of fillment 
     */
    protected function _getColor()
    {
        return self::_getColorArray( $this->_sFillColor );
    }
}
