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
 * @version 1.0.2, 2016-11-12
 */

/**
 * Class for creating dash-dot-dot dark green border (and right, which is black) for cells.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Borders_ExampleBorders extends KrscReports_Type_Excel_PHPExcel_Style_Borders
{
    /**
     * Method for getting style array for all borders.
     * @return array style array for all borders
     */
    protected function _getAllBorders()
    {
        $aOutput = array();
        $aOutput[static::KEY_STYLE] = PHPExcel_Style_Border::BORDER_DOUBLE;
        $aOutput[static::KEY_COLOR] = self::_getColorArray( PHPExcel_Style_Color::COLOR_DARKGREEN );
        return $aOutput;
    }
    
    /**
     * Method for getting style array only for right border.
     * @return array style array for only right border
     */
    protected function _getRight()
    {
        $aOutput = array();
        $aOutput[static::KEY_STYLE] = PHPExcel_Style_Border::BORDER_DOUBLE;
        $aOutput[static::KEY_COLOR] = self::_getColorArray( PHPExcel_Style_Color::COLOR_BLACK );
        return $aOutput;
    }
}