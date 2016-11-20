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
 * @package KrscReports_Logic
 * @copyright Copyright (c) 2016 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.3, 2016-11-20
 */

/**
 * Logic handling issues related to color.
 * 
 * @category KrscReports
 * @package KrscReports_Logic
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Logic_Color
{
    
    /**
     * Method filtering Rgb color string.
     * @param string $sColor rgb color to be filtered
     * @return string output of filter
     */
    public function filterRgb( $sColor )
    {
        if( strlen( $sColor ) == 8 )
        {
            return substr( $sColor, 2 );
        } else {
            return $sColor;
        }
    }
    
}

