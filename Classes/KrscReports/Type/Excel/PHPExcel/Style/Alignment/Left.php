<?php
use KrscReports\Type\Excel\PhpSpreadsheet\StyleConstantsTranslatorTrait;

/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2019 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2019 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 2.0.4, 2019-06-03
 */

/**
 * Class for creating left alignment.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Alignment_Left extends KrscReports_Type_Excel_PHPExcel_Style_Alignment
{
    use StyleConstantsTranslatorTrait;

    protected function _getHorizontal()
    {
        return $this->getTranslatedStyleConstant('PHPExcel_Style_Alignment', 'HORIZONTAL_LEFT');
    }
}
