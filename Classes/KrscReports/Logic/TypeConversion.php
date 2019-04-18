<?php
namespace KrscReports\Logic;

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
 * @package KrscReports_Logic
 * @copyright Copyright (c) 2019 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 2.0.3, 2019-04-17
 */

/**
 * Logic handling issues related to type conversion.
 *
 * @category KrscReports
 * @package KrscReports_Logic
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class TypeConversion
{
    /**
     * Method converting excel date to PHP object.
     *
     * @param object $oInputDate object with date (Excel format)
     *
     * @return \DateTime date object (PHP format)
     */
    public static function dateFromExcelToPHP($oInputDate)
    {
        switch (\KrscReports_File::getBuilderType()) {
            case \KrscReports_File::SETTINGS_PHPEXCEL:
                $returnObject = \PHPExcel_Shared_Date::ExcelToPHPObject($oInputDate);
                break;
            case \KrscReports_File::SETTINGS_PHPSPREADSHEET:
                $returnObject = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($oInputDate);
                break;
        }

        return $returnObject;
    }
}
