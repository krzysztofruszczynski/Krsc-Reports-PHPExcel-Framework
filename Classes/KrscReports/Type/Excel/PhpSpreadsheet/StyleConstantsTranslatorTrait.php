<?php
namespace KrscReports\Type\Excel\PhpSpreadsheet;

/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2020 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2020 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 2.1.0, 2020-02-13
 */

/**
 * Trait for translating constants for styling from PHPExcel to PhpSpreadsheet.
 *
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
trait StyleConstantsTranslatorTrait {
    /**
     * @var array array key is class name in PHPExcel, array value is class name in PhpSpreadsheet
     */
    protected $_aPHPExcelToPhpSpreadsheetClasses = array(
        'PHPExcel_Chart_DataSeries' => '\PhpOffice\PhpSpreadsheet\Chart\DataSeries',
        'PHPExcel_Style_Alignment' => '\PhpOffice\PhpSpreadsheet\Style\Alignment',
        'PHPExcel_Style_Border' => '\PhpOffice\PhpSpreadsheet\Style\Border',
        'PHPExcel_Style_Color' => '\PhpOffice\PhpSpreadsheet\Style\Color',
        'PHPExcel_Style_Fill' => '\PhpOffice\PhpSpreadsheet\Style\Fill',
    );

    /**
     * Method for getting proper style constant value.
     *
     * @param string $sClassName    class name in PHPExcel
     * @param string $sConstantName constant name inside selected class
     *
     * @return string value of selected constant
     *
     * @throws \Exception when builder type by \KrscReports_File is not configured
     */
    public function getTranslatedStyleConstant($sClassName, $sConstantName)
    {
        switch (\KrscReports_File::getBuilderType()) {
            case \KrscReports_File::SETTINGS_PHPEXCEL:
                $sOutputKey = constant(sprintf('%s::%s', $sClassName, $sConstantName));
                break;
            case \KrscReports_File::SETTINGS_PHPSPREADSHEET:
                $sOutputKey = constant(sprintf('%s::%s', $this->_aPHPExcelToPhpSpreadsheetClasses[$sClassName], $sConstantName));
                break;
            default:
                throw new \Exception('Builder Type not set - unable to establish constant');
       }

       return $sOutputKey;
    }
}
