<?php
namespace KrscReports\Type\Excel\PhpSpreadsheet;

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
 * Trait for translating keys for styling from PHPExcel to PhpSpreadsheet.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
trait StyleKeysTranslatorTrait {
    /**
     * @var array array key is value in PHPExcel, array value is value in PhpSpreadsheet
     */
    protected $_aPHPExcelToPhpSpreadsheetTranslations = array(
        // numberFormat:
        'code' => 'formatCode',
        // font:
        'strike' => 'strikethrough',
        'superScript' => 'superScript',
        'subScript' => 'subscript',
        // alignment:
        'rotation' => 'textRotation',
        'readorder' => 'readOrder',
        'wrap' => 'wrapText',
        // borders:
        'diagonaldirection' => 'diagonalDirection',
        'allborders' => 'allBorders',
        // inside allborders
        'style' => 'borderStyle',
        // fill:
        'type' => 'fillType',
        'startcolor' => 'startColor',
        'endcolor' => 'endColor'
    );

    /**
     * Method for getting proper style key.
     *
     * @param string $sKey style key in PHPExcel
     *
     * @return string style key in actually used vendor
     *
     * @throws \Exception when builder type by \KrscReports_File is not configured
     */
    public function getTranslatedStyleKey($sKey)
    {
        switch (\KrscReports_File::getBuilderType()) {
            case \KrscReports_File::SETTINGS_PHPEXCEL:
                $sOutputKey = $sKey;
                break;
            case \KrscReports_File::SETTINGS_PHPSPREADSHEET:
                $sOutputKey = $this->translateStyleKeyToPhpSpreadsheet($sKey);
                break;
            default:
                throw new \Exception('Builder Type not set - unable to establish style key');
       }

       return $sOutputKey;
    }

    /**
     * Method for getting style key in PhpSpreadsheet.
     *
     * @param string $sKey style key in PHPExcel
     *
     * @return string style key in PhpSpreadsheet
     */
    public function translateStyleKeyToPhpSpreadsheet($sKey)
    {
        if (isset($this->_aPHPExcelToPhpSpreadsheetTranslations[$sKey])) {
            $sOutputKey = $this->_aPHPExcelToPhpSpreadsheetTranslations[$sKey];
        } else {
            $sOutputKey = $sKey;
        }

        return $sOutputKey;
    }
}
