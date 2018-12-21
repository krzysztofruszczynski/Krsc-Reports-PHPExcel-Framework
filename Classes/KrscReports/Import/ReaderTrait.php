<?php
namespace KrscReports\Import;

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
 * @package KrscReports
 * @copyright Copyright (c) 2018 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt       LGPL
 * @version 2.0.0, 2018-09-27
 */

/**
 * Trait with functions for reading files.
 *
 * @category KrscReports
 * @package KrscReports_Import
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
trait ReaderTrait
{
    /**
     * @var KrscReports_Type_Excel_PHPExcel_Cell cell object
     */
    protected $_oCell;

    /**
     * Loading of file.
     *
     * @param string $fileName path to file
     *
     * @return KrscReports_File instance of class responsible for handling creation of file
     */
    public function loadFile($fileName)
    {
        $oFile = new \KrscReports_File();
        $oFile->setFileName($fileName);
        $oFile->setReader();
        $oFile->setWriter();
        $this->setCellObject();

        return $oFile;
    }

    /**
     * Constructor for cell object.
     *
     * @return \KrscReports_Type_Excel_PHPExcel_Cell|\KrscReports\Type\Excel\PhpSpreadsheet\Cell new object of cell
     */
    public static function getCellObject()
    {
        switch (\KrscReports_File::getBuilderType()) {
            case \KrscReports_File::SETTINGS_PHPEXCEL:
                $returnObject = new \KrscReports_Type_Excel_PHPExcel_Cell();
                break;
            case \KrscReports_File::SETTINGS_PHPSPREADSHEET:
                $returnObject = new \KrscReports\Type\Excel\PhpSpreadsheet\Cell();
                break;
        }

        return $returnObject;
    }

    /**
     * Method for setting cell object. Cell object can be constructed only when PHPExcel object was previously set.
     *
     * @param KrscReports_Type_Excel_PHPExcel_Cell $oCell stub of cell object (by default null)
     *
     * @return KrscReports_Import_TableImport object, on which this method was executed
     */
    public function setCellObject($oCell = null)
    {
        if(is_null($oCell)) {
            $this->_oCell = $this->getCellObject();
        } else {
            $this->_oCell = $oCell;
        }

        return $this;
    }
}
