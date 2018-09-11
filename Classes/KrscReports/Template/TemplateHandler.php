<?php
namespace KrscReports\Template;

use KrscReports\Import\ReaderTrait;

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
 * @package KrscReports_Template
 * @copyright Copyright (c) 2018 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt       LGPL
 * @version 1.2.8, 2018-09-11
 */

/**
 * Class for handling excel templates and populate them with data.
 *
 * @category KrscReports
 * @package KrscReports_Template
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class TemplateHandler
{
    use ReaderTrait;

    /**
     * Beginning of placeholder for template
     */
    const PLACEHOLDER_BEFORE_KEY = '<<';

    /**
     * End of placeholder for template
     */
    const PLACEHOLDER_AFTER_KEY = '>>';

    /**
     * Element separating table prefix from column in placeholder
     */
    const PLACEHOLDER_ELEMENT_SEPARATOR = '_';

    /**
     * @var \PHPExcel instance of used PHPExcel object
     */
    protected static $_oPHPExcel;

    /**
     * @var array key is name of placeholder
     */
    protected $singlePlaceholders = array();

    /**
     * @var array
     */
    protected $rowPlaceholders = array();

    /**
     * @var array first key is table prefix
     */
    protected $dataArray;

    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Cell cell object
     */
    protected $_oCell;

    /**
     *
     * @param string $key
     *
     * @return string
     */
    public function preparePlaceholderKey($key)
    {
        return sprintf('%s%s%s', self::PLACEHOLDER_BEFORE_KEY, $key, self::PLACEHOLDER_AFTER_KEY);
    }

    public function setSinglePlaceholder($key, $value)
    {
        $this->singlePlaceholders[$key] = $value;
    }

    public function setRowPlaceholder($tablePrefix, $key, $columnName)
    {
        $this->rowPlaceholders[$tablePrefix][$key] = $columnName;
    }

    public function setDataArray($tablePrefix, $dataArray)
    {
        $this->dataArray[$tablePrefix] = $dataArray;
    }

    public function replaceSinglePlaceholder($key, $value)
    {
        $foundInCells = $this->findCellWithValueInDocument(
            $this->preparePlaceholderKey($key)
        );

        foreach ($foundInCells as $foundInCell) {
            $this->_oCell->getCellByCoordinate($foundInCell)->setValue(
                $value
            );
        }
    }

    public function addRowsForCoordinate($sCoordinate, $rowIterator)
    {
        if (stripos($sCoordinate, '!')) {
            $simpleCoordinate = explode('!', $sCoordinate)[1];
        } else {
            $simpleCoordinate = $sCoordinate;
        }
        preg_match('/\d+/', $simpleCoordinate, $matches);
        $sNewCoordinate = str_replace($matches[0], $matches[0] + $rowIterator, $sCoordinate);

        return $sNewCoordinate;
    }

    public function replaceRowPlaceholders($tablePrefix)
    {
        $rowsInserted = false;
        $placeholders = $this->findPlaceholdersWithTablePrefix($tablePrefix);
        $rowIterator = 0;
        foreach ($this->dataArray[$tablePrefix] as $rowWithData) {
            foreach ($placeholders as $placeholderCoordinate => $placeholderName) {
                if (!$rowsInserted) {
                    $simpleCoordinate = $this->addRowsForCoordinate(
                        $placeholderCoordinate,
                        0
                    );
                    preg_match('/\d+/', $simpleCoordinate, $matches);
                    $this->_oCell->insertNewRowBefore($matches[0] + 1, count($this->dataArray[$tablePrefix]) - 3);
                    $rowsInserted = true;
                }
                $columnName = $this->rowPlaceholders[$tablePrefix][$placeholderName];
                $cellValue = $rowWithData[$columnName];
                $this->_oCell->getCellByCoordinate(
                    $this->addRowsForCoordinate(
                        $placeholderCoordinate,
                        $rowIterator
                    )
                )->setValue($cellValue);
            }
            $rowIterator++;
        }
    }

    public function findCellWithValueInDocument($searchValue, $likeSearch = false)
    {
        self::$_oPHPExcel = \KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject();
        $foundInCells = array();
        foreach (self::$_oPHPExcel ->getWorksheetIterator() as $worksheet) {
            $worksheetTitle = $worksheet->getTitle();
            foreach ($worksheet->getRowIterator() as $row) {
                $foundInCells = array_merge($foundInCells, $this->findCellWithValueInRow($searchValue, $row, $worksheetTitle, $likeSearch));
            }
        }

        return $foundInCells;
    }

    public function findCellWithValueInRow($searchValue, $row, $worksheetTitle, $likeSearch = false)
    {
        $foundInCells = array();
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(true);
        foreach ($cellIterator as $cell) {
            if ((!$likeSearch && $cell->getValue() == $searchValue) || ($likeSearch && stripos($cell->getValue(), $searchValue) !== false)) {
                $foundInCells[] = $worksheetTitle . '!' . $cell->getCoordinate();
            }
        }

        return $foundInCells;
    }

    public function findPlaceholdersWithTablePrefix($tablePrefix)
    {
        $searchValue = sprintf('%s%s', self::PLACEHOLDER_BEFORE_KEY, $tablePrefix);
        $foundInCells = $this->findCellWithValueInDocument($searchValue, true);
        $placeholders = array();

        foreach ($foundInCells as $foundInCell) {
            $placeholderName = explode(
                self::PLACEHOLDER_ELEMENT_SEPARATOR,
                str_replace(
                    self::PLACEHOLDER_AFTER_KEY,
                    '',
                    $this->_oCell->getCellByCoordinate($foundInCell)->getValue()
                )
            )[1];

             $placeholders[$foundInCell] = $placeholderName;
        }

        return $placeholders;
    }

    public function replaceAllPlaceholders()
    {
        foreach ($this->singlePlaceholders as $singlePlaceholderKey => $singlePlaceholderValue) {
            $this->replaceSinglePlaceholder($singlePlaceholderKey, $singlePlaceholderValue);
        }

        foreach ($this->rowPlaceholders as $tablePrefix => $tableColumns) {
            $this->replaceRowPlaceholders($tablePrefix);
        }
    }
}
