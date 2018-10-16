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
 * @version 2.0.0, 2018-10-08
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
    protected static $_oExcelObject;

    /**
     * @var array key is name of placeholder
     */
    protected $singlePlaceholders = array();

    /**
     * @var array
     */
    protected $rowPlaceholders = array();

    /**
     * @var array string[] - name of columns, for which placeholder is only part of text
     */
    protected $likePlaceholders = array();

    /**
     * @var array first key is table prefix
     */
    protected $dataArray;

    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Cell cell object
     */
    protected $_oCell;

    /**
     * @var integer number of rows in template (preferred 3)
     */
    protected $numberOfRowsInTemplate = 3;

    /**
     * Method returning placeholder, how it exactly looks in template.
     *
     * @param string $key placeholder key
     *
     * @return string placeholder key how it is presented in template (for example with brackets)
     */
    public function preparePlaceholderKey($key)
    {
        return sprintf('%s%s%s', self::PLACEHOLDER_BEFORE_KEY, $key, self::PLACEHOLDER_AFTER_KEY);
    }

    /**
     * Method setting value for single placeholder.
     *
     * @param string $key name of placeholder
     * @param string $value value of placeholder
     * @param boolean $likePlaceholder if true, placeholder can be only part of cell value; by false - only placeholder should be in cell value (by default false)
     *
     * @return $this
     */
    public function setSinglePlaceholder($key, $value, $likePlaceholder = false)
    {
        $this->singlePlaceholders[$key] = $value;
        if ($likePlaceholder) {
            $this->likePlaceholders[] = $key;
        }

        return $this;
    }

    /**
     * Method for setting row placeholder.
     *
     * @param string $tablePrefix prefix of table
     * @param string $key         name of placeholder in template
     * @param string $columnName  name of column in database output with data for specified placeholder
     *
     * @return $this
     */
    public function setRowPlaceholder($tablePrefix, $key, $columnName)
    {
        $this->rowPlaceholders[$tablePrefix][$key] = $columnName;

        return $this;
    }

    /**
     * Method for setting table data.
     *
     * @param string $tablePrefix prefix for table, for which data are set
     * @param array  $dataArray   array[] (each row is database record)
     *
     * @return $this
     */
    public function setDataArray($tablePrefix, $dataArray)
    {
        $this->dataArray[$tablePrefix] = $dataArray;

        return $this;
    }

    /**
     * Method replacing value of single placeholder in created file.
     *
     * @param string $key   name of placeholder in template
     * @param string $value value of placeholder
     *
     * @return $this
     */
    public function replaceSinglePlaceholder($key, $value)
    {
        $foundInCells = $this->findCellWithValueInDocument(
            $this->preparePlaceholderKey($key),
            in_array($key, $this->likePlaceholders)
        );

        foreach ($foundInCells as $foundInCell) {
            if (!in_array($key, $this->likePlaceholders)) {
                $valueToSet = $value;
            } else {
                $valueToSet = str_replace(
                    $this->preparePlaceholderKey($key),
                    $value,
                    $this->_oCell->getCellByCoordinate($foundInCell)->getValue()
                );
            }

            $this->_oCell->getCellByCoordinate($foundInCell)->setValue(
                $valueToSet
            );
        }

        return $this;
    }

    /**
     * Method returning coordinate with incremented row number depending of $rowIterator parameter input value.
     *
     * @param string $sCoordinate  coordinate with first row of data (can be with spreadsheet name)
     * @param integer $rowIterator number of row, for which coordinate is searched
     *
     * @return string coordinate with incremented row number
     */
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

    /**
     * Method replacing all row placeholders for each record for table with prefix given in input.
     *
     * @param string $tablePrefix prefix for table, for which placeholders are taken
     *
     * @return $this
     */
    public function replaceRowPlaceholders($tablePrefix)
    {
        $rowsInserted = false;
        $placeholders = $this->findPlaceholdersWithTablePrefix($tablePrefix);
        $rowIterator = 0;
        $numberOfRows = count($this->dataArray[$tablePrefix]);
        foreach ($this->dataArray[$tablePrefix] as $rowWithData) {
            foreach ($placeholders as $placeholderCoordinate => $placeholderName) {
                if (!$rowsInserted) {
                    $simpleCoordinate = $this->addRowsForCoordinate(
                        $placeholderCoordinate,
                        0
                    );
                    preg_match('/\d+/', $simpleCoordinate, $matches);
                    if ($numberOfRows > $this->numberOfRowsInTemplate) {
                        $this->_oCell->insertNewRowBefore($matches[0] + 1, $numberOfRows - $this->numberOfRowsInTemplate);
                    } else if ($numberOfRows < $this->numberOfRowsInTemplate) {
                        $this->_oCell->removeRow($matches[0], $this->numberOfRowsInTemplate - $numberOfRows);
                    }
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

        return $this;
    }

    /**
     * Method for searching in all worksheets given phrase
     *
     * @param string  $searchValue phrase to be searched
     * @param boolean $likeSearch  if true, searched phrase can be a part of field value, if false: it have to be full value (by default false)
     *
     * @return array string[] coordinates of cells matching search criteria
     */
    public function findCellWithValueInDocument($searchValue, $likeSearch = false)
    {
        self::$_oExcelObject = \KrscReports_Builder_Excel::getExcelObject();
        $foundInCells = array();
        foreach (self::$_oExcelObject ->getWorksheetIterator() as $worksheet) {
            $worksheetTitle = $worksheet->getTitle();
            foreach ($worksheet->getRowIterator() as $row) {
                $foundInCells = array_merge($foundInCells, $this->findCellWithValueInRow($searchValue, $row, $worksheetTitle, $likeSearch));
            }
        }

        return $foundInCells;
    }

    /**
     * Method for searching in particular row given phrase.
     *
     * @param string                  $searchValue    phrase to be searched
     * @param \PHPExcel_Worksheet_Row $row            row provided by $worksheet->getRowIterator()
     * @param string                  $worksheetTitle title of worksheet
     * @param boolean                 $likeSearch     if true, searched phrase can be a part of field value, if false: it have to be full value (by default false)
     *
     * @return array string[] coordinates of cells matching search criteria
     */
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

    /**
     * Method searching for all placeholders for given table.
     *
     * @param string $tablePrefix prefix for table, for which placeholders are searched
     *
     * @return array string[] coordinates of cells matching search criteria
     */
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

    /**
     * Method replacing all single and row placeholders
     * (have to be previously set by $this->setSinglePlaceholder, $this->setRowPlaceholder and $this->setDataArray).
     *
     * @return $this
     */
    public function replaceAllPlaceholders()
    {
        foreach ($this->singlePlaceholders as $singlePlaceholderKey => $singlePlaceholderValue) {
            $this->replaceSinglePlaceholder($singlePlaceholderKey, $singlePlaceholderValue);
        }

        foreach ($this->rowPlaceholders as $tablePrefix => $tableColumns) {
            $this->replaceRowPlaceholders($tablePrefix);
        }

        return $this;
    }
}
