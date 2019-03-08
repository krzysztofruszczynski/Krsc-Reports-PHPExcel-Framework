<?php
namespace KrscReports\Type\Excel\PhpSpreadsheet;

use KrscReports\Builder;
use KrscReports\Type\Excel;

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
 * @version 2.0.2, 2019-03-08
 */

/**
 * Class which object represents cell. It applies styles, value and type of cell.
 * At the end cell with given coordinates and properties can be constructed (methods inside class
 * use directly PhpSpreadsheet commands). Object properties can be changed multiple times and the same instance of
 * class can create as many cells as user requests.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class Cell extends Excel\Cell
{
    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    protected static $_oSpreadsheet;

    /**
     * Constructor which updates PHPExcel instance inside a class.
     * @return Void
     */
    public function __construct()
    {
        self::$_oSpreadsheet = Builder\Excel\PhpSpreadsheet::getPhpSpreadsheetObject();

        // creates default object for managing style collection to handle situations, when user does not set style object
        $this->_oStyle = new \KrscReports_Type_Excel_PHPExcel_Style();
    }

    /**
     * Getter for cell value (from loaded file or from PHPExcel object created during script execution)
     *
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     *
     * @return mixed value for selected cell
     */
    public function getCellValue( $iColumnId, $iRowId )
    {
        return self::$_oSpreadsheet->getActiveSheet()->getCellByColumnAndRow( $iColumnId + 1, $iRowId )->getValue();
    }

    /**
     * Getter for cell via coordinate (from loaded file or from PHPExcel object created during script execution)
     *
     * @param string $sCoordinate cell coordinate (can include spreadsheet reference at the beginning)
     *
     * @return mixed value for selected cell
     */
    public function getCellByCoordinate($sCoordinate)
    {
        return self::$_oSpreadsheet->getActiveSheet()->getCell($sCoordinate);
    }

    /**
     * Method for inserting rows before
     *
     * @param integer $iRowId number of row, before which new rows will be added
     * @param integer $iNumberOfRows number of added rows (by default 1)
     *
     * @return \KrscReports\Type\Excel\PhpSpreadsheet\Cell object on which method was executed
     */
    public function insertNewRowBefore($iRowId, $iNumberOfRows = 1)
    {
        self::$_oSpreadsheet->getActiveSheet()->insertNewRowBefore($iRowId, $iNumberOfRows);

        return $this;
    }

    /**
     * Method for deleting rows.
     *
     * @param integer $iRowId number of row, for which deletion is started
     * @param integer $iNumberOfRows number of deleted rows (by default 1)
     *
     * @return \KrscReports\Type\Excel\PhpSpreadsheet\Cell object on which method was executed
     */
    public function removeRow($iRowId, $iNumberOfRows = 1)
    {
        self::$_oSpreadsheet->getActiveSheet()->removeRow($iRowId, $iNumberOfRows);

        return $this;
    }

    /**
     * Method for getting column dimension.
     *
     * @param integer $iColumnId column index (starts from 0)
     *
     * @return object column dimension for selected in input columnId
     */
    protected function _getColumnDimensionByColumn( $iColumnId )
    {
        return self::$_oSpreadsheet->getActiveSheet()->getColumnDimensionByColumn( $iColumnId + 1 ); 
    }

    /**
     * Method returns letter coordinate for numeric coordinate of column.
     *
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     *
     * @return string letter coordinate of given column
     */
    public function getColumnDimension( $iColumnId )
    {
        return $this->_getColumnDimensionByColumn($iColumnId)->getColumnIndex();
    }

    /**
     * Method sets autosize parameter for specified column.
     *
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param boolean $bAutoSize flag for column autosize
     *
     * @return object column dimension for selected in input columnId
     */
    public function setColumnDimensionAutosize( $iColumnId, $bAutoSize = true )
    {
        return $this->_getColumnDimensionByColumn($iColumnId)->setAutoSize( $bAutoSize );
    }

    /**
     * Method for setting on cell object fixed size for given column.
     *
     * @param integer $iColumnId id of column (starts from 0)
     *
     * @return object column dimension for selected in input columnId
     */
    public function createColumnFixedSize($iColumnId)
    {
        return $this->createColumnSize($iColumnId, (is_array($this->_mColumnFixedSizes) ? $this->_mColumnFixedSizes[$iColumnId] : $this->_mColumnFixedSizes));
    }

    /**
     *
     * @param integer $iColumnId
     * @param integer $iSize
     *
     * @return object column dimension for selected in input columnId
     */
    public function createColumnSize($iColumnId, $iSize)
    {
        $this->_getColumnDimensionByColumn( $iColumnId )->setAutoSize(false);

        return $this->_getColumnDimensionByColumn( $iColumnId )->setWidth($iSize);
    }

    /**
     * Method setting auto filter for table.
     *
     * @param integer $iColumnIdMin begin of range with filter
     * @param integer $iColumnIdMax end of range with filter 
     * @param integer $iHeaderRowHeight absolute height of row with header
     *
     * @return object object on which method was executed
     */
    public function setAutoFilter( $iColumnIdMin, $iColumnIdMax, $iHeaderRowHeight )
    {
        self::$_oSpreadsheet->getActiveSheet()->setAutoFilterByColumnAndRow( $iColumnIdMin + 1, $iHeaderRowHeight, $iColumnIdMax + 1, $iHeaderRowHeight );

        return $this;
    }

    /**
     * Method setting comment for cell.
     *
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     * @param string $sComment content of comment
     *
     * @return object object on which method was executed
     */
    public function constructCellComment( $iColumnId, $iRowId, $sComment )
    {
        self::$_oSpreadsheet->getActiveSheet()->getCommentByColumnAndRow( $iColumnId + 1, $iRowId )->getText()->createTextRun( $sComment );

        return $this;
    }

    /**
     * Method set styles for cell.
     *
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     *
     * @return object object on which method was executed
     */
    public function constructCellStyles( $iColumnId, $iRowId )
    {
        // setting styles
        self::$_oSpreadsheet->getActiveSheet()->getStyleByColumnAndRow( $iColumnId + 1, $iRowId )->applyFromArray( $this->_oStyle->getStyleArray( $this->_sStyleKey ) );

        return $this;
    }

    /**
     * Method creates cell with previously set properties on given in method's input coordinates 
     * (with one object with set properties more than one cell can be created; set properties can be changed
     * and new cells with new properties can be constructed from the same instance of object).
     *
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     *
     * @return object object on which method was executed
     */
    public function constructCell( $iColumnId, $iRowId )
    {
        if (isset($this->_sStartCell)) {
            if (!isset($this->_aFromArrayData[$iRowId])) {
                $this->_aFromArrayData[$iRowId] = array();
            }
            $this->_aFromArrayData[$iRowId][$iColumnId] = $this->_mValue;
        } else {
            $this->constructCellStyles( $iColumnId, $iRowId );
            self::$_oSpreadsheet->getActiveSheet()->setCellValueByColumnAndRow($iColumnId + 1, $iRowId, $this->_mValue);
        }

        return $this;
    }
    
    /**
     * Method for merging cells.
     *
     * @param int $iBeginColumnId begin index of column (counts from 0)
     * @param int $iBeginRowId begin index of row (counts from 1)
     * @param int $iEndColumnId end index of column (counts from 0)
     * @param int $iEndRowId end index of row (counts from 1)
     *
     * @return $this
     */
    public function mergeCells($iBeginColumnId, $iBeginRowId, $iEndColumnId, $iEndRowId)
    {
        if ($iBeginRowId == $iEndRowId) {
            for ($iSelectedColumnId = $iBeginColumnId; $iSelectedColumnId <= $iEndColumnId; $iSelectedColumnId++) {
                $this->constructCellStyles($iSelectedColumnId, $iEndRowId);
            }
        }
        self::$_oSpreadsheet->getActiveSheet()->mergeCellsByColumnAndRow($iBeginColumnId + 1, $iBeginRowId, $iEndColumnId + 1, $iEndRowId);

        return $this;
    }

    /**
     * Method saving current data for fromArray method and disabling it for newer data.
     *
     * @param integer $iBeginColumnId index of first column with data in a table
     * @param integer $iEndRowId id of last row with data
     *
     * @return void
     */
    public function endFromArrayUsage($iBeginColumnId, $iEndRowId)
    {
        self::$_oSpreadsheet->getActiveSheet()->fromArray($this->_aFromArrayData, NULL, $this->_sStartCell);
        if (!empty($this->_aFromArrayData)) {
            $iNumberOfColumns = count(current($this->_aFromArrayData));
            self::$_oSpreadsheet->getActiveSheet()->getStyle($this->_sStartCell.':'.$this->getColumnDimension($iBeginColumnId+$iNumberOfColumns-1).$iEndRowId)->applyFromArray($this->_oStyle->getStyleArray($this->_sStyleKey));
        }

        unset($this->_aFromArrayData, $this->_sStartCell);
    }

    /**
     * Method attaching chart to active sheet.
     *
     * @param \PhpOffice\PhpSpreadsheet\Chart\Chart $oChart object with chart
     */
    public function attachChart($oChart) 
    {
        self::$_oSpreadsheet->getActiveSheet()->addChart($oChart);
    }
}
