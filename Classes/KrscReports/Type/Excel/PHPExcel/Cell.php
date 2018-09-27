<?php
use KrscReports\Type\Excel;

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
 * @version 2.0.0, 2018-09-25
 */

/**
 * Class which object represents cell. It applies styles, value and type of cell.
 * At the end cell with given coordinates and properties can be constructed (methods inside class
 * use directly PHPExcel commands). Object properties can be changed multiple times and the same instance of
 * class can create as many cells as user requests.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Cell extends Excel\Cell
{
    /**
     * @var PHPExcel instance of used PHPExcel object
     */
    protected static $_oPHPExcel;

    /**
     * Constructor which updates PHPExcel instance inside a class.
     * @return Void
     */
    public function __construct()
    {
        self::$_oPHPExcel = KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject();
        
        // creates default object for managing style collection to handle situations, when user does not set style object
        $this->_oStyle = new KrscReports_Type_Excel_PHPExcel_Style();
    }

    /**
     * Getter for cell value (from loaded file or from PHPExcel object created during script execution)
     *
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     *
     *  @return mixed value for selected cell
     */
    public function getCellValue($iColumnId, $iRowId)
    {
        return self::$_oPHPExcel->getActiveSheet()->getCellByColumnAndRow( $iColumnId, $iRowId )->getValue();
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
        return self::$_oPHPExcel->getActiveSheet()->getCell($sCoordinate);
    }

    /**
     * Method for inserting rows before
     *
     * @param integer $iRowId number of row, before which new rows will be added
     * @param integer $iNumberOfRows number of added rows (by default 1)
     *
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function insertNewRowBefore($iRowId, $iNumberOfRows = 1)
    {
        self::$_oPHPExcel->getActiveSheet()->insertNewRowBefore($iRowId, $iNumberOfRows);

        return $this;
    }

    /**
     * Method for deleting rows.
     *
     * @param integer $iRowId number of row, for which deletion is started
     * @param integer $iNumberOfRows number of deleted rows (by default 1)
     *
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function removeRow($iRowId, $iNumberOfRows = 1)
    {
        self::$_oPHPExcel->getActiveSheet()->removeRow($iRowId, $iNumberOfRows);

        return $this;
    }

    /**
     * Method for getting column dimension.
     * @param integer $iColumnId column index (starts from 0)
     * @return PHPExcel_Worksheet_ColumnDimension column dimension for selected in input columnId
     */
    protected function _getColumnDimensionByColumn( $iColumnId )
    {
        return self::$_oPHPExcel->getActiveSheet()->getColumnDimensionByColumn( $iColumnId ); 
    }

    /**
     * Method returns letter coordinate for numeric coordinate of column.
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @return string letter coordinate of given column
     */
    public function getColumnDimension( $iColumnId )
    {
        return $this->_getColumnDimensionByColumn( $iColumnId )->getColumnIndex();
    }

    /**
     * Method sets autosize parameter for specified column.
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param boolean $bAutoSize flag for column autosize
     * @return PHPExcel_Worksheet_ColumnDimension column dimension for selected in input columnId
     */
    public function setColumnDimensionAutosize( $iColumnId, $bAutoSize = true )
    {
        return $this->_getColumnDimensionByColumn( $iColumnId )->setAutoSize( $bAutoSize );
    }

    /**
     * Method for setting on cell object fixed size for given column.
     *
     * @param integer $iColumnId id of column (starts from 0)
     *
     * @return PHPExcel_Worksheet_ColumnDimension column dimension for selected in input columnId
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
     * @return PHPExcel_Worksheet_ColumnDimension column dimension for selected in input columnId
     */
    public function createColumnSize($iColumnId, $iSize)
    {
        $this->_getColumnDimensionByColumn( $iColumnId )->setAutoSize(false);

        return $this->_getColumnDimensionByColumn( $iColumnId )->setWidth($iSize);
    }

    /**
     * Method setting auto filter for table.
     * @param integer $iColumnIdMin begin of range with filter
     * @param integer $iColumnIdMax end of range with filter 
     * @param integer $iHeaderRowHeight absolute height of row with header
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function setAutoFilter( $iColumnIdMin, $iColumnIdMax, $iHeaderRowHeight )
    {
        self::$_oPHPExcel->getActiveSheet()->setAutoFilterByColumnAndRow( $iColumnIdMin, $iHeaderRowHeight, $iColumnIdMax, $iHeaderRowHeight );
        return $this;
    }
    
    /**
     * Method setting comment for cell.
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     * @param string $sComment content of comment
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function constructCellComment( $iColumnId, $iRowId, $sComment )
    {
        self::$_oPHPExcel->getActiveSheet()->getCommentByColumnAndRow( $iColumnId, $iRowId )->getText()->createTextRun( $sComment );
        return $this;
    }
    
    /**
     * Method set styles for cell.
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function constructCellStyles( $iColumnId, $iRowId )
    {
        // setting styles
        self::$_oPHPExcel->getActiveSheet()->getStyleByColumnAndRow( $iColumnId, $iRowId )->applyFromArray( $this->_oStyle->getStyleArray( $this->_sStyleKey ) );
        
        return $this;
    }
    
    /**
     * Method creates cell with previously set properties on given in method's input coordinates 
     * (with one object with set properties more than one cell can be created; set properties can be changed
     * and new cells with new properties can be constructed from the same instance of object).
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function constructCell( $iColumnId, $iRowId )
    {        
        $this->constructCellStyles( $iColumnId, $iRowId );
        
        // constructing phpexcel element
        self::$_oPHPExcel->getActiveSheet()->setCellValueByColumnAndRow( $iColumnId, $iRowId, $this->_mValue, true);
        
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
        self::$_oPHPExcel->getActiveSheet()->mergeCellsByColumnAndRow($iBeginColumnId, $iBeginRowId, $iEndColumnId, $iEndRowId);

        return $this;
    }

    /**
     * Method attaching chart to active sheet.
     *
     * @param \PHPExcel_Chart $oChart object with chart
     */
    public function attachChart($oChart)
    {
        self::$_oPHPExcel->getActiveSheet()->addChart($oChart);
    }
}
