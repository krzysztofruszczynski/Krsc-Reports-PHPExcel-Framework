<?php
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
 * @version 1.2.5, 2018-04-19
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
class KrscReports_Type_Excel_PHPExcel_Cell
{
    /**
     * @var PHPExcel instance of used PHPExcel object
     */
    protected static $_oPHPExcel;
    
    /**
     * @var mixed value of current cell
     */
    protected $_mValue;
    
    /**
     * @var string type of current cell 
     */
    protected $_sType;
    
    /**
     * @var array|integer key is columnId, value is column fixed size (if integer - value apply to every column)
     */
    protected $_mColumnFixedSizes = array();
    
    /**
     * @var array|integer key is columnId, value is column fixed size  (if integer - value apply to every column)
     */
    protected $_mColumnMaxSizes = array();

    /**
     * @var KrscReports_Type_Excel_PHPExcel_Style object for managing style collections
     */
    protected $_oStyle;
    
    /**
     * @var string currently selected key of style collection
     */
    protected $_sStyleKey = KrscReports_Type_Excel_PHPExcel_Style::KEY_DEFAULT;
    
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
     * Method returns current style object (useful for adding new styles).
     * @return KrscReports_Type_Excel_PHPExcel_Style currently set style object
     */
    public function getStyleObject()
    {
        return $this->_oStyle;
    }
    
    /**
     * Method for setting object for managing style collections.
     * @param KrscReports_Type_Excel_PHPExcel_Style $oStyle instance of object for managing style collections
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function setStyleObject( KrscReports_Type_Excel_PHPExcel_Style $oStyle )
    {
        $this->_oStyle = $oStyle;
        return $this;
    }
    
    /**
     * Method for setting value of current cell.
     * @param mixed $mValue value of current cell (can be text, numeric or formula: starts with =)
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function setValue( $mValue )
    {
        $this->_mValue = $mValue;
        return $this;
    }
    
    /**
     * Getter for cell value (from loaded file or from PHPExcel object created during script execution) 
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     * @return mixed value for selected cell
     */
    public function getCellValue( $iColumnId, $iRowId )
    {
        return self::$_oPHPExcel->getActiveSheet()->getCellByColumnAndRow( $iColumnId, $iRowId )->getValue();
    }
    
    /**
     * Method setting type of cell.
     * @param string $sType type of cell to be set
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function setType( $sType )
    {
        $this->_sType = $sType;
        return $this;
    }
    
    /**
     * Method for setting current style key.
     * @param string $sStyleKey style key to be set (key under which style collection is stored)
     * @param boolean $bIfStyleNotExistsSelectDefault if true and style key was not found, default style key is set 
     * (if it is like by default false, previously set style key remains)
     * @return boolean true if style was found, false otherwise
     */
    public function setStyleKey( $sStyleKey, $bIfStyleNotExistsSelectDefault = false )
    {
       $bIsValid = $this->_oStyle->isValidStyleKey( $sStyleKey );
       
       if( $bIsValid )
       {   // style key exists 
           $this->_sStyleKey = $sStyleKey; 
       }
       else if( $bIfStyleNotExistsSelectDefault )
       {   // style key not exists but user wanted in such situation to switch to default style (otherwise - previously set style is used) 
           $this->_oStyle = KrscReports_Type_Excel_PHPExcel_Style::KEY_DEFAULT;
       }
       
       return $bIsValid;
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
     * Method checks whether fixed size for this column is already set.
     *
     * @param integer $iColumnId numeric id of column
     *
     * @return boolean true if fixed size is already set, false otherwise
     */
    public function isColumnFixedSizeIsSet( $iColumnId )
    {
        $bIsColumnFixedSizeIsSet = false;

        if (!is_array($this->_mColumnFixedSizes) && is_numeric($this->_mColumnFixedSizes)) {
            $bIsColumnFixedSizeIsSet = true;
        } else {
            $bIsColumnFixedSizeIsSet = isset($this->_mColumnFixedSizes[$iColumnId]);
        }

        return $bIsColumnFixedSizeIsSet;
    }

    /**
     * Method setting fixed size for specific column.
     * @param integer|null $iColumnId numeric id of column (null - apply to all columns)
     * @param double $dFixedSize width of column
     *
     * @return void PHPExcel_Worksheet_ColumnDimension|void column dimension for selected in input columnId (if set)
     */
    public function setColumnFixedSize($iColumnId, $dFixedSize)
    {
        if (is_null($iColumnId)) {
            $this->_mColumnFixedSizes = $dFixedSize;
        } else {
            $this->_mColumnFixedSizes[$iColumnId] = $dFixedSize;
        }
    }

    /**
     * Method checks whether max size for this column is already set.
     *
     * @param integer $iColumnId numeric id of column
     *
     * @return boolean true if max size is already set, false otherwise
     */
    public function isColumnMaxSizeIsSet($iColumnId)
    {
        $bIsColumnMaxSizeIsSet = false;

        if (!is_array($this->_mColumnMaxSizes) && is_numeric($this->_mColumnMaxSizes)) {
            $bIsColumnMaxSizeIsSet = true;
        } else {
            $bIsColumnMaxSizeIsSet = isset($this->_mColumnMaxSizes[$iColumnId]);
        }

        return $bIsColumnMaxSizeIsSet;
    }

    /**
     * Method for getting max column size.
     *
     * @param integer $iColumnId column Id
     *
     * @return integer max size of column
     */
    public function getColumnMaxSize($iColumnId)
    {
        return $this->_mColumnMaxSizes[$iColumnId];
    }

    /**
     * Method setting max size for specific column.
     *
     * @param integer|null $iColumnId numeric id of column (null - apply to all columns)
     * @param double $dMaxSize max width of column
     *
     * @return void
     */
    public function setColumnMaxSize($iColumnId, $dMaxSize)
    {
        if (is_null($iColumnId)) {
            $this->_mColumnMaxSizes = $dMaxSize;
        } else {
            $this->_mColumnMaxSizes[$iColumnId] = $dMaxSize;
        }
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
}
