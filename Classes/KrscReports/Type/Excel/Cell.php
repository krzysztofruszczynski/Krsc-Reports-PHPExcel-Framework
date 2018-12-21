<?php
namespace KrscReports\Type\Excel;

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
 * Class which object represents cell. Have to be extended by specific tool (like PHPExcel or PhpSpreadsheet). It applies styles, value and type of cell.
 * Object properties can be changed multiple times and the same instance of
 * class can create as many cells as user requests.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class Cell 
{
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
    protected $_sStyleKey = \KrscReports_Type_Excel_PHPExcel_Style::KEY_DEFAULT;

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
    public function setStyleObject( \KrscReports_Type_Excel_PHPExcel_Style $oStyle )
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
           $this->_oStyle = \KrscReports_Type_Excel_PHPExcel_Style::KEY_DEFAULT;
       }

       return $bIsValid;
    }

    /**
     * Method setting fixed size for specific column.
     * @param integer|null $iColumnId numeric id of column (null - apply to all columns)
     * @param double $dFixedSize width of column
     *
     * @return void
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
     * Method for getting max column size.
     *
     * @param integer $iColumnId column Id
     *
     * @return integer max size of column
     */
    public function getColumnMaxSize($iColumnId)
    {
        if (is_array($this->_mColumnMaxSizes)) {
            $iMaxSize = $this->_mColumnMaxSizes[$iColumnId];
        } else {
            $iMaxSize = $this->_mColumnMaxSizes;
        }

        return $iMaxSize;
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

    abstract public function attachChart($oChart);
}
