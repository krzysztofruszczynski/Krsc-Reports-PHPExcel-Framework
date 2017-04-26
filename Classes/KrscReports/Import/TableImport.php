<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2017 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2017 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.1.0, 2017-04-25
 */

/**
 * Class handling import of data in table structure.
 * 
 * @category KrscReports
 * @package KrscReports_Import
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Import_TableImport
{
    /**
     * value for row, on which occurred error while importing 
     */
    const ROW_WITH_ERROR_VALUE = false;
    
    /**
     * @var integer row number with header for imported table (starts from 1)
     */
    protected $_iRowNumberWithHeader;
    
    /**
     * @var integer first column with table (starts from 0) 
     */
    protected $_iInitialColumnId = 0;
    
    /**
     * @var array subsequent values are subsequent names of columns 
     */
    protected $_aColumnNames;
    
    /**
     * @var string if this symbol is present in column name, that means, it is obligatory 
     */
    protected $_sRequiredSymbol;
    
    /**
     * @var KrscReports_Type_Excel_PHPExcel_Cell cell object
     */
    protected $_oCell;
    
    /**
     * Setter for row number.
     * @param integer $iRowNumberWithHeader row number with header for imported table (starts from 1)
     * @return KrscReports_Import_TableImport object, on which this method was executed
     */
    public function setRowNumberWithHeader( $iRowNumberWithHeader )
    {
        $this->_iRowNumberWithHeader = $iRowNumberWithHeader;
        return $this;
    }
    
    /**
     * Setter for initial column (for header and data).
     * @param integer $iInitialColumnId column number  (starts from 0)
     * @return KrscReports_Import_TableImport object, on which this method was executed
     */
    public function setInitialColumnId( $iInitialColumnId )
    {
        $this->_iInitialColumnId = $iInitialColumnId;
        return $this;
    }
    
    /**
     * Setter for column names.
     * @param array $aColumnNames subsequent values are subsequent names of columns 
     * @return KrscReports_Import_TableImport object, on which this method was executed
     */
    public function setColumnNames( $aColumnNames )
    {
        $this->_aColumnNames = $aColumnNames;
        return $this;
    }
    
    /**
     * Setter for required symbol.
     * @param string $sRequiredSymbol if this symbol is present in column name, that means, it is obligatory (symbol would be removed in output array) 
     * @return KrscReports_Import_TableImport object, on which this method was executed
     */
    public function setRequiredSymbol( $sRequiredSymbol )
    {
        $this->_sRequiredSymbol = $sRequiredSymbol;
        return $this;
    }
    
    /**
     * Loading of file.
     * @param string $fileName path to file
     * @return KrscReports_File instance of class responsible for handling creation of file
     */
    public function loadFile( $fileName )
    {
        $oFile = new \KrscReports_File();
        $oFile->setFileName( $fileName );
        $oFile->setReader();
        
        $this->setCellObject();
        return $oFile;
    }
    
    /**
     * Method for setting cell object. Cell object can be constructed only when PHPExcel object was previously set.
     * @param KrscReports_Type_Excel_PHPExcel_Cell $oCell stub of cell object (by default null)
     * @return KrscReports_Import_TableImport object, on which this method was executed
     */
    public function setCellObject( $oCell = null )
    {
        if( is_null( $oCell ) ) {
            $this->_oCell = new KrscReports_Type_Excel_PHPExcel_Cell();
        } else {
            $this->_oCell = $oCell;
        }
        
        return $this;
    }
    
    /**
     * Method checks, if header row have same columns as declared.
     * @return boolean true if row is correct (otherwise exception is thrown)
     * @throws KrscReports_Exception_Import_InvalidHeader thrown when there is column mismatch
     */
    protected function _checkHeaderRow()
    {
        foreach( $this->_aColumnNames as $iColumnPosition => $sColumnName ) 
        {
            $sCellValue = $this->_oCell->getCellValue( $this->_iInitialColumnId + $iColumnPosition, $this->_iRowNumberWithHeader );
            if( $sCellValue != $sColumnName ){
                $oException = new KrscReports_Exception_Import_InvalidHeader();
                throw $oException->setMessage( $sColumnName, $sCellValue );
            }
        }
            
        return true;
    }
    
    /**
     * Method checking, which columns are obligatory.
     * @return array subsequent values are booleans; true means that this column is required, otherwise false
     */
    protected function _getRequiredColumns()
    {
        $aRequiredColumns = array();
        
        foreach( $this->_aColumnNames as $iColumnPosition => $sColumnName ) {
            $bIsRequired = isset( $this->_sRequiredSymbol ) ? ( stripos( $sColumnName, $this->_sRequiredSymbol ) !== false ) : false;
            $aRequiredColumns[$iColumnPosition] = $bIsRequired;
        }
        
        return $aRequiredColumns;
    }
    
    /**
     * Method importing data.
     * IMPORTANT: FIRST COLUMN IS ALWAYS REQUIRED
     * @return array imported data
     */
    public function getImportedData()
    {
        $iSelectedRowId = $this->_iRowNumberWithHeader + 1;    
        $aRequiredColumns = $this->_getRequiredColumns();
        $this->_checkHeaderRow();
        $aImportedData = array(); 
        
        while( !empty( $this->_oCell->getCellValue( $this->_iInitialColumnId, $iSelectedRowId ) ) ) {
            $mImportedRow = array();
            
            foreach( $this->_aColumnNames as $iColumnPosition => $sColumnName ) {
                $mColumnValue = $this->_oCell->getCellValue( $this->_iInitialColumnId + $iColumnPosition, $iSelectedRowId );
                
                if( empty( $mColumnValue ) && $aRequiredColumns[$iColumnPosition] ) { // column is empty, but is required
                    $mImportedRow = self::ROW_WITH_ERROR_VALUE;
                    break;
                } else {
                    $mImportedRow[$iColumnPosition] = $mColumnValue;
                }
            }
            
            $aImportedData[] = $mImportedRow;
            $iSelectedRowId++;
        }
        
        return $aImportedData;
    }
}