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
 * @package KrscReports_Report
 * @copyright Copyright (c) 2017 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.1.0, 2017-04-25
 */

/**
 * Builder responsible for highlighting selected rows (can be used while creating or reading and modifying file).
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_Modifier_RowHighlight extends KrscReports_Builder_Excel_PHPExcel
{
    /**
     * @var integer start column for table, which we want to highlight (starts from 0)
     */
    protected $_iInitialColumnId;
    
    /**
     * @var integer start row for table, which we want to highlight (starts from 1)
     */
    protected $_iInitialRowId;
    
    /**
     * @var integer number of columns in table, which have to be highlighted 
     */
    protected $_iNumberOfColumns;
    
    /**
     * Setter for initial column (for header and data).
     * @param integer $iInitialColumnId column number  (starts from 0)
     * @return KrscReports_Builder_Excel_PHPExcel_Modifier_RowHighlight object, on which this method was executed
     */
    public function setInitialColumnId( $iInitialColumnId )
    {
        $this->_iInitialColumnId = $iInitialColumnId;
        return $this;
    }
    
    /**
     * Setter for initial row id with data
     * @param integer $iInitialRowId row number (starts from 1)
     * @return KrscReports_Builder_Excel_PHPExcel_Modifier_RowHighlight object, on which this method was executed
     */
    public function setInitialRowId( $iInitialRowId )
    {
        $this->_iInitialRowId = $iInitialRowId;
        return $this;
    }
    
    /**
     * Setter for number of columns.
     * @param integer $iNumberOfColumns number of columns
     * @return KrscReports_Builder_Excel_PHPExcel_Modifier_RowHighlight object, on which this method was executed
     */
    public function setNumberOfColumns( $iNumberOfColumns )
    {
        $this->_iNumberOfColumns = $iNumberOfColumns;
        return $this;
    }
    
    /**
     * Method setting new style for the whole row.
     * @param integer $iRowIndex index of row, which would be highlighted (starts from 0)
     * @param string $sStyleKey style key for row
     * @return KrscReports_Builder_Excel_PHPExcel_Modifier_RowHighlight object, on which this method was executed
     */
    public function setMarkedRow( $iRowIndex, $sStyleKey = null )
    {
        if ( !is_null( $sStyleKey ) ) {
            $this->setStyleKey( $sStyleKey );
        }
        
        for ( $iColumnNumber = 0 ; $iColumnNumber < $this->_iNumberOfColumns ; $iColumnNumber++ ) {
            $this->_oCell->constructCellStyles( $this->_iInitialColumnId + $iColumnNumber, $this->_iInitialRowId + $iRowIndex );
        }
        
        return $this;
    }
    
    /**
     * Action responsible for creating document (while creating table not used but has to be implemented).
     * @return void
     */
    public function constructDocument()    
    {
        
    }

}

