<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2014 Krzysztof Ruszczyński
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
 * @package KrscReports_Builder
 * @copyright Copyright (c) 2014 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.0, 2014-12-28
 */

/**
 * Table with possibility to add column sum at the end (extension from normal table builder). 
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_TableWithSums extends KrscReports_Builder_Excel_PHPExcel_ExampleTable implements KrscReports_Builder_Interface_Table
{
    /**
     * @var array columns which are to sum (each column name is next array value
     */
    protected $_aColumnsToSum;
    
    /**
     * @var integer height of first row with data (to be used by formulas) 
     */
    protected $_iStartHeightOfRows;
    
    /**
     * Method for adding columns, which will have sum formula under rows.
     * @param string $sColumnName name of the column which will have sum formula under table rows
     * @return KrscReports_Builder_Excel_PHPExcel_TableWithSums object on which method was executed
     */
    public function addColumnToSum( $sColumnName )
    {
        $this->_aColumnsToSum[] = $sColumnName;
        return $this;
    }
    
    /**
     * Action done when table rows has to be displayed (increments height with each row, remembers height of first row with data).
     * @return void
     */
    public function setRows() 
    {
        $this->_iStartHeightOfRows = $this->_iActualHeight;
        parent::setRows();
    }
    
    /**
     * Action done when table ends. If columns for sum formula have been added, it creates a field under the table with sum formula
     * (under particular column).
     * @return void
     */
    public function endTable() 
    {
        parent::endTable();
        
        
        foreach( $this->_aColumnsToSum as $sColumnName )
        {   // add sum formula for each column
            // position of summed column
            $iPosition = array_search( $sColumnName, array_keys( $this->_aData[0] ) );
            
            $sColumnCoordinate = $this->_oCell->getColumnDimension( $iPosition );
            $sFormula = sprintf( '=SUM(%s%d:%s%d)', $sColumnCoordinate, $this->_iStartHeightOfRows, $sColumnCoordinate, ($this->_iActualHeight-1) );
            
            $this->_oCell->setValue( $sFormula );
            $this->_oCell->constructCell( $this->_iActualWidth +  $iPosition, $this->_iActualHeight );
        }
        
        $this->_iActualHeight++;
    }
}
    
