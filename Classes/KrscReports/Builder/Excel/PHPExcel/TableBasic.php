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
 * @package KrscReports_Builder
 * @copyright Copyright (c) 2018 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.2.5, 2018-04-19
 */

/**
 * Builder responsible for creating example table.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_TableBasic extends KrscReports_Builder_Excel_PHPExcel implements KrscReports_Builder_Interface_Table
{
    /**
     * @var boolean flag if table is filtered (default false)
     */
    protected $_bAutoFilter = false;
    
    /**
     * Method setting auto filter property.
     * @param boolean $bAutoFilter flag if table is filtered (default: true)
     * @return KrscReports_Builder_Excel_PHPExcel_TableDifferentStyles object on which this method was executed
     */
    public function setAutoFilter( $bAutoFilter = true )
    {
        $this->_bAutoFilter = $bAutoFilter;
        return $this;
    }
    
    /**
     * Action done when table begins.
     * @return void
     */
    public function beginTable() 
    {
        $iIterator = 0;        
        
        $aColumnNames = $this->getColumnNames();
        
        if( $this->_bAutoFilter )
        {            
            $this->_oCell->setAutoFilter( $this->_iActualWidth, ($this->_iActualWidth + count( $aColumnNames ) - 1), $this->_iActualHeight );
        }
            
        foreach( $aColumnNames as $sColumnName )
        {
            if(!$this->_oCell->isColumnFixedSizeIsSet($iIterator))
            {
                $this->_oCell->setColumnDimensionAutosize( $this->_iActualWidth + $iIterator, true );
            } else {
                $this->_oCell->createColumnFixedSize($iIterator);
            }

            $iIterator++;
        }
    }

    /**
     * Action done when table ends.
     * @return void
     */
    public function endTable() 
    {
        
    }

    /**
     * Action done when table header row with name of the columns has to be displayed (increments height).
     * @return void
     */
    public function setHeaderRow() 
    {
        $iIterator = 0;
        
        foreach( $this->getColumnNames() as $sColumnName )
        {
            $this->_oCell->setValue( $sColumnName );
            $this->_oCell->constructCell( $this->_iActualWidth + $iIterator++, $this->_iActualHeight );
        }
        
        // adding one row in registry
        $this->_iActualHeight++;
    }

    /**
     * Action done when table rows has to be displayed (increments height with each row).
     * @return void
     */
    public function setRows() 
    {

        foreach( $this->_aData as $aRow )
        {   // iterating over rows
            $iIterator = 0;
            foreach( $aRow as $mColumnValue )
            {
                $this->_oCell->setValue( $mColumnValue );
                $this->_oCell->constructCell( $this->_iActualWidth + $iIterator++, $this->_iActualHeight );
            }

            // adding one row in registry
            $this->_iActualHeight++;
        }
    }

    /**
     * Action responsible for creating document (while creating table not used but has to be implemented).
     * @return void
     */
    public function constructDocument()
    {
        
    }

}
