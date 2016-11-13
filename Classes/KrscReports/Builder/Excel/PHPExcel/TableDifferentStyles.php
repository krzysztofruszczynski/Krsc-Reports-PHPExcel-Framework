<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2016 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2016 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.2, 2016-10-03
 */

/**
 * Builder responsible for creating table, where each row can have different style set.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_TableDifferentStyles extends KrscReports_Builder_Excel_PHPExcel_ExampleTable implements KrscReports_Builder_Interface_Table
{
    /**
     * style name for row with holiday
     */
    const STYLE_ROW_DIFFERENT = 'row_different';

    /**
     * name of column which has info about style of each row (in this situation: true or false for switching between styles)
     */
    const DATA_STYLE_COLUMN = 'column_style';

    /**
     * Action done when table begins.
     * @return void
     */
    public function beginTable() 
    {
        parent::beginTable();
    }

    /**
     * Action done when table ends.
     * @return void
     */
    public function endTable() 
    {
        parent::endTable();
    }

    /**
     * Action done when table header row with name of the columns has to be displayed (increments height).
     * @return void
     */
    public function setHeaderRow() 
    {
        $iIterator = 0;
        
        foreach( $this->_aData[0] as $sColumnName => $mColumnValue )
        {
            if( $sColumnName == self::DATA_STYLE_COLUMN )
            {	/* no column header for style column */
        	continue;
            }
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
            if( isset( $aRow[self::DATA_STYLE_COLUMN] ) )
            {	// if value false - default style; otherwise - style key provided by column
            	if( $aRow[self::DATA_STYLE_COLUMN] !== false )
	        {
                    $this->setStyleKey( $aRow[self::DATA_STYLE_COLUMN] );
	        } else {
                    $this->setStyleKey( KrscReports_Document_Element_Table::STYLE_ROW );
	        }
	        unset($aRow[self::DATA_STYLE_COLUMN]);
            }
            

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