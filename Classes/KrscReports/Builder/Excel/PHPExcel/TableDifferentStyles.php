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
 * Builder responsible for creating table, where each row can have different style set.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_TableDifferentStyles extends KrscReports_Builder_Excel_PHPExcel_TableBasic implements KrscReports_Builder_Interface_Table
{   
    /**
     * style name for row with holiday
     */
    const STYLE_ROW_DIFFERENT = 'row_different';

    /**
     * key in array representing default style for cell
     */
    const STYLE_CELL_DEFAULT = '';
    
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
        foreach( $this->getColumnNames() as $iColumnIndex => $sColumnName )
        {
            if( $sColumnName == self::DATA_STYLE_COLUMN )
            {	/* no column header for style column */
        	continue;
            }
            $this->_oCell->setValue( $sColumnName );
            $this->_oCell->constructCell( $this->_iActualWidth + $iColumnIndex, $this->_iActualHeight );
        }
        
        // adding one row in registry
        $this->_iActualHeight++;
    }
    
    /**
     * Method handling style for given column.
     * @param string $sColumnName name of column being actually analysed
     * @param array $aColumnStyles array for styles in analysed row
     * @return void
     */
    protected function _handleColumnStyle( $sColumnName, $aColumnStyles )
    {
        if( isset( $aColumnStyles[$sColumnName] ) )
        {   // setting style specific for that column
            $this->setStyleKey( $aColumnStyles[$sColumnName] );
        } else if( isset( $aColumnStyles[self::STYLE_CELL_DEFAULT] ))
        {   // setting style specific for that row
            $this->setStyleKey( $aColumnStyles[self::STYLE_CELL_DEFAULT] );
        } else
        {   // setting default style for row (no info about column and default row style)
            $this->setStyleKey( KrscReports_Document_Element_Table::STYLE_ROW );
        }
    }
    
    /**
     * Method extracting informations about styles for cells in a row.
     * @param array|string $mDataStyleColumn if value false - default style; 
     *      otherwise - style key provided by column as string - apply to all cells;
     *      if an array - column name is array key and value - style key;
     *      default style for columns not specified here if an array of styles is provided - under default key ( self::STYLE_CELL_DEFAULT )
     * @return array|string|boolean array with styles for row or name of style for row or false if no customized styles are provided
     */
    protected function _getColumnStyles( $mDataStyleColumn )
    {
        if( $mDataStyleColumn !== false && !is_array( $mDataStyleColumn )  ) {   // one style for whole row
            if( $mDataStyleColumn === true ) {
                $mDataStyleColumn = self::STYLE_ROW_DIFFERENT;
            }
            $this->setStyleKey( $mDataStyleColumn );
            $mColumnStyles = $mDataStyleColumn;
	} else {    // style applied when value is false or element is an array
            $this->setStyleKey( KrscReports_Document_Element_Table::STYLE_ROW );
            $mColumnStyles = false;
	}
                
        if( is_array( $mDataStyleColumn ) )
        {   // array with names of columns as key and values, which are style keys 
            $mColumnStyles = $mDataStyleColumn;
        }
        
        return $mColumnStyles;
    }
    
    /**
     * Action done when table rows has to be displayed (increments height with each row).
     * @return void
     */
    public function setRows() 
    {
        $aMaxSizeSet = array(); // subsequent keys are columnsIds (iterators) with boolean value

        foreach( $this->_aData as $aRow )
        {   // iterating over rows
            $iIterator = 0;
            if( isset( $aRow[self::DATA_STYLE_COLUMN] ) )
            {	
                $aColumnStyles = $this->_getColumnStyles( $aRow[self::DATA_STYLE_COLUMN] ); 
	        unset( $aRow[self::DATA_STYLE_COLUMN] );
            } else {    // for this row we have no info about styles
                unset( $aColumnStyles ); // if exists from previous iteration
                $this->setStyleKey( KrscReports_Document_Element_Table::STYLE_ROW );
            }
            
            $bIsStyleForEachRow = isset( $aColumnStyles ) && is_array($aColumnStyles);
            
            foreach( $aRow as $sColumnName => $mColumnValue )
            {
                if( $bIsStyleForEachRow )
                {
                    $this->_handleColumnStyle( $sColumnName, $aColumnStyles );
                }
                
                $this->_oCell->setValue( $mColumnValue );

                $columnSize = $this->getColumnSize($mColumnValue);
                if ($this->_oCell->isColumnMaxSizeIsSet($iIterator) &&
                    !isset($aMaxSizeSet[$iIterator]) &&
                     $columnSize > $this->_oCell->getColumnMaxSize($iIterator)) {
                        // max size of column reached - fixed size for column set
                        $aMaxSizeSet[$iIterator] = true;
                }

                $this->_oCell->constructCell( $this->_iActualWidth + $iIterator++, $this->_iActualHeight );
            }
            
            // adding one row in registry
            $this->_iActualHeight++;
        }

        $this->setMaxColumnsSizes($aMaxSizeSet);
    }
    
}