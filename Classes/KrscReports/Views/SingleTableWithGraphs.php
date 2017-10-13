<?php
namespace KrscReports\Views;

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
 * @version 1.2.0, 2017-10-13
 */

/**
 * Report consisting of one single table and graphs associated with it (using data from this table as source).
 * 
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class SingleTableWithGraphs extends SingleTable
{
    /**
     * key in options with array of graphs associated with this table (every graph has an array with options)
     */
    const KEY_GRAPHS = 'graphs';    
    
    /**
     * key for graph title inside graph options
     */
    const KEY_GRAPHS_TITLE = 'graphs_title';
    
    /**
     * key for category column name (x axis) inside graph options
     */
    const KEY_GRAPHS_COLUMN_CATEGORY = 'plot_category';
    
    /**
     * key for values column name (y axis) inside graph options
     */
    const KEY_GRAPHS_COLUMN_VALUES = 'plot_values';
    
    /**
     * size of graph in columns inside graph options
     */
    const KEY_GRAPHS_COLUMN_SIZE = 'column_size';
    
    /**
     * size of graph in rows inside graph options
     */
    const KEY_GRAPHS_ROW_SIZE = 'row_size';
    
    /**
     * type of graph inside graph options
     */
    const KEY_GRAPHS_PLOT_TYPE = 'plot_type';
    
    /**
     * offset for rows (if + value, then offset at the beginning, if -, than at the end) inside graph options
     */
    const KEY_GRAPHS_ROWS_OFFSET = 'offset_rows';
    
    /**
     * second offset for rows (if + value, then offset at the beginning, if -, than at the end) inside graph options (both can be used at the same time)
     */
    const KEY_GRAPHS_ROWS_OFFSET_SECOND = 'offset_rows_second';
    
    /**
     * colors of graph inside graph options (string or array when more than one color)
     */
    const KEY_GRAPHS_COLORS = 'colors';
    
    /**
     * Method creating builder for table with graphs.
     * @param string $spreadsheetName spreadsheet name (for setting column widths in proper spreadsheet)
     * @return \KrscReports_Builder_Excel_PHPExcel_TableCurrent builder for report
     */
    protected function getSingleTableWithGraphsBuilder( $spreadsheetName )
    {
        $builder = $this->getSingleTableBuilder( $spreadsheetName );
        
        foreach ( $this->options[self::KEY_GRAPHS] as $graphOptions ){
            $graph = $builder->addNewGraph();

            $graph->setPlotCategoryColumnName( $graphOptions[self::KEY_GRAPHS_COLUMN_CATEGORY] ); 
            $graph->setPlotValuesColumnName( $graphOptions[self::KEY_GRAPHS_COLUMN_VALUES] );            
            $graph->setGraphSize( $graphOptions[self::KEY_GRAPHS_COLUMN_SIZE], $graphOptions[self::KEY_GRAPHS_ROW_SIZE] );
            $graph->setDefaultSettings();
            $graph->setLayout()->setShowPercent( true );
            
            if (isset($graphOptions[self::KEY_GRAPHS_COLORS])) {               
                $graph->setFillColor($graphOptions[self::KEY_GRAPHS_COLORS]); 
            }
            
            if ( isset( $graphOptions[self::KEY_GRAPHS_TITLE] ) ) {
                $graph->setChartTitle( $graphOptions[self::KEY_GRAPHS_TITLE] );
            }
            
            if ( isset( $graphOptions[self::KEY_GRAPHS_ROWS_OFFSET] ) ) {
                $graph->setRowsOffset( $graphOptions[self::KEY_GRAPHS_ROWS_OFFSET], 
                        ( isset( $graphOptions[self::KEY_GRAPHS_ROWS_OFFSET_SECOND] ) ? $graphOptions[self::KEY_GRAPHS_ROWS_OFFSET_SECOND] : 0 ) 
                );
            } 
            
            if ( isset( $graphOptions[self::KEY_GRAPHS_PLOT_TYPE] ) ) {
                $graph->setPlotType( $graphOptions[self::KEY_GRAPHS_PLOT_TYPE] );
            }                        
        }
        
        return $builder;
    }
    
    /**
     * Method generating report.
     */
    public function getDocumentElement( $spreadsheetName = 'Results' )
    {
        $builder = $this->getSingleTableWithGraphsBuilder( $spreadsheetName );
        $this->setTableElement($builder, $spreadsheetName);

        return $this->documentElement;
    }

    public function getDescription()
    {
        return 'Single table view with graphs.';
    }
}

