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
 * @package KrscReports_Builder
 * @copyright Copyright (c) 2017 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.1.2, 2017-05-05
 */

/**
 * Table with possibility to add graphs based on data showed in table (extension from normal table builder). 
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_TableWithGraphs extends KrscReports_Builder_Excel_PHPExcel_TableWithSums
{
    /**
     * @var array subsequent values are graphs objects 
     */
    protected $_aGraphs;
    
    /**
     * Method adding new graph associated with table.
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic new graph (on which settings can be made)
     */
    public function addNewGraph()
    {
        $oNewGraph = new KrscReports_Builder_Excel_PHPExcel_Graph_Basic();
        $this->_aGraphs[] = $oNewGraph;
        return $oNewGraph;
    }
    
    /**
     * Action done when table ends. If graphs are set, then here they are constructed.
     * @return void
     */
    public function endTable() 
    {
        parent::endTable();
        
        $iSizeOfGraphs = 0;
        if ( isset( $this->_aGraphs ) ) {
            foreach ( $this->_aGraphs as $oGraph ) {    // action made for every graph
                /* @var $oGraph KrscReports_Builder_Excel_PHPExcel_GraphBasic */
                $iNumberOfColumns = count( $this->_aData[0] );
                
                
                $iPlotCategoryColumnIndex = array_search( $oGraph->getPlotCategoryColumnName(), $this->getColumnNames() );
                $iPlotValuesColumnIndex = array_search( $oGraph->getPlotValuesColumnName(), $this->getColumnNames() );
                
                $oGraph->setGroupName( $this->_sGroupName );
                $aPlotLabels = $oGraph->getPlotLabels();
                
                if ( empty( $aPlotLabels ) ) {  // labels were not previously set
                    $oGraph->setPlotLabels( $this->_iActualWidth + $iPlotValuesColumnIndex, $this->_iStartHeightOfRows - 1, $this->_iActualWidth + $iPlotValuesColumnIndex, $this->_iStartHeightOfRows - 1 );
                }
                
                $oGraph->setPlotCategory( $this->_iActualWidth + $iPlotCategoryColumnIndex, $this->_iStartHeightOfRows, $this->_iActualWidth + $iPlotCategoryColumnIndex, $this->_iStartHeightOfRows + count( $this->_aData ) - 1 );
                $oGraph->setPlotValues( $this->_iActualWidth + $iPlotValuesColumnIndex, $this->_iStartHeightOfRows, $this->_iActualWidth + $iPlotValuesColumnIndex, $this->_iStartHeightOfRows + count( $this->_aData ) - 1  );
                
                $oGraph->setGraphPosition( $this->_iActualWidth + $iSizeOfGraphs + $iNumberOfColumns + 2, $this->_iStartHeightOfRows );
                
                $iSizeOfGraphs += $oGraph->getGraphSize() + 2;                
                $oGraph->attachChart();
            }
        }        
    }
}
