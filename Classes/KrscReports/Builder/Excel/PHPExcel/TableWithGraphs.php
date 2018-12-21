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
 * @version 2.0.1, 2018-11-08
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
     * @var integer number of columns between subsequent graphs (and table and first graph)
     */
    protected $_iNumberOfColumnsBetweenGraphs = 1;
    
    /**
     * @var string suffix for worksheet name, in which charts would be showed (if not set - than showed in the same worksheet)
     */
    protected $_sChartWorksheetSuffix;
    
    /**
     * Setter for number of columns between subsequent graphs.
     * @param integer $iNumberOfColumnsBetweenGraphs number of columns between subsequent graphs (and table and first graph)
     * @return KrscReports_Builder_Excel_PHPExcel_TableWithGraphs object, on which this method was executed
     */
    public function setNumberOfColumnsBetweenGraphs( $iNumberOfColumnsBetweenGraphs )
    {
        $this->_iNumberOfColumnsBetweenGraphs = $iNumberOfColumnsBetweenGraphs;
        return $this;
    }
    
    /**
     * Setter for chart worksheet suffix.
     * @param string $sChartWorksheetSuffix suffix for worksheet name, in which charts would be showed
     * @return KrscReports_Builder_Excel_PHPExcel_TableWithGraphs object, on which this method was executed
     */
    public function setChartWorksheetSuffix( $sChartWorksheetSuffix )
    {
        $this->_sChartWorksheetSuffix = $sChartWorksheetSuffix;
        return $this;
    }
    
    /**
     * Method adding new graph associated with table.
     *
     * @return \KrscReports_Builder_Excel_PHPExcel_Graph_Basic|\KrscReports\Builder\Excel\PhpSpreadsheet\Graph\Basic new graph (on which settings can be made)
     */
    public function addNewGraph()
    {
        switch (\KrscReports_File::getBuilderType()) {
            case \KrscReports_File::SETTINGS_PHPEXCEL:
                $oNewGraph = new \KrscReports_Builder_Excel_PHPExcel_Graph_Basic();
                break;
            case \KrscReports_File::SETTINGS_PHPSPREADSHEET:
                $oNewGraph = new \KrscReports\Builder\Excel\PhpSpreadsheet\Graph\Basic();
                break;
        }

        $oNewGraph->setCellObject($this->_oCell);
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

        $iColumnSizeOfGraphs = 0;
        if ( isset( $this->_aGraphs ) ) {
            foreach ( $this->_aGraphs as $oGraph ) {// action made for every graph
                /* @var $oGraph \KrscReports\Builder\Excel\GraphBasic */
                $iNumberOfColumns = count( $this->_aData[0] );
                
                
                $iPlotCategoryColumnIndex = array_search( $oGraph->getPlotCategoryColumnName(), $this->getColumnNames() );
                $iPlotValuesColumnIndex = array_search( $oGraph->getPlotValuesColumnName(), $this->getColumnNames() );
                
                $oGraph->setGroupName( $this->_sGroupName . ( isset( $this->_sChartWorksheetSuffix ) ? $this->_sChartWorksheetSuffix : '' ) );
                $oGraph->setSourceGroupName( $this->_sGroupName );
                
                $aPlotLabels = $oGraph->getPlotLabels();
                
                if ( empty( $aPlotLabels ) ) {  // labels were not previously set
                    $oGraph->setPlotLabels( $this->_iActualWidth + $iPlotValuesColumnIndex, $this->_iStartHeightOfRows - 1, $this->_iActualWidth + $iPlotValuesColumnIndex, $this->_iStartHeightOfRows - 1 );
                }
                
                $oGraph->setPlotCategory( $this->_iActualWidth + $iPlotCategoryColumnIndex, $this->_iStartHeightOfRows, $this->_iActualWidth + $iPlotCategoryColumnIndex, $this->_iStartHeightOfRows + count( $this->_aData ) - 1 );
                $oGraph->setPlotValues( $this->_iActualWidth + $iPlotValuesColumnIndex, $this->_iStartHeightOfRows, $this->_iActualWidth + $iPlotValuesColumnIndex, $this->_iStartHeightOfRows + count( $this->_aData ) - 1  );
                
                if ( isset( $this->_sChartWorksheetSuffix ) ) {
                    $oGraph->setGraphPosition( $this->_iActualWidth + $iColumnSizeOfGraphs, $this->_iStartHeightOfRows - 1 );
                } else {
                    $oGraph->setGraphPosition( $this->_iActualWidth + $iColumnSizeOfGraphs + $iNumberOfColumns + $this->_iNumberOfColumnsBetweenGraphs, $this->_iStartHeightOfRows - 1 );
                }
                     
                $iColumnSizeOfGraphs += $oGraph->getGraphColumnSize() + $this->_iNumberOfColumnsBetweenGraphs;                
                $oGraph->attachChart();
            }
        }
    }
}
