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
 * @version 1.1.5, 2017-05-15
 */

/**
 * Builder responsible for creating graph.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_Graph_Basic extends KrscReports_Builder_Excel_PHPExcel
{
    /**
     * @var string type of plot 
     */
    protected $_sPlotType;
    
    /**
     * @var string value for grouping for chart (IF SET - MS EXCEL DON'T DISPLAY REPORTS)
     */
    protected $_sGrouping;
    
    /**
     * @var PHPExcel_Chart_Layout object with layout 
     */
    protected $_oLayout;
    
    /**
     * @var PHPExcel_Chart_Title object with graph Title 
     */
    protected $_oTitle;
    
    /**
     * @var array values are objects with chart labels 
     */
    protected $_aPlotLabels = array();
    
    /**
     * @var array values are objects with chart labels  
     */
    protected $_aPlotCategory = array();
    
    /**
     * @var array values are objects with chart labels 
     */
    protected $_aPlotValues = array();
    
    /**
     * @var string cell address, where graph begins 
     */
    protected $_sTopLeftPosition;
    
    /**
     * @var string cell address, where graph ends 
     */
    protected $_sBottomRightPosition;
    
    /**
     * @var integer width of graphs (number of columns) 
     */
    protected $_iGraphColumnSize = 8;
    
    /**
     * @var integer height of graphs (number of rows) 
     */
    protected $_iGraphRowSize = 8;
    
    /**
     * @var integer (if + value, then offset at the beginning, if -, than at the end)
     */
    protected $_iRowOffset = 0;
    
    /**
     * @var integer (if + value, then offset at the beginning, if -, than at the end) - second offset (range can be limited from both sides at the same time)
     */
    protected $_iRowOffsetSecond = 0;
    
    /**
     * @var string name of column with values for plot category (if graph associated with table)
     */
    protected $_sPlotCategoryColumnName;
    
    /**
     * @var string name of column with values for plot values (if graph associated with table)
     */
    protected $_sPlotValuesColumnName;
    
    /**
     * @var string name of group (worksheet), from which data to graph are taken
     */
    protected $_sSourceGroupName;
    
    /**
     * Setter for default settings for graph.
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setDefaultSettings()
    {
        $this->setPlotType();
        $this->setLayout();
        $this->setChartTitle('');
        
        return $this;
    }
    
    /**
     * Setter for plot type.
     * @param string $sPlotType type of plot
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setPlotType( $sPlotType = PHPExcel_Chart_DataSeries::TYPE_PIECHART )
    {
        $this->_sPlotType = $sPlotType;
        return $this;
    }
    
    /**
     * Setter for layout.
     * @param PHPExcel_Chart_Layout $oLayout previously created object with layout (optional)
     * @return PHPExcel_Chart_Layout object with layout set within this object
     */
    public function setLayout( $oLayout = null )
    {
        if ( is_null( $oLayout ) ) {
            $this->_oLayout = new PHPExcel_Chart_Layout();
        } else {
            $this->_oLayout = $oLayout;
        }
        
        return $this->_oLayout;
    }
    
    /**
     * Setter for grouping plot type. IF GROUPING NOT NULL THAN EXCEL HAVE PROBLEMS WITH DISPLAYING IT.
     * @param string $sGrouping type of grouping
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setGrouping( $sGrouping = PHPExcel_Chart_DataSeries::GROUPING_STACKED )
    {
        $this->_sGrouping = $sGrouping;
        return $this;
    }
    
    /**
     * Setter for graph size.
     * @param integer $iSize number of columns and rows for this diagram
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setGraphSize( $iColumnSize, $iRowSize )
    {
        $this->_iGraphColumnSize = $iColumnSize;
        $this->_iGraphRowSize = $iRowSize;
        return $this;
    }
    
    /**
     * Getter for graph column size.
     * @return integer number of columns for this diagram
     */
    public function getGraphColumnSize()
    {
        return $this->_iGraphColumnSize;
    }
    
    /**
     * Getter for graph row size.
     * @return integer number of rows for this diagram
     */
    public function getGraphRowSize()
    {
        return $this->_iGraphRowSize;
    }
    
    /**
     * Setting offset (can be from both sides).
     * @param integer $iRowOffset (if + value, then offset at the beginning, if -, than at the end)
     * @param integer $iRowOffsetSecond (if + value, then offset at the beginning, if -, than at the end) - second offset (range can be limited from both sides at the same time)
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setRowsOffset( $iRowOffset = 0, $iRowOffsetSecond = 0 )
    {
        $this->_iRowOffset = $iRowOffset;
        $this->_iRowOffsetSecond = $iRowOffsetSecond;
        return $this;
    }
    
    /**
     * Setter for name of group (worksheet), from which data to graph are taken.
     * @param string $sSourceGroupName name of group (worksheet), from which data to graph are taken
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setSourceGroupName( $sSourceGroupName )
    {
        $this->_sSourceGroupName = $sSourceGroupName;
        return $this;
    }
    
    /**
     * Method for setting graph position.
     * @param integer $iColumnId numeric coordinate of initial column for graph (starts from 0)
     * @param integer $iRowId numeric coordinate of initial row for graph (starts from 1)
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setGraphPosition( $iColumnId, $iRowId )
    {
        $this->setTopLeftPosition( $iColumnId, $iRowId );
        $this->setBottomRightPosition( $iColumnId + $this->_iGraphColumnSize, $iRowId + $this->_iGraphRowSize );
        return $this;
    }
    
    /**
     * Setter for begin position of graph.
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setTopLeftPosition( $iColumnId, $iRowId )
    {
        $this->_sTopLeftPosition = $this->_oCell->getColumnDimension( $iColumnId ) . $iRowId;
        return $this;
        
    }
    
    /**
     * Setter for end position of graph.
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setBottomRightPosition( $iColumnId, $iRowId )
    {
        $this->_sBottomRightPosition = $this->_oCell->getColumnDimension( $iColumnId ) . $iRowId;
        return $this;
    }
    
    /**
     * Method for setting chart title.
     * @param PHPExcel_Chart_Title|string $mTitle string with title or previously created object
     * @return PHPExcel_Chart_Title actually set title object
     */
    public function setChartTitle( $mTitle = '' )
    {
        if ( is_object( $mTitle ) ) {
            $this->_oTitle = $mTitle;
        } else {
            $this->_oTitle = new PHPExcel_Chart_Title( $mTitle, $this->_oLayout );
        }
        
        return $this->_oTitle;
    }
    
    /**
     * Method setting data series.
     * @param integer $iStartColumnId numeric coordinate of initial column for data series (starts from 0)
     * @param integer $iStartRowId numeric coordinate of initial row for data series (starts from 1)
     * @param integer $iEndColumnId numeric coordinate of end column for data series (starts from 0)
     * @param integer $iEndRowId numeric coordinate of end row for data series (starts from 1)
     * @param string $sType type of data (default: String)
     * @return PHPExcel_Chart_DataSeriesValues data series created
     */
    protected function _createDataSeriesValues( $iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType = 'String' )
    {
        $sDataSourcePattern = "'%s'!$%s$%d:$%s$%d";
        
        if ( $this->_iRowOffset < 0 ) {   // offset at the end of column
            $iEndRowId += $this->_iRowOffset;
        } else if ( $this->_iRowOffset > 0 ) {
            $iStartRowId += $this->_iRowOffset;
        }
        
        if ( $this->_iRowOffsetSecond < 0 ) {   // offset at the end of column
            $iEndRowId += $this->_iRowOffsetSecond;
        } else if ( $this->_iRowOffsetSecond > 0 ) {
            $iStartRowId += $this->_iRowOffsetSecond;
        }
        
        $aPatternArguments = array();
        $aPatternArguments[] = isset( $this->_sSourceGroupName ) ? $this->_sSourceGroupName : $this->_sGroupName;
        $aPatternArguments[] = $this->_oCell->getColumnDimension( $iStartColumnId );
        $aPatternArguments[] = $iStartRowId;
        $aPatternArguments[] = $this->_oCell->getColumnDimension( $iEndColumnId );
        $aPatternArguments[] = $iEndRowId;
        
        $sDataSource = vsprintf( $sDataSourcePattern, $aPatternArguments );
        
        
        if ( $iStartColumnId == $iEndColumnId ) {
            $iNumberOfElements = $iEndRowId - $iStartRowId;
        } else {
            $iNumberOfElements = $iEndColumnId - $iStartColumnId;
        }
        $iNumberOfElements++;
        
        return new PHPExcel_Chart_DataSeriesValues( $sType, $sDataSource, NULL, $iNumberOfElements );
    }
    
    /**
     * Method setting labels for graph.
     * @param integer $iStartColumnId numeric coordinate of initial column for labels (starts from 0)
     * @param integer $iStartRowId numeric coordinate of initial row for labels (starts from 1)
     * @param integer $iEndColumnId numeric coordinate of end column for labels (starts from 0)
     * @param integer $iEndRowId numeric coordinate of end row for labels (starts from 1)
     * @param string $sType type of data (default: String)
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setPlotLabels( $iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType = 'String' )
    {
        $this->_aPlotLabels[] = $this->_createDataSeriesValues( $iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType );
        return $this;
    }
    
    /**
     * Method returning previously set plot labels.
     * @return array subsequent elements are data series
     */
    public function getPlotLabels()
    {
        return $this->_aPlotLabels;
    }
    
    /**
     * Setter for column name, from which values for plot category will be taken.
     * @param string $sPlotCategoryColumnName name of column for plot category
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setPlotCategoryColumnName( $sPlotCategoryColumnName )
    {
        $this->_sPlotCategoryColumnName = $sPlotCategoryColumnName;
        return $this;
    }
    
    /**
     * Getter for plot category column name.
     * @return string name of column for plot category
     */
    public function getPlotCategoryColumnName()
    {
        return $this->_sPlotCategoryColumnName;
    }
    
    /**
     * Method setting category (x axis) for graph.
     * @param integer $iStartColumnId numeric coordinate of initial column for category (starts from 0)
     * @param integer $iStartRowId numeric coordinate of initial row for category (starts from 1)
     * @param integer $iEndColumnId numeric coordinate of end column for category (starts from 0)
     * @param integer $iEndRowId numeric coordinate of end row for category (starts from 1)
     * @param string $sType type of data (default: String)
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setPlotCategory( $iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType = 'String' )
    {
        $this->_aPlotCategory[] = $this->_createDataSeriesValues( $iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType );
        return $this;
    }
    
    /**
     * Setter for column name, from which values for plot values will be taken.
     * @param string $sPlotValuesColumnName name of column for plot values
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setPlotValuesColumnName( $sPlotValuesColumnName )
    {
        $this->_sPlotValuesColumnName = $sPlotValuesColumnName;
        return $this;
    }
    
    /**
     * Getter for plot values column name.
     * @return string name of column for plot value
     */
    public function getPlotValuesColumnName()
    {
        return $this->_sPlotValuesColumnName;
    }
    
    /**
     * Method setting values (y axis) for graph.
     * @param integer $iStartColumnId numeric coordinate of initial column for values (starts from 0)
     * @param integer $iStartRowId numeric coordinate of initial row for values (starts from 1)
     * @param integer $iEndColumnId numeric coordinate of end column for values (starts from 0)
     * @param integer $iEndRowId numeric coordinate of end row for values (starts from 1)
     * @param string $sType type of data (default: Number)
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setPlotValues( $iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType = 'Number' )
    {
        $this->_aPlotValues[] = $this->_createDataSeriesValues( $iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType );
        return $this;
    }
    
    /**
     * Method creating and returning chart object (need to be attached to worksheet to be visible).
     * @return PHPExcel_Chart chart object
     */
    public function getChartObject()
    {
        $oSeries = new PHPExcel_Chart_DataSeries(
            $this->_sPlotType,
            $this->_sGrouping, // IF GROUPING NOT NULL THAN EXCEL HAVE PROBLEMS WITH DISPLAYING IT
            range(0, count( $this->_aPlotValues )-1),
            $this->_aPlotLabels,
            $this->_aPlotCategory,
            $this->_aPlotValues,
            PHPExcel_Chart_DataSeries::DIRECTION_VERTICAL
        );
        
        $oPlotarea = new PHPExcel_Chart_PlotArea( $this->_oLayout, array( $oSeries ) );
        $oLegend = new PHPExcel_Chart_Legend( PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false );
        
        $oChart = new PHPExcel_Chart(
            'chart1',
             $this->_oTitle,
             $oLegend,
             $oPlotarea,
             true,
             0,
             NULL,
             NULL
        );
        
        $oChart->setTopLeftPosition( $this->_sTopLeftPosition );
        $oChart->setBottomRightPosition( $this->_sBottomRightPosition );
        
        return $oChart;
    }
    
    /**
     * Method creates and attach chart to worksheet.
     */
    public function attachChart()
    {
        $oChart = $this->getChartObject();
        KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject()->getActiveSheet()->addChart( $oChart );
    }
    
    /**
     * Action responsible for creating document (while creating table not used but has to be implemented).
     * @return void
     */
    public function constructDocument()    
    {
        
    }

}

