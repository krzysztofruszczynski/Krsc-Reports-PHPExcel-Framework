<?php
namespace KrscReports\Builder\Excel;

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
 * Class which object represents basic functionality for graphs. Have to be extended by specific tool (like PHPExcel or PhpSpreadsheet).
 *
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class GraphBasic extends \KrscReports_Builder_Excel
{
    /**
     * @var array|string color of graph (optional, not supported by PHPExcel 1.8.1, supported by https://github.com/krzysztofruszczynski/PHPExcel.git (1.8.2))
     */
    protected $_mFillColor = null;

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
     * @var integer iterator for knowing, which data series for plot values is actually being set 
     */
    protected $_iPlotValuesIterator = 0;

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
     * Setter for graph size.
     *
     * @param integer $iSize number of columns and rows for this diagram
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setGraphSize($iColumnSize, $iRowSize)
    {
        $this->_iGraphColumnSize = $iColumnSize;
        $this->_iGraphRowSize = $iRowSize;

        return $this;
    }

    /**
     * Getter for graph column size.
     *
     * @return integer number of columns for this diagram
     */
    public function getGraphColumnSize()
    {
        return $this->_iGraphColumnSize;
    }

    /**
     * Getter for graph row size.
     *
     * @return integer number of rows for this diagram
     */
    public function getGraphRowSize()
    {
        return $this->_iGraphRowSize;
    }

    /**
     * Setting offset (can be from both sides).
     *
     * @param integer $iRowOffset (if + value, then offset at the beginning, if -, than at the end)
     * @param integer $iRowOffsetSecond (if + value, then offset at the beginning, if -, than at the end) - second offset (range can be limited from both sides at the same time)
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setRowsOffset($iRowOffset = 0, $iRowOffsetSecond = 0)
    {
        $this->_iRowOffset = $iRowOffset;
        $this->_iRowOffsetSecond = $iRowOffsetSecond;

        return $this;
    }

    /**
     * Setter for name of group (worksheet), from which data to graph are taken.
     *
     * @param string $sSourceGroupName name of group (worksheet), from which data to graph are taken
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setSourceGroupName($sSourceGroupName)
    {
        $this->_sSourceGroupName = $sSourceGroupName;

        return $this;
    }

    /**
     * Method for setting graph position.
     *
     * @param integer $iColumnId numeric coordinate of initial column for graph (starts from 0)
     * @param integer $iRowId numeric coordinate of initial row for graph (starts from 1)
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setGraphPosition($iColumnId, $iRowId)
    {
        $this->setTopLeftPosition($iColumnId, $iRowId);
        $this->setBottomRightPosition($iColumnId + $this->_iGraphColumnSize, $iRowId + $this->_iGraphRowSize);

        return $this;
    }

    /**
     * Setter for begin position of graph.
     *
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setTopLeftPosition($iColumnId, $iRowId)
    {
        $this->_sTopLeftPosition = $this->_oCell->getColumnDimension($iColumnId) . $iRowId;

        return $this;
    }

    /**
     * Setter for end position of graph.
     *
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setBottomRightPosition($iColumnId, $iRowId)
    {
        $this->_sBottomRightPosition = $this->_oCell->getColumnDimension($iColumnId) . $iRowId;
        return $this;
    }

    /**
     * Method setting data series.
     *
     * @param integer $iStartColumnId numeric coordinate of initial column for data series (starts from 0)
     * @param integer $iStartRowId numeric coordinate of initial row for data series (starts from 1)
     * @param integer $iEndColumnId numeric coordinate of end column for data series (starts from 0)
     * @param integer $iEndRowId numeric coordinate of end row for data series (starts from 1)
     * @param string $sType type of data (default: String)
     * @param string|array $mFillColor color of data series (optional, not supported by PHPExcel 1.8.1, supported by https://github.com/krzysztofruszczynski/PHPExcel.git (1.8.2)))
     *
     *  @return object
     */
    abstract protected function _createDataSeriesValues($iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType = 'String', $mFillColor = null);

    /**
     * Method setting labels for graph.
     *
     * @param integer $iStartColumnId numeric coordinate of initial column for labels (starts from 0)
     * @param integer $iStartRowId numeric coordinate of initial row for labels (starts from 1)
     * @param integer $iEndColumnId numeric coordinate of end column for labels (starts from 0)
     * @param integer $iEndRowId numeric coordinate of end row for labels (starts from 1)
     * @param string $sType type of data (default: String)
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setPlotLabels($iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType = 'String')
    {
        $this->_aPlotLabels[] = $this->_createDataSeriesValues($iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType);

        return $this;
    }

    /**
     * Method returning previously set plot labels.
     *
     * @return array subsequent elements are data series
     */
    public function getPlotLabels()
    {
        return $this->_aPlotLabels;
    }

    /**
     * Setter for column name, from which values for plot category will be taken.
     *
     * @param string $sPlotCategoryColumnName name of column for plot category
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setPlotCategoryColumnName($sPlotCategoryColumnName)
    {
        $this->_sPlotCategoryColumnName = $sPlotCategoryColumnName;

        return $this;
    }

    /**
     * Getter for plot category column name.
     *
     * @return string name of column for plot category
     */
    public function getPlotCategoryColumnName()
    {
        return $this->_sPlotCategoryColumnName;
    }

    /**
     * Method setting category (x axis) for graph.
     *
     * @param integer $iStartColumnId numeric coordinate of initial column for category (starts from 0)
     * @param integer $iStartRowId numeric coordinate of initial row for category (starts from 1)
     * @param integer $iEndColumnId numeric coordinate of end column for category (starts from 0)
     * @param integer $iEndRowId numeric coordinate of end row for category (starts from 1)
     * @param string $sType type of data (default: String)
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setPlotCategory($iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType = 'String')
    {
        $this->_aPlotCategory[] = $this->_createDataSeriesValues($iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType);

        return $this;
    }

    /**
     * Setter for column name, from which values for plot values will be taken.
     *
     * @param string $sPlotValuesColumnName name of column for plot values
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setPlotValuesColumnName($sPlotValuesColumnName)
    {
        $this->_sPlotValuesColumnName = $sPlotValuesColumnName;

        return $this;
    }

    /**
     * Getter for plot values column name.
     *
     * @return string name of column for plot value
     */
    public function getPlotValuesColumnName()
    {
        return $this->_sPlotValuesColumnName;
    }

    /**
     * Set fill color. Not supported by PHPExcel 1.8.1 . Supported by https://github.com/krzysztofruszczynski/PHPExcel.git (1.8.2)).
     *
     * @param array|string $mFillColor one color for whole graph or array with colors (subsequent values of $mFillColor are data for subsequent data series - on color or array with colors
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setFillColor($mFillColor)
    {
        $this->_mFillColor = $mFillColor;

        return $this;
    }

    /**
     * Method for getting filled color. Method should be used only once for one series of data (iterator inside method).
     * 
     * @return array|string array with colors or string with color
     */
    protected function _getFillColor()
    {
        if (is_null($this->_mFillColor) || !is_array($this->_mFillColor)) {
            return $this->_mFillColor; // one color for whole graph or not set
        } else {    // array with colors (subsequent values of $this->_mFillColor are data for subsequent data series - on color or array with colors
            if (isset($this->_mFillColor[$this->_iPlotValuesIterator])) {
                return $this->_mFillColor[$this->_iPlotValuesIterator++];
            } else {
                return null;
            }
        }
    }

    /**
     * Method setting values (y axis) for graph.
     *
     * @param integer $iStartColumnId numeric coordinate of initial column for values (starts from 0)
     * @param integer $iStartRowId numeric coordinate of initial row for values (starts from 1)
     * @param integer $iEndColumnId numeric coordinate of end column for values (starts from 0)
     * @param integer $iEndRowId numeric coordinate of end row for values (starts from 1)
     * @param string $sType type of data (default: Number)
     *
     * @return \KrscReports\Builder\Excel\GraphBasic object, on which method was executed
     */
    public function setPlotValues($iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType = 'Number')
    {
        $this->_aPlotValues[] = $this->_createDataSeriesValues($iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType, $this->_getFillColor());

        return $this;
    }

    /**
     * Method creates and attach chart to worksheet.
     */
    public function attachChart()
    {
        $oChart = $this->getChartObject();
        $this->_oCell->attachChart($oChart);
    }

    /**
     * Action responsible for creating document (while creating table not used but has to be implemented).
     * @return void
     */
    public function constructDocument()
    {

    }
}
