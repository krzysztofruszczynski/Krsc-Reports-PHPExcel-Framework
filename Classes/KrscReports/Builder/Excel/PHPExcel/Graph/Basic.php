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
 * Builder responsible for creating graph with PHPExcel.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_Graph_Basic extends KrscReports\Builder\Excel\GraphBasic
{
    /**
     * Setter for default settings for graph.
     *
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
     *
     * @param string $sPlotType type of plot
     *
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setPlotType($sPlotType = PHPExcel_Chart_DataSeries::TYPE_PIECHART)
    {
        $this->_sPlotType = $sPlotType;

        return $this;
    }

    /**
     * Setter for layout.
     *
     * @param PHPExcel_Chart_Layout $oLayout previously created object with layout (optional)
     *
     * @return PHPExcel_Chart_Layout object with layout set within this object
     */
    public function setLayout($oLayout = null)
    {
        if (is_null($oLayout)) {
            $this->_oLayout = new PHPExcel_Chart_Layout();
        } else {
            $this->_oLayout = $oLayout;
        }

        return $this->_oLayout;
    }

    /**
     * Setter for grouping plot type. IF GROUPING NOT NULL THAN EXCEL HAVE PROBLEMS WITH DISPLAYING IT.
     *
     * @param string $sGrouping type of grouping
     *
     * @return KrscReports_Builder_Excel_PHPExcel_Graph_Basic object, on which method was executed
     */
    public function setGrouping($sGrouping = PHPExcel_Chart_DataSeries::GROUPING_STACKED)
    {
        $this->_sGrouping = $sGrouping;

        return $this;
    }

    /**
     * Method for setting chart title.
     *
     * @param PHPExcel_Chart_Title|string $mTitle string with title or previously created object
     *
     * @return PHPExcel_Chart_Title actually set title object
     */
    public function setChartTitle($mTitle = '')
    {
        if (is_object($mTitle)) {
            $this->_oTitle = $mTitle;
        } else {
            $this->_oTitle = new \PHPExcel_Chart_Title($mTitle, $this->_oLayout);
        }

        return $this->_oTitle;
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
     * @return PHPExcel_Chart_DataSeriesValues data series created
     */
    protected function _createDataSeriesValues($iStartColumnId, $iStartRowId, $iEndColumnId, $iEndRowId, $sType = 'String', $mFillColor = null)
    {
        $sDataSourcePattern = "'%s'!$%s$%d:$%s$%d";

        if ($this->_iRowOffset < 0) {   // offset at the end of column
            $iEndRowId += $this->_iRowOffset;
        } else if ($this->_iRowOffset > 0) {
            $iStartRowId += $this->_iRowOffset;
        }

        if ($this->_iRowOffsetSecond < 0) {   // offset at the end of column
            $iEndRowId += $this->_iRowOffsetSecond;
        } else if ($this->_iRowOffsetSecond > 0) {
            $iStartRowId += $this->_iRowOffsetSecond;
        }

        $aPatternArguments = array();
        $aPatternArguments[] = isset($this->_sSourceGroupName) ? $this->_sSourceGroupName : $this->_sGroupName;
        $aPatternArguments[] = $this->_oCell->getColumnDimension($iStartColumnId);
        $aPatternArguments[] = $iStartRowId;
        $aPatternArguments[] = $this->_oCell->getColumnDimension($iEndColumnId);
        $aPatternArguments[] = $iEndRowId;

        $sDataSource = vsprintf($sDataSourcePattern, $aPatternArguments);

        if ($iStartColumnId == $iEndColumnId) {
            $iNumberOfElements = $iEndRowId - $iStartRowId;
        } else {
            $iNumberOfElements = $iEndColumnId - $iStartColumnId;
        }
        $iNumberOfElements++;

        if (is_null($mFillColor)) { // version compatible with PHPExcel
            return new \PHPExcel_Chart_DataSeriesValues($sType, $sDataSource, NULL, $iNumberOfElements);
        } else {// version not compatible with PHPExcel (supported by https://github.com/krzysztofruszczynski/PHPExcel.git (1.8.2))
            return new \PHPExcel_Chart_DataSeriesValues($sType, $sDataSource, NULL, $iNumberOfElements, null, null, $mFillColor);
        }
    }

    /**
     * Method creating and returning chart object (need to be attached to worksheet to be visible).
     *
     * @return \PHPExcel_Chart chart object
     */
    public function getChartObject()
    {
        $oSeries = new \PHPExcel_Chart_DataSeries(
            $this->_sPlotType,
            $this->_sGrouping, // IF GROUPING NOT NULL THAN EXCEL HAVE PROBLEMS WITH DISPLAYING IT
            range(0, count($this->_aPlotValues)-1),
            $this->_aPlotLabels,
            $this->_aPlotCategory,
            $this->_aPlotValues,
            PHPExcel_Chart_DataSeries::DIRECTION_VERTICAL
        );

        $oPlotarea = new \PHPExcel_Chart_PlotArea($this->_oLayout, array($oSeries));
        $oLegend = new \PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);

        $oChart = new \PHPExcel_Chart(
            'chart1',
             $this->_oTitle,
             $oLegend,
             $oPlotarea,
             true,
             0,
             NULL,
             NULL
        );

        $oChart->setTopLeftPosition($this->_sTopLeftPosition);
        $oChart->setBottomRightPosition($this->_sBottomRightPosition);

        return $oChart;
    }
}
