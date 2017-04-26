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
 * @package KrscReports_Report
 * @copyright Copyright (c) 2017 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.1.0, 2017-04-20
 */

/**
 * Example report creating chart pie in PHPExcel.
 *  
 * @category KrscReports
 * @package KrscReports_Report
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Report_ExampleReportChartPie extends KrscReports_Report_ExampleReport
{
    /**
     * title of column with product name
     */
    const COLUMN_PRODUCT_NAME = 'Product name';
    
    /**
     * title of column with product price
     */
    const COLUMN_PRODUCT_PRICE = 'Product price (in EUR)';
    
    /**
     * title of column with product availibility
     */
    const COLUMN_PRODUCT_AVAILIBILITY = 'Product availibility';
    
    /**
     * value for availibility, when product is not available
     */
    const AVAILIBILITY_LEVEL_NOT_AVAILABLE = 0;
    
    /**
     * value for availibility, when product supply is very low
     */
    const AVAILIBILITY_LEVEL_VERY_LOW = 1;
    
    /**
     * value for availibility, when product supply is low
     */
    const AVAILIBILITY_LEVEL_LOW = 2;
    
    /**
     * value for availibility, when product supply is medium
     */
    const AVAILIBILITY_LEVEL_MEDIUM = 3;
    
    /**
     * value for availibility, when product supply is high
     */
    const AVAILIBILITY_LEVEL_HIGH = 4;
    
    /**
     * @var array key is availibility value which stores title to be used in spreadsheet 
     */
    protected $_aAvailibilityTexts = array( 
        self::AVAILIBILITY_LEVEL_NOT_AVAILABLE => 'Not Available',
        self::AVAILIBILITY_LEVEL_VERY_LOW => 'Very low',
        self::AVAILIBILITY_LEVEL_LOW => 'Low',
        self::AVAILIBILITY_LEVEL_MEDIUM => 'Medium',
        self::AVAILIBILITY_LEVEL_HIGH => 'High'
    );
    
    /**
     * @var array data for spreadsheet 
     */
    protected $_aData = array();
    
    /**
     * Method adding one row of data.
     * @param string $sProductName name of product in a row
     * @param string $sProductPrice prive of product in a row
     * @param integer $iProductAvailibility values from const from that class with prefix AVAILIBILITY_LEVEL
     * @return KrscReports_Report_ExampleReportDifferentCellStyles object, on which method was executed
     */
    protected function _addRow( $sProductName, $sProductPrice, $iProductAvailibility )
    {
        $aRow = array();
        $aRow[self::COLUMN_PRODUCT_NAME] = $sProductName;
        $aRow[self::COLUMN_PRODUCT_PRICE] = $sProductPrice;
        $aRow[self::COLUMN_PRODUCT_AVAILIBILITY] = $this->_aAvailibilityTexts[$iProductAvailibility];
        $this->_aData[] = $aRow;
        return $this;
    }
    
    /**
     * Method setting data for spreadsheet.
     * @return void
     */
    protected function _setData()
    {   
        $this->_addRow( 'Laptop', '800', self::AVAILIBILITY_LEVEL_HIGH )
        ->_addRow( 'Tablet', '200', self::AVAILIBILITY_LEVEL_MEDIUM )
        ->_addRow( 'Smartphone', '300', self::AVAILIBILITY_LEVEL_VERY_LOW )
        ->_addRow( 'Desktop Computer', '600', self::AVAILIBILITY_LEVEL_NOT_AVAILABLE )
        ->_addRow( 'USB Disk', '50', self::AVAILIBILITY_LEVEL_LOW );
    }
    
    /**
     * Method creating chart pie components.
     */
    protected function _createChartPie()
    {
        $aDataseriesLabels1 = array(
        new PHPExcel_Chart_DataSeriesValues('String', 'document1!$A$1:$B$1', NULL),
        );

        $aXAxisTickValues1 = array(
            new PHPExcel_Chart_DataSeriesValues('String', 'document1!$A$2:$A$6', NULL),
        );

        $aDataSeriesValues1 = array(
            new PHPExcel_Chart_DataSeriesValues('Number', 'document1!$B$2:$B$6', NULL),
        );
        
        $series1 = new PHPExcel_Chart_DataSeries(
            PHPExcel_Chart_DataSeries::TYPE_PIECHART,
            PHPExcel_Chart_DataSeries::GROUPING_STACKED,
            range(0, count($aDataSeriesValues1)-1),
            $aDataseriesLabels1,
            $aXAxisTickValues1,
            $aDataSeriesValues1
        );

        $oLayout1 = new PHPExcel_Chart_Layout();
        //$oLayout1->setShowBubbleSize( true );
        //$oLayout1->setShowLegendKey( true );
        //$oLayout1->setShowLeaderLines( TRUE );
        //$oLayout1->setShowVal(TRUE);
        
        $oLayout1->setShowPercent(TRUE);
        
        $oPlotarea1 = new PHPExcel_Chart_PlotArea($oLayout1, array($series1));
        $oLegend1 = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
        $oTitle1 = new PHPExcel_Chart_Title(''/*Test Pie Chart'*/);

        $oChart1 = new PHPExcel_Chart(
            'chart1',
             $oTitle1,
             $oLegend1,
             $oPlotarea1,
             true,
             0,
             NULL,
             NULL
        );

        $oChart1->setTopLeftPosition('A7');
        $oChart1->setBottomRightPosition('D15');
        
        KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject()->getActiveSheet()->addChart( $oChart1 );
    }
    
    /**
     * Method generating report.
     * @return void
     */
    public function generate()
    {      
        KrscReports_Builder_Excel_PHPExcel::setPHPExcelObject( new PHPExcel() );
        $oCell = new KrscReports_Type_Excel_PHPExcel_Cell();
        
        $oBuilder = new KrscReports_Builder_Excel_PHPExcel_TableDifferentStyles();
        $oBuilder->setCellObject( $oCell );
        
        $this->_setData();
        $oBuilder->setData( $this->_aData );
        
        // creation of element responsible for creating table
        $oElementTable = new KrscReports_Document_Element_Table();
        $oElementTable->setBuilder( $oBuilder );
        
        // adding table to spreadsheet
        $oElement = new KrscReports_Document_Element();
        $oElement->addElement( $oElementTable );
        
        $oElement->beforeConstructDocument();
        $oElement->constructDocument();
        $this->_createChartPie();
        $oElement->afterConstructDocument();
    }

    public function getDescription() 
    {
        return 'Report creating chart pie.';
    }

}

