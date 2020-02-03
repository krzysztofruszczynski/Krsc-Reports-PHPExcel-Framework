<?php
use KrscReports\Import\ReaderTrait;

/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2020 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2020 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 2.0.6, 2020-02-02
 */

/**
 * Example report creating one table with different cell's styles in PHPExcel.
 *  
 * @category KrscReports
 * @package KrscReports_Report
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Report_ExampleReportDifferentSizes extends KrscReports_Report_ExampleReport
{
    /**
     * key style with light blue fill
     */
    const STYLE_CELL_LIGHT_BLUE = 'style_light_blue';
    
    /**
     * key style with gray fill
     */
    const STYLE_CELL_GRAY = 'style_gray';
    
    /**
     * key style with dark green fill
     */
    const STYLE_CELL_DARK_GREEN = 'style_dark_green';
    
    /**
     * key style with dark yellow fill
     */
    const STYLE_CELL_YELLOW = 'style_dark_yellow';
    
    /**
     * key style with orange fill
     */
    const STYLE_CELL_ORANGE = 'style_orange';
    
    /**
     * key style for gradient fill (from orange to red)
     */
    const STYLE_CELL_DARK_RED = 'style_dark_red';
    
    /**
     * key style with red fill
     */
    const STYLE_CELL_RED = 'style_red';
    
    /**
     * title of column with product name
     */
    const COLUMN_PRODUCT_NAME = 'Product name';
    
    /**
     * title of column with product price
     */
    const COLUMN_PRODUCT_PRICE = 'Product price (in EUR)';
    
    /**
     * title of column with product availability
     */
    const COLUMN_PRODUCT_AVAILIBILITY = 'Product availability';
    
    /**
     * value for availability, when product is not available
     */
    const AVAILIBILITY_LEVEL_NOT_AVAILABLE = 0;
    
    /**
     * value for availability, when product supply is very low
     */
    const AVAILIBILITY_LEVEL_VERY_LOW = 1;
    
    /**
     * value for availability, when product supply is low
     */
    const AVAILIBILITY_LEVEL_LOW = 2;
    
    /**
     * value for availability, when product supply is medium
     */
    const AVAILIBILITY_LEVEL_MEDIUM = 3;
    
    /**
     * value for availability, when product supply is high
     */
    const AVAILIBILITY_LEVEL_HIGH = 4;
    
    /**
     * @var array key is availability value which stores style name to be used 
     */
    protected $_aAvailabilityStyles = array( 
        self::AVAILIBILITY_LEVEL_NOT_AVAILABLE => self::STYLE_CELL_DARK_RED,
        self::AVAILIBILITY_LEVEL_VERY_LOW => self::STYLE_CELL_RED,
        self::AVAILIBILITY_LEVEL_LOW => self::STYLE_CELL_ORANGE,
        self::AVAILIBILITY_LEVEL_MEDIUM => self::STYLE_CELL_YELLOW,
        self::AVAILIBILITY_LEVEL_HIGH => self::STYLE_CELL_DARK_GREEN
    );
    
    /**
     * @var array key is availability value which stores title to be used in spreadsheet 
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
        $aRow[KrscReports_Builder_Excel_PHPExcel_TableDifferentStyles::DATA_STYLE_COLUMN] = array(
            self::COLUMN_PRODUCT_NAME => self::STYLE_CELL_GRAY,
            KrscReports_Builder_Excel_PHPExcel_TableDifferentStyles::STYLE_CELL_DEFAULT => self::STYLE_CELL_LIGHT_BLUE
        );
        $aRow[KrscReports_Builder_Excel_PHPExcel_TableDifferentStyles::DATA_STYLE_COLUMN][self::COLUMN_PRODUCT_AVAILIBILITY] =
                $this->_aAvailabilityStyles[$iProductAvailibility];
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
     * Method generating report.
     * @return void
     */
    public function generate()
    {
        KrscReports_Builder_Excel::setExcelObject();
        $oCell = ReaderTrait::getCellObject();
        KrscReports_Builder_Excel::setDocumentProperties();

        $oCollectionDefault = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionDefault->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_ExampleBorders() );
        
        $oCollectionRowLightGreen = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionRowLightGreen->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic() );
        $oCollectionRowLightGreen->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        $oFillYellow = new KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $oFillYellow->setColor( PHPExcel_Style_Color::COLOR_YELLOW );
        $oCollectionYellow = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionYellow->addStyleElement( $oFillYellow );
        $oCollectionYellow->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        $oFillRed = new KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $oFillRed->setColor( PHPExcel_Style_Color::COLOR_RED );
        $oCollectionRowRed = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionRowRed->addStyleElement( $oFillRed );
        $oCollectionRowRed->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        $oFillDarkGreen = new KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $oFillDarkGreen->setColor( PHPExcel_Style_Color::COLOR_DARKGREEN );
        $oCollectionDarkGreen = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionDarkGreen->addStyleElement( $oFillDarkGreen );
        $oCollectionDarkGreen->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        $oFillOrange = new KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $oFillOrange->setColor( KrscReports_Type_Excel_PHPExcel_Style_Default::COLOR_ORANGE );
        $oCollectionOrange = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionOrange->addStyleElement( $oFillOrange );
        $oCollectionOrange->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        $oFillDarkRed = new KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $oFillDarkRed->setColor( PHPExcel_Style_Color::COLOR_DARKRED );
        $oCollectionDarkRed = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionDarkRed->addStyleElement( $oFillDarkRed );
        $oCollectionDarkRed->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        $oFillLightBlue = new KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $oFillLightBlue->setColor( KrscReports_Type_Excel_PHPExcel_Style_Default::COLOR_LIGHT_BLUE );
        $oCollectionLightBlue = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionLightBlue->addStyleElement( $oFillLightBlue );
        $oCollectionLightBlue->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        $oFillGray = new KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $oFillGray->setColor( KrscReports_Type_Excel_PHPExcel_Style_Default::COLOR_GRAY );
        $oCollectionGray = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionGray->addStyleElement( $oFillGray );
        $oCollectionGray->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        
        $oStyle = new KrscReports_Type_Excel_PHPExcel_Style();
        $oStyle->setStyleCollection( $oCollectionDefault );
        $oStyle->setStyleCollection( $oCollectionRowLightGreen, KrscReports_Document_Element_Table::STYLE_ROW );
        $oStyle->setStyleCollection( $oCollectionYellow, self::STYLE_CELL_YELLOW );
        $oStyle->setStyleCollection( $oCollectionRowRed, self::STYLE_CELL_RED );
        $oStyle->setStyleCollection( $oCollectionDarkGreen, self::STYLE_CELL_DARK_GREEN );
        $oStyle->setStyleCollection( $oCollectionOrange, self::STYLE_CELL_ORANGE );
        $oStyle->setStyleCollection( $oCollectionDarkRed, self::STYLE_CELL_DARK_RED );
        $oStyle->setStyleCollection( $oCollectionLightBlue, self::STYLE_CELL_LIGHT_BLUE );
        $oStyle->setStyleCollection( $oCollectionGray, self::STYLE_CELL_GRAY );

        $oCell->setStyleObject( $oStyle );
        $oCell->setColumnMaxSize(0, 15);
        $oCell->setColumnFixedSize(2, 15);

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
        $oElement->afterConstructDocument();    
    }

    public function getDescription() 
    {
        return 'Report with one table with fixed and max size of column set.';
    }
}
