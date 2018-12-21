<?php
namespace KrscReports\Views;

use KrscReports\Type\Excel\PHPExcel\Style;

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
 * @package KrscReports
 * @copyright Copyright (c) 2018 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.2.6, 2018-06-01
 */

/**
 * Abstract class for views in Excel.
 * 
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class AbstractView
{
    const SETTINGS_PHPEXCEL = 'PHPExcel';
    
    /**
     * key in options for auto filter
     */
    const KEY_AUTOFILTER = 'autofilter';
    
    /**
     * key in options where there is array column size (keys - column id starting from 0, values - width of column) 
     */
    const KEY_COLUMN_FIXED_SIZES = 'column_sizes';

    /**
     * key in options where there is array column size (keys - column id starting from 0, values - max width of column)
     */
    const KEY_COLUMN_MAX_SIZES = 'column_max_sizes';

    /**
     * key in options for number of rows between subsequent tables
     */
    const KEY_COLUMN_LINES_BETWEEN_ELEMENTS = 'lines_between_elements';

    /**
     * key in options for start column for table (counts from 0)
     */
    const KEY_START_COLUMN_INDEX = 'start_width_column_index';

    /**
     * key in options for start row for table (counts from 1)
     */
    const KEY_START_ROW_INDEX = 'start_height_row_index';

    /**
     * index of element in option array with name of translator domain to be used (only if required)
     */
    const KEY_TRANSLATOR_DOMAIN = 'translator_domain';
   
    /**
     * index of element in option array with file properties
     */
    const KEY_DOCUMENT_PROPERTIES = 'document_properties';
    
    /**
     * @var string name of selected settings
     */
    protected $selectedSettings;
    
    /**
     * @var Object service for translating column names
     */
    protected $columnTranslator;
     
    /**
     * @var array subsequent elements are titles for subsequent columns (if we don't want to use those from provided data or there are no data)
     */
    protected $columnNames;
    
    /**
     * @var \KrscReports\Type\Excel\PHPExcel\Style\Bundle\AbstractStyle style builder
     */
    protected $styleBuilder;
    
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Cell instance of cell
     */
    protected $cell;
    
    /**
     * @var array data for spreadsheet 
     */
    protected $data = array();
    
    /**
     * @var array options for table (for example legend) 
     */
    protected $options = array();
    
    /**
     * @var \KrscReports_Document_Element element, to which excel elements are attached 
     */
    protected $documentElement;
    
    public function __construct( $settings = self::SETTINGS_PHPEXCEL )
    {
        switch ( $settings ) { 
            case self::SETTINGS_PHPEXCEL:
                \KrscReports_Builder_Excel_PHPExcel::setPHPExcelObject( new \PHPExcel() );
                $this->documentElement = new \KrscReports_Document_Element();
                $this->selectedSettings = $settings;
                break;
        }
    }
    
    public function setColumnTranslator( $columnTranslator )
    {
        $this->columnTranslator = $columnTranslator;
    }

    /**
     * Method for processing column names.
     *
     * @param  array  $columns  subsequent elements are names of subsequent columns (to be translated)
     *
     * @return $this
     */
    public function setColumns($columns = array())
    {
        $this->columnNames = $this->columnTranslator->translateColumns($columns, (isset( $this->options[self::KEY_TRANSLATOR_DOMAIN]) ? $this->options[self::KEY_TRANSLATOR_DOMAIN] : null));

        return $this;
    }

    /**
     * Method for setting document properties. To be executed after options are set.
     */
    public function setDocumentProperties()
    {
        switch ( $this->selectedSettings ) {
            case self::SETTINGS_PHPEXCEL:
                if (isset($this->options[self::KEY_DOCUMENT_PROPERTIES])) {
                    if ( isset( $this->columnTranslator ) ) {
                        // translate properties
                        $this->options[self::KEY_DOCUMENT_PROPERTIES] = $this->columnTranslator->translateColumns( $this->options[self::KEY_DOCUMENT_PROPERTIES], ( isset( $this->options[self::KEY_TRANSLATOR_DOMAIN] ) ? $this->options[self::KEY_TRANSLATOR_DOMAIN] : '' ) );
                    }

                    \KrscReports_Builder_Excel_PHPExcel::setDocumentProperties( $this->options[self::KEY_DOCUMENT_PROPERTIES] );
                }

                break;
        }
    }

    /**
     * Getter for column style name (can be override). Works only if settings via constructor were properly set.
     * @return string name of column with styles
     */
    public static function getStyleColumnName()
    {
        return \KrscReports_Builder_Excel_PHPExcel_TableDifferentStyles::DATA_STYLE_COLUMN;
    }
    
    /**
     * Constructor for cell object.
     * @return \KrscReports_Type_Excel_PHPExcel_Cell new object of cell
     */
    public static function getCellObject()
    {
        return new \KrscReports_Type_Excel_PHPExcel_Cell();
    }
    
    /**
     * Setter for style builder.
     * @param \KrscReports\Type\Excel\PHPExcel\Style\Bundle\AbstractStyle $styleBuilder style to be set (default: Style\DefaultStyle)
     */
    public function setStyleBuilder( $styleBuilder = null )
    {
        if ( is_null( $styleBuilder ) ) {
            $styleBuilder = new Style\Bundle\DefaultStyle();
        }
        
        $this->styleBuilder = $styleBuilder;
    }
    
    /**
     * Method for getting cell associated with that object.
     * @param boolean $forceNew if true, new object is always created (need to clone every execution of method if we want to have different objects)
     * @return \KrscReports_Type_Excel_PHPExcel_Cell
     */
    public function getCell( $forceNew = false )
    {
        if ( isset( $this->cell ) && !$forceNew ) {
            return $this->cell;
        }
        
        // if previously set, it will remain, otherwise - default one
        $this->setStyleBuilder( $this->styleBuilder );
        
        $this->cell = self::getCellObject();
        $this->cell->setStyleObject( $this->styleBuilder->getStyleObject() );

        if ( isset( $this->options[self::KEY_COLUMN_FIXED_SIZES] )) {
            if (is_array( $this->options[self::KEY_COLUMN_FIXED_SIZES])) {
                foreach ( $this->options[self::KEY_COLUMN_FIXED_SIZES] as $columnId => $columnSize ) {
                    $this->cell->setColumnFixedSize( $columnId, $columnSize );
                }
            } else {
                $this->cell->setColumnFixedSize(null, $this->options[self::KEY_COLUMN_FIXED_SIZES]);
            }
        }

        if ( isset( $this->options[self::KEY_COLUMN_MAX_SIZES] )) {
            if (is_array( $this->options[self::KEY_COLUMN_MAX_SIZES])) {
                foreach ( $this->options[self::KEY_COLUMN_MAX_SIZES] as $columnId => $columnSize ) {
                    $this->cell->setColumnMaxSize($columnId, $columnSize);
                }
            } else {
                $this->cell->setColumnMaxSize(null, $this->options[self::KEY_COLUMN_MAX_SIZES]);
            }
        }

        return $this->cell;
    }

    /**
     * Method adding column names to data provided in input.
     * @param array $data
     * @param array $columnNames
     * @return array
     */
    protected function addColumnNames( $data, $columnNames = array() )
    {
        if( !empty( $columnNames ) ){
            $outputData = array();
            foreach( $data as $key => $row ){                
                $styleColumn = ''; // can be string or array
                if ( isset( $row[self::getStyleColumnName()] ) ) { // removing style element from array before array_combine
                    $styleColumn = $row[self::getStyleColumnName()];
                    unset( $row[self::getStyleColumnName()] );
                }

                if ( count( $columnNames ) !== count( $row ) ) {
                    throw new \Exception(sprintf( 'Array sizes are not the same: (columnNames: %s, row: %s, rowNo: %s)', count($columnNames), count($row), $key));
                }

                $outputData[$key] = array_combine( $columnNames, $row );  

                if ( !empty( $styleColumn ) ) { // inserting style element after array_combine
                    $outputData[$key][self::getStyleColumnName()] = $styleColumn;
                }
            }
        } else {
            $outputData = $data;
        }

        return $outputData;
    }

    /**
     * Method setting data for spreadsheet.
     *
     * @param array $data data for spreadsheet
     * @param array $columnNames subsequent elements are names of columns (optional)
     *
     * @return AbstractView
     */
    public function setData($data, $columnNames = array())
    {   
        if( !empty( $columnNames ) ){
            $this->columnNames = $columnNames;
        }

        $this->data = $this->addColumnNames($data, $this->columnNames);

        return $this;
    }
    
    /**
     * Get actual set of options associated with that view (for example for preserve from overriding)
     * @return array actual options
     */
    public function getOptions()
    {
        return $this->options;
    }
    
    public function addOptions( $options )
    {
        $this->options = array_merge( $this->options, $options );
        return $this;
    }
    
    /**
     * Method used for testing.
     * @return array previously set data
     */
    public function getData()
    {
        return $this->data;
    }
    
    /**
     * Setting document element (if we want to attach this view to existing document).
     * @param \KrscReports_Document_Element $documentElement
     * @return \KrscReports\Views\AbstractView object on which this method was executed
     */
    public function setDocumentElement( $documentElement )
    {
        $this->documentElement = $documentElement;
        return $this;
    }
    
    /**
     * Method for getting document element after adding data associated with view
     * @var string $spreadsheetName name of worksheet in which data would be placed
     * @return \KrscReports_Document_Element()
     */
    abstract public function getDocumentElement( $spreadsheetName );
    
    /**
     * Method for generating document.
     * @param string $spreadsheetName spreadsheet name (optional, if null than default name would be used)
     */
    public function generate( $spreadsheetName = null )
    {
        if ( !is_null( $spreadsheetName ) ) {
            $element = $this->getDocumentElement( $spreadsheetName );
        } else {
            $element = $this->getDocumentElement();
        }        
        
        $element->beforeConstructDocument();
        $element->constructDocument();
        $element->afterConstructDocument();  
    }
}
