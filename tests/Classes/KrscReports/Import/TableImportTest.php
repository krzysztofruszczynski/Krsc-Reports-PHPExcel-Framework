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
 * @package KrscReports_Type
 * @copyright Copyright (c) 2017 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.1.1, 2017-04-28
 */

/**
 * Classes testing import functionality.
 * 
 * @category KrscReports
 * @package KrscReports_Tests
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */

/**
 * Stub for cell object.
 */
class KrscReports_Type_Excel_PHPExcel_CellStubTest extends KrscReports_Type_Excel_PHPExcel_Cell
{   
    /**
     * @var array data to be returned by stubbed cell object
     */
    protected $_aImportData = array(
        1 => array(  'first column', 'second column*', 'third column' ),
        2 => array( 'row1 col1', 'row1 col2', '' ),
        3 => array( 'row2 col1', '', '' ),
        4 => array( 'row3 col1', 'row3 col2', '' ),
        5 => array( 'row4 col1', '', 'row4 col3' ),
        6 => array( 'row5 col1', '0', 'row5 col3' ),
        7 => array( 'row6 col1', 0, 'row6 col3' ),
        8 => array( 'row7 col1', 0, false ),
        9 => array( 'row8 col1', false,'row8 col3' ),
        10 => array( '', '', '' )
    );
    
    /**
     * Method for changing mock data for getCellValue method.
     * @param array $aImportData new data
     */
    public function setImportData( $aImportData )
    {
        $this->_aImportData = $aImportData;
    }
    
    /**
     * Override constructor so PHPExcel doesn't have to be created.
     */
    public function __construct()
    {
        
    }
    
    /**
     * Getter for cell value (from loaded file or from PHPExcel object created during script execution). 
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     * @return mixed value for selected cell
     */
    public function getCellValue( $iColumnId, $iRowId )
    {
        return $this->_aImportData[$iRowId][$iColumnId];
    }
}

/**
 * Class for table import testing.
 */
class KrscReports_Import_TableImportTest extends PHPUnit_Framework_TestCase {

    /**
     * @var array columns in import
     */
    protected $_aColumnNames = array( 'first column', 'second column*', 'third column' );
    
    /**
     * @var string symbol which is part of required column name
     */
    protected $_sRequiredSymbol = '*';
    
    /**
     * @var array values expected in import without specifying required column symbol
     */
    protected $_aImportProperExpected = array(
        array( 'row1 col1', 'row1 col2', '' ),
        array( 'row2 col1', '', '' ),
        array( 'row3 col1', 'row3 col2', '' ),
        array( 'row4 col1', '', 'row4 col3' ),
        array( 'row5 col1', '0', 'row5 col3' ),
        array( 'row6 col1', 0, 'row6 col3' ),
        array( 'row7 col1', 0, false ),
        array( 'row8 col1', false,'row8 col3' )
    );
    
    protected $_aRequiredInvalidRows = array( 
        1 => array(1 => KrscReports_Import_TableImport::IMPORT_ERROR_TEXT ), 
        3 => array(1 => KrscReports_Import_TableImport::IMPORT_ERROR_TEXT ),
        7 => array(1 => KrscReports_Import_TableImport::IMPORT_ERROR_TEXT )
    );
    
    /**
     * @var KrscReports_Import_TableImport tested object
     */
    protected $object;

    /**
     * Method creating test object with default settings.
     */
    protected function createTestObject()
    {
        $this->object = new KrscReports_Import_TableImport;
        $this->object->setCellObject( new KrscReports_Type_Excel_PHPExcel_CellStubTest() );
        $this->object->setColumnNames( $this->_aColumnNames );
        $this->object->setRowNumberWithHeader( 1 );
    }
    
    /**
     * This method is called before a test is executed.
     */
    protected function setUp() 
    {
        $this->createTestObject();
    }

    /**
     * This method is called after a test is executed.
     */
    protected function tearDown() 
    {
        
    }

    /**
     * Test with proper import data.
     * @covers KrscReports_Import_TableImport::getImportedData
     */
    public function testGetProperImportedData() 
    {
        $this->assertSame( $this->_aImportProperExpected, $this->object->getImportedData() );
        $this->assertSame( array(), $this->object->getInvalidRows() );
    }
    
    /**
     * Test with proper import data and specified required column.
     * @covers KrscReports_Import_TableImport::getImportedData
     */
    public function testGetRequiredImportedData() 
    {   
        $this->object->setRequiredSymbol( $this->_sRequiredSymbol );
        $this->assertSame( $this->_aImportProperExpected, $this->object->getImportedData() );
        $this->assertSame( $this->_aRequiredInvalidRows, $this->object->getInvalidRows() );
    }
    
    /**
     * Test with declared column names and header row mismatch.
     * @covers KrscReports_Import_TableImport::getImportedData
     * @expectedException KrscReports_Exception_Import_InvalidHeader
     */
    public function testGetInvalidHeaderImportedData() 
    {         
        $aColumnNames = $this->_aColumnNames;
        $aColumnNames[0] = $aColumnNames[0] . 'invalidHeader';
        
        $this->object->setColumnNames( $aColumnNames );
        $this->assertSame( $this->_aImportProperExpected, $this->object->getImportedData() );
    }

}
