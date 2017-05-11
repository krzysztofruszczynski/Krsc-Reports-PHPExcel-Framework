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
 * @version 1.1.4, 2017-05-11
 */

/**
 * Extension of abstract builder suitable with PHPExcel. Contains functionality closely connect with PHPExcel.
 * Method constructDocument() is still abstract and needs to be implemented further.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class KrscReports_Builder_Excel_PHPExcel extends KrscReports_Builder_Excel
{   
    /**
     * Maximum number of signs allowed in group name
     */
    const GROUP_NAME_LENGTH_LIMIT = 31;
    
    /**
     * array (
     *   '*' => '',
     *   ':' => '_',
     *   '/' => '_',
     *   '\\' => '_',
     *   '?' => '_',
     *   '[' => '_',
     *   ']' => '_',
     *   ' ' => '_',
     * )
     *
     * Array with illegal signs inside group name (key: illegal sign, value: replacement)
     * Note: space is allowed for excel spreadsheet names, but causes problems with name ranges for PHPExcel (for example with graphs)
     */
    const GROUP_NAME_ILLEGAL_SIGNS = 'a:8:{s:1:"*";s:0:"";s:1:":";s:1:"_";s:1:"/";s:1:"_";s:1:"\";s:1:"_";s:1:"?";s:1:"_";s:1:"[";s:1:"_";s:1:"]";s:1:"_";s:1:" ";s:1:"_";}';        
            
    /**
     * @var \PHPExcel instance of PHPExcel used while adding new data to spreadsheets
     */
    protected static $_oPHPExcel;
    
    /**
     * @var KrscReports_Type_Excel_PHPExcel_Cell object responsible for creating cell inside Excel spreadsheet
     */
    protected $_oCell;
    
    /**
     * Method for setting objects responsible for creating cell inside Excel spreadsheet.
     * @param KrscReports_Type_Excel_PHPExcel_Cell $oCell object to be set to create cells inside Excel spreadsheet
     * @return KrscReports_Builder_Excel_PHPExcel object on which method was executed
     */
    public function setCellObject( $oCell )
    {
        $this->_oCell = $oCell;
        return $this;
    }
    
    /**
     * Method for setting PHPExcel object inside class.
     * @param PHPExcel $oPHPExcel PHPExcel object to be used while building Excel spreadsheets
     * @return void
     */
    public static function setPHPExcelObject( PHPExcel $oPHPExcel )
    {
        self::$_oPHPExcel = $oPHPExcel;
    }
    
    /**
     * Method returning PHPExcel object stored inside a class.
     * @return PHPExcel stored PHPExcel object
     * @throws Exception thrown when PHPExcel has not been set before (possibly only via self::setPHPExcelObject)
     */
    public static function getPHPExcelObject()
    {
        if( isset( self::$_oPHPExcel ) )
        {
            return self::$_oPHPExcel;
        }
        
        throw new Exception( 'PhpExcel object not set in KrscReports_Builder_Excel_PHPExcel::setPhpExcelObject( PhpExcel $oPhpExcel )' );
    }

    /**
     * Method filtering group name in order to be compatible with Excel naming standards (length and lack of some signs).
     * @param string $sGroupName name of group (excel spreadsheet) to be filtered
     * @return string filtered group name
     */
    public static function filterGroupName( $sGroupName )
    {
        $sGroupName = strtr( $sGroupName, unserialize( self::GROUP_NAME_ILLEGAL_SIGNS ) );
        return substr( $sGroupName, 0, self::GROUP_NAME_LENGTH_LIMIT );
    }
    
    /**
     * Method for setting group name for current builder (used for Excel spreadsheet name).
     * @param string $sGroupName group name to be set (means that current builder will write to Excel spreadsheet with the same name)
     * @return KrscReports_Builder_Excel_PHPExcel object on which method was executed
     */
    public function setGroupName( $sGroupName )
    {
        $sGroupName = self::filterGroupName( $sGroupName );       
        $this->_sGroupName = $sGroupName;
        
        try
        {
            self::$_oPHPExcel->setActiveSheetIndexByName( $sGroupName );
        }
        catch ( PHPExcel_Exception $oException )
        {   // create new worksheet
            if( count( self::$_oPHPExcel->getAllSheets()) == 1 && self::$_oPHPExcel->getActiveSheet()->getTitle() == 'Worksheet' )
            {   // renaming default worksheet
                self::$_oPHPExcel->setActiveSheetIndex()->setTitle( $sGroupName );
            }
            else
            {   // second or later worksheet
                $oNewSheet = self::$_oPHPExcel->createSheet();
                $oNewSheet->setTitle( $sGroupName );                
            }
            
            self::$_oPHPExcel->setActiveSheetIndexByName( $sGroupName );
        }
        
        return $this;
    }
    
    /**
     * Method for setting file properties
     * @param array $aDocumentProperties property name is key (must have a set method), property value is array value
     * @return KrscReports_Builder_Excel_PHPExcel object on which method was executed
     */
    public static function setDocumentProperties( $aDocumentProperties )
    {
        $oProperties = self::$_oPHPExcel->getProperties();
        
        foreach ( $aDocumentProperties as $sPropertyName => $sPropertyValue ) 
        {
            $sMethodName = 'set' . $sPropertyName;
            if ( method_exists( $oProperties, $sMethodName ) ) 
            {
                $oProperties->$sMethodName( $sPropertyValue );
            }
        }
    }
    
    /**
     * Method setting actual style key (can be set from element layer).
     * @param string $sStyleKey style key to be set
     * @param boolean $bIfStyleNotExistsSelectDefault (by default false) if style key was not found and parameter is true, 
     * there would be default style set; otherwise (if false) previously used style would continue
     * @return boolean true if style was found, false otherwise
     */
    public function setStyleKey( $sStyleKey, $bIfStyleNotExistsSelectDefault = false )
    {
        return $this->_oCell->setStyleKey( $sStyleKey, $bIfStyleNotExistsSelectDefault );
    }
}
