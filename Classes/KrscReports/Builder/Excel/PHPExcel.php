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
 * @version 2.0.1, 2018-11-07
 */

/**
 * Extension of abstract builder suitable with PHPExcel. Contains functionality closely connected with PHPExcel.
 * Method constructDocument() is still abstract and needs to be implemented further.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel implements \KrscReports\Builder\Excel\ExcelBuilderTypeInterface
{
    /**
     * @var \PHPExcel instance of PHPExcel used while adding new data to spreadsheets
     */
    protected static $_oPHPExcel;

    /**
     * @var string name of group in which this content is generated
     */
    protected $_sGroupName;

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
     * Method for setting file properties
     * @param array $aDocumentProperties property name is key (must have a set method), property value is array value
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
     * Method for setting group name for current builder (used for Excel spreadsheet name).
     * @param string $sGroupName group name to be set (means that current builder will write to Excel spreadsheet with the same name)
     * @return KrscReports_Builder_Excel_PHPExcel object on which method was executed
     */
    public function setGroupName( $sGroupName )
    {
        $sGroupName = \KrscReports_Builder_Excel::filterGroupName($sGroupName);
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
     * @return string actually set group name
     */
    public function getGroupName()
    {
        return $this->_sGroupName;
    }
}
