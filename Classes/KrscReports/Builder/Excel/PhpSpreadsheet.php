<?php
namespace KrscReports\Builder\Excel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Exception;

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
 * @version 2.0.0, 2018-09-27
 */

/**
 * Extension of abstract builder suitable with PhpSpreadsheet. Contains functionality closely connected with PhpSpreadsheet.
 * Method constructDocument() is still abstract and needs to be implemented further.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class PhpSpreadsheet implements ExcelBuilderTypeInterface
{
    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    protected static $_oSpreadsheet;

    /**
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $oSpreadsheet
     */
    public static function setSpreadsheetObject(Spreadsheet $oSpreadsheet)
    {
        self::$_oSpreadsheet = $oSpreadsheet;
    }

    /**
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public static function getPhpSpreadsheetObject()
    {
        return self::$_oSpreadsheet;
    }

    /**
     * Method for setting file properties
     * @param array $aDocumentProperties property name is key (must have a set method), property value is array value
     */
    public static function setDocumentProperties( $aDocumentProperties )
    {
        $oProperties = self::$_oSpreadsheet->getProperties();

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
     * @return \KrscReports\Builder\Excel\PhpSpreadsheet object on which method was executed
     */
    public function setGroupName( $sGroupName )
    {
        $sGroupName = \KrscReports_Builder_Excel::filterGroupName( $sGroupName );       
        $this->_sGroupName = $sGroupName;

        try
        {
            self::$_oSpreadsheet->setActiveSheetIndexByName( $sGroupName );
        }
        catch (Exception $oException )
        {   // create new worksheet
            if( count( self::$_oSpreadsheet->getAllSheets()) == 1 && self::$_oSpreadsheet->getActiveSheet()->getTitle() == 'Worksheet' )
            {   // renaming default worksheet
                self::$_oSpreadsheet->setActiveSheetIndex(0)->setTitle( $sGroupName );
            }
            else
            {   // second or later worksheet
                $oNewSheet = self::$_oSpreadsheet->createSheet();
                $oNewSheet->setTitle( $sGroupName );
            }
            
            self::$_oSpreadsheet->setActiveSheetIndexByName( $sGroupName );
        }

        return $this;
    }
}
