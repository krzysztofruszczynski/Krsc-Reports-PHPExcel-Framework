<?php
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
 * @package KrscReports_Builder
 * @copyright Copyright (c) 2020 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 2.1.0, 2020-02-02
 */

/**
 * Extension of abstract builder which groups all methods for creating Excel spreadsheets
 * (for example PHPExcel).
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class KrscReports_Builder_Excel extends \KrscReports_Builder_Abstract
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
     *   '\'' => ' ', // theoretically allowed, but cause problems with formulas
     * )
     *
     * Array with illegal signs inside group name (key: illegal sign, value: replacement)
     */
    const GROUP_NAME_ILLEGAL_SIGNS = 'a:8:{s:1:"*";s:0:"";s:1:":";s:1:"_";s:1:"/";s:1:"_";s:1:"\";s:1:"_";s:1:"?";s:1:"_";s:1:"[";s:1:"_";s:1:"]";s:1:"_";s:1:"\'";s:1:" ";}';

    /**
     * @var \KrscReports\Type\Excel\Cell object responsible for creating cell inside Excel spreadsheet
     */
    protected $_oCell;

    /**
     * @var object for creating excel file
     */
    protected static $_oBuilderType;

    /**
     * Method for setting objects responsible for creating cell inside Excel spreadsheet.
     * @param \KrscReports\Type\Excel\Cell $oCell object to be set to create cells inside Excel spreadsheet
     * @return \KrscReports_Builder_Excel object on which method was executed
     */
    public function setCellObject( $oCell )
    {
        $this->_oCell = $oCell;

        return $this;
    }

    /**
     * Method filtering group name in order to be compatible with Excel naming standards (length and lack of some signs).
     * @param string $sGroupName name of group (excel spreadsheet) to be filtered
     * @param boolean $bReplaceSpaces if true, then spaces in group name are replaced with underscore (default: false)
     * @return string filtered group name
     */
    public static function filterGroupName( $sGroupName, $bReplaceSpaces = false )
    {
        $sGroupName = strtr( $sGroupName, unserialize( self::GROUP_NAME_ILLEGAL_SIGNS ) );

        if ( $bReplaceSpaces ) {
            $sGroupName = str_replace( ' ', '_', $sGroupName );
        }

        return substr( $sGroupName, 0, self::GROUP_NAME_LENGTH_LIMIT );
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

    /**
     * Method for setting builder type (used by setGroupName and setDocumentProperties).
     */
    public static function setBuilderType()
    {
        if (!isset(self::$_oBuilderType)) {
            switch (\KrscReports_File::getBuilderType()) {
                case \KrscReports_File::SETTINGS_PHPEXCEL:
                    self::$_oBuilderType = new KrscReports_Builder_Excel_PHPExcel();
                    break;
                case \KrscReports_File::SETTINGS_PHPSPREADSHEET:
                    self::$_oBuilderType = new \KrscReports\Builder\Excel\PhpSpreadsheet();
                    break;
            }
        }
    }

    /**
     * Method setting object responsible for creating excel file (depends on vendor).
     *
     * @return void
     */
    public static function setExcelObject()
    {
        switch (\KrscReports_File::getBuilderType()) {
            case \KrscReports_File::SETTINGS_PHPEXCEL:
                \KrscReports_Builder_Excel_PHPExcel::setPHPExcelObject(new \PHPExcel());
                break;
            case \KrscReports_File::SETTINGS_PHPSPREADSHEET:
                \KrscReports\Builder\Excel\PhpSpreadsheet::setSpreadsheetObject(new \PhpOffice\PhpSpreadsheet\Spreadsheet());
                break;
        }
    }

    /**
     * Method returning object responsible for creating excel file (depends on vendor).
     *
     * @return KrscReports_Builder_Excel_PHPExcel|\KrscReports\Builder\Excel\PhpSpreadsheet
     */
    public static function getExcelObject()
    {
        self::setBuilderType();

        switch (\KrscReports_File::getBuilderType()) {
            case \KrscReports_File::SETTINGS_PHPEXCEL:
                $returnObject = self::$_oBuilderType->getPHPExcelObject();
                break;
            case \KrscReports_File::SETTINGS_PHPSPREADSHEET:
                $returnObject = self::$_oBuilderType->getPhpSpreadsheetObject();
                break;
        }

        return $returnObject;
    }

    /**
     * Method for setting group name for current builder (used for Excel spreadsheet name).
     * @param string $sGroupName group name to be set (means that current builder will write to Excel spreadsheet with the same name)
     * @return \KrscReports_Builder_Excel_PHPExcel|\KrscReports\Builder\Excel\PhpSpreadsheet object on which method was executed
     */
    public function setGroupName( $sGroupName )
    {
        self::setBuilderType();

        return self::$_oBuilderType->setGroupName($sGroupName);
    }

    /**
     * Method for setting file properties
     * @param array $aDocumentProperties property name is key (must have a set method), property value is array value
     * @return \KrscReports_Builder_Excel_PHPExcel|\KrscReports\Builder\Excel\PhpSpreadsheet object on which method was executed
     */
    public static function setDocumentProperties($aDocumentProperties = array())
    {
        self::setBuilderType();

        switch (\KrscReports_File::getBuilderType()) {
            case \KrscReports_File::SETTINGS_PHPEXCEL:
                $aDocumentProperties['Creator'] = 'PHPExcel';
                break;
            case \KrscReports_File::SETTINGS_PHPSPREADSHEET:
                $aDocumentProperties['Creator'] = 'PhpSpreadsheet';
                break;
        }

        return self::$_oBuilderType->setDocumentProperties($aDocumentProperties);
    }
}
