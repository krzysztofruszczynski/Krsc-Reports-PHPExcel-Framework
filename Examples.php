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
 * @package KrscReports
 * @copyright Copyright (c) 2020 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 2.1.0, 2020-02-24
 */

/**
 * Structural code handling example reports.
 *
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
require_once('init.php');

KrscReports_Report_ExampleReport::createObjects();

if( isset( $_GET[KrscReports_Report_ExampleReport::INPUT_REPORT_ID] ) )
{
    /* creating output file */
    $oFile = new KrscReports_File();
    if(isset($_GET[KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME])) {
        if(in_array($_GET[KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME], KrscReports_File::getAllSettingsNames())) {
            KrscReports_File::setBuilderType($_GET[KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME]);
        }
    }

    $oFile->setFileName( 'Krsc_Example_' . $_GET[KrscReports_Report_ExampleReport::INPUT_REPORT_ID] );
    /* report generation */
    KrscReports_Report_ExampleReport::generateReport( $_GET[KrscReports_Report_ExampleReport::INPUT_REPORT_ID] );
    $oFile->createFile();
}
else
{
    echo '<html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><title>Example reports from KrscReports library</title></head><body>';
    echo '<p>Example reports from KrscReports library. See <a href="https://github.com/krzysztofruszczynski">https://github.com/krzysztofruszczynski</a> for more informations</p>';
    echo '<p>';
    if (class_exists('PHPExcel_Style_Color')) {
        // PHPExcel classes are present - possible to switch to PHPExcel
        if (isset($_GET[KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME]) && $_GET[KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME] === KrscReports_File::SETTINGS_PHPEXCEL) {
            echo 'Currently reports are made using PHPExcel.';
            echo '<br/><button onclick="location.href = \'?'.KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME.'='.KrscReports_File::SETTINGS_PHPSPREADSHEET.'\';">Switch to '.KrscReports_File::SETTINGS_PHPSPREADSHEET.'</button>';
        } else {
            echo 'Currently reports are made using PhpSpreadsheet.';
            echo '<br/><button onclick="location.href = \'?'.KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME.'='.KrscReports_File::SETTINGS_PHPEXCEL.'\';">Switch to '.KrscReports_File::SETTINGS_PHPEXCEL.'</button>';
        }
    } else if (!class_exists('\PhpOffice\PhpSpreadsheet\Style\Color')) {
        echo '<b>PhpSpreadsheet neither PHPExcel is installed, please install via composer at least one of them (PhpSpreadsheet recommended).</b>';
    } else {
        echo 'All reports are made only using PhpSpreadsheet because PHPExcel is not present (don\'t install if not needed for backward compatibility).';
    }
    echo '</p><ul>';

    $bReportsWithView = false;
    foreach( KrscReports_Report_ExampleReport::getReportArray() as $iPosition => $oReport )
    {
        if ($iPosition === 0) {
            echo '<p>Examples with custom configuration:</p>';
        } else if ($bReportsWithView === false && $oReport instanceof KrscReports\Report\ReportWithServiceInterface) {
            echo '<p>Examples with view architecture used:</p>';
            $bReportsWithView = true;
        }
        if (isset($_GET[KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME])) {
            echo sprintf('<li><a href="?%s=%d&%s=%s">%s</a></li><br/>', KrscReports_Report_ExampleReport::INPUT_REPORT_ID, $iPosition, KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME, $_GET[KrscReports_Report_ExampleReport::INPUT_SETTINGS_NAME], $oReport->getDescription());
        } else {
            echo sprintf('<li><a href="?%s=%d">%s</a></li><br/>', KrscReports_Report_ExampleReport::INPUT_REPORT_ID, $iPosition, $oReport->getDescription());
        }
    }

    echo '</ul>';

    echo '<p>Copyright (c) Krzysztof Ruszczyński 2014 - 2020 ( <a href="http://www.ruszczynski.eu">http://www.ruszczynski.eu</a> ). This library uses LGPL licence and works only with PhpSpreadsheet and/or PHPExcel.';
    echo '</body></html>';
}
