<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2016 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2014 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.1, 2016-11-12
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
    $oFile->setFileName( 'Krsc_Example_' . $_GET[KrscReports_Report_ExampleReport::INPUT_REPORT_ID] );
    /* report generation */
    KrscReports_Report_ExampleReport::generateReport( $_GET[KrscReports_Report_ExampleReport::INPUT_REPORT_ID] );
    $oFile->createFile();
}
else
{
    echo '<html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><title>Example reports from KrscReports library</title></head><body>';
    echo '<p>Example reports from KrscReports library. See <a href="https://github.com/krzysztofruszczynski">https://github.com/krzysztofruszczynski</a> for more informations</p>';
    echo '<ul>';
    
    foreach( KrscReports_Report_ExampleReport::getReportArray() as $iPosition => $oReport )
    {
        echo sprintf( '<li><a href="%s">%s</a></li><br/>', '?' . KrscReports_Report_ExampleReport::INPUT_REPORT_ID . '=' . $iPosition, $oReport->getDescription() );
    }
    
    echo '</ul>';
    
    echo '<p>Copyright (c) Krzysztof Ruszczyński 2014-2016 ( <a href="http://www.ruszczynski.eu">http://www.ruszczynski.eu</a> ). This library uses LGPL licence and works only with PHPExcel.';
    echo '</body></html>';
}
