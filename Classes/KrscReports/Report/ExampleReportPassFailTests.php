<?php
namespace KrscReports\Report;

use KrscReports\Service;
use KrscReports\Type\Excel\PhpSpreadsheet\Style\Bundle\PassFailStyle;
use KrscReports\Views\AbstractView;
use KrscReports\Views\SingleTable;

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
 * @version 2.0.6, 2020-02-17
 */

/**
 * Example report creating table with tests results, marking with color, if test was successful.
 *
 * @category KrscReports
 * @package KrscReports_Report
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class ExampleReportPassFailTests extends \KrscReports_Report_ExampleReport
{
    const COLUMN_TEST_NAME = 'Test name';

    const COLUMN_TEST_RESULT = 'Test result';

    const COLUMN_TEST_DATE = 'Test date';

    const COLUMN_COMMENTS = 'Comments';

    /**
     * @var integer number of rows to create
     */
    protected $_iNumberOfRows = 20;

    /**
     * Method returns description of generated by that object worksheet.
     * @return string description of report
     */
    public function getDescription()
    {
        return 'Report with table with marked tests results (made with \KrscReports\Service, random tests results and dates).';
    }

    /**
     * Method responsible for creating service with generated report.
     * @return void
     */
    public function generate()
    {
        $oExcelService = new Service();
        $oExcelService->setSpreadsheetName('Results_Spreadsheet');
        $oExcelService->setFileName('Krsc_Example_11');

        $aDataForExcel = array();
        $iIterator = 1;
        for ($i = 0; $i < $this->_iNumberOfRows; $i++) {
            $bPassFailResult = boolval(mt_rand(0, 1));
            $aDataForExcel[] = array(
                self::COLUMN_TEST_NAME => 'Test number '.$iIterator++,
                self::COLUMN_TEST_RESULT => ($bPassFailResult ? 'Pass' : 'Fail'),
                self::COLUMN_TEST_DATE => \date(\DateTime::RFC822, (time() - (mt_rand(0, 10000)*86400/100))),
                self::COLUMN_COMMENTS => '',
                AbstractView::getStyleColumnName() => array(
                    self::COLUMN_TEST_RESULT => ($bPassFailResult ? PassFailStyle::STYLE_PASS : PassFailStyle::STYLE_FAIL),
                ),
            );
        }

        $oReportView = new SingleTable();
        $oReportView->addOptions(
            array(
                AbstractView::KEY_DOCUMENT_PROPERTIES => array(),
                SingleTable::KEY_SINGLE_TABLE_TITLE => 'Tests results',
            )
        );
        $oReportView->setStyleBuilder(new PassFailStyle());
        $oReportView->setData(
            $aDataForExcel,
            array(self::COLUMN_TEST_NAME, self::COLUMN_TEST_RESULT, self::COLUMN_TEST_DATE, self::COLUMN_COMMENTS)
        );
        $oReportView->setDocumentProperties();

        $oExcelService->setReportView($oReportView);
        $oExcelService->createReport();
    }
}
