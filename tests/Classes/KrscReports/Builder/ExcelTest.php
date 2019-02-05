<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2019 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2019 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 2.0.1, 2019-02-05
 */

/**
 * Class for testing methods for creating excel spreadsheets.
 */
class KrscReports_Builder_ExcelTest extends PHPUnit_Framework_TestCase
{
    /**
     * Data provider for testFilterGroupName
     *
     * @return array
     */
    public function filterGroupNameDataProvider()
    {
        return array(
            array('abc123', 'abc123'),
            array('AB*', 'AB'),
            array('cd?', 'cd_'),
            array('a [2]', 'a _2_'),
            array('c d', 'c d', false),
            array('c d', 'c_d', true),
            array('c\'d', 'c d'),
            array("c'd", 'c d'),
        );
    }

    /**
     * Test with improper excel spreadsheet names.
     *
     * @covers KrscReports_Builder_Excel::filterGroupName
     *
     * @dataProvider filterGroupNameDataProvider
     */
    public function testFilterGroupName($sInputValue, $sExpectedOutputValue, $bReplaceSpaces = false)
    {
        $this->assertEquals(
            $sExpectedOutputValue,
            \KrscReports_Builder_Excel::filterGroupName($sInputValue, $bReplaceSpaces),
            'Output value for \KrscReports_Builder_Excel::filterGroupName not valid!'
        );
    }
}
