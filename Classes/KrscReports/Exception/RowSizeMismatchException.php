<?php
namespace KrscReports\Exception;

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
 * @version 2.0.1, 2018-12-21
 */

/**
 * Class for handling errors regarding improper size of row in comparison with column names.
 *
 * @category KrscReports
 * @package KrscReports_Exception
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class RowSizeMismatchException extends \KrscReports_Exception
{
    /**
     * Method generating exception message with details.
     *
     * @param array $columnNames array with column names
     * @param array $row         row, where error occurred
     * @param int   $rowKey      number of row, where error occurred
     *
     * @return string message for exception
     */
    public static function createExceptionMessage($columnNames, $row, $rowKey)
    {
        return sprintf(
            'Declared number of column names differs from the number of items for the analysed row!
            Array size for column names: %d;
            Array size for the items in analysed row: %d;
            RowNo: %d;
            Column names: %s',
            count($columnNames),
            count($row),
            $rowKey,
            implode(', ', $columnNames)
        );
    }
}
