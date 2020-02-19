<?php
namespace KrscReports\Report;

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
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt       LGPL
 * @version 2.0.6, 2020-02-19
 */

/**
 * Interface for reports based on service.
 *
 * @category KrscReports
 * @package KrscReports_Report
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
interface ReportWithServiceInterface {

    /**
     * Method returning prepared view.
     *
     * @return \KrscReports\Views\AbstractView
     */
    public function getReportView();
}
