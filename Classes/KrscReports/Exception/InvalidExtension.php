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
 * @copyright Copyright (c) 2016 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.2, 2016-10-02
 */

/**
 * Class for handling errors with improper extension.
 *
 * @category KrscReports
 * @package KrscReports_Exception
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Exception_InvalidExtension extends KrscReports_Exception
{
    /**
     * @var string
     */
    protected $message = 'Requested extension is not handled by application';

    /**
     * @param string  $message
     * @param integer $code
     */
    public function __construct( $message, $code = 0 )
    {
        $this->message = $this->message . ' (' . $message . ')';
        parent::__construct($message);
    }
}
