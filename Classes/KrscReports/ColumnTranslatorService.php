<?php
namespace KrscReports;

/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2017 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2017 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.2.0, 2017-10-13
 */

/**
 * Service for translating columns.
 * 
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class ColumnTranslatorService
{
    protected $translator;
    
    public function __construct($translator)
    {
        $this->translator = $translator;
    }
    
    /**
     * Method returning translator actually set in this service (no need to pass it seperately when this service is passed)
     * @return Object
     */
    public function getTranslator()
    {
        return $this->translator;
    }
    
    public function translateColumns($columns, $translatorDomain)
    {
        foreach ($columns as $key => $value) {
            $columns[$key] = $this->translator->trans($value, array(), $translatorDomain);
        }
        
        return $columns;
    }
}

