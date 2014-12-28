<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2014 Krzysztof Ruszczyński
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
 * @package KrscReports_Document
 * @copyright Copyright (c) 2014 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.0, 2014-12-28
 */

/**
 * Class for creating table elements. Use of builder design pattern.
 * 
 * @category KrscReports
 * @package KrscReports_Document
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Document_Element_Table extends KrscReports_Document_Element_Default
{
    /**
     * style name for beginning of table
     */
    const STYLE_BEGIN_TABLE = 'beginTable';
    
    /**
     * style name for header with rows
     */
    const STYLE_HEADER_ROW = 'headerRow';
    
    /**
     * style name for normal row
     */
    const STYLE_ROW = 'row';
    
    /**
     * style name for summary row (for example with sums)
     */
    const STYLE_SUMMARY = 'summary';
    
    /**
     * Method responsible for creating document. Invokes methods specific for table builders.
     * @return KrscReports_Document_Element_Default object on which method was executed
     */
    public function constructDocument() 
    {
        // setting default style
        $this->_oBuilder->setStyleKey( KrscReports_Type_Excel_PHPExcel_Style::KEY_DEFAULT );
        
        $this->_oBuilder->setGroupName( $this->_sGroupName );
        
        // begins table (for example overall description of presented data)
        $this->_oBuilder->beginTable();
        
        // creating header row
        $this->_oBuilder->setStyleKey( self::STYLE_HEADER_ROW );
        $this->_oBuilder->setHeaderRow();
        
        // insert rows with data
        $this->_oBuilder->setStyleKey( self::STYLE_ROW );
        $this->_oBuilder->setRows();
        
        // rows under the the table (for example summary)
        $this->_oBuilder->setStyleKey( self::STYLE_SUMMARY );
        $this->_oBuilder->endTable();
        
        return $this;
    }
}

