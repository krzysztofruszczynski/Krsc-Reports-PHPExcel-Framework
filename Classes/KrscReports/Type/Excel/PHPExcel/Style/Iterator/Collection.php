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
 * @package KrscReports_Type
 * @copyright Copyright (c) 2014 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.0, 2014-12-28
 */

/**
 * Class for creating collection of elements responsible for style of cells (Iterator design pattern).
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection implements KrscReports_Type_Excel_PHPExcel_Style_Iterator_IContainer
{
    /**
     * @var array subsequent elements which are part of collection 
     */
    protected $_aCollectionElements = array();
    
    /**
     * Method creating iterator for collection associated with current object.
     * @return KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator iterator for executed object
     */
    public function createIterator() 
    {
        return new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator( $this );
    }
    
    /**
     * Method for adding next element responsible for style of cell.
     * @param KrscReports_Type_Excel_PHPExcel_Style_Default $oElement style element to be added to current collection
     * @return KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection object on which method was executed
     */
    public function addStyleElement( KrscReports_Type_Excel_PHPExcel_Style_Default $oElement )
    {
        $this->_aCollectionElements[] = $oElement;
        return $this;
    }
    
    /**
     * Method returning style element with given position.
     * @param integer $iPosition position of element to return (starts from 0)
     * @return KrscReports_Type_Excel_PHPExcel_Style_Default requested element responsible for style of cell
     */
    public function getElement( $iPosition )
    {        
        return $this->_aCollectionElements[$iPosition];
    }
    
    /**
     * Method returning number of style elements associated with collection in current object.
     * @return integer number of style elements inside collection
     */
    public function getElementCount()
    {
        return count( $this->_aCollectionElements );
    }
}
