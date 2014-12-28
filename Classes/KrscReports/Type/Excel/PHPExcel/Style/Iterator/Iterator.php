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
 * Implementation of iterator on collection of elements responsible for style of cell.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator implements KrscReports_Type_Excel_PHPExcel_Style_Iterator_IIterator
{
    /**
     * @var integer current position of iterator (starts from 0) 
     */
    protected $_iCurrentPosition;
    
    /**
     * @var KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection collection on which iteration is made 
     */
    protected $_oCollection;
    
    /**
     * Constructor for object - sets collection and initial position of iterator.
     * @param KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection $oCollection collection of elements responsible for style of cell
     * @return void
     */
    public function __construct( KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection $oCollection )
    {
        $this->_oCollection = $oCollection;
        $this->_iCurrentPosition = 0;
    }
    
    /**
     * Method returning current element of collection (each time executed it will give next element).
     * @return KrscReports_Type_Excel_PHPExcel_Style_Default element responsible for style of cell
     */
    public function getStyleElement() 
    {
        return $this->_oCollection->getElement( $this->_iCurrentPosition++ );
    }

    /**
     * Method checking, whether there are more elements in collection.
     * @return boolean true if there are more elements in collection, false otherwise
     */
    public function hasNextElement() 
    {      
        return $this->_oCollection->getElementCount() > $this->_iCurrentPosition;
    }
    
    /**
     * Method setting pointer to first element of collection.
     * @return KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator object on which method was executed
     */
    public function resetIterator()
    {
        $this->_iCurrentPosition = 0;
        return $this;
    }
}
