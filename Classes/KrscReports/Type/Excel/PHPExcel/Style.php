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
 * Class for managing style collections.
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style
{
    /**
     * default key of style collection
     */
    const KEY_DEFAULT = 'default';
    
    /**
     * @var array array of KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator
     */
    protected $_aStyleIterator;
    
    /**
     * Method setting style collection on given key.
     * @param KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection $oStyleCollection style collection to be set
     * @param string $sKey key under which given style collection will be set
     * @return KrscReports_Type_Excel_PHPExcel_Style object on which method was executed
     */
    public function setStyleCollection( KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection $oStyleCollection, $sKey = self::KEY_DEFAULT )
    {
        $this->_aStyleIterator[$sKey] = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator( $oStyleCollection );
        return $this;
    }
    
    /**
     * Method checking whether given key of style collection exists.
     * @param string $sKey key of style collection to verify
     * @return boolean true if style key exists, false otherwise
     */
    public function isValidStyleKey( $sKey )
    {
        return isset( $this->_aStyleIterator[$sKey] ) ? true : false;
    }
    
    /**
     * Method returns array with style properties ready to be inserted to PHPExcel.
     * @param string $sKey key of style collection which is retrieved (if it is not set, default key is used)
     * @return array style properties ready to be implemented via applyFromArray method on PHPExcel_Style object
     */
    public function getStyleArray( $sKey = self::KEY_DEFAULT ) 
    {
        if( !isset( $this->_aStyleIterator[$sKey] ) )
        {   // requested key not found - switching to default
            $sKey = self::KEY_DEFAULT;
            if( !isset( $this->_aStyleIterator[$sKey] ) )
            {   // default style not set - returning empty style array
                return array();
            }
        }
        
        $this->_aStyleIterator[$sKey]->resetIterator();
        $aOutput = array();
        
        while( $this->_aStyleIterator[$sKey]->hasNextElement() )
        {
            $oStyleElement = $this->_aStyleIterator[$sKey]->getStyleElement();
            $aOutput[$oStyleElement->getArrayKey()] = $oStyleElement->getStyleArray();
        }
        
        return $aOutput;
    }
}
