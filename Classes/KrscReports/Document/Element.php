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
 * Class responsible for aggregating elements inside groups (use of composite design pattern).
 * Object of this class is able to construct document consisting of all elements stored in it. 
 * 
 * @category KrscReports
 * @package KrscReports_Document
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Document_Element
{
    /**
     * number of lines between elements in the spreadsheet
     */
    const LINES_BETWEEN_ELEMENTS = 2;
    
    /**
     * default name of group
     */
    const DEFAULT_GROUP_NAME = 'document1';
    
    /**
     * @var array array where group name is a key and value is an array with elements
     */
    protected $_aElements = array();
    
    /**
     * @var array array where group name is a key and value is actual width for this group
     */
    protected $_aActualWidths = array();
    
    /**
     * @var array array where group name is a key and value is actual height for this group
     */
    protected $_aActualHeights = array();
    
    /**
     * @var string name of group to which belongs this element 
     */
    protected $_sGroupName = self::DEFAULT_GROUP_NAME;
    
    /**
     * Method inserts subsequent element into specified group.
     * @param KrscReports_Document_Element $oElement element to be inserted
     * @param string $sGroupName name of group, to which inserting element belongs to (if null, default name of group is used)
     * @return KrscReports_Document_Element object on which method was executed
     */
    public function addElement( $oElement, $sGroupName = self::DEFAULT_GROUP_NAME )
    {
        $this->_aElements[$sGroupName][] = $oElement;
        return $this;
    }
    
    /**
     * Method sets group name for the whole object.
     * @param string $sGroupName name of group to be set
     * @return KrscReports_Document_Element object on which method was executed
     */
    public function setGroupName( $sGroupName )
    {
        $this->_sGroupName = $sGroupName;
        return $this;
    }
    
    /**
     * Code which is executed after creating document.
     * @return KrscReports_Document_Element object on which method was executed
     */
    public function afterConstructDocument() { return $this; }
    
    /**
     * Code which is executed before creating document.
     * @return KrscReports_Document_Element object on which method was executed
     */
    public function beforeConstructDocument() { return $this; }
    
    /**
     * Method iterating over all elements in composite. Before invoking method 
     * with the same name (composite design pattern), it sets group name, height and invokes
     * code to be executed before and after creating document.
     * @return KrscReports_Document_Element object on which method was executed
     */
    public function constructDocument()
    {
        // iteration over groups
        foreach( $this->_aElements as $sGroupName => $aGroupElements )
        {   
            // iteration over elements in group
            foreach( $aGroupElements as $oElement )
            {
                $oElement->setInnerGroupName( $sGroupName );
                
                // setting place of an element
                $oElement->setStartHeight( isset( $this->_aActualHeights[$sGroupName] ) ? $this->_aActualHeights[$sGroupName] + self::LINES_BETWEEN_ELEMENTS : 1, false );
                
                $oElement->beforeConstructDocument();
                $oElement->constructDocument();
                $oElement->afterConstructDocument();
                
                // updating actual height
                $this->_aActualHeights[$sGroupName] = $oElement->getActualHeight();
            }
        }
        
        return $this;
    }
}
