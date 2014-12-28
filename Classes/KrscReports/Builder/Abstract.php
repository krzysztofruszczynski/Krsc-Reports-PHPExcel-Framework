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
 * @package KrscReports_Builder
 * @copyright Copyright (c) 2014 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.0.0, 2014-12-28
 */

/**
 * Abstract class for builders. Contains abstract constructDocument() method.
 * Implement methods for managing height and width, setting actual group name and data.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class KrscReports_Builder_Abstract
{
    /**
     * @var string name of group in which this content is generated 
     */
    protected $_sGroupName = '';
    
    /**
     * @var integer actual width for builder (starts from 0) 
     */
    protected $_iActualWidth = 0;
    
    /**
     * @var integer actual height for builder (starts from 1) 
     */
    protected $_iActualHeight;
        
    /**
     * @var array data set by user to display 
     */
    protected $_aData;
    
    /**
     * Method for setting start width for the builder.
     * @param integer $iStartWidth start width for element created by the builder (starts from 0)
     * @return KrscReports_Builder_Abstract object on which method was executed
     */
    public function setStartWidth( $iStartWidth )
    {
        $this->_iActualWidth = $iStartWidth;
        return $this;
    }
    
    /**
     * Method for setting start height for the builder.
     * @param integer $iStartHeight start height for element created by the builder (starts with 1)
     * @param boolean $bForceHeight if set to true, height is set even if someone else set it previous manually (by default true);
     * when it is false, actual height sets only if it wasn't before 
     * @return KrscReports_Builder_Abstract object on which method was executed
     */
    public function setStartHeight( $iStartHeight, $bForceHeight = true )
    {
        if( !isset( $this->_iActualHeight ) || $bForceHeight )
        {
            $this->_iActualHeight = $iStartHeight; 
        }        
        return $this;
    }
    
    /**
     * Method setting user data to be displayed.
     * @param array $aData data set by user (pattern: array( 0 => array('col1'=>... , 'col2'=>...), 1 => array( ... ) ) )
     * @return KrscReports_Builder_Abstract object on which method was executed
     */
    public function setData( $aData )
    {
        $this->_aData = $aData;
        return $this;
        
    }
    
    /**
     * Method returning actual width of builder.
     * @return integer actual width of builder 
     */
    public function getActualWidth()
    {
        return $this->_iActualWidth;
    }
    
    /**
     * Method returning actual height of builder.
     * @return integer actual height of builder
     */
    public function getActualHeight()
    {
        return $this->_iActualHeight;
    }
    
    /**
     * Method for setting group name.
     * @param string $sGroupName group name to be set
     * @return KrscReports_Builder_Abstract object on which method was executed
     */
    public function setGroupName( $sGroupName )
    {
        $this->_sGroupName = $sGroupName;
        return $this;
    }
    
    /**
     * Abstract method where builder constructs document.
     * @return void
     */
    abstract public function constructDocument();
}