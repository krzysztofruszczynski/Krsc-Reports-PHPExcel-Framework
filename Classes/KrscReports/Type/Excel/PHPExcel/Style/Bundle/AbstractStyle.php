<?php
namespace KrscReports\Type\Excel\PHPExcel\Style\Bundle;

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
 * Abstract class for storing information about styles.
 * 
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class AbstractStyle
{
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style 
     */
    protected $style;
    
    abstract protected function setCollections();
    
    /**
     * To be override in non abstract classes.
     * @param \KrscReports_Type_Excel_PHPExcel_Style 
     */
    abstract protected function setStyleObject( $style );
    
    /**
     * 
     * @return \KrscReports_Type_Excel_PHPExcel_Style
     */
    public function getStyleObject()
    {
        if ( !isset( $this->collectionDefault ) ) {
            $this->setCollections();
        }
        
        if ( !isset( $this->style ) ) {
            $this->style = new \KrscReports_Type_Excel_PHPExcel_Style();
            $this->setStyleObject( $this->style );
        }
        
        return $this->style;
    }
}
