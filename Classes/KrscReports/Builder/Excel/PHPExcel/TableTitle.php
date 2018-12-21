<?php
/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2018 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2018 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.2.6, 2018-06-01
 */

/**
 * Builder responsible for creating table, where there is a title before table.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_TableTitle extends KrscReports_Builder_Excel_PHPExcel_TableDifferentStyles implements KrscReports_Builder_Interface_Table
{       
    /**
     * @var string title of the table 
     */
    protected $_sTitle;

    /**
     * @var int additional info, how many columns should be used title (not set - one, 0 - all for this table)
     */
    protected $_iTitleColSpan;

    /**
     * Method setting title for this table.
     * @param string $sTitle title for table
     * @param int $iTitleColSpan info, how many columns should be used title (by default not set - one, 0 - all for this table)
     *
     * @return KrscReports_Builder_Excel_PHPExcel_TableTitle object on which method was executed
     */
    public function setTitle($sTitle, $iTitleColSpan = null)
    {
        $this->_sTitle = $sTitle;

        if (!is_null($iTitleColSpan)) {
            $this->_iTitleColSpan = intval($iTitleColSpan);
        }

        return $this;
    }

    /**
     * Action done when table begins.
     * @return void
     */
    public function beginTable() 
    {
        if( isset( $this->_sTitle ) )
        {
            $this->_oCell->setValue( $this->_sTitle );
            $this->_oCell->constructCell( $this->_iActualWidth, $this->_iActualHeight );

            if (isset($this->_iTitleColSpan)) {
                if ($this->_iTitleColSpan === 0) {
                    // take all columns when column names are set
                    $numberOfColumns = isset($this->_aColumnNames) ? (count($this->_aColumnNames) -1) : 0;
                } else {
                    $numberOfColumns = $this->_iTitleColSpan - 1;
                }
                $this->_oCell->mergeCells($this->_iActualWidth, $this->_iActualHeight, $this->_iActualWidth + $numberOfColumns, $this->_iActualHeight);
            }

            $this->_iActualHeight++;
        }

        parent::beginTable();
    }
}
