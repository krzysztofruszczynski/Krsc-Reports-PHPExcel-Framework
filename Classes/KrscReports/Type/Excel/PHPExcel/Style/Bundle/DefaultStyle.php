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
 * @version 1.1.6, 2017-10-13
 */

/**
 * Class for storing information about default styles.
 * 
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class DefaultStyle extends AbstractStyle
{    
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection if collection is shared, can be set in Collection folder
     */
    protected $collectionDefault;
    
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection if collection is shared, can be set in Collection folder
     */
    protected $collectionGray;
    
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection if collection is shared, can be set in Collection folder
     */
    protected $collectionBold;
    
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection if collection is shared, can be set in Collection folder
     */
    protected $collectionLightGrayBorders;
    
    /**
     * Method for setting collections (loaded from external files are created here).
     */
    protected function setCollections()
    {
        $this->collectionDefault = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->collectionDefault->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Borders_ExampleBorders() );

        $fillGray = new \KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $fillGray->setColor( \KrscReports_Type_Excel_PHPExcel_Style_Default::COLOR_GRAY );

        $this->collectionGray = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->collectionGray->addStyleElement( $fillGray );
        $this->collectionGray->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Borders_DoubleBorders() );
        $this->collectionGray->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Alignment_Wrap() );

        $this->collectionBold = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->collectionBold->addStyleElement( $fillGray );
        $this->collectionBold->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Borders_DoubleBorders() );
        $this->collectionBold->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Font_Bold() );
        $this->collectionBold->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Alignment_Wrap() );

        $fillLightGray = new \KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $fillLightGray->setColor( \KrscReports_Type_Excel_PHPExcel_Style_Default::COLOR_LIGHT_GRAY );
        $this->collectionLightGrayBorders = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->collectionLightGrayBorders->addStyleElement( $fillLightGray );
        $this->collectionLightGrayBorders->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        $this->collectionLightGrayBorders->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Alignment_Wrap() );

    }
    
    /**
     * 
     * @param \KrscReports_Type_Excel_PHPExcel_Style $style
     */
    protected function setStyleObject( $style )
    {
        $style->setStyleCollection( $this->collectionBold );
        $style->setStyleCollection( $this->collectionLightGrayBorders, \KrscReports_Document_Element_Table::STYLE_ROW );
        $style->setStyleCollection( $this->collectionGray, \KrscReports_Document_Element_Table::STYLE_HEADER_ROW );
    }

}