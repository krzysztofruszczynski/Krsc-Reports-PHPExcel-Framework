<?php
namespace KrscReports\Type\Excel\PhpSpreadsheet\Style\Bundle;

use KrscReports\Type\Excel\PHPExcel\Style\Bundle\DefaultStyle;

/**
 * This file is part of KrscReports.
 *
 * Copyright (c) 2020 Krzysztof Ruszczyński
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
 * @copyright Copyright (c) 2020 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 2.0.6, 2020-02-17
 */

/**
 * Class for storing information about styles for tests results.
 *
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class PassFailStyle extends DefaultStyle
{
    /**
     * style key for pass style
     */
    const STYLE_PASS = 'style_pass';

    /**
     * style key for fail style
     */
    const STYLE_FAIL = 'style_fail';

    /**
     * color for fail style
     */
    const FILL_COLOR_FAIL = 'b8292f';

    /**
     * color for pass style
     */
    const FILL_COLOR_PASS = '00abb8';

    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection
     */
    protected $oCollectionPass;

    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection
     */
    protected $oCollectionFail;

    /**
     * Method for setting collections.
     */
    protected function setCollections()
    {
        parent::setCollections();

        $oFillPass = new \KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $oFillPass->setColor(self::FILL_COLOR_PASS);
        $this->oCollectionPass = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->oCollectionPass->addStyleElement($oFillPass);
        $this->oCollectionPass->addStyleElement(new \KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders());

        $oFillFail = new \KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $oFillFail->setColor(self::FILL_COLOR_FAIL);
        $this->oCollectionFail = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->oCollectionFail->addStyleElement($oFillFail);
        $this->oCollectionFail->addStyleElement(new \KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders());
    }

    protected function setStyleObject($style)
    {
        parent::setStyleObject($style);

        $style->setStyleCollection($this->oCollectionPass, self::STYLE_PASS);
        $style->setStyleCollection($this->oCollectionFail, self::STYLE_FAIL);
    }
}
