<?php
namespace KrscReports\Views;

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
 * Report consisting of multiple views inheriting from AbstractView.
 * 
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class CompositeView extends AbstractView
{
    /**
     * @var array Views hierarchy (input from user):
     * - spreadsheet name
     *      - instances of abstract view 
     */
    protected $viewArray = array();
    
    /**
     * Method adding element to composer.
     * @param \KrscReports\Views\AbstractView $view view to add
     * @param string $spreadsheetName name of spreadsheet where data should be added
     * @return CompositeView object on which this method is executed
     */
    public function addView( AbstractView $view, $spreadsheetName = 'composite' )
    {
        $this->viewArray[$spreadsheetName][] = $view;
        return $this;
    }
    
    public function getDocumentElement( $spreadsheetName = 'worksheet' )
    {
        foreach ( $this->viewArray as $spreadsheetName => $viewCollection ) {            
            foreach ( $viewCollection as $view ) {   
                $viewOptions = $view->getOptions();
                $view->addOptions( $this->options ); // overrides previously set options
                $view->addOptions( $viewOptions ); // prevent from overriding previously set options
                
                $view->setDocumentElement( $this->documentElement ); // there is a reference so object is automatically changed
                $view->getDocumentElement( $spreadsheetName );
            }
        }
        
        return $this->documentElement;
    }
}
