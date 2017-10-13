<?php
namespace KrscReports;

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
 * Service for creating Excel files.
 * 
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class Service
{
    /**
     * @var string name of created file
     */
    protected $fileName;
    
    /**
     * @var string name of spreadsheet, where report would be created
     */
    protected $spreadsheetName;
    
    /**
     * @var \KrscReports\Views\AbstractView
     */
    protected $reportView;
    
    /**
     * @var array column names
     */
    protected $columns;
    
    /**
     *
     * @param array  $columns          subsequent elements are names of subsequent columns (to be translated, if second input parameter is provided)
     * @param \KrscReports\ColumnTranslatorService $columnTranslator if translator is provided, column names are translated
     * @param string $translatorDomain domain for translator
     */
    public function __construct($columns = array(), $columnTranslator = null, $translatorDomain = '')
    {
        $this->setColumns($columns, $columnTranslator, $translatorDomain);
    }
    
    /**
     * Method for processing column names.
     *
     * @param  array  $columns  subsequent elements are names of subsequent columns (to be translated, if second input parameter is provided)
     * @param  \KrscReports\ColumnTranslatorService $columnTranslator   if translator is provided, column names are translated
     * @param  string $translatorDomain domain for translator
     * @return \KrscReports\Service object on which this method was executed
     */
    public function setColumns($columns = array(), $columnTranslator = null, $translatorDomain = '')
    {
        $this->columns = ( is_null($columnTranslator) || empty($columns) ) ? $columns : $columnTranslator->translateColumns($columns, $translatorDomain);
        return $this;
    }
    
    /**
     * Setter for file name.
     *
     * @param  string $fileName file name for generated report
     * @return \KrscReports\Service object on which this method was executed
     */
    public function setFileName($fileName)
    {
        $this->fileName = $fileName;
        return $this;
    }
    
    /**
     * Setter for spreadsheet name.
     *
     * @param  string $spreadsheetName spreadsheet name for generated report
     * @return \KrscReports\Service object on which this method was executed
     */
    public function setSpreadsheetName($spreadsheetName)
    {
        $this->spreadsheetName = $spreadsheetName;
        return $this;
    }
    
    /**
     *
     * @param \KrscReports\Views\AbstractView $reportView
     * @param array $options options for report
     */
    public function setReportView($reportView, $options = array())
    {
        $this->reportView = $reportView;
        $this->reportView->addOptions($options);
        $this->reportView->setDocumentProperties();
    }
    
    /**
     * @return \KrscReports\Views\AbstractView
     */
    public function getReportView()
    {
        return $this->reportView;
    }
    
    /**
     * @param array $data
     * @param array $columnNames
     * @return \KrscReports\Service
     */
    public function setData($data)
    {
        $this->reportView->setData($data, $this->columns);
        return $this;
    }
    
    /**
     * Creates file with Excel report.
     */
    public function createReport()
    {
        $oFile = new \KrscReports_File();
        $oFile->setFileName($this->fileName);
     
        $this->reportView->generate($this->spreadsheetName);
        $oFile->createFile();
    }
}
