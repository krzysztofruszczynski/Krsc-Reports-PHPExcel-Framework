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
 * @copyright Copyright (c) 2018 Krzysztof Ruszczyński
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.2.4, 2018-03-23
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
     * @var string path, to which file will be generated (optional, if not provided it goes to php output)
     */
    protected $path;

    /**
     * @var \KrscReports_File logic for creating file
     */
    protected $file;

    /**
     *
     * @param array  $columns          subsequent elements are names of subsequent columns (to be translated, if second input parameter is provided)
     * @param \KrscReports\ColumnTranslatorService $columnTranslator if translator is provided, column names are translated
     * @param string $translatorDomain domain for translator
     */
    public function __construct($columns = array(), ColumnTranslatorService $columnTranslator = null, $translatorDomain = null)
    {
        $this->setColumns($columns, $columnTranslator, $translatorDomain);
        $this->file = new \KrscReports_File();
    }

    /**
     * Getter for object responsible for creating file.
     * @return \KrscReports_File object creating file
     */
    public function getFileObject()
    {
        return $this->file;
    }

    /**
     * Method for processing column names.
     *
     * @param  array  $columns  subsequent elements are names of subsequent columns (to be translated, if second input parameter is provided)
     * @param  \KrscReports\ColumnTranslatorService $columnTranslator   if translator is provided, column names are translated
     * @param  string $translatorDomain domain for translator
     * @return \KrscReports\Service object on which this method was executed
     */
    public function setColumns($columns = array(), ColumnTranslatorService $columnTranslator = null, $translatorDomain = null)
    {
        $this->columns = ( is_null($columnTranslator) || empty($columns) ) ? $columns : $columnTranslator->translateColumns($columns, $translatorDomain);
        return $this;
    }

    /**
     * Method for getting processed column names.
     *
     * @return array processed column names
     */
    public function getColumns()
    {
        return $this->columns;
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
     * Getter for file name.
     *
     * @return string file name for generated report
     */
    public function getFileName()
    {
        return $this->fileName;
    }

    /**
     * Setter for path of generated report.
     * @param string $path path, to which file will be generated (with extension, optional, if not provided it goes to php output)
     * @return \KrscReports\Service object on which this method was executed
     */
    public function setPath($path, $addExtension = false)
    {
        if ($addExtension) {
            $path = sprintf('%s.%s', $path, $this->file->getExtension());
        }
        $this->path = $path;
        return $this;
    }

    /**
     * Method for getting path of generated report.
     * @return string|null path of report (with extension) or null when element is not set
     */
    public function getPath()
    {
        if (!isset($this->path)) {
            return null;
        }

        return $this->path;
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
        $this->file->setFileName($this->fileName);

        $this->reportView->generate($this->spreadsheetName);

        $path = $this->getPath();
        if (is_null($path)) {   // create report to php output
            $this->file->createFile();
        } else {
            $this->file->createFileWithPath($path);
        }
    }
}
