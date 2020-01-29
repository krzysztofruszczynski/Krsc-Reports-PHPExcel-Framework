<?php
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
 * @version 2.0.5, 2020-01-29
 */

/**
 * Class handling creation of file. For the time being only xslx is handled, in * future other formats would be also supported.
 * 
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_File
{
        /**
         * Date format, which can be used as prefix for files.
         */
        const DATE_FORMAT_SIMPLE = 'Y-m-d';

        /**
         * Value for settings associated with PHPExcel.
         */
        const SETTINGS_PHPEXCEL = 'PHPExcel';

        /**
         * Value for settings associated with PhpSpreadsheet.
         */
        const SETTINGS_PHPSPREADSHEET = 'PhpSpreadsheet';

        /**
         * value for file type for Excel 2007
         */
        const FILE_TYPE_EXCEL = 'Excel2007';

        /**
         * if true, then charts are displayed in output file
         */
        const INCLUDE_CHARTS = true;

        /**
         * @var Integer|Boolean if false, document is on the last page, 0 (integer) - first site (or any different spreadsheet index)
         */
        protected $_mReturnToSite = 0;

        /**
         * @var Boolean if true, writes header, when false - not
         */
        protected $_bWriteHeader = true;

        /**
         * @var string name of created file (if not additionally set, default value is used)
         */
        protected $_sFileName = 'report';

        /**
         * @var String extension of output file (for time being xlsx default, in future allowing more types of extensions)
         */
        protected $_sExtension = 'xlsx';

        /**
         * @var Array array where key is extension and value is content type for it
         */
        protected $_aContentTypes = array(
                'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );

        /**
         * @var string name of builder (by default PhpSpreadsheet)
         */
        protected static $_sBuilderType = self::SETTINGS_PHPSPREADSHEET;

        /**
         * @var string date format to be used as prefix for file names
         */
        protected $_sDateFormat = self::DATE_FORMAT_SIMPLE;

        /**
         * @var Object object responsible for creation of file
         */
        protected $_oWriter;

        /**
         * @var Object object responsible for reading of file
         */
        protected $_oReader;

        /**
         * @var string selected file type (used by PHPExcel writer and reader)
         */
        protected $_sFileType = self::FILE_TYPE_EXCEL;

        /**
         * Setter for file name.
         *
         * @param String $sFileName name of file to be created
         * @param bool $bPrefixData if true, name of file begin with timestamp (default: false)
         * @return KrscReports_File object on which method is executed
         */
        public function setFileName($sFileName, $bPrefixData = false)
        {
            if ($bPrefixData) {
                $sFileName = sprintf('%s_%s', date($this->_sDateFormat), $sFileName);
            }
            $this->_sFileName = $sFileName;

            return $this;
        }

        /**
         * Method for setting output extension.
         * @param String $sExtension extension to be set
         * @return KrscReports_File object on which method is executed
         * @throws KrscReports_Exception 
         */
        public function setExtension( $sExtension )
        {
            if( array_key_exists( $sExtension, $this->_aContentTypes ) )
            {
                $this->_sExtension = $sExtension;
            } else {
                throw new KrscReports_Exception_InvalidException( $sExtension );
            }

            return $this;
        }

        /**
         * Getter for actually set extension.
         * @return string extension
         */
        public function getExtension()
        {
            return $this->_sExtension;
        }

        /**
         * Method for setting builder type.
         *
         * @param string $sBuilderType
         */
        public static function setBuilderType($sBuilderType = self::SETTINGS_PHPSPREADSHEET)
        {
            self::$_sBuilderType = $sBuilderType;
        }

        /**
         * @return string type of builder
         */
        public static function getBuilderType()
        {
            return self::$_sBuilderType;
        }

        /**
         * Setter for deciding, if document want to return to given site.
         *
         * @param Integer|Boolean $mReturnToSite if false, document is on the last page, 0 (integer) - first site (or any different spreadsheet index)
         *
         * @return KrscReports_File object on which method was executed
         */
        public function setReturnToSite($mReturnToSite)
        {
            $this->_mReturnToSite = $mReturnToSite;

            return $this;
        }

        /**
         * Setter for writer.
         * @param Object $oWriter (by default null - then PHPExcel writer is used)
         * @return KrscReports_File object on which method was executed
         */
        public function setWriter( $oWriter = null )
        {
            if( isset( $this->_oWriter ) ) {   // previously set - no changes

            } else if( is_null( $oWriter ) ) {
                switch (self::getBuilderType()) {
                    case self::SETTINGS_PHPEXCEL:
                        $this->_oWriter = PHPExcel_IOFactory::createWriter(KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject(), $this->_sFileType);
                        $this->_oWriter->setIncludeCharts(self::INCLUDE_CHARTS);
                        $this->_oWriter->setPreCalculateFormulas(true);

                        if ($this->_mReturnToSite !== false) {
                            KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject()->setActiveSheetIndex($this->_mReturnToSite);
                        }
                        break;
                    case self::SETTINGS_PHPSPREADSHEET:
                        $this->_oWriter = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx(
                            \KrscReports\Builder\Excel\PhpSpreadsheet::getPhpSpreadsheetObject()
                        );
                        $this->_oWriter->setIncludeCharts(self::INCLUDE_CHARTS);
                        $this->_oWriter->setPreCalculateFormulas(true);

                        if ($this->_mReturnToSite !== false) {
                            \KrscReports\Builder\Excel\PhpSpreadsheet::getPhpSpreadsheetObject()->setActiveSheetIndex($this->_mReturnToSite);
                        }
                        break;
                }
            } else {
                $this->_oWriter = $oWriter;
            }

            return $this;
        }

        /**
         * Setter for reader.
         * @param Object $oReader (by default null - then PHPExcel reader is used)
         * @return KrscReports_File object on which method was executed
         */
        public function setReader( $oReader = null )
        {
            if( isset( $this->_oReader ) )
            {   // previously set - no changes

            } else if( is_null( $oReader ) ) {
                switch (self::getBuilderType()) {
                    case self::SETTINGS_PHPEXCEL:
                        $this->_oReader = PHPExcel_IOFactory::createReader($this->_sFileType);
                        if ($this->_oReader->canRead($this->_sFileName)) {
                            KrscReports_Builder_Excel_PHPExcel::setPHPExcelObject($this->_oReader->load($this->_sFileName));
                        } else {
                            throw new Exception('Unable to read file via PHPExcel: ' . $this->_sFileName );
                        }

                        break;
                    case self::SETTINGS_PHPSPREADSHEET:
                        $this->_oReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
                        if ($this->_oReader->canRead($this->_sFileName)) {
                            \KrscReports\Builder\Excel\PhpSpreadsheet::setSpreadsheetObject($this->_oReader->load($this->_sFileName));
                        } else {
                            throw new Exception('Unable to read file via PhpSpreadsheet: ' . $this->_sFileName );
                        }

                        break;
                }
            } else {
                $this->_oReader = $oReader;
            }

            return $this;
        }

        /**
         * Array with HTTP headers.
         * @return array each array key is header type, each array value is header value
         */
        public function createHeaderArray()
        {
            $aHeaders = array();
            // header for content type of output:
            $aHeaders['Content-type'] = $this->_aContentTypes[$this->_sExtension];
            // header for outputted filename:
            $aHeaders['Content-Disposition'] = 'attachment; filename="' . $this->_sFileName . '.' . $this->_sExtension . '"';

            return $aHeaders;
        }

        /**
         * Method setting headers for Symfony response.
         * @param Object $oHeaders $response->headers property
         */
        public function setSymfonyHeaders($oHtmlHeaders)
        {
            $aHeaders = $this->createHeaderArray();

            foreach ($aHeaders as $sHeaderKey => $sHeaderValue) {
                $oHtmlHeaders->set($sHeaderKey, $sHeaderValue);
            }

            // avoid creation of headers via createHeaders() method:
            $this->_bWriteHeader = false;
        }

        /**
         * Method creating headers using PHP header() method.
         */
        public function createHeaders()
        {
            // remove possible previously set headers:
            header_remove();
            $aHeaders = $this->createHeaderArray();

            foreach ($aHeaders as $sHeaderKey => $sHeaderValue) {
                header(sprintf('%s: %s', $sHeaderKey, $sHeaderValue));
            }
        }

        /**
         * Method for creating file (with headers if configured).
         * @return void
         */
        public function createFile()
        {
            if ($this->_bWriteHeader) {
                $this->createHeaders();
            }

            // Write file to the browser
            $this->createFileWithPath();
        }

        /**
         * Save file under specified path.
         *
         * @param string $sPath path, under which file would be created (default: php://output )
         * @param boolean $bAddExtension if true, extension is automatically added to file name ( with . )
         * @param Object object responsible for creation of file (optional)
         */
        public function createFileWithPath($sPath = 'php://output', $bAddExtension = false, $oWriter = null)
        {
            $this->setWriter($oWriter);

            $this->_oWriter->save($bAddExtension ? sprintf('%s.%s', $sPath, $this->_sExtension ) : $sPath);
        }
}
