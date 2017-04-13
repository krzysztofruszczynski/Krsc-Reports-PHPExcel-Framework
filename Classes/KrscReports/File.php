<?php
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
 * @version 1.0.9, 2017-04-12
 */

/**
 * Class handling creation of file. For the time being only xsls is handled, in * future other formats would be also supported.
 * 
 * @category KrscReports
 * @package KrscReports
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_File
{
        /**
         * value for file type for Excel 2007
         */
        const FILE_TYPE_EXCEL = 'Excel2007';
    
	/**
	 * @var Boolean if true, writes header, when false - not
	 */
	protected $_bWriteHeader = true;

	/**
	 * @var String name of created file
	 */
	protected $_sFileName;

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
	 * @param String $sFileName name of file to be created
	 * @return KrscReports_File object on which method is executed
	 */
	public function setFileName( $sFileName )
	{
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
         * Setter for writer.
         * @param Object $oWriter (by default null - then PHPExcel writer is used)
         * @return KrscReports_File object on which method was executed
         */
	public function setWriter( $oWriter = null )
	{
		if( isset( $this->_oWriter ) )
		{	// previously set - no changes

		} else if( is_null( $oWriter ) ) {
                        $this->_oWriter = PHPExcel_IOFactory::createWriter( KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject(), $this->_sFileType );
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
            {	// previously set - no changes

            } else if( is_null( $oReader ) ) {
                $this->_oReader = PHPExcel_IOFactory::createReader( $this->_sFileType );

                if( $this->_oReader->canRead( $this->_sFileName ) ) {
                    KrscReports_Builder_Excel_PHPExcel::setPHPExcelObject( $this->_oReader->load( $this->_sFileName ) );
                } else {
                    throw new Exception('Unable to read file: ' . $this->_sFileName );
                }
                
            } else {
		$this->_oReader = $oReader;
            }
            
            return $this;
        }

        /**
         * Method for creating file (with headers if configured).
         * @return void
         */
	public function createFile()
	{
            if( $this->_bWriteHeader )
            {
		// remove possible previously set headers
                header_remove();
	    
	    	// header for content type of output
	    	header( 'Content-type: ' . $this->_aContentTypes[$this->_sExtension] );
	    
	    	// header for outputted filename
	    	header('Content-Disposition: attachment; filename="' . $this->_sFileName . '.' . $this->_sExtension . '"');
            }
	
            $this->setWriter();

	    // Write file to the browser
	    $this->_oWriter->save('php://output');
	}
}