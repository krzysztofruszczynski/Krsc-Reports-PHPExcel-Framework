<?php

/**
 * Extension of abstract builder suitable with PHPExcel. Contains functionality closely connect with PHPExcel.
 * Method constructDocument() is still abstract and needs to be implemented further.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class KrscReports_Builder_Excel_PHPExcel extends KrscReports_Builder_Excel
{   
    /**
     * @var PHPExcel instance of PHPExcel used while adding new data to spreadsheets
     */
    protected static $_oPHPExcel;
    
    /**
     * @var KrscReports_Type_Excel_PHPExcel_Cell object responsible for creating cell inside Excel spreadsheet
     */
    protected $_oCell;
    
    /**
     * Method for setting objects responsible for creating cell inside Excel spreadsheet.
     * @param KrscReports_Type_Excel_PHPExcel_Cell $oCell object to be set to create cells inside Excel spreadsheet
     * @return KrscReports_Builder_Excel_PHPExcel object on which method was executed
     */
    public function setCellObject( $oCell )
    {
        $this->_oCell = $oCell;
        return $this;
    }
    
    /**
     * Method for setting PHPExcel object inside class.
     * @param PHPExcel $oPHPExcel PHPExcel object to be used while building Excel spreadsheets
     * @return void
     */
    public static function setPHPExcelObject( PHPExcel $oPHPExcel )
    {
        self::$_oPHPExcel = $oPHPExcel;
    }
    
    /**
     * Method returning PHPExcel object stored inside a class.
     * @return PHPExcel stored PHPExcel object
     * @throws Exception thrown when PHPExcel has not been set before (possibly only via self::setPHPExcelObject)
     */
    public static function getPHPExcelObject()
    {
        if( isset( self::$_oPHPExcel ) )
        {
            return self::$_oPHPExcel;
        }
        
        throw new Exception( 'PhpExcel object not set in KrscReports_Builder_Excel_PHPExcel::setPhpExcelObject( PhpExcel $oPhpExcel )' );
    }

    /**
     * Method for setting group name for current builder (used for Excel spreadsheet name).
     * @param string $sGroupName group name to be set (means that current builder will write to Excel spreadsheet with the same name)
     * @return KrscReports_Builder_Excel_PHPExcel object on which method was executed
     */
    public function setGroupName( $sGroupName )
    {
        $this->_sGroupName = $sGroupName;
        
        try
        {
            self::$_oPHPExcel->setActiveSheetIndexByName( $sGroupName );
        }
        catch ( PHPExcel_Exception $oException )
        {   // create new worksheet
            if( count( self::$_oPHPExcel->getAllSheets()) == 1 && self::$_oPHPExcel->getActiveSheet()->getTitle() == 'Worksheet' )
            {   // renaming default worksheet
                self::$_oPHPExcel->setActiveSheetIndex()->setTitle( $sGroupName );
            }
            else
            {   // second or later worksheet
                $oNewSheet = self::$_oPHPExcel->createSheet();
                $oNewSheet->setTitle( $sGroupName );                
            }
            
            self::$_oPHPExcel->setActiveSheetIndexByName( $sGroupName );
        }
        
        return $this;
    }
    
    /**
     * Method setting actual style key (can be set from element layer).
     * @param string $sStyleKey style key to be set
     * @param boolean $bIfStyleNotExistsSelectDefault (by default false) if style key was not found and parameter is true, 
     * there would be default style set; otherwise (if false) previously used style would continue
     * @return boolean true if style was found, false otherwise
     */
    public function setStyleKey( $sStyleKey, $bIfStyleNotExistsSelectDefault = false )
    {
        return $this->_oCell->setStyleKey( $sStyleKey, $bIfStyleNotExistsSelectDefault );
    }
}
