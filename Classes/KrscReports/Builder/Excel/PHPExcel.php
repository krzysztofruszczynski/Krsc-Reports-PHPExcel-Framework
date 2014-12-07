<?php

/**
 * constructDocument() is still abstract
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class KrscReports_Builder_Excel_PHPExcel extends KrscReports_Builder_Excel
{   
    /**
     * @var PHPExcel 
     */
    protected static $_oPHPExcel;
    
    /**
     * @var KrscReports_Type_Excel_PHPExcel_Cell 
     */
    protected $_oCell;
    
    /**
     * 
     * @param KrscReports_Type_Excel_PHPExcel_Cell $oCell
     */
    public function setCellObject( $oCell )
    {
        $this->_oCell = $oCell;
    }
    
    /**
     * @todo : zrobić statyczną zmienną opisującą aktualną grupę? 
     */
    public static function setPHPExcelObject( PHPExcel $oPHPExcel )
    {
        self::$_oPHPExcel = $oPHPExcel;
    }
    
    /**
     * 
     * @return PHPExcel
     * @throws Exception
     */
    public static function getPHPExcelObject()
    {
        if( isset( self::$_oPHPExcel ) )
        {
            return self::$_oPHPExcel;
        }
        
        throw new Exception( 'PhpExcel object not set in KrscReports_Builder_Excel_PHPExcel::setPhpExcelObject( PhpExcel $oPhpExcel )' );
    }

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
        
    }
    
    /**
     * Method setting actual style key (can be set from element layer)
     * @param String $sStyleKey
     * @param Boolean $bIfStyleNotExistsSelectDefault
     * @return Boolean true if style was found, false otherwise
     */
    public function setStyleKey( $sStyleKey, $bIfStyleNotExistsSelectDefault = false )
    {
        return $this->_oCell->setStyleKey( $sStyleKey, $bIfStyleNotExistsSelectDefault );
    }
}
