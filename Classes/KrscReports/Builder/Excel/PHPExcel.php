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
        self::$_oPHPExcel->setActiveSheetIndexByName( $sGroupName );
    }
}
?>
