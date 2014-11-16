<?php

/**
 * constructDocument() is still abstract
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class KrscReports_Builder_Excel_PHPExcel extends KrscReports_Builder_Excel
{
    protected static $_oPHPExcel;
    
    /**
     * @todo : zrobić statyczną zmienną opisującą aktualną grupę? 
     */
    
    public static function setPHPExcelObject( PHPExcel $oPHPExcel )
    {
        self::$_oPHPExcel = $oPHPExcel;
    }
    
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
