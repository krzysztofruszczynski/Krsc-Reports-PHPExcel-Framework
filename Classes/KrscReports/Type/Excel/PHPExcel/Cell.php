<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Cell
{
    /**
     * @var PHPExcel 
     */
    protected static $_oPHPExcel;
    
    /**
     * @var Mixed
     */
    protected $_mValue;
    
    /**
     * @var String 
     */
    protected $_sType;
    
    /**
     * @var KrscReports_Type_Excel_PHPExcel_Style 
     */
    protected $_oStyle;
    
    /**
     * Update of PHPExcel object.
     * @return Void
     */
    public function __construct()
    {
        self::$_oPHPExcel = KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject();
    }
    
    public function setStyleObject( KrscReports_Type_Excel_PHPExcel_Style $oStyle )
    {
        $this->_oStyle = $oStyle;
    }
    
    public function setValue( $mValue )
    {
        $this->_mValue = $mValue;
    }
    
    public function setFormula( $mFormula )
    {
        
    }
    
    public function setType( $sType )
    {
        $this->_sType = $sType;
    }
    
    public function constructCell( $iColumnId, $iRowId )
    {
        //die( var_dump( $this->_oStyle->getStyleArray() ) );
        
        // setting styles
        self::$_oPHPExcel->getActiveSheet()->getStyleByColumnAndRow( $iColumnId, $iRowId )->applyFromArray( $this->_oStyle->getStyleArray() );
        
        // constructing phpexcel element
        self::$_oPHPExcel->getActiveSheet()->setCellValueByColumnAndRow( $iColumnId, $iRowId, $this->_mValue );
    }
    
}