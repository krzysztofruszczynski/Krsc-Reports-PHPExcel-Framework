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
     * @var KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection 
     */
    protected $_oStyleCollection;
    
    /**
     * Update of PHPExcel object.
     * @return Void
     */
    public function __construct()
    {
        self::$_oPHPExcel = KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject();
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
    
    public function setStyleCollection( KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection $oStyleCollection )
    {
        $this->_oStyleCollection = $oStyleCollection;
    }
    
    public function constructCell( $iColumnId, $iRowId )
    {
        // setting styles
        
        // constructing phpexcel element
        self::$_oPHPExcel->getActiveSheet()->setCellValueByColumnAndRow( $iColumnId, $iRowId, $this->_mValue );
    }
    
}